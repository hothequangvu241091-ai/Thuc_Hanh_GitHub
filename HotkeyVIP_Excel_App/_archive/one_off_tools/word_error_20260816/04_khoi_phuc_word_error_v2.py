# -*- coding: utf-8 -*-
"""V2: hai Edge tải trước nội dung song song, một Word lưu tuần tự."""

import argparse
import importlib.util
import json
import queue
import sys
import threading
from datetime import datetime
from pathlib import Path

import pythoncom

APP_ROOT = Path(__file__).resolve().parents[3]
V1_PATH = Path(__file__).with_name("04_khoi_phuc_word_error.py")
DEFAULT_INPUT = APP_ROOT / "outputs" / "worker_profile_word_errors_20260816" / "word_error_worker_mapping.xlsx"
LOG_DIR = APP_ROOT / "outputs" / "word_recovery_logs"
PROGRESS_PATH = Path(__file__).with_name("TIEN_DO_KHOI_PHUC_WORD_V2.txt")
QUEUE_MAX = 6

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)


def load_v1():
    spec = importlib.util.spec_from_file_location("hotkeyvip_word_recovery_v1", V1_PATH)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def parse_args():
    parser = argparse.ArgumentParser(description="V2: hai Edge tải bài, một Word lưu.")
    parser.add_argument("--input", default=str(DEFAULT_INPUT))
    parser.add_argument("--workers", nargs="+", default=["worker_1", "worker_3"])
    parser.add_argument("--limit", type=int, default=0)
    return parser.parse_args()


class State:
    def __init__(self, total, log_path):
        self.total = total
        self.log_path = log_path
        self.lock = threading.RLock()
        self.fetched = 0
        self.saved = 0
        self.skipped = 0
        self.failed = 0
        self.detail = "Đang khởi động"

    def record(self, result):
        with self.lock:
            status = result.get("status")
            if status == "FETCHED":
                self.fetched += 1
            elif status == "OK":
                self.saved += 1
            elif status == "ALREADY_OK":
                self.skipped += 1
            elif status == "ERROR":
                self.failed += 1
            self.detail = result.get("detail") or result.get("keyword") or ""
            with self.log_path.open("a", encoding="utf-8") as handle:
                handle.write(json.dumps(result, ensure_ascii=False, default=str) + "\n")
            self.write("ĐANG CHẠY")

    def write(self, status, detail=None):
        with self.lock:
            if detail is not None:
                self.detail = detail
            PROGRESS_PATH.write_text(
                f"TRẠNG THÁI: {status}\n"
                f"Đã tải vào hàng chờ: {self.fetched}\n"
                f"Word mới đã lưu: {self.saved}\n"
                f"Word đã có, bỏ qua: {self.skipped}\n"
                f"Lỗi: {self.failed}\n"
                f"Tổng danh sách: {self.total}\n"
                f"Cập nhật: {datetime.now().isoformat(timespec='seconds')}\n"
                f"Chi tiết: {self.detail}\n",
                encoding="utf-8",
            )


def queue_put(work_queue, item, stop_event):
    while not stop_event.is_set():
        try:
            work_queue.put(item, timeout=0.5)
            return True
        except queue.Full:
            pass
    return False


def is_driver_connection_error(exc):
    text = str(exc).casefold()
    return any(marker in text for marker in (
        "httpconnectionpool(host='localhost'",
        "read timed out",
        "connection refused",
        "invalid session id",
        "disconnected",
        "chrome not reachable",
    ))


def browser_worker(v1, flow, worker, jobs, work_queue, stop_event, state):
    driver = None
    flow._THREAD_CONTEXT.worker_id = int(worker.rsplit("_", 1)[1])
    try:
        driver, _ = flow.create_shared_driver()
        for position, job in enumerate(jobs, start=1):
            if stop_event.is_set():
                break
            if flow.is_word_ok(job["word_path"]):
                state.record({
                    **job, "status": "ALREADY_OK",
                    "detail": f"{worker} {position}/{len(jobs)}: Word đã có",
                    "time": datetime.now().isoformat(timespec="seconds"),
                })
                print(f"[{worker} {position}/{len(jobs)}] BỎ QUA: {job['keyword']}")
                continue

            for fetch_attempt in (1, 2):
                try:
                    print(
                        f"[{worker} {position}/{len(jobs)}] Mở URL"
                        f"{' sau khi dựng lại Edge' if fetch_attempt == 2 else ''}: "
                        f"{job['keyword']}"
                    )
                    driver.get(job["url"])
                    turns = v1.wait_stable_turns(driver, timeout=45, stable_seconds=0.8)
                    snapshot, method = v1.select_article_snapshot(turns)
                    word_count = flow.count_words(snapshot["text"])
                    if word_count < v1.MIN_SNAPSHOT_WORDS:
                        raise RuntimeError(f"Nội dung quá ngắn: {word_count} từ")
                    if v1.contains_any(snapshot["text"], v1.BRIEF_ANSWER_MARKERS):
                        raise RuntimeError("Nội dung có marker Brief")

                    item = {
                        "job": job, "snapshot": snapshot,
                        "word_count": word_count, "selection": method,
                    }
                    if not queue_put(work_queue, item, stop_event):
                        return
                    state.record({
                        **job, "status": "FETCHED", "word_count": word_count,
                        "detail": f"{worker} {position}/{len(jobs)}: tải trước {word_count} từ",
                        "time": datetime.now().isoformat(timespec="seconds"),
                    })
                    print(f"[{worker}] ĐÃ TẢI TRƯỚC: {word_count} từ | chờ Word: {work_queue.qsize()}")
                    break
                except Exception as exc:
                    if fetch_attempt == 1 and is_driver_connection_error(exc):
                        print(f"[{worker}] EdgeDriver mất kết nối; đang dựng lại rồi thử đúng bài...")
                        try:
                            driver.quit()
                        except Exception:
                            pass
                        flow._THREAD_CONTEXT.worker_id = int(worker.rsplit("_", 1)[1])
                        driver, _ = flow.create_shared_driver()
                        continue
                    state.record({
                        **job, "status": "ERROR", "stage": "FETCH", "error": str(exc),
                        "detail": f"{worker}: lỗi lấy {job['keyword']}",
                        "time": datetime.now().isoformat(timespec="seconds"),
                    })
                    print(f"[{worker}] LỖI LẤY BÀI: {exc}")
                    break
    except Exception as exc:
        state.record({
            "worker": worker, "status": "ERROR", "stage": "BROWSER_START",
            "error": str(exc), "detail": f"{worker}: không mở được Edge",
            "time": datetime.now().isoformat(timespec="seconds"),
        })
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        queue_put(work_queue, {"producer_done": worker}, stop_event)


def save_word(flow, item):
    for attempt in (1, 2):
        try:
            flow.copy_and_save_snapshot(item["snapshot"], item["job"]["word_path"])
            if not flow.is_word_ok(item["job"]["word_path"]):
                raise RuntimeError("Word lưu xong nhưng đọc kiểm tra không đạt")
            return
        except Exception as exc:
            is_system = isinstance(exc, flow.WordSystemError) or flow.is_word_system_error(exc)
            if attempt == 1 and is_system:
                print("[WORD] COM lỗi lần 1: dọn Word, preflight rồi thử lại bài hiện tại...")
                if flow.recover_word_runtime_once():
                    continue
            raise


def main():
    args = parse_args()
    workers = list(dict.fromkeys(args.workers))
    v1 = load_v1()
    flow = v1.load_flow_module()
    jobs = v1.read_jobs(Path(args.input).resolve(), set(workers))
    if args.limit > 0:
        jobs = jobs[:args.limit]
    if not jobs:
        raise RuntimeError("Không có bài để phục hồi")

    LOG_DIR.mkdir(parents=True, exist_ok=True)
    log_path = LOG_DIR / f"word_recovery_v2_{datetime.now():%Y%m%d_%H%M%S}.jsonl"
    state = State(len(jobs), log_path)
    state.write("ĐANG KHỞI ĐỘNG")
    print(f"V2: {len(jobs)} bài | 2 Edge tải trước + 1 Word | log: {log_path}")
    print("[WORD] Preflight...")
    flow.run_word_preflight("phục hồi Word V2")
    print("[WORD] Preflight đạt.")

    work_queue = queue.Queue(maxsize=QUEUE_MAX)
    stop_event = threading.Event()
    threads = []
    for worker in workers:
        worker_jobs = [job for job in jobs if job["worker"] == worker]
        if not worker_jobs:
            continue
        thread = threading.Thread(
            target=browser_worker,
            args=(v1, flow, worker, worker_jobs, work_queue, stop_event, state),
            name=f"Browser-{worker}",
            daemon=True,
        )
        thread.start()
        threads.append(thread)

    pythoncom.CoInitialize()
    producers_done = 0
    try:
        while producers_done < len(threads):
            try:
                item = work_queue.get(timeout=2)
            except queue.Empty:
                if not any(thread.is_alive() for thread in threads):
                    break
                continue
            try:
                if "producer_done" in item:
                    producers_done += 1
                    print(f"[{item['producer_done']}] Đã tải xong danh sách.")
                    continue
                job = item["job"]
                print(f"[WORD] Lưu dòng {job['excel_row']}: {job['keyword']}")
                try:
                    save_word(flow, item)
                    state.record({
                        **job, "status": "OK", "word_count": item["word_count"],
                        "selection": item["selection"],
                        "detail": f"Word vừa lưu: {job['keyword']}",
                        "time": datetime.now().isoformat(timespec="seconds"),
                    })
                    print(f"[WORD] OK: {item['word_count']} từ")
                except Exception as exc:
                    state.record({
                        **job, "status": "ERROR", "stage": "WORD", "error": str(exc),
                        "detail": f"Word lỗi: {job['keyword']}",
                        "time": datetime.now().isoformat(timespec="seconds"),
                    })
                    is_system = isinstance(exc, flow.WordSystemError) or flow.is_word_system_error(exc)
                    if is_system:
                        stop_event.set()
                        state.write("ĐÃ DỪNG DO WORD/COM", str(exc))
                        try:
                            import winsound
                            for _ in range(5):
                                winsound.Beep(1200, 350)
                        except Exception:
                            pass
                        raise RuntimeError("Word/COM vẫn lỗi; V2 đã dừng an toàn") from exc
                    print(f"[WORD] LỖI RIÊNG BÀI, chuyển tiếp: {exc}")
            finally:
                work_queue.task_done()
    finally:
        stop_event.set()
        for thread in threads:
            thread.join(timeout=20)
        flow.cleanup_all_owned_word_processes()
        pythoncom.CoUninitialize()

    state.write("HOÀN TẤT", "Đã xử lý hết danh sách")
    print(f"HOÀN TẤT V2: {state.saved} Word mới | {state.skipped} đã có | {state.failed} lỗi")
    print(f"Log: {log_path}")


if __name__ == "__main__":
    main()
