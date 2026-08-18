from __future__ import annotations

import argparse
import ctypes
import gc
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Any


def _configure_utf8() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconfigure = getattr(stream, "reconfigure", None)
        if callable(reconfigure):
            reconfigure(encoding="utf-8", errors="replace")


def _console_python() -> str:
    executable = Path(sys.executable)
    if executable.name.casefold() == "pythonw.exe":
        candidate = executable.with_name("python.exe")
        if candidate.exists():
            return str(candidate)
    return str(executable)


def _release(value: Any) -> None:
    if value is None:
        return
    try:
        import pythoncom

        if pythoncom.IsObject(value):
            del value
    except Exception:
        pass


def _run_flow_in_process(script_path: Path, workbook: Any) -> int:
    """Chạy flow COM cùng tiến trình để truyền chính xác workbook Excel ẩn."""
    module_name = "hotkeyvip_excel_app_in_process_flow"
    spec = importlib.util.spec_from_file_location(module_name, script_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Không thể nạp flow: {script_path.name}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.APP_WORKBOOK = workbook
    try:
        return_code = int(module.main() or 0)
    finally:
        module.APP_WORKBOOK = None
        gc.collect()
    return return_code


def _cleanup_flow_python_children(parent_pid: int) -> int:
    """Dọn Python worker còn sống của một flow mà không đụng Python ngoài cây process đó."""
    try:
        import psutil
    except Exception:
        return 0

    try:
        parent = psutil.Process(parent_pid)
        children = parent.children(recursive=True)
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return 0

    targets = []
    for child in children:
        try:
            name = child.name().casefold()
            if name.startswith("python") or name in {"py.exe", "pyw.exe"}:
                targets.append(child)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    if not targets:
        return 0

    identities = []
    for child in targets:
        try:
            identities.append((child.pid, child.create_time()))
            child.terminate()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue

    _gone, alive = psutil.wait_procs(targets, timeout=3)
    for child in alive:
        try:
            child.kill()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    if alive:
        psutil.wait_procs(alive, timeout=2)

    cleaned = 0
    for pid, create_time in identities:
        try:
            process = psutil.Process(pid)
            if abs(process.create_time() - create_time) < 0.01 and process.is_running():
                continue
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
        cleaned += 1

    if cleaned:
        print(
            f"[APP] Đã dọn {cleaned} Python worker còn sót của flow đăng bài.",
            flush=True,
        )
    return cleaned


def _run_flow_subprocess(
    workbook_path: Path, script_path: Path, script_args: list[str]
) -> int:
    environment = os.environ.copy()
    environment["HOTKEYVIP_SELECTED_EXCEL"] = str(workbook_path)
    environment["HOTKEYVIP_APP_RUN"] = "1"
    environment["PYTHONIOENCODING"] = "utf-8"
    app_root = str(script_path.parent.parent)
    existing_pythonpath = environment.get("PYTHONPATH", "").strip()
    environment["PYTHONPATH"] = (
        os.pathsep.join((app_root, existing_pythonpath))
        if existing_pythonpath
        else app_root
    )
    command = [_console_python(), "-u", str(script_path), *script_args]
    print("[APP] Bắt đầu chạy flow...", flush=True)
    process = subprocess.Popen(
        command,
        cwd=str(script_path.parent),
        env=environment,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
        bufsize=1,
    )
    if process.stdout is not None:
        for line in process.stdout:
            print(line, end="", flush=True)
            if (
                script_path.name == "05_dang_bai_cms.py"
                and line.strip().startswith("Thời gian kết thúc:")
            ):
                # Flow 5 đã chạy hết khối finally của coordinator. Nếu worker
                # multiprocessing nào còn sống sau join(timeout=8), dọn đúng
                # Python con của process Flow 5 để parent không bị kẹt lúc thoát.
                _cleanup_flow_python_children(process.pid)
    return_code = int(process.wait())
    print(f"[APP] Flow kết thúc với mã {return_code}.", flush=True)
    return return_code


def _open_owned_workbook(workbook_path: Path) -> tuple[Any, Any]:
    """Mở đúng đường dẫn bằng Excel ẩn của app; lỗi thì dọn và thử lại một lần."""
    import win32com.client as win32

    expected = os.path.normcase(os.path.abspath(str(workbook_path)))
    last_error: BaseException | None = None
    for attempt in range(1, 3):
        excel = None
        workbook = None
        try:
            excel = win32.DispatchEx("Excel.Application")
            _register_owned_excel(excel)
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.ScreenUpdating = False
            excel.EnableEvents = False
            excel.AskToUpdateLinks = False
            excel.AutomationSecurity = 3
            workbook = excel.Workbooks.Open(str(workbook_path), 0, False)
            actual = os.path.normcase(os.path.abspath(str(workbook.FullName)))
            if actual != expected:
                raise RuntimeError(
                    f"Excel mở sai workbook. Cần: {workbook_path}; nhận: {workbook.FullName}"
                )
            if bool(workbook.ReadOnly):
                raise RuntimeError(
                    "Workbook đang mở chỉ đọc. Hãy đóng file này trong Excel rồi chạy lại."
                )
            return excel, workbook
        except BaseException as exc:
            last_error = exc
            if workbook is not None:
                try:
                    workbook.Close(SaveChanges=False)
                except Exception:
                    pass
            _release(workbook)
            if excel is not None:
                identity = _owned_excel_identity(excel)
                try:
                    excel.Quit()
                    _remove_owned_excel(identity)
                except Exception:
                    pass
            _release(excel)
            if attempt == 1:
                print(
                    "[APP] Excel ẩn của app mở lỗi; đã đóng và đang thử lại...",
                    flush=True,
                )
    raise RuntimeError(f"Không thể mở đúng workbook theo đường dẫn: {last_error}")


def _same_path(left: Any, right: Any) -> bool:
    try:
        return os.path.normcase(os.path.abspath(str(left))) == os.path.normcase(
            os.path.abspath(str(right))
        )
    except Exception:
        return False


def _file_is_locked(path: Path) -> bool:
    """Kiểm tra khóa file mà không mở workbook bằng Excel."""
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    create_file = kernel32.CreateFileW
    create_file.argtypes = (
        ctypes.c_wchar_p,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_void_p,
        ctypes.c_uint32,
        ctypes.c_uint32,
        ctypes.c_void_p,
    )
    create_file.restype = ctypes.c_void_p
    handle = create_file(str(path), 0x80000000 | 0x40000000, 0, None, 3, 0x80, None)
    invalid_handle = ctypes.c_void_p(-1).value
    if handle == invalid_handle:
        return True
    kernel32.CloseHandle(ctypes.c_void_p(handle))
    return False


def _find_in_active_excel(workbook_path: Path) -> tuple[Any, Any] | None:
    import win32com.client as win32

    try:
        excel = win32.GetActiveObject("Excel.Application")
        for index in range(1, excel.Workbooks.Count + 1):
            workbook = excel.Workbooks(index)
            if _same_path(workbook.FullName, workbook_path):
                return excel, workbook
    except Exception:
        pass
    return None


def _attach_locked_workbook(workbook_path: Path) -> tuple[Any, Any]:
    """Bám đúng workbook đã mở, kể cả workbook thuộc một Excel instance khác."""
    import win32com.client as win32

    try:
        workbook = win32.GetObject(str(workbook_path))
        if not _same_path(workbook.FullName, workbook_path):
            raise RuntimeError(f"COM trả về sai workbook: {workbook.FullName}")
        if bool(workbook.ReadOnly):
            raise RuntimeError("workbook đang mở ở chế độ chỉ đọc")
        return workbook.Application, workbook
    except Exception as exc:
        raise RuntimeError(
            "File đang bị khóa nhưng app không thể bám đúng workbook đang mở. "
            "Hãy lưu file Excel đang mở rồi thử lại."
        ) from exc


def _owned_excel_registry_path() -> Path:
    path = Path(__file__).resolve().parents[1] / "_runtime"
    path.mkdir(parents=True, exist_ok=True)
    return path / "app_owned_excel_processes.json"


def _owned_excel_identity(excel: Any) -> dict[str, Any] | None:
    try:
        import psutil
        import win32process
        _thread_id, process_id = win32process.GetWindowThreadProcessId(int(excel.Hwnd))
        process = psutil.Process(int(process_id))
        return {"pid": int(process_id), "create_time": float(process.create_time())}
    except Exception:
        return None


def _read_owned_excel_registry() -> list[dict[str, Any]]:
    path = _owned_excel_registry_path()
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        return [item for item in data if isinstance(item, dict)]
    except (OSError, ValueError, TypeError):
        return []


def _write_owned_excel_registry(items: list[dict[str, Any]]) -> None:
    path = _owned_excel_registry_path()
    path.write_text(json.dumps(items, ensure_ascii=False, indent=2), encoding="utf-8")


def _register_owned_excel(excel: Any) -> None:
    identity = _owned_excel_identity(excel)
    if identity is None:
        return
    items = [
        item for item in _read_owned_excel_registry()
        if int(item.get("pid", -1)) != identity["pid"]
    ]
    items.append(identity)
    _write_owned_excel_registry(items)


def _remove_owned_excel(identity: dict[str, Any] | None) -> None:
    if identity is None:
        return
    remaining = [
        item for item in _read_owned_excel_registry()
        if not (
            int(item.get("pid", -1)) == identity["pid"]
            and abs(float(item.get("create_time", 0)) - identity["create_time"]) < 0.01
        )
    ]
    _write_owned_excel_registry(remaining)


def _process_has_visible_window(process_id: int) -> bool:
    try:
        import win32gui
        import win32process
    except Exception:
        return True
    visible = False

    def inspect_window(hwnd: int, _extra: Any) -> bool:
        nonlocal visible
        try:
            if win32gui.IsWindowVisible(hwnd):
                _thread_id, owner_pid = win32process.GetWindowThreadProcessId(hwnd)
                if int(owner_pid) == int(process_id):
                    visible = True
                    return False
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(inspect_window, None)
    except Exception:
        return True
    return visible


def _has_visible_excel_process() -> bool:
    try:
        import psutil
        for process in psutil.process_iter(["pid", "name"]):
            try:
                if (
                    str(process.info.get("name") or "").casefold() == "excel.exe"
                    and _process_has_visible_window(int(process.info["pid"]))
                ):
                    return True
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
    except Exception:
        return True
    return False


def _cleanup_tracked_hidden_excel() -> int:
    """Chỉ dọn PID đã được app ghi nhận, đúng process và không có cửa sổ hiển thị."""
    try:
        import psutil
    except Exception:
        return 0

    cleaned = 0
    remaining: list[dict[str, Any]] = []
    for item in _read_owned_excel_registry():
        try:
            pid = int(item["pid"])
            create_time = float(item["create_time"])
            process = psutil.Process(pid)
            is_same_process = abs(process.create_time() - create_time) < 0.01
            is_excel = process.name().casefold() == "excel.exe"
            if not is_same_process or not is_excel:
                continue
            if _process_has_visible_window(pid):
                remaining.append(item)
                continue
            print(f"[APP] Đang dọn Excel ẩn do app sở hữu: PID {pid}", flush=True)
            process.terminate()
            try:
                process.wait(timeout=3)
            except psutil.TimeoutExpired:
                process.kill()
                process.wait(timeout=2)
            cleaned += 1
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        except Exception:
            remaining.append(item)
    _write_owned_excel_registry(remaining)
    return cleaned


def _connect_or_open_workbook(workbook_path: Path) -> tuple[Any, Any, bool]:
    active = _find_in_active_excel(workbook_path)
    if active is not None:
        excel, workbook = active
        if bool(workbook.ReadOnly):
            raise RuntimeError("Workbook đúng đường dẫn đang mở ở chế độ chỉ đọc.")
        print("[APP] Đã bám đúng workbook đang mở sẵn.", flush=True)
        return excel, workbook, False

    if _file_is_locked(workbook_path):
        cleaned_owned_excel = _cleanup_tracked_hidden_excel()
        if cleaned_owned_excel:
            for _attempt in range(10):
                if not _file_is_locked(workbook_path):
                    excel, workbook = _open_owned_workbook(workbook_path)
                    print(
                        "[APP] Đã dọn Excel ẩn bị kẹt và mở lại đúng file.",
                        flush=True,
                    )
                    return excel, workbook, True
                time.sleep(0.2)
        if not _has_visible_excel_process():
            raise RuntimeError(
                "File đang bị một Excel ẩn từ phiên bản cũ giữ khóa. App không tự tắt vì "
                "tiến trình này chưa có hồ sơ sở hữu; hãy đóng Excel ẩn đó một lần rồi chạy lại."
            )
        excel, workbook = _attach_locked_workbook(workbook_path)
        print("[APP] Đã bám workbook đang mở trong một Excel instance khác.", flush=True)
        return excel, workbook, False

    try:
        excel, workbook = _open_owned_workbook(workbook_path)
        print("[APP] File chưa mở; app đã mở một Excel ẩn riêng.", flush=True)
        return excel, workbook, True
    except RuntimeError:
        # Bảo hiểm cho trường hợp file vừa được người dùng mở đúng lúc app khởi động.
        if _file_is_locked(workbook_path):
            excel, workbook = _attach_locked_workbook(workbook_path)
            print("[APP] File vừa được mở; app đã tự chuyển sang bám workbook đó.", flush=True)
            return excel, workbook, False
        raise


def run_flow(workbook_path: Path, script_path: Path, script_args: list[str]) -> int:
    if not workbook_path.is_file():
        raise RuntimeError(f"Không tìm thấy file Excel: {workbook_path}")
    if not script_path.is_file():
        raise RuntimeError(f"Không tìm thấy file flow: {script_path}")

    # Một nguồn sự thật cho mọi flow, kể cả flow chạy in-process. Các module
    # đọc biến môi trường ngay lúc import sẽ luôn nhận đúng file app đang chọn.
    os.environ["HOTKEYVIP_SELECTED_EXCEL"] = str(workbook_path)
    os.environ["HOTKEYVIP_APP_RUN"] = "1"

    excel = None
    workbook = None
    owns_excel = False
    workbook_saved = False
    old_calculation = None
    try:
        print(f"[APP] File Excel: {workbook_path}", flush=True)
        print(f"[APP] Flow: {script_path.name}", flush=True)

        print("[APP] Đang tìm đúng workbook theo đường dẫn...", flush=True)
        excel, workbook, owns_excel = _connect_or_open_workbook(workbook_path)
        if owns_excel:
            try:
                old_calculation = excel.Calculation
                excel.Calculation = -4135
            except Exception:
                old_calculation = None
        workbook.Activate()

        in_process_flows = {
            "01_nhap_ke_hoach.py",
            "02_chuan_bi_viet_bai.py",
            "04_chuan_bi_dang_bai.py",
            "06_lay_url_cms.py",
            "08_dong_bo_url.py",
        }
        if script_path.name in in_process_flows:
            print(
                f"[APP] Chạy {script_path.name} trực tiếp trên workbook Excel ẩn...",
                flush=True,
            )
            return_code = _run_flow_in_process(script_path, workbook)
            workbook.Save()
            workbook_saved = True
            print(f"[APP] Flow kết thúc với mã {return_code}.", flush=True)
            return return_code

        return_code = _run_flow_subprocess(workbook_path, script_path, script_args)
        try:
            workbook.Save()
            workbook_saved = True
        except Exception as exc:
            print(f"[APP] Cảnh báo khi lưu workbook lần cuối: {exc}", flush=True)
        return return_code
    finally:
        if workbook is not None and owns_excel:
            try:
                workbook.Close(SaveChanges=not workbook_saved)
            except Exception:
                pass
        _release(workbook)
        if excel is not None and owns_excel:
            if old_calculation is not None:
                try:
                    excel.Calculation = old_calculation
                except Exception:
                    pass
            identity = _owned_excel_identity(excel)
            try:
                excel.Quit()
                _remove_owned_excel(identity)
            except Exception:
                pass
        _release(excel)


def main() -> int:
    _configure_utf8()
    parser = argparse.ArgumentParser(description="Chạy flow trên file Excel được app chọn")
    parser.add_argument("--workbook", required=True)
    parser.add_argument("--script", required=True)
    parser.add_argument("script_args", nargs=argparse.REMAINDER)
    arguments = parser.parse_args()
    script_args = list(arguments.script_args)
    if script_args[:1] == ["--"]:
        script_args = script_args[1:]
    try:
        return run_flow(
            Path(arguments.workbook).resolve(),
            Path(arguments.script).resolve(),
            script_args,
        )
    except BaseException as exc:
        print(f"[APP] LỖI: {exc}", file=sys.stderr, flush=True)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())