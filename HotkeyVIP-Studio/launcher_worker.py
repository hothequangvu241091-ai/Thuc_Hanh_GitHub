from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
import time
import uuid
from datetime import datetime
from pathlib import Path


MAX_LOG_BYTES = 1 * 1024 * 1024
MAX_LOGS_PER_LAUNCHER = 50
LONG_RUN_LOG_INTERVAL_SECONDS = 60 * 60


def clean_old_logs(log_dir: Path) -> None:
    logs = sorted(log_dir.glob("*.log"), key=lambda path: path.stat().st_mtime, reverse=True)
    for old_log in logs[MAX_LOGS_PER_LAUNCHER - 1:]:
        status_file = old_log.with_suffix(".json")
        old_log.unlink(missing_ok=True)
        status_file.unlink(missing_ok=True)


def wait_on_error(show_console: bool) -> None:
    if not show_console:
        return
    try:
        input("\nChương trình bị lỗi. Nhấn Enter để đóng cửa sổ này...")
    except (EOFError, KeyboardInterrupt):
        pass


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--target", required=True)
    parser.add_argument("--log-dir", required=True)
    parser.add_argument("--show-console", action="store_true")
    args = parser.parse_args()

    target = Path(args.target).resolve()
    log_dir = Path(args.log_dir).resolve()
    log_dir.mkdir(parents=True, exist_ok=True)
    clean_old_logs(log_dir)

    stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    run_id = uuid.uuid4().hex[:8]
    log_path = log_dir / f"{stamp}_{run_id}.log"
    status_path = log_path.with_suffix(".json")
    started = time.time()
    written = 0
    truncated = False
    last_long_run_log_at = started
    pending_output = ""

    def write_log(text: str, *, force: bool = False) -> None:
        nonlocal written, truncated, last_long_run_log_at
        if truncated:
            return
        now = time.time()
        if not force and now - started >= LONG_RUN_LOG_INTERVAL_SECONDS:
            if now - last_long_run_log_at < LONG_RUN_LOG_INTERVAL_SECONDS:
                return
            last_long_run_log_at = now
        encoded = text.encode("utf-8", errors="replace")
        remaining = MAX_LOG_BYTES - written
        if remaining <= 0:
            truncated = True
            return
        chunk = encoded[:remaining]
        with log_path.open("ab") as stream:
            stream.write(chunk)
        written += len(chunk)
        if len(chunk) < len(encoded):
            truncated = True

    def write_status(state: str, exit_code: int | None = None, error: str = "") -> None:
        payload = {
            "state": state,
            "target": str(target),
            "log": str(log_path),
            "startedAt": started,
            "endedAt": time.time() if state != "running" else None,
            "duration": round(time.time() - started, 2),
            "exitCode": exit_code,
            "truncated": truncated,
            "error": error,
        }
        status_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def write_program_output(text: str, *, flush: bool = False) -> None:
        """Ghi đầu ra của chương trình, thêm giờ cho mỗi dòng hoàn chỉnh."""
        nonlocal pending_output
        pending_output += text.replace("\r\n", "\n").replace("\r", "\n")
        lines = pending_output.splitlines(keepends=True)
        pending_output = ""
        if lines and not lines[-1].endswith("\n") and not flush:
            pending_output = lines.pop()

        for line in lines:
            content = line.rstrip("\n")
            if content:
                write_log(f"[{datetime.now():%H:%M:%S}] {content}\n")
            else:
                write_log("\n")

        if flush and pending_output:
            write_log(f"[{datetime.now():%H:%M:%S}] {pending_output}\n", force=True)
            pending_output = ""

    header = (
        "HOTKEYVIP STUDIO - NHẬT KÝ CHẠY PYTHON\n"
        f"File: {target}\n"
        f"Bắt đầu: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
        + "=" * 72 + "\n"
    )
    write_log(header)
    write_status("running")
    if args.show_console:
        print(header, end="", flush=True)

    python_exe = Path(sys.executable).with_name("python.exe")
    if not python_exe.exists():
        python_exe = Path(sys.executable)
    environment = os.environ.copy()
    environment["PYTHONIOENCODING"] = "utf-8"
    environment["PYTHONUNBUFFERED"] = "1"

    try:
        process = subprocess.Popen(
            [str(python_exe), "-u", str(target)],
            cwd=str(target.parent),
            stdin=None,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            env=environment,
        )
        assert process.stdout is not None
        while True:
            chunk = process.stdout.read1(4096)
            if not chunk:
                break
            text = chunk.decode("utf-8", errors="replace")
            write_program_output(text)
            if args.show_console:
                print(text, end="", flush=True)
        write_program_output("", flush=True)
        exit_code = process.wait()
    except Exception as exc:
        message = f"\n[STUDIO] Không thể chạy file: {exc}\n"
        write_log(message)
        if args.show_console:
            print(message, end="", flush=True)
        write_status("error", -1, str(exc))
        wait_on_error(args.show_console)
        return 1

    footer = (
        "\n" + "=" * 72 + "\n"
        f"Kết thúc: {datetime.now():%Y-%m-%d %H:%M:%S}\n"
        f"Mã thoát: {exit_code}\n"
        f"Thời gian: {time.time() - started:.2f} giây\n"
    )
    if truncated:
        footer += "Log đã đạt giới hạn 1 MB; chương trình vẫn tiếp tục chạy.\n"
    write_log(footer, force=True)
    if args.show_console:
        print(footer, end="", flush=True)
    state = "success" if exit_code == 0 else "error"
    write_status(state, exit_code)
    if exit_code != 0:
        wait_on_error(args.show_console)
    return exit_code


if __name__ == "__main__":
    raise SystemExit(main())
