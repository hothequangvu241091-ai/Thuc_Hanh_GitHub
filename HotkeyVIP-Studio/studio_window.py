from __future__ import annotations

import os
import subprocess
import sys
import time
import urllib.request
import webbrowser
from pathlib import Path


APP_DIR = Path(__file__).resolve().parent
URL = "http://127.0.0.1:8765/"


def server_is_ready() -> bool:
    try:
        with urllib.request.urlopen(URL + "api/health", timeout=0.6) as response:
            return response.status == 200
    except Exception:
        return False


def start_server() -> None:
    if server_is_ready():
        return
    executable = Path(sys.executable)
    pythonw = executable.with_name("pythonw.exe")
    if pythonw.exists():
        executable = pythonw
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    subprocess.Popen(
        [str(executable), str(APP_DIR / "app.py"), "--no-browser"],
        cwd=str(APP_DIR),
        creationflags=flags,
    )
    for _ in range(50):
        if server_is_ready():
            return
        time.sleep(0.2)
    raise RuntimeError("HotkeyVIP Studio không khởi động được máy chủ.")


def find_edge() -> Path | None:
    candidates = [
        Path(os.environ.get("PROGRAMFILES(X86)", "")) / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("PROGRAMFILES", "")) / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft/Edge/Application/msedge.exe",
    ]
    return next((path for path in candidates if path.is_file()), None)


def open_app_window() -> None:
    edge = find_edge()
    if edge:
        subprocess.Popen(
            [
                str(edge),
                f"--app={URL}",
                "--start-maximized",
                "--disable-features=msEdgeSidebarV2",
            ],
            cwd=str(APP_DIR),
        )
        return
    webbrowser.open(URL)


def main() -> None:
    start_server()
    open_app_window()


if __name__ == "__main__":
    main()
