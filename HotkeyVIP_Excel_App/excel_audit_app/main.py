from __future__ import annotations

import ctypes
from ctypes import wintypes

from .ui import run_app


MUTEX_NAME = "Local\\ExcelAuditApp_HotkeyVIP_SingleInstance"
WINDOW_TITLE_PREFIX = "Đối soát nội dung Excel"
ERROR_ALREADY_EXISTS = 183
SW_RESTORE = 9


def _activate_existing_window() -> bool:
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    found = False

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def callback(hwnd: int, _lparam: int) -> bool:
        nonlocal found
        length = user32.GetWindowTextLengthW(hwnd)
        if length <= 0:
            return True
        buffer = ctypes.create_unicode_buffer(length + 1)
        user32.GetWindowTextW(hwnd, buffer, len(buffer))
        if buffer.value.startswith(WINDOW_TITLE_PREFIX):
            user32.ShowWindow(hwnd, SW_RESTORE)
            user32.SetForegroundWindow(hwnd)
            found = True
            return False
        return True

    user32.EnumWindows(callback, 0)
    return found


def main() -> None:
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    kernel32.CreateMutexW.argtypes = (wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR)
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.CloseHandle.argtypes = (wintypes.HANDLE,)
    kernel32.CloseHandle.restype = wintypes.BOOL

    mutex = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if not mutex:
        raise ctypes.WinError(ctypes.get_last_error())
    already_running = ctypes.get_last_error() == ERROR_ALREADY_EXISTS
    if already_running:
        try:
            if not _activate_existing_window():
                ctypes.WinDLL("user32", use_last_error=True).MessageBoxW(
                    None,
                    "App đối soát Excel đang chạy. Hãy dùng cửa sổ đã mở.",
                    "App đã được mở",
                    0x40,
                )
        finally:
            kernel32.CloseHandle(mutex)
        return

    try:
        run_app()
    finally:
        kernel32.CloseHandle(mutex)


if __name__ == "__main__":
    main()
