import argparse
import ctypes
import time
from pathlib import Path

import cv2
import numpy as np
from PIL import ImageGrab


def set_clipboard(text: str) -> None:
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    kernel32.GlobalAlloc.argtypes = (ctypes.c_uint, ctypes.c_size_t)
    kernel32.GlobalAlloc.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = (ctypes.c_void_p,)
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = (ctypes.c_void_p,)
    user32.SetClipboardData.argtypes = (ctypes.c_uint, ctypes.c_void_p)
    user32.SetClipboardData.restype = ctypes.c_void_p
    data = ctypes.create_unicode_buffer(text)

    user32.OpenClipboard(None)
    try:
        user32.EmptyClipboard()
        handle = kernel32.GlobalAlloc(0x0002, ctypes.sizeof(data))
        pointer = kernel32.GlobalLock(handle)
        if not pointer:
            raise RuntimeError("Không khóa được bộ nhớ clipboard.")
        ctypes.memmove(pointer, ctypes.addressof(data), ctypes.sizeof(data))
        kernel32.GlobalUnlock(handle)
        user32.SetClipboardData(13, handle)
    finally:
        user32.CloseClipboard()


def press_key(key: int) -> None:
    ctypes.windll.user32.keybd_event(key, 0, 0, 0)
    ctypes.windll.user32.keybd_event(key, 0, 0x0002, 0)


def paste_and_enter(text: str) -> None:
    user32 = ctypes.windll.user32
    set_clipboard(text)

    user32.keybd_event(0x11, 0, 0, 0)
    press_key(0x41)
    user32.keybd_event(0x11, 0, 0x0002, 0)

    user32.keybd_event(0x11, 0, 0, 0)
    press_key(0x56)
    user32.keybd_event(0x11, 0, 0x0002, 0)
    time.sleep(0.2)
    press_key(0x0D)


def virtual_screen():
    user32 = ctypes.windll.user32
    left = user32.GetSystemMetrics(76)
    top = user32.GetSystemMetrics(77)
    width = user32.GetSystemMetrics(78)
    height = user32.GetSystemMetrics(79)
    return left, top, width, height


def find_template(template, timeout: float):
    left, top, width, height = virtual_screen()
    deadline = time.time() + timeout
    scales = (1.0, 0.9, 1.1, 0.8, 1.2, 0.75, 1.25)

    while time.time() < deadline:
        screen = ImageGrab.grab(
            bbox=(left, top, left + width, top + height),
            all_screens=True,
        )
        gray = cv2.cvtColor(np.array(screen), cv2.COLOR_RGB2GRAY)
        best = (-1.0, None, None)

        for scale in scales:
            candidate = cv2.resize(
                template,
                None,
                fx=scale,
                fy=scale,
                interpolation=cv2.INTER_AREA if scale < 1 else cv2.INTER_CUBIC,
            )
            if (
                candidate.shape[0] > gray.shape[0]
                or candidate.shape[1] > gray.shape[1]
            ):
                continue

            result = cv2.matchTemplate(gray, candidate, cv2.TM_CCOEFF_NORMED)
            _, score, _, location = cv2.minMaxLoc(result)
            if score > best[0]:
                best = (score, location, candidate.shape)

        score, location, shape = best
        if score >= 0.82:
            x = left + location[0] + shape[1] // 2
            y = top + location[1] + shape[0] // 2
            return x, y, score

        time.sleep(0.4)

    return None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--template", required=True)
    parser.add_argument("--text", required=True)
    parser.add_argument("--timeout", type=float, default=35)
    args = parser.parse_args()

    template_path = Path(args.template)
    template = cv2.imread(str(template_path), cv2.IMREAD_GRAYSCALE)
    if template is None:
        raise RuntimeError(f"Không đọc được ảnh mẫu: {template_path}")

    set_clipboard(args.text)
    match = find_template(template, args.timeout)
    if not match:
        return 2

    x, y, _ = match
    user32 = ctypes.windll.user32
    user32.SetCursorPos(x, y)
    user32.mouse_event(0x0002, 0, 0, 0, 0)
    user32.mouse_event(0x0004, 0, 0, 0, 0)
    time.sleep(0.35)
    paste_and_enter(args.text)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
