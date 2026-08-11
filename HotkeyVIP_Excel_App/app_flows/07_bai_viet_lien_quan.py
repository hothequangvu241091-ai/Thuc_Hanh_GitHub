# -*- coding: utf-8 -*-
"""
CHẠY HẾT BÀI VIẾT LIÊN QUAN BẰNG SELENIUM

- Chạy tuần tự toàn bộ bài chưa xử lý trong sheet "dangbai".
- Khi bắt đầu, chốt một lần toàn bộ các dòng thỏa:
    + Cột "Bài viết liên quan" đang trống.
    + Có "Tên miền".
    + Có "ID".
- Không dùng pyautogui, clipboard hoặc tọa độ INI.
- Không chiếm chuột và bàn phím.
- Edge dùng profile SeleniumData, mở ngoài màn hình rồi ẩn.
- Chỉ dùng đúng workbook Excel đang mở; không tự mở Excel ẩn.

Yêu cầu:
    pip install selenium pywin32
"""

from __future__ import annotations

import configparser
import ctypes
import os
import queue
import re
import subprocess
import sys
import threading
import time
import winsound
from pathlib import Path
from typing import Any

import win32com.client as win32
import win32con
import win32gui
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait

PROJECT_ROOT = Path(
    os.environ.get("HOTKEYVIP_RUNTIME_ROOT", r"D:\CodexProjects\Hotkeyvip")
).resolve()
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
from hotkeyvip_config import (
    EDGE_USER_DATA_DIR,
    EXCEL_FILE,
    LOGIN_INI,
    PUBLISH_HEADERS,
    SHEET_PUBLISH,
    SHEET_WEBSITE,
)


# ============================================================
# CẤU HÌNH
# ============================================================

EXCEL_PATH = Path(
    os.environ.get("HOTKEYVIP_SELECTED_EXCEL", str(EXCEL_FILE))
).resolve()
EXCEL_NAME = EXCEL_PATH.name
SHEET_NAME = SHEET_PUBLISH

NEAREST_RELATED_COUNT = 5
SAME_CATEGORY_RELATED_COUNT = 7
TARGET_RELATED_COUNT = NEAREST_RELATED_COUNT + SAME_CATEGORY_RELATED_COUNT

EDGE_USER_DATA_DIR = Path(EDGE_USER_DATA_DIR)

# Tài khoản dùng để tự đăng nhập khi session hết hạn.
LOGIN_INI = Path(LOGIN_INI)

WAIT_PAGE = 30
WAIT_DATA = 30
WAIT_AFTER_CATEGORY = 1.0
WAIT_AFTER_SELECT = 0.35
WAIT_AFTER_SAVE = 2.0

# Nếu True: lưu xong sẽ hiện Edge lại để kiểm tra.
# Nếu False: chạy đủ số bài đã chốt lúc bắt đầu rồi đóng Edge.
SHOW_EDGE_AFTER_DONE = False

# Nếu True: ghi URL vào cột "Bài viết liên quan" sau khi lưu thành công.
WRITE_STATUS_TO_EXCEL = True


# ============================================================
# BIẾN TOÀN CỤC
# ============================================================

_EDGE_HWND: int | None = None
_EDGE_IS_HIDDEN = False

_EXCEL_APP = None
_EXCEL_WB = None
_EXCEL_OPENED_BY_SCRIPT = False


# ============================================================
# HÀM CHUNG
# ============================================================

def clean(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_id(value: Any) -> str:
    text = clean(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text


def normalize_domain(value: Any) -> str:
    text = clean(value).lower()
    text = text.replace("https://", "").replace("http://", "")
    text = text.replace("www.", "")
    return text.strip("/")


def build_edit_url(create_url: str, post_id: str) -> str:
    """
    Dùng URL đăng bài trong CAU_HINH_WEBSITE và chèn ID CMS.

    Ví dụ:
        .../field/root_id=1
        -> .../field/id=90/root_id=1
    """
    source = clean(create_url).rstrip("/")
    normalized_post_id = normalize_id(post_id)
    if not source or not normalized_post_id:
        raise RuntimeError("Thiếu URL đăng bài hoặc ID CMS để tạo URL sửa bài.")

    edit_url, count = re.subn(
        r"/field(?:/id=\d+)?/root_id=",
        f"/field/id={normalized_post_id}/root_id=",
        source,
        count=1,
        flags=re.IGNORECASE,
    )
    if count != 1:
        raise RuntimeError(
            "URL trong CAU_HINH_WEBSITE không đúng cấu trúc "
            f'"/field/root_id=...": {source}'
        )
    return edit_url


def get_website_url_map(wb: Any) -> dict[str, str]:
    """Đọc một lần cột A/B của CAU_HINH_WEBSITE thành map tên miền -> URL."""
    ws = wb.Worksheets(SHEET_WEBSITE)
    last_row = int(ws.Cells(ws.Rows.Count, 1).End(-4162).Row)
    if last_row < 2:
        return {}

    domain_raw = ws.Range(ws.Cells(2, 1), ws.Cells(last_row, 1)).Value
    url_raw = ws.Range(ws.Cells(2, 2), ws.Cells(last_row, 2)).Value

    def column_values(raw: Any) -> list[Any]:
        if isinstance(raw, tuple):
            return [item[0] if isinstance(item, tuple) else item for item in raw]
        return [raw]

    result: dict[str, str] = {}
    for domain_value, url_value in zip(
        column_values(domain_raw),
        column_values(url_raw),
    ):
        domain = normalize_domain(domain_value)
        url = clean(url_value)
        if domain and url and domain not in result:
            result[domain] = url

    return result


def find_column_by_header(ws: Any, header_name: str) -> int:
    last_col = int(ws.Cells(1, ws.Columns.Count).End(-4159).Column)

    for col in range(1, last_col + 1):
        value = clean(ws.Cells(1, col).Value)
        if value.lower() == header_name.lower():
            return col

    raise RuntimeError(f'Không tìm thấy cột "{header_name}" trong sheet {SHEET_NAME}.')


# ============================================================
# EXCEL: CHỈ DÙNG ĐÚNG WORKBOOK ĐANG MỞ
# ============================================================

def _same_path(path_a: str, path_b: str) -> bool:
    try:
        return os.path.normcase(os.path.abspath(path_a)) == os.path.normcase(
            os.path.abspath(path_b)
        )
    except Exception:
        return clean(path_a).lower() == clean(path_b).lower()


def _excel_file_is_locked(path: Path) -> bool:
    kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
    create_file = kernel32.CreateFileW
    create_file.argtypes = (
        ctypes.c_wchar_p, ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p,
        ctypes.c_uint32, ctypes.c_uint32, ctypes.c_void_p,
    )
    create_file.restype = ctypes.c_void_p
    handle = create_file(str(path), 0x80000000 | 0x40000000, 0, None, 3, 0x80, None)
    if handle == ctypes.c_void_p(-1).value:
        return True
    kernel32.CloseHandle(ctypes.c_void_p(handle))
    return False


def connect_excel() -> tuple[Any, Any, bool]:
    global _EXCEL_APP, _EXCEL_WB, _EXCEL_OPENED_BY_SCRIPT

    # Ưu tiên đúng workbook đang mở. Nếu không thấy, tự mở đúng đường dẫn
    # bằng một phiên Excel ẩn do flow sở hữu và chỉ đóng phiên đó khi xong.
    try:
        app = win32.GetActiveObject("Excel.Application")
    except Exception:
        app = None

    if app is not None:
        try:
            for index in range(1, app.Workbooks.Count + 1):
                wb = app.Workbooks(index)
                if _same_path(clean(wb.FullName), str(EXCEL_PATH)):
                    print(f"Đã kết nối Excel đang mở: {wb.Name}")
                    _EXCEL_APP = app
                    _EXCEL_WB = wb
                    _EXCEL_OPENED_BY_SCRIPT = False
                    return app, wb, False
        except Exception:
            pass

    if not EXCEL_PATH.is_file():
        raise RuntimeError(f"Không tìm thấy file Excel theo đường dẫn: {EXCEL_PATH}")

    if _excel_file_is_locked(EXCEL_PATH):
        try:
            attached_wb = win32.GetObject(str(EXCEL_PATH))
            if not _same_path(clean(attached_wb.FullName), str(EXCEL_PATH)):
                raise RuntimeError(f"COM trả về sai workbook: {attached_wb.FullName}")
            if bool(attached_wb.ReadOnly):
                raise RuntimeError("Workbook đang mở ở chế độ chỉ đọc.")
            attached_app = attached_wb.Application
            print(f"Đã bám đúng file Excel đang mở: {attached_wb.Name}")
            _EXCEL_APP = attached_app
            _EXCEL_WB = attached_wb
            _EXCEL_OPENED_BY_SCRIPT = False
            return attached_app, attached_wb, False
        except Exception as exc:
            raise RuntimeError(
                f"File Excel đang mở nhưng không thể bám đúng workbook: {EXCEL_PATH}"
            ) from exc

    owned_app = None
    owned_wb = None
    try:
        owned_app = win32.DispatchEx("Excel.Application")
        owned_app.Visible = False
        owned_app.DisplayAlerts = False
        owned_app.EnableEvents = False
        owned_wb = owned_app.Workbooks.Open(str(EXCEL_PATH), 0, False)
        if not _same_path(clean(owned_wb.FullName), str(EXCEL_PATH)):
            raise RuntimeError(f"Excel đã mở sai workbook: {owned_wb.FullName}")
        if bool(owned_wb.ReadOnly):
            raise RuntimeError("Workbook được mở ở chế độ chỉ đọc.")
        print(f"Đã tự mở đúng file Excel ẩn: {owned_wb.Name}")
        _EXCEL_APP = owned_app
        _EXCEL_WB = owned_wb
        _EXCEL_OPENED_BY_SCRIPT = True
        return owned_app, owned_wb, True
    except Exception:
        if owned_wb is not None:
            try:
                owned_wb.Close(SaveChanges=False)
            except Exception:
                pass
        if owned_app is not None:
            try:
                owned_app.Quit()
            except Exception:
                pass
        raise


def close_excel_if_needed(save: bool = True) -> None:
    global _EXCEL_APP, _EXCEL_WB, _EXCEL_OPENED_BY_SCRIPT

    app = _EXCEL_APP
    wb = _EXCEL_WB
    opened_by_script = _EXCEL_OPENED_BY_SCRIPT

    _EXCEL_APP = None
    _EXCEL_WB = None
    _EXCEL_OPENED_BY_SCRIPT = False

    if not opened_by_script:
        return

    try:
        if wb is not None:
            wb.Close(SaveChanges=save)
    except Exception:
        pass

    try:
        if app is not None:
            app.Quit()
    except Exception:
        pass


def get_target_rows(
    ws: Any,
    website_url_map: dict[str, str],
) -> list[dict[str, Any]]:
    """Đọc Excel một lần và chốt cố định danh sách bài cần chạy."""
    col_domain = find_column_by_header(ws, PUBLISH_HEADERS["domain"])
    col_id = find_column_by_header(ws, PUBLISH_HEADERS["cms_id"])
    col_category = find_column_by_header(ws, PUBLISH_HEADERS["category"])
    col_bvlq = find_column_by_header(ws, PUBLISH_HEADERS["related"])

    last_row = int(ws.Cells(ws.Rows.Count, col_id).End(-4162).Row)
    if last_row < 2:
        return []

    domain_values_raw = ws.Range(
        ws.Cells(2, col_domain), ws.Cells(last_row, col_domain)
    ).Value
    id_values_raw = ws.Range(
        ws.Cells(2, col_id), ws.Cells(last_row, col_id)
    ).Value
    category_values_raw = ws.Range(
        ws.Cells(2, col_category), ws.Cells(last_row, col_category)
    ).Value
    status_values_raw = ws.Range(
        ws.Cells(2, col_bvlq), ws.Cells(last_row, col_bvlq)
    ).Value

    def column_values(raw: Any) -> list[Any]:
        if isinstance(raw, tuple):
            return [row[0] if isinstance(row, tuple) else row for row in raw]
        return [raw]

    domain_values = column_values(domain_values_raw)
    id_values = column_values(id_values_raw)
    category_values = column_values(category_values_raw)
    status_values = column_values(status_values_raw)

    targets: list[dict[str, Any]] = []
    for offset, (
        domain_value,
        id_value,
        category_value,
        status_value,
    ) in enumerate(
        zip(
            domain_values,
            id_values,
            category_values,
            status_values,
        ),
        start=2,
    ):
        domain = clean(domain_value)
        post_id = normalize_id(id_value)
        if domain and post_id and not clean(status_value):
            normalized_domain = normalize_domain(domain)
            create_url = website_url_map.get(normalized_domain, "")
            if not create_url:
                raise RuntimeError(
                    f"Không tìm thấy URL trong {SHEET_WEBSITE} "
                    f"cho tên miền: {domain}"
                )
            targets.append(
                {
                    "row": offset,
                    "domain": domain,
                    "post_id": post_id,
                    "category": clean(category_value),
                    "edit_url": build_edit_url(create_url, post_id),
                    "col_bvlq": col_bvlq,
                }
            )

    return targets


class RelatedProgressWindow:
    """Bảng tiến độ; yêu cầu dừng chỉ có hiệu lực sau khi lưu xong bài hiện tại."""

    def __init__(self, total: int) -> None:
        self.total = total
        self.stop_after_current = threading.Event()
        self._messages: queue.Queue[tuple[int, str]] = queue.Queue()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self) -> None:
        import tkinter as tk

        root = tk.Tk()
        root.title("Tiến độ bài viết liên quan")
        root.geometry("360x145+30+250")
        root.resizable(False, False)
        root.attributes("-topmost", True)

        title = tk.Label(
            root,
            text=f"Chuẩn bị chạy 0/{self.total}",
            font=("Segoe UI", 13, "bold"),
        )
        title.pack(pady=(18, 6))
        detail = tk.Label(root, text="Đang chuẩn bị Edge...", font=("Segoe UI", 10))
        detail.pack(pady=(0, 12))

        buttons = tk.Frame(root)
        buttons.pack()

        def request_stop() -> None:
            self.stop_after_current.set()
            stop_button.config(state="disabled", text="Sẽ dừng sau bài này")
            detail.config(text="Đã nhận lệnh dừng an toàn")

        tk.Button(buttons, text="Ẩn bảng", width=12, command=root.withdraw).pack(
            side="left", padx=5
        )
        stop_button = tk.Button(
            buttons,
            text="Dừng sau bài này",
            width=18,
            command=request_stop,
        )
        stop_button.pack(side="left", padx=5)

        # Nút X chỉ ẩn bảng, không cắt ngang bài đang xử lý.
        root.protocol("WM_DELETE_WINDOW", root.withdraw)

        def poll_messages() -> None:
            try:
                while True:
                    completed, message = self._messages.get_nowait()
                    if completed < 0:
                        root.destroy()
                        return
                    title.config(text=f"Đã hoàn thành {completed}/{self.total}")
                    detail.config(text=message)
            except queue.Empty:
                pass
            root.after(200, poll_messages)

        poll_messages()
        root.mainloop()

    def update(self, completed: int, message: str) -> None:
        self._messages.put((completed, message))

    def should_stop(self) -> bool:
        return self.stop_after_current.is_set()

    def notify_finished(self, completed: int, stopped: bool = False) -> None:
        if stopped:
            message = f"Đã dừng an toàn sau bài {completed}/{self.total}"
        else:
            message = f"Đã chạy xong {completed}/{self.total} bài"
        self.update(completed, message)
        try:
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except Exception:
            pass
        self.close()

    def close(self) -> None:
        """Đóng hẳn bảng tiến độ trên đúng luồng Tkinter."""
        self._messages.put((-1, ""))


# ============================================================
# KIỂM TRA VÀ TỰ ĐĂNG NHẬP TỪ FILE INI
# ============================================================

LOGIN_FORM_SELECTOR = "form#login-form"
LOGIN_BUTTON_SELECTOR = "form#login-form button[type='submit']"
POST_PAGE_SELECTOR = "#name"

USERNAME_SELECTORS = (
    "form#login-form input[type='text']",
    "form#login-form input[name='username']",
    "form#login-form input[name='user']",
    "form#login-form input[name='account']",
    "form#login-form input[id*='user']",
    "form#login-form input[placeholder*='Tài khoản']",
    "form#login-form input[placeholder*='tài khoản']",
)

PASSWORD_SELECTORS = (
    "form#login-form input[type='password']",
    "form#login-form input[name='password']",
    "form#login-form input[id*='pass']",
    "form#login-form input[placeholder*='Mật khẩu']",
    "form#login-form input[placeholder*='mật khẩu']",
)


def wait_document_ready(driver) -> None:
    WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def has_visible_element(driver, css_selector: str) -> bool:
    try:
        return any(
            element.is_displayed()
            for element in driver.find_elements(By.CSS_SELECTOR, css_selector)
        )
    except Exception:
        return False


def is_login_page(driver) -> bool:
    return (
        has_visible_element(driver, LOGIN_FORM_SELECTOR)
        and has_visible_element(driver, LOGIN_BUTTON_SELECTOR)
    )


def is_post_edit_page(driver) -> bool:
    return has_visible_element(driver, POST_PAGE_SELECTOR)


def read_login_credentials() -> tuple[str, str]:
    if not LOGIN_INI.is_file():
        raise RuntimeError(
            "Không tìm thấy file tài khoản:\n"
            f"{LOGIN_INI}\n\n"
            "Nội dung mẫu:\n"
            "[login]\n"
            "username=ten_dang_nhap\n"
            "password=mat_khau"
        )

    config = configparser.ConfigParser(interpolation=None)
    loaded = config.read(LOGIN_INI, encoding="utf-8-sig")
    if not loaded:
        raise RuntimeError(f"Không đọc được file tài khoản:\n{LOGIN_INI}")

    username_keys = ("username", "user", "id", "taikhoan", "tai_khoan")
    password_keys = ("password", "pass", "matkhau", "mat_khau")
    section_names = ("login", "taikhoan", "account")

    sections = []
    for section_name in section_names:
        if config.has_section(section_name):
            sections.append(config[section_name])
    sections.append(config.defaults())

    username = ""
    password = ""

    for section in sections:
        if not username:
            for key in username_keys:
                value = str(section.get(key, "") or "").strip()
                if value:
                    username = value
                    break

        if not password:
            for key in password_keys:
                value = str(section.get(key, "") or "")
                if value:
                    password = value
                    break

        if username and password:
            break

    if not username or not password:
        raise RuntimeError(
            "File tài khoản chưa có đủ ID và mật khẩu.\n"
            f"File: {LOGIN_INI}\n\n"
            "Nên dùng:\n"
            "[login]\n"
            "username=ten_dang_nhap\n"
            "password=mat_khau"
        )

    return username, password


def find_first_visible(driver, selectors: tuple[str, ...]):
    for selector in selectors:
        try:
            for element in driver.find_elements(By.CSS_SELECTOR, selector):
                if element.is_displayed() and element.is_enabled():
                    return element
        except Exception:
            continue
    return None


def fill_login_field(driver, element, value: str) -> None:
    element.click()
    element.send_keys(Keys.CONTROL, "a")
    element.send_keys(Keys.BACKSPACE)
    element.send_keys(value)


def ensure_post_page_ready(driver, target_url: str) -> None:
    """Nếu bị trả về trang login thì tự đăng nhập và quay lại URL bài."""
    driver.switch_to.default_content()
    wait_document_ready(driver)

    if is_post_edit_page(driver):
        print("    Đã ở đúng trang sửa bài, không cần đăng nhập.")
        return

    if not is_login_page(driver):
        raise RuntimeError(
            "Không thấy trang sửa bài hoặc form đăng nhập.\n"
            f"URL hiện tại: {driver.current_url}"
        )

    print("    Phát hiện trang đăng nhập → tự nhập ID và mật khẩu...")
    username, password = read_login_credentials()

    username_input = WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: find_first_visible(d, USERNAME_SELECTORS)
    )
    password_input = WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: find_first_visible(d, PASSWORD_SELECTORS)
    )

    fill_login_field(driver, username_input, username)
    fill_login_field(driver, password_input, password)

    login_button = WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: next(
            (
                element
                for element in d.find_elements(By.CSS_SELECTOR, LOGIN_BUTTON_SELECTOR)
                if element.is_displayed() and element.is_enabled()
            ),
            None,
        )
    )

    print("    Đã nhập ID/mật khẩu → click ĐĂNG NHẬP.")
    try:
        login_button.click()
    except Exception:
        driver.execute_script("arguments[0].click();", login_button)

    WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: (not is_login_page(d)) or is_post_edit_page(d)
    )
    wait_document_ready(driver)

    # Một số website đăng nhập xong về dashboard, nên mở lại đúng URL bài.
    if not is_post_edit_page(driver):
        driver.get(target_url)
        wait_document_ready(driver)

    if is_login_page(driver):
        raise RuntimeError(
            "Đăng nhập không thành công, website vẫn ở trang login. "
            "Hãy kiểm tra ID/mật khẩu trong taikhoan.ini."
        )

    WebDriverWait(driver, WAIT_PAGE).until(lambda d: is_post_edit_page(d))
    print("    Đăng nhập thành công, đã vào đúng trang sửa bài.")


# ============================================================
# EDGE ẨN / NGOÀI MÀN HÌNH
# ============================================================

def _find_edge_hwnd_by_marker(marker: str) -> int | None:
    matches: list[int] = []

    def callback(hwnd: int, _extra: object) -> None:
        try:
            if win32gui.IsWindow(hwnd) and marker in win32gui.GetWindowText(hwnd):
                matches.append(hwnd)
        except Exception:
            pass

    win32gui.EnumWindows(callback, None)
    return matches[0] if matches else None


def remember_edge_window(driver) -> int:
    global _EDGE_HWND

    marker = f"SELENIUM_BVLQ_{int(time.time() * 1000)}"
    old_title = str(driver.execute_script("return document.title || '';"))
    driver.execute_script("document.title = arguments[0];", marker)

    deadline = time.time() + 8
    hwnd = None

    while time.time() < deadline:
        hwnd = _find_edge_hwnd_by_marker(marker)
        if hwnd:
            break
        time.sleep(0.1)

    driver.execute_script("document.title = arguments[0];", old_title)

    if not hwnd:
        raise RuntimeError("Không xác định được cửa sổ Edge Selenium.")

    _EDGE_HWND = hwnd
    return hwnd


def hide_edge_window(driver) -> None:
    global _EDGE_IS_HIDDEN
    hwnd = _EDGE_HWND or remember_edge_window(driver)
    win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
    _EDGE_IS_HIDDEN = True
    print("Edge đã được ẩn.")


def show_edge_window() -> None:
    global _EDGE_IS_HIDDEN

    hwnd = _EDGE_HWND
    if not hwnd or not win32gui.IsWindow(hwnd):
        return

    try:
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        except Exception:
            pass

    _EDGE_IS_HIDDEN = False
    print("Edge đã hiện lại để kiểm tra.")


def close_existing_selenium_edge() -> None:
    global _EDGE_HWND, _EDGE_IS_HIDDEN

    profile = str(EDGE_USER_DATA_DIR.resolve()).replace("'", "''")
    script = f"""
    $profile = '{profile}'
    Get-CimInstance Win32_Process |
      Where-Object {{
        $_.Name -ieq 'msedge.exe' -and
        $_.CommandLine -and
        $_.CommandLine.IndexOf($profile, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
      }} |
      ForEach-Object {{
        Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue
      }}
    """

    subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ],
        check=False,
        capture_output=True,
        text=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )

    time.sleep(1.2)
    _EDGE_HWND = None
    _EDGE_IS_HIDDEN = False


def open_edge(url: str):
    print("Mở Edge Selenium...")

    EDGE_USER_DATA_DIR.mkdir(parents=True, exist_ok=True)

    options = Options()
    options.add_argument(f"--user-data-dir={EDGE_USER_DATA_DIR}")
    options.add_argument("--disable-notifications")
    options.add_argument("--window-position=-32000,-32000")
    options.add_argument("--window-size=1400,1000")

    driver = webdriver.Edge(options=options)
    driver.get(url)

    wait_document_ready(driver)
    ensure_post_page_ready(driver, url)

    remember_edge_window(driver)
    hide_edge_window(driver)
    return driver


# ============================================================
# SELENIUM: BÀI VIẾT LIÊN QUAN
# ============================================================

def element_text(element) -> str:
    return clean(
        element.text
        or element.get_attribute("textContent")
        or element.get_attribute("value")
    )


def find_visible_by_text(driver, selector: str, wanted: str):
    wanted_norm = clean(wanted).lower()

    for element in driver.find_elements(By.CSS_SELECTOR, selector):
        try:
            if not element.is_displayed():
                continue
            haystack = " ".join(
                clean(value).lower()
                for value in (
                    element.text,
                    element.get_attribute("value"),
                    element.get_attribute("title"),
                    element.get_attribute("aria-label"),
                )
                if value
            )
            if wanted_norm in haystack:
                return element
        except Exception:
            continue

    return None


def click_element(driver, element) -> None:
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    try:
        element.click()
    except Exception:
        driver.execute_script("arguments[0].click();", element)


def click_load_data(driver) -> None:
    print("[1] Click Tải dữ liệu...")

    button = WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: find_visible_by_text(
            d,
            "button.btn.btn-sm.blue, button, a.btn",
            "Tải dữ liệu",
        )
    )
    click_element(driver, button)

    # Chờ plugin multiselect được dựng và có vùng danh sách.
    WebDriverWait(driver, WAIT_DATA).until(
        lambda d: bool(
            d.find_elements(
                By.CSS_SELECTOR,
                "div[id^='ms-'][id$='multiselect'].ms-container, .ms-container",
            )
        )
    )

    print("    Đã tải dữ liệu bài liên quan.")


def get_visible_select2_search(driver):
    boxes = driver.find_elements(By.CSS_SELECTOR, "input.select2-search__field")
    visible = [box for box in boxes if box.is_displayed()]
    if not visible:
        raise RuntimeError("Không thấy ô tìm kiếm Select2 đang hiển thị.")
    return visible[-1]


def find_select2_category_control(driver):
    """
    Ưu tiên Select2 nằm ngay trong khu vực 'Bài viết liên quan'.
    Fallback dùng placeholder 'Lọc tin theo danh mục'.
    """
    return driver.execute_script(
        """
        const placeholders = Array.from(document.querySelectorAll(
            '.select2-selection__placeholder'
        ));

        for (const item of placeholders) {
            const text = String(item.textContent || '').trim().toLowerCase();
            if (text.includes('lọc tin theo danh mục')) {
                return item.closest('.select2-selection') || item;
            }
        }

        const heading = Array.from(document.querySelectorAll('h1,h2,h3,h4,label,div'))
            .find(el => String(el.textContent || '').trim().toLowerCase() === 'bài viết liên quan');

        if (heading) {
            const root = heading.parentElement || heading;
            return root.querySelector('.select2-selection');
        }

        return null;
        """
    )


def find_visible_select2_option(driver, wanted_text: str):
    wanted = clean(wanted_text).lower()
    options = driver.find_elements(
        By.CSS_SELECTOR,
        "li.select2-results__option[role='option'], li.select2-results__option",
    )
    visible = [item for item in options if item.is_displayed()]

    for item in visible:
        if element_text(item).lower() == wanted:
            return item

    for item in visible:
        text = element_text(item).lower()
        if not text:
            continue
        if wanted in text or text in wanted:
            return item

    return None


def choose_category(driver, category: str) -> None:
    target = clean(category) or "Tất cả"
    print(f"[2] Chọn danh mục: {target}")

    control = WebDriverWait(driver, WAIT_PAGE).until(find_select2_category_control)
    click_element(driver, control)

    search = WebDriverWait(driver, WAIT_PAGE).until(get_visible_select2_search)
    search.send_keys(Keys.CONTROL, "a")
    search.send_keys(Keys.BACKSPACE)
    search.send_keys(target)

    option = WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: find_visible_select2_option(d, target)
    )
    click_element(driver, option)

    WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: not any(
            item.is_displayed()
            for item in d.find_elements(By.CSS_SELECTOR, ".select2-container--open")
        )
    )

    time.sleep(WAIT_AFTER_CATEGORY)


def get_multiselect_container(driver):
    """
    Tìm đúng ms-container chứa select:
    name="list_news_id[bai-viet-lien-quan][]"
    """
    container = driver.execute_script(
        """
        const select = document.querySelector(
            'select[name="list_news_id[bai-viet-lien-quan][]"]'
        );
        if (!select) return null;

        let node = select.nextElementSibling;
        while (node) {
            if (node.classList && node.classList.contains('ms-container')) {
                return node;
            }
            node = node.nextElementSibling;
        }

        return select.parentElement
            ? select.parentElement.querySelector('.ms-container')
            : null;
        """
    )

    if not container:
        raise RuntimeError("Không tìm thấy vùng MultiSelect bài viết liên quan.")

    return container


def count_selected_related(driver) -> int:
    return int(
        driver.execute_script(
            """
            const select = document.querySelector(
                'select[name="list_news_id[bai-viet-lien-quan][]"]'
            );
            if (!select) return 0;
            return Array.from(select.options)
                .filter(option => option.selected).length;
            """,
        )
        or 0
    )


def get_selected_related_ids(driver) -> list[str]:
    values = driver.execute_script(
        """
        const select = document.querySelector(
            'select[name="list_news_id[bai-viet-lien-quan][]"]'
        );
        if (!select) return [];
        return Array.from(select.options)
            .filter(option => option.selected)
            .map(option => String(option.value || '').trim())
            .filter(Boolean);
        """
    )
    return [
        normalize_id(value)
        for value in (values or [])
        if normalize_id(value)
    ]


def get_selectable_items(driver):
    container = get_multiselect_container(driver)
    items = container.find_elements(
        By.CSS_SELECTOR,
        ".ms-selectable li.ms-elem-selectable:not(.ms-selected)",
    )

    result = []
    for item in items:
        try:
            if item.is_displayed() and element_text(item):
                result.append(item)
        except StaleElementReferenceException:
            continue

    return result


def select_until_target(driver, target_count: int) -> int:
    """
    Chọn nhanh nhiều bài liên quan trong một lượt JavaScript.
    Chỉ chọn số bài còn thiếu và kiểm tra một lần sau cùng.
    """
    selected_before = count_selected_related(driver)
    missing = target_count - selected_before

    if missing <= 0:
        return selected_before

    result = driver.execute_script(
        """
        const missing = arguments[0];

        const select = document.querySelector(
            'select[name="list_news_id[bai-viet-lien-quan][]"]'
        );

        if (!select) {
            return {
                ok: false,
                error: 'Không tìm thấy select bài viết liên quan'
            };
        }

        let container = select.nextElementSibling;

        if (
            !container
            || !container.classList
            || !container.classList.contains('ms-container')
        ) {
            container = select.parentElement
                ? select.parentElement.querySelector('.ms-container')
                : null;
        }

        if (!container) {
            return {
                ok: false,
                error: 'Không tìm thấy ms-container'
            };
        }

        const items = Array.from(
            container.querySelectorAll(
                '.ms-selectable li.ms-elem-selectable:not(.ms-selected)'
            )
        ).filter(item => {
            const style = window.getComputedStyle(item);
            const rect = item.getBoundingClientRect();

            return (
                style.display !== 'none'
                && style.visibility !== 'hidden'
                && rect.width > 0
                && rect.height > 0
            );
        });

        const chosen = items.slice(0, missing);

        for (const item of chosen) {
            item.click();
        }

        return {
            ok: true,
            clicked: chosen.length
        };
        """,
        missing,
    )

    if not result or not result.get("ok"):
        error_text = (
            result.get("error")
            if isinstance(result, dict)
            else str(result)
        )
        raise RuntimeError(
            "Không chọn nhanh được bài liên quan: " + error_text
        )

    clicked = int(result.get("clicked", 0))

    if clicked <= 0:
        return selected_before

    expected_minimum = selected_before + clicked

    try:
        WebDriverWait(driver, 10).until(
            lambda d: count_selected_related(d) >= expected_minimum
        )
    except TimeoutException:
        pass

    selected_after = count_selected_related(driver)

    print(
        f"    Đã chọn nhanh {clicked} bài, "
        f"tổng hiện tại {selected_after}/{target_count}"
    )

    return selected_after


def select_nearest_related_ids(
    driver,
    current_post_id: str,
    target_count: int,
) -> tuple[int, list[str]]:
    """
    Chọn trực tiếp theo value ID trong select:
    - ưu tiên ID lớn hơn gần nhất;
    - chưa đủ thì lấy ID nhỏ hơn gần nhất;
    - giữ nguyên các bài đã được chọn trước đó.
    """
    try:
        current_id = int(normalize_id(current_post_id))
    except (TypeError, ValueError) as exc:
        raise RuntimeError(
            f'ID bài hiện tại không hợp lệ: "{current_post_id}".'
        ) from exc

    snapshot = driver.execute_script(
        """
        const select = document.querySelector(
            'select[name="list_news_id[bai-viet-lien-quan][]"]'
        );
        if (!select) {
            return {ok:false, error:'Không tìm thấy select bài viết liên quan'};
        }

        const options = Array.from(select.options).map(option => ({
            value: String(option.value || '').trim(),
            selected: Boolean(option.selected),
            disabled: Boolean(option.disabled)
        }));
        return {ok:true, options};
        """
    )

    if not snapshot or not snapshot.get("ok"):
        error_text = (
            snapshot.get("error")
            if isinstance(snapshot, dict)
            else str(snapshot)
        )
        raise RuntimeError("Không đọc được ID bài viết liên quan: " + error_text)

    available_ids: list[int] = []
    selected_ids: list[int] = []

    for item in snapshot.get("options", []):
        if item.get("disabled"):
            continue
        try:
            value = int(normalize_id(item.get("value")))
        except (TypeError, ValueError):
            continue
        if value == current_id:
            continue
        if value not in available_ids:
            available_ids.append(value)
        if item.get("selected") and value not in selected_ids:
            selected_ids.append(value)

    if len(selected_ids) >= target_count:
        kept = [str(value) for value in selected_ids[:target_count]]
        print(f"    Bài đã có sẵn {len(selected_ids)} lựa chọn, không cần chọn thêm.")
        return len(selected_ids), kept

    selected_set = set(selected_ids)
    greater = sorted(
        value
        for value in available_ids
        if value > current_id and value not in selected_set
    )
    smaller = sorted(
        (
            value
            for value in available_ids
            if value < current_id and value not in selected_set
        ),
        reverse=True,
    )

    missing = target_count - len(selected_ids)
    new_ids = greater[:missing]
    if len(new_ids) < missing:
        new_ids.extend(smaller[: missing - len(new_ids)])

    if not new_ids and not selected_ids:
        return 0, []

    target_values = [str(value) for value in new_ids]
    result = driver.execute_script(
        """
        const values = arguments[0].map(String);
        const select = document.querySelector(
            'select[name="list_news_id[bai-viet-lien-quan][]"]'
        );
        if (!select) {
            return {ok:false, error:'Không tìm thấy select bài viết liên quan'};
        }

        const existing = new Set(
            Array.from(select.options)
                .filter(option => option.selected)
                .map(option => String(option.value))
        );
        const toSelect = values.filter(value => !existing.has(value));

        if (
            window.jQuery
            && window.jQuery.fn
            && typeof window.jQuery.fn.multiSelect === 'function'
        ) {
            window.jQuery(select).multiSelect('select', toSelect);
        } else {
            for (const option of Array.from(select.options)) {
                if (toSelect.includes(String(option.value))) {
                    option.selected = true;
                }
            }
            select.dispatchEvent(new Event('input', {bubbles:true}));
            select.dispatchEvent(new Event('change', {bubbles:true}));
        }

        const selected = Array.from(select.options)
            .filter(option => option.selected)
            .map(option => String(option.value));
        return {ok:true, selected};
        """,
        target_values,
    )

    if not result or not result.get("ok"):
        error_text = (
            result.get("error")
            if isinstance(result, dict)
            else str(result)
        )
        raise RuntimeError("Không chọn được ID bài viết liên quan: " + error_text)

    selected_after = [
        normalize_id(value)
        for value in result.get("selected", [])
        if normalize_id(value)
    ]
    missing_after = [
        value for value in target_values
        if value not in selected_after
    ]
    if missing_after:
        raise RuntimeError(
            "CMS không đánh dấu selected cho các ID: "
            + ", ".join(missing_after)
        )

    print(
        f"    ID hiện tại: {current_id} | "
        f"đã chọn thêm: {', '.join(target_values) or '(không có)'}"
    )
    print(
        f"    Tổng selected sau cùng: "
        f"{len(selected_after)}/{target_count}"
    )
    return len(selected_after), selected_after


def add_related_articles(
    driver,
    current_post_id: str,
    category: str,
) -> int:
    article_category = clean(category)
    if not article_category:
        raise RuntimeError(
            'Dòng Excel không có "Danh mục"; không chọn bài liên quan.'
        )

    click_load_data(driver)

    # Nhóm 1: lấy 5 ID gần nhất trên toàn bộ bài viết.
    choose_category(driver, "Tất cả")
    nearest_selected, nearest_ids = select_nearest_related_ids(
        driver,
        current_post_id,
        NEAREST_RELATED_COUNT,
    )

    # Nhóm 2: lọc đúng danh mục và chỉ click các mục đang hiển thị
    # trong bộ lọc đó, cho đến khi tổng cộng đủ 10.
    choose_category(driver, article_category)
    selected = select_until_target(driver, TARGET_RELATED_COUNT)
    selected_ids = get_selected_related_ids(driver)

    print(
        f"    Nhóm ID gần nhất: "
        f"{nearest_selected}/{NEAREST_RELATED_COUNT}"
    )
    if nearest_ids:
        print("    ID nhóm gần nhất:", ", ".join(nearest_ids))
    print(
        f"    Nhóm cùng danh mục \"{article_category}\": "
        f"chọn thêm đến đủ {TARGET_RELATED_COUNT}"
    )
    print(
        f"Kết quả cuối: {selected}/{TARGET_RELATED_COUNT} bài liên quan."
    )
    if selected < TARGET_RELATED_COUNT:
        print(
            f"NOTE: Khong du bai viet lien quan: "
            f"{selected}/{TARGET_RELATED_COUNT} "
            f"(thieu {TARGET_RELATED_COUNT - selected} bai). "
            "Van tiep tuc luu bai."
        )
    if selected_ids:
        print("    Các ID selected:", ", ".join(selected_ids))
    return selected


def save_article(driver) -> None:
    print("[3] Lưu bài...")

    button = WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: next(
            (
                item
                for item in d.find_elements(
                    By.CSS_SELECTOR,
                    'button.btn.btn-sm.green-jungle[type="submit"], '
                    'button.green-jungle[type="submit"]',
                )
                if item.is_displayed() and item.is_enabled()
            ),
            None,
        )
    )

    old_url = driver.current_url
    click_element(driver, button)
    time.sleep(WAIT_AFTER_SAVE)

    # Không bắt buộc URL đổi vì một số CMS lưu bằng POST rồi quay lại cùng URL.
    WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )

    print(f"    Đã click lưu. URL trước: {old_url}")
    print(f"    URL sau  : {driver.current_url}")


# Core functions are embedded for V2.2 standalone mode.


# ============================================================
# V2.2 ORCHESTRATOR (STANDALONE)
# ============================================================
# -*- coding: utf-8 -*-
"""
V2.3_dieu_phoi_bai_viet_lien_quan.py

ĐIỀU PHỐI BÀI VIẾT LIÊN QUAN ĐA LUỒNG V2.3 - BẢN SOLO
- Áp dụng toàn bộ cơ chế điều phối nâng cao từ v2.8 (Đăng bài):
  1. Khoá đồng bộ Đăng nhập tập trung (LOGIN_LOCK).
  2. Khoá Lưu theo Domain (domain_save_locks).
  3. Phân bổ bài thông minh ưu tiên khác Domain (pop_next_domain_safe_task).
  4. Quản lý sự cố WAF / HTTP 406 (tạm hoãn domain lỗi, worker chạy tiếp).
  5. Bật/Tắt ẩn hiện Edge động từ giao diện GUI (browser_visible_event).
  6. Ghi Excel COM an toàn với Retry Exponential Backoff khi Excel bận (-2146777998).
  7. Ghi Log độc lập theo từng Worker (TeeOutput & save_worker_error_log).
  8. Giao diện GUI Tkinter theo dõi tiến độ chi tiết từng Worker.
"""


from datetime import datetime
import io
import multiprocessing as mp
import os
import queue
import sys
import threading
import time
import traceback
from pathlib import Path
from typing import Any, Callable

from selenium import webdriver
from selenium.webdriver.edge.options import Options

CURRENT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = Path(os.environ.get("HOTKEYVIP_RUNTIME_ROOT", r"D:\CodexProjects\Hotkeyvip")).resolve()

PROFILE_ROOT = (
    PROJECT_ROOT
    / "02_viet_bai"
    / "du_lieu_3_workers"
    / "profiles"
)
WORKER_LOG_ROOT = (
    PROJECT_ROOT
    / "02_viet_bai"
    / "du_lieu_3_workers"
    / "log_lien_quan"
)

VERSION = "07_bai_viet_lien_quan (engine V2.3)"
EXCEL_BUSY_HRESULT = -2146777998  # 0x800AC472: Excel temporarily rejects COM calls.




# ============================================================
# TEE OUTPUT & LOGGING
# ============================================================

class TeeOutput:
    """Ghi đồng thời ra console và file log độc lập cho từng Worker."""

    def __init__(self, console, log_file) -> None:
        self.console = console
        self.log_file = log_file
        self._lock = threading.Lock()

    def write(self, text) -> int:
        value = str(text)
        with self._lock:
            if self.console is not None:
                try:
                    self.console.write(value)
                    self.console.flush()
                except Exception:
                    pass
            self.log_file.write(value)
            self.log_file.flush()
        return len(value)

    def flush(self) -> None:
        with self._lock:
            if self.console is not None:
                try:
                    self.console.flush()
                except Exception:
                    pass
            self.log_file.flush()

    def isatty(self) -> bool:
        if self.console is None:
            return False
        return bool(getattr(self.console, "isatty", lambda: False)())


def save_worker_error_log(worker_id: int, row: int, content: str) -> str:
    WORKER_LOG_ROOT.mkdir(parents=True, exist_ok=True)
    path = WORKER_LOG_ROOT / (
        f"lien_quan_worker{worker_id}_row{row}_"
        f"{datetime.now():%Y%m%d_%H%M%S}.log"
    )
    path.write_text(content, encoding="utf-8-sig")
    return str(path)


# ============================================================
# EXCEL WRITER QUEUE (V2.8 MECHANISM)
# ============================================================

class ExcelWriterQueue:
    """Đơn luồng sở hữu COM Excel; tự động Retry Exponential Backoff khi Excel bận."""

    def __init__(self) -> None:
        self.commands: queue.Queue[tuple[str, Callable[[], None] | None]] = queue.Queue()
        self.ready = threading.Event()
        self.failed = threading.Event()
        self.error: BaseException | None = None
        self.thread = threading.Thread(target=self._run, name="RelatedExcelWriter", daemon=True)

    @staticmethod
    def _is_excel_busy(exc: BaseException) -> bool:
        text = str(exc).casefold()
        return (
            str(EXCEL_BUSY_HRESULT) in text
            or "800ac472" in text
            or "rejected by callee" in text
            or "excel is busy" in text
        )

    @staticmethod
    def _retry_delay(attempt: int) -> float:
        return min(5.0, 0.5 * (2 ** min(attempt - 1, 4)))

    def start(self) -> None:
        self.thread.start()
        if not self.ready.wait(timeout=15):
            self.raise_if_failed()
            raise RuntimeError("Excel Writer không khởi động được trong 15 giây.")
        self.raise_if_failed()

    def submit(self, label: str, operation: Callable[[], None]) -> None:
        self.raise_if_failed()
        self.commands.put((label, operation))

    def raise_if_failed(self) -> None:
        if self.failed.is_set():
            raise RuntimeError("Excel Writer đã dừng do lỗi nghiêm trọng.") from self.error

    def drain_and_stop(self) -> None:
        while self.commands.unfinished_tasks:
            self.raise_if_failed()
            time.sleep(0.1)
        self.commands.put(("STOP", None))
        self.thread.join(timeout=20)
        if self.thread.is_alive():
            raise RuntimeError("Excel Writer không dừng được trong 20 giây.")
        self.raise_if_failed()

    def _run_operation(self, label: str, operation: Callable[[], None]) -> None:
        attempt = 0
        while True:
            try:
                operation()
                if attempt:
                    print(f"[EXCEL OK] {label} đã phục hồi sau {attempt} lần thử lại.")
                return
            except BaseException as exc:
                if not self._is_excel_busy(exc):
                    raise
                attempt += 1
                delay = self._retry_delay(attempt)
                print(f"[EXCEL BẬN] {label}; thử lại lần {attempt} sau {delay:.1f}s: {exc!r}")
                time.sleep(delay)

    def _run(self) -> None:
        pythoncom = None
        try:
            try:
                import pythoncom as pythoncom_module
                pythoncom = pythoncom_module
                pythoncom.CoInitialize()
            except ImportError:
                pass
            _app, workbook, _opened = connect_excel()
            sheet = workbook.Worksheets(SHEET_NAME)
            self.ready.set()
            while True:
                label, operation = self.commands.get()
                try:
                    if operation is None:
                        return
                    self._run_operation(label, operation)
                finally:
                    self.commands.task_done()
        except BaseException as exc:
            self.error = exc
            self.failed.set()
            self.ready.set()
            print(f"[EXCEL FATAL] {exc!r}")
        finally:
            if pythoncom is not None:
                pythoncom.CoUninitialize()


# ============================================================
# HELPER CẤU HÌNH & TRÌNH DUYỆT EDGE
# ============================================================

def configure_utf8() -> None:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")


def profile_path(worker_id: int) -> Path:
    return PROFILE_ROOT / f"worker_{worker_id}"


def choose_worker_count() -> int:
    return 5


def validate_profiles(worker_count: int) -> None:
    for worker_id in range(1, worker_count + 1):
        profile = profile_path(worker_id)
        if not (profile / "Local State").is_file():
            raise RuntimeError(
                f"Profile worker {worker_id} không hợp lệ hoặc chưa được tạo:\n{profile}"
            )


def set_edge_window_visible(driver, visible: bool) -> None:
    """Bật/tắt ẩn hiện cửa sổ Edge Selenium."""
    try:
        if visible:
            driver.maximize_window()
            driver.execute_script("window.focus();")
        else:
            driver.set_window_position(-32000, -32000)
    except Exception:
        pass


def create_worker_driver(worker_id: int):
    profile = profile_path(worker_id)
    profile.mkdir(parents=True, exist_ok=True)
    options = Options()
    options.add_argument(f"--user-data-dir={profile}")
    options.add_argument("--disable-notifications")
    options.add_argument("--window-position=-32000,-32000")
    options.add_argument("--window-size=1400,1000")
    options.add_experimental_option("detach", True)

    driver = webdriver.Edge(options=options)
    driver.get("about:blank")
    wait_document_ready(driver)
    return driver


# ============================================================
# CỔNG KIỂM TRA & ĐĂNG NHẬP ĐỒNG BỘ (LOGIN_LOCK)
# ============================================================

def ensure_post_page_ready_with_lock(driver, edit_url: str, login_lock: Any) -> None:
    """Đăng nhập đồng bộ: chỉ 1 worker thực hiện UI đăng nhập tại một thời điểm khi gặp form login."""
    driver.switch_to.default_content()
    wait_document_ready(driver)

    # V2.3: is_post_edit_page/is_login_page là hàm nội bộ của bản solo.
    # Không dùng biến `related` (không tồn tại trong file standalone).
    # Nếu bài đã ở đúng form chỉnh sửa, không cần xin khóa login.
    if is_post_edit_page(driver):
        return

    # Nếu phát hiện trang đăng nhập, dùng LOGIN_LOCK để xếp hàng
    if is_login_page(driver):
        print("    [LOGIN LOCK] Phát hiện trang đăng nhập → chờ khóa Đăng nhập...")
        with login_lock:
            if is_post_edit_page(driver):
                return
            ensure_post_page_ready(driver, edit_url)
        return

    # Gọi mặc định nếu lõi V1.0 tự xử lý
    ensure_post_page_ready(driver, edit_url)


# ============================================================
# PHÂN BỔ BÀI THÔNG MINH THEO DOMAIN (pop_next_domain_safe_task)
# ============================================================

def pop_next_domain_safe_task(
    pending_targets: list[dict[str, Any]],
    active_domains: set[str],
    deferred_406_domains: set[str],
    deferred_until: dict[str, float],
) -> dict[str, Any] | None:
    """Ưu tiên bài ở Domain khác để tránh nhiều worker cùng truy cập 1 website CMS."""
    # 1. Tìm bài thuộc domain hoàn toàn trống
    for index, target in enumerate(pending_targets):
        domain = str(target.get("domain", ""))
        if domain not in active_domains and domain not in deferred_406_domains:
            return pending_targets.pop(index)

    # 2. Không còn domain trống: cho phép chạy chung domain nếu domain đó không bị 406
    for index, target in enumerate(pending_targets):
        domain = str(target.get("domain", ""))
        if domain not in deferred_406_domains:
            return pending_targets.pop(index)

    # 3. Xét domain bị 406 khi hết toàn bộ bài khác
    if active_domains:
        return None

    now = time.time()
    for index, target in enumerate(pending_targets):
        domain = str(target.get("domain", ""))
        if domain in deferred_406_domains and now >= deferred_until.get(domain, 0):
            deferred_406_domains.discard(domain)
            return pending_targets.pop(index)

    return None


# ============================================================
# WORKER PROCESS MAIN
# ============================================================

def worker_main(
    worker_id: int,
    command_queue: mp.Queue,
    result_queue: mp.Queue,
    login_lock: Any,
    domain_save_locks: dict[str, Any],
    browser_visible_event: Any,
) -> None:
    configure_utf8()
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    memory_log = io.StringIO()
    sys.stdout = TeeOutput(original_stdout, memory_log)
    sys.stderr = TeeOutput(original_stderr, memory_log)

    driver = None
    had_error = False
    current_row = 0

    try:
        driver = create_worker_driver(worker_id)

        # Thread ngầm theo dõi sự kiện Bật/Tắt ẩn hiện Edge từ GUI
        def watch_browser_visibility() -> None:
            last_state = None
            while True:
                try:
                    desired = browser_visible_event.is_set()
                    if desired != last_state:
                        set_edge_window_visible(driver, desired)
                        last_state = desired
                    time.sleep(0.3)
                except Exception:
                    time.sleep(0.5)

        threading.Thread(target=watch_browser_visibility, daemon=True).start()

        result_queue.put({"type": "ready", "worker_id": worker_id})

        while True:
            command = command_queue.get()
            if command is None:
                break

            current_row = int(command["row"])
            post_id = str(command["post_id"])
            category = str(command.get("category", ""))
            edit_url = str(command["edit_url"])
            domain = str(command.get("domain", ""))
            started = time.time()

            result_queue.put({
                "type": "progress",
                "worker_id": worker_id,
                "row": current_row,
                "message": f"Đang mở ID {post_id} ({domain})",
            })

            try:
                driver.get(edit_url)
                wait_document_ready(driver)

                # 1. Đăng nhập tập trung dùng LOGIN_LOCK
                ensure_post_page_ready_with_lock(driver, edit_url, login_lock)

                result_queue.put({
                    "type": "progress",
                    "worker_id": worker_id,
                    "row": current_row,
                    "message": "Đang chọn bài liên quan",
                })

                # 2. Chọn bài liên quan
                selected_count = add_related_articles(driver, post_id, category)
                if selected_count <= 0:
                    raise RuntimeError("Không chọn được bài viết liên quan nào; không lưu bài.")

                # 3. Bấm Lưu dùng domain_save_locks (khóa theo từng Domain)
                domain_lock = domain_save_locks.get(domain)
                if domain_lock is not None:
                    with domain_lock:
                        save_article(driver)
                else:
                    save_article(driver)

                result_queue.put({
                    "type": "done",
                    "worker_id": worker_id,
                    "row": current_row,
                    "domain": domain,
                    "edit_url": edit_url,
                    "selected_count": selected_count,
                    "elapsed": round(time.time() - started, 1),
                })

            except Exception as exc:
                exc_text = str(exc).casefold()
                is_406 = "406" in exc_text or "not acceptable" in exc_text

                if is_406:
                    log_path = save_worker_error_log(worker_id, current_row, memory_log.getvalue())
                    result_queue.put({
                        "type": "blocked_406",
                        "worker_id": worker_id,
                        "row": current_row,
                        "domain": domain,
                        "error": str(exc),
                        "log_path": log_path,
                    })
                    # Reset driver cho worker để xóa trang bị block
                    try:
                        driver.quit()
                    except Exception:
                        pass
                    driver = create_worker_driver(worker_id)
                    result_queue.put({"type": "ready", "worker_id": worker_id})
                    continue

                had_error = True
                trace_text = traceback.format_exc()
                log_path = save_worker_error_log(worker_id, current_row, memory_log.getvalue())
                result_queue.put({
                    "type": "error",
                    "worker_id": worker_id,
                    "row": current_row,
                    "domain": domain,
                    "edit_url": edit_url,
                    "error": f"{type(exc).__name__}: {exc}",
                    "traceback": trace_text,
                    "log_path": log_path,
                })
                # Bật Edge lỗi lên cho người dùng kiểm tra
                try:
                    set_edge_window_visible(driver, True)
                except Exception:
                    pass
                break

    except Exception as exc:
        had_error = True
        log_path = save_worker_error_log(worker_id, current_row, memory_log.getvalue())
        result_queue.put({
            "type": "fatal",
            "worker_id": worker_id,
            "row": current_row,
            "error": f"{type(exc).__name__}: {exc}",
            "traceback": traceback.format_exc(),
            "log_path": log_path,
        })
    finally:
        if driver is not None and not had_error:
            try:
                driver.quit()
            except Exception:
                pass
        sys.stdout = original_stdout
        sys.stderr = original_stderr


# ============================================================
# GUI TIẾN ĐỘ TKINTER (MULTIPROGRESSWINDOW)
# ============================================================

class MultiProgressWindow:
    """Giao diện theo dõi tiến độ chi tiết từng Worker với nút Bật/Tắt Trình Duyệt."""

    def __init__(self, worker_count: int, total: int) -> None:
        self.worker_count = worker_count
        self.total = total
        self.stop_requested = threading.Event()
        self.browser_visible_requested = threading.Event()
        self.messages: queue.Queue[dict[str, Any]] = queue.Queue()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self) -> None:
        import tkinter as tk
        from tkinter import ttk

        root = tk.Tk()
        root.title(f"Tiến độ bài viết liên quan V2.1 — {self.worker_count} luồng")
        root.geometry(f"{max(640, 210 * self.worker_count)}x340+25+70")
        root.resizable(False, False)
        root.attributes("-topmost", True)

        tk.Label(
            root,
            text=f"Đang chạy {self.worker_count} luồng — tổng cộng {self.total} bài liên quan",
            font=("Segoe UI", 12, "bold"),
        ).pack(pady=(12, 6))

        worker_frame = tk.Frame(root)
        worker_frame.pack(fill="x", padx=10, pady=4)
        worker_labels: dict[int, tuple[Any, Any]] = {}

        for worker_id in range(1, self.worker_count + 1):
            group = tk.LabelFrame(
                worker_frame,
                text=f"Worker {worker_id}",
                padx=8,
                pady=6,
            )
            group.pack(side="left", fill="both", expand=True, padx=4)
            row_label = tk.Label(group, text="Dòng: —", font=("Segoe UI", 10, "bold"))
            row_label.pack()
            step_label = tk.Label(
                group,
                text="Đang khởi động Edge...",
                width=22,
                height=3,
                wraplength=165,
                justify="center",
            )
            step_label.pack()
            worker_labels[worker_id] = (row_label, step_label)

        progress_bar = ttk.Progressbar(
            root,
            maximum=max(self.total, 1),
            length=580,
            mode="determinate",
        )
        progress_bar.pack(pady=(10, 4))

        total_label = tk.Label(root, text=f"Đã hoàn thành: 0/{self.total}")
        total_label.pack()

        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=(8, 0))

        def toggle_browsers() -> None:
            if self.browser_visible_requested.is_set():
                self.browser_visible_requested.clear()
                browser_btn.config(text="Hiện trình duyệt")
            else:
                self.browser_visible_requested.set()
                browser_btn.config(text="Ẩn trình duyệt")

        def request_stop() -> None:
            self.stop_requested.set()
            stop_btn.config(text="Đang chờ bài hiện tại xong...", state="disabled")
            total_label.config(text="Đã yêu cầu dừng an toàn — không giao bài mới")

        browser_btn = tk.Button(btn_frame, text="Hiện trình duyệt", width=18, command=toggle_browsers)
        browser_btn.pack(side="left", padx=5)

        stop_btn = tk.Button(btn_frame, text="Dừng an toàn", width=28, command=request_stop)
        stop_btn.pack(side="left", padx=5)

        root.protocol("WM_DELETE_WINDOW", root.withdraw)

        def poll() -> None:
            try:
                while True:
                    msg = self.messages.get_nowait()
                    kind = msg.get("type")
                    worker_id = int(msg.get("worker_id", 0))

                    if kind == "complete_count":
                        count = int(msg.get("count", 0))
                        progress_bar["value"] = count
                        total_label.config(text=f"Đã hoàn thành: {count}/{self.total}")

                    elif kind == "finished":
                        total_label.config(text=str(msg.get("message")))
                        stop_btn.config(state="disabled")

                    elif worker_id in worker_labels:
                        row_lbl, step_lbl = worker_labels[worker_id]
                        row = msg.get("row")
                        if row:
                            row_lbl.config(text=f"Dòng: {row}")
                        step_lbl.config(text=str(msg.get("step") or kind))
            except queue.Empty:
                pass
            root.after(150, poll)

        poll()
        root.mainloop()

    def update_worker(self, worker_id: int, row: int | None, step: str) -> None:
        self.messages.put({"type": "worker", "worker_id": worker_id, "row": row, "step": step})

    def update_count(self, count: int) -> None:
        self.messages.put({"type": "complete_count", "count": count})

    def finish(self, message: str) -> None:
        self.messages.put({"type": "finished", "message": message})


# ============================================================
# MAIN ORCHESTRATOR
# ============================================================

def main() -> int:
    configure_utf8()
    mp.freeze_support()
    worker_count = choose_worker_count()
    validate_profiles(worker_count)

    _app, workbook, _opened = connect_excel()
    sheet = workbook.Worksheets(SHEET_NAME)
    website_url_map = get_website_url_map(workbook)
    targets = get_target_rows(sheet, website_url_map)

    if not targets:
        print("Không có bài viết liên quan nào cần chạy.")
        return 0

    # Rút danh sách unique domains để khởi tạo Locks
    unique_domains = {str(item.get("domain", "")) for item in targets if item.get("domain")}

    print("=" * 72)
    print(f"PHIÊN BẢN: {VERSION}")
    print(f"ĐÃ CHỐT {len(targets)} BÀI LIÊN QUAN | CHẠY {worker_count} WORKER")
    print(f"Profile Root: {PROFILE_ROOT}")
    print("=" * 72)

    progress = MultiProgressWindow(worker_count, len(targets))

    context = mp.get_context("spawn")
    result_queue = context.Queue()
    login_lock = context.Lock()
    browser_visible_event = context.Event()
    domain_save_locks = {domain: context.Lock() for domain in unique_domains}

    command_queues: dict[int, mp.Queue] = {}
    workers: dict[int, mp.Process] = {}

    for worker_id in range(1, worker_count + 1):
        command_queue = context.Queue(maxsize=1)
        process = context.Process(
            target=worker_main,
            args=(
                worker_id,
                command_queue,
                result_queue,
                login_lock,
                domain_save_locks,
                browser_visible_event,
            ),
            name=f"RelatedWorker-{worker_id}",
        )
        process.start()
        command_queues[worker_id] = command_queue
        workers[worker_id] = process

    free_workers: set[int] = set()
    active_rows: dict[int, int] = {}
    active_tasks: dict[int, dict[str, Any]] = {}
    active_domains: set[str] = set()

    pending_targets = list(targets)
    deferred_406_domains: set[str] = set()
    deferred_until: dict[str, float] = {}

    completed = 0
    stopping = False
    user_stopped = False
    had_error = False

    try:
        while True:
            # Đồng bộ ẩn/hiện trình duyệt từ nút bấm GUI
            if progress.browser_visible_requested.is_set():
                browser_visible_event.set()
            else:
                browser_visible_event.clear()

            if progress.stop_requested.is_set() and not stopping:
                stopping = True
                user_stopped = True
                print("\n[DỪNG AN TOÀN] Không giao bài mới; chờ các bài hiện tại hoàn tất.")
                progress.finish(f"Đang dừng an toàn — đã xong {completed}/{len(targets)}")

            try:
                message = result_queue.get(timeout=0.25)
            except queue.Empty:
                message = None

            if message:
                kind = message.get("type")
                worker_id = int(message.get("worker_id", 0))

                if kind == "ready":
                    free_workers.add(worker_id)
                    print(f"Worker {worker_id}: sẵn sàng.")
                    progress.update_worker(worker_id, None, "Sẵn sàng nhận bài")

                elif kind == "progress":
                    row = message.get("row")
                    msg = str(message.get("message"))
                    print(f"Worker {worker_id} | dòng {row}: {msg}")
                    progress.update_worker(worker_id, row, msg)

                elif kind == "done":
                    row = int(message["row"])
                    active_rows.pop(worker_id, None)
                    task = active_tasks.pop(worker_id, None)
                    if task:
                        active_domains.discard(str(task.get("domain", "")))

                    target = next(item for item in targets if int(item["row"]) == row)

                    # Ghi Excel ngay trên luồng chính, cùng luồng đã kết nối COM Excel.
                    # Không truyền đối tượng sheet/workbook sang luồng phụ.
                    cell = sheet.Cells(row, int(target["col_bvlq"]))
                    if hasattr(cell, "setValue"):
                        cell.setValue(str(message["edit_url"]))
                    else:
                        cell.Value = str(message["edit_url"])
                    workbook.Save()

                    completed += 1
                    free_workers.add(worker_id)
                    progress.update_worker(worker_id, row, "Đã lưu thành công")
                    progress.update_count(completed)
                    print(
                        f"[OK] Worker {worker_id} xong dòng {row} | "
                        f"{message.get('selected_count')} bài | {message.get('elapsed')}s"
                    )

                elif kind == "blocked_406":
                    row = int(message["row"])
                    domain = str(message["domain"])
                    active_rows.pop(worker_id, None)
                    active_tasks.pop(worker_id, None)
                    active_domains.discard(domain)
                    deferred_406_domains.add(domain)
                    deferred_until[domain] = time.time() + 60.0  # Tạm hoãn domain 60s
                    print(f"\n[HTTP 406] Worker {worker_id} bị WAF chặn ở domain {domain}. Tạm hoãn domain 60s.")
                    progress.update_worker(worker_id, row, "Gặp 406 — tạm hoãn domain")

                elif kind in {"error", "fatal"}:
                    failed_row = active_rows.pop(worker_id, None)
                    failed_task = active_tasks.pop(worker_id, None)
                    if failed_task:
                        active_domains.discard(str(failed_task.get("domain", "")))
                    had_error = True
                    print("\n" + "!" * 72)
                    error_row = message.get("row") or failed_row or "?"
                    print(f"[LỖI WORKER {worker_id}] dòng {error_row}: {message.get('error')}")
                    print("Bỏ qua worker/bài lỗi; các worker còn lại vẫn tiếp tục chạy.")
                    print(f"File Log: {message.get('log_path', '')}")
                    print("!" * 72)
                    progress.update_worker(worker_id, error_row, "LỖI — bỏ qua, worker khác tiếp tục")

            if not stopping:
                for worker_id in sorted(list(free_workers)):
                    target = pop_next_domain_safe_task(
                        pending_targets,
                        active_domains,
                        deferred_406_domains,
                        deferred_until,
                    )
                    if target is None:
                        break

                    free_workers.remove(worker_id)
                    row = int(target["row"])
                    domain = str(target.get("domain", ""))
                    active_rows[worker_id] = row
                    active_tasks[worker_id] = target
                    if domain:
                        active_domains.add(domain)

                    command_queues[worker_id].put({
                        "row": row,
                        "post_id": str(target["post_id"]),
                        "category": str(target.get("category", "")),
                        "edit_url": str(target["edit_url"]),
                        "domain": domain,
                    })

                    print(f"[GIAO] Worker {worker_id} <- dòng {row} | ID {target['post_id']} | Domain: {domain}")
                    progress.update_worker(worker_id, row, f"Đã nhận ID {target['post_id']}")

            all_dispatched = not pending_targets
            if (all_dispatched or stopping) and not active_rows:
                break

            if all(not process.is_alive() for process in workers.values()) and not active_rows:
                break

    except KeyboardInterrupt:
        stopping = True
        print("\nĐã nhận Ctrl+C: ngừng giao bài mới.")
    finally:
        for command_queue in command_queues.values():
            try:
                command_queue.put_nowait(None)
            except Exception:
                pass

        for process in workers.values():
            process.join(timeout=8)

        for process in workers.values():
            if process.is_alive():
                process.terminate()
                process.join(timeout=3)

        final_msg = f"Hoàn thành: {completed}/{len(targets)} bài."
        if had_error:
            final_msg += " (CÓ LỖI)"
        elif user_stopped:
            final_msg += " (ĐÃ DỪNG AN TOÀN)"
        else:
            final_msg += " (THÀNH CÔNG)"

        progress.finish(final_msg)

    print("=" * 72)
    print(f"Kết thúc: hoàn thành {completed}/{len(targets)} bài.")
    print("=" * 72)
    return 1 if had_error else 0


if __name__ == "__main__":
    mp.freeze_support()
    try:
        exit_code = main()
    finally:
        close_excel_if_needed(save=True)
    raise SystemExit(exit_code)


