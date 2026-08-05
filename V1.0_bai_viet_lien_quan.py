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

EXCEL_PATH = EXCEL_FILE
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


def connect_excel() -> tuple[Any, Any, bool]:
    global _EXCEL_APP, _EXCEL_WB, _EXCEL_OPENED_BY_SCRIPT

    # Chỉ dùng đúng Excel/workbook người dùng đang mở.
    # Nếu không kết nối được thì báo lỗi và dừng; tuyệt đối không mở Excel ẩn.
    try:
        app = win32.GetActiveObject("Excel.Application")
    except Exception as exc:
        raise RuntimeError(
            "Không kết nối được với Excel đang mở. "
            "Hãy mở Excel và file hotkeyvip_test.xlsm rồi chạy lại."
        ) from exc

    try:
        for index in range(1, app.Workbooks.Count + 1):
            wb = app.Workbooks(index)
            fullname = clean(wb.FullName)

            if _same_path(fullname, str(EXCEL_PATH)):
                print(f"Đã kết nối Excel đang mở: {wb.Name}")
                _EXCEL_APP = app
                _EXCEL_WB = wb
                _EXCEL_OPENED_BY_SCRIPT = False
                return app, wb, False
    except Exception as exc:
        raise RuntimeError(
            "Đã thấy Excel nhưng không đọc được danh sách workbook đang mở."
        ) from exc

    raise RuntimeError(
        "Excel đang mở nhưng chưa mở đúng file hotkeyvip_test.xlsm.\n"
        f"Hãy mở file này rồi chạy lại:\n{EXCEL_PATH}"
    )


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


# ============================================================
# MAIN: CHẠY HẾT TẤT CẢ BÀI
# ============================================================

def main() -> None:
    driver = None
    save_excel = False
    processed_count = 0
    error_count = 0
    progress_window: RelatedProgressWindow | None = None
    stopped_early = False

    try:
        _app, wb, _opened_by_script = connect_excel()
        ws = wb.Worksheets(SHEET_NAME)

        # Quét Excel đúng một lần rồi chạy cố định danh sách đã chốt.
        website_url_map = get_website_url_map(wb)
        targets = get_target_rows(ws, website_url_map)
        target_count = len(targets)
        print(f"Số bài cần chạy đã chốt: {target_count}")

        if target_count == 0:
            import tkinter as tk
            from tkinter import messagebox

            info_root = tk.Tk()
            info_root.withdraw()
            info_root.attributes("-topmost", True)
            messagebox.showinfo(
                "Bài viết liên quan",
                "Không có bài viết liên quan nào cần chạy.",
                parent=info_root,
            )
            info_root.destroy()
            return

        progress_window = RelatedProgressWindow(target_count)

        # Không tìm lại Excel sau mỗi bài và không quét tìm bài thứ target_count + 1.
        for target_index, target in enumerate(targets, start=1):
            if progress_window.should_stop():
                stopped_early = True
                print("Đã nhận lệnh dừng trước khi bắt đầu bài tiếp theo.")
                break
            row = int(target["row"])
            domain = target["domain"]
            post_id = target["post_id"]
            category = target["category"]
            col_bvlq = int(target["col_bvlq"])
            edit_url = target["edit_url"]
            progress_window.update(
                processed_count,
                f"Đang chạy bài {target_index}/{target_count} - dòng {row}",
            )

            print("\n" + "=" * 70)
            print(f"ĐANG XỬ LÝ BÀI {target_index}/{target_count}")
            print(f"Dòng Excel : {row}")
            print(f"Tên miền   : {domain}")
            print(f"ID         : {post_id}")
            print(f"Danh mục   : {category}")
            print(f"URL        : {edit_url}")
            print("=" * 70)

            try:
                if driver is None:
                    driver = open_edge(edit_url)
                else:
                    print("Mở URL bài tiếp theo trên Edge hiện tại...")
                    driver.get(edit_url)

                    wait_document_ready(driver)
                    ensure_post_page_ready(driver, edit_url)

                selected_count = add_related_articles(
                    driver,
                    post_id,
                    category,
                )

                if selected_count <= 0:
                    raise RuntimeError(
                        "Không chọn được bài viết liên quan nào; không lưu bài."
                    )

                save_article(driver)

                if WRITE_STATUS_TO_EXCEL:
                    ws.Cells(row, col_bvlq).Value = edit_url
                    wb.Save()
                    save_excel = True
                    print(f"Đã ghi trạng thái vào Excel, dòng {row}.")

                processed_count += 1
                progress_window.update(
                    processed_count,
                    f"Đã lưu xong bài {target_index}/{target_count}",
                )
                print(
                    f"HOÀN THÀNH DÒNG {row}. "
                    f"Tổng đã chạy: {processed_count} bài."
                )

                # Chỉ dừng sau khi web đã lưu và Excel đã được ghi xong.
                if progress_window.should_stop():
                    stopped_early = processed_count < target_count
                    print(
                        f"ĐÃ DỪNG AN TOÀN SAU BÀI "
                        f"{processed_count}/{target_count}."
                    )
                    break

            except Exception as row_error:
                error_count += 1
                print("\n" + "!" * 70)
                print(f"LỖI DÒNG {row}: {row_error}")
                print(
                    "Không ghi trạng thái vào Excel. "
                    "Chương trình dừng để tránh lặp vô hạn đúng dòng lỗi."
                )
                print("!" * 70)

                if driver is not None:
                    try:
                        show_edge_window()
                        print("Đã hiện Edge để kiểm tra lỗi.")
                        driver = None
                    except Exception:
                        pass

                raise

        progress_window.notify_finished(processed_count, stopped=stopped_early)

        print("=" * 70)
        if stopped_early:
            print(f"ĐÃ DỪNG AN TOÀN: {processed_count}/{target_count} bài.")
        else:
            print(f"ĐÃ CHẠY ĐỦ VÒNG ĐÃ CHỐT: {processed_count}/{target_count} bài.")
        print(f"Số bài lỗi: {error_count} bài.")
        print("=" * 70)

        if driver is not None and SHOW_EDGE_AFTER_DONE:
            show_edge_window()
            print("Đã chạy hết. Edge được giữ lại để bạn kiểm tra.")
            driver = None

    except Exception as exc:
        print("\nLỖI CHƯƠNG TRÌNH:", exc)
        try:
            import tkinter as tk
            from tkinter import messagebox

            error_root = tk.Tk()
            error_root.withdraw()
            error_root.attributes("-topmost", True)
            messagebox.showerror(
                "Bài viết liên quan gặp lỗi",
                str(exc),
                parent=error_root,
            )
            error_root.destroy()
        except Exception:
            pass
        raise

    finally:
        if progress_window is not None:
            progress_window.close()

        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass

        close_excel_if_needed(save=save_excel)


if __name__ == "__main__":
    main()
