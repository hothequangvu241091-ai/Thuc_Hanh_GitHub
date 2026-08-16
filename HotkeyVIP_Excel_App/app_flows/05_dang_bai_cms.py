# -*- coding: utf-8 -*-
"""
test_nhap_du_lieu_v1.py

BẢN TEST WORD CHẠY NỀN:
- Không gọi file mở Word cũ; Word được mở ẩn bằng COM.
- Mở URL CMS bằng Edge Selenium dùng profile cố định, sau đó ẩn cửa sổ.
- Điền 5 ô đầu bằng Selenium.
- Đưa Edge lên trước và focus CKEditor.
- Gọi F7 để AHK dán nội dung Word giữ định dạng.
- Chỉ nạp nội dung chính trước FAQ vào CKEditor.
- Chọn Danh mục và Tác giả qua giao diện Select2 bằng Selenium.
- Tự chạy, không hỏi xác nhận đăng nhập.
- Chạy xong bước gom HTML thì hiện Edge lại để kiểm tra; chưa lưu bài.

Yêu cầu:
    pip install selenium pywin32
"""

from __future__ import annotations

import configparser
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import html as html_std
import io
import json
import os
import random
import re
import queue
import subprocess
import threading
import traceback
import unicodedata
import sys
import time
import winsound
from pathlib import Path
from typing import Any

import win32com.client as win32
from docx import Document
import win32clipboard
import win32con
import win32gui
from selenium import webdriver
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait

# Source files live in D:\Dang_bai.  Existing Excel/data assets remain in the
# Hotkeyvip runtime folder until they are migrated deliberately.
PROJECT_ROOT = Path(
    os.environ.get("HOTKEYVIP_RUNTIME_ROOT", r"D:\CodexProjects\Hotkeyvip")
).resolve()
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))

from hotkeyvip_config import (
    EDGE_USER_DATA_DIR,
    EXCEL_FILE,
    LOGIN_INI,
    OUTPUT_DIR,
    PUBLISH_HEADERS,
    SHEET_AUTHOR as PROJECT_SHEET_AUTHOR,
    SHEET_CATEGORY_ID,
    SHEET_PUBLISH,
    SHEET_WEBSITE,
    ensure_runtime_directories,
)


# ============================================================
# CẤU HÌNH
# ============================================================

MO_WORD_PY = PROJECT_ROOT / "03_dang_bai" / "phu_tro" / "mo_word_tu_excel.py"
EXCEL_PATH = Path(
    os.environ.get("HOTKEYVIP_SELECTED_EXCEL", str(EXCEL_FILE))
).resolve()

# Thư mục mặc định chứa file Word nếu Excel chỉ lưu tên file.
WORD_BASE_DIR = OUTPUT_DIR

# Word chạy nền, không hiện cửa sổ.
_WORD_APP = None
_WORD_DOC = None
_WORD_PATH: Path | None = None
_WORD_ROW: int | None = None
_FAQ_RUN_NOTE = ""

# Handle cửa sổ Edge do Selenium mở.
_EDGE_HWND: int | None = None
_EDGE_IS_HIDDEN = False

# Bản điều phối đa luồng có thể gắn multiprocessing.Lock vào đây để
# các worker lần lượt mở CMS/đăng nhập. Chạy lẻ để None nên không đổi flow.
CMS_ENTRY_LOCK = None
# V1.2: Locks are attached by the multi-worker coordinator only when needed.
# LOGIN_LOCK is used only after a worker has actually detected a login form.
# SAVE_DOMAIN_LOCK is set to the lock of the current task's domain.
LOGIN_LOCK = None
SAVE_DOMAIN_LOCK = None
WORD_CLIPBOARD_LOCK = None  # V2.5 locks only Copy + Clipboard read.

SHEET_POST = SHEET_PUBLISH
SHEET_DOMAIN = SHEET_WEBSITE
SHEET_AUTHOR = PROJECT_SHEET_AUTHOR

HEADER_DOMAIN = "Tên miền"
HEADER_CATEGORY = "Danh mục"

WAIT_PAGE = 30
WAIT_AFTER_FILL = 0.5
WAIT_AFTER_SET_DATA = 1
WAIT_UPLOAD = 45


class Http406Error(RuntimeError):
    """Website/WAF từ chối request bằng trang 406 Not Acceptable."""


class PreSaveValidationError(RuntimeError):
    """Dữ liệu trên form chưa đủ an toàn để bấm Lưu bài hiện tại."""

# ============================================================
# QUY TẮC ĐƯỜNG DẪN ẢNH
# - Lấy tên file Word được mở ẩn, bỏ phần mở rộng .doc/.docx.
# - Ảnh nằm trong D:\baivietlamviec\dangbai.
# - Ảnh đại diện: <tên Word> 1.<đuôi ảnh>
# - Ảnh bài viết: <tên Word> 2.<đuôi ảnh>
# - Tự dò các đuôi: png, jpg, jpeg, webp.
# ============================================================

IMAGE_DIR = OUTPUT_DIR

IMAGE_EXTENSIONS = (
    ".png",
    ".jpg",
    ".jpeg",
    ".webp",
)

# Dùng đúng dữ liệu trình duyệt giống file code_tong_v2_bao_hiem.
# Cookie và phiên đăng nhập được lưu tại đây.
EDGE_USER_DATA_DIR = Path(EDGE_USER_DATA_DIR)
LOGIN_INI = Path(LOGIN_INI)

# Mỗi lần chạy lưu toàn bộ nội dung CMD vào một file riêng.
RUN_LOG_DIR = PROJECT_ROOT / "06_du_lieu_chay" / "log_dang_bai"


class TeeOutput:
    """Ghi đồng thời ra CMD và file log, kể cả khi chạy bằng pythonw."""

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


def run_with_log() -> int:
    """Chỉ lưu toàn bộ nội dung CMD thành file khi chương trình gặp lỗi."""
    started_at = datetime.now()
    log_path = RUN_LOG_DIR / (
        f"dang_bai_{started_at:%Y%m%d_%H%M%S}_pid{os.getpid()}.log"
    )
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    memory_log = io.StringIO()
    exit_code = 1
    unexpected_error: BaseException | None = None

    sys.stdout = TeeOutput(original_stdout, memory_log)
    sys.stderr = TeeOutput(original_stderr, memory_log)
    try:
        print(f"Thời gian bắt đầu : {started_at:%d/%m/%Y %H:%M:%S}")
        exit_code = main()
    except BaseException as exc:
        unexpected_error = exc
        print("\nLỖI NGOÀI LUỒNG MAIN:")
        traceback.print_exc()
        exit_code = 1
    finally:
        finished_at = datetime.now()
        print(f"Thời gian kết thúc: {finished_at:%d/%m/%Y %H:%M:%S}")
        sys.stdout = original_stdout
        sys.stderr = original_stderr

    if exit_code != 0:
        RUN_LOG_DIR.mkdir(parents=True, exist_ok=True)
        log_path.write_text(memory_log.getvalue(), encoding="utf-8-sig")
        message = f"Đã lưu log lỗi tại: {log_path}"
        if original_stdout is not None:
            print(message, file=original_stdout, flush=True)

    if unexpected_error is not None:
        raise unexpected_error

    return exit_code



# ============================================================
# HÀM CHUNG
# ============================================================

def clean_text(value: Any) -> str:
    if value is None:
        return ""

    text = str(value)
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def excel_local_now_serial() -> float:
    """Return local wall-clock time as an Excel serial without COM timezone conversion."""
    local_now = datetime.now()
    excel_epoch = datetime(1899, 12, 30)
    return (local_now - excel_epoch).total_seconds() / 86400


def set_faq_run_note(message: str) -> None:
    global _FAQ_RUN_NOTE
    _FAQ_RUN_NOTE = clean_text(message)


def get_faq_run_note() -> str:
    return _FAQ_RUN_NOTE


def get_target_excel_workbook() -> tuple[Any, Any]:
    """Bám workbook đang mở; nếu không thấy thì kết nối theo đường dẫn."""
    excel = None
    try:
        excel = win32.GetActiveObject("Excel.Application")
    except Exception:
        pass

    if excel is not None:
        try:
            for index in range(1, excel.Workbooks.Count + 1):
                workbook = excel.Workbooks(index)
                workbook_path = Path(str(workbook.FullName)).resolve()
                if os.path.normcase(str(workbook_path)) == os.path.normcase(str(EXCEL_PATH)):
                    return excel, workbook
        except Exception:
            # Excel có thể vừa đóng/mở lại; thử bám theo file bên dưới.
            pass

    if not EXCEL_PATH.is_file():
        raise RuntimeError(f"Không tìm thấy file Excel:\n{EXCEL_PATH}")

    try:
        workbook = win32.GetObject(str(EXCEL_PATH))
        excel = workbook.Application
        excel.Visible = True
    except Exception as exc:
        raise RuntimeError(
            "Không bám được Excel đang mở và cũng không mở được file:\n"
            f"{EXCEL_PATH}"
        ) from exc

    if bool(workbook.ReadOnly):
        raise RuntimeError(
            "Excel đã kết nối nhưng workbook đang ở chế độ chỉ đọc. "
            "Có thể file đang bị mở trùng ở một tiến trình Excel khác; "
            "hãy đóng bản trùng để tránh xung đột lưu."
        )

    print(f"    Đã kết nối Excel theo đường dẫn: {EXCEL_PATH}")
    return excel, workbook




def safe_filename(title: str) -> str:
    return re.sub(r'[\\/:*?"<>|]', "_", title).strip()[:100]


def normalize_docx_name(name: str) -> str:
    value = os.path.splitext(name)[0]
    value = re.sub(r'[\\/:*?"<>|]', "_", value)
    value = re.sub(r"_+", "_", value)
    value = value.rstrip("_")
    return value.strip().lower()


def find_next_word_file_from_excel() -> tuple[int, str, Path]:
    """
    Quy tắc chọn và mở file Word:
    - Ưu tiên dòng có Trạng thái chứa "cần mở" hoặc bằng "Ok".
    - Nếu không có, lấy dòng đầu tiên có Trạng thái trống như cũ.
    - Nếu cột dn_file_word có đường dẫn, ưu tiên dùng đường dẫn đó.
    - Nếu dn_file_word trống, tìm/tạo file Word theo Tiêu đề như cũ.
    - Chỉ sau khi xác định được file Word hợp lệ mới đổi Trạng thái thành WORD.
    """
    try:
        excel, workbook = get_target_excel_workbook()
        sheet = workbook.Worksheets(SHEET_POST)
    except Exception as exc:
        raise RuntimeError(
            'Không kết nối được Excel hoặc không tìm thấy sheet "dangbai".'
        ) from exc

    title_col = find_column_by_header(sheet, PUBLISH_HEADERS["keyword"])
    status_col = find_column_by_header(sheet, PUBLISH_HEADERS["status"])
    word_path_col = find_column_by_header(sheet, PUBLISH_HEADERS["word_path"])

    target_row: int | None = None
    title = ""

    # Ưu tiên dòng có trạng thái chứa "cần mở" hoặc bằng "Ok".
    row = 2
    while True:
        title_cell = sheet.Cells(row, title_col).Value

        if title_cell is None or not str(title_cell).strip():
            break

        status_cell = sheet.Cells(row, status_col).Value
        status = "" if status_cell is None else str(status_cell).strip()

        status_key = status.casefold()
        if "cần mở" in status_key or status_key == "ok":
            target_row = row
            title = str(title_cell).strip()
            break

        row += 1

    # Nếu không có "cần mở"/"Ok", dùng dòng trạng thái trống như logic cũ.
    if target_row is None:
        row = 2

        while True:
            title_cell = sheet.Cells(row, title_col).Value

            if title_cell is None or not str(title_cell).strip():
                break

            status_cell = sheet.Cells(row, status_col).Value
            status = "" if status_cell is None else str(status_cell).strip()

            if status == "":
                target_row = row
                title = str(title_cell).strip()
                break

            row += 1

    if target_row is None:
        raise RuntimeError(
            'Không còn bài nào có Trạng thái "cần mở", "Ok" hoặc đang trống.'
        )

    WORD_BASE_DIR.mkdir(parents=True, exist_ok=True)

    word_path_cell = sheet.Cells(target_row, word_path_col).Value
    direct_word_path = (
        ""
        if word_path_cell is None
        else str(word_path_cell).strip().strip('"')
    )

    if direct_word_path:
        direct_word_path = os.path.expandvars(
            os.path.expanduser(direct_word_path)
        )
        found_path = Path(direct_word_path).resolve()

        if not found_path.is_file():
            raise RuntimeError(
                f"Dòng {target_row} có dn_file_word nhưng đường dẫn không tồn tại:\n"
                f"{found_path}"
            )

        if found_path.suffix.lower() not in {".doc", ".docx"}:
            raise RuntimeError(
                f"Dòng {target_row} có dn_file_word nhưng không phải file Word:\n"
                f"{found_path}"
            )

        print(
            f'    Ưu tiên mở Word từ cột "{PUBLISH_HEADERS["word_path"]}" '
            "của sheet DANG_BAI."
        )

    else:
        file_name = safe_filename(title) + ".docx"
        target_name = normalize_docx_name(file_name)
        found_path: Path | None = None

        for root, _dirs, files in os.walk(WORD_BASE_DIR):
            for filename in files:
                if not filename.lower().endswith(".docx"):
                    continue

                if normalize_docx_name(filename) == target_name:
                    found_path = Path(root) / filename
                    break

            if found_path is not None:
                break

        if found_path is None:
            found_path = WORD_BASE_DIR / file_name
            new_doc = Document()
            new_doc.add_heading(title, level=1)
            new_doc.save(str(found_path))

        found_path = found_path.resolve()

    sheet.Cells(target_row, status_col).Value = "WORD"
    workbook.Save()

    return target_row, title, found_path

# ============================================================
# LƯU Ý LỖI WIN32COM / GEN_PY
# ============================================================
# Nếu chương trình báo lỗi kiểu:
#   AttributeError:
#   module 'win32com.gen_py....' has no attribute 'CLSIDToClassMap'
#
# Thì nguyên nhân KHÔNG phải sai đường dẫn hoặc file Word hỏng.
# Đây là lỗi cache COM của pywin32 trong thư mục gen_py bị hỏng/lệch.
#
# Cách sửa nhanh trong PowerShell:
#
#   python -c "import shutil; import win32com.client.gencache as g; p=g.GetGeneratePath(); print('Xóa:', p); shutil.rmtree(p, ignore_errors=True)"
#
# Sau đó chạy:
#
#   python -c "import win32com.client.gencache as g; g.Rebuild(); print('Đã tạo lại cache COM')"
#
# Rồi chạy lại chương trình.
#
# Có thể xảy ra lại sau khi:
# - Office cập nhật
# - Python hoặc pywin32 cập nhật
# - Máy/program tắt đột ngột
# - Cache gen_py được tạo dở hoặc bị lỗi
#
# Từ khóa để hỏi ChatGPT:
# "win32com gen_py CLSIDToClassMap Word COM lỗi"
# ============================================================

def open_hidden_word_document() -> tuple[int, str, Path]:
    """
    Dùng nguyên điều kiện tìm bài của file mo_word_tu_excel.py,
    nhưng mở Word bằng tiến trình COM riêng ở chế độ ẩn.
    """
    global _WORD_APP, _WORD_DOC, _WORD_PATH, _WORD_ROW

    if _WORD_DOC is not None and _WORD_PATH is not None and _WORD_ROW is not None:
        return _WORD_ROW, _WORD_PATH.stem, _WORD_PATH

    current_row, title, word_path = find_next_word_file_from_excel()

    word = None
    try:
        word = win32.DispatchEx("Word.Application")
        word.Visible = False
        word.DisplayAlerts = 0
        doc = word.Documents.Open(
            str(word_path),
            ReadOnly=True,
            AddToRecentFiles=False,
            Visible=False,
        )
    except Exception as exc:
        error_text = repr(exc)

        if "CLSIDToClassMap" in error_text:
            extra_note = (
                "\n\nĐây có thể là lỗi cache win32com/gen_py. "
                "Hãy xóa gen_py và chạy gencache.Rebuild()."
            )
        else:
            extra_note = ""

        raise RuntimeError(
            f"Không mở được file Word:\n{word_path}\n\n"
            f"Lỗi thật: {error_text}"
            f"{extra_note}"
        ) from exc

    _WORD_APP = word
    _WORD_DOC = doc
    _WORD_PATH = word_path
    _WORD_ROW = current_row

    print(f"[1] Đã chọn dòng {current_row}: {title}")
    print("    Word chạy nền, không hiện cửa sổ.")
    print(f"    File: {word_path}")
    print(f"    Dòng Excel được truyền trực tiếp: {current_row}")

    return current_row, title, word_path

def get_hidden_word_document() -> Any:
    if _WORD_DOC is None:
        raise RuntimeError("Tài liệu Word chạy nền chưa được mở.")
    return _WORD_DOC


def close_hidden_word_document() -> None:
    """Đóng tài liệu và tiến trình Word nền do script tạo."""
    global _WORD_APP, _WORD_DOC, _WORD_PATH, _WORD_ROW

    doc = _WORD_DOC
    word = _WORD_APP

    _WORD_DOC = None
    _WORD_APP = None
    _WORD_PATH = None
    _WORD_ROW = None

    if doc is not None:
        try:
            doc.Close(False)
        except Exception:
            pass

    if word is not None:
        try:
            word.Quit()
        except Exception:
            pass


def get_open_word_file_info() -> dict[str, Any]:
    """Lấy thông tin file Word đang được script mở ẩn."""
    doc = get_hidden_word_document()
    full_name = clean_text(doc.FullName)

    if not full_name:
        raise RuntimeError("Không đọc được đường dẫn file Word chạy nền.")

    word_path = Path(full_name)
    return {
        "path": word_path,
        "stem": word_path.stem,
        "name": word_path.name,
    }

def get_publish_asset_path(row: int, header_key: str) -> Path | None:
    _excel, workbook = get_target_excel_workbook()
    sheet = workbook.Worksheets(SHEET_POST)
    col = find_column_by_header(sheet, PUBLISH_HEADERS[header_key])
    raw = clean_text(sheet.Cells(row, col).Value).strip('"')
    if not raw:
        return None
    path = Path(os.path.expandvars(os.path.expanduser(raw))).resolve()
    if not path.is_file():
        raise RuntimeError(
            f'Dòng {row}, cột "{PUBLISH_HEADERS[header_key]}" trỏ tới file không tồn tại:\n{path}'
        )
    print(
        f'    Ưu tiên file từ DANG_BAI, cột "{PUBLISH_HEADERS[header_key]}": {path}'
    )
    return path


def get_image_paths_from_open_word(current_row: int) -> dict[str, Any]:
    word_info = get_open_word_file_info()
    stem = word_info["stem"]

    featured = get_publish_asset_path(current_row, "image1_path")
    content = get_publish_asset_path(current_row, "image2_path")
    if featured is None:
        featured = find_numbered_image(stem, 1, word_info["path"].parent)
    if content is None:
        content = find_numbered_image(stem, 2, word_info["path"].parent)

    print("\nĐối chiếu tài nguyên đăng bài:")
    print(f"- File Word     : {word_info['path']}")
    print(f"- Tên gốc Word : {stem}")
    print(f"- Ảnh đại diện : {featured}")
    print(f"- Ảnh bài viết : {content}")

    return {
        "word": word_info,
        "featured": featured,
        "content": content,
    }


def normalize_domain(domain: Any) -> str:
    value = clean_text(domain).lower()
    value = value.replace("https://", "")
    value = value.replace("http://", "")
    value = value.replace("www.", "")
    return value.rstrip("/")


def process_f5_line(text: str) -> str:
    """
    Giữ logic F5 hiện tại:
    - Xóa '# ' hoặc '#' nếu không phải ## / ###.
    - Nếu phần trước ':' có dưới 4 dấu cách thì lấy phần sau ':'.
    """
    value = text.strip()

    if value.startswith("# "):
        value = value[2:].strip()
    elif value.startswith("#") and not value.startswith("##"):
        value = value[1:].strip()

    pos = value.find(":")

    if pos >= 0:
        before_colon = value[:pos].strip()

        if before_colon.count(" ") < 4:
            value = value[pos + 1 :].strip()

    return value




def find_column_by_header(ws: Any, header_name: str) -> int:
    last_col = int(
        ws.Cells(
            1,
            ws.Columns.Count,
        ).End(-4159).Column
    )

    for col in range(1, last_col + 1):
        value = ws.Cells(1, col).Value

        if (
            value
            and str(value).strip().lower()
            == header_name.lower()
        ):
            return col

    raise RuntimeError(
        f'Không tìm thấy cột "{header_name}"'
    )


def format_publish_status_cell(cell: Any, status: str) -> None:
    """Tô màu trạng thái để ĐÃ ĐĂNG, LỖI ID và LỖI ĐĂNG dễ phân biệt."""
    text = clean_text(status).casefold()
    cell.Font.Bold = True

    if "lỗi id" in text:
        # Cam: CMS có thể đã lưu bài nhưng chưa bắt được ID.
        cell.Interior.Color = 255 + 192 * 256
        cell.Font.Color = 128
    elif "lỗi" in text:
        # Đỏ: lỗi đăng thật.
        cell.Interior.Color = 192
        cell.Font.Color = 255 + 255 * 256 + 255 * 65536
    elif "đã đăng" in text:
        # Xanh nhạt: hoàn thành.
        cell.Interior.Color = 198 + 239 * 256 + 206 * 65536
        cell.Font.Color = 0 + 97 * 256
    else:
        cell.Interior.ColorIndex = -4142
        cell.Font.ColorIndex = -4105
        cell.Font.Bold = False


def write_compact_error_cell(cell: Any, error: Any) -> None:
    """
    Ghi lỗi mà không làm dòng Excel tự cao lên vì Wrap Text.

    Gán Value vẫn giữ font, màu, viền và number format. Chỉ tắt Wrap Text
    ở đúng ô lỗi rồi trả lại chiều cao dòng đang có trước khi ghi.
    """
    row_height = cell.EntireRow.RowHeight
    cell.Value = clean_text(error)[:1000]
    cell.WrapText = False
    if row_height is not None:
        cell.EntireRow.RowHeight = row_height


def write_url_faq_note(row: int, note: str) -> None:
    """Gắn Note vào ô URL, không làm thay đổi giá trị URL trong ô."""
    message = clean_text(note)
    if not message:
        return
    has_h2_warning = "[H2]" in message
    has_faq_warning = not message.startswith("[H2]")
    _excel, workbook = get_target_excel_workbook()
    sheet = workbook.Worksheets(SHEET_POST)
    cell = sheet.Cells(
        row,
        find_column_by_header(
            sheet, PUBLISH_HEADERS["published_url"]
        ),
    )
    note_title = (
        "KIỂM TRA BÀI ĐĂNG"
        if has_h2_warning
        else "KIỂM TRA FAQ"
    )
    full_note = f"{note_title}: {message}"
    try:
        if cell.Comment is None:
            cell.AddComment(full_note)
        else:
            cell.Comment.Text(full_note)
        if has_h2_warning and has_faq_warning:
            # Tím nhạt: đồng thời có cảnh báo FAQ và thiếu H2.
            cell.Interior.Color = 228 + 223 * 256 + 236 * 65536
        elif has_h2_warning:
            # Xanh dương nhạt: nội dung có ít hơn 2 thẻ H2.
            cell.Interior.Color = 221 + 235 * 256 + 247 * 65536
        else:
            # Vàng nhạt: cảnh báo cấu trúc FAQ.
            cell.Interior.Color = 255 + 242 * 256 + 204 * 65536
        workbook.Save()
        print(f"    Đã gắn Note kiểm tra vào ô URL đã đăng, dòng {row}.")
    except Exception as exc:
        # Note chỉ để kiểm tra thủ công, không được làm dừng đăng bài.
        print(
            "    CẢNH BÁO: Không gắn được Note kiểm tra vào ô URL, "
            f"vẫn tiếp tục: {exc!r}"
        )


def format_publish_status_cell(cell: Any, status: str) -> None:
    """V1.0: show pre-save validation errors separately from publish errors."""
    text = clean_text(status).casefold()
    cell.Font.Bold = True
    if "lỗi kiểm tra" in text:
        cell.Interior.Color = 156 + 235 * 256 + 255 * 65536  # yellow
        cell.Font.Color = 0
    elif "lỗi id" in text:
        cell.Interior.Color = 102 + 192 * 256 + 255 * 65536  # orange
        cell.Font.Color = 0
    elif "lỗi đăng" in text:
        cell.Interior.Color = 192  # red
        cell.Font.Color = 255 + 255 * 256 + 255 * 65536
    elif "đã đăng" in text:
        cell.Interior.Color = 198 + 239 * 256 + 206 * 65536
        cell.Font.Color = 0 + 97 * 256
    else:
        cell.Interior.ColorIndex = -4142
        cell.Font.ColorIndex = -4105
        cell.Font.Bold = False


def add_publish_error_note(status_cell: Any, status: str, error: str) -> None:
    """Put the exact failed validation/publish conditions in the status Note."""
    detail = clean_text(error) or "Chưa nhận được mô tả chi tiết từ worker."
    detail = detail.replace("Kiểm tra trước Save không đạt: ", "")
    detail = detail.replace(" | ", "\n• ")
    note = (
        f"{clean_text(status)}\n"
        f"Thời điểm: {datetime.now():%d/%m/%Y %H:%M:%S}\n"
        f"Chi tiết:\n• {detail[:1500]}"
    )
    if status_cell.Comment is None:
        status_cell.AddComment(note)
    else:
        status_cell.Comment.Text(note)

    # Excel does not automatically enlarge legacy Notes after Comment.Text().
    # Size the note by its wrapped lines so all failed conditions are visible.
    visual_lines = sum(
        max(1, (len(line) + 64) // 65)
        for line in note.splitlines()
    )
    note_height = max(105, min(520, 28 + visual_lines * 19))
    try:
        shape = status_cell.Comment.Shape
        shape.Width = 440
        shape.Height = note_height
        shape.TextFrame.WordWrap = True
    except Exception as resize_error:
        # The Note text has already been written; size is only a display aid.
        print(
            "    CẢNH BÁO: Không tự giãn được khung Note lỗi: "
            f"{resize_error!r}"
        )


def write_publish_result(
    row: int,
    status: str,
    url: str = "",
    error: str = "",
    cms_id: str = "",
) -> None:
    _excel, workbook = get_target_excel_workbook()
    sheet = workbook.Worksheets(SHEET_POST)
    status_cell = sheet.Cells(
        row,
        find_column_by_header(sheet, PUBLISH_HEADERS["status"]),
    )
    status_cell.Value = status
    format_publish_status_cell(status_cell, status)
    if url:
        sheet.Cells(row, find_column_by_header(sheet, PUBLISH_HEADERS["published_url"])).Value = url
    if cms_id:
        cms_id_col = find_column_by_header(sheet, PUBLISH_HEADERS["cms_id"])
        existing_id = clean_text(sheet.Cells(row, cms_id_col).Value)
        if cms_id == "LỖI ID":
            if not existing_id or existing_id == "LỖI ID":
                sheet.Cells(row, cms_id_col).Value = "LỖI ID"
            else:
                print(
                    f'    CẢNH BÁO: Dòng {row} đã có ID CMS "{existing_id}", '
                    'không ghi đè bằng "LỖI ID".'
                )
            cms_id = ""
        try:
            normalized_existing_id = (
                str(int(float(existing_id))) if existing_id else ""
            )
        except (TypeError, ValueError):
            normalized_existing_id = existing_id
        if cms_id and normalized_existing_id and normalized_existing_id != cms_id:
            raise RuntimeError(
                f'Dòng {row} đã có ID CMS "{existing_id}", không ghi đè bằng "{cms_id}".'
            )
        if cms_id:
            sheet.Cells(row, cms_id_col).Value = int(cms_id)
    sheet.Cells(
        row, find_column_by_header(sheet, PUBLISH_HEADERS["published_at"])
    ).Value2 = excel_local_now_serial()
    write_compact_error_cell(
        sheet.Cells(
            row,
            find_column_by_header(sheet, PUBLISH_HEADERS["publish_error"]),
        ),
        error,
    )
    if "lỗi" in clean_text(status).casefold():
        try:
            add_publish_error_note(status_cell, status, error)
            workbook.Save()
        except Exception as note_error:
            print(
                "    CẢNH BÁO: Không ghi được Note mô tả lỗi vào trạng thái: "
                f"{note_error!r}"
            )
    workbook.Save()


def upsert_category_id_mapping(
    row: int,
    category: str,
    category_id: str,
) -> dict[str, str]:
    """
    Tạm thời tự bổ sung DANH_MUC_ID từ ID CMS vừa chọn.

    Khi bảng DANH_MUC_ID đã chạy ổn định đủ lâu và đầy đủ dữ liệu, có thể bỏ
    lời gọi bước phụ này mà không ảnh hưởng flow đăng bài chính.
    """
    normalized_category = clean_text(category)
    normalized_id = clean_text(category_id)
    if not normalized_category or not normalized_id:
        return {"status": "SKIP", "message": ""}

    _excel, workbook = get_target_excel_workbook()
    post_sheet = workbook.Worksheets(SHEET_POST)
    map_sheet = workbook.Worksheets(SHEET_CATEGORY_ID)

    domain_col = find_column_by_header(
        post_sheet,
        PUBLISH_HEADERS["domain"],
    )
    domain = normalize_domain(
        post_sheet.Cells(row, domain_col).Value
    )
    if not domain:
        return {
            "status": "WARNING",
            "message": f"Dòng {row} không có tên miền để lưu ID danh mục.",
        }

    map_domain_col = find_column_by_header(map_sheet, "Tên miền URL")
    map_id_col = find_column_by_header(map_sheet, "ID")
    map_category_col = find_column_by_header(map_sheet, "Danh mục")
    last_row = int(
        map_sheet.Cells(
            map_sheet.Rows.Count,
            map_domain_col,
        ).End(-4162).Row
    )
    category_key = normalized_category.casefold()

    map_first_col = min(map_domain_col, map_id_col, map_category_col)
    map_last_col = max(map_domain_col, map_id_col, map_category_col)
    map_values = None
    if last_row >= 2:
        map_values = map_sheet.Range(
            map_sheet.Cells(2, map_first_col),
            map_sheet.Cells(last_row, map_last_col),
        ).Value2

    def map_value(map_row: int, map_col: int) -> Any:
        row_offset = map_row - 2
        col_offset = map_col - map_first_col
        if isinstance(map_values, tuple):
            row_values = map_values[row_offset]
            if isinstance(row_values, tuple):
                return row_values[col_offset]
            return row_values if col_offset == 0 else None
        return map_values if row_offset == 0 and col_offset == 0 else None

    for map_row in range(2, max(last_row, 1) + 1):
        existing_domain = normalize_domain(
            map_value(map_row, map_domain_col)
        )
        existing_category = clean_text(
            map_value(map_row, map_category_col)
        ).casefold()
        if existing_domain != domain or existing_category != category_key:
            continue

        existing_id = clean_text(
            map_value(map_row, map_id_col)
        )
        try:
            same_id = int(float(existing_id)) == int(float(normalized_id))
        except (TypeError, ValueError):
            same_id = existing_id == normalized_id

        if same_id:
            print(
                f"    DANH_MUC_ID đã có: {domain} | "
                f"{normalized_category} | ID {normalized_id}"
            )
            return {"status": "EXISTS", "message": ""}

        message = (
            f'DANH_MUC_ID đang có ID "{existing_id}" nhưng CMS vừa trả '
            f'ID "{normalized_id}" cho {domain} | {normalized_category}. '
            "Không ghi đè."
        )
        print("    CẢNH BÁO:", message)
        return {"status": "WARNING", "message": message}

    new_row = max(last_row + 1, 2)
    map_sheet.Cells(new_row, map_domain_col).Value = domain
    map_sheet.Cells(new_row, map_id_col).Value = int(float(normalized_id))
    map_sheet.Cells(new_row, map_category_col).Value = normalized_category
    workbook.Save()
    print(
        f"    Đã thêm DANH_MUC_ID dòng {new_row}: "
        f"{domain} | {normalized_category} | ID {normalized_id}"
    )
    return {"status": "ADDED", "message": ""}


# ============================================================
# HÀM DỮ LIỆU DANH MỤC / TÁC GIẢ
# Giữ logic từ code cũ:
# - Nhận trực tiếp dòng Excel đang xử lý
# - Đọc sheet dangbai
# - Map Tên miền + Danh mục trong sheet tác giả
# - Chọn ngẫu nhiên một tác giả phù hợp
# ============================================================

def normalize_text(value: Any) -> str:
    return clean_text(value).lower()


def get_author_columns(ws: Any) -> list[int]:
    last_col = int(
        ws.Cells(
            1,
            ws.Columns.Count,
        ).End(-4159).Column
    )

    columns: list[int] = []

    for col in range(1, last_col + 1):
        header = normalize_text(
            ws.Cells(1, col).Value
        )

        if (
            header.startswith("tác giả")
            or header.startswith("tac gia")
        ):
            columns.append(col)

    return columns


def read_category_author_data(
    current_row: int,
) -> dict[str, Any]:
    """
    Trả về:
    {
        "domain": "...",
        "category": "...",
        "authors": [...],
        "author": "..."
    }

    Quy tắc:
    - Danh mục trống: bỏ qua cả bước, không phải lỗi.
    - Không có tác giả phù hợp: vẫn chọn Danh mục,
      sau đó bỏ qua Tác giả, không phải lỗi.
    """
    try:
        excel, wb = get_target_excel_workbook()
        sh_post = wb.Worksheets(SHEET_POST)
        sh_author = wb.Worksheets(SHEET_AUTHOR)
    except Exception as exc:
        raise RuntimeError(
            'Không kết nối được Excel hoặc thiếu sheet '
            '"dangbai" / "tác giả".'
        ) from exc

    domain_col = find_column_by_header(
        sh_post,
        HEADER_DOMAIN,
    )

    category_col = find_column_by_header(
        sh_post,
        HEADER_CATEGORY,
    )

    domain = normalize_domain(
        sh_post.Cells(
            current_row,
            domain_col,
        ).Value
    )

    category = clean_text(
        sh_post.Cells(
            current_row,
            category_col,
        ).Value
    )

    if not category:
        return {
            "domain": domain,
            "category": "",
            "authors": [],
            "author": "",
        }

    author_domain_col = find_column_by_header(
        sh_author,
        HEADER_DOMAIN,
    )

    author_category_col = find_column_by_header(
        sh_author,
        HEADER_CATEGORY,
    )

    author_cols = get_author_columns(
        sh_author
    )

    if not author_cols:
        raise RuntimeError(
            'Không tìm thấy các cột "Tác giả 1", '
            '"Tác giả 2"... trong sheet tác giả.'
        )

    last_row = int(
        sh_author.Cells(
            sh_author.Rows.Count,
            author_domain_col,
        ).End(-4162).Row
    )

    matched_authors: list[str] = []

    for row in range(2, last_row + 1):
        row_domain = normalize_domain(
            sh_author.Cells(
                row,
                author_domain_col,
            ).Value
        )

        row_category = normalize_text(
            sh_author.Cells(
                row,
                author_category_col,
            ).Value
        )

        if (
            row_domain == domain
            and row_category
            == normalize_text(category)
        ):
            for col in author_cols:
                value = clean_text(
                    sh_author.Cells(
                        row,
                        col,
                    ).Value
                )

                if value:
                    matched_authors.append(
                        value
                    )

            break

    author = (
        random.choice(matched_authors)
        if matched_authors
        else ""
    )

    return {
        "domain": domain,
        "category": category,
        "authors": matched_authors,
        "author": author,
    }


def normalize_option_text(value: Any) -> str:
    return re.sub(
        r"\s+",
        " ",
        str(value or ""),
    ).strip().lower()


def get_visible_select2_search(driver):
    """
    Lấy ô tìm kiếm Select2 đang hiển thị.
    Không dùng bàn phím/chuột thật của Windows.
    """
    search_boxes = driver.find_elements(
        By.CSS_SELECTOR,
        "input.select2-search__field",
    )

    visible_boxes = [
        element
        for element in search_boxes
        if element.is_displayed()
    ]

    if not visible_boxes:
        raise RuntimeError(
            "Select2 đã mở nhưng không thấy ô tìm kiếm."
        )

    return visible_boxes[-1]


def open_select2(
    driver,
    select_css: str,
    label: str,
) -> None:
    """
    Mở đúng giao diện Select2 tương ứng với thẻ select gốc.
    """
    selection = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: d.execute_script(
            """
            const select = document.querySelector(
                arguments[0]
            );

            if (!select) {
                return null;
            }

            let wrapper = select.nextElementSibling;

            if (
                !wrapper
                || !wrapper.classList
                || !wrapper.classList.contains('select2')
            ) {
                wrapper = select.parentElement
                    ? select.parentElement.querySelector(
                        'span.select2'
                    )
                    : null;
            }

            if (!wrapper) {
                return null;
            }

            return wrapper.querySelector(
                '.select2-selection'
            );
            """,
            select_css,
        )
    )

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block: 'center'
        });
        """,
        selection,
    )

    time.sleep(0.3)

    try:
        selection.click()
    except Exception:
        driver.execute_script(
            "arguments[0].click();",
            selection,
        )

    WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: bool(
            [
                element
                for element in d.find_elements(
                    By.CSS_SELECTOR,
                    "input.select2-search__field",
                )
                if element.is_displayed()
            ]
        )
    )

    print(f"Đã mở Select2 {label}.")


def find_visible_select2_option(
    driver,
    wanted_text: str,
):
    wanted = normalize_option_text(
        wanted_text
    )

    options = driver.find_elements(
        By.CSS_SELECTOR,
        (
            "li.select2-results__option"
            "[role='option'], "
            "li.select2-results__option"
        ),
    )

    visible_options = [
        option
        for option in options
        if option.is_displayed()
    ]

    # Ưu tiên trùng chính xác.
    for option in visible_options:
        option_text = normalize_option_text(
            option.text
        )

        if option_text == wanted:
            return option

    # Fallback chứa nhau, nhưng bỏ các dòng trạng thái Select2.
    ignored = {
        "searching…",
        "searching...",
        "đang tìm kiếm…",
        "đang tìm kiếm...",
        "no results found",
        "không tìm thấy kết quả",
    }

    for option in visible_options:
        option_text = normalize_option_text(
            option.text
        )

        if not option_text or option_text in ignored:
            continue

        if (
            wanted in option_text
            or option_text in wanted
        ):
            return option

    return None


def choose_select2_by_text(
    driver,
    select_css: str,
    visible_text: str,
    label: str,
) -> dict[str, str]:
    """
    Thao tác giống người dùng:
    - mở Select2;
    - nhập vào ô tìm kiếm;
    - click đúng kết quả.

    Đây là Selenium DOM, không dùng pyautogui,
    clipboard hoặc tọa độ.
    """
    open_select2(
        driver,
        select_css,
        label,
    )

    search_box = get_visible_select2_search(
        driver
    )

    search_box.send_keys(
        Keys.CONTROL,
        "a",
    )
    search_box.send_keys(
        Keys.BACKSPACE
    )
    search_box.send_keys(
        visible_text
    )

    option = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: find_visible_select2_option(
            d,
            visible_text,
        )
    )

    selected_text = option.text.strip()

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block: 'nearest'
        });
        """,
        option,
    )

    try:
        option.click()
    except Exception:
        driver.execute_script(
            "arguments[0].click();",
            option,
        )

    # Chờ popup Select2 đóng.
    WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: not any(
            element.is_displayed()
            for element in d.find_elements(
                By.CSS_SELECTOR,
                ".select2-container--open",
            )
        )
    )

    print(
        f"Đã chọn {label}: {selected_text}"
    )

    selected_value = clean_text(
        driver.execute_script(
            """
            const select = document.querySelector(arguments[0]);
            return select ? select.value : '';
            """,
            select_css,
        )
    )

    return {
        "text": selected_text,
        "value": selected_value,
    }


def wait_author_select_ready(
    driver,
) -> Any:
    """
    Sau khi chọn Danh mục, chỉ chờ thẻ Select2 Tác giả
    được website tạo lại.

    Không yêu cầu tên tác giả phải có sẵn trong option,
    vì Select2 có thể tải/tìm kiếm động khi mở dropdown.
    """
    def find_author_select(d):
        return d.execute_script(
            """
            const direct = document.querySelector([
                'select[name*="tac-gia" i]',
                'select[id*="tac-gia" i]',
                'select[name*="tac_gia" i]',
                'select[id*="tac_gia" i]',
                'select[name*="author" i]',
                'select[id*="author" i]'
            ].join(','));
            if (direct) return direct;

            const labels = Array.from(document.querySelectorAll(
                'label, .control-label, th, td, div, span'
            ));
            const label = labels.find(node => {
                const text = String(node.textContent || '')
                    .replace(/\\s+/g, ' ')
                    .trim()
                    .toLocaleLowerCase('vi');
                return text === 'tác giả' || text === 'tac gia';
            });
            if (!label) return null;

            if (label.htmlFor) {
                const linked = document.getElementById(label.htmlFor);
                if (linked && linked.tagName === 'SELECT') return linked;
            }
            const area = label.closest(
                '.form-group, .row, td, .control-group'
            ) || label.parentElement;
            return area ? area.querySelector('select') : null;
            """
        )

    author_select = WebDriverWait(driver, WAIT_PAGE).until(find_author_select)

    time.sleep(1)

    print("Đã tải xong ô Select2 Tác giả.")
    return author_select


def choose_random_author(driver, author_select) -> str:
    """Mở Select2 Tác giả, đọc danh sách CMS và chọn ngẫu nhiên một tên hợp lệ."""
    # Danh mục có thể làm Ajax thay mới toàn bộ thẻ select. Luôn tìm lại phần tử
    # ngay trước khi thao tác để tránh StaleElementReferenceException.
    author_select = WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: d.execute_script(
            """
            return document.querySelector([
                'select[name*="tac-gia" i]',
                'select[id*="tac-gia" i]',
                'select[name*="tac_gia" i]',
                'select[id*="tac_gia" i]',
                'select[name*="author" i]',
                'select[id*="author" i]'
            ].join(','));
            """
        )
    )
    selection = driver.execute_script(
        """
        const select = arguments[0];
        let wrapper = select.nextElementSibling;
        if (!wrapper || !wrapper.classList || !wrapper.classList.contains('select2')) {
            wrapper = select.parentElement
                ? select.parentElement.querySelector('span.select2')
                : null;
        }
        return wrapper ? wrapper.querySelector('.select2-selection') : null;
        """,
        author_select,
    )
    if selection is None:
        raise RuntimeError("Đã thấy trường Tác giả nhưng không thấy giao diện Select2.")

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        selection,
    )
    try:
        selection.click()
    except Exception:
        driver.execute_script("arguments[0].click();", selection)

    ignored = {
        "",
        "chọn tác giả",
        "-- chọn tác giả --",
        "tất cả",
        "đang tải",
        "đang tải...",
        "đang tìm kiếm...",
        "searching...",
        "no results found",
        "không tìm thấy kết quả",
    }

    def valid_options(d):
        results = []
        for option in d.find_elements(
            By.CSS_SELECTOR,
            "li.select2-results__option[role='option'], li.select2-results__option",
        ):
            if not option.is_displayed():
                continue
            text = normalize_option_text(option.text)
            if (
                text in ignored
                or option.get_attribute("aria-disabled") == "true"
                or "loading-results" in (option.get_attribute("class") or "")
            ):
                continue
            results.append(option)
        return results or False

    options = WebDriverWait(driver, WAIT_PAGE).until(valid_options)
    chosen = random.choice(options)
    chosen_text = clean_text(chosen.text)
    try:
        chosen.click()
    except Exception:
        driver.execute_script("arguments[0].click();", chosen)

    WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: not any(
            element.is_displayed()
            for element in d.find_elements(By.CSS_SELECTOR, ".select2-container--open")
        )
    )
    print(f"Đã random Tác giả từ CMS: {chosen_text}")
    return chosen_text


def select_category_author(
    driver,
    current_row: int,
) -> dict[str, str]:
    print(
        "[6] Đọc Excel và chọn Danh mục / Tác giả..."
    )

    _excel, workbook = get_target_excel_workbook()
    post_sheet = workbook.Worksheets(SHEET_POST)
    category = clean_text(
        post_sheet.Cells(
            current_row,
            find_column_by_header(post_sheet, HEADER_CATEGORY),
        ).Value
    )
    if not category:
        print(
            f'Dòng {current_row} không có "Danh mục" '
            "→ bỏ qua Danh mục và Tác giả."
        )

        return {
            "status": "SKIP",
            "category": "",
            "author": "",
        }

    category_selector = (
        'select[name="cat_id"]'
    )

    category_result = choose_select2_by_text(
        driver,
        category_selector,
        category,
        "Danh mục",
    )
    category_id = clean_text(category_result.get("value"))
    if not category_id:
        raise RuntimeError(
            f'Đã chọn Danh mục "{category}" nhưng CMS không trả ID danh mục.'
        )
    print(f"ID Danh mục từ CMS: {category_id}")

    # Sau khi Danh mục thay đổi, website tải Tác giả bằng Ajax. Nếu DOM bị
    # thay đúng lúc click thì tìm lại và thử đúng một lần, không thử vô hạn.
    for author_attempt in range(2):
        try:
            author_select = wait_author_select_ready(driver)
            author = choose_random_author(driver, author_select)
            break
        except StaleElementReferenceException as exc:
            if author_attempt == 0:
                print(
                    "    Ô Tác giả vừa bị Ajax thay mới "
                    "→ tìm lại và thử thêm một lần."
                )
                time.sleep(0.5)
                continue
            raise RuntimeError(
                "Ô Tác giả tiếp tục bị Ajax thay mới sau một lần thử lại."
            ) from exc
        except TimeoutException as exc:
            raise RuntimeError(
                "Không tìm thấy Tác giả hợp lệ sau khi chờ website tải danh sách."
            ) from exc

    return {
        "status": "OK",
        "category": category,
        "category_id": category_id,
        "author": author,
    }


# ============================================================
# BƯỚC 1: MỞ WORD THEO CODE CŨ
# ============================================================


# ============================================================
# BƯỚC 2: ĐỌC URL THEO CODE mo_url.py CŨ
# ============================================================



# ============================================================
# BƯỚC 3: ĐỌC 5 DÒNG ĐẦU TỪ WORD
# Không xóa, không cắt.
# ============================================================

def read_first_five_word_lines() -> dict[str, str]:
    data, _content_start = read_word_metadata(
        get_hidden_word_document()
    )
    return data


# ============================================================
# BƯỚC 4: LẤY NỘI DUNG WORD TỪ DÒNG 6 TRỞ ĐI DẠNG HTML
# Giữ Heading, đậm, nghiêng, bảng và liên kết.
# ============================================================

def clean_word_paragraph_text(
    raw_text: Any,
) -> str:
    value = str(raw_text or "")

    value = (
        value.replace("\r", "")
        .replace("\x07", "")
        .strip()
    )

    # Hỗ trợ trường hợp tiêu đề còn tiền tố Markdown.
    value = re.sub(
        r"^\s*#{1,6}\s*",
        "",
        value,
    )

    # Hỗ trợ trường hợp còn tiền tố H2:/H3:.
    value = re.sub(
        r"^\s*[Hh][123]\s*:\s*",
        "",
        value,
    )

    value = re.sub(
        r"\s+",
        " ",
        value,
    ).strip()

    return value


_WORD_METADATA_ORDER = (
    "title",
    "description",
    "keyword",
    "name",
    "desc_short",
)

_WORD_METADATA_LABELS = {
    "title": "title",
    "title seo": "title",
    "tiêu đề": "title",
    "tiêu đề seo": "title",
    "description": "description",
    "description seo": "description",
    "meta description": "description",
    "mô tả": "description",
    "mô tả seo": "description",
    "keyword": "keyword",
    "keywords": "keyword",
    "từ khóa": "keyword",
}


def _clean_word_metadata_text(raw: Any) -> str:
    value = str(raw or "")
    value = value.replace("\r", "").replace("\x07", " ")
    value = re.sub(r"\s+", " ", value).strip()
    if value.startswith("# "):
        return value[2:].strip()
    if value.startswith("#") and not value.startswith("##"):
        return value[1:].strip()
    return value


def _word_metadata_label(value: Any) -> str | None:
    key = unicodedata.normalize("NFC", str(value or ""))
    key = key.strip().rstrip(":").casefold()
    key = re.sub(r"\s+", " ", key)
    return _WORD_METADATA_LABELS.get(key)


def _split_word_metadata_label(text: str) -> tuple[str | None, str]:
    """Return a known label and its inline value, if this paragraph has one."""
    value = _clean_word_metadata_text(text)
    before, separator, after = value.partition(":")
    if separator:
        label = _word_metadata_label(before)
        if label:
            return label, after.strip()
    # A bare word such as "keyword" may be actual data.  Only the explicit
    # ``Label:`` form activates the "take the next paragraph" behavior.
    return None, ""


def read_word_metadata(doc: Any) -> tuple[dict[str, str], int]:
    """Read Word metadata while accepting labels on their own paragraph.

    Legacy documents contain five unlabeled values. Some documents insert
    ``Title:``, ``Description:``, or ``Keyword:`` before the value; those
    labels are metadata markers, not content values.
    """
    data: dict[str, str] = {}
    pending_label: str | None = None
    saw_label = False
    content_start: int | None = None

    for index in range(1, doc.Paragraphs.Count + 1):
        paragraph = doc.Paragraphs(index)
        raw = str(paragraph.Range.Text or "")
        text = _clean_word_metadata_text(raw)
        if not text:
            continue

        label, inline_value = _split_word_metadata_label(text)
        if label:
            if pending_label is not None:
                raise RuntimeError(
                    f"Word thiếu giá trị cho {pending_label} trước nhãn {label}."
                )
            saw_label = True
            content_start = int(paragraph.Range.End)
            if inline_value:
                # Nhãn (Title:/Description:/Keyword:) đã được tách ở trên.
                # Không chạy process_f5_line() lần hai vì dấu ':' còn lại có
                # thể là một phần hợp lệ của giá trị, ví dụ:
                # "Title: Dòng tiền gia đình: cách theo dõi...".
                data[label] = _clean_word_metadata_text(inline_value)
                pending_label = None
            else:
                pending_label = label
            continue

        value = process_f5_line(text)
        if pending_label is not None:
            data[pending_label] = value
            pending_label = None
            content_start = int(paragraph.Range.End)
            continue

        next_key = next(
            (key for key in _WORD_METADATA_ORDER if key not in data),
            None,
        )
        if next_key is None:
            break

        data[next_key] = value
        content_start = int(paragraph.Range.End)

    if pending_label is not None:
        raise RuntimeError(f"Word thiếu giá trị cho nhãn {pending_label}.")

    required = _WORD_METADATA_ORDER
    missing = [key for key in required if not clean_text(data.get(key, ""))]
    if missing:
        raise RuntimeError(
            "Word thiếu metadata bắt buộc: " + ", ".join(missing)
        )

    if content_start is None:
        raise RuntimeError("Không xác định được điểm bắt đầu nội dung Word.")

    return data, content_start


def is_faq_heading(text: str) -> bool:
    """
    Xác định dòng bắt đầu phần FAQ.

    Chấp nhận:
    - FAQ
    - FAQs
    - Câu hỏi thường gặp
    - Câu hỏi thường gặp về ...
    """
    normalized = clean_word_paragraph_text(
        text
    ).lower()

    normalized = normalized.rstrip(
        ":：-–— "
    )

    if normalized in {
        "faq",
        "faqs",
        "câu hỏi thường gặp",
    }:
        return True

    if normalized.startswith(
        "câu hỏi thường gặp "
    ):
        return True

    return False


def find_faq_heading_from_bottom(
    doc: Any,
    max_check: int = 50,
) -> int | None:
    """Giữ đúng quy tắc F7 cũ: tìm FAQ từ cuối Word đi lên."""
    paragraph_count = int(doc.Paragraphs.Count)
    first_index = max(1, paragraph_count - max_check + 1)

    for index in range(
        paragraph_count,
        first_index - 1,
        -1,
    ):
        text = clean_word_paragraph_text(
            doc.Paragraphs(index).Range.Text
        ).lower()

        # AHK F7 cũ dùng InStr với hai chuỗi này.
        if (
            "câu hỏi thường gặp" in text
            or "faq" in text
        ):
            return index

    return None


def get_word_content_range(
    doc: Any,
) -> tuple[int, int, bool]:
    """
    Tìm vùng nội dung chính:

    Bắt đầu:
    - ngay sau 5 đoạn không rỗng đầu tiên.

    Kết thúc:
    - ngay trước dòng FAQ/Câu hỏi thường gặp;
    - nếu không có FAQ thì lấy đến cuối Word.

    Trả về:
    (content_start, content_end, found_faq)
    """
    _metadata, content_start = read_word_metadata(doc)
    faq_heading_index = find_faq_heading_from_bottom(doc)
    found_faq = faq_heading_index is not None

    if faq_heading_index is not None:
        content_end = int(
            doc.Paragraphs(faq_heading_index).Range.Start
        )
    else:
        content_end = int(doc.Content.End)

    return (
        content_start,
        content_end,
        found_faq,
    )


def _read_clipboard_html_bytes(
    timeout: float = 10.0,
) -> bytes:
    """
    Đọc định dạng CF_HTML ("HTML Format") từ Windows Clipboard.

    Word tạo CF_HTML khi gọi Range.Copy(), tương tự thao tác Ctrl+C
    trực tiếp trong Word.
    """
    html_format = win32clipboard.RegisterClipboardFormat(
        "HTML Format"
    )

    deadline = time.time() + timeout
    last_error: Exception | None = None

    while time.time() < deadline:
        try:
            win32clipboard.OpenClipboard()

            try:
                if not win32clipboard.IsClipboardFormatAvailable(
                    html_format
                ):
                    raise RuntimeError(
                        "Clipboard chưa có định dạng HTML Format."
                    )

                data = win32clipboard.GetClipboardData(
                    html_format
                )

                if isinstance(data, bytes):
                    return data

                if isinstance(data, str):
                    return data.encode(
                        "utf-8",
                        errors="replace",
                    )

                raise RuntimeError(
                    "Dữ liệu HTML Clipboard có kiểu không hỗ trợ: "
                    + type(data).__name__
                )

            finally:
                win32clipboard.CloseClipboard()

        except Exception as exc:
            last_error = exc
            time.sleep(0.15)

    raise RuntimeError(
        "Không đọc được HTML từ Clipboard sau khi copy Word."
    ) from last_error


def _parse_cf_html_fragment(raw: bytes) -> str:
    """
    Tách đúng phần fragment từ CF_HTML của Windows.
    """
    header_text = raw[:4096].decode(
        "ascii",
        errors="ignore",
    )

    def read_offset(name: str) -> int | None:
        match = re.search(
            rf"(?im)^{re.escape(name)}:\\s*(\\d+)\\s*$",
            header_text,
        )

        if not match:
            return None

        try:
            return int(match.group(1))
        except ValueError:
            return None

    start_fragment = read_offset("StartFragment")
    end_fragment = read_offset("EndFragment")

    fragment_bytes: bytes | None = None

    if (
        start_fragment is not None
        and end_fragment is not None
        and 0 <= start_fragment < end_fragment <= len(raw)
    ):
        fragment_bytes = raw[start_fragment:end_fragment]

    if fragment_bytes is None:
        start_marker = b"<!--StartFragment-->"
        end_marker = b"<!--EndFragment-->"

        marker_start = raw.find(start_marker)
        marker_end = raw.find(end_marker)

        if marker_start >= 0 and marker_end > marker_start:
            marker_start += len(start_marker)
            fragment_bytes = raw[marker_start:marker_end]

    if fragment_bytes is None:
        start_html = read_offset("StartHTML")
        end_html = read_offset("EndHTML")

        if (
            start_html is not None
            and end_html is not None
            and 0 <= start_html < end_html <= len(raw)
        ):
            fragment_bytes = raw[start_html:end_html]
        else:
            fragment_bytes = raw

    html = fragment_bytes.decode(
        "utf-8",
        errors="replace",
    )

    return unicodedata.normalize(
        "NFC",
        html,
    ).strip()


def _remove_word_section_wrapper(
    html: str,
) -> str:
    """
    Bỏ div WordSection1/WordSection2 bao ngoài nhưng giữ nội dung.
    """
    value = str(html or "").strip()

    opening = re.match(
        r"(?is)^\\s*<div\\b(?=[^>]*\\bclass\\s*=\\s*[\"'][^\"']*\\bWordSection\\d*\\b[^\"']*[\"'])[^>]*>\\s*",
        value,
    )

    if not opening:
        return value

    value = value[opening.end():]

    value = re.sub(
        r"(?is)\\s*</div>\\s*$",
        "",
        value,
        count=1,
    )

    return value.strip()


def read_word_content_as_html() -> str:
    """
    Lấy nội dung Word bằng Clipboard HTML giống thao tác Ctrl+C:

    - Lấy từ sau 5 dòng đầu đến trước FAQ.
    - Gọi Range.Copy().
    - Đọc CF_HTML trực tiếp từ Windows Clipboard.
    - Không tạo file HTML tạm.
    """
    source_doc = get_hidden_word_document()

    (
        content_start,
        content_end,
        found_faq,
    ) = get_word_content_range(
        source_doc
    )

    if found_faq:
        print(
            "Đã tìm thấy phần FAQ → "
            "chỉ lấy nội dung chính trước FAQ."
        )
    else:
        print(
            "Không thấy tiêu đề FAQ → "
            "lấy nội dung đến cuối Word."
        )

    if content_start >= content_end:
        raise RuntimeError(
            "Không có nội dung Word từ dòng 6 trở đi."
        )

    source_range = source_doc.Range(
        Start=content_start,
        End=content_end,
    )

    clipboard_lock = WORD_CLIPBOARD_LOCK
    if clipboard_lock is not None:
        print("    Waiting for Word Clipboard transfer...")
        clipboard_lock.acquire()
    try:
        try:
            source_range.Copy()
        except Exception as exc:
            raise RuntimeError(
                "Khong copy duoc vung noi dung Word."
            ) from exc

        raw_html = _read_clipboard_html_bytes(timeout=10.0)
    finally:
        if clipboard_lock is not None:
            clipboard_lock.release()

    # The lock is released as soon as the CF_HTML bytes are privately held.
    # HTML parsing and every CMS step continue concurrently.
    """
    try:
        source_range.Copy()
    except Exception as exc:
        raise RuntimeError(
            "Không copy được vùng nội dung Word."
        ) from exc

    raw_html = _read_clipboard_html_bytes(timeout=10.0)
    """

    body_html = _parse_cf_html_fragment(
        raw_html
    )

    body_html = _remove_word_section_wrapper(
        body_html
    )

    body_html = unicodedata.normalize(
        "NFC",
        body_html,
    ).strip()

    if not body_html:
        raise RuntimeError(
            "HTML lấy từ Clipboard Word bị rỗng."
        )

    print(
        "Đã lấy HTML trực tiếp từ Clipboard Word, "
        f"độ dài: {len(body_html)} ký tự."
    )

    return body_html


# ============================================================
# BƯỚC 4: MỞ FIREFOX SELENIUM
# Selenium Manager tự tìm / tải driver nếu cần.
# ============================================================



# ============================================================
# BƯỚC 5: ĐIỀN 5 Ô
# ============================================================

def set_value(
    driver,
    by: By,
    selector: str,
    value: str,
) -> None:
    element = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: d.find_element(
            by,
            selector,
        )
    )

    driver.execute_script(
        """
        const el = arguments[0];
        const value = arguments[1];

        el.focus();
        el.value = value;

        el.dispatchEvent(
            new Event('input', {bubbles: true})
        );
        el.dispatchEvent(
            new Event('change', {bubbles: true})
        );
        el.dispatchEvent(
            new Event('blur', {bubbles: true})
        );
        """,
        element,
        value,
    )


def fill_five_fields(
    driver,
    data: dict[str, str],
) -> None:
    print("[3] Điền 5 ô đầu...")

    set_value(
        driver,
        By.ID,
        "name",
        data["name"],
    )

    set_value(
        driver,
        By.ID,
        "title",
        data["title"],
    )

    set_value(
        driver,
        By.ID,
        "description",
        data["description"],
    )

    set_value(
        driver,
        By.ID,
        "keyword",
        data["keyword"],
    )

    set_value(
        driver,
        By.ID,
        "desc_short",
        data["desc_short"],
    )

    time.sleep(WAIT_AFTER_FILL)


# ============================================================
# BƯỚC 6: FOCUS CKEDITOR VÀ NẠP HTML WORD
# ============================================================

def focus_ckeditor(driver) -> None:
    print("[4] Đưa Edge lên trước và focus CKEditor detail...")

    # Selenium vẫn điều khiển DOM bình thường dù cửa sổ Edge đang ẩn.
    driver.switch_to.window(
        driver.current_window_handle
    )

    if not _EDGE_IS_HIDDEN:
        driver.maximize_window()
        driver.execute_script("window.focus();")
        time.sleep(0.5)

    # Chờ đúng CKEditor của textarea name="detail".
    WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: d.execute_script(
            """
            return (
                typeof CKEDITOR !== 'undefined'
                && CKEDITOR.instances
                && CKEDITOR.instances.detail
                && CKEDITOR.instances.detail.status === 'ready'
            );
            """
        )
    )

    # Focus trực tiếp đúng instance detail và vùng body contenteditable
    # nằm bên trong iframe cke_detail.
    driver.execute_script(
        """
        const editor = CKEDITOR.instances.detail;

        editor.focus();

        const editable = editor.editable();

        if (editable) {
            editable.focus();

            if (editable.$) {
                editable.$.focus();
                editable.$.click();
            }
        }
        """
    )

    time.sleep(1)


def set_ckeditor_html(
    driver,
    html_content: str,
) -> None:
    print("[5] Đưa HTML Word vào CKEditor detail...")

    result = driver.execute_async_script(
        """
        const html = arguments[0];
        const done = arguments[arguments.length - 1];

        try {
            if (
                typeof CKEDITOR === 'undefined'
                || !CKEDITOR.instances
                || !CKEDITOR.instances.detail
            ) {
                done({
                    ok: false,
                    error: 'Không tìm thấy CKEDITOR.instances.detail'
                });
                return;
            }

            const editor = CKEDITOR.instances.detail;

            editor.setData(html, {
                callback: function () {
                    editor.focus();
                    editor.fire('change');

                    done({
                        ok: true,
                        length: editor.getData().length
                    });
                }
            });
        } catch (error) {
            done({
                ok: false,
                error: String(error)
            });
        }
        """,
        html_content,
    )

    if not result or not result.get("ok"):
        error_text = (
            result.get("error")
            if isinstance(result, dict)
            else str(result)
        )

        raise RuntimeError(
            "Không đưa được HTML vào CKEditor: "
            + error_text
        )

    print(
        "Đã nạp CKEditor, độ dài HTML:",
        result.get("length", 0),
    )

    time.sleep(WAIT_AFTER_SET_DATA)



# ============================================================
# BƯỚC 7: ĐỌC VÀ NHẬP FAQ BẰNG SELENIUM
# Quy tắc giữ giống AHK cũ:
# - Tìm dòng FAQ / Câu hỏi thường gặp trong Word.
# - Sau dòng đó, cứ 2 đoạn không rỗng là 1 cặp:
#       Câu hỏi
#       Trả lời
# - Không cắt hoặc xóa nội dung trong Word.
# ============================================================

def is_word_bullet_paragraph(paragraph, raw_text: str = "") -> bool:
    """Nhận diện bullet thật của Word hoặc bullet được gõ thành ký tự."""
    try:
        # Microsoft Word: wdListBullet = 2.
        if int(paragraph.Range.ListFormat.ListType) == 2:
            return True
    except Exception:
        pass

    return bool(
        re.match(
            r"^\s*[•●▪◦‣⁃]\s+",
            str(raw_text or ""),
        )
    )


def get_word_paragraph_style_name(paragraph) -> str:
    """Đọc tên style Word COM, hỗ trợ cả NameLocal và giá trị chuỗi."""
    try:
        style = paragraph.Range.Style
        name = getattr(style, "NameLocal", None)
        if name:
            return clean_text(name).casefold()
        return clean_text(style).casefold()
    except Exception:
        return ""


def is_faq_question_heading(paragraph) -> bool:
    style = get_word_paragraph_style_name(paragraph)
    return (
        style.startswith("heading ")
        or style.startswith("tiêu đề ")
        or style.startswith("title ")
    )


def read_faq_pairs_from_word() -> list[dict[str, str]]:
    """
    Đọc FAQ từ Word theo đúng cấu trúc AHK cũ.

    Ví dụ:
        FAQ
        Câu hỏi 1
        Trả lời 1
        Câu hỏi 2
        Trả lời 2

    Trả về:
        [
            {"question": "...", "answer": "..."},
            ...
        ]
    """
    set_faq_run_note("")
    doc = get_hidden_word_document()

    faq_heading_index = find_faq_heading_from_bottom(doc)
    if faq_heading_index is None:
        print(
            "Không có FAQ trong 50 đoạn cuối Word → bỏ qua bước FAQ."
        )
        return []

    faq_items: list[dict[str, Any]] = []
    for index in range(
        faq_heading_index + 1,
        doc.Paragraphs.Count + 1,
    ):
        paragraph = doc.Paragraphs(index)
        raw = str(
            paragraph.Range.Text or ""
        )

        clean = clean_word_paragraph_text(raw)

        if not clean:
            continue

        # FAQ có bullet/list không đúng cấu trúc cặp câu hỏi - trả lời.
        # Giữ nguyên Word và nội dung chính, nhưng bỏ toàn bộ bước đăng FAQ.
        if is_word_bullet_paragraph(paragraph, raw):
            set_faq_run_note(
                "FAQ có bullet/list nên đã bỏ qua toàn bộ phần FAQ."
            )
            print(
                "FAQ có bullet/list → bỏ qua toàn bộ bước đăng FAQ."
            )
            return []

        value = process_f5_line(clean)

        if value:
            faq_items.append(
                {
                    "text": value,
                    "is_heading": is_faq_question_heading(paragraph),
                }
            )

    if not faq_items:
        set_faq_run_note(
            "Có tiêu đề FAQ nhưng không có câu hỏi/trả lời."
        )
        print(
            "Có tiêu đề FAQ nhưng không có câu hỏi/trả lời "
            "→ bỏ qua bước FAQ."
        )
        return []

    # Nếu Word có Heading trong phần FAQ thì Heading là mốc chắc chắn nhất.
    # Chỉ khi hoàn toàn không có Heading mới dùng dấu ? làm phương án dự phòng.
    has_heading_markers = any(item["is_heading"] for item in faq_items)

    pairs: list[dict[str, str]] = []
    skipped_questions: list[str] = []
    orphan_lines: list[str] = []
    current_question = ""
    current_answers: list[str] = []

    def finish_current_question() -> None:
        nonlocal current_question, current_answers
        if not current_question:
            return
        if current_answers:
            pairs.append(
                {
                    "question": current_question,
                    # Ô CMS là textarea nên giữ được nhiều đoạn bằng xuống dòng.
                    "answer": "\n\n".join(current_answers),
                }
            )
        else:
            skipped_questions.append(current_question)
        current_question = ""
        current_answers = []

    for item in faq_items:
        text = clean_text(item["text"])
        is_question = (
            bool(item["is_heading"])
            if has_heading_markers
            else text.rstrip().endswith("?")
        )
        if is_question:
            finish_current_question()
            current_question = text
        elif current_question:
            current_answers.append(text)
        else:
            orphan_lines.append(text)

    finish_current_question()

    if not pairs:
        set_faq_run_note(
            "Không nhận diện được cặp câu hỏi/trả lời FAQ chắc chắn."
        )
        print(
            "Không nhận diện được cặp FAQ chắc chắn "
            "→ bỏ qua toàn bộ bước FAQ."
        )
        return []

    if orphan_lines:
        print(
            f"CẢNH BÁO: Bỏ qua {len(orphan_lines)} đoạn FAQ "
            "nằm trước câu hỏi đầu tiên."
        )
    if skipped_questions:
        print(
            f"CẢNH BÁO: Bỏ qua {len(skipped_questions)} câu hỏi FAQ "
            "không có phần trả lời."
        )

    notes: list[str] = []
    if orphan_lines:
        notes.append(
            f"Đã bỏ {len(orphan_lines)} đoạn nằm trước câu hỏi đầu tiên"
        )
    if skipped_questions:
        notes.append(
            f"Đã bỏ {len(skipped_questions)} câu hỏi không có trả lời"
        )

    last_answer = pairs[-1]["answer"]
    last_answer_paragraphs = [
        paragraph
        for paragraph in last_answer.split("\n\n")
        if clean_text(paragraph)
    ]
    last_answer_word_count = len(
        re.findall(r"\S+", last_answer)
    )
    previous_answer_word_counts = [
        len(re.findall(r"\S+", pair["answer"]))
        for pair in pairs[:-1]
    ]
    last_answer_warnings: list[str] = []

    if len(last_answer_paragraphs) >= 2:
        last_answer_warnings.append(
            f"có {len(last_answer_paragraphs)} đoạn trả lời"
        )
    if last_answer_word_count > 200:
        last_answer_warnings.append(
            f"dài {last_answer_word_count} từ (vượt 200 từ)"
        )
    if previous_answer_word_counts:
        previous_average = (
            sum(previous_answer_word_counts)
            / len(previous_answer_word_counts)
        )
        if (
            previous_average > 0
            and last_answer_word_count > previous_average * 2.5
        ):
            last_answer_warnings.append(
                "dài gấp "
                f"{last_answer_word_count / previous_average:.1f} lần "
                "trung bình các câu trả lời trước"
            )
    if last_answer_warnings:
        notes.append(
            "Câu trả lời FAQ cuối cần kiểm tra: "
            + "; ".join(last_answer_warnings)
        )

    if notes:
        set_faq_run_note("; ".join(notes) + ".")

    print(
        f"Đã đọc được {len(pairs)} cặp FAQ hợp lệ từ Word."
    )

    return pairs


def inspect_word_structure_before_close() -> dict[str, Any]:
    """
    Chốt cấu trúc Word để dùng tại cổng kiểm tra ngay trước Save CMS.

    Chỉ xét vùng nội dung sau 5 dòng dữ liệu đầu và trước FAQ.
    """
    doc = get_hidden_word_document()
    content_start, content_end, has_faq = get_word_content_range(doc)
    h2_count = 0
    last_non_empty = None

    for index in range(1, doc.Paragraphs.Count + 1):
        paragraph = doc.Paragraphs(index)
        paragraph_start = int(paragraph.Range.Start)
        paragraph_end = int(paragraph.Range.End)
        if paragraph_end <= content_start:
            continue
        if paragraph_start >= content_end:
            break

        raw = str(paragraph.Range.Text or "")
        text = clean_word_paragraph_text(raw)
        if not text:
            continue

        style_name = get_word_paragraph_style_name(paragraph)
        is_h2 = bool(
            re.match(
                r"^(?:heading|tiêu đề|title)\s*2(?:\b|$)",
                style_name,
                re.IGNORECASE,
            )
        )
        if not is_h2:
            try:
                # Word: wdOutlineLevel2 = 2.
                is_h2 = int(
                    paragraph.Range.ParagraphFormat.OutlineLevel
                ) == 2
            except Exception:
                pass
        if is_h2:
            h2_count += 1

        last_non_empty = (paragraph, raw, text)

    ends_with_bullet = False
    last_text = ""
    if last_non_empty is not None:
        paragraph, raw, last_text = last_non_empty
        if not has_faq:
            ends_with_bullet = is_word_bullet_paragraph(paragraph, raw)

    return {
        "h2_count": h2_count,
        "has_faq": has_faq,
        "ends_with_bullet": ends_with_bullet,
        "last_text": last_text,
    }


def get_visible_elements(
    driver,
    css_selector: str,
):
    return [
        element
        for element in driver.find_elements(
            By.CSS_SELECTOR,
            css_selector,
        )
        if element.is_displayed()
    ]


def find_add_faq_button(driver):
    """
    Tìm nút có nội dung '+ Thêm Hỏi đáp'.
    Không dùng tọa độ, chuột thật hoặc ảnh mẫu.
    """
    return driver.execute_script(
        """
        const elements = Array.from(
            document.querySelectorAll(
                'button, a, input[type="button"], input[type="submit"]'
            )
        );

        function normalize(value) {
            return String(value || '')
                .replace(/\\s+/g, ' ')
                .trim()
                .toLowerCase();
        }

        for (const element of elements) {
            const text = normalize(
                element.innerText
                || element.textContent
                || element.value
            );

            if (
                text.includes('thêm hỏi đáp')
                || text.includes('them hoi dap')
            ) {
                const rect = element.getBoundingClientRect();
                const style = window.getComputedStyle(element);

                if (
                    rect.width > 0
                    && rect.height > 0
                    && style.display !== 'none'
                    && style.visibility !== 'hidden'
                ) {
                    return element;
                }
            }
        }

        return null;
        """
    )


def click_add_faq(driver) -> None:
    button = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: find_add_faq_button(d)
    )

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block: 'center'
        });
        """,
        button,
    )

    time.sleep(0.3)

    try:
        button.click()
    except Exception:
        driver.execute_script(
            "arguments[0].click();",
            button,
        )


def set_element_value(
    driver,
    element,
    value: str,
) -> None:
    driver.execute_script(
        """
        const element = arguments[0];
        const value = arguments[1];

        element.scrollIntoView({
            block: 'center'
        });
        element.focus();
        element.value = value;

        element.dispatchEvent(
            new Event('input', {bubbles: true})
        );
        element.dispatchEvent(
            new Event('change', {bubbles: true})
        );
        element.dispatchEvent(
            new Event('blur', {bubbles: true})
        );
        """,
        element,
        value,
    )


def create_all_faq_blocks(
    driver,
    count: int,
) -> tuple[int, int]:
    """
    Tạo đủ toàn bộ block FAQ trước, chưa nhập dữ liệu.

    Trả về số ô Tiêu đề và Mô tả đã có trước khi tạo thêm,
    để bước dán theo lô chỉ ghi đúng các block vừa tạo.
    """
    if count <= 0:
        return 0, 0

    title_selector = (
        'input.form-control[placeholder="Tiêu đề"]'
    )
    answer_selector = (
        'textarea.form-control.margin-top-10[placeholder="Mô tả"], '
        'textarea.form-control[placeholder="Mô tả"]'
    )

    old_title_count = len(
        get_visible_elements(
            driver,
            title_selector,
        )
    )
    old_answer_count = len(
        get_visible_elements(
            driver,
            answer_selector,
        )
    )

    button = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: find_add_faq_button(d)
    )

    driver.execute_script(
        """
        arguments[0].scrollIntoView({
            block: 'center'
        });
        """,
        button,
    )

    # Click đủ số lần trong một lượt JavaScript.
    result = driver.execute_async_script(
        """
        const button = arguments[0];
        const count = arguments[1];
        const done = arguments[arguments.length - 1];

        let clicked = 0;

        function clickNext() {
            if (clicked >= count) {
                done({ok: true, clicked: clicked});
                return;
            }

            try {
                button.click();
                clicked += 1;

                // Nhường một nhịp ngắn để CMS dựng block mới.
                setTimeout(clickNext, 40);
            } catch (error) {
                done({
                    ok: false,
                    clicked: clicked,
                    error: String(error)
                });
            }
        }

        clickNext();
        """,
        button,
        count,
    )

    if not result or not result.get("ok"):
        error_text = (
            result.get("error")
            if isinstance(result, dict)
            else str(result)
        )
        raise RuntimeError(
            "Không tạo đủ block FAQ: " + error_text
        )

    expected_titles = old_title_count + count
    expected_answers = old_answer_count + count

    WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: (
            len(get_visible_elements(d, title_selector))
            >= expected_titles
            and len(get_visible_elements(d, answer_selector))
            >= expected_answers
        )
    )

    print(
        f"Đã tạo đủ {count} block FAQ."
    )

    return old_title_count, old_answer_count


def paste_all_faqs_once(
    driver,
    faq_pairs: list[dict[str, str]],
    title_start_index: int,
    answer_start_index: int,
) -> int:
    """
    Dán toàn bộ câu hỏi và câu trả lời bằng một lần JavaScript.
    """
    result = driver.execute_script(
        """
        const faqPairs = arguments[0];
        const titleStart = arguments[1];
        const answerStart = arguments[2];

        function isVisible(element) {
            if (!element) return false;

            const rect = element.getBoundingClientRect();
            const style = window.getComputedStyle(element);

            return (
                rect.width > 0
                && rect.height > 0
                && style.display !== 'none'
                && style.visibility !== 'hidden'
            );
        }

        function setValue(element, value) {
            element.focus();
            element.value = value;

            element.dispatchEvent(
                new Event('input', {bubbles: true})
            );
            element.dispatchEvent(
                new Event('change', {bubbles: true})
            );
            element.dispatchEvent(
                new Event('blur', {bubbles: true})
            );
        }

        const titles = Array.from(
            document.querySelectorAll(
                'input.form-control[placeholder="Tiêu đề"]'
            )
        ).filter(isVisible);

        const answers = Array.from(
            document.querySelectorAll(
                'textarea.form-control.margin-top-10[placeholder="Mô tả"], '
                + 'textarea.form-control[placeholder="Mô tả"]'
            )
        ).filter(isVisible);

        if (
            titles.length < titleStart + faqPairs.length
            || answers.length < answerStart + faqPairs.length
        ) {
            return {
                ok: false,
                error: 'Số ô FAQ hiện có không đủ',
                titleCount: titles.length,
                answerCount: answers.length
            };
        }

        for (let index = 0; index < faqPairs.length; index += 1) {
            setValue(
                titles[titleStart + index],
                faqPairs[index].question || ''
            );

            setValue(
                answers[answerStart + index],
                faqPairs[index].answer || ''
            );
        }

        return {
            ok: true,
            count: faqPairs.length
        };
        """,
        faq_pairs,
        title_start_index,
        answer_start_index,
    )

    if not result or not result.get("ok"):
        error_text = (
            result.get("error")
            if isinstance(result, dict)
            else str(result)
        )
        raise RuntimeError(
            "Không dán được FAQ theo lô: " + error_text
        )

    count = int(result.get("count", 0))
    print(
        f"Đã dán một lượt {count} cặp FAQ."
    )
    return count


def fill_faqs(
    driver,
    faq_pairs: list[dict[str, str]],
) -> dict[str, int | str]:
    print("[7] Tạo đủ block và dán FAQ một lượt...")

    if not faq_pairs:
        return {
            "status": "SKIP",
            "count": 0,
        }

    old_title_count, old_answer_count = (
        create_all_faq_blocks(
            driver,
            len(faq_pairs),
        )
    )

    inserted_count = paste_all_faqs_once(
        driver,
        faq_pairs,
        old_title_count,
        old_answer_count,
    )

    return {
        "status": "OK",
        "count": inserted_count,
    }

def click_js(driver, element) -> None:
    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        element,
    )

    try:
        element.click()
    except Exception:
        driver.execute_script(
            "arguments[0].click();",
            element,
        )


def find_visible_by_text(
    driver,
    selector: str,
    wanted_text: str,
):
    wanted = clean_text(wanted_text).lower()

    for element in driver.find_elements(
        By.CSS_SELECTOR,
        selector,
    ):
        try:
            if not element.is_displayed():
                continue

            hay = " ".join(
                clean_text(value).lower()
                for value in (
                    element.text,
                    element.get_attribute("value"),
                    element.get_attribute("title"),
                    element.get_attribute("aria-label"),
                )
                if value
            )

            if wanted in hay:
                return element

        except Exception:
            continue

    return None


# ============================================================
# WORD / EXCEL / ẢNH
# ============================================================





def find_numbered_image(
    stem: str,
    number: int,
    image_dir: Path | None = None,
) -> Path:
    image_dir = image_dir or IMAGE_DIR
    candidates: list[Path] = []

    for base_name in (
        f"{stem} {number}",
        f"{stem}_{number}",
        f"{stem}{number}",
    ):
        for extension in IMAGE_EXTENSIONS:
            candidates.append(
                image_dir / f"{base_name}{extension}"
            )

    for path in candidates:
        if path.is_file():
            return path.resolve()

    raise RuntimeError(
        f"Không tìm thấy ảnh số {number}:\n"
        + "\n".join(str(path) for path in candidates)
    )


def get_post_url(current_row: int) -> tuple[int, str]:
    """Lấy URL đăng bài từ đúng dòng được truyền vào, không qua file INI."""

    excel, workbook = get_target_excel_workbook()
    post_sheet = workbook.Worksheets(SHEET_POST)
    domain_sheet = workbook.Worksheets(SHEET_DOMAIN)

    domain_col = find_column_by_header(
        post_sheet,
        "Tên miền",
    )

    domain = normalize_domain(
        post_sheet.Cells(
            current_row,
            domain_col,
        ).Value
    )

    if not domain:
        raise RuntimeError(
            f'Dòng {current_row} chưa có "Tên miền"'
        )

    last_row = int(
        domain_sheet.Cells(
            domain_sheet.Rows.Count,
            1,
        ).End(-4162).Row
    )

    for row in range(2, last_row + 1):
        row_domain = normalize_domain(
            domain_sheet.Cells(row, 1).Value
        )

        if row_domain == domain:
            url = clean_text(
                domain_sheet.Cells(row, 2).Value
            )

            if url:
                return current_row, url

    raise RuntimeError(
        f"Không tìm thấy URL cho domain: {domain}"
    )


# ============================================================
# EDGE: MỞ BÌNH THƯỜNG, ẨN TRONG LÚC CHẠY, XONG THÌ HIỆN LẠI
# ============================================================

def _find_edge_hwnd_by_marker(marker: str) -> int | None:
    matched: list[int] = []

    def enum_callback(hwnd: int, _extra: object) -> None:
        try:
            if not win32gui.IsWindow(hwnd):
                return

            title = win32gui.GetWindowText(hwnd)
            if marker in title:
                matched.append(hwnd)
        except Exception:
            pass

    win32gui.EnumWindows(enum_callback, None)
    return matched[0] if matched else None


def remember_edge_window(driver) -> int:
    """Đánh dấu title tạm để tìm đúng cửa sổ Edge của Selenium."""
    global _EDGE_HWND

    marker = f"SELENIUM_EDGE_{int(time.time() * 1000)}"
    original_title = str(
        driver.execute_script("return document.title || '';") or ""
    )

    driver.execute_script(
        "document.title = arguments[0];",
        marker,
    )

    deadline = time.time() + 8
    hwnd = None

    while time.time() < deadline:
        hwnd = _find_edge_hwnd_by_marker(marker)
        if hwnd:
            break
        time.sleep(0.1)

    driver.execute_script(
        "document.title = arguments[0];",
        original_title,
    )

    if not hwnd:
        raise RuntimeError(
            "Không xác định được cửa sổ Edge để ẩn."
        )

    _EDGE_HWND = hwnd
    return hwnd


def hide_edge_window(driver) -> None:
    global _EDGE_IS_HIDDEN

    hwnd = _EDGE_HWND or remember_edge_window(driver)
    win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
    _EDGE_IS_HIDDEN = True
    print("    Edge đã được ẩn trong lúc chạy.")


def show_edge_window() -> None:
    global _EDGE_IS_HIDDEN

    hwnd = _EDGE_HWND
    if not hwnd or not win32gui.IsWindow(hwnd):
        return

    try:
        # Đưa cửa sổ từ ngoài màn hình về màn hình chính trước khi hiện.
        win32gui.SetWindowPos(
            hwnd,
            win32con.HWND_TOP,
            0,
            0,
            1200,
            900,
            win32con.SWP_SHOWWINDOW,
        )
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        # Có máy Windows chặn SetForegroundWindow; cửa sổ vẫn được hiện.
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        except Exception:
            pass

    _EDGE_IS_HIDDEN = False
    print("    Edge đã hiện lại để kiểm tra.")


def set_edge_window_visible(visible: bool) -> None:
    """Hiện/ẩn cửa sổ Edge hiện tại theo lựa chọn trên bảng tiến độ."""
    global _EDGE_IS_HIDDEN

    hwnd = _EDGE_HWND
    if not hwnd or not win32gui.IsWindow(hwnd):
        return

    if visible:
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            # Windows đôi khi chặn ứng dụng nền giành focus; Edge vẫn được hiện.
            pass
        _EDGE_IS_HIDDEN = False
        print("    Edge đã được hiện theo lựa chọn trên bảng tiến độ.")
    else:
        win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
        _EDGE_IS_HIDDEN = True
        print("    Edge đã được ẩn theo lựa chọn trên bảng tiến độ.")


def close_existing_selenium_edge() -> None:
    """
    Đóng Edge cũ đang dùng đúng profile SeleniumData trước khi tạo session mới.

    Chỉ nhắm tới msedge.exe có CommandLine chứa đường dẫn EDGE_USER_DATA_DIR,
    không đóng Edge cá nhân dùng profile mặc định.
    """
    global _EDGE_HWND, _EDGE_IS_HIDDEN

    profile_path = str(EDGE_USER_DATA_DIR.resolve())
    escaped_profile = profile_path.replace("'", "''")

    powershell_script = f"""
    $profile = '{escaped_profile}'
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

    result = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            powershell_script,
        ],
        check=False,
        capture_output=True,
        text=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )

    if result.returncode != 0:
        error_text = clean_text(result.stderr)
        raise RuntimeError(
            "Không kiểm tra/đóng được Edge Selenium cũ."
            + (f" Chi tiết: {error_text}" if error_text else "")
        )

    # Chờ Windows giải phóng khóa profile trước khi EdgeDriver mở lại.
    time.sleep(1.5)
    _EDGE_HWND = None
    _EDGE_IS_HIDDEN = False
    print("    Đã đóng Edge Selenium cũ nếu đang mở.")


def open_edge(url: str):
    """
    Mở một Edge automation riêng:
    - Không đóng Edge cá nhân.
    - Không dùng profile SeleniumData.
    - Giữ cửa sổ hiện để kiểm tra.
    - Dùng lại cùng driver cho toàn bộ các bài.
    """
    print("[2] Mở Edge automation riêng, không tắt trình duyệt đang có...")

    options = Options()
    options.add_argument("--disable-notifications")
    options.add_argument("--start-maximized")

    # Giữ cửa sổ Edge tồn tại độc lập nếu Python dừng do lỗi.
    # Khi chạy thành công, code vẫn chủ động gọi driver.quit().
    options.add_experimental_option("detach", True)

    driver = webdriver.Edge(options=options)
    driver.maximize_window()
    driver.get(url)
    wait_document_ready(driver)

    print(f"    URL hiện tại: {driver.current_url}")
    return driver



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


def is_forbidden_page(driver) -> bool:
    """Nhận diện trang 403 do web server/WAF trả về trước khi CMS tải."""
    try:
        title = clean_text(driver.title).casefold()
        body = clean_text(
            driver.execute_script(
                "return document.body ? document.body.innerText : '';"
            )
        ).casefold()
        return (
            "403 forbidden" in title
            or (
                "403" in body
                and (
                    "forbidden" in body
                    or "access to this resource" in body
                )
            )
        )
    except Exception:
        return False


def is_not_acceptable_page(driver) -> bool:
    """Nhận diện trang 406 nginx/WAF qua title và nội dung hiển thị."""
    try:
        title = clean_text(driver.title).casefold()
        body = clean_text(
            driver.execute_script(
                "return document.body ? document.body.innerText : '';"
            )
        ).casefold()
        return (
            "406 not acceptable" in title
            or "406 not acceptable" in body
        )
    except Exception:
        return False


def wait_document_ready(driver) -> None:
    WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )


def read_login_credentials() -> tuple[str, str]:
    if not LOGIN_INI.is_file():
        raise RuntimeError(
            "Không tìm thấy file tài khoản:\n"
            f"{LOGIN_INI}\n\n"
            "Ví dụ nội dung:\n"
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
    driver.switch_to.default_content()
    wait_document_ready(driver)

    if is_not_acceptable_page(driver):
        raise Http406Error(
            "HTTP 406 Not Acceptable: website/WAF từ chối trang đăng bài."
        )

    if is_forbidden_page(driver):
        raise RuntimeError(
            "Website trả về 403 Forbidden và từ chối truy cập trang đăng bài. "
            "Chương trình dừng, không thử đăng nhập lại."
        )

    if is_post_edit_page(driver):
        print("    Đã ở đúng trang đăng bài, không cần đăng nhập.")
        return

    if not is_login_page(driver):
        raise RuntimeError(
            "Không thấy trang đăng bài hoặc form đăng nhập.\n"
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

    print("    Đã nhập đủ ID/mật khẩu → click ĐĂNG NHẬP một lần.")

    try:
        login_button.click()
    except Exception:
        driver.execute_script("arguments[0].click();", login_button)

    try:
        WebDriverWait(driver, WAIT_PAGE).until(
            lambda d: (
                is_not_acceptable_page(d)
                or is_forbidden_page(d)
                or (not is_login_page(d))
                or is_post_edit_page(d)
            )
        )
    except TimeoutException as exc:
        raise RuntimeError(
            "Đăng nhập không thành công sau một lần thử. "
            "Website vẫn ở trang đăng nhập hoặc không phản hồi; chương trình dừng."
        ) from exc

    wait_document_ready(driver)

    if is_not_acceptable_page(driver):
        raise Http406Error(
            "HTTP 406 Not Acceptable sau đăng nhập: website/WAF từ chối request."
        )

    if is_forbidden_page(driver):
        raise RuntimeError(
            "Đăng nhập không thành công: website trả về 403 Forbidden. "
            "Chương trình dừng và không thử đăng nhập lại."
        )

    if not is_post_edit_page(driver):
        driver.get(target_url)
        wait_document_ready(driver)

    if is_not_acceptable_page(driver):
        raise Http406Error(
            "HTTP 406 Not Acceptable khi mở trang đăng bài sau đăng nhập."
        )

    if is_forbidden_page(driver):
        raise RuntimeError(
            "Đăng nhập xong nhưng website trả về 403 Forbidden khi mở trang đăng bài. "
            "Chương trình dừng và không thử lại."
        )

    if is_login_page(driver):
        raise RuntimeError(
            "Đăng nhập không thành công sau một lần thử, website vẫn trả về trang login. "
            "Chương trình dừng; hãy kiểm tra ID/mật khẩu trong file INI."
        )

    WebDriverWait(driver, WAIT_PAGE).until(lambda d: is_post_edit_page(d))
    print("    Đăng nhập thành công, đã vào đúng trang đăng bài.")


# ============================================================
# ẢNH ĐẠI DIỆN
# ============================================================

_ensure_post_page_ready_unlocked = ensure_post_page_ready


def ensure_post_page_ready(driver, target_url: str) -> None:
    """Only serialize the actual login action; opening an already-ready CMS is free."""
    driver.switch_to.default_content()
    wait_document_ready(driver)
    if is_post_edit_page(driver):
        print("    Đã ở đúng trang đăng bài, không cần đăng nhập.")
        return

    if LOGIN_LOCK is not None and is_login_page(driver):
        print("    Phát hiện trang đăng nhập → chờ khóa Login ngắn...")
        with LOGIN_LOCK:
            # Re-check after obtaining the lock; do not serialize workers that
            # already reached the post form while another login was running.
            if is_post_edit_page(driver):
                print("    Đã ở đúng trang đăng bài, không cần đăng nhập.")
                return
            _ensure_post_page_ready_unlocked(driver, target_url)
        return

    _ensure_post_page_ready_unlocked(driver, target_url)


def save_and_confirm_article(driver, titles: list[str]) -> tuple[str, str, str, str]:
    """Serialize Save + CMS ID lookup only for the current domain."""
    def operation() -> tuple[str, str, str, str]:
        saved_url = save_article(driver)
        cms_id, edit_url, public_url = find_saved_article_info(driver, titles)
        return saved_url, cms_id, edit_url, public_url

    if SAVE_DOMAIN_LOCK is None:
        return operation()

    print("    Đang chờ khóa Save/Xác nhận ID của đúng tên miền...")
    with SAVE_DOMAIN_LOCK:
        print("    Đã đến lượt Save/Xác nhận ID của tên miền này.")
        return operation()


def upload_featured_image(
    driver,
    image_path: Path,
) -> None:
    print("[3] Upload ảnh đại diện...")

    if not image_path.is_file():
        raise RuntimeError(f"Không tìm thấy ảnh đại diện:\n{image_path}")

    driver.switch_to.default_content()

    btn = WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: d.find_element(
            By.CSS_SELECTOR,
            "span.input-group-addon.btn.btn-file",
        )
    )

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        btn,
    )
    driver.execute_script("arguments[0].click();", btn)

    file_input = btn.find_element(
        By.CSS_SELECTOR,
        "input[type='file'][name='thumbnail']",
    )

    driver.execute_script(
        """
        arguments[0].style.display = 'block';
        arguments[0].style.visibility = 'visible';
        arguments[0].style.opacity = '1';
        arguments[0].removeAttribute('hidden');
        """,
        file_input,
    )

    file_input.send_keys(str(image_path.resolve()))

    # CMS tự xử lý upload qua onchange của input.
    # Chuyển ngay sang bước upload ảnh 2, không chờ xác nhận.
    print("    OK: đã gửi ảnh đại diện")


# ============================================================
# ẢNH CKEDITOR
# ============================================================

def open_image_dialog(driver) -> None:
    print("[4] Mở popup ảnh CKEditor...")

    icon = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: d.find_element(
            By.CSS_SELECTOR,
            "span.cke_button_icon.cke_button__image_icon",
        )
    )

    click_js(driver, icon)

    WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: any(
            element.is_displayed()
            for element in d.find_elements(
                By.CSS_SELECTOR,
                ".cke_dialog",
            )
        )
    )


def click_upload_tab(driver) -> None:
    print("[5] Click tab Tải lên...")

    tab = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: find_visible_by_text(
            d,
            ".cke_dialog_tab, a.cke_dialog_tab",
            "Tải lên",
        )
    )

    click_js(driver, tab)

    WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: d.find_elements(
            By.CSS_SELECTOR,
            "iframe.cke_dialog_ui_input_file",
        )
    )


def select_content_image_file(
    driver,
    image_path: Path,
) -> None:
    print("[6] Chọn ảnh bài viết...")
    print(f"    {image_path}")

    upload_iframe = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: next(
            (
                frame
                for frame in d.find_elements(
                    By.CSS_SELECTOR,
                    "iframe.cke_dialog_ui_input_file",
                )
                if frame.is_displayed()
            ),
            None,
        )
    )

    driver.switch_to.frame(upload_iframe)

    try:
        file_input = WebDriverWait(
            driver,
            WAIT_PAGE,
        ).until(
            lambda d: d.find_element(
                By.CSS_SELECTOR,
                'input[type="file"]',
            )
        )

        file_input.send_keys(str(image_path))

    finally:
        driver.switch_to.default_content()

    print("    OK: đã chọn file trong popup")


def click_upload_to_server(driver) -> None:
    print("[7] Click Tải lên máy chủ...")

    button = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: find_visible_by_text(
            d,
            (
                ".cke_dialog_ui_button, "
                "a.cke_dialog_ui_button, "
                "button, input[type='button']"
            ),
            "Tải lên máy chủ",
        )
    )

    click_js(driver, button)


def wait_for_uploaded_url(driver) -> str:
    print("[8] Chờ server trả URL ảnh...")

    def get_url(d):
        try:
            return d.execute_script(
                """
                try {
                    const dialog =
                        CKEDITOR.dialog.getCurrent();

                    if (!dialog) return '';

                    const field =
                        dialog.getContentElement(
                            'info',
                            'txtUrl'
                        );

                    return field
                        ? (field.getValue() || '')
                        : '';
                } catch (e) {
                    return '';
                }
                """
            )
        except Exception:
            return ""

    url = WebDriverWait(
        driver,
        WAIT_UPLOAD,
    ).until(
        lambda d: get_url(d) or False
    )

    print(f"    URL: {url}")
    return str(url)


def click_dialog_ok(driver) -> None:
    print("[9] Click Đồng ý...")

    button = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: find_visible_by_text(
            d,
            (
                ".cke_dialog_ui_button, "
                "a.cke_dialog_ui_button, "
                "button"
            ),
            "Đồng ý",
        )
    )

    click_js(driver, button)

    WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: not any(
            element.is_displayed()
            for element in d.find_elements(
                By.CSS_SELECTOR,
                ".cke_dialog",
            )
        )
    )

    print("    OK: ảnh đã được chèn vào CKEditor")


def upload_content_image(
    driver,
    image_path: Path,
) -> None:
    open_image_dialog(driver)
    click_upload_tab(driver)
    select_content_image_file(
        driver,
        image_path,
    )
    click_upload_to_server(driver)
    wait_for_uploaded_url(driver)
    click_dialog_ok(driver)


# ============================================================
# BƯỚC 13: LẤY, LÀM SẠCH VÀ GOM HTML TRỰC TIẾP TRONG CKEDITOR
# ============================================================

def get_ckeditor_detail_html(driver) -> str:
    html = WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: d.execute_script(
            """
            if (
                typeof CKEDITOR === 'undefined'
                || !CKEDITOR.instances
                || !CKEDITOR.instances.detail
                || CKEDITOR.instances.detail.status !== 'ready'
            ) return null;
            return CKEDITOR.instances.detail.getData();
            """
        )
    )
    html = str(html or "")
    if not html.strip():
        raise RuntimeError("CKEditor detail đang rỗng, không có HTML để xử lý.")
    return html


def _strip_direct_strong_from_heading(match: re.Match[str]) -> str:
    tag = match.group(1).lower()
    attrs = match.group(2) or ""
    inner = match.group(3) or ""
    direct = re.fullmatch(
        r"(?is)\s*<strong(?:\s[^>]*)?>(.*?)</strong>\s*",
        inner,
    )
    if direct:
        inner = direct.group(1)
    return f"<{tag}{attrs}>{inner}</{tag}>"


def _normalize_heading(match: re.Match[str]) -> str:
    tag = match.group(1).lower()
    inner = (match.group(2) or "").strip()
    return f"<{tag}><strong>{inner}</strong></{tag}>"


def _move_images_out_of_h2(match: re.Match[str]) -> str:
    inner = match.group(1) or ""
    images = re.findall(r"(?is)<img\b[^>]*>", inner)
    if not images:
        return match.group(0)
    clean_inner = re.sub(r"(?is)<img\b[^>]*>", "", inner).strip()
    return "\n".join(images) + f"\n<h2>{clean_inner}</h2>"


def _remove_image_only_wrappers(html: str) -> str:
    pattern = (
        r"(?is)<(?:p|div)\b[^>]*>\s*"
        r"(<img\b[^>]*>)\s*(?:<br\s*/?>\s*)?"
        r"</(?:p|div)>"
    )
    previous = None
    while previous != html:
        previous = html
        html = re.sub(pattern, r"\1", html)
    return html


def _extract_centered_images(html: str) -> tuple[str, list[str]]:
    images: list[str] = []
    pattern = re.compile(
        r"(?is)<p\s+style=[\"']text-align\s*:\s*center\s*;?[\"']\s*>"
        r"\s*(<img\b[^>]*>)\s*</p>"
    )
    def collect(match: re.Match[str]) -> str:
        images.append(f'<p style="text-align:center;">{match.group(1)}</p>')
        return "\n"
    return pattern.sub(collect, html), images


def _insert_images_before_h2_from_second(html: str, images: list[str]) -> str:
    if not images:
        return html
    h2_count = 0
    image_index = 0
    def inject(match: re.Match[str]) -> str:
        nonlocal h2_count, image_index
        h2_count += 1
        if h2_count >= 2 and image_index < len(images):
            block = images[image_index]
            image_index += 1
            return f"{block}\n{match.group(0)}"
        return match.group(0)
    output = re.sub(r"(?is)<h2\b[^>]*>.*?</h2>", inject, html)
    if image_index < len(images):
        output = output.rstrip() + "\n" + "\n".join(images[image_index:])
    return output


def clean_article_html(html: str) -> str:
    original = str(html or "")
    if not original.strip():
        raise RuntimeError("HTML đầu vào bị rỗng.")

    # H2/H3: bỏ strong cũ bao trọn rồi chuẩn hóa lại.
    original = re.sub(
        r"(?is)<(h2|h3)([^>]*)>(.*?)</\1>",
        _strip_direct_strong_from_heading,
        original,
    )
    original = re.sub(
        r"(?is)<(h2|h3)\b[^>]*>\s*(.*?)\s*</\1>",
        _normalize_heading,
        original,
    )

    # Bỏ span rác nhưng giữ nội dung.
    original = re.sub(r"(?is)<span\b[^>]*>", "", original)
    original = re.sub(r"(?is)</span\s*>", "", original)

    # Bỏ anchor Word dạng name=... nhưng giữ nội dung.
    original = re.sub(
        r"(?is)<a\s+[^>]*\bname\s*=\s*([\"']).*?\1[^>]*>(.*?)</a>",
        r"\2",
        original,
    )

    # Chuẩn hóa bảng.
    original = re.sub(
        r"(?is)<table\b[^>]*>",
        '<table border="1" cellpadding="5" cellspacing="5" '
        'style="border-collapse:collapse; width:100%; text-align:center">',
        original,
    )

    # Tách ảnh khỏi H2.
    original = re.sub(
        r"(?is)<h2\b[^>]*>(.*?)</h2>",
        _move_images_out_of_h2,
        original,
    )
    original = re.sub(
        r"(?is)<(h2|h3)\b[^>]*>\s*(.*?)\s*</\1>",
        _normalize_heading,
        original,
    )

    # Chuẩn hóa wrapper ảnh rồi căn giữa ảnh như AHK cũ.
    original = _remove_image_only_wrappers(original)
    original = re.sub(
        r"(?is)(<img\b[^>]*>)",
        r'<p style="text-align:center;">\1</p>',
        original,
    )

    # Xóa H2 rỗng.
    original = re.sub(
        r"(?is)<h2\b[^>]*>\s*(?:<strong\b[^>]*>)?"
        r"(?:&nbsp;|\s)*(?:</strong>)?\s*</h2>",
        "",
        original,
    )

    # Gom ảnh: ảnh đầu trước H2 thứ 2, ảnh tiếp theo trước H2 sau.
    original, image_blocks = _extract_centered_images(original)
    output = _insert_images_before_h2_from_second(original, image_blocks)

    # Xóa thẻ rỗng và rác.
    output = re.sub(
        r"(?is)<p\b[^>]*>(?:\s|&nbsp;|<br\s*/?>)*</p>",
        "",
        output,
    )
    output = re.sub(
        r"(?is)<div\b[^>]*>(?:\s|&nbsp;|<br\s*/?>)*</div>",
        "",
        output,
    )
    output = re.sub(r"(?is)<(?:p|div)\b[^>]*/>", "", output)
    output = re.sub(r"(?im)^\s*&nbsp;\s*$", "", output)
    output = re.sub(
        r"(?is)<div\b[^>]*(?:class|id)\s*=\s*([\"'])[^\"']*"
        r"simple-translate[^\"']*\1[^>]*>.*?</div>",
        "",
        output,
    )
    output = re.sub(r"(?m)^\s*$\n?", "", output).strip()

    if not output:
        raise RuntimeError("HTML bị rỗng sau khi làm sạch.")
    return output


def set_ckeditor_detail_html(driver, html_content: str) -> dict[str, Any]:
    result = driver.execute_async_script(
        """
        const html = arguments[0];
        const done = arguments[arguments.length - 1];
        try {
            if (
                typeof CKEDITOR === 'undefined'
                || !CKEDITOR.instances
                || !CKEDITOR.instances.detail
            ) {
                done({ok:false, error:'Không tìm thấy CKEDITOR.instances.detail'});
                return;
            }
            const editor = CKEDITOR.instances.detail;
            editor.setData(html, {
                callback: function () {
                    editor.fire('change');
                    editor.updateElement();
                    done({ok:true, length:editor.getData().length});
                }
            });
        } catch (error) {
            done({ok:false, error:String(error)});
        }
        """,
        html_content,
    )
    if not result or not result.get("ok"):
        error = result.get("error") if isinstance(result, dict) else str(result)
        raise RuntimeError("Không trả được HTML đã gom vào CKEditor: " + error)
    return result


def clean_and_reorganize_ckeditor_html(driver) -> dict[str, int]:
    print("[13] Lấy và gom HTML trực tiếp trong CKEditor...")
    driver.switch_to.default_content()
    html_before = get_ckeditor_detail_html(driver)
    html_after = clean_article_html(html_before)
    set_ckeditor_detail_html(driver, html_after)
    confirmed = get_ckeditor_detail_html(driver)
    stats = {
        "before": len(html_before),
        "after": len(confirmed),
        "images": len(re.findall(r"(?is)<img\b[^>]*>", confirmed)),
        "h2": len(re.findall(r"(?is)<h2\b[^>]*>", confirmed)),
        "h3": len(re.findall(r"(?is)<h3\b[^>]*>", confirmed)),
    }
    print("    OK: đã gom HTML bằng Python")
    print(f"    HTML trước : {stats['before']} ký tự")
    print(f"    HTML sau   : {stats['after']} ký tự")
    print(f"    Số ảnh     : {stats['images']}")
    print(f"    Số H2/H3   : {stats['h2']}/{stats['h3']}")
    return stats


def get_pre_save_expected_values(row: int) -> dict[str, str]:
    """Đọc các giá trị chuẩn của đúng dòng Excel ngay trước Save."""
    _excel, workbook = get_target_excel_workbook()
    sheet = workbook.Worksheets(SHEET_POST)

    def read(header_key: str) -> str:
        column = find_column_by_header(
            sheet,
            PUBLISH_HEADERS[header_key],
        )
        return clean_text(sheet.Cells(row, column).Value)

    return {
        "domain": read("domain"),
        "category": read("category"),
        "title": read("title"),
        "h1": read("h1"),
        "keyword": read("keyword"),
        "image1_path": read("image1_path"),
        "image2_path": read("image2_path"),
    }


def normalize_validation_text(value: Any) -> str:
    return unicodedata.normalize("NFC", clean_text(value))


def read_cms_values_before_save(driver) -> dict[str, str]:
    result = driver.execute_script(
        """
        const valueOf = id => {
            const element = document.getElementById(id);
            return element ? String(element.value || '') : '';
        };
        const category = document.querySelector('select[name="cat_id"]');
        const selectedCategory = category
            ? category.options[category.selectedIndex]
            : null;
        return {
            title: valueOf('title'),
            h1: valueOf('name'),
            keyword: valueOf('keyword'),
            category: selectedCategory
                ? String(selectedCategory.textContent || '').trim()
                : ''
        };
        """
    )
    return {
        key: normalize_validation_text((result or {}).get(key, ""))
        for key in ("title", "h1", "keyword", "category")
    }


def _resolved_path_text(value: Any) -> str:
    raw = clean_text(value).strip('"')
    if not raw:
        return ""
    path = Path(
        os.path.expandvars(
            os.path.expanduser(raw)
        )
    ).resolve()
    return os.path.normcase(str(path))


def validate_before_save(
    driver,
    current_row: int,
    word_info: dict[str, Any],
    image_paths: dict[str, Any],
    word_structure: dict[str, Any],
    html_result: dict[str, int],
) -> None:
    """Cổng bảo hiểm cuối cùng: lỗi thì tuyệt đối không click Save bài đó."""
    print("[14] Kiểm tra bảo hiểm ngay trước khi Save...")
    expected = get_pre_save_expected_values(current_row)
    actual = read_cms_values_before_save(driver)
    errors: list[str] = []

    required_excel = {
        "Tên miền": expected["domain"],
        "Danh mục": expected["category"],
        "Tiêu đề SEO": expected["title"],
        "H1": expected["h1"],
        'Keyword (cột Excel "Tiêu đề")': expected["keyword"],
    }
    for label, value in required_excel.items():
        if not normalize_validation_text(value):
            errors.append(f'Excel thiếu "{label}"')

    comparisons = (
        ("Title", actual["title"], expected["title"]),
        ("H1", actual["h1"], expected["h1"]),
        (
            'Keyword (cột Excel "Tiêu đề")',
            actual["keyword"],
            expected["keyword"],
        ),
        ("Danh mục", actual["category"], expected["category"]),
    )
    for label, cms_value, excel_value in comparisons:
        # Không phân biệt chữ hoa/chữ thường khi đối chiếu CMS với Excel.
        cms_normalized = normalize_validation_text(cms_value).casefold()
        excel_normalized = normalize_validation_text(excel_value).casefold()
        if cms_normalized != excel_normalized:
            errors.append(
                f'{label} CMS="{cms_normalized}" khác Excel="{excel_normalized}"'
            )

    word_h2_count = int(word_structure.get("h2_count", 0))
    cms_h2_count = int(html_result.get("h2", 0))
    if word_h2_count < 2:
        errors.append(
            f"Word chỉ có {word_h2_count} Heading 2 thật (cần ít nhất 2)"
        )
    if cms_h2_count < 2:
        errors.append(
            f"CMS chỉ còn {cms_h2_count} thẻ H2 sau khi dán (cần ít nhất 2)"
        )

    if (
        not bool(word_structure.get("has_faq"))
        and bool(word_structure.get("ends_with_bullet"))
    ):
        errors.append(
            "Bài không có FAQ nhưng đoạn cuối nội dung Word vẫn là bullet: "
            + clean_text(word_structure.get("last_text", ""))[:160]
        )

    word_stem = normalize_validation_text(word_info.get("stem", "")).casefold()
    for number, image_key, excel_key, label in (
        (1, "featured", "image1_path", "Ảnh đại diện"),
        (2, "content", "image2_path", "Ảnh nội dung"),
    ):
        raw_path = image_paths.get(image_key)
        image_path = Path(str(raw_path)).resolve() if raw_path else None
        if image_path is None or not image_path.is_file():
            errors.append(f"{label} không tồn tại: {raw_path}")
            continue

        image_stem = normalize_validation_text(image_path.stem).casefold()
        # Bỏ hậu tố lần/đợt (ví dụ: _dot_2, _lan_2, 1, 2) khi so tên ảnh.
        # Vì cùng một bài có thể được xuất lại và ảnh bị đổi hậu tố.
        def image_base(value: str) -> str:
            return re.sub(r"(?:[_\s-]*(?:dot|lan|lần)?[_\s-]*\d+)+$", "", value).strip(" _-")

        if image_base(image_stem) != image_base(word_stem):
            errors.append(
                f'{label} "{image_path.name}" không khớp file Word '
                f'"{word_info.get("name", "")}"'
            )

        excel_path = _resolved_path_text(expected.get(excel_key, ""))
        actual_path = os.path.normcase(str(image_path))
        if excel_path and actual_path != excel_path:
            errors.append(
                f"{label} thực tế khác đường dẫn Excel: "
                f'"{image_path}" != "{expected[excel_key]}"'
            )

    if errors:
        print("    KHÔNG SAVE BÀI NÀY:")
        for error in errors:
            print(f"    - {error}")
        raise PreSaveValidationError(
            "Kiểm tra trước Save không đạt: " + " | ".join(errors)
        )

    print(
        "    OK: Title, H1, Keyword, Danh mục, ảnh, H2 và bullet đều hợp lệ."
    )


# ============================================================
# BƯỚC 14: LƯU BÀI
# Nút xác nhận từ DOM:
# <button class="btn btn-sm green-jungle" type="submit">...</button>
# ============================================================

def save_article(driver) -> str:
    print("[14] Click nút lưu bài...")
    driver.switch_to.default_content()

    # Đồng bộ dữ liệu CKEditor về textarea trước khi submit form.
    driver.execute_script(
        """
        if (
            typeof CKEDITOR !== 'undefined'
            && CKEDITOR.instances
            && CKEDITOR.instances.detail
        ) {
            CKEDITOR.instances.detail.updateElement();
        }
        """
    )

    save_button = WebDriverWait(
        driver,
        WAIT_PAGE,
    ).until(
        lambda d: next(
            (
                button
                for button in d.find_elements(
                    By.CSS_SELECTOR,
                    'button.btn.btn-sm.green-jungle[type="submit"]',
                )
                if button.is_displayed() and button.is_enabled()
            ),
            None,
        )
    )

    driver.execute_script(
        "arguments[0].scrollIntoView({block:'center'});",
        save_button,
    )

    old_url = driver.current_url

    try:
        save_button.click()
    except Exception:
        driver.execute_script("arguments[0].click();", save_button)

    # Cho form có thời gian submit/chuyển trang, sau đó chờ trang ổn định.
    time.sleep(0.5)
    WebDriverWait(driver, WAIT_PAGE).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    time.sleep(0.5)

    new_url = driver.current_url
    if new_url != old_url:
        print(f"    OK: đã chuyển sang URL sau khi lưu: {new_url}")
    else:
        print(
            "    Đã click nút lưu nhưng URL không đổi. "
            "Edge sẽ hiện lại để kiểm tra thông báo của CMS."
        )

    return new_url


def extract_cms_id_from_url(url: str) -> str:
    """Tách ID bài từ URL sửa dạng /field/id=4077/root_id=999."""
    match = re.search(r"/field/id=(\d+)(?:/|$|[?#])", str(url or ""), re.IGNORECASE)
    return match.group(1) if match else ""


def find_saved_article_info(driver, titles: list[str]) -> tuple[str, str, str]:
    """
    Lấy ID bài vừa lưu:
    1) ưu tiên URL hiện tại nếu đang ở trang sửa;
    2) nếu CMS trả về danh sách, chỉ lấy link sửa trong hàng khớp tiêu đề.
    """
    current_url = str(driver.current_url or "")
    post_id = extract_cms_id_from_url(current_url)
    if post_id:
        return post_id, current_url, ""

    normalized_titles = [
        clean_text(title).casefold()
        for title in titles
        if clean_text(title)
    ]
    matches = driver.execute_script(
        """
        const titles = arguments[0];
        const clean = value => String(value || '')
            .replace(/\\s+/g, ' ')
            .trim()
            .toLocaleLowerCase('vi');
        const results = [];
        for (const link of document.querySelectorAll('a[href*="/field/id="]')) {
            const row = link.closest('tr');
            if (!row) continue;
            const rowText = clean(row.innerText || row.textContent);
            if (!titles.some(title => title && rowText.includes(title))) continue;
            results.push({
                href: link.href,
                text: rowText,
                links: Array.from(row.querySelectorAll('a[href]'))
                    .map(item => item.href)
            });
        }
        return results;
        """,
        normalized_titles,
    ) or []

    unique: dict[str, dict[str, Any]] = {}
    for item in matches:
        href = str(item.get("href", "") if isinstance(item, dict) else "")
        found_id = extract_cms_id_from_url(href)
        if found_id:
            unique[found_id] = {
                "edit_url": href,
                "links": list(item.get("links", [])),
            }

    if len(unique) != 1:
        raise RuntimeError(
            "Đã lưu bài nhưng không xác định chắc chắn được ID CMS trên trang danh sách. "
            f"Số ID khớp tiêu đề tìm thấy: {len(unique)}."
        )

    post_id, row_info = next(iter(unique.items()))
    edit_url = clean_text(row_info["edit_url"])

    public_candidates: list[str] = []
    public_pattern = re.compile(
        rf"-{re.escape(post_id)}\.html(?:$|[?#])",
        re.IGNORECASE,
    )
    for raw_url in row_info["links"]:
        candidate = clean_text(raw_url)
        lowered = candidate.lower()
        if (
            candidate
            and lowered.startswith(("http://", "https://"))
            and "/admin/" not in lowered
            and "/linkrutgon-" not in lowered
            and public_pattern.search(candidate)
            and candidate not in public_candidates
        ):
            public_candidates.append(candidate)

    public_url = public_candidates[0] if len(public_candidates) == 1 else ""
    if public_url:
        print(f"    URL công khai đọc từ hàng danh sách: {public_url}")
    else:
        print(
            "    Không thấy đúng một URL công khai chắc chắn trong hàng bài "
            "→ để trống URL đã đăng."
        )

    return post_id, edit_url, public_url


# ============================================================
# MAIN - HỎI SỐ BÀI, CHẠY LẶP; CHỈ HIỆN EDGE Ở BÀI CUỐI
# ============================================================

def ask_article_count() -> int:
    """Hiện hộp nhập số bài cần chạy trước khi bắt đầu."""
    try:
        import tkinter as tk
        from tkinter import simpledialog, messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        value = simpledialog.askinteger(
            "Số bài cần đăng",
            "Nhập số bài cần chạy:",
            parent=root,
            minvalue=1,
        )

        if value is None:
            root.destroy()
            raise RuntimeError("Bạn đã hủy nhập số bài.")

        if value < 1:
            messagebox.showerror(
                "Lỗi",
                "Số bài phải lớn hơn hoặc bằng 1.",
                parent=root,
            )
            root.destroy()
            raise RuntimeError("Số bài không hợp lệ.")

        root.destroy()
        return int(value)

    except RuntimeError:
        raise
    except Exception as exc:
        raise RuntimeError(
            "Không mở được hộp nhập số bài."
        ) from exc


def reset_edge_state() -> None:
    """Xóa handle Edge cũ sau khi đã đóng driver."""
    global _EDGE_HWND, _EDGE_IS_HIDDEN
    _EDGE_HWND = None
    _EDGE_IS_HIDDEN = False


def close_driver_safely(driver) -> None:
    """Đóng Edge khi chạy thành công; mọi lỗi cleanup đều được bỏ qua."""
    if driver is not None:
        try:
            driver.quit()
        except Exception as cleanup_error:
            print(
                "Bỏ qua lỗi cleanup Edge:",
                repr(cleanup_error),
            )

    try:
        reset_edge_state()
    except Exception as cleanup_error:
        print(
            "Bỏ qua lỗi reset trạng thái Edge:",
            repr(cleanup_error),
        )

    try:
        time.sleep(1.5)
    except Exception:
        pass


def close_word_safely() -> None:
    """Đóng Word nền; lỗi cleanup không được tính là lỗi đăng bài."""
    try:
        close_hidden_word_document()
    except Exception as cleanup_error:
        print(
            "Bỏ qua lỗi cleanup Word:",
            repr(cleanup_error),
        )


class PublishProgressWindow:
    """Bảng tiến độ độc lập; đóng bảng không làm dừng chương trình."""

    def __init__(self, total: int) -> None:
        self.total = total
        self.stop_after_current = threading.Event()
        # Mặc định không bật: Edge chạy ẩn.
        self.show_browser = threading.Event()
        if os.environ.get("VIP_SHOW_EDGE", "").strip() == "1":
            self.show_browser.set()
        self._messages: queue.Queue[tuple[int, str]] = queue.Queue()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._thread.start()

    def _run(self) -> None:
        import tkinter as tk

        root = tk.Tk()
        root.title("Tiến độ đăng bài")
        root.geometry("360x190+30+80")
        root.resizable(False, False)
        root.attributes("-topmost", True)

        title = tk.Label(root, text=f"Chuẩn bị chạy 0/{self.total}", font=("Segoe UI", 13, "bold"))
        title.pack(pady=(18, 6))
        detail = tk.Label(root, text="Đang chuẩn bị Edge...", font=("Segoe UI", 10))
        detail.pack(pady=(0, 8))

        show_browser_var = tk.BooleanVar(value=self.show_browser.is_set())

        def change_browser_visibility() -> None:
            visible = bool(show_browser_var.get())
            if visible:
                self.show_browser.set()
            else:
                self.show_browser.clear()
            set_edge_window_visible(visible)

        tk.Checkbutton(
            root,
            text="Hiện trình duyệt khi chạy",
            variable=show_browser_var,
            command=change_browser_visibility,
            font=("Segoe UI", 10),
        ).pack(pady=(0, 10))

        buttons = tk.Frame(root)
        buttons.pack()

        def request_stop() -> None:
            self.stop_after_current.set()
            stop_button.config(state="disabled", text="Sẽ dừng sau bài này")
            detail.config(text="Đã nhận lệnh dừng an toàn")

        tk.Button(buttons, text="Ẩn bảng", width=12, command=root.withdraw).pack(side="left", padx=5)
        stop_button = tk.Button(buttons, text="Dừng sau bài này", width=18, command=request_stop)
        stop_button.pack(side="left", padx=5)

        # Nút X chỉ ẩn bảng; tiến trình đăng bài vẫn tiếp tục.
        root.protocol("WM_DELETE_WINDOW", root.withdraw)

        def poll_messages() -> None:
            try:
                while True:
                    completed, message = self._messages.get_nowait()
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

    def should_show_browser(self) -> bool:
        return self.show_browser.is_set()

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


# V1.2 keeps the current domain lock only from Save until its ID has been
# identified. Earlier form/CKEditor/upload work remains concurrent.
_save_article_unlocked = save_article
_find_saved_article_info_unlocked = find_saved_article_info
_SAVE_LOCK_HELD = False


def save_article(driver) -> str:
    global _SAVE_LOCK_HELD
    if SAVE_DOMAIN_LOCK is None:
        return _save_article_unlocked(driver)
    print("    Đang chờ khóa Save/Xác nhận ID của đúng tên miền...")
    SAVE_DOMAIN_LOCK.acquire()
    _SAVE_LOCK_HELD = True
    print("    Đã đến lượt Save/Xác nhận ID của tên miền này.")
    try:
        return _save_article_unlocked(driver)
    except Exception:
        SAVE_DOMAIN_LOCK.release()
        _SAVE_LOCK_HELD = False
        raise


def find_saved_article_info(driver, titles: list[str]) -> tuple[str, str, str]:
    global _SAVE_LOCK_HELD
    try:
        return _find_saved_article_info_unlocked(driver, titles)
    finally:
        if _SAVE_LOCK_HELD and SAVE_DOMAIN_LOCK is not None:
            SAVE_DOMAIN_LOCK.release()
            _SAVE_LOCK_HELD = False


def _normalize_content_marker(value: str) -> str:
    value = unicodedata.normalize("NFD", html_std.unescape(value or "").lower())
    value = value.replace(chr(0x0111), "d")
    return "".join(char for char in value if char.isascii() and char.isalnum())


def build_word_content_markers(content_html: str) -> list[str]:
    """Create distinct Word markers spread across the main article body."""
    blocks = re.split(r"</(?:p|h[1-6]|li|div|tr)>|<br\s*/?>", content_html, flags=re.I)
    normalized = [_normalize_content_marker(re.sub(r"<[^>]+>", " ", block)) for block in blocks]
    normalized = [item for item in normalized if len(item) >= 45]
    if not normalized:
        raise PreSaveValidationError("Khong tao duoc moc noi dung tu Word.")
    positions = {0, len(normalized) - 1}
    for index in range(12):
        positions.add(round(index * (len(normalized) - 1) / 11))
    markers: list[str] = []
    for index in sorted(positions):
        marker = normalized[index][:140]
        if marker and marker not in markers:
            markers.append(marker)
    return markers


def validate_ckeditor_content_markers(driver, markers: list[str]) -> None:
    """Last safety gate: Word body markers must still exist in CKEditor."""
    raw_html = driver.execute_script(
        "return (window.CKEDITOR && CKEDITOR.instances && CKEDITOR.instances.detail) ? CKEDITOR.instances.detail.getData() : '';"
    ) or ""
    editor_text = _normalize_content_marker(re.sub(r"<[^>]+>", " ", raw_html))
    missing = [marker for marker in markers if marker not in editor_text]
    if missing:
        raise PreSaveValidationError(
            "NOI DUNG CKEditor thieu moc Word: "
            + " | ".join(marker[:70] for marker in missing[:4])
        )
    print(f"[CONTENT CHECK] OK: {len(markers)} Word markers found in CKEditor.")


def prepare_article_from_word() -> dict[str, Any]:
    """Prepare exactly one next article in RAM; never opens or saves CMS."""
    current_row = None
    # Ham nay chay trong ThreadPoolExecutor cua worker. COM (Word/Excel) bat
    # buoc phai duoc khoi tao rieng tren moi thread, neu khong DispatchEx
    # se bao: "CoInitialize has not been called".
    import pythoncom
    pythoncom.CoInitialize()
    try:
        current_row, title, word_path = open_hidden_word_document()
        row_from_url, post_url = get_post_url(current_row)
        if row_from_url != current_row:
            raise RuntimeError(
                f"Word row ({current_row}) does not match CMS URL row ({row_from_url})."
            )
        word_info = get_open_word_file_info()
        image_paths = get_image_paths_from_open_word(current_row)
        data = read_first_five_word_lines()
        content_html = read_word_content_as_html()
        content_markers = build_word_content_markers(content_html)
        faq_pairs = read_faq_pairs_from_word()
        faq_note = get_faq_run_note()
        word_structure = inspect_word_structure_before_close()
        print(
            f"[PIPELINE] Word ready in RAM: row {current_row} | {word_info['path']}"
        )
        return {
            "current_row": current_row,
            "title": title,
            "word_path": word_path,
            "post_url": post_url,
            "word_info": word_info,
            "image_paths": image_paths,
            "data": data,
            "content_html": content_html,
            "content_markers": content_markers,
            "faq_pairs": faq_pairs,
            "faq_note": faq_note,
            "word_structure": word_structure,
        }
    finally:
        # The prefetch never leaves Word open and never writes a publish result.
        try:
            close_hidden_word_document()
        finally:
            pythoncom.CoUninitialize()


def process_one_article(
    driver,
    article_index: int,
    total_articles: int,
    prepared: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Chạy một bài bằng driver đang mở và tái sử dụng cho toàn bộ vòng lặp."""
    current_row = None
    cms_id = ""
    uses_prepared_word = prepared is not None
    try:
        print("\n" + "=" * 72)
        print(f"BẮT ĐẦU BÀI {article_index}/{total_articles}")
        print("=" * 72)

        if prepared is None:
            prepared = prepare_article_from_word()

        current_row = int(prepared["current_row"])
        _title = str(prepared["title"])
        post_url = str(prepared["post_url"])
        word_info = prepared["word_info"]
        image_paths = prepared["image_paths"]
        data = prepared["data"]
        content_html = str(prepared["content_html"])
        content_markers = list(prepared["content_markers"])
        faq_pairs = prepared["faq_pairs"]
        faq_note = str(prepared["faq_note"])
        word_structure = prepared["word_structure"]

        print("[PIPELINE] Using Word data already prepared in RAM.")
        print(f"    Excel row : {current_row}")
        print(f"    Word file : {word_info['path']}")
        print(f"    CMS URL   : {post_url}")

        if CMS_ENTRY_LOCK is None:
            if driver.current_url != post_url:
                driver.get(post_url)
                wait_document_ready(driver)
            ensure_post_page_ready(driver, post_url)
        else:
            print(
                "    Đang chờ lượt mở CMS/đăng nhập "
                "(không cho nhiều worker vào website cùng lúc)..."
            )
            with CMS_ENTRY_LOCK:
                print("    Đã đến lượt worker mở CMS/đăng nhập.")
                if driver.current_url != post_url:
                    driver.get(post_url)
                    wait_document_ready(driver)
                ensure_post_page_ready(driver, post_url)
                # Tạo khoảng nghỉ nhỏ trước khi worker tiếp theo chạm CMS.
                time.sleep(random.uniform(2.0, 3.0))

        fill_five_fields(driver, data)

        category_author_result = select_category_author(driver, current_row)
        print("\nKết quả Danh mục/Tác giả:")
        print("- Trạng thái:", category_author_result["status"])
        print("- Danh mục  :", category_author_result["category"])
        print("- Tác giả   :", category_author_result["author"])

        focus_ckeditor(driver)
        set_ckeditor_html(driver, content_html)

        faq_result = fill_faqs(driver, faq_pairs)
        print("\nKết quả FAQ:")
        print("- Trạng thái:", faq_result["status"])
        print("- Số cặp    :", faq_result["count"])

        upload_featured_image(driver, image_paths["featured"])
        upload_content_image(driver, image_paths["content"])

        html_result = clean_and_reorganize_ckeditor_html(driver)
        validate_ckeditor_content_markers(driver, content_markers)
        validate_before_save(
            driver,
            current_row,
            word_info,
            image_paths,
            word_structure,
            html_result,
        )
        saved_url = save_article(driver)

        # Bước phụ tự học ID danh mục: mọi lỗi chỉ cảnh báo, không được làm
        # dừng flow hoặc biến bài đã đăng thành lỗi.
        try:
            upsert_category_id_mapping(
                current_row,
                category_author_result.get("category", ""),
                category_author_result.get("category_id", ""),
            )
        except Exception as mapping_error:
            print(
                "    CẢNH BÁO: Không cập nhật được DANH_MUC_ID, "
                f"vẫn tiếp tục hoàn tất bài đăng: {mapping_error!r}"
            )

        try:
            cms_id, edit_url, public_url = find_saved_article_info(
                driver,
                # Chi doi chieu theo tieu de hien thi (H1) cua bai. Khong
                # dung ten file Word hay Title SEO vi chung de gay trung voi
                # cac hang bai khac tren trang danh sach CMS.
                [data["name"]],
            )
        except RuntimeError as id_error:
            if "không xác định chắc chắn được ID CMS" not in str(id_error):
                raise

            write_publish_result(
                current_row,
                "LỖI ĐĂNG",
                "",
                "Sau khi bấm Lưu, CMS không xác nhận đã tạo bài: "
                + str(id_error),
                "",
            )
            write_url_faq_note(current_row, faq_note)
            print(
                "    LỖI: Sau khi bấm Lưu không tìm thấy ID hoặc đúng bài "
                "trong CMS → ghi LỖI ĐĂNG và giữ trình duyệt để kiểm tra."
            )
            raise RuntimeError(
                "CMS không xác nhận đã tạo bài sau khi bấm Lưu. "
                "Giữ nguyên Edge để kiểm tra thông báo trên website."
            ) from id_error

        print(f"    ID CMS vừa đăng: {cms_id}")

        # Ghi ID ngay khi CMS đã tạo bài. Lỗi lấy URL sau đó không được làm bài
        # đã đăng bị hiểu nhầm thành chưa đăng và bị chạy lại.
        write_publish_result(
            current_row,
            "ĐÃ ĐĂNG",
            public_url,
            "",
            cms_id,
        )
        write_url_faq_note(current_row, faq_note)

        # Nếu trang danh sách không có link công khai chắc chắn, public_url
        # vẫn trống và có thể xử lý riêng bằng phu_tro/xuatID_URL.py.

        result = {
            "current_row": current_row,
            "word_path": str(word_info["path"]),
            "saved_url": saved_url,
            "cms_id": cms_id,
            "edit_url": edit_url,
            "public_url": public_url,
            "html_before": html_result["before"],
            "html_after": html_result["after"],
        }

        print("\n" + "=" * 72)
        print(f"ĐÃ LƯU XONG BÀI {article_index}/{total_articles}")
        print(f"- Dòng Excel: {current_row}")
        print(f"- HTML trước/sau: {html_result['before']}/{html_result['after']} ký tự")
        print(f"- URL CMS sau khi lưu: {saved_url}")
        print("=" * 72)

        return result

    except Exception as exc:
        # Chỉ ném lại lỗi thật của quá trình đăng bài.
        # Lỗi phát sinh khi đóng Word nền sẽ bị bỏ qua.
        # In pipeline mode the Word prefetcher may be preparing the following
        # article in another thread.  This Edge-side failure must not close it.
        if not uses_prepared_word:
            close_word_safely()
        if current_row is not None:
            try:
                if cms_id:
                    write_publish_result(
                        current_row,
                        "ĐÃ ĐĂNG - LỖI URL",
                        "",
                        str(exc),
                        cms_id,
                    )
                else:
                    status = (
                        "LỖI KIỂM TRA"
                        if isinstance(exc, PreSaveValidationError)
                        else "LỖI ĐĂNG"
                    )
                    write_publish_result(current_row, status, "", str(exc))
            except Exception:
                pass
        raise


def main() -> int:
    ensure_runtime_directories()
    current_driver = None
    had_error = False
    progress_window: PublishProgressWindow | None = None

    try:
        print("=" * 72)
        print("ĐĂNG BÀI TỰ ĐỘNG - DÙNG MỘT EDGE CHO TOÀN BỘ CÁC BÀI")
        print("TỰ ĐĂNG NHẬP TỪ FILE INI, KHÔNG TẮT EDGE CÁ NHÂN")
        print("=" * 72)

        total_articles = ask_article_count()
        print(f"Số bài cần chạy: {total_articles}")
        progress_window = PublishProgressWindow(total_articles)

        # Mở Edge ở trang trắng trước.
        # Không mở Word tại đây vì open_hidden_word_document() sẽ:
        # - chọn bài đầu tiên;
        # - ghi trạng thái WORD;
        # - làm vòng lặp sau nhảy sang bài thứ hai.
        #
        # Mỗi bài chỉ được mở Word đúng một lần bên trong process_one_article().
        current_driver = open_edge("about:blank")
        hide_edge_window(current_driver)
        if progress_window.should_show_browser():
            set_edge_window_visible(True)
            print("    Edge đang hiện để theo dõi tiến trình.")
        else:
            print("    Edge sẽ chạy ẩn; tích 'Hiện trình duyệt khi chạy' để kiểm tra.")

        completed_results: list[dict[str, Any]] = []
        validation_skips: list[dict[str, Any]] = []

        # One Word prefetch thread prepares only the next reserved row.  Edge
        # remains controlled exclusively by this main thread.
        with ThreadPoolExecutor(max_workers=1, thread_name_prefix="WordPrefetch") as prefetch:
            prepared_future = prefetch.submit(prepare_article_from_word)

            for article_index in range(1, total_articles + 1):
                if progress_window.should_stop():
                    print("Stopped before starting the next article.")
                    break

                progress_window.update(
                    len(completed_results),
                    f"Preparing article {article_index}/{total_articles}",
                )
                prepared = prepared_future.result()

                # Reserve and prepare one following article while Edge publishes
                # the current one.  No CMS form for that next article is opened.
                if article_index < total_articles:
                    prepared_future = prefetch.submit(prepare_article_from_word)

                progress_window.update(
                    len(completed_results),
                    f"Publishing article {article_index}/{total_articles}",
                )
                try:
                    result = process_one_article(
                        current_driver,
                        article_index,
                        total_articles,
                        prepared=prepared,
                    )
                    completed_results.append(result)
                except PreSaveValidationError as exc:
                    validation_skips.append(
                        {
                            "article_index": article_index,
                            "error": str(exc),
                        }
                    )
                    progress_window.update(
                        article_index,
                        f"Skipped article {article_index}: pre-save validation failed",
                    )
                    print(
                        f"[SKIP {article_index}] Not saved; continuing with the next article."
                    )
                    continue

                progress_window.update(
                    article_index,
                    f"Saved article {article_index}/{total_articles}",
                )

                if progress_window.should_stop():
                    print(f"Stopped safely after article {article_index}/{total_articles}.")
                    break

        processed_count = len(completed_results) + len(validation_skips)
        stopped_early = (
            processed_count < total_articles
            and progress_window.should_stop()
        )
        progress_window.notify_finished(processed_count, stopped=stopped_early)

        print("\n" + "=" * 72)
        if stopped_early:
            print(f"ĐÃ DỪNG AN TOÀN SAU {processed_count}/{total_articles} BÀI")
        else:
            print(f"ĐÃ XỬ LÝ XONG {processed_count}/{total_articles} BÀI")
        for index, result in enumerate(completed_results, start=1):
            print(
                f"- Bài {index}: dòng {result['current_row']} | "
                f"{result['saved_url']}"
            )
        if validation_skips:
            print(
                f"- Bỏ qua {len(validation_skips)} bài do kiểm tra "
                "trước Save không đạt; các bài còn lại vẫn tiếp tục."
            )
        else:
            print("- Toàn bộ quá trình đã hoàn tất, không phát hiện lỗi.")
        print("=" * 72)
        return 0

    except Exception as exc:
        had_error = True
        print("\n" + "=" * 72)
        print("LỖI:")
        print(repr(exc))
        print("\nTRACEBACK CHI TIẾT:")
        traceback.print_exc()
        print("=" * 72)

        if current_driver is not None:
            try:
                current_driver.switch_to.default_content()
                show_edge_window()
                current_driver.maximize_window()
            except Exception:
                pass

        # Không dùng input() vì khi chạy bằng pythonw hoặc nhấp đúp file,
        # console có thể không nhận được bàn phím và tự thoát ngay.
        try:
            from tkinter import Tk, messagebox

            error_root = Tk()
            error_root.withdraw()
            error_root.attributes("-topmost", True)

            messagebox.showerror(
                "Đăng bài gặp lỗi",
                "Chương trình đã dừng tại bài đang lỗi.\n\n"
                f"Lỗi: {exc!r}\n\n"
                "Edge đã được hiện và sẽ được giữ nguyên để kiểm tra.",
                parent=error_root,
            )

            error_root.destroy()
        except Exception:
            # Nếu hộp thoại Windows không mở được, vẫn giữ Edge nhờ detach=True.
            pass

        return 1

    finally:
        # Cleanup không được làm một phiên chạy thành công biến thành lỗi.
        close_word_safely()

        # Chạy thành công: đóng Edge nền.
        # Có lỗi thật khi đăng bài: giữ Edge mở tại đúng trạng thái lỗi.
        if current_driver is not None and not had_error:
            close_driver_safely(current_driver)

# ============================================================
# V2.9 MULTI-WORKER ORCHESTRATOR (embedded publish core above)
# ============================================================

# -*- coding: utf-8 -*-
"""
Bản thử nghiệm đăng bài 3 hoặc 5 worker.

- Mỗi worker là một process, một Word COM và một Edge/profile độc lập.
- Không dùng current_row.ini.
- Tiến trình chính chọn bài, giao đúng dòng và là nơi duy nhất ghi Excel.
- Worker chỉ nhận dữ liệu đã chốt và thao tác CMS.
- Có lỗi: ngừng giao bài mới, giữ Edge lỗi và ghi log riêng.
"""

from datetime import datetime
import importlib.util
import io
import multiprocessing as mp
import os
from pathlib import Path
import queue
import random
import sys
import threading
import time
import traceback
from typing import Any, Callable

from selenium import webdriver
from selenium.webdriver.edge.options import Options


# The publishing core is embedded above: V2.9 is a single-file release.
publish = sys.modules[__name__]


PROJECT_ROOT = Path(
    os.environ.get("HOTKEYVIP_RUNTIME_ROOT", r"D:\CodexProjects\Hotkeyvip")
).resolve()
PROFILE_ROOT = (
    PROJECT_ROOT
    / "06_du_lieu_chay"
    / "dang_bai_workers"
    / "profiles"
)
WORKER_LOG_ROOT = (
    PROJECT_ROOT
    / "06_du_lieu_chay"
    / "log_dang_bai"
)

VERSION = "05_dang_bai_cms (engine V2.11 fast-load)"
EXCEL_BUSY_HRESULT = -2146777998  # 0x800AC472: Excel temporarily rejects COM calls.


class ExcelWriterQueue:
    """The only COM writer; retries Excel's temporary busy response."""

    def __init__(self) -> None:
        self.commands: queue.Queue[tuple[str, Callable[[], None] | None]] = queue.Queue()
        self.ready = threading.Event()
        self.failed = threading.Event()
        self.error: BaseException | None = None
        self.thread = threading.Thread(target=self._run, name="ExcelWriter", daemon=True)

    @staticmethod
    def _is_excel_busy(exc: BaseException) -> bool:
        text = str(exc).casefold()
        return (
            str(EXCEL_BUSY_HRESULT) in text
            or "800ac472" in text
            or "call was rejected by callee" in text
            or "excel is busy" in text
        )

    @staticmethod
    def _retry_delay(attempt: int) -> float:
        return min(5.0, 0.5 * (2 ** min(attempt - 1, 4)))

    def start(self) -> None:
        self.thread.start()
        if not self.ready.wait(timeout=15):
            self.raise_if_failed()
            raise RuntimeError("Excel Writer did not start within 15 seconds.")
        self.raise_if_failed()

    def submit(self, label: str, operation: Callable[[], None]) -> None:
        self.raise_if_failed()
        self.commands.put((label, operation))

    def raise_if_failed(self) -> None:
        if self.failed.is_set():
            raise RuntimeError("Excel Writer stopped unexpectedly.") from self.error

    def drain_and_stop(self) -> None:
        while self.commands.unfinished_tasks:
            self.raise_if_failed()
            time.sleep(0.1)
        self.commands.put(("STOP", None))
        self.thread.join(timeout=20)
        if self.thread.is_alive():
            raise RuntimeError("Excel Writer did not stop within 20 seconds.")
        self.raise_if_failed()

    def _run_operation(self, label: str, operation: Callable[[], None]) -> None:
        attempt = 0
        while True:
            try:
                operation()
                if attempt:
                    print(f"[EXCEL] {label} recovered after {attempt} retry/retries.")
                return
            except BaseException as exc:
                if not self._is_excel_busy(exc):
                    raise
                attempt += 1
                delay = self._retry_delay(attempt)
                print(f"[EXCEL BUSY] {label}; retry {attempt} in {delay:.1f}s: {exc!r}")
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


_EXCEL_WRITER: ExcelWriterQueue | None = None


def queue_excel(label: str, operation: Callable[[], None]) -> None:
    if _EXCEL_WRITER is None:
        operation()
        return
    _EXCEL_WRITER.submit(label, operation)


def queue_excel_and_wait(label: str, operation: Callable[[], Any]) -> Any:
    """Run a COM operation on the single Excel thread and return its result."""
    if _EXCEL_WRITER is None:
        return operation()

    completed = threading.Event()
    outcome: dict[str, Any] = {}

    def wrapped_operation() -> None:
        try:
            outcome["value"] = operation()
        except BaseException as exc:
            # The writer retries temporary Excel-busy errors. Validation
            # errors return to the coordinator without killing that thread.
            if ExcelWriterQueue._is_excel_busy(exc):
                raise
            outcome["error"] = exc
        completed.set()

    _EXCEL_WRITER.submit(label, wrapped_operation)
    while not completed.wait(timeout=0.1):
        _EXCEL_WRITER.raise_if_failed()
    if "error" in outcome:
        raise outcome["error"]
    return outcome.get("value")


def write_publish_result_v22(
    row: int,
    status: str,
    url: str = "",
    error: str = "",
    cms_id: str = "",
) -> None:
    """Delegate result, status colour and exact error Note to the paired V1 core."""
    publish.write_publish_result(row, status, url, error, cms_id)


# V1.0 owns the shared status-colour and Note rules for both modes.


def ask_worker_count() -> int:
    raw = os.environ.get("HOTKEYVIP_PUBLISH_WORKER_COUNT", "").strip()
    return max(1, int(raw)) if raw else 5


def clean_cell(value: Any) -> str:
    return publish.clean_text(value)


def resolve_existing_path(raw: Any) -> str:
    value = clean_cell(raw).strip('"')
    if not value:
        return ""
    return str(
        Path(
            os.path.expandvars(
                os.path.expanduser(value)
            )
        ).resolve()
    )


def load_app_publish_plan() -> dict[str, Any] | None:
    raw = os.environ.get("HOTKEYVIP_PUBLISH_PLAN", "").strip()
    if not raw:
        return None
    try:
        request = json.loads(raw)
        if request.get("mode") == "explicit_error_rows":
            selected_rows = request.get("selected_rows") or []
            if not selected_rows:
                raise ValueError("Danh sách dòng LỖI KIỂM TRA đang trống")
            return {
                "mode": "explicit_error_rows",
                "selected_rows": selected_rows,
                "selected_total": len(selected_rows),
                "groups": [],
                "per_domain_limit": 0,
            }
        if request.get("mode") not in {"balanced_one_category", "selected_domains"}:
            raise ValueError("Chế độ batch không được hỗ trợ")
        per_domain_limit = max(1, int(request["per_domain_limit"]))
        from excel_audit_app.publish_plan import (
            build_balanced_publish_plan,
            inspect_publish_queue,
        )

        return build_balanced_publish_plan(
            inspect_publish_queue(EXCEL_PATH),
            per_domain_limit,
            request.get("category_overrides") or {},
            request.get("domain_limits") or {},
            request.get("mode") == "selected_domains",
        )
    except Exception as exc:
        raise RuntimeError(f"Không thể tạo batch đăng bài từ app: {exc}") from exc


def load_target_tasks(
    limit: int | None,
    app_plan: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Chốt danh sách: Cần mở/Ok trước, trạng thái trống sau."""
    _excel, workbook = publish.get_target_excel_workbook()
    sheet = workbook.Worksheets(publish.SHEET_POST)
    website_sheet = workbook.Worksheets(publish.SHEET_DOMAIN)

    title_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["keyword"]
    )
    seo_title_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["title"]
    )
    h1_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["h1"]
    )
    status_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["status"]
    )
    domain_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["domain"]
    )
    category_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["category"]
    )
    word_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["word_path"]
    )
    image1_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["image1_path"]
    )
    image2_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["image2_path"]
    )

    website_urls: dict[str, str] = {}
    last_website_row = int(
        website_sheet.Cells(
            website_sheet.Rows.Count, 1
        ).End(-4162).Row
    )
    for row in range(2, last_website_row + 1):
        domain = publish.normalize_domain(
            website_sheet.Cells(row, 1).Value
        )
        url = clean_cell(website_sheet.Cells(row, 2).Value)
        if domain and url:
            website_urls[domain] = url

    expected_by_row: dict[int, dict[str, Any]] = {}
    if app_plan is not None:
        expected_by_row = {
            int(item["row"]): dict(item)
            for item in app_plan.get("selected_rows", [])
        }
        selected_rows = list(expected_by_row)
    else:
        need_open: list[int] = []
        blank_status: list[int] = []
        row = 2
        while True:
            title = clean_cell(sheet.Cells(row, title_col).Value)
            if not title:
                break
            status = clean_cell(sheet.Cells(row, status_col).Value)
            status_key = status.casefold()
            if (
                "cần mở" in status_key
                or status_key == "ok"
                or "lỗi kiểm tra" in status_key
                or "lỗi đăng" in status_key
            ):
                need_open.append(row)
            elif not status:
                blank_status.append(row)
            row += 1
        selected_rows = (need_open + blank_status)[: int(limit or 0)]
    if not selected_rows:
        raise RuntimeError(
            'Không còn bài có trạng thái "Cần mở", "Ok" hoặc trạng thái trống.'
        )

    tasks: list[dict[str, Any]] = []
    for row in selected_rows:
        title = clean_cell(sheet.Cells(row, title_col).Value)
        domain = publish.normalize_domain(
            sheet.Cells(row, domain_col).Value
        )
        post_url = website_urls.get(domain, "")
        if not post_url:
            raise RuntimeError(
                f"Dòng {row}: không tìm thấy URL CMS cho domain {domain!r}."
            )

        word_path = resolve_existing_path(
            sheet.Cells(row, word_col).Value
        )
        if not word_path or not Path(word_path).is_file():
            raise RuntimeError(
                f"Dòng {row}: đường dẫn Word không tồn tại: {word_path!r}."
            )

        task = {
                "row": row,
                "title": title,
                "seo_title": clean_cell(
                    sheet.Cells(row, seo_title_col).Value
                ),
                "h1": clean_cell(
                    sheet.Cells(row, h1_col).Value
                ),
                "domain": domain,
                "post_url": post_url,
                "category": clean_cell(
                    sheet.Cells(row, category_col).Value
                ),
                "word_path": word_path,
                "image1_path": resolve_existing_path(
                    sheet.Cells(row, image1_col).Value
                ),
                "image2_path": resolve_existing_path(
                    sheet.Cells(row, image2_col).Value
                ),
            }
        expected = expected_by_row.get(row)
        if expected is not None:
            actual_identity = {
                "domain": task["domain"],
                "category": task["category"],
                "title": task["title"],
                "seo_title": task["seo_title"],
                "h1": task["h1"],
            }
            mismatches = [
                key for key, actual in actual_identity.items()
                if publish.clean_text(actual).casefold()
                != publish.clean_text(expected.get(key, "")).casefold()
            ]
            if mismatches:
                raise RuntimeError(
                    f"Dòng {row} đã thay đổi sau preview ({', '.join(mismatches)}). "
                    "Đã dừng trước khi đăng để tránh nhầm bài."
                )
            task["_expected_identity"] = actual_identity
        tasks.append(task)

    return tasks


def load_target_tasks_fast(
    limit: int | None,
    app_plan: dict[str, Any] | None = None,
) -> list[dict[str, Any]]:
    """Load the publish snapshot with two bulk Excel reads."""
    _excel, workbook = publish.get_target_excel_workbook()
    sheet = workbook.Worksheets(publish.SHEET_POST)
    website_sheet = workbook.Worksheets(publish.SHEET_DOMAIN)

    title_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["keyword"]
    )
    seo_title_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["title"]
    )
    h1_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["h1"]
    )
    status_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["status"]
    )
    domain_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["domain"]
    )
    category_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["category"]
    )
    word_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["word_path"]
    )
    image1_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["image1_path"]
    )
    image2_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["image2_path"]
    )

    def matrix_value(matrix: Any, row: int, col: int) -> Any:
        if isinstance(matrix, tuple):
            row_values = matrix[row - 1]
            if isinstance(row_values, tuple):
                return row_values[col - 1]
            return row_values if col == 1 else None
        return matrix if row == 1 and col == 1 else None

    last_website_row = max(
        1,
        int(website_sheet.Cells(website_sheet.Rows.Count, 1).End(-4162).Row),
    )
    website_matrix = website_sheet.Range(
        website_sheet.Cells(1, 1),
        website_sheet.Cells(last_website_row, 2),
    ).Value2
    website_urls: dict[str, str] = {}
    for row in range(2, last_website_row + 1):
        domain = publish.normalize_domain(matrix_value(website_matrix, row, 1))
        url = clean_cell(matrix_value(website_matrix, row, 2))
        if domain and url:
            website_urls[domain] = url

    expected_by_row: dict[int, dict[str, Any]] = {}
    if app_plan is not None:
        expected_by_row = {
            int(item["row"]): dict(item)
            for item in app_plan.get("selected_rows", [])
        }

    last_post_row = max(
        1,
        int(sheet.Cells(sheet.Rows.Count, title_col).End(-4162).Row),
        max(expected_by_row, default=1),
    )
    max_col = max(
        title_col,
        seo_title_col,
        h1_col,
        status_col,
        domain_col,
        category_col,
        word_col,
        image1_col,
        image2_col,
    )
    post_matrix = sheet.Range(
        sheet.Cells(1, 1),
        sheet.Cells(last_post_row, max_col),
    ).Value2

    if expected_by_row:
        selected_rows = list(expected_by_row)
    else:
        need_open: list[int] = []
        blank_status: list[int] = []
        for row in range(2, last_post_row + 1):
            title = clean_cell(matrix_value(post_matrix, row, title_col))
            if not title:
                break
            status = clean_cell(matrix_value(post_matrix, row, status_col))
            status_key = status.casefold()
            if (
                "cần mở" in status_key
                or status_key == "ok"
                or "lỗi kiểm tra" in status_key
                or "lỗi đăng" in status_key
            ):
                need_open.append(row)
            elif not status:
                blank_status.append(row)
        selected_rows = (need_open + blank_status)[: int(limit or 0)]

    if not selected_rows:
        raise RuntimeError(
            'Không còn bài có trạng thái "Cần mở", "Ok" hoặc trạng thái trống.'
        )

    tasks: list[dict[str, Any]] = []
    for row in selected_rows:
        title = clean_cell(matrix_value(post_matrix, row, title_col))
        domain = publish.normalize_domain(
            matrix_value(post_matrix, row, domain_col)
        )
        post_url = website_urls.get(domain, "")
        if not post_url:
            raise RuntimeError(
                f"Dòng {row}: không tìm thấy URL CMS cho domain {domain!r}."
            )

        word_path = resolve_existing_path(
            matrix_value(post_matrix, row, word_col)
        )
        if not word_path or not Path(word_path).is_file():
            raise RuntimeError(
                f"Dòng {row}: đường dẫn Word không tồn tại: {word_path!r}."
            )

        task = {
            "row": row,
            "title": title,
            "seo_title": clean_cell(
                matrix_value(post_matrix, row, seo_title_col)
            ),
            "h1": clean_cell(matrix_value(post_matrix, row, h1_col)),
            "domain": domain,
            "post_url": post_url,
            "category": clean_cell(
                matrix_value(post_matrix, row, category_col)
            ),
            "word_path": word_path,
            "image1_path": resolve_existing_path(
                matrix_value(post_matrix, row, image1_col)
            ),
            "image2_path": resolve_existing_path(
                matrix_value(post_matrix, row, image2_col)
            ),
        }
        expected = expected_by_row.get(row)
        if expected is not None:
            actual_identity = {
                "domain": task["domain"],
                "category": task["category"],
                "title": task["title"],
                "seo_title": task["seo_title"],
                "h1": task["h1"],
            }
            mismatches = [
                key
                for key, actual in actual_identity.items()
                if publish.clean_text(actual).casefold()
                != publish.clean_text(expected.get(key, "")).casefold()
            ]
            if mismatches:
                raise RuntimeError(
                    f"Dòng {row} đã thay đổi sau preview ({', '.join(mismatches)}). "
                    "Đã dừng trước khi đăng để tránh nhầm bài."
                )
            task["_expected_identity"] = actual_identity
        tasks.append(task)

    print(
        f"[FAST LOAD] Đã nạp {len(tasks)} bài bằng 2 lần đọc vùng Excel "
        f"(DANG_BAI đến dòng {last_post_row})."
    )
    return tasks


def verify_dispatch_task(task: dict[str, Any]) -> None:
    """Kiểm tra lại danh tính dòng ngay trước khi giao Worker đăng thật."""
    expected = task.get("_expected_identity")
    if not expected:
        return

    row = int(task["row"])

    def queued_verification() -> None:
        _excel, workbook = publish.get_target_excel_workbook()
        sheet = workbook.Worksheets(publish.SHEET_POST)
        columns = {
            "domain": publish.find_column_by_header(sheet, publish.PUBLISH_HEADERS["domain"]),
            "category": publish.find_column_by_header(sheet, publish.PUBLISH_HEADERS["category"]),
            "title": publish.find_column_by_header(sheet, publish.PUBLISH_HEADERS["keyword"]),
            "seo_title": publish.find_column_by_header(sheet, publish.PUBLISH_HEADERS["title"]),
            "h1": publish.find_column_by_header(sheet, publish.PUBLISH_HEADERS["h1"]),
        }
        actual = {
            key: publish.clean_text(sheet.Cells(row, column).Value)
            for key, column in columns.items()
        }
        mismatches = [
            key for key, value in actual.items()
            if value.casefold() != publish.clean_text(expected.get(key, "")).casefold()
        ]
        if mismatches:
            raise RuntimeError(
                f"[BẢO HIỂM DÒNG] Dòng {row} đã thay đổi trước khi đăng "
                f"({', '.join(mismatches)}). Không giao bài cho Worker."
            )

    queue_excel_and_wait(f"verify identity row {row}", queued_verification)


def set_dispatch_status(row: int, worker_id: int) -> None:
    _excel, workbook = publish.get_target_excel_workbook()
    sheet = workbook.Worksheets(publish.SHEET_POST)
    status_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["status"]
    )
    sheet.Cells(row, status_col).Value = f"ĐANG ĐĂNG - W{worker_id}"
    workbook.Save()


def set_deferred_406_status(row: int, message: str) -> None:
    _excel, workbook = publish.get_target_excel_workbook()
    sheet = workbook.Worksheets(publish.SHEET_POST)
    sheet.Cells(
        row,
        publish.find_column_by_header(
            sheet, publish.PUBLISH_HEADERS["status"]
        ),
    ).Value = "Cần mở"
    publish.write_compact_error_cell(
        sheet.Cells(
            row,
            publish.find_column_by_header(
                sheet, publish.PUBLISH_HEADERS["publish_error"]
            ),
        ),
        message,
    )
    sheet.Cells(
        row,
        publish.find_column_by_header(
            sheet, publish.PUBLISH_HEADERS["published_at"]
        ),
    ).Value2 = excel_local_now_serial()
    workbook.Save()


def mark_domain_skipped(
    rows: list[int],
    message: str,
) -> None:
    """Giữ các bài domain lỗi ở trạng thái Cần mở để xử lý ở phiên sau."""
    if not rows:
        return
    _excel, workbook = publish.get_target_excel_workbook()
    sheet = workbook.Worksheets(publish.SHEET_POST)
    status_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["status"]
    )
    error_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["publish_error"]
    )
    time_col = publish.find_column_by_header(
        sheet, publish.PUBLISH_HEADERS["published_at"]
    )
    for row in rows:
        sheet.Cells(row, status_col).Value = "Cần mở"
        publish.write_compact_error_cell(
            sheet.Cells(row, error_col),
            message,
        )
        sheet.Cells(row, time_col).Value2 = excel_local_now_serial()
    workbook.Save()


def write_worker_failure(row: int, error: str) -> None:
    publish.write_publish_result(
        row,
        "LỖI ĐĂNG",
        "",
        error,
        "",
    )


def apply_publish_writes(
    writes: list[dict[str, str]],
    fallback_row: int,
) -> None:
    if not writes:
        raise RuntimeError(
            f"Worker không trả thao tác ghi Excel cho dòng {fallback_row}."
        )
    for item in writes:
        publish.write_publish_result(
            int(item["row"]),
            item["status"],
            item["url"],
            item["error"],
            item["cms_id"],
        )


def apply_category_mappings(
    mappings: list[dict[str, str]],
) -> None:
    """Bước phụ: lỗi mapping chỉ cảnh báo, không dừng đăng bài."""
    for item in mappings:
        try:
            publish.upsert_category_id_mapping(
                int(item["row"]),
                item["category"],
                item["category_id"],
            )
        except Exception as exc:
            print(
                "    CẢNH BÁO: Không cập nhật được DANH_MUC_ID, "
                f"bỏ qua và tiếp tục: {exc!r}"
            )


def apply_url_faq_notes(
    notes: list[dict[str, str]],
) -> None:
    for item in notes:
        publish.write_url_faq_note(
            int(item["row"]),
            item["note"],
        )


# V2.0 overrides: all Excel writes from the coordinator are queued.  The
# original functions above are intentionally left intact for comparison with
# the TEST version.
def set_dispatch_status(row: int, worker_id: int) -> None:
    def operation() -> None:
        _excel, workbook = publish.get_target_excel_workbook()
        sheet = workbook.Worksheets(publish.SHEET_POST)
        status_col = publish.find_column_by_header(
            sheet, publish.PUBLISH_HEADERS["status"]
        )
        sheet.Cells(row, status_col).Value = f"ĐANG ĐĂNG - W{worker_id}"
        workbook.Save()

    queue_excel(f"dispatch status row {row}", operation)


def set_deferred_406_status(row: int, message: str) -> None:
    def operation() -> None:
        _excel, workbook = publish.get_target_excel_workbook()
        sheet = workbook.Worksheets(publish.SHEET_POST)
        sheet.Cells(row, publish.find_column_by_header(
            sheet, publish.PUBLISH_HEADERS["status"]
        )).Value = "Cần mở"
        publish.write_compact_error_cell(sheet.Cells(row, publish.find_column_by_header(
            sheet, publish.PUBLISH_HEADERS["publish_error"]
        )), message)
        sheet.Cells(row, publish.find_column_by_header(
            sheet, publish.PUBLISH_HEADERS["published_at"]
        )).Value2 = excel_local_now_serial()
        workbook.Save()

    queue_excel(f"defer 406 row {row}", operation)


def mark_domain_skipped(rows: list[int], message: str) -> None:
    if not rows:
        return

    def operation() -> None:
        _excel, workbook = publish.get_target_excel_workbook()
        sheet = workbook.Worksheets(publish.SHEET_POST)
        status_col = publish.find_column_by_header(sheet, publish.PUBLISH_HEADERS["status"])
        error_col = publish.find_column_by_header(sheet, publish.PUBLISH_HEADERS["publish_error"])
        time_col = publish.find_column_by_header(sheet, publish.PUBLISH_HEADERS["published_at"])
        for row in rows:
            sheet.Cells(row, status_col).Value = "Cần mở"
            publish.write_compact_error_cell(sheet.Cells(row, error_col), message)
            sheet.Cells(row, time_col).Value2 = excel_local_now_serial()
        workbook.Save()

    queue_excel(f"skip domain rows {','.join(map(str, rows))}", operation)


def write_worker_failure(row: int, error: str) -> None:
    queue_excel(
        f"write failure row {row}",
        lambda: write_publish_result_v22(row, "LỖI ĐĂNG", "", error, ""),
    )


def write_validation_failure(row: int, error: str) -> None:
    queue_excel(
        f"write validation failure row {row}",
        lambda: write_publish_result_v22(row, "LỖI KIỂM TRA", "", error, ""),
    )


def apply_publish_writes(writes: list[dict[str, str]], fallback_row: int) -> None:
    if not writes:
        raise RuntimeError(f"Worker did not return Excel writes for row {fallback_row}.")
    for item in writes:
        row = int(item["row"])
        status, url = item["status"], item["url"]
        error, cms_id = item["error"], item["cms_id"]
        queue_excel_and_wait(
            f"publish result row {row}",
            lambda row=row, status=status, url=url, error=error, cms_id=cms_id:
                write_publish_result_v22(row, status, url, error, cms_id),
        )
        print(
            f"[EXCEL SAVED] Dòng {row} | trạng thái: {status} "
            f"| CMS ID: {cms_id or '-'}"
        )


def apply_category_mappings(mappings: list[dict[str, str]]) -> None:
    for item in mappings:
        row = int(item["row"])
        category, category_id = item["category"], item["category_id"]

        def operation(row: int = row, category: str = category, category_id: str = category_id) -> None:
            try:
                publish.upsert_category_id_mapping(row, category, category_id)
            except Exception as exc:
                print(f"[EXCEL WARNING] category mapping row {row}: {exc!r}")

        queue_excel(f"category mapping row {row}", operation)


def apply_url_faq_notes(notes: list[dict[str, str]]) -> None:
    for item in notes:
        row, note = int(item["row"]), item["note"]
        queue_excel(
            f"URL/FAQ note row {row}",
            lambda row=row, note=note: publish.write_url_faq_note(row, note),
        )


def create_worker_driver(worker_id: int):
    profile = PROFILE_ROOT / f"worker_{worker_id}"
    profile.mkdir(parents=True, exist_ok=True)

    options = Options()
    options.add_argument(f"--user-data-dir={profile}")
    options.add_argument("--disable-notifications")
    options.add_argument("--start-maximized")
    options.add_experimental_option("detach", True)

    driver = webdriver.Edge(options=options)
    driver.maximize_window()
    driver.get("about:blank")
    publish.wait_document_ready(driver)
    publish.hide_edge_window(driver)
    return driver


def install_task_adapters(
    task: dict[str, Any],
    captured_writes: list[dict[str, str]],
    captured_mappings: list[dict[str, str]],
    captured_notes: list[dict[str, str]],
) -> None:
    """Thay các điểm chạm Excel bằng dữ liệu task đã được tiến trình chính chốt."""
    row = int(task["row"])
    title = str(task["title"])
    word_path = Path(str(task["word_path"])).resolve()

    publish.find_next_word_file_from_excel = (
        lambda: (row, title, word_path)
    )

    def get_asset(_row: int, header_key: str) -> Path | None:
        key = (
            "image1_path"
            if header_key == "image1_path"
            else "image2_path"
        )
        raw = str(task.get(key, "") or "")
        if not raw:
            return None
        path = Path(raw).resolve()
        if not path.is_file():
            raise RuntimeError(
                f'Dòng {row}, ảnh "{header_key}" không tồn tại:\n{path}'
            )
        return path

    publish.get_publish_asset_path = get_asset
    publish.get_post_url = (
        lambda current_row: (row, str(task["post_url"]))
    )
    publish.get_pre_save_expected_values = lambda _row: {
        "domain": publish.clean_text(task.get("domain", "")),
        "category": publish.clean_text(task.get("category", "")),
        "title": publish.clean_text(task.get("seo_title", "")),
        "h1": publish.clean_text(task.get("h1", "")),
        # Cột Excel "Tiêu đề" hiện đang là Keyword chính.
        "keyword": publish.clean_text(task.get("title", "")),
        "image1_path": publish.clean_text(task.get("image1_path", "")),
        "image2_path": publish.clean_text(task.get("image2_path", "")),
    }

    def select_category_author(driver, current_row: int):
        category = publish.clean_text(task.get("category", ""))
        print("[6] Chọn Danh mục / Tác giả từ dữ liệu đã chốt...")
        if not category:
            return {
                "status": "SKIP",
                "category": "",
                "author": "",
            }

        result = publish.choose_select2_by_text(
            driver,
            'select[name="cat_id"]',
            category,
            "Danh mục",
        )
        category_id = publish.clean_text(result.get("value"))
        if not category_id:
            raise RuntimeError(
                f'Đã chọn Danh mục "{category}" nhưng CMS không trả ID.'
            )
        for author_attempt in range(2):
            try:
                author_select = publish.wait_author_select_ready(driver)
                author = publish.choose_random_author(
                    driver,
                    author_select,
                )
                break
            except publish.StaleElementReferenceException as exc:
                if author_attempt == 0:
                    print(
                        "    Ô Tác giả vừa bị Ajax thay mới "
                        "→ tìm lại và thử thêm một lần."
                    )
                    time.sleep(0.5)
                    continue
                raise RuntimeError(
                    "Ô Tác giả tiếp tục bị Ajax thay mới "
                    "sau một lần thử lại."
                ) from exc
            except publish.TimeoutException as exc:
                raise RuntimeError(
                    "Không tìm thấy Tác giả hợp lệ sau khi "
                    "chờ website tải danh sách."
                ) from exc
        return {
            "status": "OK",
            "category": category,
            "category_id": category_id,
            "author": author,
        }

    publish.select_category_author = select_category_author

    def capture_write(
        write_row: int,
        status: str,
        url: str = "",
        error: str = "",
        cms_id: str = "",
    ) -> None:
        captured_writes.append(
            {
                "row": str(write_row),
                "status": str(status),
                "url": str(url),
                "error": str(error),
                "cms_id": str(cms_id),
            }
        )

    publish.write_publish_result = capture_write

    def capture_mapping(
        mapping_row: int,
        category: str,
        category_id: str,
    ) -> dict[str, str]:
        if publish.clean_text(category) and publish.clean_text(category_id):
            captured_mappings.append(
                {
                    "row": str(mapping_row),
                    "category": str(category),
                    "category_id": str(category_id),
                }
            )
        return {"status": "QUEUED", "message": ""}

    publish.upsert_category_id_mapping = capture_mapping

    def capture_faq_note(note_row: int, note: str) -> None:
        if publish.clean_text(note):
            captured_notes.append(
                {"row": str(note_row), "note": str(note)}
            )

    publish.write_url_faq_note = capture_faq_note


def save_worker_error_log(
    worker_id: int,
    row: int,
    content: str,
) -> str:
    WORKER_LOG_ROOT.mkdir(parents=True, exist_ok=True)
    path = WORKER_LOG_ROOT / (
        f"dang_bai_worker{worker_id}_row{row}_"
        f"{datetime.now():%Y%m%d_%H%M%S}.log"
    )
    path.write_text(content, encoding="utf-8-sig")
    return str(path)


def worker_main(
    worker_id: int,
    command_queue,
    result_queue,
    login_lock,
    domain_save_locks,
    word_clipboard_lock,
    browser_visible_event,
) -> None:
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    memory_log = io.StringIO()
    sys.stdout = publish.TeeOutput(original_stdout, memory_log)
    sys.stderr = publish.TeeOutput(original_stderr, memory_log)
    driver = None
    had_error = False
    current_row = 0

    try:
        # V2.4: opening a CMS page is concurrent.  The core takes LOGIN_LOCK
        # only when it sees an actual login form, and receives the current
        # domain's lock only for Save + ID confirmation.
        publish.CMS_ENTRY_LOCK = None
        publish.LOGIN_LOCK = login_lock
        publish.WORD_CLIPBOARD_LOCK = word_clipboard_lock
        driver = create_worker_driver(worker_id)

        def watch_browser_visibility() -> None:
            last_state = None
            last_hwnd = None
            while True:
                try:
                    desired = browser_visible_event.is_set()
                    current_hwnd = publish._EDGE_HWND
                    if desired != last_state or current_hwnd != last_hwnd:
                        publish.set_edge_window_visible(desired)
                        last_state = desired
                        last_hwnd = current_hwnd
                    time.sleep(0.2)
                except Exception:
                    time.sleep(0.5)

        threading.Thread(
            target=watch_browser_visibility,
            daemon=True,
        ).start()
        result_queue.put(
            {"type": "ready", "worker_id": worker_id}
        )

        while True:
            task = command_queue.get()
            if task is None:
                break

            current_row = int(task["row"])
            publish.SAVE_DOMAIN_LOCK = domain_save_locks.get(str(task["domain"]))
            captured_writes: list[dict[str, str]] = []
            captured_mappings: list[dict[str, str]] = []
            captured_notes: list[dict[str, str]] = []
            install_task_adapters(
                task,
                captured_writes,
                captured_mappings,
                captured_notes,
            )
            print(
                f"\n[WORKER {worker_id}] BẮT ĐẦU DÒNG {current_row}"
            )

            try:
                # V2.9: this worker already owns the task.  Prepare its Word
                # data in one thread while the main worker thread opens the
                # same task's CMS page and logs in only if necessary.
                with ThreadPoolExecutor(
                    max_workers=1,
                    thread_name_prefix=f"WordTask{worker_id}",
                ) as word_preparer:
                    prepared_future = word_preparer.submit(
                        publish.prepare_article_from_word
                    )
                    post_url = str(task["post_url"])
                    if driver.current_url != post_url:
                        driver.get(post_url)
                        publish.wait_document_ready(driver)
                    publish.ensure_post_page_ready(driver, post_url)
                    prepared = prepared_future.result()

                result = publish.process_one_article(
                    driver,
                    int(task["article_index"]),
                    int(task["total_articles"]),
                    prepared=prepared,
                )
                result_queue.put(
                    {
                        "type": "done",
                        "worker_id": worker_id,
                        "row": current_row,
                        "result": result,
                        "writes": captured_writes,
                        "mappings": captured_mappings,
                        "notes": captured_notes,
                    }
                )
            except publish.PreSaveValidationError as exc:
                log_path = save_worker_error_log(
                    worker_id,
                    current_row,
                    memory_log.getvalue(),
                )
                result_queue.put(
                    {
                        "type": "validation_error",
                        "worker_id": worker_id,
                        "row": current_row,
                        "domain": str(task["domain"]),
                        "error": str(exc),
                        "writes": captured_writes,
                        "mappings": captured_mappings,
                        "notes": captured_notes,
                        "log_path": log_path,
                    }
                )
                print(
                    f"[WORKER {worker_id}] Bỏ qua dòng {current_row}; "
                    "không Save và tiếp tục nhận bài khác."
                )
                continue

            except Exception as exc:
                if isinstance(exc, publish.Http406Error):
                    traceback.print_exc()
                    log_path = save_worker_error_log(
                        worker_id,
                        current_row,
                        memory_log.getvalue(),
                    )
                    result_queue.put(
                        {
                            "type": "blocked_406",
                            "worker_id": worker_id,
                            "row": current_row,
                            "domain": str(task["domain"]),
                            "task": task,
                            "error": str(exc),
                            "log_path": log_path,
                        }
                    )
                    publish.close_driver_safely(driver)
                    driver = create_worker_driver(worker_id)
                    result_queue.put(
                        {
                            "type": "ready",
                            "worker_id": worker_id,
                        }
                    )
                    continue

                trace_text = traceback.format_exc()
                print(trace_text, end="")
                log_path = save_worker_error_log(
                    worker_id,
                    current_row,
                    memory_log.getvalue(),
                )
                result_queue.put(
                    {
                        "type": "task_error",
                        "worker_id": worker_id,
                        "row": current_row,
                        "domain": str(task["domain"]),
                        "task": task,
                        "domain_access_error": (
                            "ensure_post_page_ready" in trace_text
                        ),
                        "error": repr(exc),
                        "writes": captured_writes,
                        "mappings": captured_mappings,
                        "notes": captured_notes,
                        "log_path": log_path,
                    }
                )
                # Lỗi của riêng bài này đã được gửi về tiến trình chính để
                # ghi trạng thái/Note vào Excel. Worker vẫn giữ Edge và nhận
                # bài kế tiếp; chỉ lỗi truy cập trang đăng nhập/form mới làm
                # worker dừng qua nhánh fatal bên ngoài.
                continue
    except Exception as exc:
        had_error = True
        traceback.print_exc()
        log_path = save_worker_error_log(
            worker_id,
            current_row,
            memory_log.getvalue(),
        )
        result_queue.put(
            {
                "type": "fatal",
                "worker_id": worker_id,
                "row": current_row,
                "error": repr(exc),
                "log_path": log_path,
            }
        )
    finally:
        publish.close_word_safely()
        if driver is not None and not had_error:
            publish.close_driver_safely(driver)
        sys.stdout = original_stdout
        sys.stderr = original_stderr


def stop_workers(command_queues, workers) -> None:
    for command_queue in command_queues.values():
        try:
            command_queue.put_nowait(None)
        except Exception:
            pass
    for process in workers.values():
        process.join(timeout=8)


class MultiProgressWindow:
    """Bảng tiến độ; nút dừng chỉ dừng giao bài mới."""

    def __init__(self, worker_count: int, total: int) -> None:
        self.worker_count = worker_count
        self.total = total
        self.stop_requested = threading.Event()
        self.browser_visible_requested = threading.Event()
        self.messages: queue.Queue[dict[str, Any]] = queue.Queue()
        self._thread = threading.Thread(
            target=self._run,
            daemon=True,
        )
        self._thread.start()

    def _run(self) -> None:
        import tkinter as tk

        root = tk.Tk()
        root.title(
            f"Tiến độ đăng bài — {self.worker_count} luồng"
        )
        root.geometry(
            f"{max(620, 210 * self.worker_count)}x320+25+70"
        )
        root.resizable(False, False)
        root.attributes("-topmost", True)

        heading = tk.Label(
            root,
            text=(
                f"Đang chạy {self.worker_count} luồng — "
                f"tổng cộng {self.total} bài"
            ),
            font=("Segoe UI", 13, "bold"),
        )
        heading.pack(pady=(16, 10))

        worker_frame = tk.Frame(root)
        worker_frame.pack(fill="x", padx=12)
        worker_labels: dict[int, tuple[Any, Any]] = {}

        for worker_id in range(1, self.worker_count + 1):
            group = tk.LabelFrame(
                worker_frame,
                text=f"Worker {worker_id}",
                padx=8,
                pady=8,
            )
            group.pack(
                side="left",
                fill="both",
                expand=True,
                padx=4,
            )
            row_label = tk.Label(
                group,
                text="Dòng: —",
                font=("Segoe UI", 10, "bold"),
            )
            row_label.pack(pady=(2, 6))
            step_label = tk.Label(
                group,
                text="Đang mở Edge...",
                wraplength=175,
                justify="center",
            )
            step_label.pack()
            worker_labels[worker_id] = (
                row_label,
                step_label,
            )

        total_label = tk.Label(
            root,
            text=f"Đã hoàn thành 0/{self.total} bài",
            font=("Segoe UI", 11, "bold"),
        )
        total_label.pack(pady=(18, 10))

        button_frame = tk.Frame(root)
        button_frame.pack()

        def request_stop() -> None:
            self.stop_requested.set()
            stop_button.config(
                state="disabled",
                text="Đang chờ các bài hiện tại hoàn tất",
            )
            total_label.config(
                text="Đã yêu cầu dừng an toàn — không giao bài mới"
            )

        def toggle_browsers() -> None:
            if self.browser_visible_requested.is_set():
                self.browser_visible_requested.clear()
                browser_button.config(text="Hiện trình duyệt")
            else:
                self.browser_visible_requested.set()
                browser_button.config(text="Ẩn trình duyệt")

        tk.Button(
            button_frame,
            text="Ẩn bảng",
            width=12,
            command=root.withdraw,
        ).pack(side="left", padx=5)
        browser_button = tk.Button(
            button_frame,
            text="Hiện trình duyệt",
            width=18,
            command=toggle_browsers,
        )
        browser_button.pack(side="left", padx=5)
        stop_button = tk.Button(
            button_frame,
            text="Dừng an toàn",
            width=30,
            command=request_stop,
        )
        stop_button.pack(side="left", padx=5)

        root.protocol("WM_DELETE_WINDOW", root.withdraw)

        def poll_messages() -> None:
            try:
                while True:
                    message = self.messages.get_nowait()
                    message_type = message.get("type")
                    if message_type == "worker":
                        worker_id = int(
                            message.get("worker_id", 0)
                        )
                        if worker_id in worker_labels:
                            row_label, step_label = worker_labels[
                                worker_id
                            ]
                            row = message.get("row")
                            row_label.config(
                                text=(
                                    f"Dòng: {row}"
                                    if row
                                    else "Dòng: —"
                                )
                            )
                            step_label.config(
                                text=str(message.get("step", ""))
                            )
                    elif message_type == "total":
                        completed = int(
                            message.get("completed", 0)
                        )
                        total_label.config(
                            text=(
                                f"Đã hoàn thành {completed}/"
                                f"{self.total} bài"
                            )
                        )
                    elif message_type == "finished":
                        total_label.config(
                            text=str(message.get("text", "Đã kết thúc"))
                        )
                        stop_button.config(state="disabled")
            except queue.Empty:
                pass
            root.after(150, poll_messages)

        poll_messages()
        root.mainloop()

    def worker(
        self,
        worker_id: int,
        row: int | None,
        step: str,
    ) -> None:
        self.messages.put(
            {
                "type": "worker",
                "worker_id": worker_id,
                "row": row,
                "step": step,
            }
        )

    def total_done(self, completed: int) -> None:
        self.messages.put(
            {"type": "total", "completed": completed}
        )

    def finish(self, completed: int, stopped: bool) -> None:
        if stopped:
            text = (
                f"Đã dừng an toàn sau {completed}/{self.total} bài"
            )
        else:
            text = (
                f"Đã kết thúc {completed}/{self.total} bài"
            )
        self.messages.put(
            {"type": "finished", "text": text}
        )


def pop_next_domain_safe_task(
    pending_tasks: list[dict[str, Any]],
    active_domains: set[str],
    deferred_406_domains: set[str],
    deferred_until: dict[str, float],
) -> dict[str, Any] | None:
    """Ưu tiên domain khác; hết domain trống thì cho chạy chung domain."""
    for index, task in enumerate(pending_tasks):
        domain = str(task["domain"])
        if (
            domain not in active_domains
            and domain not in deferred_406_domains
        ):
            selected = pending_tasks.pop(index)
            selected["_retrying_406"] = False
            return selected

    # Không còn domain trống: cho worker rảnh lấy bài thường thuộc domain
    # đang chạy. Khoảng cách giao bài 1-2 giây vẫn được giữ ở run_multi().
    # Domain đã gặp 406 vẫn bị hoãn tới cuối hàng đợi.
    for index, task in enumerate(pending_tasks):
        domain = str(task["domain"])
        if domain not in deferred_406_domains:
            selected = pending_tasks.pop(index)
            selected["_retrying_406"] = False
            return selected

    # Chỉ khi mọi domain bình thường đã làm xong mới xét domain bị 406.
    if active_domains:
        return None

    now = time.time()
    retry_candidates = [
        (index, task)
        for index, task in enumerate(pending_tasks)
        if (
            str(task["domain"]) not in active_domains
            and str(task["domain"]) in deferred_406_domains
            and now >= deferred_until.get(str(task["domain"]), 0)
        )
    ]
    retry_candidates.sort(
        key=lambda item: (
            int(item[1].get("_406_attempt", 0)) == 0,
            int(item[1].get("_sequence", 0)),
        )
    )
    if retry_candidates:
        index, _task = retry_candidates[0]
        selected = pending_tasks.pop(index)
        domain = str(selected["domain"])
        deferred_406_domains.discard(domain)
        selected["_retrying_406"] = True
        return selected

    return None


def run_multi() -> int:
    global _EXCEL_WRITER
    publish.ensure_runtime_directories()
    worker_count = ask_worker_count()
    app_plan = load_app_publish_plan()
    if app_plan is not None:
        tasks = load_target_tasks_fast(None, app_plan)
        if app_plan.get("mode") == "explicit_error_rows":
            print(f"[APP ĐĂNG LẠI] {len(tasks)} bài LỖI KIỂM TRA | 1 worker.")
        else:
            print(
                f"[APP BATCH] {len(app_plan['groups'])} tên miền | "
                f"tối đa {app_plan['per_domain_limit']} bài/tên miền | "
                f"tổng {len(tasks)} bài | mỗi tên miền đúng một danh mục."
            )
    else:
        requested_total = publish.ask_article_count()
        tasks = load_target_tasks_fast(requested_total)
    total_articles = len(tasks)
    progress = MultiProgressWindow(
        worker_count,
        total_articles,
    )
    _EXCEL_WRITER = ExcelWriterQueue()
    _EXCEL_WRITER.start()

    print("=" * 72)
    print(f"PHIEN BAN: {VERSION}")
    print(
        f"ĐĂNG BÀI ĐA LUỒNG | {worker_count} WORKER | "
        f"{total_articles} BÀI"
    )
    print("Không dùng current_row.ini; chỉ tiến trình chính ghi Excel.")
    print("=" * 72)

    context = mp.get_context("spawn")
    result_queue = context.Queue()
    login_lock = context.Lock()
    word_clipboard_lock = context.Lock()
    domain_save_locks = {
        str(task["domain"]): context.Lock()
        for task in tasks
    }
    browser_visible_event = context.Event()
    command_queues: dict[int, Any] = {}
    workers: dict[int, Any] = {}

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
                word_clipboard_lock,
                browser_visible_event,
            ),
            name=f"PublishWorker-{worker_id}",
        )
        process.start()
        command_queues[worker_id] = command_queue
        workers[worker_id] = process

    free_workers: set[int] = set()
    active_rows: dict[int, int] = {}
    active_tasks: dict[int, dict[str, Any]] = {}
    active_domains: set[str] = set()
    pending_tasks: list[dict[str, Any]] = []
    for sequence, source_task in enumerate(tasks, start=1):
        task = dict(source_task)
        task["_sequence"] = sequence
        task["_406_attempt"] = 0
        task["article_index"] = sequence
        task["total_articles"] = total_articles
        pending_tasks.append(task)
    deferred_406_domains: set[str] = set()
    deferred_until: dict[str, float] = {}
    completed = 0
    stopping = False
    stopped_safely = False
    had_error = False
    next_dispatch_at = 0.0

    try:
        while True:
            _EXCEL_WRITER.raise_if_failed()
            if progress.browser_visible_requested.is_set():
                browser_visible_event.set()
            else:
                browser_visible_event.clear()

            if progress.stop_requested.is_set() and not stopping:
                stopping = True
                stopped_safely = True
                print(
                    "[DỪNG AN TOÀN] Không giao bài mới; "
                    "đang chờ các bài hiện tại hoàn tất."
                )

            try:
                message = result_queue.get(timeout=0.25)
            except queue.Empty:
                message = None

            if message:
                message_type = message.get("type")
                worker_id = int(message.get("worker_id", 0))

                if message_type == "ready":
                    free_workers.add(worker_id)
                    print(f"Worker {worker_id}: sẵn sàng.")
                    progress.worker(
                        worker_id,
                        None,
                        "Sẵn sàng nhận bài",
                    )

                elif message_type == "done":
                    row = int(message["row"])
                    active_rows.pop(worker_id, None)
                    finished_task = active_tasks.pop(worker_id, None)
                    if finished_task:
                        active_domains.discard(
                            str(finished_task["domain"])
                        )
                    apply_category_mappings(
                        list(message.get("mappings", []))
                    )
                    apply_publish_writes(message["writes"], row)
                    apply_url_faq_notes(
                        list(message.get("notes", []))
                    )
                    completed += 1
                    free_workers.add(worker_id)
                    progress.worker(
                        worker_id,
                        row,
                        "Đã đăng và ghi Excel xong",
                    )
                    progress.total_done(completed)
                    print(
                        f"[OK] Worker {worker_id} xong dòng {row} "
                        f"| {completed}/{total_articles}"
                    )

                elif message_type == "validation_error":
                    row = int(message["row"])
                    active_rows.pop(worker_id, None)
                    finished_task = active_tasks.pop(worker_id, None)
                    if finished_task:
                        active_domains.discard(
                            str(finished_task["domain"])
                        )
                    apply_category_mappings(
                        list(message.get("mappings", []))
                    )
                    writes = list(message.get("writes", []))
                    if writes:
                        apply_publish_writes(writes, row)
                    else:
                        write_validation_failure(
                            row,
                            "LỖI KIỂM TRA: "
                            + str(message.get("error", "")),
                        )
                    apply_url_faq_notes(
                        list(message.get("notes", []))
                    )
                    completed += 1
                    had_error = True
                    free_workers.add(worker_id)
                    progress.worker(
                        worker_id,
                        row,
                        "Bỏ qua: kiểm tra trước Save không đạt",
                    )
                    progress.total_done(completed)
                    print(
                        f"[BỎ QUA — KHÔNG SAVE] Worker {worker_id}, "
                        f"dòng {row}: {message.get('error', '')}"
                    )
                    print(f"Log: {message.get('log_path', '')}")

                elif message_type == "blocked_406":
                    row = int(message["row"])
                    domain = str(message["domain"])
                    active_rows.pop(worker_id, None)
                    active_tasks.pop(
                        worker_id,
                        dict(message["task"]),
                    )
                    active_domains.discard(domain)
                    free_workers.discard(worker_id)
                    had_error = True
                    skipped_rows = [row]
                    kept_tasks: list[dict[str, Any]] = []
                    for pending_task in pending_tasks:
                        if str(pending_task["domain"]) == domain:
                            skipped_rows.append(int(pending_task["row"]))
                        else:
                            kept_tasks.append(pending_task)
                    pending_tasks[:] = kept_tasks
                    deferred_406_domains.discard(domain)
                    deferred_until.pop(domain, None)
                    mark_domain_skipped(
                        sorted(set(skipped_rows)),
                        "Bỏ qua domain trong phiên này vì website/WAF "
                        "trả về HTTP 406 Not Acceptable.",
                    )
                    progress.worker(
                        worker_id,
                        row,
                        "HTTP 406 — đã bỏ qua domain",
                    )
                    print(
                        f"[SKIP DOMAIN 406] {domain}: bỏ qua "
                        f"{len(set(skipped_rows))} bài trong phiên này; "
                        "các domain khác tiếp tục."
                    )
                    print(f"Log: {message.get('log_path', '')}")

                elif message_type == "task_error":
                    row = int(message.get("row", 0) or 0)
                    active_rows.pop(worker_id, None)
                    failed_task = active_tasks.pop(worker_id, None)
                    if failed_task:
                        active_domains.discard(
                            str(failed_task["domain"])
                        )
                    had_error = True
                    domain = str(message.get("domain", ""))
                    if bool(message.get("domain_access_error")):
                        stopping = True
                        if row:
                            try:
                                write_worker_failure(
                                    row,
                                    "Không xác nhận được trang đăng nhập/"
                                    f"form đăng bài: {message.get('error')}",
                                )
                            except Exception as write_error:
                                print(
                                    "Không ghi được lỗi truy cập CMS "
                                    f"dòng {row}: {write_error!r}"
                                )
                        print(
                            f"[DỪNG AN TOÀN] {domain}: không xác nhận được "
                            "trang đăng nhập/form đăng bài. "
                            "Đã khóa giao bài mới để người dùng kiểm tra."
                        )
                        progress.worker(
                            worker_id,
                            row or None,
                            "Không thấy login/form đăng bài — dừng an toàn",
                        )
                    elif row:
                        try:
                            writes = list(message.get("writes", []))
                            apply_category_mappings(
                                list(message.get("mappings", []))
                            )
                            if writes:
                                apply_publish_writes(writes, row)
                            else:
                                write_worker_failure(
                                    row, str(message.get("error", ""))
                                )
                        except Exception as write_error:
                            print(
                                f"Không ghi được lỗi dòng {row}: {write_error!r}"
                            )
                    if not bool(message.get("domain_access_error")):
                        print(
                            f"[LỖI BÀI — DỪNG AN TOÀN] Worker {worker_id}, "
                            f"dòng {row}: {message.get('error')}"
                        )
                        progress.worker(
                            worker_id,
                            row or None,
                            "Lỗi bài — giữ Edge để kiểm tra",
                        )
                    print(f"Log: {message.get('log_path', '')}")
                    if bool(message.get("domain_access_error")):
                        print(
                            "Các bài đang chạy sẽ hoàn tất; "
                            "không giao thêm bài mới."
                        )
                    else:
                        # Lỗi của một bài (kể cả không xác nhận được ID CMS)
                        # không khóa hàng đợi. Dòng lỗi đã được ghi Excel;
                        # trả worker về trạng thái rảnh để nhận bài tiếp theo.
                        stopping = False
                        free_workers.add(worker_id)
                        print(
                            "Đã ghi lỗi vào Excel; worker tiếp tục nhận bài mới."
                        )

                elif message_type == "fatal":
                    row = int(message.get("row", 0) or 0)
                    active_rows.pop(worker_id, None)
                    failed_task = active_tasks.pop(worker_id, None)
                    if failed_task:
                        active_domains.discard(
                            str(failed_task["domain"])
                        )
                    free_workers.discard(worker_id)
                    had_error = True
                    if row:
                        try:
                            write_worker_failure(
                                row, str(message.get("error", ""))
                            )
                        except Exception as write_error:
                            print(
                                f"Không ghi được lỗi dòng {row}: {write_error!r}"
                            )
                    print(
                        f"[WORKER {worker_id} ĐÃ DỪNG] "
                        f"{message.get('error')}"
                    )
                    progress.worker(
                        worker_id,
                        row or None,
                        "Worker đã dừng — worker khác tiếp tục",
                    )
                    print(f"Log: {message.get('log_path', '')}")

            if progress.stop_requested.is_set() and not stopping:
                stopping = True
                stopped_safely = True
                print(
                    "[DỪNG AN TOÀN] Đã khóa giao bài mới."
                )

            if not stopping:
                if time.time() >= next_dispatch_at:
                    for worker_id in sorted(list(free_workers)):
                        task = pop_next_domain_safe_task(
                            pending_tasks,
                            active_domains,
                            deferred_406_domains,
                            deferred_until,
                        )
                        if task is None:
                            break
                        verify_dispatch_task(task)
                        set_dispatch_status(int(task["row"]), worker_id)
                        command_queues[worker_id].put(task)
                        active_rows[worker_id] = int(task["row"])
                        active_tasks[worker_id] = task
                        active_domains.add(str(task["domain"]))
                        free_workers.remove(worker_id)
                        # Workers receive the next task immediately.  CMS
                        # protection now happens only at conditional Login and
                        # per-domain Save/ID confirmation.
                        delay_seconds = 0.0
                        next_dispatch_at = time.time()
                        progress.worker(
                            worker_id,
                            int(task["row"]),
                            f"Đang đăng: {task['domain']}",
                        )
                        print(
                            f"[GIAO] Worker {worker_id} <- "
                            f"dòng {task['row']} | {task['domain']} "
                            "| giao ngay cho worker rảnh"
                        )
                        continue

            all_dispatched = not pending_tasks
            if (all_dispatched or stopping) and not active_rows:
                break

            dead_active_workers = [
                (worker_id, row)
                for worker_id, row in active_rows.items()
                if not workers[worker_id].is_alive()
            ]
            if dead_active_workers:
                had_error = True
                for worker_id, row in dead_active_workers:
                    active_rows.pop(worker_id, None)
                    dead_task = active_tasks.pop(worker_id, None)
                    if dead_task:
                        active_domains.discard(
                            str(dead_task["domain"])
                        )
                    try:
                        write_worker_failure(
                            row,
                            f"Worker {worker_id} đã tắt bất thường, không trả kết quả.",
                        )
                    except Exception as write_error:
                        print(
                            f"Không ghi được lỗi dòng {row}: {write_error!r}"
                        )
                    print(
                        f"[LỖI] Worker {worker_id} tắt bất thường tại dòng {row}."
                    )
                    progress.worker(
                        worker_id,
                        row,
                        "Worker tắt bất thường — worker khác tiếp tục",
                    )

            if not any(process.is_alive() for process in workers.values()):
                if active_rows or pending_tasks:
                    had_error = True
                break
    finally:
        stop_workers(command_queues, workers)
        try:
            _EXCEL_WRITER.drain_and_stop()
        finally:
            _EXCEL_WRITER = None
        progress.finish(
            completed,
            stopped=stopped_safely,
        )

    print("=" * 72)
    print(
        f"KẾT THÚC: hoàn thành {completed}/{total_articles} bài"
    )
    print("=" * 72)
    return 1 if had_error else 0


def run_multi_with_error_log() -> int:
    """Chỉ tạo log điều phối khi phiên đa luồng gặp lỗi."""
    started_at = datetime.now()
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    memory_log = io.StringIO()
    exit_code = 1

    sys.stdout = publish.TeeOutput(original_stdout, memory_log)
    sys.stderr = publish.TeeOutput(original_stderr, memory_log)
    try:
        print(f"Thời gian bắt đầu: {started_at:%d/%m/%Y %H:%M:%S}")
        exit_code = run_multi()
    except BaseException:
        traceback.print_exc()
        exit_code = 1
    finally:
        print(
            f"Thời gian kết thúc: {datetime.now():%d/%m/%Y %H:%M:%S}"
        )
        sys.stdout = original_stdout
        sys.stderr = original_stderr

    if exit_code != 0:
        WORKER_LOG_ROOT.mkdir(parents=True, exist_ok=True)
        log_path = WORKER_LOG_ROOT / (
            f"dang_bai_dieu_phoi_{started_at:%Y%m%d_%H%M%S}.log"
        )
        log_path.write_text(
            memory_log.getvalue(),
            encoding="utf-8-sig",
        )
        if original_stdout is not None:
            print(
                f"Đã lưu log lỗi điều phối tại: {log_path}",
                file=original_stdout,
                flush=True,
            )

    return exit_code


if __name__ == "__main__":
    mp.freeze_support()
    raise SystemExit(run_multi_with_error_log())
