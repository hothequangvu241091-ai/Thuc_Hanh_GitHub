# -*- coding: utf-8 -*-
"""V1.5 - Kế thừa quy luật đối chiếu 3 cột (Tiêu đề/Main Keyword + Tiêu đề SEO
+ H1) và các xử lý "Bài đã xóa" / "Đã viết" của V1.2, đồng thời khôi phục các
cơ chế an toàn đã có ở V1.0 mà V1.2 vô tình bỏ mất:

- Ghi file đích an toàn trên Google Drive ảo: SaveCopyAs ra ổ cục bộ, đóng
  workbook để nhả khóa, rồi copy đè lên file đích. KHÔNG Save() thẳng trên
  đường dẫn G:\\... (đây là nguyên nhân chính gây lỗi/khóa file ở V1.2).
- Tự dọn các tiến trình Excel ẩn (không cửa sổ) còn sót trước khi chạy.
- Nếu 1 file đích KHÔNG có cột "Main Keyword" thì tự động lùi về đối chiếu
  2 cột (Tiêu đề SEO + H1) như V1.0 cho riêng file đó, thay vì báo lỗi toàn
  bộ file / toàn bộ domain.
- Ghi log ra file (để xem lại khi chạy nền / double-click, không có console).
- Backup xoay vòng, không phình thư mục vô hạn theo từng lần chạy.
- Khóa tạm thời tương tác chuột/bàn phím trên Excel nguồn trong lúc chạy
  (Application.Interactive = False), để thao tác vô tình của người dùng
  không làm gián đoạn các lệnh COM đang chạy.
- Tự động thử lại (retry) khi gặp lỗi COM tạm thời (Excel "bận") thay vì
  chết ngang giữa chừng.

- Ưu tiên trạng thái xóa từ file đích: nếu ô "URL Page" chứa "Bài đã xóa"
  (ví dụ "Bài đã xóa (đợt 2)") thì cập nhật nguồn thành "Bài đã xóa" và
  không ghi đè URL Page, không chuyển URL, không ghi "Đã viết".

- Khi sheet đích được bảo vệ và chỉ mở khóa vùng nhập liệu, chỉ ghi các ô URL Page
  thực sự thay đổi. Không ghi lại ô tiêu đề hoặc toàn bộ cột URL Page, tránh Excel
  từ chối cả lệnh chỉ vì phạm vi ghi có chứa ô bị khóa.
"""
from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any
import gc
import os
import re
import shutil
import subprocess
import sys
import time
import traceback
import unicodedata
from urllib.parse import urlsplit

APP_WORKBOOK: Any = None

import pywintypes
import win32com.client as win32

# ----------------------------------------------------------------------------
# CẤU HÌNH
# ----------------------------------------------------------------------------
SOURCE_FILE = Path(
    os.environ.get(
        "HOTKEYVIP_SELECTED_EXCEL",
        r"D:\CodexProjects\Hotkeyvip\04_excel\hotkeyvip_test.xlsm",
    )
).resolve()
SOURCE_SHEET = "DANG_BAI"
TARGET_DIRECTORY = Path(
    r"G:\.shortcut-targets-by-id"
    r"\1Emi7P7uNkYpOjfn6wOnl9VeQbJesvcwP"
    r"\1.New-Content_Kênh SEO"
)

PROJECT_ROOT = SOURCE_FILE.parents[1]
BACKUP_ROOT = PROJECT_ROOT / "06_du_lieu_chay" / "backup_cap_nhat_url_cong_ty_v13"
LOG_ROOT = PROJECT_ROOT / "06_du_lieu_chay" / "log_cap_nhat_url_cong_ty_v13"
MAX_BACKUPS = 30

# Cột nguồn bắt buộc phải có (nếu thiếu 1 trong các cột này -> dừng hẳn,
# vì không thể chạy đúng logic đối chiếu / cập nhật status).
SOURCE_REQUIRED_HEADERS = {
    "keyword": "Tiêu đề",
    "title": "Tiêu đề SEO",
    "domain": "Tên miền",
    "h1": "H1",
    "url": "URL đã đăng",
}
# Cột nguồn tùy chọn: nếu không thấy thì coi như trống / tự thêm header.
SOURCE_OPTIONAL_HEADERS = {
    "word": "Đường dẫn Word",
    "status": "Đã chuyển",
}

# Cột đích: thử bộ đủ 4 cột (đúng quy luật V2) trước, nếu file đích không có
# "Main Keyword" thì lùi về bộ 3 cột (đúng quy luật V1.0) cho riêng file đó.
TARGET_HEADERS_FULL = {
    "keyword": "Main Keyword",
    "title": "Title [SEO]",
    "h1": "Article Name [H1]",
    "url": "URL Page",
}
TARGET_HEADERS_FALLBACK = {
    "title": "Title [SEO]",
    "h1": "Article Name [H1]",
    "url": "URL Page",
}

STATUS_OK = "OK"
STATUS_TRANSFERRED = "Đã chuyển"
STATUS_WRITTEN = "Đã viết"
STATUS_DELETED = "Bài đã xóa"
URL_WRITTEN_MARKER = "Đã viết"


# ----------------------------------------------------------------------------
# TIỆN ÍCH VĂN BẢN
# ----------------------------------------------------------------------------
def clean_text(value: Any) -> str:
    return re.sub(r"\s+", " ", "" if value is None else str(value)).strip()


def match_text(value: Any) -> str:
    return unicodedata.normalize("NFC", clean_text(value)).casefold()


def normalize_domain(value: Any) -> str:
    raw = clean_text(value).lower()
    if not raw:
        return ""
    parsed = urlsplit(raw if "://" in raw else "https://" + raw)
    domain = parsed.netloc or parsed.path.split("/", 1)[0]
    return domain.removeprefix("www.").rstrip("/")


def is_real_url(value: Any) -> bool:
    return match_text(value).startswith(("http://", "https://"))


def is_deleted(value: Any) -> bool:
    return "bài đã xóa" in match_text(value)


def com_retry(func, attempts: int = 5, delay_seconds: float = 1.5, label: str = ""):
    """Gọi func() và tự thử lại nếu Excel ném lỗi COM tạm thời (đang bận,
    tạm thời từ chối lệnh...). Sau khi hết số lần thử vẫn lỗi thì ném lại
    lỗi gốc để nơi gọi xử lý (không nuốt lỗi thật)."""
    last_exc: Exception | None = None
    for attempt in range(1, attempts + 1):
        try:
            return func()
        except pywintypes.com_error as exc:
            last_exc = exc
            if attempt == attempts:
                break
            print(f"    (COM tạm thời lỗi{' - ' + label if label else ''}, thử lại lần {attempt}/{attempts - 1}...)")
            time.sleep(delay_seconds)
    raise last_exc  # type: ignore[misc]


def matrix_from_com(value: Any) -> list[list[Any]]:
    if isinstance(value, tuple) and value and isinstance(value[0], tuple):
        return [list(row) for row in value]
    if isinstance(value, tuple):
        return [list(value)]
    return [[value]]


# ----------------------------------------------------------------------------
# LOG RA FILE (TEE) - giống V1.0
# ----------------------------------------------------------------------------
class TeeOutput:
    def __init__(self, console, log_file) -> None:
        self.console = console
        self.log_file = log_file

    def write(self, value) -> int:
        text_value = str(value)
        if self.console is not None:
            self.console.write(text_value)
            self.console.flush()
        self.log_file.write(text_value)
        self.log_file.flush()
        return len(text_value)

    def flush(self) -> None:
        if self.console is not None:
            self.console.flush()
        self.log_file.flush()

    def isatty(self) -> bool:
        return bool(self.console is not None and getattr(self.console, "isatty", lambda: False)())


# ----------------------------------------------------------------------------
# DỌN MÔI TRƯỜNG TRƯỚC KHI CHẠY - giống V1.0 (V1.2 đã bỏ mất phần này)
# ----------------------------------------------------------------------------
def close_orphan_hidden_excel() -> None:
    """Tắt các Excel chạy nền không có cửa sổ (bị treo từ lần chạy trước)."""
    command = (
        "$ids=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | "
        "Where-Object { $_.MainWindowHandle -eq 0 } | "
        "ForEach-Object { $id=$_.Id; "
        "Stop-Process -Id $id -Force -ErrorAction SilentlyContinue; $id }); "
        "$ids -join ','"
    )
    try:
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", command],
            capture_output=True, text=True, timeout=15, check=False,
        )
    except Exception:
        return
    closed_ids = result.stdout.strip()
    if closed_ids:
        print(f"Đã tự đóng Excel chạy nền bị sót: PID {closed_ids}")


def cleanup_old_backups() -> None:
    if not BACKUP_ROOT.is_dir():
        return
    backup_dirs = sorted(
        (item for item in BACKUP_ROOT.iterdir() if item.is_dir()),
        key=lambda item: item.name, reverse=True,
    )
    for old_dir in backup_dirs[MAX_BACKUPS:]:
        shutil.rmtree(old_dir, ignore_errors=True)


# ----------------------------------------------------------------------------
# TÌM HEADER
# ----------------------------------------------------------------------------
def find_headers(values: list[list[Any]], required: dict[str, str], max_rows: int = 10):
    """Trả về (row_index, {key: column_index}) nếu tìm đủ TẤT CẢ cột `required`
    trên cùng 1 dòng trong `max_rows` dòng đầu. Nếu không thấy đủ -> None
    (không raise, để nơi gọi tự quyết định lùi phương án nào)."""
    wanted = {match_text(label): name for name, label in required.items()}
    for row_index, row in enumerate(values[:max_rows]):
        found: dict[str, int] = {}
        for column_index, value in enumerate(row):
            normalized = match_text(value)
            if normalized in wanted:
                found[wanted[normalized]] = column_index
        if len(found) == len(required):
            return row_index, found
    return None


def find_header_column(values: list[list[Any]], header_row_index: int, header_name: str) -> int | None:
    expected = match_text(header_name)
    if header_row_index >= len(values):
        return None
    for column_index, value in enumerate(values[header_row_index]):
        if match_text(value) == expected:
            return column_index
    return None


# ----------------------------------------------------------------------------
# MỞ FILE NGUỒN (đúng cách của V1.0 / V1.2: ưu tiên workbook đang mở sẵn)
# ----------------------------------------------------------------------------
def get_source_workbook():
    expected_path = os.path.normcase(str(SOURCE_FILE.resolve()))
    if APP_WORKBOOK is not None:
        workbook_path = os.path.normcase(
            str(Path(str(APP_WORKBOOK.FullName)).resolve())
        )
        if workbook_path != expected_path:
            raise RuntimeError(
                f"App truyền sai workbook. Cần: {SOURCE_FILE}; nhận: {APP_WORKBOOK.FullName}"
            )
        return APP_WORKBOOK.Application, APP_WORKBOOK, False

    try:
        active_excel = win32.GetActiveObject("Excel.Application")
        for index in range(1, active_excel.Workbooks.Count + 1):
            workbook = active_excel.Workbooks(index)
            workbook_path = os.path.normcase(str(Path(str(workbook.FullName)).resolve()))
            if workbook_path == expected_path:
                return active_excel, workbook, False
    except Exception:
        pass

    hidden_excel = win32.DispatchEx("Excel.Application")
    hidden_excel.Visible = False
    hidden_excel.DisplayAlerts = False
    workbook = hidden_excel.Workbooks.Open(str(SOURCE_FILE.resolve()), 0, False)
    return hidden_excel, workbook, True


def confirm_run(row_count: int, domain_count: int) -> bool:
    from tkinter import Tk, messagebox
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return bool(messagebox.askyesno(
            "Xác nhận đồng bộ URL V1.5",
            f"Có {row_count} dòng thuộc {domain_count} tên miền cần kiểm tra.\n\n"
            "Chương trình sẽ backup trước, ghi file đích an toàn qua ổ cục bộ "
            "(không Save thẳng trên Google Drive), rồi mới cập nhật file nguồn.\n"
            "Bắt đầu?",
            parent=root,
        ))
    finally:
        root.destroy()


# ----------------------------------------------------------------------------
# ĐỌC NGUỒN
# ----------------------------------------------------------------------------
def read_source(source_workbook):
    worksheet = source_workbook.Worksheets(SOURCE_SHEET)
    used_range = worksheet.UsedRange
    values = matrix_from_com(used_range.Value)

    header_info = find_headers(values, SOURCE_REQUIRED_HEADERS, max_rows=10)
    if header_info is None:
        raise RuntimeError(
            "Không tìm thấy đủ cột bắt buộc ở nguồn: "
            + ", ".join(SOURCE_REQUIRED_HEADERS.values())
        )
    header_row_index, columns = header_info

    word_col = find_header_column(values, header_row_index, SOURCE_OPTIONAL_HEADERS["word"])
    status_col = find_header_column(values, header_row_index, SOURCE_OPTIONAL_HEADERS["status"])
    needs_status_header = status_col is None
    if status_col is None:
        status_col = len(values[header_row_index])

    return worksheet, used_range, values, header_row_index, columns, word_col, status_col, needs_status_header


# ----------------------------------------------------------------------------
# XỬ LÝ 1 FILE ĐÍCH
# ----------------------------------------------------------------------------
def find_target_headers(target_workbook):
    """Duyệt các sheet, thử bộ cột đầy đủ (4 cột, đúng quy luật V2) trước,
    nếu không có sheet nào đủ thì thử bộ cột rút gọn (3 cột, như V1.0).
    Trả về (sheet, used_range, values, header_row_index, columns, mode)
    với mode = "3cot" (có Main Keyword) hoặc "2cot" (fallback)."""
    for required, mode in ((TARGET_HEADERS_FULL, "3cot"), (TARGET_HEADERS_FALLBACK, "2cot")):
        for sheet_index in range(1, target_workbook.Worksheets.Count + 1):
            sheet = target_workbook.Worksheets(sheet_index)
            used_range = sheet.UsedRange
            values = matrix_from_com(used_range.Value)
            header_info = find_headers(values, required, max_rows=10)
            if header_info is not None:
                header_row_index, columns = header_info
                return sheet, used_range, values, header_row_index, columns, mode
    return None


def combo_key(row: list[Any], columns: dict[str, int], mode: str) -> tuple[str, ...]:
    if mode == "3cot":
        return (match_text(row[columns["keyword"]]), match_text(row[columns["title"]]), match_text(row[columns["h1"]]))
    return (match_text(row[columns["title"]]), match_text(row[columns["h1"]]))


def item_combo_key(item: dict[str, Any], mode: str) -> tuple[str, ...]:
    return item["combo3"] if mode == "3cot" else item["combo2"]


def process_domain(target_excel, site: str, items: list[dict[str, Any]], backup_directory: Path) -> tuple[Counter, dict[int, str]]:
    local_counts: Counter = Counter()
    status_updates: dict[int, str] = {}

    target_path = (TARGET_DIRECTORY / f"{site}.xlsx").resolve()
    if not target_path.is_file():
        local_counts["THIẾU FILE"] += len(items)
        print("    THIẾU FILE")
        return local_counts, status_updates

    target_workbook = None
    try:
        target_workbook = target_excel.Workbooks.Open(str(target_path), 0, False)
        if bool(target_workbook.ReadOnly):
            raise RuntimeError(f"File đích đang ReadOnly hoặc bị khóa: {target_path}")

        found = find_target_headers(target_workbook)
        if found is None:
            raise RuntimeError("Không tìm thấy sheet có đủ cột đích (kể cả bộ rút gọn 2 cột)")
        target_sheet, target_range, target_values, target_header_row, target_columns, mode = found
        print(f"    (đối chiếu theo {'3 cột: Tiêu đề+Tiêu đề SEO+H1' if mode == '3cot' else '2 cột: Tiêu đề SEO+H1 (không có Main Keyword)'})")

        # đếm trùng theo đúng combo đang dùng cho domain này
        combo_counts_this_mode: Counter = Counter(item_combo_key(item, mode) for item in items)

        target_index: dict[tuple[str, ...], list[int]] = defaultdict(list)
        for row_index in range(target_header_row + 1, len(target_values)):
            row = target_values[row_index]
            key = combo_key(row, target_columns, mode)
            if all(key):
                target_index[key].append(row_index)

        url_values = [row[target_columns["url"]] for row in target_values]
        changed_rows: set[int] = set()

        for item in items:
            key = item_combo_key(item, mode)
            if combo_counts_this_mode[key] > 1:
                local_counts["TRÙNG NGUỒN"] += 1
                continue
            matches = target_index.get(key, [])
            if len(matches) != 1:
                local_counts["KHÔNG TÌM THẤY" if not matches else "TRÙNG ĐÍCH"] += 1
                continue

            row_index = matches[0]
            current_target_url = clean_text(url_values[row_index])

            # Ưu tiên tuyệt đối trạng thái xóa ở file đích. Chỉ cần URL Page có
            # chứa "Bài đã xóa" (ví dụ "Bài đã xóa đợt 2") thì đồng bộ trạng
            # thái về nguồn và không ghi đè hay xử lý tiếp dòng này.
            if is_deleted(current_target_url):
                status_updates[item["source_row"]] = STATUS_DELETED
                local_counts["BÀI ĐÃ XÓA"] += 1
            elif is_real_url(current_target_url):
                status_updates[item["source_row"]] = STATUS_OK
                local_counts["ĐÃ CÓ URL"] += 1
            elif is_real_url(item["url"]):
                url_values[row_index] = item["url"]
                status_updates[item["source_row"]] = STATUS_TRANSFERRED
                local_counts["ĐÃ CHUYỂN URL"] += 1
                changed_rows.add(row_index)
            elif item["word"]:
                if match_text(current_target_url) != match_text(URL_WRITTEN_MARKER):
                    url_values[row_index] = URL_WRITTEN_MARKER
                    changed_rows.add(row_index)
                status_updates[item["source_row"]] = STATUS_WRITTEN
                local_counts["ĐÃ VIẾT"] += 1

        if changed_rows:
            # --- Ghi an toàn qua ổ cục bộ, KHÔNG Save() thẳng trên Drive ảo ---
            url_column_number = target_range.Column + target_columns["url"]

            # Chỉ ghi những hàng dữ liệu thực sự đổi. V1.4 từng ghi lại toàn bộ cột,
            # gồm cả ô tiêu đề; nếu tiêu đề bị khóa trên protected sheet thì Excel
            # từ chối cả lệnh dù các ô nhập liệu bên dưới đã được mở khóa.
            sorted_rows = sorted(changed_rows)
            changed_blocks: list[tuple[int, int]] = []
            block_start = block_end = sorted_rows[0]
            for row_index in sorted_rows[1:]:
                if row_index == block_end + 1:
                    block_end = row_index
                else:
                    changed_blocks.append((block_start, block_end))
                    block_start = block_end = row_index
            changed_blocks.append((block_start, block_end))

            def _write_url_column():
                for start_index, end_index in changed_blocks:
                    target_sheet.Range(
                        target_sheet.Cells(target_range.Row + start_index, url_column_number),
                        target_sheet.Cells(target_range.Row + end_index, url_column_number),
                    ).Value = tuple(
                        (url_values[index],)
                        for index in range(start_index, end_index + 1)
                    )

            com_retry(_write_url_column, label=f"ghi cột URL {target_path.name}")

            staged_path = backup_directory / f"DA_CAP_NHAT_{target_path.name}"
            target_workbook.SaveCopyAs(str(staged_path))
            target_workbook.Close(False)
            target_workbook = None
            # backup bản gốc (trước khi bị ghi đè) - chép ra ngay từ file đích hiện có trên đĩa
            shutil.copy2(target_path, backup_directory / target_path.name)
            shutil.copy2(staged_path, target_path)

        return local_counts, status_updates
    finally:
        if target_workbook is not None:
            target_workbook.Close(False)


# ----------------------------------------------------------------------------
# MAIN
# ----------------------------------------------------------------------------
def main() -> int:
    if os.environ.get("HOTKEYVIP_APP_RUN") != "1":
        close_orphan_hidden_excel()

    source_excel = None
    source_workbook = None
    owns_source = False
    target_excel = None
    interactive_locked = False
    try:
        source_excel, source_workbook, owns_source = get_source_workbook()
        if os.environ.get("HOTKEYVIP_APP_RUN") == "1":
            # Workbook do flow_host mở ẩn và quản lý. Tự Save qua COM thay vì
            # yêu cầu người dùng mở Excel chỉ để bấm Save.
            source_workbook.Save()
            if not bool(source_workbook.Saved):
                raise RuntimeError(
                    "App không thể lưu workbook ẩn trước khi đồng bộ URL."
                )
        elif not bool(source_workbook.Saved):
            raise RuntimeError("File hotkeyvip_test.xlsm đang có thay đổi chưa Save. Hãy Save rồi chạy lại.")

        (worksheet, used_range, values, header_row_index, columns,
         word_col, status_col, needs_status_header) = read_source(source_workbook)

        counts: Counter = Counter()
        status_updates: dict[int, str] = {}
        items_by_domain: dict[str, list[dict[str, Any]]] = defaultdict(list)

        for row_index in range(header_row_index + 1, len(values)):
            row = values[row_index]
            source_row = used_range.Row + row_index

            current_status = clean_text(row[status_col]) if status_col < len(row) else ""
            if match_text(current_status) == match_text(STATUS_OK):
                counts["BỎ QUA OK"] += 1
                continue

            source_url = clean_text(row[columns["url"]])
            if is_deleted(source_url):
                status_updates[source_row] = STATUS_DELETED
                counts["BÀI ĐÃ XÓA"] += 1
                continue

            word_path = clean_text(row[word_col]) if word_col is not None and word_col < len(row) else ""
            if not source_url and not word_path:
                counts["KHÔNG CÓ URL/WORD"] += 1
                continue

            site = normalize_domain(row[columns["domain"]])
            keyword = match_text(row[columns["keyword"]])
            title = match_text(row[columns["title"]])
            h1 = match_text(row[columns["h1"]])
            combo3 = (keyword, title, h1)
            combo2 = (title, h1)
            if not site or not title or not h1:
                counts["THIẾU COMBO"] += 1
                continue

            items_by_domain[site].append({
                "source_row": source_row, "domain": site,
                "combo3": combo3, "combo2": combo2,
                "url": source_url, "word": word_path,
            })

        total_items = sum(len(items) for items in items_by_domain.values())
        if not items_by_domain:
            print("Không có dòng nào cần xử lý.")
            return 0
        if not confirm_run(total_items, len(items_by_domain)):
            print("Đã hủy trước khi sửa dữ liệu.")
            return 0

        # Khóa tạm thời chuột/bàn phím trên Excel nguồn kể từ đây, để thao
        # tác vô tình của người dùng không làm gián đoạn các lệnh COM ghi
        # dữ liệu bên dưới. Luôn mở khóa lại ở khối finally, kể cả khi lỗi.
        try:
            source_excel.Interactive = False
            interactive_locked = True
        except Exception:
            pass

        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_directory = BACKUP_ROOT / stamp
        backup_directory.mkdir(parents=True, exist_ok=True)
        cleanup_old_backups()
        source_workbook.SaveCopyAs(str(backup_directory / SOURCE_FILE.name))

        if needs_status_header:
            worksheet.Cells(used_range.Row + header_row_index, used_range.Column + status_col).Value = SOURCE_OPTIONAL_HEADERS["status"]
            source_workbook.Save()

        target_excel = win32.DispatchEx("Excel.Application")
        target_excel.Visible = False
        target_excel.DisplayAlerts = False
        target_excel.AskToUpdateLinks = False

        domains = sorted(items_by_domain)
        for number, site in enumerate(domains, start=1):
            print(f"\n[{number}/{len(domains)}] {site}.xlsx")
            try:
                local_counts, local_status_updates = process_domain(target_excel, site, items_by_domain[site], backup_directory)
                counts.update(local_counts)
                status_updates.update(local_status_updates)
                print("    " + " | ".join(f"{name}: {amount}" for name, amount in sorted(local_counts.items())))
            except Exception as exc:
                counts["LỖI FILE ĐÍCH"] += 1
                print(f"    BỎ QUA DO LỖI: {type(exc).__name__}: {exc}")
                print("    Các tên miền khác vẫn tiếp tục; file này sẽ được thử lại ở lần chạy sau.")

        # Ghi status vào file nguồn 1 LẦN DUY NHẤT bằng mảng (không lặp COM
        # theo từng ô — với vài nghìn dòng, lặp Cells().Value từng ô rất dễ
        # bị Excel từ chối giữa chừng, đúng như lỗi đã gặp ở V1.3 bản trước),
        # có tự thử lại nếu Excel báo bận tạm thời.
        if status_updates:
            status_values = [row[status_col] if status_col < len(row) else "" for row in values]
            for source_row, status in status_updates.items():
                status_values[source_row - used_range.Row] = status
            status_column_number = used_range.Column + status_col

            def _write_status_column():
                worksheet.Range(
                    worksheet.Cells(used_range.Row, status_column_number),
                    worksheet.Cells(used_range.Row + len(status_values) - 1, status_column_number),
                ).Value = tuple((value,) for value in status_values)

            com_retry(_write_status_column, label="ghi cột Đã chuyển vào nguồn")
        com_retry(source_workbook.Save, label="Save file nguồn")

        print("\n" + "=" * 72)
        print(f"Backup: {backup_directory}")
        for name, amount in sorted(counts.items()):
            print(f"- {name}: {amount}")
        return 0
    finally:
        if interactive_locked and source_excel is not None:
            try:
                source_excel.Interactive = True
            except Exception:
                pass
        if target_excel is not None:
            target_excel.Quit()
        if owns_source and source_workbook is not None:
            source_workbook.Close(False)
        if owns_source and source_excel is not None:
            source_excel.Quit()
        gc.collect()


def run_with_log() -> int:
    LOG_ROOT.mkdir(parents=True, exist_ok=True)
    log_path = LOG_ROOT / f"dong_bo_url_v15_{datetime.now():%Y%m%d_%H%M%S}.log"
    original_stdout = sys.stdout
    original_stderr = sys.stderr
    exit_code = 1
    with log_path.open("w", encoding="utf-8-sig") as log_file:
        sys.stdout = TeeOutput(original_stdout, log_file)
        sys.stderr = TeeOutput(original_stderr, log_file)
        try:
            print(f"File log: {log_path}")
            exit_code = main()
        except BaseException as exc:
            print("\nDỪNG DO LỖI:")
            print(repr(exc))
            traceback.print_exc()
            exit_code = 1
        finally:
            sys.stdout = original_stdout
            sys.stderr = original_stderr
    print(f"Log: {log_path}")
    if os.environ.get("HOTKEYVIP_APP_RUN") != "1":
        try:
            input("\nNhấn Enter để đóng chương trình...")
        except (EOFError, KeyboardInterrupt):
            pass
    return exit_code


if __name__ == "__main__":
    raise SystemExit(run_with_log())
