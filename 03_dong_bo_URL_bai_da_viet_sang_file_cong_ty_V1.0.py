# -*- coding: utf-8 -*-
"""V1.0: Cập nhật URL Page thật, chạy độc lập, có backup và log."""

from __future__ import annotations

from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
import gc
import os
import re
import shutil
import subprocess
import sys
import traceback
import unicodedata
from typing import Any
from urllib.parse import urlsplit

import win32com.client as win32


STATUS_HEADER = "Đã chuyển"
STATUS_TRANSFERRED = "Đã chuyển"
STATUS_ALREADY_PRESENT = "OK"
STATUS_WRITTEN = "Đã viết"
URL_WRITTEN_MARKER = "Đã viết"
MAX_BACKUPS = 30

# Cấu hình và hàm hỗ trợ được nhúng tại đây để file chạy hoàn toàn độc lập.
SOURCE_FILE = Path(
    r"D:\CodexProjects\Hotkeyvip\04_excel\hotkeyvip_test.xlsm"
)
SOURCE_SHEET = "DANG_BAI"
SOURCE_HEADERS = {
    "domain": "Tên miền",
    "title": "Tiêu đề SEO",
    "h1": "H1",
    "url": "URL đã đăng",
}
TARGET_HEADERS = {
    "title": "Title [SEO]",
    "h1": "Article Name [H1]",
    "url": "URL Page",
}


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
        return bool(
            self.console is not None
            and getattr(self.console, "isatty", lambda: False)()
        )


def clean_text(value) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def match_text(value) -> str:
    """Chuẩn hóa Unicode, khoảng trắng và hoa/thường khi so khớp."""
    return unicodedata.normalize("NFC", clean_text(value)).casefold()


def normalize_domain(value) -> str:
    raw = clean_text(value).lower()
    if not raw:
        return ""
    if "://" not in raw:
        raw = "https://" + raw
    parsed = urlsplit(raw)
    domain = parsed.netloc or parsed.path.split("/", 1)[0]
    if domain.startswith("www."):
        domain = domain[4:]
    return domain.rstrip("/")

PROJECT_ROOT = Path(__file__).resolve().parents[1]
BACKUP_ROOT = (
    PROJECT_ROOT
    / "06_du_lieu_chay"
    / "backup_cap_nhat_url_cong_ty"
)
LOG_ROOT = (
    PROJECT_ROOT
    / "06_du_lieu_chay"
    / "log_cap_nhat_url_cong_ty"
)

# CHỈNH THƯ MỤC FILE EXCEL ĐÍCH TẠI ĐÂY.
# Mỗi tên miền sẽ tương ứng với một file: <tên-miền>.xlsx
TARGET_DIRECTORY = Path(
    r"G:\.shortcut-targets-by-id"
    r"\1Emi7P7uNkYpOjfn6wOnl9VeQbJesvcwP"
    r"\1.New-Content_Kênh SEO"
)


def cleanup_old_backups() -> None:
    """Chỉ giữ lại MAX_BACKUPS thư mục backup mới nhất."""
    if not BACKUP_ROOT.is_dir():
        return
    backup_dirs = sorted(
        (item for item in BACKUP_ROOT.iterdir() if item.is_dir()),
        key=lambda item: item.name,
        reverse=True,
    )
    for old_dir in backup_dirs[MAX_BACKUPS:]:
        shutil.rmtree(old_dir)

_OWNED_SOURCE_EXCEL = None
_OWNED_SOURCE_WORKBOOK = None
_OWNED_TARGET_EXCEL = None


def cleanup_hidden_excel() -> None:
    """Đóng đúng các Excel ẩn do chính chương trình này tạo."""
    global _OWNED_SOURCE_EXCEL
    global _OWNED_SOURCE_WORKBOOK
    global _OWNED_TARGET_EXCEL

    if _OWNED_TARGET_EXCEL is not None:
        try:
            _OWNED_TARGET_EXCEL.Quit()
        except Exception:
            pass
        _OWNED_TARGET_EXCEL = None

    if _OWNED_SOURCE_WORKBOOK is not None:
        try:
            _OWNED_SOURCE_WORKBOOK.Close(False)
        except Exception:
            pass
        _OWNED_SOURCE_WORKBOOK = None

    if _OWNED_SOURCE_EXCEL is not None:
        try:
            _OWNED_SOURCE_EXCEL.Quit()
        except Exception:
            pass
        _OWNED_SOURCE_EXCEL = None

    gc.collect()


def close_orphan_hidden_excel() -> None:
    """Tắt Excel chạy nền không có cửa sổ trước khi bắt đầu."""
    command = (
        "$ids=@(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | "
        "Where-Object { $_.MainWindowHandle -eq 0 } | "
        "ForEach-Object { $id=$_.Id; "
        "Stop-Process -Id $id -Force -ErrorAction SilentlyContinue; $id }); "
        "$ids -join ','"
    )
    result = subprocess.run(
        [
            "powershell.exe",
            "-NoProfile",
            "-Command",
            command,
        ],
        capture_output=True,
        text=True,
        timeout=15,
        check=False,
    )
    closed_ids = result.stdout.strip()
    if closed_ids:
        print(f"Đã tự đóng Excel chạy nền bị sót: PID {closed_ids}")


def matrix_from_com(value: Any) -> list[list[Any]]:
    if isinstance(value, tuple):
        if value and isinstance(value[0], tuple):
            return [list(row) for row in value]
        return [list(value)]
    return [[value]]


def find_headers_in_matrix(
    values: list[list[Any]],
    required: dict[str, str],
    max_rows: int = 10,
) -> tuple[int, dict[str, int]] | None:
    expected = {
        key: match_text(header)
        for key, header in required.items()
    }
    for row_index, row_values in enumerate(values[:max_rows]):
        found: dict[str, int] = {}
        for column_index, value in enumerate(row_values):
            normalized = match_text(value)
            for key, expected_header in expected.items():
                if normalized == expected_header:
                    found[key] = column_index
        if len(found) == len(required):
            return row_index, found
    return None


def find_header_column(
    values: list[list[Any]],
    header_row_index: int,
    header_name: str,
) -> int | None:
    expected = match_text(header_name)
    for column_index, value in enumerate(values[header_row_index]):
        if match_text(value) == expected:
            return column_index
    return None


def get_source_workbook():
    expected_path = os.path.normcase(
        str(SOURCE_FILE.resolve())
    )
    try:
        active_excel = win32.GetActiveObject("Excel.Application")
        for index in range(1, active_excel.Workbooks.Count + 1):
            workbook = active_excel.Workbooks(index)
            workbook_path = os.path.normcase(
                str(Path(str(workbook.FullName)).resolve())
            )
            if workbook_path == expected_path:
                return active_excel, workbook, False
    except Exception:
        pass

    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = excel.Workbooks.Open(
        str(SOURCE_FILE.resolve()),
        0,
        False,
    )
    return excel, workbook, True


def confirm_run(total_rows: int, total_domains: int) -> bool:
    from tkinter import Tk, messagebox

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        return bool(
            messagebox.askyesno(
                "Xác nhận cập nhật URL thật",
                f"Có {total_rows} dòng cần kiểm tra thuộc "
                f"{total_domains} tên miền.\n\n"
                "Chương trình sẽ:\n"
                "- Tạo backup trước khi sửa.\n"
                "- Ghi URL Page dạng Values.\n"
                "- Save từng file đích.\n"
                "- Chuyển thành công: ghi Đã chuyển.\n"
                "- URL đích đã có: ghi OK.\n"
                "- Save file nguồn.\n\n"
                "Bắt đầu chạy thật?",
                parent=root,
            )
        )
    finally:
        root.destroy()


def read_source_items(source_workbook):
    worksheet = source_workbook.Worksheets(SOURCE_SHEET)
    used_range = worksheet.UsedRange
    values = matrix_from_com(used_range.Value)
    header_info = find_headers_in_matrix(
        values,
        SOURCE_HEADERS,
        max_rows=5,
    )
    if header_info is None:
        raise RuntimeError(
            'Không tìm thấy đủ cột nguồn trong sheet "DANG_BAI".'
        )

    header_row_index, columns = header_info
    status_column_index = find_header_column(
        values,
        header_row_index,
        STATUS_HEADER,
    )
    word_column_index = find_header_column(
        values,
        header_row_index,
        "Đường dẫn Word",
    )
    status_column_number = (
        used_range.Column + status_column_index
        if status_column_index is not None
        else used_range.Column + len(values[header_row_index])
    )
    header_row_number = used_range.Row + header_row_index

    items: list[dict[str, Any]] = []
    skipped = Counter()
    for row_index in range(header_row_index + 1, len(values)):
        row_values = values[row_index]
        source_row_number = used_range.Row + row_index

        current_status = (
            clean_text(row_values[status_column_index])
            if status_column_index is not None
            and status_column_index < len(row_values)
            else ""
        )
        if current_status in {
            STATUS_TRANSFERRED,
            STATUS_ALREADY_PRESENT,
        }:
            skipped["CỘT ĐÃ CHUYỂN CÓ NỘI DUNG"] += 1
            continue

        source_url = clean_text(row_values[columns["url"]])
        word_path = (
            clean_text(row_values[word_column_index])
            if word_column_index is not None
            and word_column_index < len(row_values)
            else ""
        )
        if not source_url and not word_path:
            skipped["URL NGUỒN TRỐNG"] += 1
            continue

        domain = normalize_domain(row_values[columns["domain"]])
        title = clean_text(row_values[columns["title"]])
        h1 = clean_text(row_values[columns["h1"]])
        items.append(
            {
                "source_row": source_row_number,
                "domain": domain,
                "title": title,
                "h1": h1,
                "source_url": source_url,
                "word_path": word_path,
                "match_key": (
                    match_text(title),
                    match_text(h1),
                ),
            }
        )

    return (
        worksheet,
        header_row_number,
        status_column_number,
        status_column_index is None,
        items,
        skipped,
    )


def find_target_worksheet(target_workbook):
    sheet_names = [
        str(target_workbook.Worksheets(index).Name)
        for index in range(1, target_workbook.Worksheets.Count + 1)
    ]
    ordered_names = (
        ["Article"] + [name for name in sheet_names if name != "Article"]
        if "Article" in sheet_names
        else sheet_names
    )
    for sheet_name in ordered_names:
        worksheet = target_workbook.Worksheets(sheet_name)
        used_range = worksheet.UsedRange
        values = matrix_from_com(used_range.Value)
        header_info = find_headers_in_matrix(
            values,
            TARGET_HEADERS,
        )
        if header_info is not None:
            return worksheet, used_range, values, header_info
    return None


def process_domain(
    target_excel,
    domain: str,
    items: list[dict[str, Any]],
    source_key_counts: Counter,
    source_worksheet,
    status_column_number: int,
    backup_directory: Path,
) -> Counter:
    counts = Counter()
    target_path = (
        TARGET_DIRECTORY / f"{domain}.xlsx"
    ).resolve()
    if not target_path.is_file():
        print(f"    THIẾU FILE: {target_path.name}")
        counts["THIẾU FILE"] += len(items)
        return counts

    target_workbook = None
    try:
        target_workbook = target_excel.Workbooks.Open(
            str(target_path),
            0,
            False,
        )
        if bool(target_workbook.ReadOnly):
            raise RuntimeError(
                f"File đích đang ReadOnly hoặc bị khóa: {target_path}"
            )

        target_info = find_target_worksheet(target_workbook)
        if target_info is None:
            raise RuntimeError(
                f"Không tìm thấy đủ cột đích trong {target_path.name}"
            )
        worksheet, used_range, values, header_info = target_info
        header_row_index, columns = header_info

        index: dict[tuple[str, str], list[dict[str, Any]]] = defaultdict(list)
        for row_index in range(header_row_index + 1, len(values)):
            row_values = values[row_index]
            title = clean_text(row_values[columns["title"]])
            h1 = clean_text(row_values[columns["h1"]])
            if not title and not h1:
                continue
            target_row_number = used_range.Row + row_index
            target_url = clean_text(row_values[columns["url"]])
            index[
                (
                    match_text(title),
                    match_text(h1),
                )
            ].append(
                {
                    "row": target_row_number,
                    "url": target_url,
                }
            )

        changes: list[tuple[dict[str, Any], int, str, str]] = []
        completed_without_change: list[tuple[dict[str, Any], str]] = []
        for item in items:
            key = item["match_key"]
            if source_key_counts[(domain, key)] > 1:
                print(
                    f"    TRÙNG NGUỒN dòng {item['source_row']}: "
                    f"{item['title']}"
                )
                counts["TRÙNG NGUỒN"] += 1
                continue

            matches = index.get(key, [])
            if len(matches) != 1:
                result = (
                    "KHÔNG TÌM THẤY"
                    if not matches
                    else "TRÙNG ĐÍCH"
                )
                print(
                    f"    {result} dòng {item['source_row']}: "
                    f"{item['title']}"
                )
                counts[result] += 1
                continue

            match = matches[0]
            target_url = match["url"]
            is_written_marker = (
                match_text(target_url)
                == match_text(URL_WRITTEN_MARKER)
            )
            if target_url and not is_written_marker:
                completed_without_change.append(
                    (item, STATUS_ALREADY_PRESENT)
                )
                counts["ĐÍCH ĐÃ CÓ URL"] += 1
            elif item["source_url"]:
                # URL mới thay thế ô trống hoặc marker "Đã viết".
                changes.append(
                    (
                        item,
                        int(match["row"]),
                        item["source_url"],
                        STATUS_TRANSFERRED,
                    )
                )
            elif item["word_path"]:
                if is_written_marker:
                    completed_without_change.append(
                        (item, STATUS_WRITTEN)
                    )
                else:
                    changes.append(
                        (
                            item,
                            int(match["row"]),
                            URL_WRITTEN_MARKER,
                            STATUS_WRITTEN,
                        )
                    )
                counts["ĐÃ VIẾT CHƯA ĐĂNG"] += 1

        if changes:
            backup_path = backup_directory / target_path.name
            target_workbook.SaveCopyAs(str(backup_path))
            target_url_column_number = (
                used_range.Column + columns["url"]
            )
            for item, target_row_number, desired_url, _status in changes:
                # Gán Value thuần, tương đương Paste Values (123).
                worksheet.Cells(
                    target_row_number,
                    target_url_column_number,
                ).Value = desired_url

            # Excel không Save ổn định trực tiếp trên Google Drive ảo.
            # Lưu bản hoàn chỉnh ở ổ cục bộ, đóng workbook để nhả khóa,
            # rồi mới chép đè lên file đích. Backup gốc đã có ở trên.
            staged_path = (
                backup_directory / f"DA_CAP_NHAT_{target_path.name}"
            )
            target_workbook.SaveCopyAs(str(staged_path))
            target_workbook.Close(False)
            target_workbook = None
            shutil.copy2(staged_path, target_path)

            for item, _target_row_number, _desired_url, status in changes:
                source_worksheet.Cells(
                    item["source_row"],
                    status_column_number,
                ).Value = status
            counts["ĐÃ CHUYỂN"] += len(changes)

        for item, status in completed_without_change:
            source_worksheet.Cells(
                item["source_row"],
                status_column_number,
            ).Value = status
        return counts
    finally:
        if target_workbook is not None:
            target_workbook.Close(False)


def main() -> int:
    global _OWNED_SOURCE_EXCEL
    global _OWNED_SOURCE_WORKBOOK
    global _OWNED_TARGET_EXCEL

    source_excel = None
    source_workbook = None
    owns_source = False
    target_excel = None

    close_orphan_hidden_excel()
    source_excel, source_workbook, owns_source = get_source_workbook()
    if owns_source:
        _OWNED_SOURCE_EXCEL = source_excel
        _OWNED_SOURCE_WORKBOOK = source_workbook
    if not bool(source_workbook.Saved):
        raise RuntimeError(
            "File nguồn đang có thay đổi chưa Save. "
            "Hãy Save hotkeyvip_test.xlsm trước khi chạy lại."
        )

    (
        source_worksheet,
        header_row_number,
        status_column_number,
        needs_status_header,
        items,
        skipped,
    ) = read_source_items(source_workbook)
    items_by_domain: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for item in items:
        items_by_domain[item["domain"]].append(item)

    if not items:
        print("Không có dòng nào cần xử lý.")
        return 0
    if not confirm_run(len(items), len(items_by_domain)):
        print("Đã hủy trước khi sửa dữ liệu.")
        return 0

    run_stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_directory = BACKUP_ROOT / run_stamp
    backup_directory.mkdir(parents=True, exist_ok=True)
    cleanup_old_backups()
    source_workbook.SaveCopyAs(
        str(backup_directory / SOURCE_FILE.name)
    )

    if needs_status_header:
        source_worksheet.Cells(
            header_row_number,
            status_column_number,
        ).Value = STATUS_HEADER
        source_workbook.Save()

    target_excel = win32.DispatchEx("Excel.Application")
    _OWNED_TARGET_EXCEL = target_excel
    target_excel.Visible = False
    target_excel.DisplayAlerts = False
    target_excel.AskToUpdateLinks = False

    source_key_counts = Counter(
        (item["domain"], item["match_key"])
        for item in items
    )
    total_counts = Counter(skipped)
    domains = sorted(items_by_domain)
    try:
        for index, domain in enumerate(domains, start=1):
            print(
                f"\n[{index}/{len(domains)}] "
                f"Đang xử lý {domain}.xlsx"
            )
            counts = process_domain(
                target_excel,
                domain,
                items_by_domain[domain],
                source_key_counts,
                source_worksheet,
                status_column_number,
                backup_directory,
            )
            source_workbook.Save()
            total_counts.update(counts)
            print(
                "    "
                + " | ".join(
                    f"{name}: {count}"
                    for name, count in sorted(counts.items())
                )
            )
    finally:
        if target_excel is not None:
            target_excel.Quit()
            target_excel = None
            _OWNED_TARGET_EXCEL = None
            gc.collect()

    print("\n" + "=" * 72)
    print("HOÀN TẤT")
    print(f"Backup: {backup_directory}")
    for name, count in sorted(total_counts.items()):
        print(f"- {name}: {count}")
    return 0


def run_with_log() -> int:
    LOG_ROOT.mkdir(parents=True, exist_ok=True)
    log_path = LOG_ROOT / (
        f"cap_nhat_url_THUC_TE_{datetime.now():%Y%m%d_%H%M%S}.log"
    )
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
            cleanup_hidden_excel()
            sys.stdout = original_stdout
            sys.stderr = original_stderr

    print(f"Log: {log_path}")
    try:
        input("\nNhấn Enter để đóng chương trình...")
    except (EOFError, KeyboardInterrupt):
        pass
    return exit_code


if __name__ == "__main__":
    raise SystemExit(run_with_log())
