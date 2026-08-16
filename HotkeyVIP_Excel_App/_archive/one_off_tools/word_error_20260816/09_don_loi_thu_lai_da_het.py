# -*- coding: utf-8 -*-
"""Dọn mã lỗi retry đã hết, giữ nguyên các lỗi còn thiếu file thật."""

import os
import re
import shutil
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import xlwings as xw
from docx import Document
from PIL import Image

TARGET = Path(r"D:\CodexProjects\Hotkeyvip\04_excel\hotkeyvip_test.xlsm")
SHEET_NAME = "VIET_BAI"
RETRY_HEADERS = ["Số lần thử lại", "Bước thử lại", "Lỗi thử lại", "Thời gian thử lại"]

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)


def valid_word(path):
    return bool(
        path
        and os.path.isfile(str(path).strip())
        and os.path.getsize(str(path).strip()) >= 10_000
    )


def valid_image(path):
    return bool(
        path
        and os.path.isfile(str(path).strip())
        and os.path.getsize(str(path).strip()) >= 10_000
    )


def find_open_book():
    target = os.path.normcase(os.path.abspath(TARGET))
    for app in xw.apps:
        for book in app.books:
            if os.path.normcase(os.path.abspath(book.fullname)) == target:
                return book, False
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    return app.books.open(str(TARGET), update_links=False, read_only=False), True


def clear_retry(sheet, row, headers):
    for name in RETRY_HEADERS:
        sheet.cells(row, headers[name]).value = ""


def main():
    book, owns_app = find_open_book()
    app = book.app
    sheet = book.sheets[SHEET_NAME]
    try:
        book.save()
        backup = TARGET.with_name(
            f"{TARGET.stem}_backup_truoc_don_loi_retry_{datetime.now():%Y%m%d_%H%M%S}{TARGET.suffix}"
        )
        shutil.copy2(TARGET, backup)

        used_values = sheet.used_range.value
        header_values = used_values[0]
        headers = {
            str(value or "").strip(): index
            for index, value in enumerate(header_values)
            if str(value or "").strip()
        }
        required = [
            "Tên Miền", "Từ khóa", "Đường dẫn Word", "Trạng thái viết",
            "Đường dẫn ảnh 1", "Đường dẫn ảnh 2", "Trạng thái hoàn tất",
            *RETRY_HEADERS,
        ]
        missing = [name for name in required if name not in headers]
        if missing:
            raise RuntimeError("Thiếu cột: " + ", ".join(missing))

        cleared = Counter()
        kept = Counter()
        cleared_rows = []
        kept_rows = []
        for value_index, values in enumerate(used_values[1:], start=1):
            row = value_index + 1
            error = str(values[headers["Lỗi thử lại"]] or "").strip()
            if not error:
                continue
            domain = str(values[headers["Tên Miền"]] or "").strip()
            keyword = str(values[headers["Từ khóa"]] or "").strip()
            article_status = str(values[headers["Trạng thái viết"]] or "").strip().upper()
            done = str(values[headers["Trạng thái hoàn tất"]] or "").strip().upper()
            word_path = values[headers["Đường dẫn Word"]]
            image_1 = values[headers["Đường dẫn ảnh 1"]]
            image_2 = values[headers["Đường dẫn ảnh 2"]]

            reason = None
            upper_error = error.upper()
            if done == "OK":
                reason = "DONE_OK"
            elif upper_error.startswith(("WORD_QUEUE_ERROR", "WORD_SYSTEM_ERROR", "WORD_SYSTEM_STOPPED")):
                if article_status == "OK" and valid_word(word_path):
                    reason = "WORD_DA_HOP_LE"
            elif "GEMINI" in upper_error or upper_error.startswith(("IMG1", "IMG2", "IMAGE")):
                if valid_image(image_1) and valid_image(image_2):
                    reason = "HAI_ANH_DA_HOP_LE"

            if reason:
                for name in RETRY_HEADERS:
                    sheet.cells(row, headers[name] + 1).value = ""
                cleared[domain] += 1
                cleared_rows.append((row, keyword, reason))
            else:
                kept[domain] += 1
                kept_rows.append((row, keyword, error[:100]))

        book.save()
        print(f"Đã xóa mã lỗi cũ: {len(cleared_rows)}")
        print(f"Giữ lỗi còn thực: {len(kept_rows)}")
        print("Đã xóa theo tên miền:", dict(cleared))
        print("Còn giữ theo tên miền:", dict(kept))
        print(f"Backup: {backup}")
        for row, keyword, error in kept_rows[:30]:
            print(f"GIỮ dòng {row}: {keyword} | {error}")
    finally:
        if owns_app:
            book.close()
            app.quit()


if __name__ == "__main__":
    main()
