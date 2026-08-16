# -*- coding: utf-8 -*-
from pathlib import Path
import sys
import xlwings as xw

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

TARGET = Path(r"D:\CodexProjects\Hotkeyvip\04_excel\hotkeyvip_test.xlsm")
KEYWORD = "tính nhất quán phương pháp luận"
WORD_PATH = r"D:\CodexProjects\Hotkeyvip\07_ket_qua\bai_viet\bantinkhoahoc.com\tính nhất quán phương pháp luận.docx"

book = None
for app in xw.apps:
    for candidate in app.books:
        if Path(candidate.fullname).resolve() == TARGET.resolve():
            book = candidate
            break
    if book is not None:
        break
if book is None:
    raise RuntimeError("Không tìm thấy workbook đích trong Excel đang mở")

sheet = book.sheets["VIET_BAI"]
last_col = sheet.used_range.last_cell.column
header_values = sheet.range((1, 1), (1, last_col)).value
headers = {
    str(value or "").strip(): index + 1
    for index, value in enumerate(header_values)
    if str(value or "").strip()
}
target_row = None
for row in range(2, sheet.used_range.last_cell.row + 1):
    if str(sheet.cells(row, headers["Từ khóa"]).value or "").strip() == KEYWORD:
        target_row = row
        break
if target_row is None:
    raise RuntimeError("Không tìm thấy dòng từ khóa")

sheet.cells(target_row, headers["Đường dẫn Word"]).value = WORD_PATH
sheet.cells(target_row, headers["Trạng thái viết"]).value = "OK"
sheet.cells(target_row, headers["Lỗi viết"]).value = ""
sheet.cells(target_row, headers["Số từ Word"]).value = 3222
book.save()
print(f"Đã cập nhật trực tiếp workbook đang mở: dòng {target_row}")
print(
    "Trạng thái tại chỗ:",
    sheet.cells(target_row, headers["Trạng thái viết"]).value,
    "| Lỗi:",
    repr(sheet.cells(target_row, headers["Lỗi viết"]).value),
)
