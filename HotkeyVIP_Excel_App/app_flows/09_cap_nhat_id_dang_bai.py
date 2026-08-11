from __future__ import annotations

import json
import os
from typing import Any


APP_WORKBOOK: Any = None


def _normalize(value: Any) -> str:
    return " ".join(str(value or "").strip().casefold().split())


def _find_column(sheet: Any, wanted: str) -> int:
    target = _normalize(wanted)
    last_column = int(sheet.Cells(1, sheet.Columns.Count).End(-4159).Column)
    for column in range(1, last_column + 1):
        if _normalize(sheet.Cells(1, column).Value) == target:
            return column
    raise RuntimeError(f'DANG_BAI không có cột "{wanted}".')


def main() -> int:
    if APP_WORKBOOK is None:
        raise RuntimeError("Flow cập nhật ID phải chạy từ app.")
    raw = os.environ.get("HOTKEYVIP_PUBLISH_ID_UPDATES", "").strip()
    updates = json.loads(raw) if raw else []
    if not isinstance(updates, list):
        raise RuntimeError("Danh sách cập nhật ID không hợp lệ.")

    sheet = APP_WORKBOOK.Worksheets("DANG_BAI")
    status_col = _find_column(sheet, "Trạng thái đăng")
    id_col = _find_column(sheet, "ID CMS")
    error_col = _find_column(sheet, "Lỗi đăng")
    for item in updates:
        row = int(item["row"])
        cms_id = str(item["cms_id"]).strip()
        if not cms_id.isdigit() or int(cms_id) <= 0:
            raise RuntimeError(f"Dòng {row}: ID CMS phải là số nguyên lớn hơn 0.")
        sheet.Cells(row, id_col).Value = int(cms_id)
        sheet.Cells(row, status_col).Value = "ĐÃ ĐĂNG"
        sheet.Cells(row, error_col).Value = ""
        print(f"[CẬP NHẬT ID] Dòng {row} -> ID CMS {cms_id}", flush=True)
    APP_WORKBOOK.Save()
    print(f"Đã cập nhật {len(updates)} ID CMS.", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
