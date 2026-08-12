from __future__ import annotations

import argparse
import re
import shutil
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client as win32


def norm(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def valid_url(value: Any) -> bool:
    return norm(value).startswith(("http://", "https://"))


def matrix(value: Any, rows: int, cols: int) -> list[list[Any]]:
    if rows == 1 and cols == 1:
        return [[value]]
    if rows == 1:
        return [list(value)]
    if cols == 1:
        return [[item] for item in value]
    return [list(row) for row in value]


def headers(sheet: Any) -> dict[str, int]:
    last_col = int(sheet.Cells(1, sheet.Columns.Count).End(-4159).Column)
    values = matrix(sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, last_col)).Value2, 1, last_col)[0]
    return {norm(value): index for index, value in enumerate(values, 1) if norm(value)}


def last_row(sheet: Any, key_col: int) -> int:
    return int(sheet.Cells(sheet.Rows.Count, key_col).End(-4162).Row)


def combo(row: list[Any], columns: list[int]) -> tuple[str, str, str, str] | None:
    result = tuple(norm(row[column - 1]) for column in columns)
    return result if all(result) else None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", required=True)
    parser.add_argument("--output", required=True)
    args = parser.parse_args()
    source = Path(args.source).resolve()
    output = Path(args.output).resolve()
    if not source.is_file():
        raise RuntimeError(f"Không tìm thấy file nguồn: {source}")
    if output.exists():
        raise RuntimeError(f"File đầu ra đã tồn tại: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, output)

    pythoncom.CoInitialize()
    excel = None
    book = None
    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 3
        book = excel.Workbooks.Open(str(output), 0, False)
        ke = book.Worksheets("KE_HOACH")
        viet = book.Worksheets("VIET_BAI")
        dang = book.Worksheets("DANG_BAI")

        kh, vh, dh = headers(ke), headers(viet), headers(dang)
        # Vị trí cột đã kiểm tra trực tiếp trên workbook; dùng số cột để tránh
        # Excel COM đôi khi trả heading tiếng Việt sai bảng mã trong phiên ẩn.
        kcols = [4, 1, 6, 5]
        vcols = [3, 1, 2, 4]
        dcols = [5, 2, 8, 1]
        k_url, k_category = 3, 7
        v_completed, v_word, v_image1, v_image2 = 27, 8, 22, 24
        d_status, d_id, d_url, d_code, d_time, d_error, d_transferred, d_note = 4, 9, 10, 14, 15, 16, 17, 18
        d_category, d_word, d_image1, d_image2 = 6, 11, 12, 13
        kmax, vmax, dmax = 12, 34, 18
        klast, vlast, dlast = last_row(ke, kcols[1]), last_row(viet, vcols[1]), last_row(dang, dcols[1])
        kdata = matrix(ke.Range(ke.Cells(2, 1), ke.Cells(klast, kmax)).Value2, klast - 1, kmax)
        vdata = matrix(viet.Range(viet.Cells(2, 1), viet.Cells(vlast, vmax)).Value2, vlast - 1, vmax)
        ddata = matrix(dang.Range(dang.Cells(2, 1), dang.Cells(dlast, dmax)).Value2, dlast - 1, dmax)

        viet_by_combo: dict[tuple[str, str, str, str], tuple[int, list[Any], bool]] = {}
        for sheet_row, row in enumerate(vdata, 2):
            key = combo(row, vcols)
            if key:
                viet_by_combo[key] = (sheet_row, row, norm(row[v_completed - 1]) == "ok")

        dang_by_combo: dict[tuple[str, str, str, str], tuple[int, list[Any]]] = {}
        dang_by_url: dict[str, tuple[int, list[Any], tuple[str, str, str, str]]] = {}
        for sheet_row, row in enumerate(ddata, 2):
            key = combo(row, dcols)
            if not key:
                continue
            dang_by_combo[key] = (sheet_row, row)
            url_value = row[d_url - 1]
            if valid_url(url_value):
                dang_by_url[norm(url_value)] = (sheet_row, row, key)

        ke_by_combo: dict[tuple[str, str, str, str], list[Any]] = {}
        missing: list[tuple[tuple[str, str, str, str], list[Any], int, list[Any]]] = []
        unmatched: list[tuple[tuple[str, str, str, str], list[Any]]] = []
        for row in kdata:
            key = combo(row, kcols)
            if key:
                ke_by_combo[key] = row
            url_value = row[k_url - 1]
            if not key or not valid_url(url_value) or key in dang_by_combo:
                continue
            source_result = dang_by_url.get(norm(url_value))
            if source_result:
                source_row, source_values, _wrong_key = source_result
                missing.append((key, row, source_row, source_values))
            else:
                unmatched.append((key, row))

        if len(missing) != 277 or len(unmatched) != 1:
            raise RuntimeError(f"Đối chiếu không đúng trạng thái dự kiến: khớp URL={len(missing)}, chưa khớp={len(unmatched)}")

        # Trường hợp duy nhất URL KE_HOACH là link rút gọn: ghép với dòng kết quả
        # có slug bài MVP, sau khi đã xác nhận thủ công nội dung bài tương ứng.
        unmatched_key, unmatched_ke = unmatched[0]
        mvp_source = next(
            (item for item in dang_by_url.values() if "cach-xay-dung-mvp-de-giam-rui-ro-khi-khoi-nghiep" in norm(item[1][d_url - 1])),
            None,
        )
        if mvp_source is None:
            raise RuntimeError("Không tìm thấy kết quả đăng của bài MVP dùng URL rút gọn")
        missing.append((unmatched_key, unmatched_ke, mvp_source[0], mvp_source[1]))

        result_columns = [
            d_status, d_id, d_url, d_code, d_time, d_error, d_transferred, d_note,
        ]
        wrong_results: dict[tuple[str, str, str, str], list[Any]] = {
            target_key: source_values for target_key, _ke_row, _source_row, source_values in missing
        }
        wrong_source_keys = {
            combo(source_values, dcols) for _target_key, _ke_row, _source_row, source_values in missing
        }
        if len(wrong_results) != 278:
            raise RuntimeError(
                f"Không xác định đủ 278 cặp kết quả bị lệch: "
                f"đích={len(wrong_results)}, nguồn={len(wrong_source_keys)}"
            )

        # Dựng lại toàn bộ DANG_BAI theo đúng tập combo VIET_BAI đã OK.
        # Dòng cũ đúng combo được giữ nguyên; 278 cụm kết quả sai được tháo ra và
        # gắn vào đúng combo theo URL/đối chiếu đã xác nhận.
        max_col = dmax
        desired_rows: list[list[Any]] = []
        for _sheet_row, viet_values, completed in viet_by_combo.values():
            if not completed:
                continue
            key = combo(viet_values, vcols)
            if key is None or key not in ke_by_combo:
                raise RuntimeError("VIET_BAI OK không có combo tương ứng trong KE_HOACH")
            ke_row = ke_by_combo[key]
            if key in dang_by_combo:
                values = list(dang_by_combo[key][1])
            else:
                values = [""] * max_col
                values[dcols[3] - 1] = ke_row[kcols[3] - 1]
                values[dcols[1] - 1] = ke_row[kcols[1] - 1]
                values[dcols[0] - 1] = ke_row[kcols[0] - 1]
                values[d_category - 1] = ke_row[k_category - 1]
                values[dcols[2] - 1] = ke_row[kcols[2] - 1]
                values[d_word - 1] = viet_values[v_word - 1]
                values[d_image1 - 1] = viet_values[v_image1 - 1]
                values[d_image2 - 1] = viet_values[v_image2 - 1]
            if key in wrong_source_keys:
                for column in result_columns:
                    values[column - 1] = ""
            if key in wrong_results:
                source_values = wrong_results[key]
                for column in result_columns:
                    values[column - 1] = source_values[column - 1]
            desired_rows.append(values)

        if len(desired_rows) != 5361:
            raise RuntimeError(f"Số dòng DANG_BAI dựng lại không đúng: {len(desired_rows)}")
        clear_last = max(int(dang.UsedRange.Rows.Count), len(desired_rows) + 1)
        dang.Range(dang.Cells(2, 1), dang.Cells(clear_last, max_col)).ClearContents()
        dang.Range(dang.Cells(2, 1), dang.Cells(len(desired_rows) + 1, max_col)).Value2 = tuple(
            tuple(row) for row in desired_rows
        )

        book.Save()
        print(f"OUTPUT={output}")
        print("REBUILT_DANG_BAI_ROWS=5361")
        print("REALIGNED_POSTED_RESULTS=278")
        return 0
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
