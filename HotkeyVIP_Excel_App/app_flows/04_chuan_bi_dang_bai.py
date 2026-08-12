# -*- coding: utf-8 -*-
"""Đưa bài OK từ VIET_BAI sang DANG_BAI theo combo 4 cột, không nhân dòng."""

from __future__ import annotations

import re
from typing import Any

import win32com.client as win32


# Khi chạy từ app, flow_host gán trực tiếp workbook Excel ẩn vào biến này.
# Chạy độc lập vẫn dùng workbook đang mở như trước.
APP_WORKBOOK: Any = None


SHEET_PUBLISH = "DANG_BAI"
SHEET_PLAN = "KE_HOACH"
SHEET_WRITE = "VIET_BAI"

HEADER_SEO_PUBLISH = "Tiêu đề SEO"
HEADER_TITLE_PUBLISH = "Tiêu đề"
HEADER_DOMAIN_PUBLISH = "Tên miền"
HEADER_CATEGORY_PUBLISH = "Danh mục"
HEADER_H1_PUBLISH = "H1"
HEADER_WORD_PUBLISH = "Đường dẫn Word"
HEADER_IMAGE1_PUBLISH = "Đường dẫn ảnh 1"
HEADER_IMAGE2_PUBLISH = "Đường dẫn ảnh 2"
HEADER_DUPLICATE_NOTE_PUBLISH = "Ghi chú kiểm tra"

HEADER_SEO_PLAN = "Title [SEO]"
HEADER_DOMAIN_PLAN = "Tên Miền"
HEADER_KEYWORD_PLAN = "Main Keyword"
HEADER_URL_PAGE_PLAN = "URL Page"
HEADER_CATEGORY_PLAN = "POST / UPDATE"
HEADER_CATEGORY_PLAN_ALTERNATIVE = "CATE [POST]"
HEADER_H1_PLAN = "Article Name [H1]"

HEADER_SEO_WRITE = "Tiêu đề SEO"
HEADER_DOMAIN_WRITE = "Tên Miền"
HEADER_TITLE_WRITE = "Từ khóa"
HEADER_H1_WRITE = "H1"
HEADER_COMPLETED_STATUS_WRITE = "Trạng thái hoàn tất"
HEADER_WORD_WRITE = "Đường dẫn Word"
HEADER_OLD_WORD_COUNT_WRITE = "Số từ Word"
HEADER_NEW_WORD_WRITE = "Đường dẫn bài viết mới"
HEADER_NEW_WORD_COUNT_WRITE = "Số từ bài viết mới"
HEADER_IMAGE1_WRITE = "Đường dẫn ảnh 1"
HEADER_IMAGE2_WRITE = "Đường dẫn ảnh 2"

MIN_EXTRA_WORDS = 50
MAX_SAFE_PUBLISH_ROW = 50_000
# Excel/VBA RGB(255, 255, 153): vàng nhạt, báo đang dùng Word đợt 2.
NEW_WORD_HIGHLIGHT_COLOR = 10092543
DUPLICATE_NOTE = "Có trùng"
DELETED_MARKER = "bài đã xóa"


def normalize(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip()).casefold()


def is_usable_key(value: str) -> bool:
    return bool(value) and value not in {"#", "-", "#.n/a", "#n/a", "n/a", "na"}


def has_value(value: Any) -> bool:
    return value is not None and bool(str(value).strip())


def is_real_url(value: Any) -> bool:
    return bool(re.match(r"^https?://", str(value or "").strip(), re.IGNORECASE))


def to_number(value: Any) -> float | None:
    if not has_value(value):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def headers(ws: Any) -> dict[str, int]:
    last_col = int(ws.Cells(1, ws.Columns.Count).End(-4159).Column)
    result: dict[str, int] = {}
    for col in range(1, last_col + 1):
        value = ws.Cells(1, col).Value
        if has_value(value):
            result[normalize(value)] = col
    return result


def require_columns(ws: Any, names: list[str]) -> dict[str, int]:
    available = headers(ws)
    missing = [name for name in names if normalize(name) not in available]
    if missing:
        raise RuntimeError(f"Sheet {ws.Name} thiếu tiêu đề: " + ", ".join(missing))
    return {name: available[normalize(name)] for name in names}


def require_any_column(ws: Any, names: list[str]) -> int:
    available = headers(ws)
    matches = [(available[normalize(name)], name) for name in names if normalize(name) in available]
    if not matches:
        raise RuntimeError(f"Sheet {ws.Name} thiếu tiêu đề: " + " hoặc ".join(names))
    matches.sort()
    return matches[0][0]


def ensure_column(ws: Any, name: str) -> int:
    available = headers(ws)
    existing = available.get(normalize(name))
    if existing:
        return existing

    last_col = int(ws.Cells(1, ws.Columns.Count).End(-4159).Column)
    new_col = last_col + 1
    ws.Cells(1, new_col).Value = name
    # Giữ định dạng header nhất quán với cột ngay trước đó.
    ws.Cells(1, last_col).Copy(ws.Cells(1, new_col))
    ws.Cells(1, new_col).Value = name
    return new_col


def last_data_row(ws: Any, col: int) -> int:
    found = ws.Columns(col).Find(
        What="*", After=ws.Cells(1, col), LookIn=-4163, LookAt=2,
        SearchOrder=1, SearchDirection=2, MatchCase=False,
    )
    return 1 if found is None else int(found.Row)


def assert_publish_range_is_safe(ws: Any, last_row: int) -> None:
    if last_row > MAX_SAFE_PUBLISH_ROW:
        raise RuntimeError(
            f"Sheet {ws.Name} đang nhận dòng dữ liệu cuối là {last_row:,}, "
            f"vượt ngưỡng an toàn {MAX_SAFE_PUBLISH_ROW:,}. Chưa ghi dữ liệu."
        )
    for index in range(1, int(ws.ListObjects.Count) + 1):
        table = ws.ListObjects.Item(index)
        table_last = int(table.Range.Row + table.Range.Rows.Count - 1)
        if table_last > MAX_SAFE_PUBLISH_ROW:
            raise RuntimeError(
                f"Excel Table '{table.Name}' trên sheet {ws.Name} đang kéo tới dòng "
                f"{table_last:,}. Hãy thu Table về vùng dữ liệu thật trước khi chạy lại."
            )


def read_matrix(ws: Any, first_row: int, last_row: int, first_col: int, last_col: int) -> tuple:
    values = ws.Range(ws.Cells(first_row, first_col), ws.Cells(last_row, last_col)).Value2
    if not isinstance(values, tuple):
        return ((values,),)
    return (values,) if values and not isinstance(values[0], tuple) else values


def make_key(
    seo: Any, domain: Any, title: Any, h1: Any
) -> tuple[str, str, str, str] | None:
    key = (normalize(seo), normalize(domain), normalize(title), normalize(h1))
    return key if all(is_usable_key(part) for part in key) else None


def read_grouped_rows(
    ws: Any,
    key_cols: tuple[int, int, int, int],
    value_cols: dict[str, int],
) -> dict[tuple[str, str, str, str], list[dict[str, Any]]]:
    last_row = max(last_data_row(ws, col) for col in key_cols)
    if last_row < 2:
        return {}
    all_cols = [*key_cols, *value_cols.values()]
    first_col, final_col = min(all_cols), max(all_cols)
    rows: dict[tuple[str, str, str, str], list[dict[str, Any]]] = {}
    for row_number, values in enumerate(read_matrix(ws, 2, last_row, first_col, final_col), start=2):
        key = make_key(
            values[key_cols[0] - first_col],
            values[key_cols[1] - first_col],
            values[key_cols[2] - first_col],
            values[key_cols[3] - first_col],
        )
        if key is None:
            continue
        item = {"row": row_number, "key": key}
        item.update({name: values[col - first_col] for name, col in value_cols.items()})
        rows.setdefault(key, []).append(item)
    return rows


def select_word_path(write_item: dict[str, Any]) -> tuple[Any, bool]:
    old_count = to_number(write_item["old_word_count"])
    new_count = to_number(write_item["new_word_count"])
    use_new = (
        has_value(write_item["new_word_path"])
        and old_count is not None
        and new_count is not None
        and new_count >= old_count + MIN_EXTRA_WORDS
    )
    return (write_item["new_word_path"] if use_new else write_item["old_word_path"], use_new)


def plan_item_can_publish(plan_item: dict[str, Any]) -> bool:
    url_page = normalize(plan_item["url_page"])
    return url_page != DELETED_MARKER and not is_real_url(plan_item["url_page"])


def main() -> None:
    print("=" * 70)
    print("CHUẨN BỊ DANG_BAI V1.2 - ĐỒNG BỘ BÀI OK, GHI CÓ TRÙNG TẠI ĐÍCH")
    print("=" * 70)

    workbook = APP_WORKBOOK
    if workbook is None:
        excel = win32.GetActiveObject("Excel.Application")
        workbook = excel.ActiveWorkbook
    else:
        excel = workbook.Application
    if workbook is None:
        raise RuntimeError("Không tìm thấy workbook Excel đang mở.")

    ws_publish = workbook.Worksheets(SHEET_PUBLISH)
    ws_plan = workbook.Worksheets(SHEET_PLAN)
    ws_write = workbook.Worksheets(SHEET_WRITE)

    publish_cols = require_columns(ws_publish, [
        HEADER_SEO_PUBLISH, HEADER_TITLE_PUBLISH, HEADER_DOMAIN_PUBLISH,
        HEADER_CATEGORY_PUBLISH, HEADER_H1_PUBLISH, HEADER_WORD_PUBLISH,
        HEADER_IMAGE1_PUBLISH, HEADER_IMAGE2_PUBLISH,
    ])
    publish_cols[HEADER_DUPLICATE_NOTE_PUBLISH] = ensure_column(
        ws_publish, HEADER_DUPLICATE_NOTE_PUBLISH
    )
    plan_cols = require_columns(ws_plan, [
        HEADER_SEO_PLAN, HEADER_DOMAIN_PLAN, HEADER_KEYWORD_PLAN,
        HEADER_URL_PAGE_PLAN, HEADER_H1_PLAN,
    ])
    plan_cols[HEADER_CATEGORY_PLAN] = require_any_column(
        ws_plan, [HEADER_CATEGORY_PLAN, HEADER_CATEGORY_PLAN_ALTERNATIVE]
    )
    write_cols = require_columns(ws_write, [
        HEADER_SEO_WRITE, HEADER_DOMAIN_WRITE, HEADER_TITLE_WRITE, HEADER_H1_WRITE,
        HEADER_COMPLETED_STATUS_WRITE, HEADER_WORD_WRITE,
        HEADER_OLD_WORD_COUNT_WRITE, HEADER_NEW_WORD_WRITE,
        HEADER_NEW_WORD_COUNT_WRITE, HEADER_IMAGE1_WRITE, HEADER_IMAGE2_WRITE,
    ])

    plan_rows = read_grouped_rows(
        ws_plan,
        (plan_cols[HEADER_SEO_PLAN], plan_cols[HEADER_DOMAIN_PLAN],
         plan_cols[HEADER_KEYWORD_PLAN], plan_cols[HEADER_H1_PLAN]),
        {"category": plan_cols[HEADER_CATEGORY_PLAN],
         "url_page": plan_cols[HEADER_URL_PAGE_PLAN]},
    )
    write_rows = read_grouped_rows(
        ws_write,
        (write_cols[HEADER_SEO_WRITE], write_cols[HEADER_DOMAIN_WRITE],
         write_cols[HEADER_TITLE_WRITE], write_cols[HEADER_H1_WRITE]),
        {"seo": write_cols[HEADER_SEO_WRITE], "domain": write_cols[HEADER_DOMAIN_WRITE],
         "title": write_cols[HEADER_TITLE_WRITE], "h1": write_cols[HEADER_H1_WRITE],
         "completed_status": write_cols[HEADER_COMPLETED_STATUS_WRITE],
         "old_word_path": write_cols[HEADER_WORD_WRITE], "old_word_count": write_cols[HEADER_OLD_WORD_COUNT_WRITE],
         "new_word_path": write_cols[HEADER_NEW_WORD_WRITE], "new_word_count": write_cols[HEADER_NEW_WORD_COUNT_WRITE],
         "image1": write_cols[HEADER_IMAGE1_WRITE], "image2": write_cols[HEADER_IMAGE2_WRITE]},
    )
    publish_rows = read_grouped_rows(
        ws_publish,
        (publish_cols[HEADER_SEO_PUBLISH], publish_cols[HEADER_DOMAIN_PUBLISH],
         publish_cols[HEADER_TITLE_PUBLISH], publish_cols[HEADER_H1_PUBLISH]),
        {"duplicate_note": publish_cols[HEADER_DUPLICATE_NOTE_PUBLISH]},
    )

    last_row = max(
        last_data_row(ws_publish, publish_cols[HEADER_SEO_PUBLISH]),
        last_data_row(ws_publish, publish_cols[HEADER_DOMAIN_PUBLISH]),
        last_data_row(ws_publish, publish_cols[HEADER_TITLE_PUBLISH]),
    )
    assert_publish_range_is_safe(ws_publish, last_row)

    added_rows: list[dict[str, Any]] = []
    added_by_key: dict[tuple[str, str, str, str], list[dict[str, Any]]] = {}
    used_new_word = 0

    for key, items in write_rows.items():
        # Một dòng OK tương ứng một vị trí cần có ở DANG_BAI.
        completed_items = [item for item in items if normalize(item["completed_status"]) == "ok"]
        if not completed_items:
            continue

        # Các dòng Bài đã xóa / có URL thật không thể tạo nhiệm vụ đăng mới.
        active_plan_items = [item for item in plan_rows.get(key, []) if plan_item_can_publish(item)]
        if len(active_plan_items) != 1:
            continue
        plan_item = active_plan_items[0]

        existing_count = len(publish_rows.get(key, []))
        for write_item in completed_items[existing_count:]:
            word_path, use_new_word = select_word_path(write_item)
            if use_new_word:
                used_new_word += 1
            item = {
                "key": key,
                HEADER_SEO_PUBLISH: write_item["seo"],
                HEADER_TITLE_PUBLISH: write_item["title"],
                HEADER_DOMAIN_PUBLISH: write_item["domain"],
                HEADER_CATEGORY_PUBLISH: plan_item["category"],
                HEADER_H1_PUBLISH: write_item["h1"],
                HEADER_WORD_PUBLISH: word_path,
                HEADER_IMAGE1_PUBLISH: write_item["image1"],
                HEADER_IMAGE2_PUBLISH: write_item["image2"],
                "use_new_word": use_new_word,
            }
            added_rows.append(item)
            added_by_key.setdefault(key, []).append(item)

    final_last_row = last_row + len(added_rows)
    assert_publish_range_is_safe(ws_publish, final_last_row)

    old_screen_updating = excel.ScreenUpdating
    old_enable_events = excel.EnableEvents
    try:
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        output_headers = [
            HEADER_SEO_PUBLISH, HEADER_TITLE_PUBLISH, HEADER_DOMAIN_PUBLISH,
            HEADER_CATEGORY_PUBLISH, HEADER_H1_PUBLISH, HEADER_WORD_PUBLISH,
            HEADER_IMAGE1_PUBLISH, HEADER_IMAGE2_PUBLISH,
        ]
        if added_rows:
            first_new_row = last_row + 1
            for header in output_headers:
                ws_publish.Range(
                    ws_publish.Cells(first_new_row, publish_cols[header]),
                    ws_publish.Cells(final_last_row, publish_cols[header]),
                ).Value2 = tuple((item[header],) for item in added_rows)
            for offset, item in enumerate(added_rows):
                row = first_new_row + offset
                item["row"] = row
                if item["use_new_word"]:
                    ws_publish.Cells(row, publish_cols[HEADER_WORD_PUBLISH]).Interior.Color = NEW_WORD_HIGHLIGHT_COLOR

        # Chỉ ghi/cập nhật note ở DANG_BAI. Không sửa nội dung ở VIET_BAI hoặc KE_HOACH.
        note_col = publish_cols[HEADER_DUPLICATE_NOTE_PUBLISH]
        all_publish_keys = set(publish_rows) | set(added_by_key)
        marked_duplicate_rows = 0
        for key in all_publish_keys:
            final_items = [*publish_rows.get(key, []), *added_by_key.get(key, [])]
            is_duplicate = len(final_items) >= 2
            for item in final_items:
                if is_duplicate:
                    ws_publish.Cells(item["row"], note_col).Value2 = DUPLICATE_NOTE
                    marked_duplicate_rows += 1
                elif normalize(item.get("duplicate_note")) == normalize(DUPLICATE_NOTE):
                    ws_publish.Cells(item["row"], note_col).ClearContents()

        workbook.Save()
    finally:
        excel.ScreenUpdating = old_screen_updating
        excel.EnableEvents = old_enable_events

    print(f"Đã thêm mới vào DANG_BAI: {len(added_rows)} dòng.")
    print(f"Đã ghi '{DUPLICATE_NOTE}' tại DANG_BAI: {marked_duplicate_rows} dòng.")
    print(f"Đã chọn Word đợt 2: {used_new_word} dòng.")


if __name__ == "__main__":
    main()
