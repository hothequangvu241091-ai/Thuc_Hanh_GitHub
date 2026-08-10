# -*- coding: utf-8 -*-
"""V3.19: VIET_BAI có combo4 nhưng KE_HOACH không có thì báo khả năng đã bị xóa."""

from __future__ import annotations

import re
from datetime import date
from typing import Any
from urllib.parse import urlparse

import win32com.client as win32


# Khi chạy từ app, flow_host gán trực tiếp workbook Excel ẩn vào biến này.
# Chạy độc lập vẫn dùng workbook đang mở như trước.
APP_WORKBOOK: Any = None


SHEET_WRITE = "VIET_BAI"
SHEET_PLAN = "KE_HOACH"
SHEET_GPT = "CAU_HINH_GPT"

WRAP_TEXT = False
MAX_SAFE_WRITE_ROW = 20000

HEADER_SEO_WRITE = "Tiêu đề SEO"
HEADER_H1_WRITE = "H1"
HEADER_FOLDER_WRITE = "Tên Miền"
HEADER_KEYWORD_WRITE = "Từ khóa"
HEADER_PROMPT_WRITE = "Prompt viết bài"
HEADER_GPT_TYPE_WRITE = "Loại GPT"
HEADER_GPT_URL_WRITE = "URL GPT gốc"
HEADER_CHECK_STATUS_WRITE = "Trạng thái kiểm tra"
HEADER_COMPLETE_STATUS_WRITE = "Trạng thái hoàn tất"
HEADER_TRANSFER_NOTE_WRITE = "Ghi chú chuyển dữ liệu"

HEADER_SEO_PLAN = "Title [SEO]"
HEADER_H1_PLAN = "Article Name [H1]"
HEADER_FOLDER_PLAN = "Tên Miền"
HEADER_KEYWORD_PLAN = "Main Keyword"
HEADER_PROMPT_PLAN = """ĐẦU VÀO CỦA PROMPT
[Copy > Dán đúng loai Prompt]"""
HEADER_PROMPT_PLAN_FALLBACK = """ĐẦU VÀO CỦA PROMPT
[Copy > Dán đúng Prompt]"""
HEADER_GPT_TYPE_PLAN = """CHÚ Ý ĐẶC BIỆT (!!!)
[PROMPT SỬ DỤNG]"""
HEADER_URL_PAGE_PLAN = "URL Page"

DELETED_MARKER = "Bài đã xóa"
DELETE_MANUAL_MARKER = "Bài đã xóa - cần xóa thủ công"
DUPLICATE_ARTICLE_MARKER = "Bài viết trùng"
HAS_URL_MARKER = "Đã đăng"
WRITTEN_MARKER = "Đã viết"
NOT_IN_PLAN_MARKER = "Khả năng đã bị xóa"
NOT_IN_PLAN_UNFINISHED_MARKER = "Không còn trong kế hoạch - chưa hoàn tất"
NEWLY_WRITTEN_MARKER = "Mới viết xong"
DATA_PROBLEM_MARKER = "Dữ liệu này có vấn đề"


def normalize(value: Any) -> str:
    """Chuẩn hóa nội dung dùng để tìm tiêu đề cột và đối chiếu khóa."""
    text = re.sub(r"\s+", " ", str(value or "").strip())
    return text.casefold()


def is_usable_key(key: str) -> bool:
    return bool(key) and key not in {
        "#",
        "-",
        "#.n/a",
        "#n/a",
        "n/a",
        "na",
    }


def is_valid_url(value: Any) -> bool:
    return bool(
        re.fullmatch(r"https?://\S+", str(value or "").strip(), re.IGNORECASE)
    )


def url_matches_domain(value: Any, domain: Any) -> bool:
    """Chỉ nhận URL http(s) có hostname khớp đúng tên miền kế hoạch."""
    url_text = str(value or "").strip()
    if not is_valid_url(url_text):
        return False
    host = (urlparse(url_text).hostname or "").casefold()
    wanted = str(domain or "").strip().casefold()
    if host.startswith("www."):
        host = host[4:]
    if wanted.startswith("www."):
        wanted = wanted[4:]
    return bool(host and wanted and host == wanted)


def headers(ws: Any) -> dict[str, int]:
    """Trả về vị trí cột theo tên tiêu đề ở hàng 1."""
    last_col = int(ws.Cells(1, ws.Columns.Count).End(-4159).Column)
    result: dict[str, int] = {}
    for col in range(1, last_col + 1):
        value = ws.Cells(1, col).Value
        if value is not None and str(value).strip():
            result[normalize(value)] = col
    return result


def require_columns(ws: Any, names: list[str]) -> dict[str, int]:
    """Kiểm tra sheet có đủ các tiêu đề bắt buộc."""
    available = headers(ws)
    missing = [name for name in names if normalize(name) not in available]
    if missing:
        raise RuntimeError(
            f"Sheet {ws.Name} thiếu tiêu đề: " + ", ".join(missing)
        )
    return {name: available[normalize(name)] for name in names}


def require_preferred_column(
    ws: Any,
    preferred_name: str,
    fallback_name: str,
) -> int:
    """
    Tìm cột ưu tiên; chỉ dùng cột thay thế khi không có cột ưu tiên.

    Nếu cả hai cùng xuất hiện, cột ưu tiên được chọn.
    """
    available = headers(ws)
    preferred_key = normalize(preferred_name)
    fallback_key = normalize(fallback_name)

    if preferred_key in available:
        return available[preferred_key]
    if fallback_key in available:
        print(
            f"Sheet {ws.Name}: không có '{preferred_name}', "
            f"đang dùng cột thay thế '{fallback_name}'."
        )
        return available[fallback_key]

    raise RuntimeError(
        f"Sheet {ws.Name} thiếu tiêu đề: "
        f"{preferred_name} (hoặc {fallback_name})"
    )


def require_any_column(ws: Any, names: list[str]) -> int:
    """
    Tìm một trong các tên cột tương đương.

    Nếu nhiều tên cùng xuất hiện, chọn cột nằm trước từ trái sang phải.
    """
    available = headers(ws)
    matches = [
        (available[normalize(name)], name)
        for name in names
        if normalize(name) in available
    ]
    if not matches:
        raise RuntimeError(
            f"Sheet {ws.Name} thiếu tiêu đề: " + " hoặc ".join(names)
        )

    matches.sort(key=lambda item: item[0])
    selected_col, selected_name = matches[0]
    if len(matches) > 1:
        print(
            f"Sheet {ws.Name}: có nhiều heading tương đương cho Prompt; "
            f"đang dùng cột '{selected_name}' (cột {selected_col})."
        )
    return selected_col


def ensure_column(ws: Any, header: str) -> int:
    """Lấy cột có sẵn hoặc tự thêm heading mới ở cuối sheet."""
    available = headers(ws)
    header_key = normalize(header)
    if header_key in available:
        return available[header_key]
    last_col = int(ws.Cells(1, ws.Columns.Count).End(-4159).Column)
    next_col = last_col + 1 if last_col >= 1 else 1
    ws.Cells(1, next_col).Value2 = header
    print(f"Sheet {ws.Name}: đã thêm cột '{header}'.")
    return next_col


def last_data_row(ws: Any, col: int) -> int:
    """
    Tìm dòng cuối theo giá trị hiển thị thật.

    Không dùng End(xlUp), vì ô có công thức trả về chuỗi rỗng vẫn có thể làm
    Excel nhận nhầm dòng cuối ở rất xa.
    """
    found = ws.Columns(col).Find(
        What="*",
        After=ws.Cells(1, col),
        LookIn=-4163,       # xlValues
        LookAt=2,           # xlPart
        SearchOrder=1,      # xlByRows
        SearchDirection=2,  # xlPrevious
        MatchCase=False,
    )
    if found is None:
        return 1
    return int(found.Row)


def assert_write_range_is_safe(ws: Any, last_row: int) -> None:
    """Dừng trước khi ghi/định dạng nếu vùng dữ liệu hoặc Excel Table bị phình."""
    if last_row > MAX_SAFE_WRITE_ROW:
        raise RuntimeError(
            f"Sheet {ws.Name} đang nhận dòng dữ liệu cuối là {last_row:,}, "
            f"vượt ngưỡng an toàn {MAX_SAFE_WRITE_ROW:,}. "
            "Dừng để tránh ghi và Wrap Text hàng nghìn dòng trống."
        )

    list_objects = ws.ListObjects
    for index in range(1, int(list_objects.Count) + 1):
        table = list_objects.Item(index)
        table_last_row = int(
            table.Range.Row + table.Range.Rows.Count - 1
        )
        if table_last_row > MAX_SAFE_WRITE_ROW:
            raise RuntimeError(
                f"Excel Table '{table.Name}' trên sheet {ws.Name} đang kéo "
                f"tới dòng {table_last_row:,}. Hãy dùng Table Design > "
                f"Resize Table về dòng dữ liệu cuối thật ({last_row:,}) "
                "rồi chạy lại. Chương trình chưa ghi hay định dạng dữ liệu."
            )


def read_matrix(
    ws: Any,
    first_row: int,
    last_row: int,
    first_col: int,
    last_col: int,
) -> tuple:
    """Đọc vùng Excel bằng một lệnh COM và luôn trả về ma trận 2 chiều."""
    values = ws.Range(
        ws.Cells(first_row, first_col),
        ws.Cells(last_row, last_col),
    ).Value2
    if not isinstance(values, tuple):
        return ((values,),)
    if values and not isinstance(values[0], tuple):
        return (values,)
    return values


def plan_record_priority(url_page: Any) -> str:
    """Các trạng thái KE_HOACH có quyền chặn việc tạo/cập nhật bài viết."""
    if normalize(url_page) == normalize(DELETED_MARKER):
        return DELETED_MARKER
    if is_valid_url(url_page):
        return HAS_URL_MARKER
    # "Đã viết" là trạng thái của bài trong kế hoạch, không phải URL lỗi.
    # Phải nhận diện riêng để không bị coi là nhiệm vụ chưa xử lý và thêm lại
    # vào VIET_BAI ở mỗi lần chạy.
    if normalize(url_page) == normalize(WRITTEN_MARKER):
        return WRITTEN_MARKER
    if normalize(url_page) == normalize(DUPLICATE_ARTICLE_MARKER):
        return DUPLICATE_ARTICLE_MARKER
    return ""


def build_plan_records(
    ws: Any,
    title_col: int,
    h1_col: int,
    url_page_col: int,
    folder_col: int,
    keyword_col: int,
    prompt_col: int,
    gpt_type_col: int,
) -> tuple[list[dict[str, Any]], list[str]]:
    """Đọc KE_HOACH, nhận dạng mỗi bài theo Title + Tên Miền + Keyword."""
    last_row = last_data_row(ws, title_col)
    if last_row < 2:
        return [], []

    first_col = min(
        title_col,
        h1_col,
        url_page_col,
        folder_col,
        keyword_col,
        prompt_col,
        gpt_type_col,
    )
    last_col = max(
        title_col,
        h1_col,
        url_page_col,
        folder_col,
        keyword_col,
        prompt_col,
        gpt_type_col,
    )
    matrix = read_matrix(ws, 2, last_row, first_col, last_col)
    column_indexes = {
        "title": title_col - first_col,
        "h1": h1_col - first_col,
        "url_page": url_page_col - first_col,
        "folder": folder_col - first_col,
        "keyword": keyword_col - first_col,
        "prompt": prompt_col - first_col,
        "gpt_type": gpt_type_col - first_col,
    }
    records: list[dict[str, Any]] = []
    missing_identity: list[str] = []

    for row_number, values in enumerate(matrix, start=2):
        title = values[column_indexes["title"]]
        title_key = normalize(title)
        if not is_usable_key(title_key):
            continue
        folder = values[column_indexes["folder"]]
        folder_key = normalize(folder)
        if not is_usable_key(folder_key):
            missing_identity.append(f"Dòng {row_number}: {str(title).strip()}")
            continue
        keyword = values[column_indexes["keyword"]]
        keyword_key = normalize(keyword)
        if not is_usable_key(keyword_key):
            missing_identity.append(f"Dòng {row_number}: {str(title).strip()}")
            continue
        h1 = values[column_indexes["h1"]]
        h1_key = normalize(h1)
        if not is_usable_key(h1_key):
            missing_identity.append(f"Dòng {row_number}: {str(title).strip()}")
            continue
        key = (title_key, folder_key, keyword_key, h1_key)
        url_page = values[column_indexes["url_page"]]
        records.append(
            {
                "row": row_number,
                "key": key,
                "title": title,
                "h1": h1,
                "url_page": url_page,
                "folder": folder,
                "keyword": keyword,
                "prompt": values[column_indexes["prompt"]],
                "gpt_type": values[column_indexes["gpt_type"]],
                "priority": plan_record_priority(url_page),
            }
        )
    return records, missing_identity


def build_write_rows(
    ws: Any,
    title_col: int,
    h1_col: int,
    folder_col: int,
    keyword_col: int,
    check_status_col: int,
    complete_status_col: int,
) -> tuple[dict[tuple[str, str, str, str], list[dict[str, Any]]], list[dict[str, Any]]]:
    """Đọc VIET_BAI để đối chiếu theo Tiêu đề SEO + Thư mục + Từ khóa."""
    last_row = last_data_row(ws, title_col)
    if last_row < 2:
        return {}, []

    first_col = min(
        title_col, h1_col, folder_col, keyword_col, check_status_col, complete_status_col
    )
    last_col = max(
        title_col, h1_col, folder_col, keyword_col, check_status_col, complete_status_col
    )
    matrix = read_matrix(ws, 2, last_row, first_col, last_col)
    title_index = title_col - first_col
    h1_index = h1_col - first_col
    folder_index = folder_col - first_col
    keyword_index = keyword_col - first_col
    check_index = check_status_col - first_col
    complete_index = complete_status_col - first_col
    rows_by_key: dict[tuple[str, str, str, str], list[dict[str, Any]]] = {}
    all_rows: list[dict[str, Any]] = []

    for row_number, values in enumerate(matrix, start=2):
        title = values[title_index]
        title_key = normalize(title)
        if not is_usable_key(title_key):
            continue
        folder = values[folder_index]
        folder_key = normalize(folder)
        h1 = values[h1_index]
        h1_key = normalize(h1)
        keyword = values[keyword_index]
        keyword_key = normalize(keyword)
        item = {
            "row": row_number,
            "title_key": title_key,
            "folder_key": folder_key,
            "h1_key": h1_key,
            "keyword_key": keyword_key,
            "check_status": values[check_index],
            "complete_status": values[complete_index],
        }
        all_rows.append(item)
        if not is_usable_key(folder_key) or not is_usable_key(keyword_key) or not is_usable_key(h1_key):
            continue
        key = (title_key, folder_key, keyword_key, h1_key)
        rows_by_key.setdefault(key, []).append(item)
    return rows_by_key, all_rows


def build_gpt_lookup(ws: Any) -> dict[str, Any]:
    """Đọc CAU_HINH_GPT: cột A là Loại GPT, cột B là URL GPT gốc."""
    lookup: dict[str, Any] = {}
    duplicates: list[str] = []
    last_row = last_data_row(ws, 1)
    if last_row < 1:
        return lookup

    matrix = read_matrix(ws, 1, last_row, 1, 2)
    for gpt_type, gpt_url in matrix:
        key = normalize(gpt_type)
        if not is_usable_key(key):
            continue
        if key in lookup:
            duplicates.append(str(gpt_type).strip())
            continue
        lookup[key] = gpt_url

    if duplicates:
        preview = ", ".join(duplicates[:10])
        raise RuntimeError(
            f"{SHEET_GPT} có Loại GPT bị trùng: {preview}"
        )
    return lookup


def main() -> None:
    print("=" * 70)
    print("CHUẨN BỊ VIET_BAI V3.19 - BÁO COMBO VIET_BAI KHÔNG CÒN TRONG KE_HOACH")
    print("=" * 70)

    workbook = APP_WORKBOOK
    if workbook is None:
        excel = win32.GetActiveObject("Excel.Application")
        workbook = excel.ActiveWorkbook
    else:
        excel = workbook.Application
    if workbook is None:
        raise RuntimeError("Không tìm thấy workbook Excel đang mở.")

    ws_write = workbook.Worksheets(SHEET_WRITE)
    ws_plan = workbook.Worksheets(SHEET_PLAN)
    ws_gpt = workbook.Worksheets(SHEET_GPT)

    write_cols = require_columns(
        ws_write,
        [
            HEADER_SEO_WRITE,
            HEADER_H1_WRITE,
            HEADER_FOLDER_WRITE,
            HEADER_KEYWORD_WRITE,
            HEADER_PROMPT_WRITE,
            HEADER_GPT_TYPE_WRITE,
            HEADER_GPT_URL_WRITE,
        ],
    )
    write_cols[HEADER_CHECK_STATUS_WRITE] = ensure_column(
        ws_write,
        HEADER_CHECK_STATUS_WRITE,
    )
    write_cols[HEADER_COMPLETE_STATUS_WRITE] = ensure_column(
        ws_write,
        HEADER_COMPLETE_STATUS_WRITE,
    )
    write_cols[HEADER_TRANSFER_NOTE_WRITE] = ensure_column(
        ws_write,
        HEADER_TRANSFER_NOTE_WRITE,
    )
    plan_cols = require_columns(
        ws_plan,
        [
            HEADER_SEO_PLAN,
            HEADER_H1_PLAN,
            HEADER_KEYWORD_PLAN,
            HEADER_GPT_TYPE_PLAN,
            HEADER_URL_PAGE_PLAN,
        ],
    )
    plan_available = headers(ws_plan)
    folder_plan_col = plan_available.get(normalize(HEADER_FOLDER_PLAN))
    if folder_plan_col is None:
        raise RuntimeError(
            f"Sheet {ws_plan.Name} thiếu tiêu đề: {HEADER_FOLDER_PLAN}"
        )
    plan_cols[HEADER_PROMPT_PLAN] = require_any_column(
        ws_plan,
        [
            HEADER_PROMPT_PLAN,
            HEADER_PROMPT_PLAN_FALLBACK,
        ],
    )

    plan_records, missing_identity = build_plan_records(
        ws_plan,
        plan_cols[HEADER_SEO_PLAN],
        plan_cols[HEADER_H1_PLAN],
        plan_cols[HEADER_URL_PAGE_PLAN],
        folder_plan_col,
        plan_cols[HEADER_KEYWORD_PLAN],
        plan_cols[HEADER_PROMPT_PLAN],
        plan_cols[HEADER_GPT_TYPE_PLAN],
    )
    gpt_lookup = build_gpt_lookup(ws_gpt)
    if not gpt_lookup:
        raise RuntimeError(
            f"Sheet {SHEET_GPT} không có cấu hình Loại GPT và URL."
        )

    seo_col = write_cols[HEADER_SEO_WRITE]
    last_row = last_data_row(ws_write, seo_col)
    assert_write_range_is_safe(ws_write, last_row)
    write_rows_by_key, all_write_rows = build_write_rows(
        ws_write,
        seo_col,
        write_cols[HEADER_H1_WRITE],
        write_cols[HEADER_FOLDER_WRITE],
        write_cols[HEADER_KEYWORD_WRITE],
        write_cols[HEADER_CHECK_STATUS_WRITE],
        write_cols[HEADER_COMPLETE_STATUS_WRITE],
    )

    updates_by_col: dict[int, dict[int, Any]] = {}
    updated = 0
    added = 0
    missing_gpt: list[str] = []
    next_row = max(last_row + 1, 2)
    plan_duplicate_keys: set[tuple[str, str, str, str]] = set()
    write_duplicate_keys: set[tuple[str, str, str, str]] = set()
    incomplete_write_rows: list[int] = []
    status_counts: dict[str, int] = {}

    def set_value(column: int, row: int, value: Any) -> None:
        updates_by_col.setdefault(column, {})[row] = value

    def write_plan_data(
        row: int,
        record: dict[str, Any],
        is_new: bool = False,
        new_status: str = "",
        complete_status: str = "",
        transfer_note: str = "",
    ) -> None:
        nonlocal updated
        gpt_type = record["gpt_type"]
        gpt_url = gpt_lookup.get(normalize(gpt_type))
        if gpt_url is None:
            missing_gpt.append(
                f"Dòng KE_HOACH {record['row']}: {str(record['title']).strip()} "
                f"(Loại GPT: {str(gpt_type or '').strip() or '[trống]'})"
            )
            gpt_url = ""
        values = {
            HEADER_SEO_WRITE: record["title"],
            HEADER_H1_WRITE: record["h1"],
            HEADER_FOLDER_WRITE: record["folder"],
            HEADER_KEYWORD_WRITE: record["keyword"],
            HEADER_PROMPT_WRITE: record["prompt"],
            HEADER_GPT_TYPE_WRITE: gpt_type,
            HEADER_GPT_URL_WRITE: gpt_url,
            HEADER_CHECK_STATUS_WRITE: (
                new_status
                if is_new
                else ""
            ),
        }
        if is_new and complete_status:
            values[HEADER_COMPLETE_STATUS_WRITE] = complete_status
        if is_new and transfer_note:
            values[HEADER_TRANSFER_NOTE_WRITE] = transfer_note
        for header, value in values.items():
            set_value(write_cols[header], row, "" if value is None else value)
        updated += 1

    # V3.15: lập bản đồ combo4 duy nhất của KE_HOACH. Combo trùng bị khóa,
    # không tự chọn dòng và không ghi hàng loạt.
    plan_rows_by_key: dict[
        tuple[str, str, str, str], list[dict[str, Any]]
    ] = {}
    for record in plan_records:
        plan_rows_by_key.setdefault(record["key"], []).append(record)
    for key, records in plan_rows_by_key.items():
        if len(records) > 1:
            plan_duplicate_keys.add(key)

    for key, rows in write_rows_by_key.items():
        if len(rows) > 1:
            write_duplicate_keys.add(key)

    def queue_status(row: int, status: str) -> None:
        set_value(write_cols[HEADER_CHECK_STATUS_WRITE], row, status)
        status_counts[status] = status_counts.get(status, 0) + 1

    # LƯỢT 1: VIET_BAI -> KE_HOACH. Đây là hướng xử lý chính.
    for key, rows in write_rows_by_key.items():
        if key in write_duplicate_keys or key in plan_duplicate_keys:
            continue
        item = rows[0]
        plan_matches = plan_rows_by_key.get(key, [])
        complete_ok = normalize(item["complete_status"]) == normalize("OK")

        if not plan_matches:
            # V3.19: VIET_BAI vẫn còn combo4 nhưng KE_HOACH không còn combo đó.
            # Không tự xóa hay thêm ngược vào KE_HOACH; chỉ cảnh báo để kiểm tra.
            queue_status(item["row"], NOT_IN_PLAN_MARKER)
            continue

        record = plan_matches[0]
        url_page = str(record["url_page"] or "").strip()
        normalized_url_page = normalize(url_page)
        if normalized_url_page == normalize(DELETED_MARKER):
            check_status = DELETE_MANUAL_MARKER
        elif normalized_url_page == normalize(WRITTEN_MARKER):
            check_status = WRITTEN_MARKER
        elif url_matches_domain(url_page, record["folder"]):
            check_status = HAS_URL_MARKER
        elif url_page == "":
            check_status = NEWLY_WRITTEN_MARKER if complete_ok else ""
        else:
            check_status = DATA_PROBLEM_MARKER
        queue_status(item["row"], check_status)

    # Dòng VIET_BAI thiếu combo4 chắc chắn là dữ liệu có vấn đề.
    for item in all_write_rows:
        if not (
            is_usable_key(item["folder_key"])
            and is_usable_key(item["keyword_key"])
            and is_usable_key(item["h1_key"])
        ):
            incomplete_write_rows.append(item["row"])
            queue_status(item["row"], DATA_PROBLEM_MARKER)

    # LƯỢT 2: KE_HOACH -> VIET_BAI.
    # Chỉ thêm combo4 còn thiếu khi URL Page trống hoặc là URL đúng tên miền.
    existing_write_keys = set(write_rows_by_key)
    for key, records in plan_rows_by_key.items():
        if key in plan_duplicate_keys or key in existing_write_keys:
            continue
        record = records[0]
        url_page = str(record["url_page"] or "").strip()
        complete_status = ""
        transfer_note = ""
        if url_page == "":
            new_status = f"Mới thêm vào {date.today().isoformat()}"
        elif url_matches_domain(url_page, record["folder"]):
            new_status = HAS_URL_MARKER
            complete_status = "OK"
            transfer_note = "Mới thêm vào đã đăng"
        else:
            continue
        target_row = next_row
        next_row += 1
        added += 1
        write_plan_data(
            target_row,
            record,
            is_new=True,
            new_status=new_status,
            complete_status=complete_status,
            transfer_note=transfer_note,
        )
        existing_write_keys.add(key)

    final_last_row = next_row - 1
    assert_write_range_is_safe(ws_write, final_last_row)
    old_screen_updating = excel.ScreenUpdating
    old_enable_events = excel.EnableEvents

    try:
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        for col, row_values in updates_by_col.items():
            if final_last_row < 2:
                continue
            values = [
                list(item)
                for item in read_matrix(ws_write, 2, final_last_row, col, col)
            ]
            for row, value in row_values.items():
                values[row - 2][0] = value
            ws_write.Range(
                ws_write.Cells(2, col),
                ws_write.Cells(final_last_row, col),
            ).Value2 = tuple(tuple(item) for item in values)

        # Hiển thị nội dung trên một dòng như bảng dữ liệu thông thường.
        # Gán định dạng cho cả vùng bằng một lệnh COM; không duyệt từng dòng.
        first_display_col = min(
            write_cols[HEADER_SEO_WRITE],
            *(
                write_cols[header]
                for header in (
                    HEADER_FOLDER_WRITE,
                    HEADER_KEYWORD_WRITE,
                    HEADER_PROMPT_WRITE,
                    HEADER_GPT_TYPE_WRITE,
                    HEADER_GPT_URL_WRITE,
                    HEADER_CHECK_STATUS_WRITE,
                )
            ),
        )
        last_display_col = max(
            write_cols[HEADER_SEO_WRITE],
            *(
                write_cols[header]
                for header in (
                    HEADER_FOLDER_WRITE,
                    HEADER_KEYWORD_WRITE,
                    HEADER_PROMPT_WRITE,
                    HEADER_GPT_TYPE_WRITE,
                    HEADER_GPT_URL_WRITE,
                    HEADER_CHECK_STATUS_WRITE,
                )
            ),
        )
        if final_last_row >= 2:
            display_range = ws_write.Range(
                ws_write.Cells(2, first_display_col),
                ws_write.Cells(final_last_row, last_display_col),
            )
            display_range.WrapText = WRAP_TEXT

        workbook.Save()
    finally:
        excel.ScreenUpdating = old_screen_updating
        excel.EnableEvents = old_enable_events

    print(f"Đã thêm dòng mới: {added}.")
    print(f"Đã nhập dữ liệu cho dòng mới: {updated} dòng.")
    for status, count in sorted(status_counts.items()):
        label = status or "[TRỐNG]"
        print(f"Đã ghi Trạng thái kiểm tra '{label}': {count} dòng.")
    if missing_identity:
        print(
            "KE_HOACH thiếu Title [SEO], Tên Miền hoặc Main Keyword: "
            f"{len(missing_identity)} dòng."
        )
        for item in missing_identity[:30]:
            print("- " + item)
    if missing_gpt:
        print(
            f"Không tìm thấy URL trong {SHEET_GPT}: "
            f"{len(missing_gpt)} dòng."
        )
        for item in missing_gpt[:30]:
            print("- " + item)
    if plan_duplicate_keys:
        print(
            "KE_HOACH có combo4 bị lặp; không ghi và không thêm: "
            f"{len(plan_duplicate_keys)} combo."
        )
    if write_duplicate_keys:
        print(
            "VIET_BAI có combo4 bị lặp; không ghi trạng thái: "
            f"{len(write_duplicate_keys)} combo."
        )
    if incomplete_write_rows:
        print(
            "VIET_BAI thiếu Thư mục hoặc Từ khóa nên không thể đối chiếu: "
            f"{len(incomplete_write_rows)} dòng."
        )
    print("Các cột khác trong VIET_BAI được giữ nguyên.")


if __name__ == "__main__":
    main()
