# -*- coding: utf-8 -*-
"""Điền VIET_BAI; chỉ báo trùng khi cùng Title [SEO] và File nguồn."""

from __future__ import annotations

import re
from typing import Any

import win32com.client as win32


SHEET_WRITE = "VIET_BAI"
SHEET_PLAN = "KE_HOACH"
SHEET_GPT = "CAU_HINH_GPT"

WRAP_TEXT = False
MAX_SAFE_WRITE_ROW = 5000

HEADER_SEO_WRITE = "Tiêu đề SEO"
HEADER_FOLDER_WRITE = "Thư mục"
HEADER_KEYWORD_WRITE = "Từ khóa"
HEADER_PROMPT_WRITE = "Prompt viết bài"
HEADER_GPT_TYPE_WRITE = "Loại GPT"
HEADER_GPT_URL_WRITE = "URL GPT gốc"
HEADER_COMPLETED_STATUS_WRITE = "Trạng thái hoàn tất"

HEADER_SEO_PLAN = "Title [SEO]"
HEADER_FOLDER_PLAN = "Tên Miền"
HEADER_KEYWORD_PLAN = "Main Keyword"
HEADER_PROMPT_PLAN = """ĐẦU VÀO CỦA PROMPT
[Copy > Dán đúng loai Prompt]"""
HEADER_PROMPT_PLAN_FALLBACK = """ĐẦU VÀO CỦA PROMPT
[Copy > Dán đúng Prompt]"""
HEADER_GPT_TYPE_PLAN = """CHÚ Ý ĐẶC BIỆT (!!!)
[PROMPT SỬ DỤNG]"""
HEADER_SOURCE_FILE_PLAN = "File nguồn"


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


def extract_domain_from_source_file(value: Any) -> str:
    """
    Tách tên miền từ cột File nguồn.

    Ví dụ:
    - 1.old\\baothuonggia.com_Content-Web.xlsx -> baothuonggia.com
    - baothuonggia.com.xlsx -> baothuonggia.com
    - baothuonggia.com -> baothuonggia.com
    """
    text = str(value or "").strip()
    if not text:
        return ""

    # Chỉ bỏ phần mở rộng file Excel; không dùng Path.stem vì chuỗi chỉ có
    # "domain.com" sẽ bị hiểu nhầm .com là phần mở rộng.
    text = re.sub(r"(?i)\.(?:xlsx|xlsm|xls)$", "", text)
    matches = re.findall(
        r"(?i)(?<![a-z0-9-])"
        r"(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+"
        r"[a-z]{2,63}",
        text,
    )
    if not matches:
        return ""

    # File nguồn có thể nằm dưới thư mục như "1.old"; tên miền thật thường
    # là kết quả cuối cùng, nằm trong chính tên file.
    domain = matches[-1].lower().strip(".")
    if domain.startswith("www."):
        domain = domain[4:]
    return domain


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


def build_unique_lookup(
    ws: Any,
    key_col: int,
    source_file_col: int,
    value_cols: list[int],
) -> dict[str, tuple[Any, ...]]:
    """Tra theo Title; chỉ dừng nếu cùng Title + File nguồn bị lặp."""
    lookup: dict[str, tuple[Any, ...]] = {}
    duplicates: list[str] = []
    seen_composite: dict[tuple[str, str], int] = {}
    last_row = last_data_row(ws, key_col)
    if last_row < 2:
        return lookup

    first_col = min([key_col, source_file_col, *value_cols])
    last_col = max([key_col, source_file_col, *value_cols])
    matrix = read_matrix(ws, 2, last_row, first_col, last_col)
    key_index = key_col - first_col
    source_file_index = source_file_col - first_col
    value_indexes = [col - first_col for col in value_cols]

    for row, values in enumerate(matrix, start=2):
        raw_key = values[key_index]
        key = normalize(raw_key)
        if not is_usable_key(key):
            continue

        raw_source_file = values[source_file_index]
        source_file_key = normalize(raw_source_file)
        composite_key = (key, source_file_key)
        if composite_key in seen_composite:
            duplicates.append(
                f"dòng {seen_composite[composite_key]} và {row}: "
                f"{str(raw_key).strip()} | File nguồn: "
                f"{str(raw_source_file or '').strip() or '[trống]'}"
            )
            continue
        seen_composite[composite_key] = row

        # VIET_BAI hiện không có File nguồn để tạo khóa kép lúc tra cứu.
        # Nếu Title có ở file khác, giữ bản xuất hiện đầu tiên trong KE_HOACH.
        if key not in lookup:
            lookup[key] = tuple(values[index] for index in value_indexes)

    if duplicates:
        preview = "\n- " + "\n- ".join(duplicates[:10])
        raise RuntimeError(
            f"{SHEET_PLAN} có Title [SEO] + File nguồn bị trùng:{preview}"
        )
    return lookup


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
    print("CHUẨN BỊ DỮ LIỆU VIẾT BÀI THEO TIÊU ĐỀ SEO")
    print("=" * 70)

    excel = win32.GetActiveObject("Excel.Application")
    workbook = excel.ActiveWorkbook
    if workbook is None:
        raise RuntimeError("Không tìm thấy workbook Excel đang mở.")

    ws_write = workbook.Worksheets(SHEET_WRITE)
    ws_plan = workbook.Worksheets(SHEET_PLAN)
    ws_gpt = workbook.Worksheets(SHEET_GPT)

    write_cols = require_columns(
        ws_write,
        [
            HEADER_SEO_WRITE,
            HEADER_FOLDER_WRITE,
            HEADER_KEYWORD_WRITE,
            HEADER_PROMPT_WRITE,
            HEADER_GPT_TYPE_WRITE,
            HEADER_GPT_URL_WRITE,
            HEADER_COMPLETED_STATUS_WRITE,
        ],
    )
    plan_cols = require_columns(
        ws_plan,
        [
            HEADER_SEO_PLAN,
            HEADER_KEYWORD_PLAN,
            HEADER_GPT_TYPE_PLAN,
        ],
    )
    plan_available = headers(ws_plan)
    folder_plan_col = plan_available.get(normalize(HEADER_FOLDER_PLAN))
    if folder_plan_col is None:
        raise RuntimeError(
            f"Sheet {ws_plan.Name} thiếu tiêu đề: {HEADER_FOLDER_PLAN}"
        )
    source_file_plan_col = plan_available.get(
        normalize(HEADER_SOURCE_FILE_PLAN)
    )
    if source_file_plan_col is None:
        raise RuntimeError(
            f"Sheet {ws_plan.Name} thiếu tiêu đề: "
            f"{HEADER_SOURCE_FILE_PLAN}"
        )
    plan_cols[HEADER_PROMPT_PLAN] = require_any_column(
        ws_plan,
        [
            HEADER_PROMPT_PLAN,
            HEADER_PROMPT_PLAN_FALLBACK,
        ],
    )

    plan_lookup = build_unique_lookup(
        ws_plan,
        plan_cols[HEADER_SEO_PLAN],
        source_file_plan_col,
        [
            folder_plan_col,
            plan_cols[HEADER_KEYWORD_PLAN],
            plan_cols[HEADER_PROMPT_PLAN],
            plan_cols[HEADER_GPT_TYPE_PLAN],
        ],
    )
    gpt_lookup = build_gpt_lookup(ws_gpt)
    if not gpt_lookup:
        raise RuntimeError(
            f"Sheet {SHEET_GPT} không có cấu hình Loại GPT và URL."
        )

    seo_col = write_cols[HEADER_SEO_WRITE]
    last_row = last_data_row(ws_write, seo_col)
    assert_write_range_is_safe(ws_write, last_row)
    if last_row < 2:
        print("VIET_BAI chưa có Tiêu đề SEO để xử lý.")
        return

    updated = 0
    skipped_completed = 0
    missing_plan: list[str] = []
    missing_gpt: list[str] = []
    old_screen_updating = excel.ScreenUpdating
    old_enable_events = excel.EnableEvents

    try:
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        seo_values = read_matrix(ws_write, 2, last_row, seo_col, seo_col)
        completed_status_values = read_matrix(
            ws_write,
            2,
            last_row,
            write_cols[HEADER_COMPLETED_STATUS_WRITE],
            write_cols[HEADER_COMPLETED_STATUS_WRITE],
        )
        output_headers = [
            HEADER_FOLDER_WRITE,
            HEADER_KEYWORD_WRITE,
            HEADER_PROMPT_WRITE,
            HEADER_GPT_TYPE_WRITE,
            HEADER_GPT_URL_WRITE,
        ]
        output_values = {
            header: [
                list(item)
                for item in read_matrix(
                    ws_write,
                    2,
                    last_row,
                    write_cols[header],
                    write_cols[header],
                )
            ]
            for header in output_headers
        }

        for offset, seo_item in enumerate(seo_values):
            row = offset + 2
            raw_seo = seo_item[0]
            key = normalize(raw_seo)

            # Dòng đã hoàn tất được bảo vệ toàn bộ, kể cả khi KE_HOACH
            # không còn Tiêu đề SEO tương ứng.
            if normalize(completed_status_values[offset][0]) == "ok":
                skipped_completed += 1
                continue

            # Tiêu đề SEO trống: giữ nguyên dữ liệu hiện có.
            if not is_usable_key(key):
                continue

            plan_data = plan_lookup.get(key)
            if plan_data is None:
                # Không tìm thấy trong KE_HOACH: báo thiếu và giữ nguyên
                # toàn bộ dữ liệu hiện có của dòng VIET_BAI.
                missing_plan.append(f"Dòng {row}: {str(raw_seo).strip()}")
                continue

            folder, keyword, prompt, gpt_type = plan_data
            gpt_url = gpt_lookup.get(normalize(gpt_type))
            if gpt_url is None:
                missing_gpt.append(
                    f"Dòng {row}: {str(raw_seo).strip()} "
                    f"(Loại GPT: {str(gpt_type or '').strip() or '[trống]'})"
                )
                gpt_url = ""

            new_values = (folder, keyword, prompt, gpt_type, gpt_url)
            for header, value in zip(output_headers, new_values):
                output_values[header][offset][0] = value
            updated += 1

        for header in output_headers:
            ws_write.Range(
                ws_write.Cells(2, write_cols[header]),
                ws_write.Cells(last_row, write_cols[header]),
            ).Value2 = tuple(tuple(item) for item in output_values[header])

        # Hiển thị nội dung trên một dòng như bảng dữ liệu thông thường.
        # Gán định dạng cho cả vùng bằng một lệnh COM; không duyệt từng dòng.
        first_display_col = min(
            write_cols[HEADER_SEO_WRITE],
            *(write_cols[header] for header in output_headers),
        )
        last_display_col = max(
            write_cols[HEADER_SEO_WRITE],
            *(write_cols[header] for header in output_headers),
        )
        display_range = ws_write.Range(
            ws_write.Cells(2, first_display_col),
            ws_write.Cells(last_row, last_display_col),
        )
        display_range.WrapText = WRAP_TEXT

        workbook.Save()
    finally:
        excel.ScreenUpdating = old_screen_updating
        excel.EnableEvents = old_enable_events

    print(f"Đã cập nhật: {updated} dòng.")
    print(
        f"Đã bỏ qua và giữ nguyên: {skipped_completed} dòng "
        f"có '{HEADER_COMPLETED_STATUS_WRITE}' = OK."
    )
    if missing_plan:
        print(
            f"Không tìm thấy trong {SHEET_PLAN}: "
            f"{len(missing_plan)} dòng."
        )
        for item in missing_plan[:30]:
            print("- " + item)
    else:
        print("Tất cả Tiêu đề SEO đều đối chiếu thành công.")
    if missing_gpt:
        print(
            f"Không tìm thấy URL trong {SHEET_GPT}: "
            f"{len(missing_gpt)} dòng."
        )
        for item in missing_gpt[:30]:
            print("- " + item)
    print("Các cột khác trong VIET_BAI được giữ nguyên.")


if __name__ == "__main__":
    main()
