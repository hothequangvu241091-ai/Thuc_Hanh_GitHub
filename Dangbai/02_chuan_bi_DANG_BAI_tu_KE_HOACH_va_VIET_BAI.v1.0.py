# -*- coding: utf-8 -*-
"""Điền DANG_BAI theo khóa Tiêu đề SEO + File nguồn (v1.0)."""

from __future__ import annotations

import re
from typing import Any

import win32com.client as win32


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

HEADER_SEO_PLAN = "Title [SEO]"
HEADER_DOMAIN_PLAN = "Tên Miền"
HEADER_SOURCE_FILE_PLAN = "File nguồn"
HEADER_CATEGORY_PLAN = "POST / UPDATE"
HEADER_CATEGORY_PLAN_ALTERNATIVE = "CATE [POST]"
HEADER_H1_PLAN = "Article Name [H1]"

HEADER_SEO_WRITE = "Tiêu đề SEO"
HEADER_TITLE_WRITE = "Từ khóa"
HEADER_WORD_WRITE = "Đường dẫn Word"
HEADER_OLD_WORD_COUNT_WRITE = "Số từ Word"
HEADER_NEW_WORD_WRITE = "Đường dẫn bài viết mới"
HEADER_NEW_WORD_COUNT_WRITE = "Số từ bài viết mới"
HEADER_IMAGE1_WRITE = "Đường dẫn ảnh 1"
HEADER_IMAGE2_WRITE = "Đường dẫn ảnh 2"

MIN_EXTRA_WORDS = 50
MAX_SAFE_PUBLISH_ROW = 5000
# Excel/VBA RGB(255, 255, 153): vàng nhạt để đánh dấu ô đang dùng Word đợt 2.
NEW_WORD_HIGHLIGHT_COLOR = 10092543


def normalize(value: Any) -> str:
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


def to_number(value: Any) -> float | None:
    """Đổi dữ liệu số từ trong Excel thành số; dữ liệu trống/sai trả về None."""
    if value is None or str(value).strip() == "":
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def has_value(value: Any) -> bool:
    return value is not None and bool(str(value).strip())


def domain_from_source_file(value: Any) -> str:
    """
    Lấy domain từ cột File nguồn.

    Ví dụ:
    - dananghitech.com.vn.xlsx -> dananghitech.com.vn
    - G:\\ThuMuc\\dananghitech.com.vn.xlsm -> dananghitech.com.vn
    """
    if not has_value(value):
        return ""

    filename = (
        str(value)
        .strip()
        .strip('"')
        .replace("\\", "/")
        .rsplit("/", 1)[-1]
    )
    domain = re.sub(
        r"\.(?:xlsx|xlsm|xls|csv)$",
        "",
        filename,
        flags=re.IGNORECASE,
    )
    domain = domain.strip().strip(".").casefold()

    # Không nhận mã như [01] hoặc một tên file thông thường.
    if not re.fullmatch(
        r"(?=.{1,253}$)"
        r"(?:[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?\.)+"
        r"[a-z0-9](?:[a-z0-9-]{0,61}[a-z0-9])?",
        domain,
    ):
        return ""
    return domain


def source_key(value: Any) -> str:
    """Chuẩn hóa File nguồn/Tên miền thành cùng một khóa để đối chiếu."""
    return domain_from_source_file(value) or normalize(value)


def headers(ws: Any) -> dict[str, int]:
    last_col = int(ws.Cells(1, ws.Columns.Count).End(-4159).Column)
    result: dict[str, int] = {}
    for col in range(1, last_col + 1):
        value = ws.Cells(1, col).Value
        if value is not None and str(value).strip():
            result[normalize(value)] = col
    return result


def require_columns(ws: Any, names: list[str]) -> dict[str, int]:
    available = headers(ws)
    missing = [name for name in names if normalize(name) not in available]
    if missing:
        raise RuntimeError(
            f"Sheet {ws.Name} thiếu tiêu đề: " + ", ".join(missing)
        )
    return {name: available[normalize(name)] for name in names}


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
            f"Sheet {ws.Name}: có nhiều heading tương đương cho Danh mục; "
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


def assert_publish_range_is_safe(ws: Any, last_row: int) -> None:
    """Dừng trước khi ghi nếu dữ liệu hoặc Excel Table phình bất thường."""
    if last_row > MAX_SAFE_PUBLISH_ROW:
        raise RuntimeError(
            f"Sheet {ws.Name} đang nhận dòng dữ liệu cuối là {last_row:,}, "
            f"vượt ngưỡng an toàn {MAX_SAFE_PUBLISH_ROW:,}. "
            "Dừng để tránh ghi/tô màu hàng trăm nghìn dòng."
        )

    list_objects = ws.ListObjects
    for index in range(1, int(list_objects.Count) + 1):
        table = list_objects.Item(index)
        table_last_row = int(
            table.Range.Row + table.Range.Rows.Count - 1
        )
        if table_last_row > MAX_SAFE_PUBLISH_ROW:
            raise RuntimeError(
                f"Excel Table '{table.Name}' trên sheet {ws.Name} đang kéo "
                f"tới dòng {table_last_row:,}. Hãy dùng Table Design > "
                f"Resize Table về dòng dữ liệu cuối thật ({last_row:,}) "
                "rồi chạy lại. Chương trình chưa ghi hay tô màu dữ liệu."
            )


def read_matrix(ws: Any, first_row: int, last_row: int, first_col: int, last_col: int) -> tuple:
    """Đọc một vùng Excel bằng đúng một lệnh COM và luôn trả về ma trận 2 chiều."""
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
    value_cols: list[int],
    label: str,
) -> dict[str, tuple[Any, ...]]:
    lookup: dict[str, tuple[Any, ...]] = {}
    duplicates: list[str] = []
    last_row = last_data_row(ws, key_col)
    if last_row < 2:
        return lookup

    first_col = min([key_col, *value_cols])
    last_col = max([key_col, *value_cols])
    matrix = read_matrix(ws, 2, last_row, first_col, last_col)
    key_index = key_col - first_col
    value_indexes = [col - first_col for col in value_cols]

    for values in matrix:
        raw_key = values[key_index]
        key = normalize(raw_key)
        if not is_usable_key(key):
            continue
        if key in lookup:
            duplicates.append(str(raw_key).strip())
            continue
        lookup[key] = tuple(values[index] for index in value_indexes)
    if duplicates:
        preview = ", ".join(duplicates[:10])
        raise RuntimeError(f"{label} có Tiêu đề SEO bị trùng: {preview}")
    return lookup


def build_plan_lookup(
    ws: Any,
    seo_col: int,
    source_file_col: int,
    value_cols: list[int],
    label: str,
) -> tuple[
    dict[tuple[str, str], tuple[Any, ...]],
    dict[str, list[tuple[str, str]]],
]:
    """Tạo lookup KE_HOACH bằng khóa kép Title [SEO] + File nguồn."""
    lookup: dict[tuple[str, str], tuple[Any, ...]] = {}
    keys_by_title: dict[str, list[tuple[str, str]]] = {}
    duplicates: list[str] = []
    last_row = max(
        last_data_row(ws, seo_col),
        last_data_row(ws, source_file_col),
    )
    if last_row < 2:
        return lookup, keys_by_title

    first_col = min([seo_col, source_file_col, *value_cols])
    last_col = max([seo_col, source_file_col, *value_cols])
    matrix = read_matrix(ws, 2, last_row, first_col, last_col)
    seo_index = seo_col - first_col
    source_index = source_file_col - first_col
    value_indexes = [col - first_col for col in value_cols]

    for values in matrix:
        raw_seo = values[seo_index]
        raw_source = values[source_index]
        seo = normalize(raw_seo)
        source = source_key(raw_source)
        if not is_usable_key(seo):
            continue
        if not source:
            raise RuntimeError(
                f"{label} có Title [SEO] nhưng thiếu File nguồn: "
                f"{str(raw_seo).strip()}"
            )

        key = (seo, source)
        if key in lookup:
            duplicates.append(
                f"{str(raw_seo).strip()} | File nguồn: "
                f"{str(raw_source).strip()}"
            )
            continue

        lookup[key] = tuple(values[index] for index in value_indexes)
        keys_by_title.setdefault(seo, []).append(key)

    if duplicates:
        preview = ", ".join(duplicates[:10])
        raise RuntimeError(
            f"{label} bị trùng cả Title [SEO] và File nguồn: {preview}"
        )
    return lookup, keys_by_title


def main() -> None:
    print("=" * 70)
    print("CHUẨN BỊ DỮ LIỆU ĐĂNG BÀI THEO TIÊU ĐỀ SEO")
    print("=" * 70)

    excel = win32.GetActiveObject("Excel.Application")
    workbook = excel.ActiveWorkbook
    if workbook is None:
        raise RuntimeError("Không tìm thấy workbook Excel đang mở.")

    ws_publish = workbook.Worksheets(SHEET_PUBLISH)
    ws_plan = workbook.Worksheets(SHEET_PLAN)
    ws_write = workbook.Worksheets(SHEET_WRITE)

    publish_cols = require_columns(
        ws_publish,
        [
            HEADER_SEO_PUBLISH,
            HEADER_TITLE_PUBLISH,
            HEADER_CATEGORY_PUBLISH,
            HEADER_H1_PUBLISH,
            HEADER_WORD_PUBLISH,
            HEADER_IMAGE1_PUBLISH,
            HEADER_IMAGE2_PUBLISH,
        ],
    )
    plan_cols = require_columns(
        ws_plan,
        [HEADER_SEO_PLAN, HEADER_H1_PLAN, HEADER_SOURCE_FILE_PLAN],
    )
    plan_cols[HEADER_CATEGORY_PLAN] = require_any_column(
        ws_plan,
        [
            HEADER_CATEGORY_PLAN,
            HEADER_CATEGORY_PLAN_ALTERNATIVE,
        ],
    )
    plan_available = headers(ws_plan)
    domain_plan_col = plan_available.get(normalize(HEADER_DOMAIN_PLAN))
    source_file_plan_col = plan_cols[HEADER_SOURCE_FILE_PLAN]
    should_fill_domain = (
        source_file_plan_col is not None or domain_plan_col is not None
    )
    if should_fill_domain:
        publish_cols.update(
            require_columns(ws_publish, [HEADER_DOMAIN_PUBLISH])
        )
    write_cols = require_columns(
        ws_write,
        [
            HEADER_SEO_WRITE,
            HEADER_TITLE_WRITE,
            HEADER_WORD_WRITE,
            HEADER_OLD_WORD_COUNT_WRITE,
            HEADER_NEW_WORD_WRITE,
            HEADER_NEW_WORD_COUNT_WRITE,
            HEADER_IMAGE1_WRITE,
            HEADER_IMAGE2_WRITE,
        ],
    )

    plan_value_cols = [
        plan_cols[HEADER_CATEGORY_PLAN],
        plan_cols[HEADER_H1_PLAN],
    ]
    source_file_value_index = None
    domain_value_index = None
    plan_value_cols.append(source_file_plan_col)
    source_file_value_index = len(plan_value_cols) - 1
    if domain_plan_col is not None:
        plan_value_cols.append(domain_plan_col)
        domain_value_index = len(plan_value_cols) - 1

    plan_lookup, plan_keys_by_title = build_plan_lookup(
        ws_plan,
        plan_cols[HEADER_SEO_PLAN],
        source_file_plan_col,
        plan_value_cols,
        SHEET_PLAN,
    )
    write_lookup = build_unique_lookup(
        ws_write,
        write_cols[HEADER_SEO_WRITE],
        [
            write_cols[HEADER_TITLE_WRITE],
            write_cols[HEADER_WORD_WRITE],
            write_cols[HEADER_OLD_WORD_COUNT_WRITE],
            write_cols[HEADER_NEW_WORD_WRITE],
            write_cols[HEADER_NEW_WORD_COUNT_WRITE],
            write_cols[HEADER_IMAGE1_WRITE],
            write_cols[HEADER_IMAGE2_WRITE],
        ],
        SHEET_WRITE,
    )

    seo_col = publish_cols[HEADER_SEO_PUBLISH]
    last_row = last_data_row(ws_publish, seo_col)
    assert_publish_range_is_safe(ws_publish, last_row)
    if last_row < 2:
        print("DANG_BAI chưa có Tiêu đề SEO để xử lý.")
        return

    updated = 0
    used_new_word = 0
    new_word_rows: list[int] = []
    missing: list[str] = []
    old_screen_updating = excel.ScreenUpdating
    old_enable_events = excel.EnableEvents
    try:
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        # Đọc SEO và 6 cột đầu ra theo khối, xử lý trong bộ nhớ rồi ghi mỗi cột đúng một lần.
        seo_values = read_matrix(ws_publish, 2, last_row, seo_col, seo_col)
        output_headers = [
            HEADER_TITLE_PUBLISH,
            HEADER_CATEGORY_PUBLISH,
            HEADER_H1_PUBLISH,
            HEADER_WORD_PUBLISH,
            HEADER_IMAGE1_PUBLISH,
            HEADER_IMAGE2_PUBLISH,
        ]
        if should_fill_domain:
            output_headers.append(HEADER_DOMAIN_PUBLISH)
        output_values = {
            header: [list(item) for item in read_matrix(
                ws_publish,
                2,
                last_row,
                publish_cols[header],
                publish_cols[header],
            )]
            for header in output_headers
        }

        for offset, seo_item in enumerate(seo_values):
            row = offset + 2
            raw_seo = seo_item[0]
            key = normalize(raw_seo)
            if not is_usable_key(key):
                continue

            current_domain = output_values[
                HEADER_DOMAIN_PUBLISH
            ][offset][0]
            plan_data = None
            plan_candidates = plan_keys_by_title.get(key, [])
            current_source_key = source_key(current_domain)
            if current_source_key:
                plan_data = plan_lookup.get((key, current_source_key))
            if plan_data is None and len(plan_candidates) == 1:
                # Tương thích dữ liệu cũ: tiêu đề chỉ có một File nguồn thì
                # vẫn đối chiếu được dù Tên miền ở DANG_BAI đang trống.
                plan_data = plan_lookup[plan_candidates[0]]
            write_data = write_lookup.get(key)
            if plan_data is None or write_data is None:
                sources = []
                if plan_data is None:
                    if len(plan_candidates) > 1 and not current_source_key:
                        sources.append("KE_HOACH: cần Tên miền để chọn File nguồn")
                    else:
                        sources.append(SHEET_PLAN)
                if write_data is None:
                    sources.append(SHEET_WRITE)
                missing.append(
                    f"Dòng {row}: {str(raw_seo).strip()} (thiếu {', '.join(sources)})"
                )
                continue

            category, h1 = plan_data[:2]
            domain = ""
            if source_file_value_index is not None:
                domain = domain_from_source_file(
                    plan_data[source_file_value_index]
                )
            if not domain and domain_value_index is not None:
                fallback_domain = plan_data[domain_value_index]
                if has_value(fallback_domain):
                    domain = str(fallback_domain).strip()
            (
                title,
                old_word_path,
                old_word_count_raw,
                new_word_path,
                new_word_count_raw,
                image1_path,
                image2_path,
            ) = write_data

            old_word_count = to_number(old_word_count_raw)
            new_word_count = to_number(new_word_count_raw)
            use_new_word = (
                has_value(new_word_path)
                and old_word_count is not None
                and new_word_count is not None
                and new_word_count >= old_word_count + MIN_EXTRA_WORDS
            )
            selected_word_path = new_word_path if use_new_word else old_word_path
            if use_new_word:
                used_new_word += 1
                new_word_rows.append(row)

            new_values = (
                title,
                category,
                h1,
                selected_word_path,
                image1_path,
                image2_path,
            )
            if should_fill_domain:
                new_values += (
                    domain if domain else current_domain,
                )
            for header, value in zip(output_headers, new_values):
                output_values[header][offset][0] = value
            updated += 1

        for header in output_headers:
            ws_publish.Range(
                ws_publish.Cells(2, publish_cols[header]),
                ws_publish.Cells(last_row, publish_cols[header]),
            ).Value2 = tuple(tuple(item) for item in output_values[header])

        # Chỉ cột Đường dẫn Word ở sheet DANG_BAI được đổi màu.
        # Xóa dấu màu từ lần chạy trước, sau đó tô vàng các dòng đang dùng Word đợt 2.
        word_publish_col = publish_cols[HEADER_WORD_PUBLISH]
        word_output_range = ws_publish.Range(
            ws_publish.Cells(2, word_publish_col),
            ws_publish.Cells(last_row, word_publish_col),
        )
        word_output_range.Interior.Pattern = -4142  # xlPatternNone

        # Gom các dòng liên tiếp để giảm số lệnh COM khi có hàng nghìn dòng.
        if new_word_rows:
            range_start = new_word_rows[0]
            range_end = range_start
            for marked_row in new_word_rows[1:] + [None]:
                if marked_row is not None and marked_row == range_end + 1:
                    range_end = marked_row
                    continue
                ws_publish.Range(
                    ws_publish.Cells(range_start, word_publish_col),
                    ws_publish.Cells(range_end, word_publish_col),
                ).Interior.Color = NEW_WORD_HIGHLIGHT_COLOR
                if marked_row is not None:
                    range_start = marked_row
                    range_end = marked_row

        workbook.Save()
    finally:
        excel.ScreenUpdating = old_screen_updating
        excel.EnableEvents = old_enable_events

    print(f"Đã cập nhật: {updated} dòng.")
    print(
        f"Đã chọn Word đợt 2: {used_new_word} dòng "
        f"(nhiều hơn Word cũ ít nhất {MIN_EXTRA_WORDS} từ, ô đích được tô vàng)."
    )
    if missing:
        print(f"Không đối chiếu được: {len(missing)} dòng.")
        for item in missing[:30]:
            print("- " + item)
    else:
        print("Tất cả Tiêu đề SEO đều đối chiếu thành công.")
    if should_fill_domain:
        print(
            "Đối chiếu KE_HOACH bằng Title [SEO] + File nguồn; "
            "DANG_BAI dùng Tiêu đề SEO + Tên miền. "
            "Tên miền: ưu tiên lấy domain từ KE_HOACH.File nguồn; "
            "nếu không có thì dùng KE_HOACH.Tên Miền; "
            "nếu cả hai trống thì giữ nguyên DANG_BAI."
        )
    else:
        print(
            "KE_HOACH không có cột File nguồn hoặc Tên Miền; "
            "Tên miền và các cột nhập tay "
            "trong DANG_BAI được giữ nguyên."
        )


if __name__ == "__main__":
    main()
