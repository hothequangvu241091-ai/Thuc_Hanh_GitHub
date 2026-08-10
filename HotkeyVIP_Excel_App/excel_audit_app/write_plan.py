from __future__ import annotations

from collections import Counter
from pathlib import Path
from typing import Any

from .excel_io import OpenXmlWorkbook, normalize_spaces, normalize_text


START_ROW = 2
END_ROW = 10000
MANUAL_MARK_TEXT = "OK OK"

ALIASES = {
    "domain": ["Tên Miền", "Tên miền"],
    "keyword": ["Từ khóa", "Main Keyword"],
    "done": ["Trạng thái hoàn tất"],
    "manual_mark": ["Mốc bắt đầu"],
    "article_status": ["Trạng thái viết"],
    "article_error": ["Lỗi viết"],
    "brief_status": ["Trạng thái brief", "Trạng thái Brief"],
    "retry_error": ["Lỗi thử lại"],
}


def _is_open(item: dict[str, Any]) -> bool:
    return bool(normalize_spaces(item.get("keyword"))) and (
        normalize_text(item.get("done")) != "ok"
    )


def _is_error(item: dict[str, Any]) -> bool:
    if not _is_open(item):
        return False
    article_status = normalize_text(item.get("article_status"))
    brief_status = normalize_text(item.get("brief_status"))
    done = normalize_text(item.get("done"))
    return bool(
        normalize_spaces(item.get("article_error"))
        or normalize_spaces(item.get("retry_error"))
        or article_status in {"error", "word_error"}
        or "lỗi" in brief_status
        or "error" in brief_status
        or "lỗi" in done
        or "error" in done
        or "tạm bỏ lượt" in done
    )


def inspect_write_queue(path: str | Path) -> dict[str, Any]:
    """Đọc preview VIET_BAI bằng Open XML, không mở Microsoft Excel."""
    workbook = OpenXmlWorkbook(path)
    sheet = workbook.find_sheet("VIET_BAI")
    if sheet is None:
        raise RuntimeError("Workbook không có sheet VIET_BAI.")
    table = workbook.read_sheet(sheet)
    columns = {
        name: table.find_column(aliases, required=name in {"domain", "keyword", "done"})
        for name, aliases in ALIASES.items()
    }

    items: list[dict[str, Any]] = []
    for row_number, row in table.rows:
        if row_number < START_ROW or row_number > END_ROW:
            continue
        item = {
            "row": row_number,
            **{
                name: table.value(row, column)
                for name, column in columns.items()
            },
        }
        item["domain"] = normalize_spaces(item.get("domain"))
        item["keyword"] = normalize_spaces(item.get("keyword"))
        item["is_open"] = _is_open(item)
        item["is_error"] = _is_error(item)
        items.append(item)

    manual_rows = [
        int(item["row"])
        for item in items
        if normalize_spaces(item.get("manual_mark")) == MANUAL_MARK_TEXT
    ]
    manual_row = max(manual_rows) if manual_rows else None
    normal_start_row = (manual_row + 1) if manual_row is not None else START_ROW

    open_items = [item for item in items if item["is_open"]]
    error_items = [item for item in open_items if item["is_error"]]
    domain_counts: Counter[str] = Counter(
        str(item["domain"] or "(Không có tên miền)") for item in open_items
    )
    domain_error_counts: Counter[str] = Counter(
        str(item["domain"] or "(Không có tên miền)") for item in error_items
    )
    normal_items = [
        item for item in open_items if int(item["row"]) >= normal_start_row
    ]
    return {
        "items": items,
        "open_items": open_items,
        "error_items": error_items,
        "normal_items": normal_items,
        "manual_row": manual_row,
        "normal_start_row": normal_start_row,
        "domain_counts": dict(domain_counts),
        "domain_error_counts": dict(domain_error_counts),
    }


def build_write_queue_preview(snapshot: dict[str, Any], plan: dict[str, Any]) -> list[dict[str, Any]]:
    """Dựng đúng thứ tự preview: lỗi -> domain ưu tiên -> bình thường."""
    selected_rows: set[int] = set()
    result: list[dict[str, Any]] = []

    def add(items: list[dict[str, Any]]) -> None:
        for item in items:
            row = int(item["row"])
            if row in selected_rows:
                continue
            selected_rows.add(row)
            result.append(item)

    if plan.get("retry_errors_first", True):
        add(list(snapshot["error_items"]))

    priority_domain = normalize_text(plan.get("priority_domain"))
    try:
        priority_count = max(0, int(plan.get("priority_count", 0)))
    except (TypeError, ValueError):
        priority_count = 0
    if priority_domain and priority_count:
        candidates = [
            item
            for item in snapshot["open_items"]
            if int(item["row"]) not in selected_rows
            and normalize_text(item.get("domain")) == priority_domain
        ]
        add(candidates[:priority_count])

    if plan.get("continue_normal", True):
        add(list(snapshot["normal_items"]))
    return result
