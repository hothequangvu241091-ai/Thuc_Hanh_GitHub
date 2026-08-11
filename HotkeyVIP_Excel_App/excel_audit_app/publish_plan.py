from __future__ import annotations

import os
from collections import defaultdict
from pathlib import Path
from typing import Any

from .excel_io import OpenXmlWorkbook, normalize_spaces, normalize_text


ALIASES = {
    "title": ["Tiêu đề", "Từ khóa", "Main Keyword"],
    "seo_title": ["Tiêu đề SEO", "Title [SEO]", "Title SEO"],
    "h1": ["H1"],
    "status": ["Trạng thái đăng"],
    "domain": ["Tên Miền", "Tên miền"],
    "category": ["Danh mục", "CATE [POST]"],
    "word_path": ["Đường dẫn Word"],
}


def _status_priority(value: Any) -> int | None:
    status = normalize_text(value)
    if "lỗi đăng" in status or "lỗi kiểm tra" in status:
        return 0
    if "cần mở" in status or status == "ok":
        return 1
    if not status:
        return 2
    return None


def _resolved_path(raw: Any) -> str:
    value = normalize_spaces(raw).strip('"')
    if not value:
        return ""
    return str(Path(os.path.expandvars(os.path.expanduser(value))).resolve())


def inspect_publish_queue(path: str | Path) -> dict[str, Any]:
    """Đọc các bài có thể đưa vào batch đăng mà không mở Microsoft Excel."""
    workbook = OpenXmlWorkbook(path)
    sheet = workbook.find_sheet("DANG_BAI")
    if sheet is None:
        raise RuntimeError("Workbook không có sheet DANG_BAI.")
    table = workbook.read_sheet(sheet)
    columns = {
        name: table.find_column(aliases, required=True)
        for name, aliases in ALIASES.items()
    }

    eligible: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    for row_number, row in table.rows:
        title = normalize_spaces(table.value(row, columns["title"]))
        if not title:
            continue
        item = {
            "row": row_number,
            **{
                name: normalize_spaces(table.value(row, column))
                for name, column in columns.items()
            },
        }
        priority = _status_priority(item["status"])
        if priority is None:
            continue
        item["status_priority"] = priority
        item["word_path_resolved"] = _resolved_path(item["word_path"])
        reasons = []
        if not item["domain"]:
            reasons.append("thiếu tên miền")
        if not item["category"]:
            reasons.append("thiếu danh mục")
        if not item["word_path_resolved"] or not Path(item["word_path_resolved"]).is_file():
            reasons.append("Word không tồn tại")
        if reasons:
            skipped.append({**item, "reason": ", ".join(reasons)})
            continue
        eligible.append(item)

    eligible.sort(key=lambda item: (int(item["status_priority"]), int(item["row"])))
    return {"eligible": eligible, "skipped": skipped}


def build_balanced_publish_plan(
    snapshot: dict[str, Any],
    per_domain_limit: int,
    category_overrides: dict[str, str] | None = None,
) -> dict[str, Any]:
    """Mỗi domain chọn một danh mục; mặc định lớn nhất, có thể ghi đè từ app."""
    limit = max(1, int(per_domain_limit))
    overrides = {
        normalize_text(domain): normalize_text(category)
        for domain, category in (category_overrides or {}).items()
        if normalize_text(domain) and normalize_text(category)
    }
    grouped: dict[str, dict[str, list[dict[str, Any]]]] = defaultdict(lambda: defaultdict(list))
    displays: dict[tuple[str, str], tuple[str, str]] = {}
    first_rows: dict[tuple[str, str], int] = {}

    for item in snapshot["eligible"]:
        domain_key = normalize_text(item["domain"])
        category_key = normalize_text(item["category"])
        grouped[domain_key][category_key].append(item)
        displays.setdefault((domain_key, category_key), (item["domain"], item["category"]))
        first_rows.setdefault((domain_key, category_key), int(item["row"]))

    groups: list[dict[str, Any]] = []
    selected_rows: list[dict[str, Any]] = []
    domain_order = sorted(
        grouped,
        key=lambda key: min(
            first_rows[(key, category_key)] for category_key in grouped[key]
        ),
    )
    for domain_key in domain_order:
        categories = grouped[domain_key]
        ordered_categories = sorted(
            categories,
            key=lambda category_key: (
                -len(categories[category_key]),
                first_rows[(domain_key, category_key)],
            ),
        )
        requested_category = overrides.get(domain_key, "")
        selected_category = (
            requested_category
            if requested_category in categories
            else ordered_categories[0]
        )
        candidates = categories[selected_category]
        chosen = candidates[:limit]
        domain, category = displays[(domain_key, selected_category)]
        category_options = [
            {
                "key": category_key,
                "label": displays[(domain_key, category_key)][1],
                "available": len(categories[category_key]),
            }
            for category_key in ordered_categories
        ]
        groups.append(
            {
                "domain": domain,
                "domain_key": domain_key,
                "category": category,
                "category_key": selected_category,
                "category_options": category_options,
                "available": len(candidates),
                "selected": len(chosen),
                "first_row": first_rows[(domain_key, selected_category)],
            }
        )
        for item in chosen:
            selected_rows.append(
                {
                    "row": int(item["row"]),
                    "domain": item["domain"],
                    "category": item["category"],
                    "title": item["title"],
                    "seo_title": item["seo_title"],
                    "h1": item["h1"],
                }
            )

    return {
        "mode": "balanced_one_category",
        "per_domain_limit": limit,
        "groups": groups,
        "selected_rows": selected_rows,
        "selected_total": len(selected_rows),
        "eligible_total": len(snapshot["eligible"]),
        "skipped_invalid_total": len(snapshot["skipped"]),
        "category_overrides": {
            group["domain_key"]: group["category_key"] for group in groups
        },
    }
