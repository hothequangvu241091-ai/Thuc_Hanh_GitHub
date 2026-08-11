from __future__ import annotations

import os
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Iterable

from .excel_io import OpenXmlWorkbook, normalize_spaces, normalize_text


DANG_ALIASES = {
    "keyword": ["Tiêu đề", "Từ khóa", "Main Keyword"],
    "seo_title": ["Tiêu đề SEO", "Title [SEO]", "Title SEO"],
    "h1": ["H1", "Article Name [H1]"],
    "status": ["Trạng thái đăng"],
    "domain": ["Tên miền", "Tên Miền"],
    "category": ["Danh mục", "CATE [POST]"],
    "slug": ["Slug"],
    "related": ["Bài viết liên quan"],
    "cms_id": ["ID CMS"],
    "published_url": ["URL đã đăng", "URL"],
    "word_path": ["Đường dẫn Word"],
    "image1_path": ["Đường dẫn ảnh 1"],
    "image2_path": ["Đường dẫn ảnh 2"],
    "article_id": ["Mã bài"],
    "published_at": ["Thời gian đăng"],
    "publish_error": ["Lỗi đăng"],
}

VIET_ALIASES = {
    "keyword": ["Từ khóa", "Main Keyword"],
    "seo_title": ["Tiêu đề SEO", "Title [SEO]", "Title SEO"],
    "h1": ["H1", "Article Name [H1]"],
    "domain": ["Tên miền", "Tên Miền"],
    "gpt_url": ["URL GPT gốc"],
    "chat_url": ["URL ChatGPT"],
}


def _columns(table: Any, aliases: dict[str, list[str]]) -> dict[str, int | None]:
    return {
        key: table.find_column(names, required=False)
        for key, names in aliases.items()
    }


def _combo(values: Iterable[Any]) -> tuple[str, ...] | None:
    result = tuple(normalize_text(value) for value in values)
    return result if all(result) else None


def _resolved_path(raw: Any) -> str:
    value = normalize_spaces(raw).strip('"')
    if not value:
        return ""
    return str(Path(os.path.expandvars(os.path.expanduser(value))).resolve())


def _excel_date(value: Any) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
        except (OverflowError, ValueError):
            return None
    text = normalize_spaces(value)
    for fmt in (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    try:
        return datetime.fromisoformat(text).date()
    except ValueError:
        return None


def _display_datetime(value: Any) -> str:
    if isinstance(value, (int, float)):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).strftime(
                "%d/%m/%Y %H:%M:%S"
            )
        except (OverflowError, ValueError):
            pass
    return normalize_spaces(value)


def _is_error_item(status: str, cms_id: str) -> bool:
    key = normalize_text(status)
    if "lỗi" in key or "lỗi" in normalize_text(cms_id):
        return True
    return "đã đăng" in key and not normalize_spaces(cms_id)


def _is_posted_today(status: str, cms_id: str, published_at: Any, today: date) -> bool:
    normalized_id = normalize_spaces(cms_id)
    return (
        "đã đăng" in normalize_text(status)
        and normalized_id.isdigit()
        and _excel_date(published_at) == today
    )


def build_publish_review_from_tables(dang: Any, viet: Any, today: date | None = None) -> dict[str, Any]:
    """Build the review from tables already read by the main analysis."""
    dc = _columns(dang, DANG_ALIASES)
    vc = _columns(viet, VIET_ALIASES)
    required = ("keyword", "status", "domain", "cms_id", "word_path", "published_at")
    missing = [key for key in required if dc.get(key) is None]
    if missing:
        raise RuntimeError("DANG_BAI thiếu cột cần thiết: " + ", ".join(missing))

    viet_urls: dict[tuple[str, ...], dict[str, str]] = {}
    for _row_number, row in viet.rows:
        combo = _combo(
            (
                viet.value(row, vc["domain"]),
                viet.value(row, vc["seo_title"]),
                viet.value(row, vc["h1"]),
                viet.value(row, vc["keyword"]),
            )
        )
        if combo and combo not in viet_urls:
            viet_urls[combo] = {
                "gpt_url": normalize_spaces(viet.value(row, vc["gpt_url"])),
                "chat_url": normalize_spaces(viet.value(row, vc["chat_url"])),
            }

    error_rows: list[dict[str, Any]] = []
    posted_today: list[dict[str, Any]] = []
    target_date = today or date.today()
    for row_number, row in dang.rows:
        keyword = normalize_spaces(dang.value(row, dc["keyword"]))
        if not keyword:
            continue
        item = {
            "row": row_number,
            **{
                key: normalize_spaces(dang.value(row, column))
                for key, column in dc.items()
                if key != "published_at"
            },
        }
        raw_published_at = dang.value(row, dc["published_at"])
        item["published_at"] = _display_datetime(raw_published_at)
        item["word_path_resolved"] = _resolved_path(item.get("word_path", ""))
        status_key = normalize_text(item.get("status", ""))
        cms_id_key = normalize_text(item.get("cms_id", ""))
        if "lỗi" in cms_id_key:
            item["display_status"] = "LỖI ID"
        elif "đã đăng" in status_key and not normalize_spaces(item.get("cms_id", "")):
            item["display_status"] = "THIẾU ID CMS"
        else:
            item["display_status"] = item.get("status", "")
        combo = _combo((item["domain"], item["seo_title"], item["h1"], item["keyword"]))
        item.update(viet_urls.get(combo or (), {"gpt_url": "", "chat_url": ""}))
        if _is_error_item(item["status"], item["cms_id"]):
            error_rows.append(item)
        if _is_posted_today(item["status"], item["cms_id"], raw_published_at, target_date):
            posted_today.append(item)

    retry_rows = [
        item for item in error_rows
        if "lỗi kiểm tra" in normalize_text(item.get("status", ""))
    ]
    return {
        "errors": error_rows,
        "posted_today": posted_today,
        "retry_rows": retry_rows,
        "today": target_date.isoformat(),
    }


def inspect_publish_review(path: str | Path, today: date | None = None) -> dict[str, Any]:
    """Read error rows and today's posted rows without opening Microsoft Excel."""
    workbook = OpenXmlWorkbook(path)
    dang_sheet = workbook.find_sheet("DANG_BAI")
    viet_sheet = workbook.find_sheet("VIET_BAI")
    if dang_sheet is None or viet_sheet is None:
        raise RuntimeError("Workbook phải có cả sheet DANG_BAI và VIET_BAI.")
    return build_publish_review_from_tables(
        workbook.read_sheet(dang_sheet),
        workbook.read_sheet(viet_sheet),
        today,
    )


def build_retry_publish_plan(review: dict[str, Any]) -> dict[str, Any]:
    selected = []
    for item in review.get("retry_rows", []):
        selected.append(
            {
                "row": int(item["row"]),
                "domain": item.get("domain", ""),
                "category": item.get("category", ""),
                "title": item.get("keyword", ""),
                "seo_title": item.get("seo_title", ""),
                "h1": item.get("h1", ""),
            }
        )
    return {
        "mode": "explicit_error_rows",
        "selected_rows": selected,
        "selected_total": len(selected),
    }
