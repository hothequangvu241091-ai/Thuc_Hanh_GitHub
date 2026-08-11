from __future__ import annotations

from collections import defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any, Callable, Iterable

from .excel_io import (
    OpenXmlWorkbook,
    SheetTable,
    file_fingerprint,
    is_blank,
    is_valid_url,
    normalize_spaces,
    normalize_text,
)
from .publish_review import build_publish_review_from_tables


APP_VERSION = "1.5.0"
EMPTY_DOMAIN = "(Không có tên miền)"


ALIASES = {
    "ke_hoach": {
        "title": ["Title [SEO]", "Title SEO", "Tiêu đề SEO"],
        "url": ["URL Page"],
        "domain": ["Tên Miền", "Tên miền"],
        "keyword": ["Main Keyword", "Từ khóa"],
        "h1": ["Article Name [H1]", "H1"],
        "category": ["CATE [POST]", "Danh mục"],
        "source_status": ["Trạng thái nguồn"],
    },
    "viet_bai": {
        "title": ["Tiêu đề SEO", "Title [SEO]", "Title SEO"],
        "h1": ["H1"],
        "domain": ["Tên Miền", "Tên miền"],
        "keyword": ["Từ khóa", "Main Keyword"],
        "word": ["Đường dẫn Word"],
        "image1": ["Đường dẫn ảnh 1"],
        "image2": ["Đường dẫn ảnh 2"],
        "completed": ["Trạng thái hoàn tất"],
    },
    "dang_bai": {
        "keyword": ["Tiêu đề"],
        "title": ["Tiêu đề SEO"],
        "post_status": ["Trạng thái đăng"],
        "domain": ["Tên Miền", "Tên miền"],
        "h1": ["H1"],
        "url": ["URL đã đăng", "URL"],
        "word": ["Đường dẫn Word"],
        "image1": ["Đường dẫn ảnh 1"],
        "image2": ["Đường dẫn ảnh 2"],
        "category": ["Danh mục", "CATE [POST]"],
    },
}


OPTIONAL_COLUMNS = {
    "ke_hoach": {"source_status", "category"},
    "viet_bai": set(),
    "dang_bai": {"category"},
}


class DomainRegistry:
    def __init__(self) -> None:
        self._display: dict[str, str] = {}

    def register(self, value: Any) -> tuple[str, str]:
        display = normalize_spaces(value)
        if not display:
            return "", EMPTY_DOMAIN
        key = normalize_text(display)
        self._display.setdefault(key, display)
        return key, self._display[key]

    def display(self, key: str) -> str:
        if key == "":
            return EMPTY_DOMAIN
        return self._display.get(key, key)

    def ordered_keys(self, keys: Iterable[str]) -> list[str]:
        return sorted(
            set(keys),
            key=lambda key: (key == "", self.display(key).casefold()),
        )


def _columns(table: SheetTable, schema: str) -> dict[str, int | None]:
    optional = OPTIONAL_COLUMNS[schema]
    return {
        logical: table.find_column(aliases, required=logical not in optional)
        for logical, aliases in ALIASES[schema].items()
    }


def _combo(values: Iterable[Any]) -> tuple[str, ...] | None:
    normalized = tuple(normalize_text(value) for value in values)
    if any(not value for value in normalized):
        return None
    return normalized


def _row_has_content(table: SheetTable, row: dict[int, Any], ignored: set[int | None]) -> bool:
    for column in table.headers:
        if column not in ignored and not is_blank(row.get(column, "")):
            return True
    return False


def _new_ke_metrics() -> dict[str, int]:
    return {
        "total_rows": 0,
        "combo4_complete": 0,
        "combo4_missing": 0,
        "url_valid": 0,
        "url_written": 0,
        "url_blank": 0,
        "url_other": 0,
        "problem_rows": 0,
        "duplicate_groups": 0,
        "duplicate_rows": 0,
        "missing_in_viet": 0,
    }


def _new_viet_metrics() -> dict[str, int]:
    return {
        "total_rows": 0,
        "combo4_complete": 0,
        "combo4_missing": 0,
        "completed_ok": 0,
        "has_word": 0,
        "has_word_images": 0,
        "missing_word_images": 0,
        "duplicate_groups": 0,
        "duplicate_rows": 0,
        "missing_in_ke": 0,
        "in_dang": 0,
        "pending_dang": 0,
        "recovery_dang": 0,
        "completed_with_assets": 0,
        "archived_posted_no_assets": 0,
        "recovery_no_assets": 0,
        "unexplained_no_assets": 0,
        "not_completed": 0,
        "classified_total": 0,
        "classification_difference": 0,
    }


def _new_dang_metrics() -> dict[str, int]:
    return {
        "total_rows": 0,
        "combo4_complete": 0,
        "combo4_missing": 0,
        "posted": 0,
        "url_not_posted_full_assets": 0,
        "dang_missing_viet": 0,
        "viet_missing_dang": 0,
        "ke_url_missing_dang": 0,
        "in_viet": 0,
        "classified_total": 0,
        "classification_difference": 0,
    }


def _issue(
    issues: list[dict[str, Any]],
    *,
    category: str,
    sheet: str,
    row: int,
    domain: str,
    title: Any,
    h1: Any,
    keyword: Any,
    detail: str,
    level: str = "error",
    target_sheet: str = "",
    target_row: int | str = "",
) -> None:
    issues.append(
        {
            "category": category,
            "sheet": sheet,
            "row": row,
            "domain": domain,
            "title": normalize_spaces(title),
            "h1": normalize_spaces(h1),
            "keyword": normalize_spaces(keyword),
            "detail": detail,
            "level": level,
            "target_sheet": target_sheet,
            "target_row": target_row,
        }
    )


def _summary_rows(
    metrics: dict[str, dict[str, int]],
    registry: DomainRegistry,
    factory: Callable[[], dict[str, int]],
    total_overrides: dict[str, int] | None = None,
) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    total = factory()
    for domain_key in registry.ordered_keys(metrics.keys()):
        values = metrics[domain_key]
        rows.append({"domain": registry.display(domain_key), "domain_key": domain_key, **values})
        for key, value in values.items():
            total[key] += value
    if total_overrides:
        total.update(total_overrides)
    return {"rows": rows, "total": {"domain": "TỔNG TẤT CẢ", "domain_key": "__total__", **total}}


def analyze_workbook(path: str | Path) -> dict[str, Any]:
    workbook = OpenXmlWorkbook(path)
    required = workbook.require_sheets(["KE_HOACH", "VIET_BAI", "DANG_BAI"])
    ke_table = workbook.read_sheet(required["KE_HOACH"])
    viet_table = workbook.read_sheet(required["VIET_BAI"])
    dang_table = workbook.read_sheet(required["DANG_BAI"])

    ke_cols = _columns(ke_table, "ke_hoach")
    viet_cols = _columns(viet_table, "viet_bai")
    dang_cols = _columns(dang_table, "dang_bai")

    registry = DomainRegistry()
    issues: list[dict[str, Any]] = []
    ke_metrics: dict[str, dict[str, int]] = defaultdict(_new_ke_metrics)
    viet_metrics: dict[str, dict[str, int]] = defaultdict(_new_viet_metrics)
    dang_metrics: dict[str, dict[str, int]] = defaultdict(_new_dang_metrics)

    ke_records: list[dict[str, Any]] = []
    viet_records: list[dict[str, Any]] = []
    dang_records: list[dict[str, Any]] = []

    # KE_HOACH
    ke_ignored = {ke_cols["source_status"]}
    required_data_columns = [
        column
        for column in ke_table.headers
        if column not in {ke_cols["url"], ke_cols["source_status"]}
    ]
    for row_number, row in ke_table.rows:
        if not _row_has_content(ke_table, row, ke_ignored):
            continue
        domain_key, domain = registry.register(ke_table.value(row, ke_cols["domain"]))
        title = ke_table.value(row, ke_cols["title"])
        h1 = ke_table.value(row, ke_cols["h1"])
        keyword = ke_table.value(row, ke_cols["keyword"])
        url = ke_table.value(row, ke_cols["url"])
        category = ke_table.value(row, ke_cols["category"])
        combo4 = _combo((domain, title, h1, keyword)) if domain_key else None
        missing_fields = [
            ke_table.headers[column]
            for column in required_data_columns
            if is_blank(row.get(column, ""))
        ]
        if is_blank(url):
            url_class = "blank"
        elif normalize_text(url) == "đã viết":
            url_class = "written"
        elif is_valid_url(url):
            url_class = "valid"
        else:
            url_class = "other"
        record = {
            "row": row_number,
            "domain_key": domain_key,
            "domain": domain,
            "title": title,
            "h1": h1,
            "keyword": keyword,
            "url": url,
            "category": category,
            "url_class": url_class,
            "combo4": combo4,
            "missing_fields": missing_fields,
            "duplicate": False,
        }
        ke_records.append(record)
        metric = ke_metrics[domain_key]
        metric["total_rows"] += 1
        metric["combo4_complete" if combo4 else "combo4_missing"] += 1
        metric[f"url_{url_class}"] += 1

    ke_duplicate_map: dict[tuple[str, ...], list[dict[str, Any]]] = defaultdict(list)
    for record in ke_records:
        if record["combo4"]:
            ke_duplicate_map[record["combo4"]].append(record)
    ke_duplicate_groups = [group for group in ke_duplicate_map.values() if len(group) > 1]
    for group in ke_duplicate_groups:
        domain_key = group[0]["domain_key"]
        ke_metrics[domain_key]["duplicate_groups"] += 1
        ke_metrics[domain_key]["duplicate_rows"] += len(group)
        for record in group:
            record["duplicate"] = True

    status_updates: dict[str, str] = {}
    for record in ke_records:
        has_problem = bool(record["missing_fields"]) or record["url_class"] == "other"
        metric = ke_metrics[record["domain_key"]]
        if record["duplicate"]:
            status = "Bài viết trùng"
            detail = "Combo 4 bị trùng"
            if record["missing_fields"]:
                detail += "; đồng thời thiếu: " + ", ".join(record["missing_fields"])
            if record["url_class"] == "other":
                detail += "; URL Page không hợp lệ"
            _issue(
                issues,
                category="Bài viết trùng",
                sheet=ke_table.name,
                row=record["row"],
                domain=record["domain"],
                title=record["title"],
                h1=record["h1"],
                keyword=record["keyword"],
                detail=detail,
            )
        elif has_problem:
            status = "Dữ liệu có vấn đề"
            parts: list[str] = []
            if record["missing_fields"]:
                parts.append("Thiếu: " + ", ".join(record["missing_fields"]))
            if record["url_class"] == "other":
                parts.append("URL Page không phải URL, 'Đã viết' hoặc ô trống")
            _issue(
                issues,
                category="Dữ liệu có vấn đề",
                sheet=ke_table.name,
                row=record["row"],
                domain=record["domain"],
                title=record["title"],
                h1=record["h1"],
                keyword=record["keyword"],
                detail="; ".join(parts),
            )
        else:
            status = ""
        if has_problem:
            metric["problem_rows"] += 1
        status_updates[str(record["row"])] = status

    # VIET_BAI
    for row_number, row in viet_table.rows:
        if not _row_has_content(viet_table, row, set()):
            continue
        domain_key, domain = registry.register(viet_table.value(row, viet_cols["domain"]))
        title = viet_table.value(row, viet_cols["title"])
        h1 = viet_table.value(row, viet_cols["h1"])
        keyword = viet_table.value(row, viet_cols["keyword"])
        combo4 = _combo((domain, title, h1, keyword)) if domain_key else None
        completed = normalize_text(viet_table.value(row, viet_cols["completed"])) == "ok"
        has_word = not is_blank(viet_table.value(row, viet_cols["word"]))
        word = viet_table.value(row, viet_cols["word"])
        image1 = viet_table.value(row, viet_cols["image1"])
        image2 = viet_table.value(row, viet_cols["image2"])
        has_word_images = has_word and all(
            not is_blank(viet_table.value(row, viet_cols[key])) for key in ("image1", "image2")
        )
        record = {
            "row": row_number,
            "domain_key": domain_key,
            "domain": domain,
            "title": title,
            "h1": h1,
            "keyword": keyword,
            "combo4": combo4,
            "completed": completed,
            "has_word_images": has_word_images,
            "word": word,
            "image1": image1,
            "image2": image2,
        }
        viet_records.append(record)
        metric = viet_metrics[domain_key]
        metric["total_rows"] += 1
        metric["combo4_complete" if combo4 else "combo4_missing"] += 1
        metric["completed_ok"] += int(completed)
        metric["has_word"] += int(has_word)
        metric["has_word_images"] += int(has_word_images)
        metric["missing_word_images"] += int(not has_word_images)

        if combo4 is None:
            _issue(
                issues,
                category="VIET_BAI thiếu Combo 4",
                sheet=viet_table.name,
                row=row_number,
                domain=domain,
                title=title,
                h1=h1,
                keyword=keyword,
                detail="Thiếu ít nhất một trường trong Tên miền + Tiêu đề SEO + H1 + Từ khóa",
                target_sheet=ke_table.name,
            )

    viet_duplicate_map: dict[tuple[str, ...], list[dict[str, Any]]] = defaultdict(list)
    for record in viet_records:
        if record["combo4"]:
            viet_duplicate_map[record["combo4"]].append(record)
    viet_duplicate_groups = [group for group in viet_duplicate_map.values() if len(group) > 1]
    for group in viet_duplicate_groups:
        touched_domains: set[str] = set()
        for record in group:
            viet_metrics[record["domain_key"]]["duplicate_rows"] += 1
            touched_domains.add(record["domain_key"])
            _issue(
                issues,
                category="Combo 4 trùng",
                sheet=viet_table.name,
                row=record["row"],
                domain=record["domain"],
                title=record["title"],
                h1=record["h1"],
                keyword=record["keyword"],
                detail="Tên miền + Tiêu đề SEO + H1 + Từ khóa bị trùng trong VIET_BAI",
            )
        for domain_key in touched_domains:
            viet_metrics[domain_key]["duplicate_groups"] += 1

    # DANG_BAI
    for row_number, row in dang_table.rows:
        if not _row_has_content(dang_table, row, set()):
            continue
        domain_key, domain = registry.register(dang_table.value(row, dang_cols["domain"]))
        title = dang_table.value(row, dang_cols["title"])
        h1 = dang_table.value(row, dang_cols["h1"])
        keyword = dang_table.value(row, dang_cols["keyword"])
        combo4 = _combo((domain, title, h1, keyword)) if domain_key else None
        posted = normalize_text(dang_table.value(row, dang_cols["post_status"])) == "đã đăng"
        url = dang_table.value(row, dang_cols["url"])
        full_assets = all(
            not is_blank(dang_table.value(row, dang_cols[key])) for key in ("word", "image1", "image2")
        )
        url_not_posted_full = is_valid_url(url) and not posted and full_assets
        record = {
            "row": row_number,
            "domain_key": domain_key,
            "domain": domain,
            "title": title,
            "h1": h1,
            "keyword": keyword,
            "combo4": combo4,
            "posted": posted,
            "url_not_posted_full": url_not_posted_full,
        }
        dang_records.append(record)
        metric = dang_metrics[domain_key]
        metric["total_rows"] += 1
        metric["combo4_complete" if combo4 else "combo4_missing"] += 1
        metric["posted"] += int(posted)
        metric["url_not_posted_full_assets"] += int(url_not_posted_full)

        if combo4 is None:
            _issue(
                issues,
                category="DANG_BAI thiếu Combo 4",
                sheet=dang_table.name,
                row=row_number,
                domain=domain,
                title=title,
                h1=h1,
                keyword=keyword,
                detail="Thiếu ít nhất một trường trong Tên miền + Tiêu đề SEO + H1 + Tiêu đề",
                target_sheet=viet_table.name,
            )

    ke_keys = {record["combo4"] for record in ke_records if record["combo4"]}
    viet_keys = {record["combo4"] for record in viet_records if record["combo4"]}
    dang_keys = {record["combo4"] for record in dang_records if record["combo4"]}
    ke_url_keys = {
        record["combo4"]
        for record in ke_records
        if record["combo4"] and record["url_class"] == "valid"
    }
    viet_by_key: dict[tuple[str, ...], list[dict[str, Any]]] = defaultdict(list)
    dang_by_key: dict[tuple[str, ...], list[dict[str, Any]]] = defaultdict(list)
    for record in viet_records:
        if record["combo4"]:
            viet_by_key[record["combo4"]].append(record)
    for record in dang_records:
        if record["combo4"]:
            dang_by_key[record["combo4"]].append(record)

    # Đối chiếu KE_HOACH <-> VIET_BAI theo hai chiều.
    for record in ke_records:
        if record["combo4"] and record["combo4"] not in viet_keys:
            ke_metrics[record["domain_key"]]["missing_in_viet"] += 1
            _issue(
                issues,
                category="KE_HOACH có - VIET_BAI thiếu",
                sheet=ke_table.name,
                row=record["row"],
                domain=record["domain"],
                title=record["title"],
                h1=record["h1"],
                keyword=record["keyword"],
                detail="Combo 4 có trong KE_HOACH nhưng không tồn tại trong VIET_BAI",
                target_sheet=viet_table.name,
            )

    for record in viet_records:
        if record["combo4"] and record["combo4"] not in ke_keys:
            viet_metrics[record["domain_key"]]["missing_in_ke"] += 1
            _issue(
                issues,
                category="VIET_BAI có - KE_HOACH thiếu",
                sheet=viet_table.name,
                row=record["row"],
                domain=record["domain"],
                title=record["title"],
                h1=record["h1"],
                keyword=record["keyword"],
                detail="Combo 4 có trong VIET_BAI nhưng không tồn tại trong KE_HOACH",
                target_sheet=ke_table.name,
            )

    for record in dang_records:
        metric = dang_metrics[record["domain_key"]]
        if not record["combo4"]:
            continue
        if record["combo4"] in viet_keys:
            metric["in_viet"] += 1
        else:
            metric["dang_missing_viet"] += 1
            _issue(
                issues,
                category="DANG_BAI có - VIET_BAI thiếu",
                sheet=dang_table.name,
                row=record["row"],
                domain=record["domain"],
                title=record["title"],
                h1=record["h1"],
                keyword=record["keyword"],
                detail="Combo 4 có trong DANG_BAI nhưng không có trong VIET_BAI",
                target_sheet=viet_table.name,
            )

    # Mỗi dòng VIET_BAI được phân vào đúng một nhóm tiến độ và một nhóm tài nguyên.
    for record in viet_records:
        metric = viet_metrics[record["domain_key"]]
        combo4 = record["combo4"]
        if combo4:
            if combo4 in dang_keys:
                metric["in_dang"] += 1
            elif combo4 in ke_url_keys:
                metric["recovery_dang"] += 1
                dang_metrics[record["domain_key"]]["viet_missing_dang"] += 1
            else:
                metric["pending_dang"] += 1
                dang_metrics[record["domain_key"]]["viet_missing_dang"] += 1
                _issue(
                    issues,
                    category="Chưa chuyển sang DANG_BAI",
                    sheet=viet_table.name,
                    row=record["row"],
                    domain=record["domain"],
                    title=record["title"],
                    h1=record["h1"],
                    keyword=record["keyword"],
                    detail="Bài viết hợp lệ nhưng chưa được chuyển sang DANG_BAI",
                    level="pending",
                    target_sheet=dang_table.name,
                )

        if not record["completed"]:
            metric["not_completed"] += 1
        elif record["has_word_images"]:
            metric["completed_with_assets"] += 1
        else:
            posted_matches = (
                [item for item in dang_by_key.get(combo4, []) if item["posted"]]
                if combo4 else []
            )
            if posted_matches:
                metric["archived_posted_no_assets"] += 1
                _issue(
                    issues,
                    category="Đã đăng - đã xóa tài nguyên",
                    sheet=viet_table.name,
                    row=record["row"],
                    domain=record["domain"],
                    title=record["title"],
                    h1=record["h1"],
                    keyword=record["keyword"],
                    detail="Đã OK, thiếu Word/ảnh nhưng có trạng thái ĐÃ ĐĂNG trong DANG_BAI",
                    level="info",
                    target_sheet=dang_table.name,
                    target_row=posted_matches[0]["row"],
                )
            elif combo4 and combo4 in ke_url_keys and combo4 not in dang_keys:
                metric["recovery_no_assets"] += 1
            else:
                metric["unexplained_no_assets"] += 1
                if combo4:
                    _issue(
                        issues,
                        category="VIET_BAI OK nhưng thiếu tài nguyên",
                        sheet=viet_table.name,
                        row=record["row"],
                        domain=record["domain"],
                        title=record["title"],
                        h1=record["h1"],
                        keyword=record["keyword"],
                        detail="Đã OK nhưng thiếu Word/ảnh và chưa có bằng chứng đã đăng",
                        target_sheet=dang_table.name,
                    )
        metric["classified_total"] += 1

    # Dựng sẵn các dòng đúng thứ tự cột DANG_BAI để có thể sao chép/dán trực tiếp.
    recovery_rows: list[dict[str, Any]] = []
    dang_headers = [
        dang_table.headers.get(index, "") for index in range(1, dang_table.max_column + 1)
    ]

    def build_recovery_values(
        ke_record: dict[str, Any], viet_record: dict[str, Any] | None
    ) -> list[Any]:
        values: list[Any] = [""] * dang_table.max_column

        def put(logical: str, value: Any) -> None:
            column = dang_cols.get(logical)
            if column:
                values[column - 1] = value

        put("keyword", ke_record["keyword"])
        put("title", ke_record["title"])
        put("post_status", "ĐÃ ĐĂNG")
        put("domain", ke_record["domain"])
        put("category", ke_record.get("category", ""))
        put("h1", ke_record["h1"])
        put("url", ke_record["url"])
        if viet_record:
            put("word", viet_record.get("word", ""))
            put("image1", viet_record.get("image1", ""))
            put("image2", viet_record.get("image2", ""))
        return values

    for record in ke_records:
        if not (
            record["combo4"]
            and record["url_class"] == "valid"
            and record["combo4"] not in dang_keys
        ):
            continue
        dang_metrics[record["domain_key"]]["ke_url_missing_dang"] += 1
        viet_match = viet_by_key.get(record["combo4"], [None])[0]
        recovery_rows.append(
            {
                "ke_row": record["row"],
                "viet_row": viet_match["row"] if viet_match else "",
                "domain": record["domain"],
                "title": normalize_spaces(record["title"]),
                "h1": normalize_spaces(record["h1"]),
                "keyword": normalize_spaces(record["keyword"]),
                "url": normalize_spaces(record["url"]),
                "values": build_recovery_values(record, viet_match),
            }
        )
        _issue(
            issues,
            category="Cần khôi phục DANG_BAI",
            sheet=ke_table.name,
            row=record["row"],
            domain=record["domain"],
            title=record["title"],
            h1=record["h1"],
            keyword=record["keyword"],
            detail="KE_HOACH có URL hợp lệ nhưng combo 4 không còn trong DANG_BAI",
            level="recovery",
            target_sheet=dang_table.name,
        )

    for values in viet_metrics.values():
        values["classification_difference"] = values["total_rows"] - values["classified_total"]
    for values in dang_metrics.values():
        values["classified_total"] = (
            values["in_viet"] + values["dang_missing_viet"] + values["combo4_missing"]
        )
        values["classification_difference"] = values["total_rows"] - values["classified_total"]

    ke_summary = _summary_rows(ke_metrics, registry, _new_ke_metrics)
    viet_summary = _summary_rows(
        viet_metrics,
        registry,
        _new_viet_metrics,
        total_overrides={"duplicate_groups": len(viet_duplicate_groups)},
    )
    dang_summary = _summary_rows(dang_metrics, registry, _new_dang_metrics)

    all_domain_keys = set(ke_metrics) | set(viet_metrics) | set(dang_metrics)
    reconciliation_rows: list[dict[str, Any]] = []
    for domain_key in registry.ordered_keys(all_domain_keys):
        ke_values = ke_metrics[domain_key]
        viet_values = viet_metrics[domain_key]
        dang_values = dang_metrics[domain_key]
        classified_total = (
            viet_values["in_dang"]
            + viet_values["recovery_dang"]
            + viet_values["pending_dang"]
            + viet_values["combo4_missing"]
        )
        difference = viet_values["total_rows"] - classified_total
        has_mismatch = any(
            (
                ke_values["total_rows"] != viet_values["total_rows"],
                ke_values["missing_in_viet"] != 0,
                viet_values["missing_in_ke"] != 0,
                viet_values["combo4_missing"] != 0,
                dang_values["dang_missing_viet"] != 0,
                dang_values["combo4_missing"] != 0,
                difference != 0,
            )
        )
        reconciliation_rows.append(
            {
                "domain": registry.display(domain_key),
                "domain_key": domain_key,
                "ke_total": ke_values["total_rows"],
                "viet_total": viet_values["total_rows"],
                "ke_combo4": ke_values["combo4_complete"],
                "viet_combo4": viet_values["combo4_complete"],
                "dang_combo4": dang_values["combo4_complete"],
                "ke_missing_viet": ke_values["missing_in_viet"],
                "viet_missing_ke": viet_values["missing_in_ke"],
                "in_dang": viet_values["in_dang"],
                "recovery_dang": viet_values["recovery_dang"],
                "pending_dang": viet_values["pending_dang"],
                "ke_url_deleted": viet_values["recovery_dang"],
                "viet_missing_remaining": viet_values["pending_dang"],
                "viet_combo4_missing": viet_values["combo4_missing"],
                "classified_total": classified_total,
                "difference": difference,
                "dang_missing_viet": dang_values["dang_missing_viet"],
                "status": "LỆCH" if has_mismatch else "KHỚP",
            }
        )

    rec_sum_keys = [
        "ke_total", "viet_total", "ke_combo4", "viet_combo4", "dang_combo4",
        "ke_missing_viet", "viet_missing_ke", "in_dang", "recovery_dang",
        "pending_dang", "ke_url_deleted", "viet_missing_remaining",
        "viet_combo4_missing", "classified_total", "difference", "dang_missing_viet",
    ]
    rec_totals = {
        key: sum(int(row[key]) for row in reconciliation_rows) for key in rec_sum_keys
    }
    total_mismatch = any(
        (
            rec_totals["ke_total"] != rec_totals["viet_total"],
            rec_totals["ke_missing_viet"] != 0,
            rec_totals["viet_missing_ke"] != 0,
            rec_totals["viet_combo4_missing"] != 0,
            rec_totals["dang_missing_viet"] != 0,
            dang_summary["total"]["combo4_missing"] != 0,
            rec_totals["difference"] != 0,
        )
    )
    reconciliation_total = {
        "domain": "TỔNG TẤT CẢ",
        "domain_key": "__total__",
        **rec_totals,
        "status": "LỆCH" if total_mismatch else "KHỚP",
    }

    level_order = {"error": 0, "recovery": 1, "pending": 2, "info": 3}
    issues.sort(
        key=lambda item: (
            level_order.get(item.get("level", "error"), 9),
            item["category"],
            item["sheet"],
            item["row"],
        )
    )
    error_count = sum(item.get("level", "error") == "error" for item in issues)
    recovery_count = len(recovery_rows)
    pending_count = sum(item.get("level") == "pending" for item in issues)
    archived_count = sum(item.get("level") == "info" for item in issues)
    write_completed_ok = int(viet_summary["total"]["completed_ok"])
    dang_total_rows = int(dang_summary["total"]["total_rows"])
    write_dang_difference = write_completed_ok - dang_total_rows
    ke_url_valid = int(ke_summary["total"]["url_valid"])
    dang_posted = int(dang_summary["total"]["posted"])
    url_posted_difference = ke_url_valid - dang_posted
    comparison_details: list[str] = []
    if write_dang_difference > 0:
        comparison_details.append(
            f"Hoàn tất OK {write_completed_ok} ↔ Tổng ĐĂNG_BÀI {dang_total_rows}: "
            f"ĐĂNG_BÀI thiếu {write_dang_difference} dòng"
        )
    elif write_dang_difference < 0:
        comparison_details.append(
            f"Hoàn tất OK {write_completed_ok} ↔ Tổng ĐĂNG_BÀI {dang_total_rows}: "
            f"ĐĂNG_BÀI dư {abs(write_dang_difference)} dòng"
        )
    if url_posted_difference > 0:
        comparison_details.append(
            f"URL hợp lệ {ke_url_valid} ↔ Đã đăng {dang_posted}: "
            f"ĐĂNG_BÀI thiếu {url_posted_difference} bài đã đăng"
        )
    elif url_posted_difference < 0:
        comparison_details.append(
            f"URL hợp lệ {ke_url_valid} ↔ Đã đăng {dang_posted}: "
            f"ĐĂNG_BÀI dư {abs(url_posted_difference)} bài đã đăng"
        )
    if comparison_details:
        if write_dang_difference and url_posted_difference:
            overall_status = "LỆCH 2 CẶP ĐỐI SOÁT"
        elif write_dang_difference:
            overall_status = "LỆCH VIẾT_BÀI ↔ ĐĂNG_BÀI"
        else:
            overall_status = "LỆCH URL HỢP LỆ ↔ ĐÃ ĐĂNG"
        overall_level = "error"
        overall_detail = " • ".join(comparison_details)
    elif error_count:
        overall_status = "CẦN KIỂM TRA"
        overall_level = "error"
        overall_detail = f"Có {error_count} lỗi dữ liệu cần mở danh sách để xử lý"
    elif recovery_count:
        overall_status = "CẦN KHÔI PHỤC"
        overall_level = "recovery"
        overall_detail = f"Dữ liệu đối chiếu khớp, có {recovery_count} bài cần thêm lại vào DANG_BAI"
    elif pending_count:
        overall_status = "DỮ LIỆU KHỚP"
        overall_level = "pending"
        overall_detail = f"Không có lỗi; còn {pending_count} bài trong tiến độ chưa chuyển sang DANG_BAI"
    else:
        overall_status = "ỔN"
        overall_level = "ok"
        overall_detail = "Không phát hiện lỗi dữ liệu hoặc chênh lệch đối soát"
    analyzed_at = datetime.now().astimezone().isoformat(timespec="seconds")
    fingerprint = file_fingerprint(workbook.path, include_hash=True)
    return {
        "app_version": APP_VERSION,
        "source_path": str(workbook.path),
        "source_name": workbook.path.name,
        "source_fingerprint": fingerprint,
        "analyzed_at": analyzed_at,
        "sheet_names": workbook.sheet_names,
        "publish_review": build_publish_review_from_tables(dang_table, viet_table),
        "resolved_sheets": {
            "ke_hoach": ke_table.name,
            "viet_bai": viet_table.name,
            "dang_bai": dang_table.name,
        },
        "sheet_info": {
            "ke_hoach": {
                "name": ke_table.name,
                "xml_path": ke_table.xml_path,
                "header_row": ke_table.header_row,
                "status_column": ke_cols["source_status"],
                "status_column_header": "Trạng thái nguồn",
                "max_column": ke_table.max_column,
            }
        },
        "summaries": {
            "ke_hoach": ke_summary,
            "viet_bai": viet_summary,
            "dang_bai": dang_summary,
            "reconciliation": {
                "rows": reconciliation_rows,
                "total": reconciliation_total,
            },
        },
        "overall": {
            "status": overall_status,
            "level": overall_level,
            "detail": overall_detail,
            "error_count": error_count,
            "recovery_count": recovery_count,
            "pending_count": pending_count,
            "archived_count": archived_count,
            "write_completed_ok": write_completed_ok,
            "dang_total_rows": dang_total_rows,
            "write_dang_difference": write_dang_difference,
            "ke_url_valid": ke_url_valid,
            "dang_posted": dang_posted,
            "url_posted_difference": url_posted_difference,
        },
        "recovery": {
            "headers": dang_headers,
            "rows": recovery_rows,
        },
        "status_updates": status_updates,
        "details": issues,
        "issues": issues,
        "issue_count": error_count,
    }
