from __future__ import annotations

import json
import os
import shutil
import tempfile
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any
from urllib.parse import urlsplit

from .excel_io import OpenXmlWorkbook, normalize_spaces, normalize_text


DEFAULT_SUBMIT_SYSTEM_DIR = Path(
    r"D:\CodexProjects\Hotkeyvip\06_du_lieu_chay\submit_edge_profiles\_he_thong"
)
ALIASES = {
    "status": ["Trạng thái đăng"],
    "url": ["URL đã đăng", "URL"],
    "published_at": ["Thời gian đăng"],
    "domain": ["Tên miền", "Tên Miền"],
    "title": ["Tiêu đề", "Từ khóa", "Main Keyword"],
}


def submit_system_dir() -> Path:
    override = os.environ.get("HOTKEYVIP_SUBMIT_SYSTEM_DIR", "").strip()
    return Path(override).resolve() if override else DEFAULT_SUBMIT_SYSTEM_DIR


def submit_launcher_path() -> Path:
    return submit_system_dir().parent / "QUAN_LY_PROFILE_SUBMIT.bat"


def _parse_excel_date(value: Any) -> date | None:
    if value in (None, ""):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return (datetime(1899, 12, 30) + timedelta(days=float(value))).date()
        except (OverflowError, ValueError):
            return None

    text = normalize_spaces(value)
    if not text:
        return None
    try:
        return (datetime(1899, 12, 30) + timedelta(days=float(text))).date()
    except (OverflowError, ValueError):
        pass
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date()
    except ValueError:
        pass
    for pattern in (
        "%d/%m/%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            continue
    return None


def _public_url(value: Any) -> str:
    url = normalize_spaces(value)
    if not url:
        return ""
    try:
        parsed = urlsplit(url)
    except ValueError:
        return ""
    if parsed.scheme.casefold() not in {"http", "https"} or not parsed.hostname:
        return ""
    path = parsed.path.casefold()
    if path == "/admin" or path.startswith("/admin/"):
        return ""
    return url


def _url_key(value: Any) -> str:
    return normalize_spaces(value).rstrip("/").casefold()


def _load_history(path: Path) -> dict[str, Any]:
    if not path.exists():
        return {"version": 1, "urls": []}
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Không đọc được lịch sử Submit: {path}") from exc
    if not isinstance(data, dict) or not isinstance(data.get("urls"), list):
        raise RuntimeError(f"File lịch sử Submit không đúng cấu trúc: {path}")
    return data


def _assert_submit_not_running(system_dir: Path) -> None:
    progress_path = system_dir / "auto_submit_progress.json"
    if not progress_path.exists():
        return
    try:
        progress = json.loads(progress_path.read_text(encoding="utf-8-sig"))
    except (OSError, json.JSONDecodeError):
        return
    state = normalize_spaces(progress.get("state", "")).upper()
    if state in {"STARTING", "RUNNING", "STOPPING"}:
        raise RuntimeError(
            "App Submit đang chạy tự động. Hãy chờ phiên Submit kết thúc rồi chuyển URL."
        )


def inspect_latest_published_urls(path: str | Path) -> dict[str, Any]:
    source = Path(path).resolve()
    workbook = OpenXmlWorkbook(source)
    sheet = workbook.find_sheet("DANG_BAI")
    if sheet is None:
        raise RuntimeError("Workbook không có sheet DANG_BAI.")
    table = workbook.read_sheet(sheet)
    columns = {
        name: table.find_column(aliases, required=True)
        for name, aliases in ALIASES.items()
    }

    published_rows: list[dict[str, Any]] = []
    for row_number, row in table.rows:
        status = normalize_spaces(table.value(row, columns["status"]))
        if "đã đăng" not in normalize_text(status):
            continue
        published_date = _parse_excel_date(table.value(row, columns["published_at"]))
        if published_date is None:
            continue
        published_rows.append(
            {
                "row": row_number,
                "date": published_date,
                "url": _public_url(table.value(row, columns["url"])),
                "domain": normalize_spaces(table.value(row, columns["domain"])),
                "title": normalize_spaces(table.value(row, columns["title"])),
            }
        )

    if not published_rows:
        return {
            "source_path": str(source),
            "latest_date": None,
            "published_total": 0,
            "valid_total": 0,
            "missing_url_total": 0,
            "duplicate_total": 0,
            "history_total": 0,
            "new_total": 0,
            "new_rows": [],
            "history_path": str(submit_system_dir() / "submit_url_history.json"),
        }

    latest_date = max(item["date"] for item in published_rows)
    latest_rows = [item for item in published_rows if item["date"] == latest_date]
    valid_rows: list[dict[str, Any]] = []
    seen_in_excel: set[str] = set()
    duplicate_inside_excel = 0
    for item in latest_rows:
        key = _url_key(item["url"])
        if not key:
            continue
        if key in seen_in_excel:
            duplicate_inside_excel += 1
            continue
        seen_in_excel.add(key)
        valid_rows.append(item)

    system_dir = submit_system_dir()
    history_path = system_dir / "submit_url_history.json"
    history = _load_history(history_path)
    return {
        "source_path": str(source),
        "latest_date": latest_date.isoformat(),
        "published_total": len(latest_rows),
        "valid_total": len(valid_rows),
        "missing_url_total": sum(1 for item in latest_rows if not item["url"]),
        "duplicate_total": duplicate_inside_excel,
        "history_total": len(history["urls"]),
        "new_total": len(valid_rows),
        "new_rows": valid_rows,
        "history_path": str(history_path),
    }


def transfer_latest_published_urls(path: str | Path) -> dict[str, Any]:
    snapshot = inspect_latest_published_urls(path)
    if not snapshot["latest_date"] or not snapshot["new_rows"]:
        return snapshot

    system_dir = submit_system_dir()
    if not system_dir.is_dir():
        raise RuntimeError(f"Không tìm thấy hệ thống Submit: {system_dir}")
    _assert_submit_not_running(system_dir)
    history_path = system_dir / "submit_url_history.json"
    history = _load_history(history_path)
    removed_total = len(history["urls"])
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    replacement: list[dict[str, Any]] = []
    for item in snapshot["new_rows"]:
        replacement.append(
            {
                "url": item["url"],
                "status": "PENDING",
                "message": (
                    "Nhập từ app Excel; ngày đăng mới nhất "
                    f"{snapshot['latest_date']}; dòng {item['row']}."
                ),
                "updatedAt": now,
                "sourceExcel": snapshot["source_path"],
                "sourceRow": int(item["row"]),
                "publishedDate": snapshot["latest_date"],
            }
        )

    backup_path = history_path.with_name(history_path.name + ".excel_app.bak")
    if history_path.exists():
        shutil.copy2(history_path, backup_path)
    history["urls"] = replacement
    temporary_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(
            mode="w",
            encoding="utf-8",
            dir=history_path.parent,
            prefix=history_path.name + ".",
            suffix=".tmp",
            delete=False,
        ) as handle:
            json.dump(history, handle, ensure_ascii=False, indent=2)
            temporary_path = Path(handle.name)
        os.replace(temporary_path, history_path)
    finally:
        if temporary_path is not None and temporary_path.exists():
            temporary_path.unlink(missing_ok=True)

    snapshot["added_total"] = len(replacement)
    snapshot["removed_total"] = removed_total
    snapshot["backup_path"] = str(backup_path)
    return snapshot
