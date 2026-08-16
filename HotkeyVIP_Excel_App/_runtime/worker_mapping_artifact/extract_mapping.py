import json
import sqlite3
from pathlib import Path
from urllib.parse import urlsplit, urlunsplit

import openpyxl

SOURCE = Path(r"D:\CodexProjects\Hotkeyvip\04_excel\hotkeyvip_test.xlsm")
OUTPUT = Path(r"D:\HotkeyVIP_Excel_App\_runtime\worker_mapping_artifact\mapping.json")
PROFILE_ROOT = Path(r"D:\CodexProjects\Hotkeyvip\02_viet_bai\du_lieu_3_workers\profiles")


def normalize_url(value):
    if not value:
        return ""
    raw = str(value).strip()
    try:
        parts = urlsplit(raw)
        path = parts.path.rstrip("/") or "/"
        return urlunsplit((parts.scheme.lower(), parts.netloc.lower(), path, "", ""))
    except Exception:
        return raw.rstrip("/")


def history_urls(worker):
    db = PROFILE_ROOT / worker / "Default" / "History"
    connection = sqlite3.connect(f"file:{db.as_posix()}?mode=ro", uri=True)
    try:
        return {normalize_url(row[0]) for row in connection.execute("SELECT url FROM urls") if row[0]}
    finally:
        connection.close()


book = openpyxl.load_workbook(SOURCE, read_only=True, data_only=True)
sheet = book["VIET_BAI"]
headers = {str(cell.value).strip(): cell.column for cell in sheet[1] if cell.value is not None}

aliases = {
    "domain": ["Tên Miền", "Tên miền"],
    "seo_title": ["Tiêu đề SEO"],
    "h1": ["H1"],
    "keyword": ["Từ khóa"],
    "url": ["URL ChatGPT", "URL chatgpt"],
    "status": ["Trạng thái viết"],
    "error": ["Lỗi viết"],
    "word_path": ["Đường dẫn Word"],
    "image_1": ["Đường dẫn ảnh 1"],
    "image_2": ["Đường dẫn ảnh 2"],
}


def find_col(names):
    for name in names:
        if name in headers:
            return headers[name]
    raise KeyError(f"Không tìm thấy cột {names}; các cột hiện có: {list(headers)}")


cols = {key: find_col(names) for key, names in aliases.items()}
histories = {worker: history_urls(worker) for worker in ("worker_1", "worker_3")}
records = []

for row_idx, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
    def value(key):
        return values[cols[key] - 1]

    status = value("status")
    if str(status or "").strip() != "WORD_ERROR":
        continue
    url = value("url")
    normalized = normalize_url(url)
    matches = [worker for worker, urls in histories.items() if normalized in urls]
    if len(matches) != 1:
        raise RuntimeError(f"Dòng {row_idx}: URL phải khớp đúng 1 worker, thực tế={matches}, URL={url}")
    image_1 = value("image_1")
    image_2 = value("image_2")
    records.append({
        "stt": len(records) + 1,
        "excel_row": row_idx,
        "domain": value("domain"),
        "seo_title": value("seo_title"),
        "h1": value("h1"),
        "keyword": value("keyword"),
        "url": url,
        "worker": matches[0],
        "status": status,
        "error": value("error"),
        "word_path": value("word_path"),
        "image_1_exists": bool(image_1),
        "image_2_exists": bool(image_2),
    })

payload = {
    "source": str(SOURCE),
    "sheet": sheet.title,
    "total": len(records),
    "worker_1": sum(r["worker"] == "worker_1" for r in records),
    "worker_3": sum(r["worker"] == "worker_3" for r in records),
    "records": records,
}
OUTPUT.parent.mkdir(parents=True, exist_ok=True)
OUTPUT.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
print(json.dumps({k: payload[k] for k in ("total", "worker_1", "worker_3")}, ensure_ascii=False))
