# -*- coding: utf-8 -*-
"""Tạo Word bài tính nhất quán từ pasted-text người dùng cung cấp."""

import html
import importlib.util
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

import pythoncom
import xlwings as xw

APP_ROOT = Path(__file__).resolve().parents[3]
FLOW_PATH = APP_ROOT / "app_flows" / "03_viet_bai_tao_anh.py"
SOURCE = Path(r"D:\CodexProjects\Hotkeyvip\04_excel\hotkeyvip_test.xlsm")
INPUT = Path(r"C:\Users\Admin\.codex\attachments\c237df59-6c7f-4507-a0a3-4fd7dd4b76aa\pasted-text.txt")
TARGET_KEYWORD = "tính nhất quán phương pháp luận"
WORD_PATH = Path(r"D:\CodexProjects\Hotkeyvip\07_ket_qua\bai_viet\bantinkhoahoc.com\tính nhất quán phương pháp luận.docx")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)


def load_flow():
    sys.path.insert(0, str(FLOW_PATH.parent))
    spec = importlib.util.spec_from_file_location("pasted_article_flow", FLOW_PATH)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def inline_markup(text):
    escaped = html.escape(text.strip())
    escaped = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", escaped)
    escaped = re.sub(r"(?<!\*)\*([^*]+?)\*(?!\*)", r"<em>\1</em>", escaped)
    return escaped


def markdown_to_html(markdown_text):
    lines = markdown_text.replace("\r\n", "\n").split("\n")
    output = []
    paragraph = []
    index = 0

    def flush_paragraph():
        if paragraph:
            joined = " ".join(part.strip() for part in paragraph if part.strip())
            if joined:
                output.append(f"<p>{inline_markup(joined)}</p>")
            paragraph.clear()

    while index < len(lines):
        raw = lines[index]
        stripped = raw.strip()
        if not stripped:
            flush_paragraph()
            index += 1
            continue
        if stripped == "---":
            flush_paragraph()
            index += 1
            continue
        if stripped.startswith("### "):
            flush_paragraph()
            output.append(f"<h3>{inline_markup(stripped[4:])}</h3>")
            index += 1
            continue
        if stripped.startswith("## "):
            flush_paragraph()
            output.append(f"<h2>{inline_markup(stripped[3:])}</h2>")
            index += 1
            continue
        if stripped.startswith("# "):
            flush_paragraph()
            output.append(f"<h1>{inline_markup(stripped[2:])}</h1>")
            index += 1
            continue
        if "|" in stripped and index + 1 < len(lines):
            separator = lines[index + 1].strip()
            if "|" in separator and re.fullmatch(r"[\s|:\-]+", separator):
                flush_paragraph()
                headers = [cell.strip() for cell in stripped.strip("|").split("|")]
                index += 2
                rows = []
                while index < len(lines) and "|" in lines[index]:
                    rows.append([cell.strip() for cell in lines[index].strip().strip("|").split("|")])
                    index += 1
                output.append("<table><thead><tr>")
                output.extend(f"<th>{inline_markup(cell)}</th>" for cell in headers)
                output.append("</tr></thead><tbody>")
                for row in rows:
                    output.append("<tr>")
                    output.extend(f"<td>{inline_markup(cell)}</td>" for cell in row)
                    output.append("</tr>")
                output.append("</tbody></table>")
                continue
        if stripped.startswith(("- ", "* ")):
            flush_paragraph()
            items = []
            while index < len(lines) and lines[index].strip().startswith(("- ", "* ")):
                items.append(lines[index].strip()[2:].strip())
                index += 1
            output.append("<ul>")
            output.extend(f"<li>{inline_markup(item)}</li>" for item in items)
            output.append("</ul>")
            continue
        paragraph.append(stripped)
        index += 1

    flush_paragraph()
    return "\n".join(output)


def update_excel(flow, word_count):
    backup = SOURCE.with_name(
        f"{SOURCE.stem}_backup_truoc_noi_dung_dan_{datetime.now():%Y%m%d_%H%M%S}{SOURCE.suffix}"
    )
    shutil.copy2(SOURCE, backup)
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    book = None
    try:
        book = app.books.open(str(SOURCE), update_links=False, read_only=False)
        sheet = book.sheets["VIET_BAI"]
        last_col = sheet.used_range.last_cell.column
        values = sheet.range((1, 1), (1, last_col)).value
        headers = {
            str(value or "").strip(): index + 1
            for index, value in enumerate(values)
            if str(value or "").strip()
        }
        target_row = None
        for row in range(2, sheet.used_range.last_cell.row + 1):
            if str(sheet.cells(row, headers["Từ khóa"]).value or "").strip() == TARGET_KEYWORD:
                target_row = row
                break
        if target_row is None:
            raise RuntimeError("Không tìm thấy dòng từ khóa để cập nhật")
        sheet.cells(target_row, headers["Đường dẫn Word"]).value = str(WORD_PATH)
        sheet.cells(target_row, headers["Trạng thái viết"]).value = "OK"
        sheet.cells(target_row, headers["Lỗi viết"]).value = ""
        sheet.cells(target_row, headers["Số từ Word"]).value = int(word_count)
        book.save()
        print(f"Excel dòng {target_row}: OK")
        print(f"Backup: {backup}")
    finally:
        if book is not None:
            book.close()
        app.quit()


def main():
    flow = load_flow()
    markdown_text = INPUT.read_text(encoding="utf-8")
    article_html = markdown_to_html(markdown_text)
    snapshot = {"html": article_html}
    snapshot["text"] = flow.get_word_source_text(snapshot)
    word_count = flow.count_words(snapshot["text"])
    if "--update-only" in sys.argv:
        if not flow.is_word_ok(str(WORD_PATH)):
            raise RuntimeError("Không cập nhật Excel vì file Word chưa hợp lệ")
        update_excel(flow, word_count)
        print(f"CHỈ CẬP NHẬT EXCEL: {word_count} từ")
        return
    if word_count < 80:
        raise RuntimeError(f"Nội dung chỉ có {word_count} từ")
    if TARGET_KEYWORD.casefold() not in snapshot["text"].casefold():
        raise RuntimeError("Nội dung không chứa từ khóa mục tiêu")

    pythoncom.CoInitialize()
    try:
        flow.run_word_preflight("nội dung người dùng cung cấp")
        for attempt in (1, 2):
            try:
                flow.copy_and_save_snapshot(snapshot, str(WORD_PATH))
                break
            except Exception as exc:
                if attempt == 1 and (
                    isinstance(exc, flow.WordSystemError) or flow.is_word_system_error(exc)
                ):
                    if flow.recover_word_runtime_once():
                        continue
                raise
        if not flow.is_word_ok(str(WORD_PATH)):
            raise RuntimeError("File Word tạo xong nhưng không đạt kiểm tra")
        update_excel(flow, word_count)
        print(f"HOÀN TẤT: {word_count} từ")
        print(f"Word: {WORD_PATH}")
    finally:
        flow.cleanup_all_owned_word_processes()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
