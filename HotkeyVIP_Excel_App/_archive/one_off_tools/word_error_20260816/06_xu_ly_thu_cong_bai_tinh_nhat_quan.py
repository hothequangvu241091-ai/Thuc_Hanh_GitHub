# -*- coding: utf-8 -*-
"""Xử lý riêng bài 'tính nhất quán phương pháp luận' bị hỏi nhầm Brief hai lần."""

import importlib.util
import shutil
import sys
import time
from datetime import datetime
from pathlib import Path

import openpyxl
import pythoncom
import xlwings as xw

APP_ROOT = Path(__file__).resolve().parents[3]
FLOW_PATH = APP_ROOT / "app_flows" / "03_viet_bai_tao_anh.py"
SOURCE = Path(r"D:\CodexProjects\Hotkeyvip\04_excel\hotkeyvip_test.xlsm")
SHEET_NAME = "VIET_BAI"
TARGET_KEYWORD = "tính nhất quán phương pháp luận"

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace", line_buffering=True)


def load_flow():
    sys.path.insert(0, str(FLOW_PATH.parent))
    spec = importlib.util.spec_from_file_location("manual_special_article_flow", FLOW_PATH)
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def read_target():
    book = openpyxl.load_workbook(SOURCE, read_only=True, data_only=True)
    sheet = book[SHEET_NAME]
    headers = {str(cell.value).strip(): cell.column for cell in sheet[1] if cell.value}
    for row_index, values in enumerate(sheet.iter_rows(min_row=2, values_only=True), start=2):
        keyword = str(values[headers["Từ khóa"] - 1] or "").strip()
        if keyword == TARGET_KEYWORD:
            result = {
                "row": row_index,
                "keyword": keyword,
                "prompt": str(values[headers["Prompt viết bài"] - 1] or ""),
                "gpt_url": str(values[headers["URL GPT gốc"] - 1] or "").strip(),
                "word_path": str(values[headers["Đường dẫn Word"] - 1] or "").strip(),
            }
            book.close()
            return result
    book.close()
    raise RuntimeError(f"Không tìm thấy bài: {TARGET_KEYWORD}")


def update_excel(task, chat_url, word_count):
    backup = SOURCE.with_name(
        f"{SOURCE.stem}_backup_truoc_bai_thu_cong_{datetime.now():%Y%m%d_%H%M%S}{SOURCE.suffix}"
    )
    shutil.copy2(SOURCE, backup)
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    book = None
    try:
        book = app.books.open(str(SOURCE), update_links=False, read_only=False)
        sheet = book.sheets[SHEET_NAME]
        last_col = sheet.used_range.last_cell.column
        header_values = sheet.range((1, 1), (1, last_col)).value
        headers = {
            str(value or "").strip(): index + 1
            for index, value in enumerate(header_values)
            if str(value or "").strip()
        }
        target_row = None
        last_row = sheet.used_range.last_cell.row
        keyword_col = headers["Từ khóa"]
        for row in range(2, last_row + 1):
            if str(sheet.cells(row, keyword_col).value or "").strip() == TARGET_KEYWORD:
                target_row = row
                break
        if target_row is None:
            raise RuntimeError("Không tìm thấy đúng từ khóa khi cập nhật Excel")
        sheet.cells(target_row, headers["Đường dẫn Word"]).value = task["word_path"]
        sheet.cells(target_row, headers["URL ChatGPT"]).value = chat_url
        sheet.cells(target_row, headers["Trạng thái viết"]).value = "OK"
        sheet.cells(target_row, headers["Lỗi viết"]).value = ""
        sheet.cells(target_row, headers["Số từ Word"]).value = int(word_count)
        book.save()
        print(f"Đã cập nhật Excel dòng hiện tại {target_row} sang OK.")
        print(f"Backup: {backup}")
    finally:
        if book is not None:
            book.close()
        app.quit()


def main():
    task = read_target()
    flow = load_flow()
    driver = None
    pythoncom.CoInitialize()
    try:
        print(f"Xử lý thủ công: {task['keyword']} | dòng hiện tại {task['row']}")
        flow.run_word_preflight("bài thủ công tính nhất quán")
        flow._THREAD_CONTEXT.worker_id = 3
        driver, wait = flow.create_shared_driver()
        driver.get(task["gpt_url"])
        flow.wait_chatgpt_page_ready(driver, wait, timeout=45)
        flow.send_once_unless_present(
            driver, wait, task["prompt"], flow.send_prompt_by_js, "GỬI PROMPT VIẾT BÀI THỦ CÔNG"
        )
        content = flow.get_gpt_content_after_wait(
            driver, flow.ARTICLE_WAIT_SECONDS, "Bài thủ công tính nhất quán"
        )
        if not content:
            raise RuntimeError("ChatGPT không trả về nội dung bài viết")
        snapshot = flow._THREAD_CONTEXT.last_stable_article
        word_count = flow.count_words(content)
        if word_count < flow.MIN_WORDS:
            raise RuntimeError(
                f"Bài chỉ có {word_count} từ, dưới mức yêu cầu {flow.MIN_WORDS}; không lưu"
            )
        if TARGET_KEYWORD.casefold() not in content.casefold():
            raise RuntimeError("Bài không chứa từ khóa mục tiêu; không lưu")
        if "===brief_1===" in content.casefold() or "===brief_2===" in content.casefold():
            raise RuntimeError("Phản hồi vẫn là Brief; không lưu")

        for attempt in (1, 2):
            try:
                flow.copy_and_save_snapshot(snapshot, task["word_path"])
                break
            except Exception as exc:
                if attempt == 1 and (
                    isinstance(exc, flow.WordSystemError) or flow.is_word_system_error(exc)
                ):
                    if flow.recover_word_runtime_once():
                        continue
                raise
        if not flow.is_word_ok(task["word_path"]):
            raise RuntimeError("File Word tạo xong nhưng đọc kiểm tra không đạt")
        chat_url = driver.current_url
        update_excel(task, chat_url, word_count)
        print(f"HOÀN TẤT: {word_count} từ")
        print(f"Word: {task['word_path']}")
        print(f"Chat mới: {chat_url}")
    finally:
        if driver is not None:
            try:
                driver.quit()
            except Exception:
                pass
        flow.cleanup_all_owned_word_processes()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
