# -*- coding: utf-8 -*-
import importlib.util
import json
import sys
import time
from pathlib import Path

ROOT = Path(r"D:\HotkeyVIP_Excel_App")
V1 = Path(__file__).with_name("04_khoi_phuc_word_error.py")
URL = "https://chatgpt.com/g/g-6a27de6c03b48191abf9a9902635d9c5-article-detail-chi-tiet-bai-viet/c/6a80e702-acf8-83ec-b6a9-6d622c070f9b"
OUT = Path(__file__).with_name("manual_row3")

spec = importlib.util.spec_from_file_location("manual_v1", V1)
v1 = importlib.util.module_from_spec(spec)
sys.modules[spec.name] = v1
spec.loader.exec_module(v1)
flow = v1.load_flow_module()
flow._THREAD_CONTEXT.worker_id = 3
OUT.mkdir(parents=True, exist_ok=True)

driver = None
try:
    driver, _ = flow.create_shared_driver()
    driver.get(URL)
    time.sleep(8)
    data = driver.execute_script(r"""
        const selectors = [
          '[data-message-author-role]',
          '.markdown',
          'article',
          '[contenteditable="true"]',
          '.ProseMirror',
          '[data-testid*="canvas"]',
          '[class*="canvas"]',
          '[class*="artifact"]',
          'main'
        ];
        const result = {};
        for (const selector of selectors) {
          result[selector] = Array.from(document.querySelectorAll(selector)).map((node, index) => ({
            index,
            tag: node.tagName,
            role: node.getAttribute('data-message-author-role') || '',
            testid: node.getAttribute('data-testid') || '',
            cls: String(node.className || '').slice(0, 300),
            text: (node.innerText || node.textContent || '').trim().slice(0, 30000),
            html: (node.innerHTML || '').slice(0, 60000),
          }));
        }
        result.title = document.title;
        result.url = location.href;
        result.bodyText = (document.body.innerText || '').slice(0, 100000);
        return result;
    """)
    (OUT / "dom.json").write_text(
        json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    driver.save_screenshot(str(OUT / "page.png"))
    print(json.dumps({
        "title": data.get("title"),
        "roles": len(data.get("[data-message-author-role]", [])),
        "markdown": len(data.get(".markdown", [])),
        "articles": len(data.get("article", [])),
        "prosemirror": len(data.get(".ProseMirror", [])),
        "canvas": len(data.get('[data-testid*="canvas"]', [])),
        "artifact": len(data.get('[class*="artifact"]', [])),
        "body_chars": len(data.get("bodyText", "")),
    }, ensure_ascii=False))
finally:
    if driver is not None:
        driver.quit()
