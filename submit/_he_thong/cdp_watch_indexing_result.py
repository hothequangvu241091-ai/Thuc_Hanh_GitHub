import argparse
import json
import time

from cdp_fill_search_console_url import Cdp, find_search_console_page


def classify(text: str):
    lowered = text.casefold()
    if (
        "đã yêu cầu lập chỉ mục" in lowered
        or "url đã được thêm vào hàng đợi ưu tiên" in lowered
    ):
        return "SUCCESS"
    if "vượt quá hạn mức" in lowered or "quota exceeded" in lowered:
        return "QUOTA"
    return None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", required=True, type=int)
    parser.add_argument("--timeout", type=float, default=55)
    args = parser.parse_args()

    deadline = time.time() + args.timeout
    page = find_search_console_page(args.port, deadline)
    if not page:
        print(json.dumps({"status": "NO_PAGE"}))
        return 2

    cdp = Cdp(page["webSocketDebuggerUrl"])
    last_text = ""
    try:
        cdp.call("Runtime.enable")
        script = r"""
(() => {
  const selectors = [
    '[role="dialog"]',
    '[role="alertdialog"]',
    '[aria-live="assertive"]',
    '[aria-live="polite"]'
  ];
  const visible = (element) => {
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return (
      style.display !== 'none' &&
      style.visibility !== 'hidden' &&
      rect.width > 0 &&
      rect.height > 0
    );
  };
  return Array.from(document.querySelectorAll(selectors.join(',')))
    .filter(visible)
    .map((element) => (element.innerText || '').trim())
    .filter(Boolean)
    .filter((text, index, all) => all.indexOf(text) === index);
})()
"""

        while time.time() < deadline:
            try:
                result = cdp.call(
                    "Runtime.evaluate",
                    {"expression": script, "returnByValue": True},
                )
                texts = result.get("result", {}).get("value") or []
                for text in texts:
                    last_text = text
                    status = classify(text)
                    if status:
                        print(
                            json.dumps(
                                {"status": status, "text": text},
                                ensure_ascii=True,
                            )
                        )
                        return 0
            except RuntimeError:
                pass
            time.sleep(0.5)

        print(
            json.dumps(
                {"status": "TIMEOUT", "last_text": last_text},
                ensure_ascii=True,
            )
        )
        return 3
    finally:
        cdp.close()


if __name__ == "__main__":
    raise SystemExit(main())
