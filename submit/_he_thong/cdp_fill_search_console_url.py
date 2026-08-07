import argparse
import json
import time

import requests
import websocket


class Cdp:
    def __init__(self, socket_url: str):
        self.socket = websocket.create_connection(
            socket_url,
            timeout=5,
            suppress_origin=True,
        )
        self.message_id = 0

    def call(self, method: str, params=None):
        self.message_id += 1
        request_id = self.message_id
        self.socket.send(
            json.dumps(
                {
                    "id": request_id,
                    "method": method,
                    "params": params or {},
                }
            )
        )
        while True:
            response = json.loads(self.socket.recv())
            if response.get("id") == request_id:
                if "error" in response:
                    raise RuntimeError(response["error"])
                return response.get("result", {})

    def close(self):
        self.socket.close()


def find_search_console_page(port: int, deadline: float):
    endpoint = f"http://127.0.0.1:{port}/json/list"
    while time.time() < deadline:
        try:
            pages = requests.get(endpoint, timeout=1).json()
            matches = [
                page
                for page in pages
                if page.get("type") == "page"
                and "search.google.com" in page.get("url", "")
            ]
            if matches:
                return matches[0]
        except (requests.RequestException, ValueError):
            pass
        time.sleep(0.4)
    return None


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--port", required=True, type=int)
    parser.add_argument("--text", required=True)
    parser.add_argument("--timeout", type=float, default=35)
    parser.add_argument("--click-request-indexing", action="store_true")
    parser.add_argument("--request-timeout", type=float, default=45)
    args = parser.parse_args()

    deadline = time.time() + args.timeout
    page = find_search_console_page(args.port, deadline)
    if not page:
        print("Không tìm thấy trang Search Console qua CDP.")
        return 2

    cdp = Cdp(page["webSocketDebuggerUrl"])
    try:
        cdp.call("Runtime.enable")

        selector_script = r"""
(() => {
  const selector =
    'input[role="combobox"][aria-label^="Kiểm tra mọi URL trong"]:not([disabled])';
  const input = document.querySelector(selector);
  if (!input) return false;
  input.click();
  input.focus();
  const setter =
    Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value').set;
  setter.call(input, '');
  input.dispatchEvent(new Event('input', { bubbles: true }));
  return document.activeElement === input;
})()
"""

        focused = False
        while time.time() < deadline:
            result = cdp.call(
                "Runtime.evaluate",
                {
                    "expression": selector_script,
                    "returnByValue": True,
                },
            )
            focused = bool(result.get("result", {}).get("value"))
            if focused:
                break
            time.sleep(0.4)

        if not focused:
            print("Không tìm thấy ô URL đang hoạt động trong DOM.")
            return 3

        cdp.call("Input.insertText", {"text": args.text})
        cdp.call(
            "Input.dispatchKeyEvent",
            {
                "type": "keyDown",
                "key": "Enter",
                "code": "Enter",
                "windowsVirtualKeyCode": 13,
                "nativeVirtualKeyCode": 13,
            },
        )
        cdp.call(
            "Input.dispatchKeyEvent",
            {
                "type": "keyUp",
                "key": "Enter",
                "code": "Enter",
                "windowsVirtualKeyCode": 13,
                "nativeVirtualKeyCode": 13,
            },
        )

        if not args.click_request_indexing:
            print("DOM_CDP_FILL_AND_ENTER_OK")
            return 0

        request_deadline = time.time() + args.request_timeout
        click_script = r"""
(() => {
  const candidates = Array.from(
    document.querySelectorAll('[role="button"], button')
  );
  const button = candidates.find((element) => {
    const label = (element.getAttribute('aria-label') || '').trim();
    const text = (element.textContent || '').trim();
    return (
      label.startsWith('Yêu cầu lập chỉ mục') ||
      text.startsWith('Yêu cầu lập chỉ mục')
    );
  });
  if (!button) return { clicked: false, reason: 'not-found' };
  if (
    button.disabled ||
    button.getAttribute('aria-disabled') === 'true'
  ) {
    return { clicked: false, reason: 'disabled' };
  }
  button.click();
  return {
    clicked: true,
    label: button.getAttribute('aria-label') || button.textContent
  };
})()
"""

        while time.time() < request_deadline:
            try:
                result = cdp.call(
                    "Runtime.evaluate",
                    {
                        "expression": click_script,
                        "returnByValue": True,
                    },
                )
                value = result.get("result", {}).get("value") or {}
                if value.get("clicked"):
                    print("REQUEST_INDEXING_CLICKED")
                    return 0
            except RuntimeError:
                pass
            time.sleep(0.5)

        print("REQUEST_INDEXING_BUTTON_NOT_FOUND")
        return 4
    finally:
        cdp.close()


if __name__ == "__main__":
    raise SystemExit(main())
