import argparse
import base64
import json
import os
import shutil
import subprocess
import sys
import time
import uuid
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from urllib.parse import quote, urlsplit

import requests

from cdp_fill_search_console_url import Cdp


SYSTEM_DIR = Path(__file__).resolve().parent
ROOT_DIR = SYSTEM_DIR.parent
CONFIG_PATH = SYSTEM_DIR / "submit_edge_profiles.json"
HISTORY_PATH = SYSTEM_DIR / "submit_url_history.json"
PROGRESS_PATH = SYSTEM_DIR / "auto_submit_progress.json"
STOP_FLAG_PATH = SYSTEM_DIR / "stop_auto_submit.flag"
LOGS_DIR = SYSTEM_DIR / "logs"
LOG_LINE_LIMIT = 2000

PROPERTY_TIMEOUT = 50
INSPECTION_TIMEOUT = 60
ENTER_RETRY_INTERVAL = 4
ENTER_MAX_ATTEMPTS = 3
RESULT_TIMEOUT = 180


def now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def print_progress(message):
    """Never let a legacy Windows console encoding abort an active submission."""
    try:
        print(message, flush=True)
    except UnicodeEncodeError:
        encoding = getattr(sys.stdout, "encoding", None) or "ascii"
        safe_message = str(message).encode(encoding, errors="replace").decode(encoding)
        print(safe_message, flush=True)


def quota_ready_for_retry(entry):
    """A quota result is held only for its calendar day, then retried."""
    quota_date = str(
        entry.get("quotaDate") or entry.get("updatedAt") or ""
    )[:10]
    try:
        return datetime.strptime(quota_date, "%Y-%m-%d").date() < datetime.now().date()
    except ValueError:
        # Older records without a usable date are safe to retry once.
        return True


def prune_old_logs():
    """Keep only the newest LOG_LINE_LIMIT JSONL events across all runs."""
    try:
        log_files = sorted(
            LOGS_DIR.rglob("run_*.jsonl"),
            key=lambda path: (path.stat().st_mtime, str(path)),
        )
        file_lines = []
        total_lines = 0
        for path in log_files:
            with open(path, "r", encoding="utf-8", errors="replace") as handle:
                lines = handle.readlines()
            file_lines.append((path, lines))
            total_lines += len(lines)

        excess = total_lines - LOG_LINE_LIMIT
        for path, lines in file_lines:
            if excess <= 0:
                break
            if excess >= len(lines):
                excess -= len(lines)
                run_id = path.stem.removeprefix("run_")
                path.unlink(missing_ok=True)
                for artifact in path.parent.glob(f"{run_id}_*"):
                    artifact.unlink(missing_ok=True)
                continue

            temp_path = path.with_suffix(path.suffix + ".tmp")
            with open(temp_path, "w", encoding="utf-8", newline="\n") as handle:
                handle.writelines(lines[excess:])
            temp_path.replace(path)
            break
    except OSError:
        # Logging must never interfere with the actual submission run.
        pass


def load_json(path, default):
    try:
        with open(path, "r", encoding="utf-8-sig") as handle:
            return json.load(handle)
    except (OSError, ValueError):
        return default


def atomic_write_json(path, data):
    path = Path(path)
    temp_path = path.with_suffix(path.suffix + ".tmp")
    with open(temp_path, "w", encoding="utf-8", newline="\n") as handle:
        json.dump(data, handle, ensure_ascii=False, indent=2)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temp_path, path)


def save_history(data):
    if HISTORY_PATH.exists():
        try:
            shutil.copy2(HISTORY_PATH, str(HISTORY_PATH) + ".bak")
        except OSError:
            pass
    atomic_write_json(HISTORY_PATH, data)


def domain_from_url(url):
    try:
        host = (urlsplit(url).hostname or "").lower().strip(".")
    except ValueError:
        return ""
    if host.startswith("www."):
        return host[4:]
    return host


def profile_number(profile_id):
    try:
        return int(profile_id.rsplit("_", 1)[1])
    except (ValueError, IndexError):
        return 999


def find_edge():
    candidates = [
        Path(os.environ.get("ProgramFiles(x86)", ""))
        / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("ProgramFiles", ""))
        / "Microsoft/Edge/Application/msedge.exe",
        Path(os.environ.get("LOCALAPPDATA", ""))
        / "Microsoft/Edge/Application/msedge.exe",
    ]
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    raise RuntimeError("Không tìm thấy Microsoft Edge.")


def get_pages(port):
    response = requests.get(f"http://127.0.0.1:{port}/json/list", timeout=2)
    response.raise_for_status()
    return [
        page
        for page in response.json()
        if page.get("type") == "page" and page.get("webSocketDebuggerUrl")
    ]


def cdp_ready(port):
    try:
        return bool(get_pages(port))
    except (requests.RequestException, ValueError):
        return False


def ensure_profile_browser(profile_id, first_url, browser_mode="visible"):
    number = profile_number(profile_id)
    if number < 1 or number > 14:
        raise RuntimeError(f"Profile không hợp lệ: {profile_id}")
    port = 9300 + number
    if cdp_ready(port):
        return port, None

    edge_path = find_edge()
    profile_path = ROOT_DIR / profile_id
    profile_path.mkdir(parents=True, exist_ok=True)
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    edge_args = [
        str(edge_path),
        f"--user-data-dir={profile_path}",
        "--profile-directory=Default",
        "--remote-debugging-address=127.0.0.1",
        f"--remote-debugging-port={port}",
        "--remote-allow-origins=*",
        "--no-first-run",
        "--no-default-browser-check",
    ]
    if browser_mode == "hidden":
        edge_args += ["--headless=new", "--window-size=1440,1000"]
    else:
        edge_args += ["--start-maximized", "--new-window"]
    edge_args.append(first_url)
    browser_process = subprocess.Popen(
        edge_args,
        creationflags=creation_flags,
    )

    deadline = time.time() + 25
    while time.time() < deadline:
        if cdp_ready(port):
            return port, browser_process
        time.sleep(0.5)
    raise RuntimeError(
        f"Không kết nối được CDP của {profile_id}. "
        "Có thể profile đang mở từ trước nhưng không có cổng điều khiển."
    )


def get_page_for_port(port):
    pages = get_pages(port)
    search_console = [
        page for page in pages if "search.google.com" in page.get("url", "")
    ]
    if search_console:
        return search_console[0]
    if pages:
        return pages[0]
    raise RuntimeError("Không tìm thấy tab Edge để điều khiển.")


def evaluate(cdp, expression):
    result = cdp.call(
        "Runtime.evaluate",
        {
            "expression": expression,
            "returnByValue": True,
            "awaitPromise": True,
        },
    )
    inner = result.get("result", {})
    if inner.get("subtype") == "error":
        raise RuntimeError(inner.get("description", "Lỗi JavaScript trong trang."))
    return inner.get("value")


def maximize_browser_window(cdp, target_id=""):
    try:
        params = {"targetId": target_id} if target_id else {}
        window = cdp.call("Browser.getWindowForTarget", params)
        window_id = window.get("windowId")
        if window_id is None:
            return False
        bounds = window.get("bounds", {})
        if bounds.get("windowState") == "minimized":
            cdp.call(
                "Browser.setWindowBounds",
                {
                    "windowId": window_id,
                    "bounds": {"windowState": "normal"},
                },
            )
        cdp.call(
            "Browser.setWindowBounds",
            {
                "windowId": window_id,
                "bounds": {"windowState": "maximized"},
            },
        )
        return True
    except (RuntimeError, TypeError, ValueError):
        return False


def visible_page_state(cdp):
    expression = r"""
(() => {
  const visible = (element) => {
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== 'none' &&
      style.visibility !== 'hidden' &&
      rect.width > 0 &&
      rect.height > 0;
  };
  const inputs = Array.from(document.querySelectorAll(
    'input[role="combobox"][aria-label^="Kiểm tra mọi URL trong"]'
  )).filter(visible).map((element) => ({
    label: element.getAttribute('aria-label') || '',
    disabled: element.disabled || element.getAttribute('aria-disabled') === 'true'
  }));
  return {
    href: location.href,
    title: document.title || '',
    body: (document.body && document.body.innerText || '').trim(),
    inputs
  };
})()
"""
    return evaluate(cdp, expression) or {}


def validate_property(cdp, expected_domain):
    deadline = time.time() + PROPERTY_TIMEOUT
    last_state = {}
    mismatch_labels = []
    expected = expected_domain.casefold()
    permission_phrases = (
        "không có quyền",
        "bạn không có quyền",
        "you don't have permission",
        "you do not have permission",
        "property not in account",
        "không thể truy cập tài sản",
    )

    while time.time() < deadline:
        try:
            last_state = visible_page_state(cdp)
        except RuntimeError:
            time.sleep(0.5)
            continue
        href = str(last_state.get("href", "")).casefold()
        body = str(last_state.get("body", ""))
        lowered_body = body.casefold()
        if "accounts.google.com" in href:
            return {
                "status": "SYSTEM_ERROR",
                "message": "Profile chưa đăng nhập Google.",
                "raw": body,
            }
        if any(phrase in lowered_body for phrase in permission_phrases):
            return {
                "status": "MAPPING_ERROR",
                "message": "Gmail/profile không có quyền với property.",
                "raw": body,
            }

        active_inputs = [
            item
            for item in last_state.get("inputs", [])
            if not item.get("disabled")
        ]
        if active_inputs:
            labels = [str(item.get("label", "")) for item in active_inputs]
            if any(expected in label.casefold() for label in labels):
                return {
                    "status": "OK",
                    "message": labels[0],
                    "raw": body,
                }
            mismatch_labels = labels
        time.sleep(0.5)

    if mismatch_labels:
        return {
            "status": "MAPPING_ERROR",
            "message": (
                "Property đang mở không đúng tên miền. "
                f"Cần: {expected_domain}; thấy: {' | '.join(mismatch_labels)}"
            ),
            "raw": str(last_state.get("body", "")),
        }
    return {
        "status": "SYSTEM_ERROR",
        "message": "Hết thời gian xác minh property.",
        "raw": str(last_state.get("body", "")),
    }


def enter_inspection_url(cdp, url):
    focus_script = r"""
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
    if not evaluate(cdp, focus_script):
        raise RuntimeError("Không tìm thấy ô Kiểm tra URL đang hoạt động.")
    cdp.call("Input.insertText", {"text": url})
    entered_value = evaluate(
        cdp,
        r"""
(() => {
  const input = document.querySelector(
    'input[role="combobox"][aria-label^="Kiểm tra mọi URL trong"]:not([disabled])'
  );
  return input ? input.value : '';
})()
""",
    )
    if entered_value != url:
        raise RuntimeError("Ô Kiểm tra URL chưa nhận đủ URL cần nhập.")
    for event_type in ("keyDown", "keyUp"):
        cdp.call(
            "Input.dispatchKeyEvent",
            {
                "type": event_type,
                "key": "Enter",
                "code": "Enter",
                "windowsVirtualKeyCode": 13,
                "nativeVirtualKeyCode": 13,
            },
        )


def inspection_state(cdp, expected_url):
    expression = r"""
(() => {
  const expectedUrl = __EXPECTED_URL__;
  const visible = (element) => {
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== 'none' &&
      style.visibility !== 'hidden' &&
      rect.width > 0 &&
      rect.height > 0;
  };
  const candidates = Array.from(
    document.querySelectorAll('[role="button"], button')
  ).filter(visible);
  const button = candidates.find((element) => {
    const label = (element.getAttribute('aria-label') || '').trim();
    const text = (element.textContent || '').trim();
    return label.startsWith('Yêu cầu lập chỉ mục') ||
      text.startsWith('Yêu cầu lập chỉ mục');
  });
  const input = Array.from(document.querySelectorAll(
    'input[role="combobox"][aria-label^="Kiểm tra mọi URL trong"]'
  )).find(visible);
  const body = (document.body && document.body.innerText || '').trim();
  return {
    href: location.href,
    body,
    targetVisible: body.includes(expectedUrl),
    inputValue: input ? input.value : '',
    inputDisabled: input ? (
      input.disabled || input.getAttribute('aria-disabled') === 'true'
    ) : true,
    requestFound: Boolean(button),
    requestDisabled: button ? (
      button.disabled || button.getAttribute('aria-disabled') === 'true'
    ) : true
  };
})()
"""
    expression = expression.replace(
        "__EXPECTED_URL__",
        json.dumps(expected_url, ensure_ascii=False),
    )
    return evaluate(cdp, expression) or {}


def wait_for_inspection(cdp, expected_url, initial_href):
    deadline = time.time() + INSPECTION_TIMEOUT
    next_enter_retry = time.time() + ENTER_RETRY_INTERVAL
    enter_attempts = 1
    inspection_started = False
    last_state = {}
    while time.time() < deadline:
        try:
            last_state = inspection_state(cdp, expected_url)
        except RuntimeError:
            time.sleep(0.7)
            continue
        current_href = str(last_state.get("href", ""))
        body = str(last_state.get("body", ""))
        lowered = body.casefold()
        target_visible = bool(last_state.get("targetVisible"))
        inspection_started = inspection_started or (
            current_href != initial_href
            or (
                target_visible
                and (
                    last_state.get("inputDisabled")
                    or "kiểm tra url đang hoạt động" in lowered
                    or "url nằm trên google" in lowered
                    or "url không nằm trên google" in lowered
                    or last_state.get("requestFound")
                )
            )
        )
        if (
            target_visible
            and "url nằm trên google" in lowered
            and "url không nằm trên google" not in lowered
        ):
            return {
                "status": "ALREADY_INDEXED",
                "message": "URL nằm trên Google.",
                "raw": body,
                "enterAttempts": enter_attempts,
            }
        if (
            target_visible
            and last_state.get("requestFound")
            and not last_state.get("requestDisabled")
        ):
            return {
                "status": "READY",
                "message": "Đã sẵn sàng yêu cầu lập chỉ mục.",
                "raw": body,
                "enterAttempts": enter_attempts,
            }
        if (
            not inspection_started
            and enter_attempts < ENTER_MAX_ATTEMPTS
            and time.time() >= next_enter_retry
        ):
            enter_inspection_url(cdp, expected_url)
            enter_attempts += 1
            next_enter_retry = time.time() + ENTER_RETRY_INTERVAL
        time.sleep(0.7)
    return {
        "status": "INSPECTION_TIMEOUT",
        "message": (
            f"Quá {INSPECTION_TIMEOUT} giây chưa thấy kết quả hoặc nút "
            "Yêu cầu lập chỉ mục; đã bỏ qua URL."
        ),
        "raw": str(last_state.get("body", "")),
        "enterAttempts": enter_attempts,
    }


def click_request_indexing(cdp, expected_url):
    expression = r"""
(() => {
  const expectedUrl = __EXPECTED_URL__;
  const body = (document.body && document.body.innerText || '').trim();
  if (!body.includes(expectedUrl)) {
    return { clicked: false, reason: 'wrong-url' };
  }
  const candidates = Array.from(
    document.querySelectorAll('[role="button"], button')
  );
  const button = candidates.find((element) => {
    const label = (element.getAttribute('aria-label') || '').trim();
    const text = (element.textContent || '').trim();
    return label.startsWith('Yêu cầu lập chỉ mục') ||
      text.startsWith('Yêu cầu lập chỉ mục');
  });
  if (!button) return { clicked: false, reason: 'not-found' };
  if (button.disabled || button.getAttribute('aria-disabled') === 'true') {
    return { clicked: false, reason: 'disabled' };
  }
  button.click();
  return {
    clicked: true,
    label: button.getAttribute('aria-label') || button.textContent || ''
  };
})()
"""
    expression = expression.replace(
        "__EXPECTED_URL__",
        json.dumps(expected_url, ensure_ascii=False),
    )
    result = evaluate(cdp, expression) or {}
    if not result.get("clicked"):
        raise RuntimeError(
            "Không click được Yêu cầu lập chỉ mục: "
            + str(result.get("reason", "không rõ"))
        )


def classify_result(text):
    lowered = text.casefold()
    if (
        "đã yêu cầu lập chỉ mục" in lowered
        or "url đã được thêm vào hàng đợi ưu tiên" in lowered
    ):
        return "SUCCESS"
    quota_phrases = (
        "đã vượt hạn ngạch",
        "vượt quá hạn mức",
        "vượt quá hạn ngạch hàng ngày",
        "vượt hạn ngạch",
        "đã đạt đến hạn mức",
        "đã đạt giới hạn",
        "quota exceeded",
        "daily quota",
    )
    if any(phrase in lowered for phrase in quota_phrases):
        return "QUOTA"
    error_phrases = (
        "rất tiếc! đã xảy ra sự cố",
        "chúng tôi gặp sự cố khi gửi yêu cầu lập chỉ mục",
    )
    if any(phrase in lowered for phrase in error_phrases):
        return "ERROR"
    return None


def wait_for_result(cdp):
    expression = r"""
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
    return style.display !== 'none' &&
      style.visibility !== 'hidden' &&
      rect.width > 0 &&
      rect.height > 0;
  };
  const texts = Array.from(document.querySelectorAll(selectors.join(',')))
    .filter(visible)
    .map((element) => (element.innerText || '').trim())
    .filter(Boolean);
  // Một số bản GSC hiển thị cảnh báo quota nhưng không gắn role=dialog
  // hoặc aria-live vào khung cảnh báo. Luôn thêm nội dung trang làm dự phòng.
  const pageText = (document.body && document.body.innerText || '').trim();
  if (pageText) texts.push(pageText);
  return texts.filter((text, index, all) => all.indexOf(text) === index);
})()
"""
    deadline = time.time() + RESULT_TIMEOUT
    last_text = ""
    while time.time() < deadline:
        try:
            texts = evaluate(cdp, expression) or []
        except RuntimeError:
            time.sleep(0.7)
            continue
        for text in texts:
            last_text = text
            status = classify_result(text)
            if status:
                return {"status": status, "message": text, "raw": text}
        time.sleep(0.7)
    if last_text:
        return {
            "status": "UNKNOWN",
            "message": "Có nội dung kết quả nhưng chưa nhận diện được.",
            "raw": last_text,
        }
    return {
        "status": "TIMEOUT",
        "message": "Không thấy kết quả sau 180 giây.",
        "raw": last_text,
    }


def capture_screenshot(cdp, destination):
    try:
        cdp.call("Page.enable")
        result = cdp.call(
            "Page.captureScreenshot",
            {"format": "png", "captureBeyondViewport": False},
        )
        data = result.get("data")
        if data:
            with open(destination, "wb") as handle:
                handle.write(base64.b64decode(data))
            return str(destination)
    except Exception:
        return ""
    return ""


class Runner:
    def __init__(self, limit, target_url="", browser_mode="visible"):
        self.limit = limit
        self.target_url = target_url.strip()
        self.browser_mode = browser_mode
        self.run_id = datetime.now().strftime("%Y%m%d_%H%M%S") + "_" + uuid.uuid4().hex[:6]
        self.run_dir = LOGS_DIR / datetime.now().strftime("%Y-%m-%d")
        self.run_dir.mkdir(parents=True, exist_ok=True)
        self.log_path = self.run_dir / f"run_{self.run_id}.jsonl"
        self.config = load_json(CONFIG_PATH, {})
        self.history = load_json(HISTORY_PATH, {"version": 1, "urls": []})
        self.history.setdefault("version", 1)
        self.history.setdefault("urls", [])

        self.blocked_gmails = set()
        self.blocked_domain_gmails = set()
        self.domain_gmail_errors = defaultdict(int)
        self.gmail_failed_domain_streak = defaultdict(int)
        self.gmail_last_failed_domain = {}
        self.profile_processes = {}
        self.summary = defaultdict(int)
        self.started_at = time.time()
        self.started_at_text = now_text()
        self.current_position = 0
        self.current_total = 0
        self.last_progress = ""

    def log(self, step, **fields):
        event = {
            "timestamp": now_text(),
            "runId": self.run_id,
            "step": step,
            **fields,
        }
        with open(self.log_path, "a", encoding="utf-8", newline="\n") as handle:
            handle.write(json.dumps(event, ensure_ascii=False) + "\n")
            handle.flush()
            os.fsync(handle.fileno())
        prune_old_logs()

    def progress(self, message, state="RUNNING", current_url="", **extra):
        data = {
            "runId": self.run_id,
            "state": state,
            "message": message,
            "currentUrl": current_url,
            "position": self.current_position,
            "total": self.current_total,
            "processId": os.getpid(),
            "startedAt": self.started_at_text,
            "elapsedSeconds": round(time.time() - self.started_at, 1),
            "updatedAt": now_text(),
            "logPath": str(self.log_path),
            "summary": dict(self.summary),
            **extra,
        }
        atomic_write_json(PROGRESS_PATH, data)
        self.last_progress = message
        print_progress(message)

    def save_entry(self, entry, status, message="", **extra):
        entry["status"] = status
        entry["message"] = message
        entry["updatedAt"] = now_text()
        for key, value in extra.items():
            entry[key] = value
        save_history(self.history)

    def candidates_for_domain(self, domain):
        account_emails = []
        for account in self.config.get("accounts", []):
            email = str(account.get("email", "")).strip()
            domains = [
                str(value).lower().strip(".")
                for value in account.get("domains", [])
            ]
            matched = any(
                domain == value or domain == f"www.{value}"
                for value in domains
            )
            if email and matched:
                account_emails.append((email, next(
                    value
                    for value in domains
                    if domain == value or domain == f"www.{value}"
                )))

        candidates = []
        for profile in self.config.get("profiles", []):
            profile_email = str(profile.get("email", "")).strip()
            for email, property_domain in account_emails:
                if profile_email.casefold() == email.casefold():
                    candidates.append(
                        {
                            "profile": str(profile.get("id", "")),
                            "email": email,
                            "domain": property_domain,
                        }
                    )
        candidates.sort(key=lambda item: profile_number(item["profile"]))
        return candidates

    def property_url(self, domain):
        resource = f"https://{domain}/"
        return (
            "https://search.google.com/u/0/search-console"
            f"?resource_id={quote(resource, safe='')}&hl=vi"
        )

    def submit_attempt(self, url, domain, candidate, attempt):
        profile_id = candidate["profile"]
        email = candidate["email"]
        property_domain = candidate["domain"]
        gsc_url = self.property_url(property_domain)
        started = time.time()
        cdp = None

        self.log(
            "PROFILE_SELECTED",
            url=url,
            domain=domain,
            gmail=email,
            profile=profile_id,
            attempt=attempt,
            propertyDomain=property_domain,
        )
        self.progress(
            f"{domain} → {profile_id} → đang mở GSC",
            current_url=url,
        )

        try:
            port, browser_process = ensure_profile_browser(
                profile_id, gsc_url, self.browser_mode
            )
            if browser_process:
                self.profile_processes[profile_id] = browser_process
            page = get_page_for_port(port)
            cdp = Cdp(page["webSocketDebuggerUrl"])
            cdp.call("Runtime.enable")
            cdp.call("Page.enable")
            window_maximized = False
            if self.browser_mode == "visible":
                window_maximized = maximize_browser_window(
                    cdp,
                    str(page.get("id", "")),
                )
            self.log(
                "BROWSER_WINDOW_STATE",
                url=url,
                domain=domain,
                gmail=email,
                profile=profile_id,
                attempt=attempt,
                maximized=window_maximized,
                browserMode=self.browser_mode,
            )
            cdp.call("Page.navigate", {"url": gsc_url})
            self.log(
                "GSC_OPENED",
                url=url,
                domain=domain,
                gmail=email,
                profile=profile_id,
                attempt=attempt,
                port=port,
            )

            validation = validate_property(cdp, property_domain)
            if validation["status"] != "OK":
                return {
                    **validation,
                    "elapsed": round(time.time() - started, 2),
                    "cdp": cdp,
                }
            self.log(
                "PROPERTY_VERIFIED",
                url=url,
                domain=domain,
                gmail=email,
                profile=profile_id,
                attempt=attempt,
                rawGoogleMessage=validation.get("message", ""),
            )

            initial_href = str(evaluate(cdp, "location.href") or "")
            enter_inspection_url(cdp, url)
            self.log(
                "URL_ENTERED",
                url=url,
                domain=domain,
                gmail=email,
                profile=profile_id,
                attempt=attempt,
            )

            inspection = wait_for_inspection(cdp, url, initial_href)
            if inspection["status"] == "ALREADY_INDEXED":
                return {
                    **inspection,
                    "elapsed": round(time.time() - started, 2),
                    "cdp": cdp,
                }
            if inspection["status"] != "READY":
                return {
                    **inspection,
                    "elapsed": round(time.time() - started, 2),
                    "cdp": cdp,
                }

            click_request_indexing(cdp, url)
            self.log(
                "INDEX_REQUEST_CLICKED",
                url=url,
                domain=domain,
                gmail=email,
                profile=profile_id,
                attempt=attempt,
            )
            result = wait_for_result(cdp)
            return {
                **result,
                "elapsed": round(time.time() - started, 2),
                "cdp": cdp,
            }
        except Exception as exc:
            return {
                "status": "SYSTEM_ERROR",
                "message": str(exc),
                "raw": "",
                "elapsed": round(time.time() - started, 2),
                "cdp": cdp,
            }

    def reset_success_counters(self, domain, email):
        key = (domain, email.casefold())
        self.domain_gmail_errors[key] = 0
        self.gmail_failed_domain_streak[email.casefold()] = 0
        self.gmail_last_failed_domain.pop(email.casefold(), None)

    def register_domain_failure(self, domain, email):
        email_key = email.casefold()
        previous = self.gmail_last_failed_domain.get(email_key)
        if previous != domain:
            self.gmail_failed_domain_streak[email_key] += 1
            self.gmail_last_failed_domain[email_key] = domain
        if self.gmail_failed_domain_streak[email_key] >= 2:
            self.blocked_gmails.add(email_key)
            self.log(
                "GMAIL_PAUSED",
                domain=domain,
                gmail=email,
                reason="Hai tên miền khác nhau liên tiếp đạt hai lỗi.",
            )
            return True
        return False

    def save_failure_artifacts(self, cdp, url, profile_id, status, raw):
        safe_name = "".join(
            char if char.isalnum() or char in "-_" else "_"
            for char in domain_from_url(url)
        )
        prefix = f"{self.run_id}_{safe_name}_{profile_id}_{status.lower()}"
        text_path = self.run_dir / f"{prefix}.txt"
        with open(text_path, "w", encoding="utf-8") as handle:
            handle.write(raw or "")
        screenshot_path = ""
        if cdp:
            screenshot_path = capture_screenshot(
                cdp, self.run_dir / f"{prefix}.png"
            )
        return str(text_path), screenshot_path

    def release_profile_browser(self, profile_id, cdp, close_browser=False):
        """Close only an Edge window started by this automatic run."""
        if cdp:
            try:
                cdp.close()
            except Exception:
                pass

        if not close_browser:
            return

        browser_process = self.profile_processes.pop(profile_id, None)
        if not browser_process or browser_process.poll() is not None:
            return

        self.log(
            "BROWSER_CLOSING",
            profile=profile_id,
            pid=browser_process.pid,
            reason="URL đã có kết quả; đóng Edge do submit tự mở.",
        )
        try:
            subprocess.run(
                [
                    "taskkill",
                    "/PID",
                    str(browser_process.pid),
                    "/T",
                    "/F",
                ],
                check=False,
                capture_output=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        finally:
            self.log(
                "BROWSER_CLOSED",
                profile=profile_id,
                pid=browser_process.pid,
            )

    def close_all_profile_browsers(self):
        for profile_id in list(self.profile_processes):
            self.release_profile_browser(profile_id, None, close_browser=True)

    def process_entry(self, entry):
        url = str(entry.get("url", "")).strip()
        domain = domain_from_url(url)
        if not domain:
            self.save_entry(entry, "UNMAPPED", "URL không hợp lệ.")
            self.summary["unmapped"] += 1
            return "CONTINUE"

        all_candidates = self.candidates_for_domain(domain)
        if not all_candidates:
            self.save_entry(
                entry,
                "UNMAPPED",
                "Không tìm thấy Gmail/profile được gán cho tên miền.",
            )
            self.summary["unmapped"] += 1
            self.log(
                "URL_UNMAPPED",
                url=url,
                domain=domain,
                result="UNMAPPED",
            )
            return "CONTINUE"

        self.save_entry(entry, "RUNNING", "Đang xử lý.")
        self.log("URL_SELECTED", url=url, domain=domain)
        attempt = 0
        attempted_submission = False
        last_result = None

        while True:
            available = [
                item
                for item in all_candidates
                if item["email"].casefold() not in self.blocked_gmails
                and (
                    domain,
                    item["email"].casefold(),
                )
                not in self.blocked_domain_gmails
            ]
            if not available:
                if not attempted_submission:
                    if last_result == "MAPPING_ERROR":
                        self.save_entry(
                            entry,
                            "UNMAPPED",
                            "Các profile được gán không có quyền với tên miền.",
                        )
                        self.summary["unmapped"] += 1
                    else:
                        self.save_entry(
                            entry,
                            "PENDING",
                            "Tạm dừng trong phiên vì không còn Gmail khả dụng.",
                        )
                        self.summary["paused"] += 1
                return "CONTINUE"

            candidate = available[0]
            attempt += 1
            email = candidate["email"]
            profile_id = candidate["profile"]
            self.save_entry(
                entry,
                "RUNNING",
                f"Đang thử bằng {profile_id}.",
                lastGmail=email,
                lastProfile=profile_id,
                attempts=int(entry.get("attempts", 0)) + 1,
            )
            result = self.submit_attempt(
                url, domain, candidate, attempt
            )
            status = result["status"]
            message = result.get("message", "")
            raw = result.get("raw", "")
            elapsed = result.get("elapsed", 0)
            cdp = result.pop("cdp", None)

            self.log(
                "RESULT_DETECTED",
                url=url,
                domain=domain,
                gmail=email,
                profile=profile_id,
                attempt=attempt,
                result=status,
                rawGoogleMessage=raw,
                elapsedSeconds=elapsed,
                enterAttempts=result.get("enterAttempts", 0),
                domainGmailConsecutiveErrors=self.domain_gmail_errors[
                    (domain, email.casefold())
                ],
                gmailConsecutiveFailedDomains=self.gmail_failed_domain_streak[
                    email.casefold()
                ],
            )

            if status == "SUCCESS":
                attempted_submission = True
                self.reset_success_counters(domain, email)
                self.save_entry(
                    entry,
                    "SUBMITTED",
                    message,
                    lastGmail=email,
                    lastProfile=profile_id,
                )
                self.summary["completed"] += 1
                self.progress(
                    f"{domain} → HOÀN THÀNH bằng {profile_id}",
                    current_url=url,
                )
                self.release_profile_browser(profile_id, cdp)
                return "CONTINUE"

            if status == "ALREADY_INDEXED":
                self.reset_success_counters(domain, email)
                self.save_entry(
                    entry,
                    "SKIPPED",
                    "Đã index, không dùng lượt submit.",
                    lastGmail=email,
                    lastProfile=profile_id,
                )
                self.summary["skipped"] += 1
                self.progress(
                    f"{domain} → BỎ QUA vì URL đã index",
                    current_url=url,
                )
                self.release_profile_browser(profile_id, cdp)
                return "CONTINUE"

            if status == "INSPECTION_TIMEOUT":
                self.save_entry(
                    entry,
                    "SKIPPED",
                    message,
                    lastGmail=email,
                    lastProfile=profile_id,
                )
                self.summary["skipped"] += 1
                self.log(
                    "URL_SKIPPED",
                    url=url,
                    domain=domain,
                    gmail=email,
                    profile=profile_id,
                    reason=message,
                    enterAttempts=result.get("enterAttempts", 0),
                )
                self.progress(
                    f"{domain} → BỎ QUA vì quá 60 giây chưa có nút lập chỉ mục",
                    current_url=url,
                )
                self.release_profile_browser(profile_id, cdp)
                return "CONTINUE"

            if status == "MAPPING_ERROR":
                self.blocked_domain_gmails.add(
                    (domain, email.casefold())
                )
                last_result = "MAPPING_ERROR"
                self.log(
                    "PROFILE_MAPPING_REJECTED",
                    url=url,
                    domain=domain,
                    gmail=email,
                    profile=profile_id,
                    reason=message,
                )
                self.progress(
                    f"{domain} → {profile_id} không có quyền; đang tìm profile khác",
                    current_url=url,
                )
                self.release_profile_browser(profile_id, cdp)
                continue

            if status == "QUOTA":
                attempted_submission = True
                self.blocked_gmails.add(email.casefold())
                last_result = "QUOTA"
                text_path, image_path = self.save_failure_artifacts(
                    cdp, url, profile_id, status, raw
                )
                self.save_entry(
                    entry,
                    "QUOTA",
                    message,
                    lastGmail=email,
                    lastProfile=profile_id,
                    quotaDate=datetime.now().strftime("%Y-%m-%d"),
                    retryAfter=(datetime.now() + timedelta(days=1)).strftime(
                        "%Y-%m-%d"
                    ),
                )
                self.log(
                    "GMAIL_PAUSED",
                    url=url,
                    domain=domain,
                    gmail=email,
                    profile=profile_id,
                    reason="Vượt hạn ngạch trong phiên.",
                    textPath=text_path,
                    screenshotPath=image_path,
                )
                self.progress(
                    f"{email} → VƯỢT HẠN NGẠCH; đang chuyển Gmail khác",
                    current_url=url,
                )
                self.release_profile_browser(profile_id, cdp)
                continue

            if status == "ERROR":
                attempted_submission = True
                last_result = "ERROR"
                key = (domain, email.casefold())
                self.domain_gmail_errors[key] += 1
                current_streak = self.domain_gmail_errors[key]
                text_path, image_path = self.save_failure_artifacts(
                    cdp, url, profile_id, status, raw
                )
                self.save_entry(
                    entry,
                    "ERROR",
                    message,
                    lastGmail=email,
                    lastProfile=profile_id,
                )
                self.summary["errors"] += 1
                self.log(
                    "URL_ERROR",
                    url=url,
                    domain=domain,
                    gmail=email,
                    profile=profile_id,
                    consecutiveErrors=current_streak,
                    textPath=text_path,
                    screenshotPath=image_path,
                )
                self.progress(
                    f"{domain} → LỖI {current_streak}/2 trên {profile_id}",
                    current_url=url,
                )
                self.release_profile_browser(profile_id, cdp)

                if current_streak < 2:
                    return "CONTINUE"

                self.blocked_domain_gmails.add(
                    (domain, email.casefold())
                )
                gmail_paused = self.register_domain_failure(domain, email)
                self.log(
                    "PROFILE_FALLBACK",
                    url=url,
                    domain=domain,
                    gmail=email,
                    profile=profile_id,
                    reason=(
                        "Hai URL lỗi liên tiếp; chuyển Gmail phụ."
                        if not gmail_paused
                        else "Hai tên miền lỗi liên tiếp; dừng Gmail."
                    ),
                )
                continue

            text_path, image_path = self.save_failure_artifacts(
                cdp, url, profile_id, status, raw
            )
            self.save_entry(
                entry,
                "ERROR",
                f"{status}: {message}",
                lastGmail=email,
                lastProfile=profile_id,
            )
            self.log(
                "SYSTEM_ERROR",
                url=url,
                domain=domain,
                gmail=email,
                profile=profile_id,
                result=status,
                reason=message,
                textPath=text_path,
                screenshotPath=image_path,
            )
            self.progress(
                f"Lỗi hệ thống tại {domain}: {message}",
                state="RUNNING",
                current_url=url,
            )
            self.release_profile_browser(profile_id, cdp)
            self.summary["systemErrors"] += 1
            return "STOP"

    def run(self):
        if STOP_FLAG_PATH.exists():
            STOP_FLAG_PATH.unlink()

        if self.target_url:
            target_key = self.target_url.casefold()
            pending = [
                entry
                for entry in self.history.get("urls", [])
                if str(entry.get("url", "")).strip().casefold() == target_key
            ][:1]
            if not pending:
                raise RuntimeError("Không tìm thấy URL đã chọn trong hàng đợi.")
        else:
            pending = [
                entry
                for entry in self.history.get("urls", [])
                if (
                    str(entry.get("status", "PENDING")).upper()
                    in ("PENDING", "PRIORITY")
                    or (
                        str(entry.get("status", "")).upper() == "QUOTA"
                        and quota_ready_for_retry(entry)
                    )
                )
            ]
            pending.sort(
                key=lambda entry: (
                    0
                    if str(entry.get("status", "")).upper() == "PRIORITY"
                    else 1
                )
            )
            pending = pending[: self.limit]

        self.summary["selected"] = len(pending)
        self.current_total = len(pending)
        self.log(
            "RUN_STARTED",
            limit=self.limit,
            selected=len(pending),
        )
        self.progress(
            f"Bắt đầu phiên {self.run_id}: {len(pending)} URL.",
            selected=len(pending),
        )

        stopped = False
        for index, entry in enumerate(pending, start=1):
            if STOP_FLAG_PATH.exists():
                stopped = True
                self.log(
                    "RUN_STOPPED",
                    reason="Người dùng yêu cầu dừng sau URL hiện tại.",
                    processed=index - 1,
                )
                break
            self.current_position = index
            url = str(entry.get("url", ""))
            self.progress(
                f"URL {index}/{len(pending)}: {url}",
                current_url=url,
                position=index,
                total=len(pending),
            )
            action = self.process_entry(entry)
            self.summary["processed"] += 1
            if action == "STOP":
                stopped = True
                break

        final_statuses = [
            str(entry.get("status", "")).upper() for entry in pending
        ]
        self.summary["completed"] = final_statuses.count("SUBMITTED")
        self.summary["errors"] = final_statuses.count("ERROR")
        self.summary["skipped"] = final_statuses.count("SKIPPED")
        self.summary["quota"] = final_statuses.count("QUOTA")
        self.summary["unmapped"] = final_statuses.count("UNMAPPED")
        self.summary["pending"] = final_statuses.count("PENDING")

        elapsed = round(time.time() - self.started_at, 2)
        final_state = "STOPPED" if stopped else "FINISHED"
        final_message = (
            "Đã dừng phiên."
            if stopped
            else "Đã hoàn tất danh sách của phiên."
        )
        self.close_all_profile_browsers()
        self.log(
            "RUN_FINISHED",
            state=final_state,
            elapsedSeconds=elapsed,
            summary=dict(self.summary),
            blockedGmails=sorted(self.blocked_gmails),
        )
        self.progress(
            final_message,
            state=final_state,
            elapsedSeconds=elapsed,
        )
        return 0 if not stopped else 3


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, required=True)
    parser.add_argument("--url", default="")
    parser.add_argument(
        "--browser-mode", choices=("visible", "hidden"), default="visible"
    )
    args = parser.parse_args()
    if args.limit < 1:
        print("Số URL phải lớn hơn 0.", file=sys.stderr)
        return 2
    runner = Runner(args.limit, args.url, args.browser_mode)
    try:
        return runner.run()
    except Exception as exc:
        runner.close_all_profile_browsers()
        runner.log("SYSTEM_ERROR", reason=str(exc))
        runner.progress(
            f"Lỗi hệ thống: {exc}",
            state="STOPPED",
        )
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
