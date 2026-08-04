from __future__ import annotations

import json
import mimetypes
import os
import re
import socket
import subprocess
import sys
import threading
import time
import uuid
import webbrowser
from dataclasses import dataclass, asdict
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse


APP_DIR = Path(__file__).resolve().parent
STATIC_DIR = APP_DIR / "static"
LAUNCHERS_FILE = APP_DIR / "launchers.json"
LAUNCHER_GROUPS_FILE = APP_DIR / "launcher_groups.json"
LOGS_DIR = APP_DIR / "logs"
VOICE_CONTROL_CONFIG_FILE = APP_DIR / "voice_control.json"
VOICE_CONTROL_PORT = 8766
SUBMIT_PROFILE_MANAGER = Path(
    r"D:\CodexProjects\Hotkeyvip\06_du_lieu_chay\submit_edge_profiles\QUAN_LY_PROFILE_SUBMIT.bat"
)
DEFAULT_REPOSITORY = Path(r"D:\CodexProjects\Hotkeyvip")
EXCLUDED_TOP_LEVEL = {"07_ket_qua"}
HOST = "127.0.0.1"
PORT = 8765
LAUNCHER_EXTENSIONS = {
    ".py", ".pyw", ".ahk",
    ".ini", ".xlsx", ".xlsm", ".xls", ".docx", ".doc",
}
VOICE_PROCESS = None
VOICE_OVERLAY_HWND = None
RUNNING_LAUNCHERS: dict[str, list[dict]] = {}
RUNNING_LAUNCHERS_LOCK = threading.Lock()


def load_voice_control_config() -> dict:
    default = {
        "launcherPath": r"D:\CodexProjects\VoiceControlV3_HoanChinh\CODEX_THO_MAY\CHAY_THO_MAY.bat",
        "url": f"http://127.0.0.1:{VOICE_CONTROL_PORT}/voice",
    }
    try:
        stored = json.loads(VOICE_CONTROL_CONFIG_FILE.read_text(encoding="utf-8"))
        if isinstance(stored, dict):
            default.update(stored)
    except Exception:
        pass
    return default


def voice_control_is_ready() -> bool:
    try:
        with socket.create_connection(("127.0.0.1", VOICE_CONTROL_PORT), timeout=0.25):
            return True
    except OSError:
        return False


def voice_control_pid() -> int | None:
    if os.name != "nt":
        return VOICE_PROCESS.pid if VOICE_PROCESS and VOICE_PROCESS.poll() is None else None
    try:
        command = (
            f"$row=Get-NetTCPConnection -LocalAddress 127.0.0.1 -LocalPort {VOICE_CONTROL_PORT} "
            "-State Listen -ErrorAction SilentlyContinue | Select-Object -First 1; "
            "if($row){$row.OwningProcess}"
        )
        result = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", command],
            capture_output=True,
            text=True,
            timeout=4,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        value = result.stdout.strip()
        return int(value) if value.isdigit() else None
    except Exception:
        return None


def set_voice_overlay_visible(visible: bool) -> bool:
    global VOICE_OVERLAY_HWND
    if os.name != "nt":
        return False
    pid = voice_control_pid()
    if pid is None:
        return False
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    matched = []
    process_windows = []

    @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    def enum_callback(hwnd, _lparam):
        process_id = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))
        if process_id.value == pid:
            process_windows.append(hwnd)
            length = user32.GetWindowTextLengthW(hwnd)
            title = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, title, length + 1)
            if "Voice Control" in title.value:
                matched.append(hwnd)
        return True

    user32.EnumWindows(enum_callback, 0)
    if matched:
        VOICE_OVERLAY_HWND = matched[0]
    elif visible and VOICE_OVERLAY_HWND and user32.IsWindow(VOICE_OVERLAY_HWND):
        matched = [VOICE_OVERLAY_HWND]
    elif visible and process_windows:
        matched = process_windows
    command = 4 if visible else 0  # SW_SHOWNOACTIVATE / SW_HIDE
    for hwnd in matched:
        user32.ShowWindow(hwnd, command)
    return bool(matched)

CORE_EXTENSIONS = {".py", ".ahk", ".xlsm", ".xlsx", ".ini", ".md", ".txt", ".bat", ".ps1"}
TEMP_MARKERS = (
    "cache", "temp", "tmp", "__pycache__", ".pyc", "logs", "filelog_anhloi",
    "crashpad", "component_crx_cache", "code cache", "gpu cache",
)
HISTORY_MARKERS = (
    "backup", "bacup", "copy", "bản cũ", "ban cu", "bancu", "du phong",
    "chua can sai", "chưa cần xài",
)
TEST_MARKERS = ("test", "thu_nghiem", "thử nghiệm", "ket_qua_test", "diagnostic")


DEFAULT_LAUNCHERS = [
    {
        "id": "viet-bai-3-luong",
        "name": "Viết bài tự động 3 luồng",
        "path": r"D:\CodexProjects\Hotkeyvip\02_viet_bai\TUDONG_3_LUONG_WORD_BRIEF_VA_ANH.py",
        "description": "Flow viết Word, tạo brief và ảnh nhanh.",
        "showConsole": True,
        "group": "Tự động",
        "order": 10,
    },
    {
        "id": "viet-bai-le",
        "name": "Viết bài chạy lẻ",
        "path": r"D:\CodexProjects\Hotkeyvip\02_viet_bai\code_tong_v3_no_mouse_CHAY_HIDE.py",
        "description": "Chạy tuần tự từng bài.",
        "showConsole": True,
        "group": "Tự động",
        "order": 20,
    },
    {
        "id": "mo-word-tu-excel",
        "name": "Mở Word từ Excel",
        "path": r"D:\CodexProjects\Hotkeyvip\03_dang_bai\phu_tro\mo_word_tu_excel.py",
        "description": "Mở đúng file Word theo dòng Excel đang chọn.",
        "showConsole": False,
        "group": "Cứu hộ thủ công",
        "order": 10,
    },
    {
        "id": "mo-url-dang-bai",
        "name": "Mở URL đăng bài",
        "path": r"D:\CodexProjects\Hotkeyvip\03_dang_bai\phu_tro\mo_url_dang_bai.py",
        "description": "Mở trang CMS theo dữ liệu hiện tại.",
        "showConsole": False,
        "group": "Cứu hộ thủ công",
        "order": 20,
    },
    {
        "id": "xu-ly-anh",
        "name": "Tổng hợp và xử lý ảnh",
        "path": r"D:\CodexProjects\Hotkeyvip\03_dang_bai\phu_tro\tong_hop_va_xu_ly_anh.py",
        "description": "Gom và chuẩn hóa ảnh trước khi đăng.",
        "showConsole": True,
        "group": "Công cụ",
        "order": 10,
    },
    {
        "id": "bai-viet-lien-quan",
        "name": "Bài viết liên quan",
        "path": r"D:\CodexProjects\Hotkeyvip\03_dang_bai\phu_tro\bai_viet_lien_quan.py",
        "description": "Thêm bài viết liên quan sau khi đăng.",
        "showConsole": True,
        "group": "Sau khi đăng",
        "order": 10,
    },
    {
        "id": "xuat-url-cms",
        "name": "Xuất URL từ ID CMS",
        "path": r"D:\CodexProjects\Hotkeyvip\03_dang_bai\phu_tro\xuat_url_tu_id_cms.py",
        "description": "Lấy URL thật từ ID bài đăng.",
        "showConsole": True,
        "group": "Sau khi đăng",
        "order": 20,
    },
]
DEFAULT_LAUNCHER_GROUPS = ["Chưa phân nhóm"]


def load_launcher_groups() -> list[str]:
    if LAUNCHER_GROUPS_FILE.exists():
        try:
            stored = json.loads(LAUNCHER_GROUPS_FILE.read_text(encoding="utf-8"))
            groups = [str(name).strip() for name in stored if str(name).strip()]
        except Exception:
            groups = []
    else:
        groups = []
    if not groups:
        groups = list(DEFAULT_LAUNCHER_GROUPS)
    if LAUNCHERS_FILE.exists():
        try:
            for item in json.loads(LAUNCHERS_FILE.read_text(encoding="utf-8")):
                name = str(item.get("group", "")).strip()
                if name and name not in groups:
                    groups.append(name)
        except Exception:
            pass
    save_launcher_groups(groups)
    return groups


def save_launcher_groups(groups: list[str]) -> None:
    unique = []
    for name in groups:
        cleaned = str(name).strip()
        if cleaned and cleaned not in unique:
            unique.append(cleaned)
    temporary = LAUNCHER_GROUPS_FILE.with_suffix(".json.tmp")
    temporary.write_text(json.dumps(unique, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(LAUNCHER_GROUPS_FILE)


def load_launchers() -> list[dict]:
    if not LAUNCHERS_FILE.exists():
        save_launchers(DEFAULT_LAUNCHERS)
    try:
        data = json.loads(LAUNCHERS_FILE.read_text(encoding="utf-8"))
        if not isinstance(data, list):
            return []
        defaults = {item["id"]: item for item in DEFAULT_LAUNCHERS}
        changed = False
        for index, item in enumerate(data):
            fallback = defaults.get(item.get("id"), {})
            if not item.get("group"):
                item["group"] = fallback.get("group", "Chưa phân nhóm")
                changed = True
            if "order" not in item:
                item["order"] = fallback.get("order", (index + 1) * 10)
                changed = True
        if changed:
            save_launchers(data)
        return data
    except Exception:
        return []


def save_launchers(items: list[dict]) -> None:
    temporary = LAUNCHERS_FILE.with_suffix(".json.tmp")
    temporary.write_text(
        json.dumps(items, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    temporary.replace(LAUNCHERS_FILE)


def launcher_view(item: dict) -> dict:
    result = dict(item)
    path = Path(str(item.get("path", "")))
    result["exists"] = path.is_file()
    result["validType"] = path.suffix.lower() in LAUNCHER_EXTENSIONS
    result["actionLabel"] = "Chạy" if path.suffix.lower() in {".py", ".pyw", ".ahk"} else "Mở"
    result["hasLog"] = False
    result["lastLogState"] = ""
    result["lastLogDuration"] = 0
    result["lastLogStartedAt"] = 0
    result["runningPids"] = []
    launcher_id = str(item.get("id", "unknown"))
    with RUNNING_LAUNCHERS_LOCK:
        active = [
            entry for entry in RUNNING_LAUNCHERS.get(launcher_id, [])
            if entry["process"].poll() is None
        ]
        RUNNING_LAUNCHERS[launcher_id] = active
    if active:
        result["runningPids"] = [entry["process"].pid for entry in active]
        result["lastLogState"] = "running"
        result["lastLogStartedAt"] = min(entry["startedAt"] for entry in active)
    if path.suffix.lower() in {".py", ".pyw"}:
        log_dir = LOGS_DIR / re.sub(r"[^A-Za-z0-9_.-]+", "_", str(item.get("id", "unknown")))
        statuses = sorted(log_dir.glob("*.json"), key=lambda file: file.stat().st_mtime, reverse=True) if log_dir.exists() else []
        if statuses:
            result["hasLog"] = True
            try:
                result["lastLogState"] = json.loads(statuses[0].read_text(encoding="utf-8")).get("state", "")
                latest = json.loads(statuses[0].read_text(encoding="utf-8"))
                if not active:
                    result["lastLogState"] = latest.get("state", "")
                if not active:
                    result["lastLogDuration"] = latest.get("duration", 0)
                    result["lastLogStartedAt"] = latest.get("startedAt", 0)
            except Exception:
                result["lastLogState"] = "unknown"
    return result


@dataclass
class FileItem:
    path: str
    name: str
    folder: str
    extension: str
    size: int
    modified: float
    category: str
    reason: str


class RepositoryIndex:
    def __init__(self, root: Path):
        self.root = root
        self.files: list[FileItem] = []
        self.folders: list[dict] = []
        self.scanned_at = 0.0
        self.duration = 0.0
        self.lock = threading.Lock()
        self.error = ""
        self.scanning = False

    def classify(self, relative: str, suffix: str) -> tuple[str, str]:
        lower = relative.lower()
        if any(marker in lower for marker in TEMP_MARKERS):
            return "Tạm & có thể tạo lại", "Tên hoặc vị trí cho thấy đây là cache, log hay kết quả sinh tự động."
        if any(marker in lower for marker in HISTORY_MARKERS):
            return "Lịch sử / bản cũ", "Tên hoặc thư mục cho thấy đây là bản sao hay phiên bản cũ."
        if any(marker in lower for marker in TEST_MARKERS):
            return "Thử nghiệm", "Tên hoặc vị trí có dấu hiệu là dữ liệu kiểm tra."
        if relative.startswith(("01_hotkey", "02_viet_bai", "03_dang_bai", "04_excel", "05_cau_hinh")) and suffix in CORE_EXTENSIONS:
            return "Cần xem xét", "Nằm trong vùng chức năng chính; cần xác nhận trước khi phân loại."
        if relative.startswith(("06_du_lieu_chay", "07_ket_qua")):
            return "Dữ liệu làm việc", "Dữ liệu runtime hoặc kết quả công việc, không tự động coi là rác."
        return "Chưa xác định", "Chưa đủ bằng chứng để phân loại an toàn."

    def scan(self) -> None:
        started = time.time()
        self.scanning = True
        items: list[FileItem] = []
        folder_stats: dict[str, dict] = {}
        try:
            for current, dirs, names in os.walk(self.root):
                dirs[:] = [d for d in dirs if d != ".git"]
                current_path = Path(current)
                if current_path == self.root:
                    dirs[:] = [d for d in dirs if d not in EXCLUDED_TOP_LEVEL]
                for name in names:
                    path = current_path / name
                    try:
                        stat = path.stat()
                    except OSError:
                        continue
                    relative = str(path.relative_to(self.root)).replace("/", "\\")
                    top = relative.split("\\", 1)[0]
                    category, reason = self.classify(relative, path.suffix.lower())
                    items.append(FileItem(
                        path=relative,
                        name=name,
                        folder=top,
                        extension=path.suffix.lower() or "(không có)",
                        size=stat.st_size,
                        modified=stat.st_mtime,
                        category=category,
                        reason=reason,
                    ))
                    bucket = folder_stats.setdefault(top, {"name": top, "files": 0, "size": 0})
                    bucket["files"] += 1
                    bucket["size"] += stat.st_size
            with self.lock:
                self.files = items
                self.folders = sorted(folder_stats.values(), key=lambda row: row["size"], reverse=True)
                self.scanned_at = time.time()
                self.duration = self.scanned_at - started
                self.error = ""
        except Exception as exc:
            self.error = str(exc)
        finally:
            self.scanning = False

    def summary(self) -> dict:
        with self.lock:
            categories: dict[str, dict] = {}
            extensions: dict[str, int] = {}
            for item in self.files:
                bucket = categories.setdefault(item.category, {"files": 0, "size": 0})
                bucket["files"] += 1
                bucket["size"] += item.size
                extensions[item.extension] = extensions.get(item.extension, 0) + 1
            largest = sorted(self.files, key=lambda item: item.size, reverse=True)[:8]
            return {
                "root": str(self.root),
                "exists": self.root.exists(),
                "files": len(self.files),
                "size": sum(item.size for item in self.files),
                "folders": self.folders,
                "categories": categories,
                "largest": [asdict(item) for item in largest],
                "extensions": sorted(
                    ({"name": key, "count": value} for key, value in extensions.items()),
                    key=lambda row: row["count"],
                    reverse=True,
                )[:10],
                "scannedAt": self.scanned_at,
                "duration": self.duration,
                "error": self.error,
                "mode": "Chỉ đọc",
                "scanning": self.scanning,
                "excluded": sorted(EXCLUDED_TOP_LEVEL),
            }

    def search(self, query: str, category: str, limit: int = 100) -> list[dict]:
        query = query.casefold().strip()
        with self.lock:
            result = []
            for item in self.files:
                if category and category != "Tất cả" and item.category != category:
                    continue
                if query and query not in item.path.casefold():
                    continue
                result.append(asdict(item))
                if len(result) >= limit:
                    break
            return result


INDEX = RepositoryIndex(DEFAULT_REPOSITORY)


def path_state(relative: str) -> dict:
    target = INDEX.root / Path(relative)
    return {
        "path": relative,
        "exists": target.exists(),
        "size": target.stat().st_size if target.exists() and target.is_file() else 0,
    }


def catalog() -> list[dict]:
    return [
        {
            "id": "control",
            "name": "Điều khiển bằng hotkey",
            "description": "Điểm vào điều khiển Word, Excel và các công cụ đăng bài.",
            "color": "violet",
            "primary": path_state(r"01_hotkey\dangbaitag.ahk"),
            "tools": [],
            "steps": ["Mở công cụ", "Điều phối thao tác", "Gọi bước viết/đăng"],
        },
        {
            "id": "prepare-writing",
            "name": "Chuẩn bị dữ liệu viết bài",
            "description": "Đưa kế hoạch và cấu hình sang vùng dữ liệu sẵn sàng viết.",
            "color": "cyan",
            "primary": path_state(r"02_viet_bai\chuan_bi_du_lieu_viet_bai_ten_mien.py"),
            "tools": [],
            "steps": ["Đọc kế hoạch", "Kiểm tra cấu hình", "Tạo hàng chờ"],
        },
        {
            "id": "writing",
            "name": "Viết bài & tạo ảnh",
            "description": "Flow nhiều luồng dùng ChatGPT, Gemini và Word.",
            "color": "blue",
            "primary": path_state(r"02_viet_bai\TUDONG-CHAY_2_HOAC_3_LUONG_NO_MOUSE.py"),
            "tools": [
                path_state(r"02_viet_bai\viet_lai_word_dot_2.py"),
                path_state(r"02_viet_bai\TUDONG_3_LUONG_WORD_BRIEF_VA_ANH.py"),
            ],
            "steps": ["Lấy bài", "Viết nội dung", "Lưu Word", "Tạo brief", "Tạo ảnh", "Ghi kết quả"],
        },
        {
            "id": "prepare-publishing",
            "name": "Chuẩn bị đăng bài",
            "description": "Chuẩn bị dữ liệu Word, ảnh, danh mục và tác giả cho CMS.",
            "color": "amber",
            "primary": path_state(r"03_dang_bai\chuan_bi_du_lieu_dang_bai_moi_cate_post.py"),
            "tools": [],
            "steps": ["Chọn bài", "Chọn Word", "Ghép ảnh", "Kiểm tra dữ liệu"],
        },
        {
            "id": "publishing",
            "name": "Đăng bài",
            "description": "Flow tự động nhiều luồng và bộ công cụ cứu hộ theo từng bước.",
            "color": "green",
            "primary": path_state(r"03_dang_bai\VIP_tudongdangbai_3_5_luong_TEST.py"),
            "tools": [
                path_state(r"03_dang_bai\VIP_CHAY_1_BAI_DUNG_TRUOC_LUU.py"),
                path_state(r"03_dang_bai\phu_tro\mofileword_trongexcel.py"),
                path_state(r"03_dang_bai\phu_tro\mo_url.py"),
                path_state(r"03_dang_bai\phu_tro\chinhdungchuanv2.py"),
                path_state(r"03_dang_bai\phu_tro\xuatID_URL.py"),
                path_state(r"03_dang_bai\phu_tro\Vip_baivietlienquan.py"),
            ],
            "steps": ["Mở Word", "Mở CMS", "Điền nội dung", "Chọn danh mục", "Tải ảnh", "Lưu", "Lấy ID/URL"],
        },
        {
            "id": "workbook",
            "name": "Trung tâm dữ liệu Excel",
            "description": "Kế hoạch, hàng chờ viết bài, đăng bài và cấu hình website.",
            "color": "lime",
            "primary": path_state(r"04_excel\hotkeyvip_test.xlsm"),
            "tools": [path_state(r"04_excel\nhap_du_lieu_ke_hoach_tu_thu_muc.py")],
            "steps": ["Kế hoạch", "Viết bài", "Đăng bài", "Cấu hình"],
        },
    ]


def risk_report() -> list[dict]:
    risks = []
    hardcoded_count = 0
    missing_refs = []
    ahk = INDEX.root / r"01_hotkey\dangbaitag.ahk"
    if ahk.exists():
        text = ahk.read_text(encoding="utf-8", errors="ignore")
        hardcoded_count = len(re.findall(r"[A-Za-z]:\\[^\"\r\n]+", text))
        for match in re.findall(r'"([A-Za-z]:\\[^"]+)"', text):
            if not Path(match).exists() and match not in missing_refs:
                missing_refs.append(match)
    risks.append({
        "level": "high",
        "title": "Đường dẫn đang buộc vào ổ D:",
        "detail": f"Phát hiện khoảng {hardcoded_count} tham chiếu đường dẫn tuyệt đối trong hotkey. Đây là trở ngại chính khi chuyển máy.",
        "action": "Đưa đường dẫn về một sổ địa chỉ trung tâm.",
    })
    risks.append({
        "level": "medium",
        "title": "Điểm chạy chính chưa thống nhất",
        "detail": "Một số tài liệu, hotkey và file đang dùng gọi các tên phiên bản khác nhau.",
        "action": "Tạo điểm vào ổn định cho từng chức năng trước khi cho phép cập nhật thật.",
    })
    risks.append({
        "level": "medium",
        "title": "Dữ liệu trình duyệt đang lặp theo worker",
        "detail": "Các profile và cache chiếm phần lớn dung lượng kho; nhiều thành phần có thể tải lại.",
        "action": "Tách dữ liệu đăng nhập cần giữ khỏi cache có thể tạo lại.",
    })
    if missing_refs:
        risks.append({
            "level": "high",
            "title": f"Có ít nhất {len(missing_refs)} đường dẫn được gọi nhưng không tồn tại",
            "detail": "Flow có thể đang phụ thuộc vào tên cũ hoặc file đã được chuyển vị trí.",
            "action": "Xác nhận file thay thế trước khi đóng gói chuyển máy.",
        })
    return risks


class Handler(BaseHTTPRequestHandler):
    server_version = "HotkeyVIP Studio"

    def log_message(self, *_args) -> None:
        return

    def json_response(self, data, status=HTTPStatus.OK):
        payload = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(payload)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == "/api/summary":
            return self.json_response(INDEX.summary())
        if parsed.path == "/api/catalog":
            return self.json_response(catalog())
        if parsed.path == "/api/risks":
            return self.json_response(risk_report())
        if parsed.path == "/api/files":
            params = parse_qs(parsed.query)
            return self.json_response(INDEX.search(
                params.get("q", [""])[0],
                params.get("category", ["Tất cả"])[0],
                min(int(params.get("limit", ["100"])[0]), 250),
            ))
        if parsed.path == "/api/rescan":
            INDEX.scan()
            return self.json_response(INDEX.summary())
        if parsed.path == "/api/health":
            return self.json_response({"ok": True, "repository": str(INDEX.root), "readOnly": True})
        if parsed.path == "/api/launchers":
            return self.json_response([launcher_view(item) for item in load_launchers()])
        if parsed.path == "/api/launchers/history":
            launcher_id = parse_qs(parsed.query).get("id", [""])[0]
            safe_id = re.sub(r"[^A-Za-z0-9_.-]+", "_", launcher_id or "unknown")
            log_dir = LOGS_DIR / safe_id
            rows = []
            status_files = sorted(
                log_dir.glob("*.json"),
                key=lambda file: file.stat().st_mtime,
                reverse=True,
            )[:10] if log_dir.exists() else []
            for status_file in status_files:
                try:
                    row = json.loads(status_file.read_text(encoding="utf-8"))
                    row["name"] = status_file.stem
                    rows.append(row)
                except Exception:
                    continue
            return self.json_response(rows)
        if parsed.path == "/api/launchers/log":
            params = parse_qs(parsed.query)
            launcher_id = params.get("id", [""])[0]
            run_name = params.get("name", [""])[0].strip()
            if not run_name or not re.fullmatch(r"[A-Za-z0-9_.-]+", run_name):
                return self.json_response({"error": "Tên lượt chạy không hợp lệ"}, HTTPStatus.BAD_REQUEST)
            safe_id = re.sub(r"[^A-Za-z0-9_.-]+", "_", launcher_id or "unknown")
            log_path = LOGS_DIR / safe_id / f"{run_name}.log"
            if not log_path.is_file():
                return self.json_response({"error": "Không tìm thấy nội dung log"}, HTTPStatus.NOT_FOUND)
            try:
                content = log_path.read_text(encoding="utf-8", errors="replace")
                stat = log_path.stat()
            except Exception as exc:
                return self.json_response(
                    {"error": f"Không đọc được log: {exc}"},
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                )
            return self.json_response({
                "name": run_name,
                "content": content,
                "size": stat.st_size,
                "modifiedAt": stat.st_mtime,
            })
        if parsed.path == "/api/launcher-groups":
            return self.json_response(load_launcher_groups())
        if parsed.path == "/api/voice/status":
            config = load_voice_control_config()
            launcher = Path(str(config.get("launcherPath", "")))
            return self.json_response({
                "running": voice_control_is_ready(),
                "pid": voice_control_pid(),
                "launcherExists": launcher.is_file(),
                "launcherPath": str(launcher),
                "url": str(config.get("url", f"http://127.0.0.1:{VOICE_CONTROL_PORT}/voice")),
            })
        if parsed.path == "/api/voice/pick-launcher":
            try:
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                selected = filedialog.askopenfilename(
                    title="Chọn file chạy Voice Control",
                    filetypes=[
                        ("File chạy", "*.bat *.cmd *.py *.pyw *.exe"),
                        ("Batch", "*.bat *.cmd"),
                        ("Python", "*.py *.pyw"),
                        ("Ứng dụng", "*.exe"),
                        ("Tất cả file", "*.*"),
                    ],
                )
                root.destroy()
                return self.json_response({"path": selected or ""})
            except Exception as exc:
                return self.json_response({"error": f"Không mở được hộp chọn file: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)
        if parsed.path == "/api/pick-python-file":
            try:
                import tkinter as tk
                from tkinter import filedialog
                root = tk.Tk()
                root.withdraw()
                root.attributes("-topmost", True)
                selected = filedialog.askopenfilename(
                    title="Chọn file để chạy hoặc mở nhanh",
                    filetypes=[
                        ("File được hỗ trợ", "*.py *.pyw *.ahk *.ini *.xlsx *.xlsm *.xls *.docx *.doc"),
                        ("Python", "*.py *.pyw"), ("AutoHotkey", "*.ahk"),
                        ("Excel", "*.xlsx *.xlsm *.xls"), ("Word", "*.docx *.doc"),
                        ("Cấu hình INI", "*.ini"), ("Tất cả file", "*.*"),
                    ],
                )
                root.destroy()
                return self.json_response({"path": selected or ""})
            except Exception as exc:
                return self.json_response(
                    {"error": f"Không mở được hộp chọn file: {exc}"},
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                )
        return self.serve_static(parsed.path)

    def do_POST(self):
        parsed = urlparse(self.path)
        length = int(self.headers.get("Content-Length", "0"))
        body = self.rfile.read(length) if length else b"{}"
        try:
            data = json.loads(body.decode("utf-8"))
        except Exception:
            data = {}
        if parsed.path == "/api/restart":
            self.json_response({"ok": True, "message": "Studio dang tu cap nhat"})
            self.server.restart_requested = True
            threading.Thread(target=self.server.shutdown, daemon=True).start()
            return
        if parsed.path == "/api/submit-profiles/open":
            if not SUBMIT_PROFILE_MANAGER.is_file():
                return self.json_response(
                    {"error": "Không tìm thấy QUAN_LY_PROFILE_SUBMIT.bat"},
                    HTTPStatus.NOT_FOUND,
                )
            try:
                flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
                subprocess.Popen(
                    ["cmd.exe", "/c", str(SUBMIT_PROFILE_MANAGER)],
                    cwd=str(SUBMIT_PROFILE_MANAGER.parent),
                    creationflags=flags,
                )
            except Exception as exc:
                return self.json_response(
                    {"error": f"Không mở được quản lý profile Submit: {exc}"},
                    HTTPStatus.INTERNAL_SERVER_ERROR,
                )
            return self.json_response({"ok": True})
        if parsed.path == "/api/voice/start":
            global VOICE_PROCESS
            config = load_voice_control_config()
            launcher = Path(str(config.get("launcherPath", "")))
            if voice_control_is_ready():
                return self.json_response({"ok": True, "running": True, "url": config["url"]})
            if not launcher.is_file():
                return self.json_response({"error": "Không tìm thấy file chạy Voice Control"}, HTTPStatus.BAD_REQUEST)
            try:
                flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
                if launcher.suffix.lower() in {".bat", ".cmd"}:
                    command = ["cmd.exe", "/c", str(launcher), "--studio"]
                elif launcher.suffix.lower() in {".py", ".pyw"}:
                    executable = Path(sys.executable).with_name("python.exe")
                    if not executable.exists():
                        executable = Path(sys.executable)
                    command = [
                        str(executable),
                        str(launcher),
                        "--port", str(VOICE_CONTROL_PORT),
                        "--no-browser",
                    ]
                else:
                    command = [str(launcher)]
                VOICE_PROCESS = subprocess.Popen(
                    command,
                    cwd=str(launcher.parent),
                    creationflags=flags,
                )
            except Exception as exc:
                return self.json_response({"error": f"Không khởi động được Voice Control: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return self.json_response({"ok": True, "running": False, "url": config["url"]})
        if parsed.path == "/api/voice/config":
            launcher_path = str(data.get("launcherPath", "")).strip().strip('"')
            url = str(data.get("url", "")).strip()
            if not launcher_path:
                return self.json_response({"error": "Đường dẫn file chạy không được để trống"}, HTTPStatus.BAD_REQUEST)
            if not Path(launcher_path).is_file():
                return self.json_response({"error": "File chạy không tồn tại"}, HTTPStatus.BAD_REQUEST)
            if not re.match(r"^https?://", url, re.IGNORECASE):
                return self.json_response({"error": "Địa chỉ giao diện phải bắt đầu bằng http:// hoặc https://"}, HTTPStatus.BAD_REQUEST)
            payload = {"launcherPath": launcher_path, "url": url}
            temporary = VOICE_CONTROL_CONFIG_FILE.with_suffix(".json.tmp")
            temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
            temporary.replace(VOICE_CONTROL_CONFIG_FILE)
            return self.json_response({"ok": True, **payload, "restartRequired": voice_control_is_ready()})
        if parsed.path == "/api/voice/overlay":
            action = str(data.get("action", "")).strip().lower()
            if action not in {"hide", "show"}:
                return self.json_response({"error": "Action phải là hide hoặc show"}, HTTPStatus.BAD_REQUEST)
            if not voice_control_is_ready():
                return self.json_response({"error": "Voice Control chưa chạy"}, HTTPStatus.BAD_REQUEST)
            try:
                changed = set_voice_overlay_visible(action == "show")
            except Exception as exc:
                return self.json_response({"error": f"Không đổi được bảng nổi: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)
            if not changed:
                return self.json_response({"error": "Không tìm thấy cửa sổ nổi Voice Control"}, HTTPStatus.NOT_FOUND)
            return self.json_response({"ok": True, "action": action})
        if parsed.path == "/api/voice/stop":
            pid = voice_control_pid()
            if pid is None and VOICE_PROCESS is not None and VOICE_PROCESS.poll() is None:
                pid = VOICE_PROCESS.pid
            if pid is None:
                VOICE_PROCESS = None
                return self.json_response({"ok": True, "alreadyStopped": True})
            try:
                if os.name == "nt":
                    subprocess.run(
                        ["taskkill", "/PID", str(pid), "/T", "/F"],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.DEVNULL,
                        check=False,
                    )
                else:
                    VOICE_PROCESS.terminate()
                VOICE_PROCESS = None
            except Exception as exc:
                return self.json_response({"error": f"Không dừng được Voice Control: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)
            for _ in range(20):
                if not voice_control_is_ready():
                    return self.json_response({"ok": True})
                time.sleep(0.1)
            return self.json_response({"error": "Voice Control vẫn chưa tắt hoàn toàn"}, HTTPStatus.INTERNAL_SERVER_ERROR)
        if parsed.path == "/api/update-preview":
            name = Path(str(data.get("name", "file_moi.py"))).name
            lower = name.lower()
            if "dang" in lower or "vip_" in lower:
                target = "Đăng bài"
            elif "viet" in lower or "word" in lower or "gemini" in lower:
                target = "Viết bài & tạo ảnh"
            elif name.endswith(".ahk"):
                target = "Điều khiển bằng hotkey"
            elif name.endswith((".xlsm", ".xlsx")):
                target = "Trung tâm dữ liệu Excel"
            else:
                target = "Chưa xác định"
            return self.json_response({
                "safe": True,
                "mode": "Xem trước — không lưu file",
                "name": name,
                "size": int(data.get("size", 0)),
                "suggestedTarget": target,
                "checks": [
                    {"label": "Bản đang chạy sẽ được giữ lại", "state": "pass"},
                    {"label": "Tên điểm chạy bên ngoài không thay đổi", "state": "pass"},
                    {"label": "Cần xác nhận file đi kèm", "state": "warn"},
                    {"label": "Cần chạy thử trước khi công nhận", "state": "warn"},
                ],
            })
        if parsed.path == "/api/package-preview":
            mode = data.get("mode", "core")
            summary = INDEX.summary()
            core_folders = {"01_hotkey", "02_viet_bai", "03_dang_bai", "04_excel", "05_cau_hinh"}
            core_size = sum(row["size"] for row in summary["folders"] if row["name"] in core_folders)
            runtime_size = sum(row["size"] for row in summary["folders"] if row["name"] == "06_du_lieu_chay")
            include = core_size
            if mode == "full":
                include += runtime_size
            return self.json_response({
                "mode": mode,
                "includeSize": include,
                "excludedSize": max(summary["size"] - include, 0),
                "readOnly": True,
                "items": [
                    {"name": "Chương trình và hotkey", "included": True},
                    {"name": "Workbook và cấu hình", "included": True},
                    {"name": "Kho bài viết 07_ket_qua (được bảo vệ, không quét)", "included": False},
                    {"name": "Profile trình duyệt", "included": mode == "full"},
                    {"name": "Cache, log và kết quả test", "included": False},
                ],
            })
        if parsed.path == "/api/launchers/save":
            items = load_launchers()
            launcher_id = str(data.get("id") or uuid.uuid4().hex)
            name = str(data.get("name", "")).strip()
            path = str(data.get("path", "")).strip().strip('"')
            description = str(data.get("description", "")).strip()
            if not name:
                return self.json_response({"error": "Tên nút không được để trống"}, HTTPStatus.BAD_REQUEST)
            if not path:
                return self.json_response({"error": "Đường dẫn không được để trống"}, HTTPStatus.BAD_REQUEST)
            if Path(path).suffix.lower() not in LAUNCHER_EXTENSIONS:
                return self.json_response({"error": "Loại file này chưa được hỗ trợ"}, HTTPStatus.BAD_REQUEST)
            item = {
                "id": launcher_id,
                "name": name,
                "path": path,
                "description": description,
                "showConsole": bool(data.get("showConsole", True)),
                "group": str(data.get("group", "Chưa phân nhóm")).strip() or "Chưa phân nhóm",
                "order": int(data.get("order", 9999)),
            }
            existing = next((index for index, row in enumerate(items) if row.get("id") == launcher_id), None)
            if existing is None:
                item["favorite"] = bool(data.get("favorite", False))
                items.append(item)
            else:
                item["favorite"] = bool(items[existing].get("favorite", False))
                items[existing] = item
            save_launchers(items)
            groups = load_launcher_groups()
            if item["group"] not in groups:
                groups.append(item["group"])
                save_launcher_groups(groups)
            return self.json_response(launcher_view(item))
        if parsed.path == "/api/launcher-groups/add":
            name = str(data.get("name", "")).strip()
            if not name:
                return self.json_response({"error": "Tên nhóm không được để trống"}, HTTPStatus.BAD_REQUEST)
            groups = load_launcher_groups()
            if name in groups:
                return self.json_response({"error": "Tên nhóm đã tồn tại"}, HTTPStatus.BAD_REQUEST)
            groups.append(name)
            save_launcher_groups(groups)
            return self.json_response({"ok": True, "name": name})
        if parsed.path == "/api/launcher-groups/delete":
            name = str(data.get("name", "")).strip()
            if not name or name == "Chưa phân nhóm":
                return self.json_response({"error": "Khong the xoa nhom nay"}, HTTPStatus.BAD_REQUEST)
            groups = [group for group in load_launcher_groups() if group != name]
            if "Chưa phân nhóm" not in groups:
                groups.append("Chưa phân nhóm")
            items = load_launchers()
            for item in items:
                if item.get("group") == name:
                    item["group"] = "Chưa phân nhóm"
                    item["order"] = 9999
            save_launchers(items)
            save_launcher_groups(groups)
            return self.json_response({"ok": True})
        if parsed.path == "/api/launcher-groups/reorder":
            requested = data.get("groups", [])
            if not isinstance(requested, list):
                return self.json_response({"error": "Danh sach nhom khong hop le"}, HTTPStatus.BAD_REQUEST)
            current = load_launcher_groups()
            ordered = []
            for name in requested:
                name = str(name).strip()
                if name in current and name not in ordered:
                    ordered.append(name)
            ordered.extend(name for name in current if name not in ordered)
            save_launcher_groups(ordered)
            return self.json_response({"ok": True, "groups": ordered})
        if parsed.path == "/api/launchers/reorder":
            updates = data.get("items", [])
            if not isinstance(updates, list):
                return self.json_response({"error": "Danh sách thứ tự không hợp lệ"}, HTTPStatus.BAD_REQUEST)
            items = load_launchers()
            positions = {
                str(row.get("id")): {
                    "group": str(row.get("group", "")).strip(),
                    "order": int(row.get("order", 9999)),
                }
                for row in updates if row.get("id")
            }
            for item in items:
                position = positions.get(str(item.get("id")))
                if position:
                    item["group"] = position["group"] or "Chưa phân nhóm"
                    item["order"] = position["order"]
            save_launchers(items)
            return self.json_response({"ok": True})
        if parsed.path == "/api/launchers/favorite":
            launcher_id = str(data.get("id", ""))
            items = load_launchers()
            item = next((row for row in items if row.get("id") == launcher_id), None)
            if not item:
                return self.json_response({"error": "Khong tim thay nut"}, HTTPStatus.NOT_FOUND)
            item["favorite"] = bool(data.get("favorite", True))
            save_launchers(items)
            return self.json_response({"ok": True, "favorite": item["favorite"]})
        if parsed.path == "/api/launchers/favorite-reorder":
            updates = data.get("items", [])
            if not isinstance(updates, list):
                return self.json_response({"error": "Danh sach thu tu khong hop le"}, HTTPStatus.BAD_REQUEST)
            positions = {
                str(row.get("id")): int(row.get("order", 9999))
                for row in updates if row.get("id")
            }
            items = load_launchers()
            for item in items:
                if str(item.get("id")) in positions:
                    item["favoriteOrder"] = positions[str(item.get("id"))]
            save_launchers(items)
            return self.json_response({"ok": True})
        if parsed.path == "/api/launchers/rename-group":
            old_name = str(data.get("oldName", "")).strip()
            new_name = str(data.get("newName", "")).strip()
            if not old_name or not new_name:
                return self.json_response({"error": "Tên nhóm không được để trống"}, HTTPStatus.BAD_REQUEST)
            items = load_launchers()
            if not any(item.get("group") == old_name for item in items):
                return self.json_response({"error": "Không tìm thấy nhóm"}, HTTPStatus.NOT_FOUND)
            for item in items:
                if item.get("group") == old_name:
                    item["group"] = new_name
            save_launchers(items)
            groups = [new_name if name == old_name else name for name in load_launcher_groups()]
            if new_name not in groups:
                groups.append(new_name)
            save_launcher_groups(groups)
            return self.json_response({"ok": True, "name": new_name})
        if parsed.path == "/api/launchers/delete":
            launcher_id = str(data.get("id", ""))
            items = load_launchers()
            updated = [item for item in items if item.get("id") != launcher_id]
            if len(items) == len(updated):
                return self.json_response({"error": "Không tìm thấy nút"}, HTTPStatus.NOT_FOUND)
            save_launchers(updated)
            return self.json_response({"ok": True})
        if parsed.path == "/api/launchers/run":
            launcher_id = str(data.get("id", ""))
            item = next((row for row in load_launchers() if row.get("id") == launcher_id), None)
            if not item:
                return self.json_response({"error": "Không tìm thấy nút"}, HTTPStatus.NOT_FOUND)
            path = Path(str(item.get("path", "")))
            if not path.is_file():
                return self.json_response({"error": "File không tồn tại tại đường dẫn đã lưu"}, HTTPStatus.BAD_REQUEST)
            if path.suffix.lower() not in LAUNCHER_EXTENSIONS:
                return self.json_response({"error": "Loại file này chưa được hỗ trợ"}, HTTPStatus.BAD_REQUEST)
            if path.suffix.lower() not in {".py", ".pyw"}:
                try:
                    os.startfile(str(path))
                except Exception as exc:
                    return self.json_response({"error": f"Windows không mở được file: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)
                return self.json_response({
                    "ok": True, "pid": None, "name": item["name"], "path": str(path),
                    "showConsole": False, "launchMode": "windows-default",
                })
            flags = 0
            executable = Path(sys.executable).with_name("python.exe")
            if not executable.exists():
                executable = Path(sys.executable)
            show_console = bool(item.get("showConsole", True))
            if os.name == "nt":
                flags = subprocess.CREATE_NEW_CONSOLE if show_console else subprocess.CREATE_NO_WINDOW
            safe_id = re.sub(r"[^A-Za-z0-9_.-]+", "_", launcher_id or "unknown")
            command = [
                str(executable),
                str(APP_DIR / "launcher_worker.py"),
                "--target", str(path),
                "--log-dir", str(LOGS_DIR / safe_id),
            ]
            if show_console:
                command.append("--show-console")
            try:
                process = subprocess.Popen(
                    command,
                    cwd=str(APP_DIR),
                    creationflags=flags,
                )
            except Exception as exc:
                return self.json_response({"error": f"Không chạy được file: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)
            with RUNNING_LAUNCHERS_LOCK:
                RUNNING_LAUNCHERS.setdefault(launcher_id, []).append({
                    "process": process,
                    "startedAt": time.time(),
                })
            return self.json_response({
                "ok": True,
                "pid": process.pid,
                "name": item["name"],
                "path": str(path),
                "showConsole": show_console,
                "launchMode": "studio-logged",
                "logged": True,
            })
        if parsed.path == "/api/launchers/stop":
            launcher_id = str(data.get("id", ""))
            with RUNNING_LAUNCHERS_LOCK:
                active = [
                    entry for entry in RUNNING_LAUNCHERS.get(launcher_id, [])
                    if entry["process"].poll() is None
                ]
            if not active:
                return self.json_response({"error": "Nút này không có tiến trình đang chạy"}, HTTPStatus.NOT_FOUND)
            for entry in active:
                pid = entry["process"].pid
                if os.name == "nt":
                    subprocess.run(
                        ["taskkill", "/PID", str(pid), "/T", "/F"],
                        capture_output=True,
                        creationflags=subprocess.CREATE_NO_WINDOW,
                    )
                else:
                    entry["process"].terminate()
            with RUNNING_LAUNCHERS_LOCK:
                RUNNING_LAUNCHERS[launcher_id] = []
            safe_id = re.sub(r"[^A-Za-z0-9_.-]+", "_", launcher_id or "unknown")
            log_dir = LOGS_DIR / safe_id
            for status_file in log_dir.glob("*.json") if log_dir.exists() else []:
                try:
                    payload = json.loads(status_file.read_text(encoding="utf-8"))
                    if payload.get("state") != "running":
                        continue
                    payload["state"] = "stopped"
                    payload["endedAt"] = time.time()
                    payload["duration"] = round(time.time() - float(payload.get("startedAt", time.time())), 2)
                    payload["error"] = "Người dùng dừng từ Studio"
                    status_file.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
                except Exception:
                    continue
            return self.json_response({"ok": True, "stopped": len(active)})
        if parsed.path == "/api/launchers/open-log":
            launcher_id = str(data.get("id", ""))
            item = next((row for row in load_launchers() if str(row.get("id")) == launcher_id), None)
            if not item:
                return self.json_response({"error": "Không tìm thấy nút"}, HTTPStatus.NOT_FOUND)
            safe_id = re.sub(r"[^A-Za-z0-9_.-]+", "_", launcher_id or "unknown")
            log_dir = LOGS_DIR / safe_id
            run_name = str(data.get("name", "")).strip()
            if run_name:
                if not re.fullmatch(r"[A-Za-z0-9_.-]+", run_name):
                    return self.json_response({"error": "Tên lượt chạy không hợp lệ"}, HTTPStatus.BAD_REQUEST)
                requested_log = log_dir / f"{run_name}.log"
                logs = [requested_log] if requested_log.is_file() else []
            else:
                logs = sorted(log_dir.glob("*.log"), key=lambda file: file.stat().st_mtime, reverse=True) if log_dir.exists() else []
            if not logs:
                return self.json_response({"error": "Nút này chưa có log"}, HTTPStatus.NOT_FOUND)
            try:
                os.startfile(str(logs[0]))
            except Exception as exc:
                return self.json_response({"error": f"Không mở được log: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return self.json_response({"ok": True, "path": str(logs[0])})
        if parsed.path == "/api/open-containing-folder":
            path = Path(str(data.get("path", "")).strip().strip('"'))
            if not path.exists():
                return self.json_response({"error": "File không tồn tại"}, HTTPStatus.BAD_REQUEST)
            try:
                if os.name == "nt":
                    subprocess.Popen(["explorer.exe", "/select,", str(path)])
                else:
                    webbrowser.open(path.parent.as_uri())
            except Exception as exc:
                return self.json_response({"error": f"Không mở được thư mục: {exc}"}, HTTPStatus.INTERNAL_SERVER_ERROR)
            return self.json_response({"ok": True})
        return self.json_response({"error": "Không tìm thấy thao tác"}, HTTPStatus.NOT_FOUND)

    def serve_static(self, request_path: str):
        relative = "index.html" if request_path in ("", "/") else request_path.lstrip("/")
        target = (STATIC_DIR / relative).resolve()
        if STATIC_DIR.resolve() not in target.parents and target != STATIC_DIR.resolve():
            self.send_error(HTTPStatus.FORBIDDEN)
            return
        if not target.exists() or not target.is_file():
            target = STATIC_DIR / "index.html"
        payload = target.read_bytes()
        content_type = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", content_type + ("; charset=utf-8" if content_type.startswith("text/") else ""))
        self.send_header("Content-Length", str(len(payload)))
        self.end_headers()
        self.wfile.write(payload)


def main() -> None:
    global INDEX
    if len(sys.argv) > 1:
        candidate = Path(sys.argv[1]).resolve()
        if candidate.exists():
            INDEX = RepositoryIndex(candidate)
    server = ThreadingHTTPServer((HOST, PORT), Handler)
    server.restart_requested = False
    url = f"http://{HOST}:{PORT}"
    print(f"HotkeyVIP Studio: {url}")
    print(f"Repository: {INDEX.root}")
    print("Mode: READ ONLY")
    # Trình chạy không cần quét toàn kho lúc khởi động.
    # Các khu phân tích vẫn còn trong code và có thể quét khi được bật lại.
    if "--no-browser" not in sys.argv:
        threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()
        if server.restart_requested:
            subprocess.Popen(
                [sys.executable, str(Path(__file__).resolve()), "--no-browser"],
                cwd=str(APP_DIR),
            )


if __name__ == "__main__":
    main()
