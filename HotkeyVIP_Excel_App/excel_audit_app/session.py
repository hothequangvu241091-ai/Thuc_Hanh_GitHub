from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any


class SessionStore:
    def __init__(self) -> None:
        override = os.environ.get("EXCEL_AUDIT_APPDATA", "").strip()
        if override:
            root = Path(override)
        else:
            local_app_data = os.environ.get("LOCALAPPDATA")
            root = Path(local_app_data) if local_app_data else Path.home() / "AppData" / "Local"
            root = root / "ExcelAuditApp"
        self.root = root
        self.session_path = self.root / "session.json"

    def save(self, result: dict[str, Any]) -> None:
        self.root.mkdir(parents=True, exist_ok=True)
        temporary = self.session_path.with_suffix(".tmp")
        temporary.write_text(
            json.dumps(result, ensure_ascii=False, separators=(",", ":")),
            encoding="utf-8",
        )
        os.replace(temporary, self.session_path)

    def load(self) -> dict[str, Any] | None:
        if not self.session_path.exists():
            return None
        try:
            value = json.loads(self.session_path.read_text(encoding="utf-8"))
            return value if isinstance(value, dict) else None
        except (OSError, json.JSONDecodeError):
            return None

    def clear(self) -> None:
        self.session_path.unlink(missing_ok=True)
