from __future__ import annotations

import json
import os
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, X, Y, filedialog, messagebox
import tkinter as tk
from tkinter import ttk
from typing import Any

from .analysis import APP_VERSION, analyze_workbook
from .excel_io import file_fingerprint, same_fingerprint
from .flow_catalog import ACTIVE_FLOWS, flow_by_key
from .report_export import (
    ExportError,
    SourceChangedError,
    export_result,
    recover_dang_bai,
    suggested_output_path,
    suggested_recovery_path,
)
from .publish_plan import build_balanced_publish_plan, inspect_publish_queue
from .publish_review import build_retry_publish_plan
from .session import SessionStore
from .submit_transfer import (
    inspect_latest_published_urls,
    submit_launcher_path,
    transfer_latest_published_urls,
)
from .write_plan import build_write_queue_preview, inspect_write_queue


COLORS = {
    "navy": "#16324F",
    "teal": "#0F766E",
    "blue": "#315B7D",
    "bg": "#F4F7FA",
    "card": "#FFFFFF",
    "text": "#243047",
    "muted": "#66758A",
    "border": "#DCE4EC",
    "green": "#15803D",
    "red": "#B42318",
    "amber": "#B45309",
}

LEVEL_LABELS = {
    "error": "Lỗi dữ liệu",
    "recovery": "Cần khôi phục",
    "pending": "Chưa chuyển",
    "info": "Đã đăng",
}
LEVEL_FILTERS = {label: key for key, label in LEVEL_LABELS.items()}


def format_number(value: Any) -> str:
    if isinstance(value, bool):
        return str(value)
    if isinstance(value, int):
        return f"{value:,}".replace(",", ".")
    return str(value)


class ExcelAuditApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title(f"Đối soát nội dung Excel v{APP_VERSION}")
        self.geometry("1420x880")
        self.minsize(1120, 720)
        self.configure(bg=COLORS["bg"])
        self.option_add("*Font", "{Segoe UI} 10")

        self.store = SessionStore()
        self.result: dict[str, Any] | None = None
        self.selected_path = tk.StringVar()
        self.status_text = tk.StringVar(value="Sẵn sàng")
        self.filter_category = tk.StringVar(value="Tất cả")
        self.filter_level = tk.StringVar(value="Tất cả")
        self.filter_text = tk.StringVar()
        self.card_vars: dict[str, tk.StringVar] = {}
        self.write_dang_value_label: tk.Label | None = None
        self.write_dang_eyebrow_label: tk.Label | None = None
        self.write_dang_detail_label: tk.Label | None = None
        self.url_posted_value_label: tk.Label | None = None
        self.url_posted_eyebrow_label: tk.Label | None = None
        self.url_posted_detail_label: tk.Label | None = None
        self.health_title = tk.StringVar(value="CHƯA PHÂN TÍCH")
        self.health_detail = tk.StringVar(value="Chọn file Excel và bấm Phân tích")
        self.pipeline_text = tk.StringVar(value="—")
        self.recovery_count_text = tk.StringVar(value="0 dòng sẵn sàng sao chép")
        self.flow_status_text = tk.StringVar(value="Chưa có flow đang chạy")
        self.flow_buttons: dict[str, ttk.Button] = {}
        self._flow_process: subprocess.Popen[str] | None = None
        self._flow_running_key: str | None = None
        self._busy = False
        self.publish_review: dict[str, Any] = {"errors": [], "posted_today": [], "retry_rows": []}
        self.publish_review_items: dict[str, dict[str, Any]] = {}
        self.publish_id_updates: dict[int, str] = {}
        self.publish_retry_rows: set[int] = set()
        self.publish_detail_vars: dict[str, tk.StringVar] = {}
        self.publish_selected_item: dict[str, Any] | None = None

        self._configure_styles()
        self._build_layout()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self.after(150, self._restore_session)

    def _configure_styles(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("TFrame", background=COLORS["bg"])
        style.configure("Card.TFrame", background=COLORS["card"], relief="flat")
        style.configure(
            "TNotebook",
            background=COLORS["bg"],
            borderwidth=0,
            tabmargins=(0, 8, 0, 0),
        )
        style.configure(
            "TNotebook.Tab",
            background="#E7EDF3",
            foreground=COLORS["muted"],
            padding=(18, 10),
            font=("Segoe UI Semibold", 10),
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", COLORS["card"])],
            foreground=[("selected", COLORS["navy"])],
        )
        style.configure(
            "Treeview",
            background="#FFFFFF",
            fieldbackground="#FFFFFF",
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            rowheight=30,
        )
        style.configure(
            "Treeview.Heading",
            background=COLORS["blue"],
            foreground="#FFFFFF",
            font=("Segoe UI Semibold", 9),
            padding=(8, 8),
            relief="flat",
        )
        style.map("Treeview.Heading", background=[("active", COLORS["navy"])])
        style.configure(
            "Primary.TButton",
            background=COLORS["teal"],
            foreground="#FFFFFF",
            padding=(16, 9),
            font=("Segoe UI Semibold", 10),
            borderwidth=0,
        )
        style.map(
            "Primary.TButton",
            background=[("active", "#0B5F59"), ("disabled", "#A9B8C5")],
        )
        style.configure(
            "WriteFlow.TButton",
            background="#2563EB",
            foreground="#FFFFFF",
            padding=(16, 9),
            font=("Segoe UI Semibold", 10),
            borderwidth=0,
        )
        style.map(
            "WriteFlow.TButton",
            background=[("active", "#1D4ED8"), ("disabled", "#A9B8C5")],
        )
        style.configure(
            "PublishFlow.TButton",
            background="#D97706",
            foreground="#FFFFFF",
            padding=(16, 9),
            font=("Segoe UI Semibold", 10),
            borderwidth=0,
        )
        style.map(
            "PublishFlow.TButton",
            background=[("active", "#B45309"), ("disabled", "#A9B8C5")],
        )
        style.configure(
            "SubmitTransfer.TButton",
            background="#7C3AED",
            foreground="#FFFFFF",
            padding=(16, 9),
            font=("Segoe UI Semibold", 10),
            borderwidth=0,
        )
        style.map(
            "SubmitTransfer.TButton",
            background=[("active", "#6D28D9"), ("disabled", "#A9B8C5")],
        )
        style.configure(
            "Secondary.TButton",
            background="#E8EEF4",
            foreground=COLORS["navy"],
            padding=(14, 9),
            borderwidth=0,
        )
        style.map("Secondary.TButton", background=[("active", "#DCE6EF")])
        style.configure(
            "Danger.TButton",
            background="#FEE2E2",
            foreground=COLORS["red"],
            padding=(12, 9),
            borderwidth=0,
        )
        style.configure("TCombobox", padding=6)
        style.configure("TEntry", padding=7)

    def _build_layout(self) -> None:
        header = tk.Frame(self, bg=COLORS["navy"], height=82)
        header.pack(fill=X)
        header.pack_propagate(False)
        tk.Label(
            header,
            text="ĐỐI SOÁT NỘI DUNG EXCEL",
            bg=COLORS["navy"],
            fg="#FFFFFF",
            font=("Segoe UI Semibold", 19),
        ).pack(anchor="w", padx=28, pady=(16, 0))
        tk.Label(
            header,
            text=f"KE_HOACH  •  VIET_BAI  •  DANG_BAI  •  v{APP_VERSION}",
            bg=COLORS["navy"],
            fg="#BFD0DF",
            font=("Segoe UI", 9),
        ).pack(anchor="w", padx=29, pady=(2, 0))

        toolbar = tk.Frame(self, bg=COLORS["card"], highlightthickness=1, highlightbackground=COLORS["border"])
        toolbar.pack(fill=X, padx=20, pady=(16, 8))
        tk.Label(
            toolbar,
            text="File Excel",
            bg=COLORS["card"],
            fg=COLORS["muted"],
            font=("Segoe UI Semibold", 9),
        ).pack(side=LEFT, padx=(16, 8), pady=14)
        path_entry = ttk.Entry(toolbar, textvariable=self.selected_path)
        path_entry.pack(side=LEFT, fill=X, expand=True, padx=(0, 8), pady=10)
        self.choose_button = ttk.Button(
            toolbar, text="Chọn file", command=self._choose_file, style="Secondary.TButton"
        )
        self.choose_button.pack(side=LEFT, padx=4, pady=8)
        self.analyze_button = ttk.Button(
            toolbar, text="Phân tích", command=self._start_analysis, style="Primary.TButton"
        )
        self.analyze_button.pack(side=LEFT, padx=4, pady=8)
        self.open_excel_button = ttk.Button(
            toolbar,
            text="Mở Excel",
            command=self._open_selected_excel,
            style="Secondary.TButton",
            state="disabled",
        )
        self.open_excel_button.pack(side=LEFT, padx=4, pady=8)
        self.export_button = ttk.Button(
            toolbar,
            text="Xuất kết quả",
            command=self._start_export,
            style="Primary.TButton",
            state="disabled",
        )
        self.export_button.pack(side=LEFT, padx=4, pady=8)
        self.clear_button = ttk.Button(
            toolbar, text="Xóa phiên", command=self._clear_session, style="Danger.TButton"
        )
        self.clear_button.pack(side=LEFT, padx=(4, 14), pady=8)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=BOTH, expand=True, padx=20, pady=(0, 8))

        self.overview_tab = ttk.Frame(self.notebook)
        self.ke_tab = ttk.Frame(self.notebook)
        self.viet_tab = ttk.Frame(self.notebook)
        self.dang_tab = ttk.Frame(self.notebook)
        self.issues_tab = ttk.Frame(self.notebook)
        self.recovery_tab = ttk.Frame(self.notebook)
        self.publish_review_tab = ttk.Frame(self.notebook)
        self.flows_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.overview_tab, text="Tổng quan")
        self.notebook.add(self.ke_tab, text="Kế hoạch")
        self.notebook.add(self.viet_tab, text="Viết bài")
        self.notebook.add(self.dang_tab, text="Đăng bài")
        self.notebook.add(self.issues_tab, text="Chi tiết đối soát")
        self.notebook.add(self.recovery_tab, text="Khôi phục DANG_BAI")
        self.notebook.add(self.publish_review_tab, text="Theo dõi đăng bài")
        self.notebook.add(self.flows_tab, text="Công việc")

        self._build_overview()
        self.ke_tree = self._build_summary_tree(
            self.ke_tab,
            [
                ("domain", "Tên miền", 180, "w"),
                ("total_rows", "Tổng bài", 90, "e"),
                ("combo4_complete", "Combo 4 đủ", 100, "e"),
                ("combo4_missing", "Combo 4 thiếu", 110, "e"),
                ("url_valid", "URL hợp lệ", 100, "e"),
                ("url_written", "Đã viết", 90, "e"),
                ("url_blank", "URL trống", 90, "e"),
                ("url_other", "URL sai/khác", 110, "e"),
                ("problem_rows", "Dữ liệu lỗi", 100, "e"),
                ("duplicate_groups", "Nhóm trùng", 100, "e"),
                ("duplicate_rows", "Dòng Combo 4 trùng", 135, "e"),
                ("missing_in_viet", "Thiếu trong VIET", 120, "e"),
            ],
        )
        self.viet_tree = self._build_summary_tree(
            self.viet_tab,
            [
                ("domain", "Tên miền", 180, "w"),
                ("total_rows", "Tổng dòng", 90, "e"),
                ("combo4_complete", "Combo 4 đủ", 100, "e"),
                ("combo4_missing", "Combo 4 thiếu", 110, "e"),
                ("completed_ok", "Hoàn tất OK", 105, "e"),
                ("completed_with_assets", "OK + đủ tài nguyên", 140, "e"),
                ("archived_posted_no_assets", "Đã đăng, đã xóa tài nguyên", 175, "e"),
                ("recovery_no_assets", "Cần khôi phục DANG", 145, "e"),
                ("unexplained_no_assets", "Thiếu tài nguyên bất thường", 175, "e"),
                ("not_completed", "Chưa hoàn tất", 110, "e"),
                ("duplicate_rows", "Dòng Combo 4 trùng", 135, "e"),
            ],
        )
        self.dang_tree = self._build_summary_tree(
            self.dang_tab,
            [
                ("domain", "Tên miền", 180, "w"),
                ("total_rows", "Tổng dòng", 90, "e"),
                ("combo4_complete", "Combo 4 đủ", 100, "e"),
                ("combo4_missing", "Combo 4 thiếu", 110, "e"),
                ("in_viet", "Có trong VIET", 110, "e"),
                ("posted", "Đã đăng", 90, "e"),
                ("url_not_posted_full_assets", "Có URL, chưa đăng, đủ tài nguyên", 210, "e"),
                ("dang_missing_viet", "DANG có - VIET thiếu", 160, "e"),
                ("classification_difference", "Chênh lệch", 100, "e"),
            ],
        )
        self._build_issues()
        self._build_recovery()
        self._build_publish_review()
        self._build_flows()

        status_bar = tk.Frame(self, bg="#E9EFF5", height=30)
        status_bar.pack(fill=X, side="bottom")
        status_bar.pack_propagate(False)
        self.status_label = tk.Label(
            status_bar,
            textvariable=self.status_text,
            bg="#E9EFF5",
            fg=COLORS["muted"],
            anchor="w",
            font=("Segoe UI", 9),
        )
        self.status_label.pack(fill=BOTH, padx=20)

    def _build_overview(self) -> None:
        container = ttk.Frame(self.overview_tab)
        container.pack(fill=BOTH, expand=True, padx=4, pady=8)

        self.health_frame = tk.Frame(
            container,
            bg="#E8EEF4",
            highlightthickness=1,
            highlightbackground=COLORS["border"],
        )
        self.health_frame.pack(fill=X, padx=5, pady=(2, 7))
        self.health_title_label = tk.Label(
            self.health_frame,
            textvariable=self.health_title,
            bg="#E8EEF4",
            fg=COLORS["navy"],
            font=("Segoe UI Semibold", 18),
        )
        self.health_title_label.pack(anchor="w", padx=18, pady=(12, 0))
        self.health_detail_label = tk.Label(
            self.health_frame,
            textvariable=self.health_detail,
            bg="#E8EEF4",
            fg=COLORS["muted"],
            font=("Segoe UI", 10),
        )
        self.health_detail_label.pack(anchor="w", padx=18, pady=(2, 12))

        cards = ttk.Frame(container)
        cards.pack(fill=X)
        definitions = [
            ("errors", "Lỗi dữ liệu", "Cần xử lý", COLORS["red"], "Lỗi dữ liệu"),
            ("recovery", "Khôi phục", "Cần thêm vào DANG_BAI", COLORS["amber"], "Cần khôi phục"),
            ("pending", "Tiến độ", "Chưa chuyển sang DANG_BAI", COLORS["blue"], "Chưa chuyển"),
            ("assets", "Viết bài", "OK + đủ Word và 2 ảnh", COLORS["green"], ""),
            ("url_posted_match", "URL HỢP LỆ ↔ ĐÃ ĐĂNG", "Đang đối chiếu", COLORS["teal"], ""),
            ("write_dang_match", "VIẾT OK ↔ ĐĂNG_BÀI", "Đang đối chiếu", COLORS["teal"], ""),
        ]
        for index, (key, eyebrow, label, color, detail_level) in enumerate(definitions):
            row, column = divmod(index, 3)
            cards.columnconfigure(column, weight=1, uniform="cards")
            frame = tk.Frame(
                cards,
                bg=COLORS["card"],
                highlightthickness=1,
                highlightbackground=COLORS["border"],
            )
            frame.grid(row=row, column=column, sticky="nsew", padx=5, pady=5)
            eyebrow_label = tk.Label(
                frame,
                text=eyebrow.upper(),
                bg=COLORS["card"],
                fg=color,
                font=("Segoe UI Semibold", 8),
            )
            eyebrow_label.pack(anchor="w", padx=14, pady=(12, 2))
            variable = tk.StringVar(value="—")
            self.card_vars[key] = variable
            value_label = tk.Label(
                frame,
                textvariable=variable,
                bg=COLORS["card"],
                fg=COLORS["navy"],
                font=("Segoe UI Semibold", 22),
            )
            value_label.pack(anchor="w", padx=14)
            detail_label = tk.Label(
                frame,
                text=label,
                bg=COLORS["card"],
                fg=COLORS["muted"],
                font=("Segoe UI", 9),
            )
            detail_label.pack(anchor="w", padx=14, pady=(0, 12))
            if key == "write_dang_match":
                self.write_dang_value_label = value_label
                self.write_dang_eyebrow_label = eyebrow_label
                self.write_dang_detail_label = detail_label
            elif key == "url_posted_match":
                self.url_posted_value_label = value_label
                self.url_posted_eyebrow_label = eyebrow_label
                self.url_posted_detail_label = detail_label
            if detail_level:
                for widget in (frame, eyebrow_label, value_label, detail_label):
                    widget.configure(cursor="hand2")
                    widget.bind(
                        "<Button-1>",
                        lambda _event, level=detail_level: self._show_detail_level(level),
                    )

        pipeline = tk.Frame(container, bg="#EAF2F8")
        pipeline.pack(fill=X, padx=5, pady=(7, 4))
        tk.Label(
            pipeline,
            text="PHÉP KIỂM TRA VIET_BAI",
            bg="#EAF2F8",
            fg=COLORS["blue"],
            font=("Segoe UI Semibold", 8),
        ).pack(anchor="w", padx=14, pady=(9, 1))
        tk.Label(
            pipeline,
            textvariable=self.pipeline_text,
            bg="#EAF2F8",
            fg=COLORS["navy"],
            font=("Segoe UI Semibold", 11),
        ).pack(anchor="w", padx=14, pady=(0, 9))

        title = tk.Label(
            container,
            text="Đối soát theo tên miền",
            bg=COLORS["bg"],
            fg=COLORS["navy"],
            font=("Segoe UI Semibold", 12),
        )
        title.pack(anchor="w", padx=6, pady=(16, 6))
        self.reconciliation_tree = self._build_tree(
            container,
            [
                ("domain", "Tên miền", 210, "w"),
                ("ke_total", "Tổng KE", 90, "e"),
                ("viet_total", "Tổng VIET", 95, "e"),
                ("ke_missing_viet", "KE có - VIET thiếu", 145, "e"),
                ("viet_missing_ke", "VIET có - KE thiếu", 145, "e"),
                ("in_dang", "Đã có trong DANG", 135, "e"),
                ("recovery_dang", "Cần khôi phục DANG", 155, "e"),
                ("pending_dang", "Chưa chuyển DANG", 135, "e"),
                ("dang_missing_viet", "DANG có - VIET thiếu", 155, "e"),
                ("viet_combo4_missing", "VIET thiếu Combo 4", 145, "e"),
                ("difference", "Chênh lệch", 105, "e"),
                ("status", "Trạng thái", 110, "center"),
            ],
        )
        self.reconciliation_tree.bind("<Double-1>", self._open_reconciliation_detail)

    def _build_summary_tree(
        self, parent: ttk.Frame, columns: list[tuple[str, str, int, str]]
    ) -> ttk.Treeview:
        wrapper = ttk.Frame(parent)
        wrapper.pack(fill=BOTH, expand=True, padx=8, pady=10)
        return self._build_tree(wrapper, columns)

    def _build_tree(
        self, parent: tk.Misc, columns: list[tuple[str, str, int, str]]
    ) -> ttk.Treeview:
        frame = ttk.Frame(parent)
        frame.pack(fill=BOTH, expand=True)
        identifiers = [item[0] for item in columns]
        tree = ttk.Treeview(frame, columns=identifiers, show="headings")
        for identifier, heading, width, anchor in columns:
            tree.heading(identifier, text=heading)
            display_anchor = "center" if anchor == "e" else anchor
            tree.column(
                identifier,
                width=width,
                minwidth=70,
                anchor=display_anchor,
                stretch=True,
            )
        tree.tag_configure("total", background="#DCEAF5", font=("Segoe UI Semibold", 10))
        tree.tag_configure("warning", foreground=COLORS["red"], font=("Segoe UI Semibold", 10))
        tree.tag_configure("error", background="#FEF2F2", foreground=COLORS["red"])
        tree.tag_configure("recovery", background="#FFF7ED", foreground=COLORS["amber"])
        tree.tag_configure("pending", background="#EFF6FF", foreground=COLORS["blue"])
        tree.tag_configure("info", background="#ECFDF5", foreground=COLORS["green"])
        vertical = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        horizontal = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vertical.set, xscrollcommand=horizontal.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vertical.grid(row=0, column=1, sticky="ns")
        horizontal.grid(row=1, column=0, sticky="ew")
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        return tree

    def _build_issues(self) -> None:
        toolbar = ttk.Frame(self.issues_tab)
        toolbar.pack(fill=X, padx=8, pady=(10, 4))
        tk.Label(
            toolbar,
            text="Nhóm",
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Segoe UI Semibold", 9),
        ).pack(side=LEFT, padx=(0, 6))
        self.level_combo = ttk.Combobox(
            toolbar,
            textvariable=self.filter_level,
            values=["Tất cả", "Lỗi dữ liệu", "Cần khôi phục", "Chưa chuyển", "Đã đăng"],
            state="readonly",
            width=18,
        )
        self.level_combo.pack(side=LEFT, padx=(0, 12))
        self.level_combo.bind("<<ComboboxSelected>>", lambda _event: self._render_issues())
        tk.Label(
            toolbar,
            text="Loại chi tiết",
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Segoe UI Semibold", 9),
        ).pack(side=LEFT, padx=(0, 6))
        self.category_combo = ttk.Combobox(
            toolbar,
            textvariable=self.filter_category,
            values=["Tất cả"],
            state="readonly",
            width=28,
        )
        self.category_combo.pack(side=LEFT, padx=(0, 12))
        self.category_combo.bind("<<ComboboxSelected>>", lambda _event: self._render_issues())
        tk.Label(
            toolbar,
            text="Tìm kiếm",
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Segoe UI Semibold", 9),
        ).pack(side=LEFT, padx=(0, 6))
        entry = ttk.Entry(toolbar, textvariable=self.filter_text, width=26)
        entry.pack(side=LEFT, padx=(0, 6))
        entry.bind("<Return>", lambda _event: self._render_issues())
        ttk.Button(
            toolbar, text="Lọc", command=self._render_issues, style="Secondary.TButton"
        ).pack(side=LEFT)
        self.issue_count_label = tk.Label(
            toolbar,
            text="0 dòng",
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Segoe UI Semibold", 9),
        )
        self.issue_count_label.pack(side=RIGHT)

        self.issue_tree = self._build_summary_tree(
            self.issues_tab,
            [
                ("level", "Nhóm", 105, "w"),
                ("category", "Loại chi tiết", 220, "w"),
                ("sheet", "Sheet nguồn", 105, "w"),
                ("row", "Dòng nguồn", 90, "e"),
                ("target_sheet", "Thiếu/đối chiếu ở", 125, "w"),
                ("target_row", "Dòng đối chiếu", 100, "e"),
                ("domain", "Tên miền", 180, "w"),
                ("title", "Tiêu đề SEO", 270, "w"),
                ("h1", "H1", 270, "w"),
                ("keyword", "Từ khóa/Tiêu đề", 220, "w"),
                ("detail", "Chi tiết", 420, "w"),
            ],
        )
        self.issue_tree.bind("<Double-1>", self._copy_issue)

    def _build_recovery(self) -> None:
        toolbar = ttk.Frame(self.recovery_tab)
        toolbar.pack(fill=X, padx=8, pady=(10, 4))
        tk.Label(
            toolbar,
            text="Có thể sao chép thủ công hoặc tạo một file mới và tự thêm các dòng vào DANG_BAI.",
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
        ).pack(side=LEFT)
        self.recover_button = ttk.Button(
            toolbar,
            text="Tạo file mới + khôi phục",
            command=self._start_recovery,
            style="Primary.TButton",
            state="disabled",
        )
        self.recover_button.pack(side=RIGHT, padx=(8, 0))
        ttk.Button(
            toolbar,
            text="Sao chép toàn bộ để dán vào DANG_BAI",
            command=self._copy_all_recovery,
            style="Secondary.TButton",
        ).pack(side=RIGHT, padx=(8, 0))
        tk.Label(
            toolbar,
            textvariable=self.recovery_count_text,
            bg=COLORS["bg"],
            fg=COLORS["amber"],
            font=("Segoe UI Semibold", 9),
        ).pack(side=RIGHT)

        self.recovery_tree = self._build_summary_tree(
            self.recovery_tab,
            [
                ("ke_row", "Dòng KE", 80, "e"),
                ("viet_row", "Dòng VIET", 85, "e"),
                ("domain", "Tên miền", 175, "w"),
                ("title", "Tiêu đề SEO", 285, "w"),
                ("h1", "H1", 285, "w"),
                ("keyword", "Từ khóa/Tiêu đề", 210, "w"),
                ("url", "URL đã đăng", 330, "w"),
            ],
        )
        self.recovery_tree.bind("<Double-1>", self._copy_one_recovery)

    def _build_publish_review(self) -> None:
        outer = ttk.Frame(self.publish_review_tab)
        outer.pack(fill=BOTH, expand=True, padx=8, pady=10)

        toolbar = ttk.Frame(outer)
        toolbar.pack(fill=X, pady=(0, 8))
        self.publish_review_count = tk.StringVar(value="Bấm Phân tích để đọc lỗi đăng bài")
        tk.Label(
            toolbar,
            textvariable=self.publish_review_count,
            bg=COLORS["bg"],
            fg=COLORS["navy"],
            font=("Segoe UI Semibold", 10),
        ).pack(side=LEFT)
        self.publish_review_action_button = ttk.Button(
            toolbar,
            text="Cập nhật ID & đăng lại LỖI KIỂM TRA",
            command=self._start_publish_review_action,
            style="PublishFlow.TButton",
            state="disabled",
        )
        self.publish_review_action_button.pack(side=RIGHT)

        content = ttk.Panedwindow(outer, orient="horizontal")
        content.pack(fill=BOTH, expand=True)
        list_frame = ttk.Frame(content)
        detail_frame = ttk.Frame(content, padding=(12, 4))
        content.add(list_frame, weight=3)
        content.add(detail_frame, weight=2)

        lists = ttk.Notebook(list_frame)
        lists.pack(fill=BOTH, expand=True)
        errors_tab = ttk.Frame(lists)
        posted_tab = ttk.Frame(lists)
        lists.add(errors_tab, text="Cần xử lý")
        lists.add(posted_tab, text="Đã đăng hôm nay")
        columns = [
            ("row", "Dòng", 65, "e"),
            ("status", "Trạng thái", 135, "w"),
            ("domain", "Tên miền", 165, "w"),
            ("keyword", "Tiêu đề", 240, "w"),
            ("cms_id", "ID CMS", 85, "w"),
            ("published_at", "Thời gian đăng", 145, "w"),
        ]
        self.publish_error_tree = self._build_summary_tree(errors_tab, columns)
        self.publish_today_tree = self._build_summary_tree(posted_tab, columns)
        self.publish_error_tree.bind("<<TreeviewSelect>>", self._select_publish_review_row)
        self.publish_today_tree.bind("<<TreeviewSelect>>", self._select_publish_review_row)

        tk.Label(
            detail_frame,
            text="THÔNG TIN BÀI ĐĂNG",
            bg=COLORS["bg"],
            fg=COLORS["navy"],
            font=("Segoe UI Semibold", 11),
        ).pack(anchor="w", pady=(0, 8))
        id_editor = ttk.Frame(detail_frame)
        id_editor.pack(fill=X, pady=(0, 10))
        tk.Label(
            id_editor,
            text="ID CMS mới",
            bg=COLORS["bg"],
            fg=COLORS["amber"],
            font=("Segoe UI Semibold", 9),
        ).pack(side=LEFT, padx=(0, 8))
        self.publish_new_id = tk.StringVar()
        ttk.Entry(id_editor, textvariable=self.publish_new_id, width=18).pack(side=LEFT)
        ttk.Button(
            id_editor,
            text="Ghi nhớ ID",
            command=self._stage_publish_id,
            style="Secondary.TButton",
        ).pack(side=LEFT, padx=(6, 0))
        ttk.Button(
            id_editor,
            text="Chuyển sang LỖI KIỂM TRA & chạy lại",
            command=self._retry_selected_publish_error,
            style="PublishFlow.TButton",
        ).pack(side=LEFT, padx=(6, 0))
        fields = [
            ("row", "Dòng Excel"),
            ("status", "Trạng thái"),
            ("domain", "Tên miền"),
            ("category", "Danh mục"),
            ("keyword", "Tiêu đề / từ khóa"),
            ("seo_title", "Tiêu đề SEO"),
            ("h1", "H1"),
            ("publish_error", "Lỗi đăng"),
            ("published_at", "Thời gian đăng"),
            ("published_url", "URL đã đăng"),
            ("related", "Bài viết liên quan"),
            ("word_path_resolved", "Đường dẫn Word"),
            ("chat_url", "URL ChatGPT"),
        ]
        form = ttk.Frame(detail_frame)
        form.pack(fill=BOTH, expand=True)
        form.columnconfigure(1, weight=1)
        for index, (key, label) in enumerate(fields):
            variable = tk.StringVar()
            self.publish_detail_vars[key] = variable
            tk.Label(
                form,
                text=label,
                bg=COLORS["bg"],
                fg=COLORS["muted"],
                font=("Segoe UI Semibold", 8),
            ).grid(row=index, column=0, sticky="w", padx=(0, 8), pady=2)
            entry = ttk.Entry(form, textvariable=variable, state="readonly")
            entry.grid(row=index, column=1, sticky="ew", pady=2)

        action_row = len(fields)
        ttk.Button(
            form, text="Mở Word", command=self._open_publish_word, style="Secondary.TButton"
        ).grid(row=action_row, column=0, sticky="ew", pady=(8, 3))
        url_buttons = ttk.Frame(form)
        url_buttons.grid(row=action_row, column=1, sticky="w", pady=(8, 3))
        ttk.Button(
            url_buttons, text="Mở ChatGPT", command=lambda: self._open_publish_url("chat_url"),
            style="Secondary.TButton",
        ).pack(side=LEFT)

    def _render_publish_review(self) -> None:
        for tree in (self.publish_error_tree, self.publish_today_tree):
            tree.delete(*tree.get_children())
        self.publish_review_items.clear()
        for tree, key in (
            (self.publish_error_tree, "errors"),
            (self.publish_today_tree, "posted_today"),
        ):
            for item in self.publish_review.get(key, []):
                item_id = tree.insert(
                    "",
                    END,
                    values=(
                        item.get("row", ""),
                        item.get("display_status", item.get("status", "")),
                        item.get("domain", ""),
                        item.get("keyword", ""),
                        self.publish_id_updates.get(int(item["row"]), item.get("cms_id", "")),
                        item.get("published_at", ""),
                    ),
                    tags=("error",) if key == "errors" else ("info",),
                )
                self.publish_review_items[f"{str(tree)}::{item_id}"] = item
        error_count = len(self.publish_review.get("errors", []))
        retry_count = len(self.publish_review.get("retry_rows", []))
        today_count = len(self.publish_review.get("posted_today", []))
        self.publish_review_count.set(
            f"{error_count} dòng cần xử lý • {retry_count} LỖI KIỂM TRA • {today_count} bài đã đăng hôm nay"
        )
        self._update_publish_review_button()

    def _select_publish_review_row(self, event: tk.Event | None = None) -> None:
        selected_tree = event.widget if event is not None else None
        if selected_tree not in (self.publish_error_tree, self.publish_today_tree):
            selected_tree = None
        if selected_tree is not None:
            other = self.publish_today_tree if selected_tree is self.publish_error_tree else self.publish_error_tree
            other.selection_remove(*other.selection())
        if selected_tree is None:
            return
        item_id = selected_tree.selection()[0]
        item = self.publish_review_items.get(f"{str(selected_tree)}::{item_id}")
        if item is None:
            return
        self.publish_selected_item = item
        for key, variable in self.publish_detail_vars.items():
            value = (
                item.get("display_status", item.get("status", ""))
                if key == "status"
                else item.get(key, "")
            )
            variable.set(str(value))
        self.publish_new_id.set(self.publish_id_updates.get(int(item["row"]), ""))

    def _stage_publish_id(self) -> None:
        item = self.publish_selected_item
        if item is None:
            messagebox.showinfo("Chưa chọn bài", "Hãy chọn một dòng cần xử lý trước.")
            return
        if "lỗi kiểm tra" in str(item.get("status", "")).casefold():
            messagebox.showwarning(
                "Không nhập ID cho LỖI KIỂM TRA",
                "Bài này phải sửa Word rồi đăng lại; không chuyển thủ công sang ĐÃ ĐĂNG.",
            )
            return
        cms_id = self.publish_new_id.get().strip()
        if not cms_id.isdigit() or int(cms_id) <= 0:
            messagebox.showwarning("ID không hợp lệ", "ID CMS phải là số nguyên lớn hơn 0.")
            return
        self.publish_id_updates[int(item["row"])] = cms_id
        self._render_publish_review()
        self.status_text.set(f"Đã ghi nhớ ID CMS {cms_id} cho dòng {item['row']}; chưa ghi vào Excel.")

    def _retry_selected_publish_error(self) -> None:
        if self._busy or self._flow_process is not None:
            return
        item = self.publish_selected_item
        if item is None:
            messagebox.showinfo("Chưa chọn bài", "Hãy chọn một dòng LỖI ĐĂNG trước.")
            return
        if "lỗi đăng" not in str(item.get("status", "")).casefold():
            messagebox.showinfo(
                "Không phải LỖI ĐĂNG",
                "Nút này chỉ dùng để chuyển một dòng LỖI ĐĂNG sang LỖI KIỂM TRA rồi chạy lại.",
            )
            return
        source = Path(self.selected_path.get().strip())
        if not source.is_file():
            messagebox.showwarning("Không tìm thấy Excel", "Hãy chọn và Phân tích đúng file Excel trước.")
            return

        retry_plan = build_retry_publish_plan(self.publish_review)
        selected_rows = list(retry_plan["selected_rows"])
        row = int(item["row"])
        if not any(int(candidate["row"]) == row for candidate in selected_rows):
            selected_rows.append(
                {
                    "row": row,
                    "domain": item.get("domain", ""),
                    "category": item.get("category", ""),
                    "title": item.get("keyword", ""),
                    "seo_title": item.get("seo_title", ""),
                    "h1": item.get("h1", ""),
                }
            )
        retry_plan["selected_rows"] = selected_rows
        retry_plan["selected_total"] = len(selected_rows)
        if not messagebox.askyesno(
            "Chuyển và đăng lại bài lỗi",
            f"Dòng {row} sẽ được chuyển từ LỖI ĐĂNG sang LỖI KIỂM TRA trong Excel, "
            f"sau đó đăng lại {retry_plan['selected_total']} bài bằng 1 worker.\n\nTiếp tục?",
        ):
            return
        self._launch_publish_id_update(source, [], retry_plan, retry_rows=[row])

    def _open_selected_excel(self) -> None:
        if self._busy or self._flow_process is not None:
            return
        path = Path(self.selected_path.get().strip())
        if not path.is_file():
            messagebox.showwarning("Không tìm thấy Excel", "Hãy chọn đúng file Excel trước.")
            return
        try:
            os.startfile(str(path))
        except OSError as exc:
            messagebox.showerror("Không mở được Excel", str(exc))

    def _open_publish_word(self) -> None:
        item = self.publish_selected_item or {}
        path = Path(str(item.get("word_path_resolved", "")))
        if not path.is_file():
            messagebox.showwarning("Không tìm thấy Word", f"File Word không tồn tại:\n{path}")
            return
        try:
            os.startfile(str(path))
        except OSError as exc:
            messagebox.showerror("Không mở được Word", str(exc))

    def _open_publish_url(self, key: str) -> None:
        import webbrowser

        url = str((self.publish_selected_item or {}).get(key, "")).strip()
        if not url.lower().startswith(("http://", "https://")):
            messagebox.showinfo("Không có URL", "Dòng này không có URL hợp lệ.")
            return
        webbrowser.open(url)

    def _update_publish_review_button(self) -> None:
        can_run = bool(
            (self.publish_id_updates or self.publish_review.get("retry_rows"))
            and not self._busy
            and self._flow_process is None
            and Path(self.selected_path.get().strip()).is_file()
        )
        if hasattr(self, "publish_review_action_button"):
            self.publish_review_action_button.configure(state="normal" if can_run else "disabled")

    def _start_publish_review_action(self) -> None:
        if self._busy or self._flow_process is not None:
            return
        source = Path(self.selected_path.get().strip())
        if not source.is_file():
            messagebox.showwarning("Không tìm thấy Excel", "Hãy chọn và Phân tích đúng file Excel trước.")
            return

        typed_id = self.publish_new_id.get().strip()
        if typed_id and self.publish_selected_item is not None:
            if "lỗi kiểm tra" in str(self.publish_selected_item.get("status", "")).casefold():
                messagebox.showwarning(
                    "Không nhập ID cho LỖI KIỂM TRA",
                    "Hãy để trống ID của bài này; app sẽ đăng lại bằng 1 worker.",
                )
                return
            if not typed_id.isdigit() or int(typed_id) <= 0:
                messagebox.showwarning("ID không hợp lệ", "ID CMS phải là số nguyên lớn hơn 0.")
                return
            self.publish_id_updates[int(self.publish_selected_item["row"])] = typed_id

        updates = [
            {"row": row, "cms_id": cms_id}
            for row, cms_id in sorted(self.publish_id_updates.items())
        ]
        retry_plan = build_retry_publish_plan(self.publish_review)
        retry_count = int(retry_plan.get("selected_total", 0))
        if not updates and not retry_count:
            messagebox.showinfo("Không có việc cần làm", "Không có ID mới hoặc bài LỖI KIỂM TRA.")
            return
        if not messagebox.askyesno(
            "Xác nhận cập nhật đăng bài",
            f"Sẽ ghi {len(updates)} ID CMS vào Excel và đăng lại {retry_count} bài "
            "LỖI KIỂM TRA bằng 1 worker.\n\nTiếp tục?",
        ):
            return
        if updates:
            self._launch_publish_id_update(source, updates, retry_plan)
        else:
            self._launch_review_publish(source, retry_plan)

    def _launch_publish_id_update(
        self,
        source: Path,
        updates: list[dict[str, Any]],
        retry_plan: dict[str, Any],
        retry_rows: list[int] | None = None,
    ) -> None:
        project_root = Path(__file__).resolve().parents[1]
        script_path = project_root / "app_flows" / "09_cap_nhat_id_dang_bai.py"
        command = [
            self._console_python(), "-u", "-m", "excel_audit_app.flow_host",
            "--workbook", str(source), "--script", str(script_path), "--",
        ]
        environment = os.environ.copy()
        environment["PYTHONIOENCODING"] = "utf-8"
        environment["HOTKEYVIP_PUBLISH_ID_UPDATES"] = json.dumps(updates, ensure_ascii=False)
        environment["HOTKEYVIP_PUBLISH_RETRY_ROWS"] = json.dumps(retry_rows or [])
        try:
            process = subprocess.Popen(
                command,
                cwd=str(project_root),
                env=environment,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except OSError as exc:
            messagebox.showerror("Không cập nhật được ID", str(exc))
            return
        self._flow_process = process
        self._flow_running_key = "publish_articles"
        self._set_busy(True, "Đang cập nhật ID CMS vào Excel...")
        self.flow_status_text.set("Đang cập nhật ID CMS")
        self.notebook.select(self.flows_tab)
        self._append_flow_log("\n=== CẬP NHẬT ID CMS ===\n")
        threading.Thread(
            target=self._publish_id_update_worker,
            args=(process, source, retry_plan),
            daemon=True,
        ).start()

    def _publish_id_update_worker(
        self,
        process: subprocess.Popen[str],
        source: Path,
        retry_plan: dict[str, Any],
    ) -> None:
        if process.stdout is not None:
            for line in process.stdout:
                self.after(0, lambda text=line: self._append_flow_log(text))
        return_code = process.wait()
        self.after(0, lambda: self._publish_id_update_finished(return_code, source, retry_plan))

    def _publish_id_update_finished(
        self, return_code: int, source: Path, retry_plan: dict[str, Any]
    ) -> None:
        self._flow_process = None
        self._flow_running_key = None
        if return_code != 0:
            self._set_busy(False, "Cập nhật ID CMS thất bại")
            messagebox.showerror(
                "Cập nhật ID thất bại",
                "Không ghi được ID CMS vào Excel. Xem nhật ký trong tab Công việc.",
            )
            return
        self.publish_id_updates.clear()
        if int(retry_plan.get("selected_total", 0)):
            self._launch_review_publish(source, retry_plan)
            return
        self._busy = False
        self._start_analysis()

    def _launch_review_publish(self, source: Path, plan: dict[str, Any]) -> None:
        project_root = Path(__file__).resolve().parents[1]
        flow = flow_by_key("publish_articles")
        script_path = flow.script_path(project_root)
        command = [
            self._console_python(), "-u", "-m", "excel_audit_app.flow_host",
            "--workbook", str(source), "--script", str(script_path), "--", *flow.script_args,
        ]
        environment = os.environ.copy()
        environment["PYTHONIOENCODING"] = "utf-8"
        environment["HOTKEYVIP_PUBLISH_PLAN"] = json.dumps(plan, ensure_ascii=False)
        environment["HOTKEYVIP_PUBLISH_WORKER_COUNT"] = "1"
        try:
            process = subprocess.Popen(
                command,
                cwd=str(project_root),
                env=environment,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
        except OSError as exc:
            self._set_busy(False)
            messagebox.showerror("Không chạy được đăng lại", str(exc))
            return
        self._flow_process = process
        self._flow_running_key = flow.key
        self._set_busy(True, f"Đang đăng lại {plan['selected_total']} bài LỖI KIỂM TRA...")
        self.flow_status_text.set(f"Đang đăng lại {plan['selected_total']} bài bằng 1 worker")
        self.notebook.select(self.flows_tab)
        self._append_flow_log(
            f"\n=== ĐĂNG LẠI {plan['selected_total']} BÀI LỖI KIỂM TRA • 1 WORKER ===\n"
        )
        threading.Thread(
            target=self._flow_worker,
            args=(process, flow.key, source),
            daemon=True,
        ).start()

    def _build_flows(self) -> None:
        outer = ttk.Frame(self.flows_tab)
        outer.pack(fill=BOTH, expand=True, padx=10, pady=10)

        left = ttk.Frame(outer)
        left.pack(side=LEFT, fill=Y, padx=(0, 10))
        tk.Label(
            left,
            text="FLOW DÙNG FILE EXCEL ĐANG CHỌN",
            bg=COLORS["bg"],
            fg=COLORS["navy"],
            font=("Segoe UI Semibold", 11),
        ).pack(anchor="w", pady=(0, 4))
        tk.Label(
            left,
            text="Excel được mở ẩn. Mỗi lần chỉ chạy một flow.",
            bg=COLORS["bg"],
            fg=COLORS["muted"],
            font=("Segoe UI", 9),
        ).pack(anchor="w", pady=(0, 8))

        for flow in ACTIVE_FLOWS:
            button_style = {
                "write_articles": "WriteFlow.TButton",
                "publish_articles": "PublishFlow.TButton",
            }.get(flow.key, "Primary.TButton")
            row = tk.Frame(
                left,
                bg=COLORS["card"],
                highlightthickness=1,
                highlightbackground=COLORS["border"],
                width=490,
                height=58,
            )
            row.pack(fill=X, pady=2)
            row.pack_propagate(False)
            text_frame = tk.Frame(row, bg=COLORS["card"])
            text_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=7)
            tk.Label(
                text_frame,
                text=flow.name,
                bg=COLORS["card"],
                fg=COLORS["navy"],
                font=("Segoe UI Semibold", 9),
            ).pack(anchor="w")
            tk.Label(
                text_frame,
                text=flow.description,
                bg=COLORS["card"],
                fg=COLORS["muted"],
                font=("Segoe UI", 8),
                wraplength=330,
                justify="left",
            ).pack(anchor="w")
            button = ttk.Button(
                row,
                text="Chạy",
                command=lambda key=flow.key: self._start_flow(key),
                style=button_style,
                state="disabled",
                width=10,
            )
            button.pack(side=RIGHT, padx=8)
            self.flow_buttons[flow.key] = button

        transfer_row = tk.Frame(
            left,
            bg=COLORS["card"],
            highlightthickness=1,
            highlightbackground="#C4B5FD",
            width=490,
            height=58,
        )
        transfer_row.pack(fill=X, pady=(5, 2))
        transfer_row.pack_propagate(False)
        transfer_text = tk.Frame(transfer_row, bg=COLORS["card"])
        transfer_text.pack(side=LEFT, fill=BOTH, expand=True, padx=10, pady=7)
        tk.Label(
            transfer_text,
            text="Chuyển URL ngày đăng mới nhất",
            bg=COLORS["card"],
            fg=COLORS["navy"],
            font=("Segoe UI Semibold", 9),
        ).pack(anchor="w")
        tk.Label(
            transfer_text,
            text="Thay danh sách Submit bằng URL trong ngày đăng mới nhất; có sao lưu danh sách cũ.",
            bg=COLORS["card"],
            fg=COLORS["muted"],
            font=("Segoe UI", 8),
            wraplength=245,
            justify="left",
        ).pack(anchor="w")
        self.transfer_submit_button = ttk.Button(
            transfer_row,
            text="Chuyển URL",
            command=self._start_submit_transfer,
            style="SubmitTransfer.TButton",
            state="disabled",
            width=10,
        )
        self.transfer_submit_button.pack(side=RIGHT, padx=8)
        self.open_submit_button = ttk.Button(
            transfer_row,
            text="Mở Submit",
            command=self._open_submit_app,
            style="Secondary.TButton",
            width=10,
        )
        self.open_submit_button.pack(side=RIGHT, padx=(0, 2))

        right = ttk.Frame(outer)
        right.pack(side=LEFT, fill=BOTH, expand=True)
        log_toolbar = ttk.Frame(right)
        log_toolbar.pack(fill=X, pady=(0, 5))
        tk.Label(
            log_toolbar,
            textvariable=self.flow_status_text,
            bg=COLORS["bg"],
            fg=COLORS["blue"],
            font=("Segoe UI Semibold", 9),
        ).pack(side=LEFT)
        ttk.Button(
            log_toolbar,
            text="Xóa nhật ký hiển thị",
            command=self._clear_flow_log,
            style="Secondary.TButton",
        ).pack(side=RIGHT)
        ttk.Button(
            log_toolbar,
            text="Lưu nhật ký",
            command=self._save_flow_log,
            style="Secondary.TButton",
        ).pack(side=RIGHT, padx=(0, 8))
        self.flow_log = tk.Text(
            right,
            bg="#101923",
            fg="#D8E4EE",
            insertbackground="#FFFFFF",
            font=("Consolas", 9),
            wrap="word",
            relief="flat",
            padx=10,
            pady=10,
            state="disabled",
        )
        self.flow_log.pack(fill=BOTH, expand=True)

    def _choose_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Chọn file Excel cần phân tích",
            filetypes=[("Excel", "*.xlsx *.xlsm"), ("Tất cả file", "*.*")],
        )
        if not path:
            return
        self.selected_path.set(path)
        self.export_button.configure(state="disabled")
        self.recover_button.configure(state="disabled")
        self._update_flow_buttons()
        self.status_text.set("Đã chọn file. Bấm Phân tích để đọc dữ liệu mới.")

    def _set_busy(self, busy: bool, message: str = "") -> None:
        self._busy = busy
        state = "disabled" if busy else "normal"
        self.choose_button.configure(state=state)
        self.analyze_button.configure(state=state)
        self.clear_button.configure(state=state)
        self.open_excel_button.configure(
            state="normal" if (not busy and Path(self.selected_path.get().strip()).is_file()) else "disabled"
        )
        if busy:
            self.export_button.configure(state="disabled")
            self.recover_button.configure(state="disabled")
        elif self.result and self._source_is_current(self.result):
            self.export_button.configure(state="normal")
            self._update_recovery_button()
        self._update_flow_buttons()
        self._update_publish_review_button()
        if message:
            self.status_text.set(message)
        self.configure(cursor="watch" if busy else "")

    def _start_analysis(self) -> None:
        if self._busy:
            return
        path = self.selected_path.get().strip()
        if not path:
            messagebox.showwarning("Chưa chọn file", "Hãy chọn file Excel trước khi phân tích.")
            return
        self._set_busy(True, "Đang đọc và đối soát ba sheet...")
        threading.Thread(target=self._analysis_worker, args=(path,), daemon=True).start()

    def _analysis_worker(self, path: str) -> None:
        try:
            result = analyze_workbook(path)
            self.store.save(result)
        except Exception as exc:  # noqa: BLE001 - cần chuyển mọi lỗi sang UI
            self.after(0, lambda: self._analysis_failed(exc))
            return
        self.after(0, lambda: self._analysis_succeeded(result))

    def _analysis_failed(self, error: Exception) -> None:
        self._set_busy(False, "Phân tích thất bại")
        messagebox.showerror("Không thể phân tích", str(error))

    def _analysis_succeeded(self, result: dict[str, Any]) -> None:
        self.result = result
        self.publish_review = result.get("publish_review", {"errors": [], "posted_today": [], "retry_rows": []})
        self.publish_id_updates.clear()
        self.selected_path.set(result.get("source_path", ""))
        self._render_result()
        overall = result.get("overall", {})
        self._set_busy(
            False,
            "Phân tích xong: "
            f"{format_number(overall.get('error_count', 0))} lỗi, "
            f"{format_number(overall.get('recovery_count', 0))} cần khôi phục, "
            f"{format_number(overall.get('pending_count', 0))} chưa chuyển DANG_BAI",
        )
        self.export_button.configure(state="normal")
        self._update_recovery_button()
        self._update_flow_buttons()

    def _source_is_current(self, result: dict[str, Any]) -> bool:
        path = Path(result.get("source_path", ""))
        if not path.exists():
            return False
        try:
            current = file_fingerprint(path, include_hash=False)
        except OSError:
            return False
        return same_fingerprint(current, result.get("source_fingerprint", {}))

    def _restore_session(self) -> None:
        result = self.store.load()
        if not result:
            return
        source_path = str(result.get("source_path", ""))
        self.selected_path.set(source_path)
        if result.get("app_version") != APP_VERSION:
            self.result = result
            self.export_button.configure(state="disabled")
            if Path(source_path).exists():
                self._render_result()
                self.health_title.set("TỔNG QUAN CHƯA CẬP NHẬT")
                self.health_detail.set(
                    f"App đã lên v{APP_VERSION}. Flow vẫn chạy ngay; bấm Phân tích khi cần số liệu mới."
                )
                self._render_health(
                    {
                        "status": self.health_title.get(),
                        "detail": self.health_detail.get(),
                        "level": "pending",
                    }
                )
                self._update_flow_buttons()
                self.status_text.set("Đã nhớ đường dẫn Excel. Không tự phân tích khi mở app.")
            else:
                self.status_text.set(
                    "Đã nhớ file của phiên trước nhưng file không còn ở đường dẫn cũ."
                )
            return
        if self._source_is_current(result):
            self.result = result
            self._render_result()
            self.export_button.configure(state="normal")
            self._update_recovery_button()
            self._update_flow_buttons()
            self.status_text.set("Đã khôi phục phiên phân tích gần nhất. File nguồn chưa thay đổi.")
        elif Path(result.get("source_path", "")).exists():
            self.result = result
            self._render_result()
            self.health_title.set("TỔNG QUAN CHƯA CẬP NHẬT")
            self.health_detail.set(
                "File đã thay đổi. Flow vẫn chạy ngay; bấm Phân tích khi cần cập nhật Tổng quan."
            )
            self._render_health(
                {
                    "status": self.health_title.get(),
                    "detail": self.health_detail.get(),
                    "level": "pending",
                }
            )
            self.export_button.configure(state="disabled")
            self.recover_button.configure(state="disabled")
            self._update_flow_buttons()
            self.status_text.set(
                "Đã nhớ đường dẫn Excel. Không tự phân tích; các flow vẫn chạy độc lập."
            )
        else:
            self.result = result
            self._render_result()
            self.export_button.configure(state="disabled")
            self.recover_button.configure(state="disabled")
            self._update_flow_buttons()
            self.status_text.set("Đã khôi phục kết quả cũ. File nguồn không còn ở đường dẫn ban đầu.")

    def _render_result(self) -> None:
        if not self.result:
            return
        summaries = self.result["summaries"]
        ke_total = summaries["ke_hoach"]["total"]
        viet_total = summaries["viet_bai"]["total"]
        rec_total = summaries["reconciliation"]["total"]
        overall = self.result.get("overall", {})
        values = {
            "errors": overall.get("error_count", 0),
            "recovery": overall.get("recovery_count", 0),
            "pending": overall.get("pending_count", 0),
            "assets": viet_total.get("completed_with_assets", 0),
        }
        for key, value in values.items():
            self.card_vars[key].set(format_number(value))

        dang_total = summaries["dang_bai"]["total"]
        write_ok = int(viet_total.get("completed_ok", 0))
        dang_rows = int(dang_total.get("total_rows", 0))
        write_dang_difference = write_ok - dang_rows
        comparison_symbol = "=" if write_dang_difference == 0 else "≠"
        self.card_vars["write_dang_match"].set(
            f"{format_number(write_ok)} {comparison_symbol} {format_number(dang_rows)}"
        )
        if write_dang_difference == 0:
            comparison_color = COLORS["green"]
            comparison_detail = "KHỚP • Chênh lệch 0"
        elif write_dang_difference > 0:
            comparison_color = COLORS["red"]
            comparison_detail = (
                f"LỆCH • ĐĂNG_BÀI thiếu {format_number(write_dang_difference)} dòng"
            )
        else:
            comparison_color = COLORS["red"]
            comparison_detail = (
                f"LỆCH • ĐĂNG_BÀI dư {format_number(abs(write_dang_difference))} dòng"
            )
        if self.write_dang_value_label is not None:
            self.write_dang_value_label.configure(fg=comparison_color)
        if self.write_dang_eyebrow_label is not None:
            self.write_dang_eyebrow_label.configure(fg=comparison_color)
        if self.write_dang_detail_label is not None:
            self.write_dang_detail_label.configure(
                text=comparison_detail,
                fg=comparison_color,
            )

        ke_url_valid = int(ke_total.get("url_valid", 0))
        dang_posted = int(dang_total.get("posted", 0))
        url_posted_difference = ke_url_valid - dang_posted
        url_comparison_symbol = "=" if url_posted_difference == 0 else "≠"
        self.card_vars["url_posted_match"].set(
            f"{format_number(ke_url_valid)} {url_comparison_symbol} {format_number(dang_posted)}"
        )
        if url_posted_difference == 0:
            url_comparison_color = COLORS["green"]
            url_comparison_detail = "KHỚP • Chênh lệch 0"
        elif url_posted_difference > 0:
            url_comparison_color = COLORS["red"]
            url_comparison_detail = (
                f"LỆCH • ĐÃ ĐĂNG thiếu {format_number(url_posted_difference)} bài"
            )
        else:
            url_comparison_color = COLORS["red"]
            url_comparison_detail = (
                f"LỆCH • ĐÃ ĐĂNG dư {format_number(abs(url_posted_difference))} bài"
            )
        if self.url_posted_value_label is not None:
            self.url_posted_value_label.configure(fg=url_comparison_color)
        if self.url_posted_eyebrow_label is not None:
            self.url_posted_eyebrow_label.configure(fg=url_comparison_color)
        if self.url_posted_detail_label is not None:
            self.url_posted_detail_label.configure(
                text=url_comparison_detail,
                fg=url_comparison_color,
            )

        self._render_health(overall)
        self.pipeline_text.set(
            f"{format_number(viet_total['total_rows'])} tổng = "
            f"{format_number(rec_total['in_dang'])} đã có trong DANG + "
            f"{format_number(rec_total['recovery_dang'])} cần khôi phục + "
            f"{format_number(rec_total['pending_dang'])} chưa chuyển + "
            f"{format_number(rec_total['viet_combo4_missing'])} thiếu Combo 4  •  "
            f"Chênh lệch: {format_number(rec_total['difference'])}"
        )

        self._fill_tree(
            self.ke_tree,
            summaries["ke_hoach"],
            [
                "domain", "total_rows", "combo4_complete", "combo4_missing", "url_valid",
                "url_written", "url_blank", "url_other", "problem_rows",
                "duplicate_groups", "duplicate_rows", "missing_in_viet",
            ],
        )
        self._fill_tree(
            self.viet_tree,
            summaries["viet_bai"],
            [
                "domain", "total_rows", "combo4_complete", "combo4_missing", "completed_ok",
                "completed_with_assets", "archived_posted_no_assets", "recovery_no_assets",
                "unexplained_no_assets", "not_completed", "duplicate_rows",
            ],
        )
        self._fill_tree(
            self.dang_tree,
            summaries["dang_bai"],
            [
                "domain", "total_rows", "combo4_complete", "combo4_missing", "in_viet",
                "posted", "url_not_posted_full_assets", "dang_missing_viet",
                "classification_difference",
            ],
        )
        self._fill_tree(
            self.reconciliation_tree,
            summaries["reconciliation"],
            [
                "domain", "ke_total", "viet_total", "ke_missing_viet", "viet_missing_ke",
                "in_dang", "recovery_dang", "pending_dang", "dang_missing_viet",
                "viet_combo4_missing", "difference", "status",
            ],
            warning_key="status",
        )
        categories = sorted({item["category"] for item in self.result.get("details", [])})
        self.category_combo.configure(values=["Tất cả", *categories])
        if self.filter_category.get() not in ["Tất cả", *categories]:
            self.filter_category.set("Tất cả")
        self._render_issues()
        self._render_recovery()
        self.publish_review = self.result.get(
            "publish_review", {"errors": [], "posted_today": [], "retry_rows": []}
        )
        self._render_publish_review()

    def _render_health(self, overall: dict[str, Any]) -> None:
        level = overall.get("level", "ok")
        palette = {
            "error": ("#FEE2E2", COLORS["red"]),
            "recovery": ("#FFEDD5", COLORS["amber"]),
            "pending": ("#DBEAFE", COLORS["blue"]),
            "ok": ("#DCFCE7", COLORS["green"]),
        }
        background, foreground = palette.get(level, palette["ok"])
        self.health_title.set(overall.get("status", "ỔN"))
        self.health_detail.set(overall.get("detail", ""))
        self.health_frame.configure(bg=background, highlightbackground=foreground)
        self.health_title_label.configure(bg=background, fg=foreground)
        self.health_detail_label.configure(bg=background, fg=COLORS["text"])

    def _show_detail_level(self, level_label: str) -> None:
        self.filter_level.set(level_label)
        self.filter_category.set("Tất cả")
        self.filter_text.set("")
        self.notebook.select(self.issues_tab)
        self._render_issues()

    def _open_reconciliation_detail(self, event: tk.Event) -> None:
        item_id = self.reconciliation_tree.identify_row(event.y)
        column_id = self.reconciliation_tree.identify_column(event.x)
        if not item_id or not column_id:
            return
        columns = list(self.reconciliation_tree["columns"])
        try:
            identifier = columns[int(column_id.removeprefix("#")) - 1]
        except (ValueError, IndexError):
            return
        category_map = {
            "ke_missing_viet": "KE_HOACH có - VIET_BAI thiếu",
            "viet_missing_ke": "VIET_BAI có - KE_HOACH thiếu",
            "recovery_dang": "Cần khôi phục DANG_BAI",
            "pending_dang": "Chưa chuyển sang DANG_BAI",
            "dang_missing_viet": "DANG_BAI có - VIET_BAI thiếu",
            "viet_combo4_missing": "VIET_BAI thiếu Combo 4",
        }
        category = category_map.get(identifier)
        if not category:
            return
        values = self.reconciliation_tree.item(item_id, "values")
        domain = str(values[0]) if values and str(values[0]) != "TỔNG TẤT CẢ" else ""
        self.filter_level.set("Tất cả")
        self.filter_category.set(category)
        self.filter_text.set(domain)
        self.notebook.select(self.issues_tab)
        self._render_issues()

    def _render_recovery(self) -> None:
        self.recovery_tree.delete(*self.recovery_tree.get_children())
        rows = self.result.get("recovery", {}).get("rows", []) if self.result else []
        for item in rows:
            self.recovery_tree.insert(
                "",
                END,
                values=[
                    format_number(item.get("ke_row", "")),
                    format_number(item.get("viet_row", "")),
                    item.get("domain", ""),
                    item.get("title", ""),
                    item.get("h1", ""),
                    item.get("keyword", ""),
                    item.get("url", ""),
                ],
            )
        self.recovery_count_text.set(f"{format_number(len(rows))} dòng sẵn sàng sao chép")
        self._update_recovery_button()

    def _update_recovery_button(self) -> None:
        if not hasattr(self, "recover_button"):
            return
        rows = self.result.get("recovery", {}).get("rows", []) if self.result else []
        error_count = int(self.result.get("overall", {}).get("error_count", 0)) if self.result else 0
        can_recover = bool(
            rows
            and error_count == 0
            and not self._busy
            and self.result
            and self._source_is_current(self.result)
        )
        self.recover_button.configure(state="normal" if can_recover else "disabled")

    def _update_flow_buttons(self) -> None:
        selected_file = Path(self.selected_path.get().strip())
        can_run = bool(
            selected_file.is_file()
            and not self._busy
            and self._flow_process is None
        )
        for button in self.flow_buttons.values():
            button.configure(state="normal" if can_run else "disabled")
        if hasattr(self, "transfer_submit_button"):
            self.transfer_submit_button.configure(
                state="normal" if can_run else "disabled"
            )
        if hasattr(self, "open_submit_button"):
            self.open_submit_button.configure(state="normal")
        if hasattr(self, "open_excel_button"):
            self.open_excel_button.configure(state="normal" if can_run else "disabled")
        self._update_publish_review_button()

    def _clear_flow_log(self) -> None:
        self.flow_log.configure(state="normal")
        self.flow_log.delete("1.0", END)
        self.flow_log.configure(state="disabled")

    def _save_flow_log(self) -> None:
        content = self.flow_log.get("1.0", "end-1c")
        if not content.strip():
            messagebox.showinfo("Lưu nhật ký", "Chưa có nội dung nhật ký để lưu.")
            return

        selected = Path(self.selected_path.get().strip())
        initial_directory = selected.parent if selected.parent.exists() else Path.cwd()
        destination = filedialog.asksaveasfilename(
            title="Lưu nhật ký flow",
            initialdir=str(initial_directory),
            initialfile=f"flow_{datetime.now():%Y%m%d_%H%M%S}.log",
            defaultextension=".log",
            filetypes=[("File nhật ký", "*.log"), ("File văn bản", "*.txt")],
        )
        if not destination:
            return
        try:
            Path(destination).write_text(content, encoding="utf-8-sig")
        except OSError as exc:
            messagebox.showerror("Không lưu được nhật ký", str(exc))
            return
        self.status_text.set(f"Đã lưu nhật ký: {destination}")
        messagebox.showinfo("Đã lưu nhật ký", destination)

    def _append_flow_log(self, value: str) -> None:
        self.flow_log.configure(state="normal")
        self.flow_log.insert(END, value)
        self.flow_log.see(END)
        self.flow_log.configure(state="disabled")

    def _start_submit_transfer(self) -> None:
        if self._busy or self._flow_process is not None:
            return
        source = Path(self.selected_path.get().strip())
        if not source.is_file():
            messagebox.showwarning("Chưa chọn file", "Hãy chọn đúng file Excel trước.")
            return
        self._set_busy(True, "Đang tìm URL thuộc ngày đăng mới nhất...")
        threading.Thread(
            target=self._submit_transfer_preview_worker,
            args=(source,),
            daemon=True,
        ).start()

    def _open_submit_app(self) -> None:
        launcher = submit_launcher_path()
        try:
            if not launcher.is_file():
                raise FileNotFoundError(f"Không tìm thấy: {launcher}")
            os.startfile(str(launcher))
        except OSError as exc:
            messagebox.showerror("Không mở được app Submit", str(exc))
            return
        self.status_text.set("Đã mở app Submit")

    def _submit_transfer_preview_worker(self, source: Path) -> None:
        try:
            snapshot = inspect_latest_published_urls(source)
        except Exception as exc:  # noqa: BLE001 - chuyển lỗi sang giao diện
            self.after(0, lambda: self._submit_transfer_failed(exc))
            return
        self.after(0, lambda: self._show_submit_transfer_preview(source, snapshot))

    def _show_submit_transfer_preview(
        self,
        source: Path,
        snapshot: dict[str, Any],
    ) -> None:
        self._set_busy(False)
        latest = snapshot.get("latest_date")
        if not latest:
            messagebox.showinfo(
                "Không có URL để chuyển",
                "Không tìm thấy bài ĐÃ ĐĂNG có Thời gian đăng trong DANG_BAI.",
            )
            return
        latest_display = datetime.strptime(str(latest), "%Y-%m-%d").strftime("%d/%m/%Y")
        summary = (
            f"Ngày đăng mới nhất: {latest_display}\n\n"
            f"Bài đã đăng trong ngày: {snapshot['published_total']}\n"
            f"Có URL công khai hợp lệ: {snapshot['valid_total']}\n"
            f"Thiếu URL công khai: {snapshot['missing_url_total']}\n"
            f"URL trùng trong Excel đã bỏ: {snapshot['duplicate_total']}\n"
            f"URL đang lưu bên Submit sẽ xóa: {snapshot['history_total']}\n"
            f"Sẽ chuyển sang Submit: {snapshot['new_total']}"
        )
        if not snapshot["new_total"]:
            messagebox.showinfo(
                "Không có URL mới",
                summary + "\n\nKhông thay đổi danh sách URL bên Submit.",
            )
            return
        if not messagebox.askyesno(
            "Chuyển URL sang app Submit",
            summary
            + "\n\nDanh sách URL đang lưu sẽ được sao lưu rồi thay bằng danh sách mới."
            + "\nHãy đóng app Submit trước khi chuyển."
            + "\n\nBạn chắc chắn muốn tiếp tục?",
        ):
            return
        self._set_busy(True, "Đang thêm URL mới vào hàng chờ Submit...")
        threading.Thread(
            target=self._submit_transfer_worker,
            args=(source,),
            daemon=True,
        ).start()

    def _submit_transfer_worker(self, source: Path) -> None:
        try:
            result = transfer_latest_published_urls(source)
        except Exception as exc:  # noqa: BLE001 - chuyển lỗi sang giao diện
            self.after(0, lambda: self._submit_transfer_failed(exc))
            return
        self.after(0, lambda: self._submit_transfer_succeeded(result))

    def _submit_transfer_failed(self, error: Exception) -> None:
        self._set_busy(False, "Chuyển URL sang Submit thất bại")
        messagebox.showerror("Không thể chuyển URL", str(error))

    def _submit_transfer_succeeded(self, result: dict[str, Any]) -> None:
        added = int(result.get("added_total", 0))
        removed = int(result.get("removed_total", 0))
        self._set_busy(False, f"Đã chuyển {added} URL sang app Submit")
        messagebox.showinfo(
            "Đã chuyển URL",
            f"Đã xóa {removed} URL cũ và chuyển {added} URL của ngày mới nhất.\n\n"
            f"Bản sao danh sách cũ: {result['backup_path']}\n"
            f"File: {result['history_path']}\n\n"
            "Bấm nút Mở Submit khi muốn chạy app Submit.",
        )

    @staticmethod
    def _console_python() -> str:
        executable = Path(sys.executable)
        if executable.name.casefold() == "pythonw.exe":
            candidate = executable.with_name("python.exe")
            if candidate.exists():
                return str(candidate)
        return str(executable)

    def _prompt_write_plan(self, source: Path) -> dict[str, Any] | None:
        """Đọc workbook và hiển thị đầy đủ kế hoạch trước khi chạy Flow 3."""
        try:
            snapshot = inspect_write_queue(source)
        except Exception as exc:
            messagebox.showerror(
                "Không đọc được kế hoạch viết bài",
                f"Không thể đọc hàng chờ VIET_BAI:\n\n{exc}",
                parent=self,
            )
            return None

        domain_counts = snapshot["domain_counts"]
        domain_error_counts = snapshot["domain_error_counts"]
        domains = sorted(domain_counts, key=lambda value: (-domain_counts[value], value.casefold()))

        dialog = tk.Toplevel(self)
        dialog.title("Thiết lập ưu tiên viết bài")
        dialog.transient(self)
        dialog.resizable(False, False)
        dialog.configure(bg=COLORS["bg"])
        dialog.grab_set()

        body = ttk.Frame(dialog, padding=18)
        body.grid(row=0, column=0, sticky="nsew")
        ttk.Label(
            body,
            text="XẾP HÀNG VIẾT BÀI",
            foreground=COLORS["navy"],
            font=("Segoe UI Semibold", 12),
        ).grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Label(
            body,
            text=(
                "Dữ liệu không bị di chuyển trong Excel. App chỉ đổi thứ tự bài "
                "được đưa vào hàng chờ của lần chạy này."
            ),
            foreground=COLORS["muted"],
            wraplength=520,
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(0, 14))

        no_priority = "— Không ưu tiên tên miền —"
        domain_labels: dict[str, str] = {no_priority: ""}
        for domain in domains:
            remaining = int(domain_counts.get(domain, 0))
            errors = int(domain_error_counts.get(domain, 0))
            error_text = f", lỗi {errors}" if errors else ""
            domain_labels[f"{domain} — còn {remaining} bài{error_text}"] = domain
        domain_var = tk.StringVar(value=no_priority)
        count_var = tk.StringVar(value="100")
        retry_var = tk.BooleanVar(value=True)
        continue_var = tk.BooleanVar(value=True)

        ttk.Label(body, text="Tên miền ưu tiên").grid(row=2, column=0, sticky="w", padx=(0, 12), pady=5)
        ttk.Combobox(
            body,
            textvariable=domain_var,
            values=list(domain_labels),
            state="readonly",
            width=58,
        ).grid(row=2, column=1, sticky="ew", pady=5)

        ttk.Label(body, text="Số bài ưu tiên").grid(row=3, column=0, sticky="w", padx=(0, 12), pady=5)
        ttk.Spinbox(body, from_=1, to=10000, textvariable=count_var, width=12).grid(
            row=3, column=1, sticky="w", pady=5
        )
        ttk.Checkbutton(
            body,
            text="Ưu tiên chạy lại các bài đang lỗi trước",
            variable=retry_var,
        ).grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 4))
        ttk.Checkbutton(
            body,
            text="Xong nhóm ưu tiên thì tiếp tục chạy bình thường sau OK OK",
            variable=continue_var,
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=4)

        manual_row = snapshot.get("manual_row")
        normal_start = int(snapshot["normal_start_row"])
        normal_count = len(snapshot["normal_items"])
        marker_text = (
            f"OK OK cuối cùng: dòng {manual_row} → chạy thường từ dòng {normal_start} "
            f"({normal_count} bài chưa hoàn tất phía sau)."
            if manual_row is not None
            else f"Không tìm thấy OK OK → chạy thường từ dòng {normal_start} ({normal_count} bài)."
        )
        ttk.Label(
            body,
            text=f"Bài lỗi đang chờ chạy lại: {len(snapshot['error_items'])}\n{marker_text}",
            foreground=COLORS["navy"],
            wraplength=620,
            font=("Segoe UI Semibold", 9),
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(12, 6))

        preview_var = tk.StringVar()
        ttk.Label(
            body,
            textvariable=preview_var,
            foreground=COLORS["text"],
            background="#FFFFFF",
            relief="solid",
            borderwidth=1,
            justify="left",
            anchor="nw",
            wraplength=620,
            padding=10,
        ).grid(row=7, column=0, columnspan=2, sticky="ew", pady=(4, 0))

        def current_plan() -> dict[str, Any]:
            selected_domain = domain_labels.get(domain_var.get().strip(), "")
            try:
                priority_count = max(0, int(count_var.get().strip()))
            except ValueError:
                priority_count = 0
            return {
                "retry_errors_first": bool(retry_var.get()),
                "priority_domain": selected_domain,
                "priority_count": priority_count if selected_domain else 0,
                "continue_normal": bool(continue_var.get()),
            }

        def refresh_preview(*_args: Any) -> None:
            plan = current_plan()
            queue_preview = build_write_queue_preview(snapshot, plan)
            lines = [f"Tổng hàng chờ dự kiến: {len(queue_preview)} bài. Những bài chạy đầu tiên:"]
            error_prefix = len(snapshot["error_items"]) if plan["retry_errors_first"] else 0
            if error_prefix and len(queue_preview) > error_prefix:
                next_item = queue_preview[error_prefix]
                lines.append(
                    f"Sau {error_prefix} bài lỗi → dòng {next_item['row']} | "
                    f"{next_item.get('domain') or '(không domain)'} | {next_item.get('keyword')}"
                )
            for index, item in enumerate(queue_preview[:6], start=1):
                flag = "LỖI" if item.get("is_error") else "BÀI"
                lines.append(
                    f"{index}. [{flag}] dòng {item['row']} | "
                    f"{item.get('domain') or '(không domain)'} | {item.get('keyword')}"
                )
            if not queue_preview:
                lines.append("Không có bài nào phù hợp với thiết lập hiện tại.")
            elif len(queue_preview) > 6:
                lines.append(f"… và {len(queue_preview) - 6} bài tiếp theo.")
            preview_var.set("\n".join(lines))

        domain_var.trace_add("write", refresh_preview)
        count_var.trace_add("write", refresh_preview)
        retry_var.trace_add("write", refresh_preview)
        continue_var.trace_add("write", refresh_preview)
        refresh_preview()

        result: dict[str, Any] = {}

        def accept() -> None:
            selected_domain = domain_labels.get(domain_var.get().strip(), "")
            try:
                priority_count = int(count_var.get().strip())
            except ValueError:
                messagebox.showwarning(
                    "Số lượng không hợp lệ",
                    "Số bài ưu tiên phải là số nguyên.",
                    parent=dialog,
                )
                return
            if selected_domain and priority_count < 1:
                messagebox.showwarning(
                    "Số lượng không hợp lệ",
                    "Số bài ưu tiên phải lớn hơn 0.",
                    parent=dialog,
                )
                return
            result.update(current_plan())
            result["preview_total"] = len(build_write_queue_preview(snapshot, result))
            result["manual_mark_row"] = manual_row
            result["normal_start_row"] = normal_start
            dialog.destroy()

        def cancel() -> None:
            dialog.destroy()

        buttons = ttk.Frame(body)
        buttons.grid(row=8, column=0, columnspan=2, sticky="e", pady=(18, 0))
        ttk.Button(buttons, text="Hủy", command=cancel).pack(side=LEFT, padx=(0, 8))
        ttk.Button(buttons, text="Tiếp tục", command=accept, style="Primary.TButton").pack(side=LEFT)
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Escape>", lambda _event: cancel())
        dialog.bind("<Return>", lambda _event: accept())
        dialog.update_idletasks()
        dialog.geometry(
            f"+{self.winfo_rootx() + max(40, (self.winfo_width() - dialog.winfo_reqwidth()) // 2)}"
            f"+{self.winfo_rooty() + max(40, (self.winfo_height() - dialog.winfo_reqheight()) // 2)}"
        )
        self.wait_window(dialog)
        return result or None

    def _prompt_publish_plan(self, source: Path) -> dict[str, Any] | None:
        """Hiển thị batch domain + một danh mục trước khi chạy Flow 5."""
        try:
            snapshot = inspect_publish_queue(source)
        except Exception as exc:
            messagebox.showerror(
                "Không đọc được kế hoạch đăng bài",
                f"Không thể đọc hàng chờ DANG_BAI:\n\n{exc}",
                parent=self,
            )
            return None

        dialog = tk.Toplevel(self)
        dialog.title("Thiết lập batch đăng bài")
        dialog.transient(self)
        dialog_width = min(900, max(720, dialog.winfo_screenwidth() - 160))
        dialog_height = min(720, max(560, dialog.winfo_screenheight() - 140))
        dialog.geometry(f"{dialog_width}x{dialog_height}")
        dialog.minsize(700, 540)
        dialog.configure(bg=COLORS["bg"])
        dialog.grab_set()

        body = ttk.Frame(dialog, padding=16)
        body.pack(fill=BOTH, expand=True)
        body.columnconfigure(0, weight=1)
        body.rowconfigure(3, weight=1)
        ttk.Label(
            body,
            text="ĐĂNG THEO TÊN MIỀN + MỘT DANH MỤC",
            foreground=COLORS["navy"],
            font=("Segoe UI Semibold", 12),
        ).grid(row=0, column=0, sticky="w")
        ttk.Label(
            body,
            text=(
                "Mỗi tên miền mặc định chọn danh mục có nhiều bài hợp lệ nhất. "
                "Chọn một dòng bên dưới để đổi sang danh mục khác. "
                "Nếu danh mục chỉ có 5 bài thì đăng 5; không ghép danh mục khác."
            ),
            foreground=COLORS["muted"],
            wraplength=730,
        ).grid(row=1, column=0, sticky="w", pady=(4, 12))

        controls = ttk.Frame(body)
        controls.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        ttk.Label(controls, text="Số bài tối đa mỗi tên miền").pack(side=LEFT)
        limit_var = tk.StringVar(value="7")
        ttk.Spinbox(controls, from_=1, to=10000, textvariable=limit_var, width=10).pack(
            side=LEFT, padx=(10, 0)
        )
        ttk.Label(controls, text="Danh mục của dòng đang chọn").pack(
            side=LEFT, padx=(28, 8)
        )
        category_var = tk.StringVar(value="— Chọn một tên miền bên dưới —")
        category_picker = ttk.Combobox(
            controls,
            textvariable=category_var,
            state="disabled",
            width=34,
        )
        category_picker.pack(side=LEFT)

        columns = ("domain", "category", "available", "selected")
        tree = ttk.Treeview(body, columns=columns, show="headings", height=8)
        tree.heading("domain", text="Tên miền")
        tree.heading("category", text="Danh mục được chọn")
        tree.heading("available", text="Có thể đăng")
        tree.heading("selected", text="Sẽ đăng")
        tree.column("domain", width=220, anchor="w")
        tree.column("category", width=260, anchor="w")
        tree.column("available", width=95, anchor="center")
        tree.column("selected", width=85, anchor="center")
        tree.grid(row=3, column=0, sticky="nsew")

        summary_var = tk.StringVar()
        ttk.Label(
            body,
            textvariable=summary_var,
            foreground=COLORS["navy"],
            font=("Segoe UI Semibold", 9),
        ).grid(row=4, column=0, sticky="w", pady=(10, 0))

        current_plan: dict[str, Any] = {}
        category_overrides: dict[str, str] = {}
        tree_domain_keys: dict[str, str] = {}
        category_choice_keys: dict[str, str] = {}
        selected_domain_key = tk.StringVar(value="")

        def refresh(*_args: Any) -> None:
            try:
                limit = int(limit_var.get().strip())
            except ValueError:
                summary_var.set("Số bài phải là số nguyên lớn hơn 0.")
                return
            if limit < 1:
                summary_var.set("Số bài phải lớn hơn 0.")
                return
            plan = build_balanced_publish_plan(snapshot, limit, category_overrides)
            current_plan.clear()
            current_plan.update(plan)
            tree.delete(*tree.get_children())
            tree_domain_keys.clear()
            for group in plan["groups"]:
                item_id = tree.insert(
                    "",
                    END,
                    values=(
                        group["domain"],
                        group["category"],
                        format_number(group["available"]),
                        format_number(group["selected"]),
                    ),
                )
                tree_domain_keys[item_id] = group["domain_key"]
                if group["domain_key"] == selected_domain_key.get():
                    tree.selection_set(item_id)
                    tree.focus(item_id)
            summary_var.set(
                f"{len(plan['groups'])} tên miền • "
                f"Sẽ đăng {format_number(plan['selected_total'])} bài • "
                f"Bỏ ngoài batch {format_number(plan['skipped_invalid_total'])} dòng thiếu dữ liệu/Word"
            )

        def selected_group() -> dict[str, Any] | None:
            selection = tree.selection()
            if not selection:
                return None
            domain_key = tree_domain_keys.get(selection[0], "")
            return next(
                (
                    group
                    for group in current_plan.get("groups", [])
                    if group.get("domain_key") == domain_key
                ),
                None,
            )

        def show_category_choices(*_args: Any) -> None:
            group = selected_group()
            category_choice_keys.clear()
            if group is None:
                selected_domain_key.set("")
                category_var.set("— Chọn một tên miền bên dưới —")
                category_picker.configure(state="disabled", values=[])
                return
            selected_domain_key.set(group["domain_key"])
            labels: list[str] = []
            selected_label = ""
            for option in group.get("category_options", []):
                label = f"{option['label']} ({format_number(option['available'])} bài)"
                labels.append(label)
                category_choice_keys[label] = option["key"]
                if option["key"] == group.get("category_key"):
                    selected_label = label
            category_picker.configure(state="readonly", values=labels)
            category_var.set(selected_label or (labels[0] if labels else ""))

        def change_category(*_args: Any) -> None:
            domain_key = selected_domain_key.get()
            category_key = category_choice_keys.get(category_var.get(), "")
            if not domain_key or not category_key:
                return
            category_overrides[domain_key] = category_key
            refresh()
            show_category_choices()

        limit_var.trace_add("write", refresh)
        tree.bind("<<TreeviewSelect>>", show_category_choices)
        category_picker.bind("<<ComboboxSelected>>", change_category)
        refresh()
        result: dict[str, Any] = {}

        def accept() -> None:
            if not current_plan.get("selected_total"):
                messagebox.showwarning(
                    "Không có bài để đăng",
                    "Thiết lập hiện tại không tạo được bài nào đủ điều kiện.",
                    parent=dialog,
                )
                return
            result.update(current_plan)
            dialog.destroy()

        def cancel() -> None:
            dialog.destroy()

        buttons = ttk.Frame(body)
        buttons.grid(row=5, column=0, sticky="e", pady=(14, 0))
        ttk.Button(buttons, text="Hủy", command=cancel).pack(side=LEFT, padx=(0, 8))
        ttk.Button(
            buttons,
            text="Xác nhận batch và đăng",
            command=accept,
            style="Primary.TButton",
            width=24,
        ).pack(side=LEFT)
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        dialog.bind("<Escape>", lambda _event: cancel())
        self.wait_window(dialog)
        return result or None

    def _start_flow(self, flow_key: str) -> None:
        if self._busy or self._flow_process is not None:
            return
        source = Path(self.selected_path.get().strip())
        if not source.is_file():
            messagebox.showwarning("Chưa chọn file", "Hãy chọn đúng file Excel trước khi chạy flow.")
            return
        flow = flow_by_key(flow_key)
        write_plan: dict[str, Any] | None = None
        publish_plan: dict[str, Any] | None = None
        if flow.key == "write_articles":
            write_plan = self._prompt_write_plan(source)
            if write_plan is None:
                return
        if flow.key == "publish_articles":
            publish_plan = self._prompt_publish_plan(source)
            if publish_plan is None:
                return
        project_root = Path(__file__).resolve().parents[1]
        script_path = flow.script_path(project_root)
        if not script_path.exists():
            messagebox.showerror("Thiếu file flow", f"Không tìm thấy:\n{script_path}")
            return
        warning = "\n\nFlow này có tác động ra hệ thống bên ngoài." if flow.external_effects else ""
        plan_text = ""
        if write_plan is not None:
            priority_domain = str(write_plan.get("priority_domain", ""))
            priority_text = (
                f"Ưu tiên {write_plan['priority_count']} bài của {priority_domain}."
                if priority_domain
                else "Không ưu tiên tên miền."
            )
            retry_text = "Có" if write_plan.get("retry_errors_first") else "Không"
            continue_text = "Có" if write_plan.get("continue_normal") else "Không"
            marker_row = write_plan.get("manual_mark_row")
            marker_text = (
                f"dòng {marker_row}, chạy thường từ dòng {write_plan.get('normal_start_row')}"
                if marker_row is not None
                else f"không có, chạy thường từ dòng {write_plan.get('normal_start_row')}"
            )
            plan_text = (
                f"\n\nKế hoạch viết:\n- Bài lỗi trước: {retry_text}"
                f"\n- {priority_text}\n- Mốc OK OK: {marker_text}"
                f"\n- Tiếp tục sau OK OK: {continue_text}"
                f"\n- Tổng hàng chờ dự kiến: {write_plan.get('preview_total', 0)} bài"
            )
        if publish_plan is not None:
            plan_text = (
                f"\n\nKế hoạch đăng:\n- Tối đa mỗi tên miền: "
                f"{publish_plan['per_domain_limit']} bài"
                f"\n- Mỗi tên miền chỉ dùng một danh mục"
                f"\n- Số tên miền: {len(publish_plan['groups'])}"
                f"\n- Tổng batch sẽ đăng: {publish_plan['selected_total']} bài"
            )
        if not messagebox.askyesno(
            "Xác nhận chạy flow",
            f"{flow.name}\n\n{flow.confirmation}{warning}\n\n"
            f"File Excel:\n{source}{plan_text}\n\nTiếp tục?",
        ):
            return

        command = [
            self._console_python(),
            "-u",
            "-m",
            "excel_audit_app.flow_host",
            "--workbook",
            str(source),
            "--script",
            str(script_path),
            "--",
            *flow.script_args,
        ]
        environment = os.environ.copy()
        environment["PYTHONIOENCODING"] = "utf-8"
        if write_plan is not None:
            environment["HOTKEYVIP_WRITE_PLAN"] = json.dumps(write_plan, ensure_ascii=False)
        if publish_plan is not None:
            environment["HOTKEYVIP_PUBLISH_PLAN"] = json.dumps(
                {
                    "mode": publish_plan["mode"],
                    "per_domain_limit": publish_plan["per_domain_limit"],
                    "category_overrides": publish_plan.get("category_overrides", {}),
                },
                ensure_ascii=False,
            )
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            process = subprocess.Popen(
                command,
                cwd=str(project_root),
                env=environment,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                creationflags=creation_flags,
            )
        except OSError as exc:
            messagebox.showerror("Không thể chạy flow", str(exc))
            return

        self._flow_process = process
        self._flow_running_key = flow.key
        self._busy = True
        self.choose_button.configure(state="disabled")
        self.analyze_button.configure(state="disabled")
        self.export_button.configure(state="disabled")
        self.clear_button.configure(state="disabled")
        self.recover_button.configure(state="disabled")
        self._update_flow_buttons()
        self.flow_status_text.set(f"Đang chạy: {flow.name}")
        self.status_text.set(f"Đang chạy flow: {flow.name}")
        self.notebook.select(self.flows_tab)
        self._append_flow_log(
            f"\n{'=' * 72}\n{flow.name}\nFile: {source}\n{'=' * 72}\n"
        )
        threading.Thread(
            target=self._flow_worker,
            args=(process, flow.key, source),
            daemon=True,
        ).start()

    def _flow_worker(
        self,
        process: subprocess.Popen[str],
        flow_key: str,
        source: Path,
    ) -> None:
        if process.stdout is not None:
            for line in process.stdout:
                self.after(0, lambda text=line: self._append_flow_log(text))
        return_code = process.wait()
        refreshed: dict[str, Any] | None = None
        analysis_error: Exception | None = None
        try:
            refreshed = analyze_workbook(source)
            self.store.save(refreshed)
        except Exception as exc:  # noqa: BLE001 - hiển thị lỗi làm mới sau flow
            analysis_error = exc
        self.after(
            0,
            lambda: self._flow_finished(
                flow_key,
                return_code,
                refreshed,
                analysis_error,
            ),
        )

    def _flow_finished(
        self,
        flow_key: str,
        return_code: int,
        refreshed: dict[str, Any] | None,
        analysis_error: Exception | None,
    ) -> None:
        flow = flow_by_key(flow_key)
        self._flow_process = None
        self._flow_running_key = None
        self._busy = False
        if refreshed is not None:
            self.result = refreshed
            self.publish_review = refreshed.get(
                "publish_review", {"errors": [], "posted_today": [], "retry_rows": []}
            )
            self.publish_id_updates.clear()
            self.selected_path.set(refreshed.get("source_path", ""))
            self._render_result()
        self.choose_button.configure(state="normal")
        self.analyze_button.configure(state="normal")
        self.clear_button.configure(state="normal")
        self._set_busy(False)
        self._update_flow_buttons()
        if self.result is not None and refreshed is None:
            self.health_title.set("TỔNG QUAN CHƯA CẬP NHẬT")
            self.health_detail.set(
                "Flow vừa thay đổi Excel. Vẫn có thể chạy flow tiếp; bấm Phân tích khi cần xem số liệu mới."
            )
            self._render_health(
                {
                    "status": self.health_title.get(),
                    "detail": self.health_detail.get(),
                    "level": "pending",
                }
            )
        if return_code == 0 and analysis_error is None:
            self.flow_status_text.set(f"Hoàn thành: {flow.name}")
            self.status_text.set(f"Flow hoàn thành: {flow.name} • Danh sách lỗi đã cập nhật")
            messagebox.showinfo(
                "Flow đã hoàn thành",
                f"{flow.name}\n\nFile Excel đã được lưu và task Theo dõi đăng bài đã được cập nhật.",
            )
        else:
            detail = f"Mã thoát: {return_code}"
            if analysis_error is not None:
                detail += f"\nKhông thể phân tích lại: {analysis_error}"
            self.flow_status_text.set(f"Có lỗi: {flow.name}")
            self.status_text.set(f"Flow kết thúc có lỗi: {flow.name}")
            messagebox.showerror(
                "Flow kết thúc có lỗi",
                f"{flow.name}\n\n{detail}\n\nXem nhật ký trong tab Công việc.",
            )

    @staticmethod
    def _clipboard_cell(value: Any) -> str:
        return str(value if value is not None else "").replace("\t", " ").replace("\r", " ").replace("\n", " ")

    def _copy_all_recovery(self) -> None:
        rows = self.result.get("recovery", {}).get("rows", []) if self.result else []
        if not rows:
            messagebox.showinfo("Không có dữ liệu", "Không có dòng nào cần khôi phục vào DANG_BAI.")
            return
        text = "\r\n".join(
            "\t".join(self._clipboard_cell(value) for value in item.get("values", []))
            for item in rows
        )
        self.clipboard_clear()
        self.clipboard_append(text)
        self.update_idletasks()
        self.status_text.set(
            f"Đã sao chép {format_number(len(rows))} dòng theo đúng thứ tự cột DANG_BAI"
        )

    def _copy_one_recovery(self, event: tk.Event) -> None:
        item_id = self.recovery_tree.identify_row(event.y)
        if not item_id or not self.result:
            return
        index = self.recovery_tree.index(item_id)
        rows = self.result.get("recovery", {}).get("rows", [])
        if index >= len(rows):
            return
        values = rows[index].get("values", [])
        self.clipboard_clear()
        self.clipboard_append("\t".join(self._clipboard_cell(value) for value in values))
        self.status_text.set("Đã sao chép 1 dòng theo đúng thứ tự cột DANG_BAI")

    def _start_recovery(self) -> None:
        if not self.result or self._busy:
            return
        rows = self.result.get("recovery", {}).get("rows", [])
        if not rows:
            messagebox.showinfo("Không có dữ liệu", "Không có dòng nào cần khôi phục vào DANG_BAI.")
            return
        if int(self.result.get("overall", {}).get("error_count", 0)) != 0:
            messagebox.showwarning(
                "Còn lỗi dữ liệu",
                "Hãy xử lý lỗi dữ liệu trước khi tự động khôi phục DANG_BAI.",
            )
            return
        source = Path(self.result["source_path"])
        if not source.exists():
            messagebox.showerror("Không tìm thấy file", "File nguồn không còn ở đường dẫn ban đầu.")
            return
        if not self._source_is_current(self.result):
            messagebox.showwarning(
                "File đã thay đổi",
                "File nguồn đã thay đổi sau lần phân tích. Hãy bấm Phân tích lại trước khi khôi phục.",
            )
            return
        suggested = suggested_recovery_path(source)
        destination = filedialog.asksaveasfilename(
            title="Tạo bản Excel mới và khôi phục DANG_BAI",
            initialdir=str(suggested.parent),
            initialfile=suggested.name,
            defaultextension=source.suffix,
            filetypes=[(f"Excel {source.suffix.upper()}", f"*{source.suffix}")],
        )
        if not destination:
            return
        if Path(destination).resolve() == source.resolve():
            messagebox.showerror("Không thể ghi đè", "Hãy chọn tên file khác file nguồn.")
            return
        if not messagebox.askyesno(
            "Xác nhận khôi phục",
            f"App sẽ tạo file mới và thêm {format_number(len(rows))} dòng vào DANG_BAI.\n\n"
            "File Excel gốc sẽ không bị thay đổi. Tiếp tục?",
        ):
            return
        self._set_busy(
            True,
            f"Đang tạo file mới và khôi phục {format_number(len(rows))} dòng vào DANG_BAI...",
        )
        threading.Thread(
            target=self._recovery_worker,
            args=(str(source), destination, self.result, len(rows)),
            daemon=True,
        ).start()

    def _recovery_worker(
        self,
        source: str,
        destination: str,
        result: dict[str, Any],
        recovery_count: int,
    ) -> None:
        try:
            output = recover_dang_bai(source, destination, result)
            refreshed = analyze_workbook(output)
            self.store.save(refreshed)
        except (ExportError, SourceChangedError, OSError, ValueError) as exc:
            self.after(0, lambda error=exc: self._recovery_failed(error))
            return
        self.after(
            0,
            lambda: self._recovery_succeeded(output, refreshed, recovery_count),
        )

    def _recovery_failed(self, error: Exception) -> None:
        self._set_busy(False, "Khôi phục DANG_BAI thất bại")
        messagebox.showerror("Không thể khôi phục DANG_BAI", str(error))

    def _recovery_succeeded(
        self,
        output: Path,
        refreshed: dict[str, Any],
        recovery_count: int,
    ) -> None:
        self.result = refreshed
        self.selected_path.set(str(output))
        self._render_result()
        self._set_busy(
            False,
            f"Đã tạo bản mới và khôi phục {format_number(recovery_count)} dòng: {output}",
        )
        messagebox.showinfo(
            "Khôi phục DANG_BAI thành công",
            f"Đã tạo file mới:\n{output}\n\n"
            f"Đã thêm {format_number(recovery_count)} dòng vào DANG_BAI và kiểm tra lại thành công.\n"
            "File Excel gốc không bị thay đổi.",
        )

    @staticmethod
    def _fill_tree(
        tree: ttk.Treeview,
        summary: dict[str, Any],
        keys: list[str],
        warning_key: str | None = None,
    ) -> None:
        tree.delete(*tree.get_children())
        for row in summary.get("rows", []):
            tags = ()
            if warning_key and str(row.get(warning_key, "")).upper() != "KHỚP":
                tags = ("warning",)
            tree.insert("", END, values=[format_number(row.get(key, "")) for key in keys], tags=tags)
        total = summary.get("total")
        if total:
            tags = ["total"]
            if warning_key and str(total.get(warning_key, "")).upper() != "KHỚP":
                tags.append("warning")
            tree.insert(
                "", END, values=[format_number(total.get(key, "")) for key in keys], tags=tuple(tags)
            )

    def _render_issues(self) -> None:
        self.issue_tree.delete(*self.issue_tree.get_children())
        if not self.result:
            self.issue_count_label.configure(text="0 dòng")
            return
        category = self.filter_category.get()
        level_filter = LEVEL_FILTERS.get(self.filter_level.get())
        query = self.filter_text.get().strip().casefold()
        shown = 0
        for item in self.result.get("details", self.result.get("issues", [])):
            level = item.get("level", "error")
            if level_filter and level != level_filter:
                continue
            if category != "Tất cả" and item["category"] != category:
                continue
            haystack = " ".join(str(value) for value in item.values()).casefold()
            if query and query not in haystack:
                continue
            self.issue_tree.insert(
                "",
                END,
                values=[
                    LEVEL_LABELS.get(level, level), item["category"], item["sheet"],
                    format_number(item["row"]), item.get("target_sheet", ""),
                    format_number(item.get("target_row", "")), item["domain"],
                    item["title"], item["h1"], item["keyword"], item["detail"],
                ],
                tags=(level,),
            )
            shown += 1
        self.issue_count_label.configure(text=f"{format_number(shown)} dòng")

    def _copy_issue(self, _event: tk.Event) -> None:
        selected = self.issue_tree.selection()
        if not selected:
            return
        values = self.issue_tree.item(selected[0], "values")
        self.clipboard_clear()
        self.clipboard_append("\t".join(str(value) for value in values))
        self.status_text.set("Đã sao chép dòng chi tiết vào clipboard")

    def _start_export(self) -> None:
        if not self.result or self._busy:
            return
        source = Path(self.result["source_path"])
        if not source.exists():
            messagebox.showerror("Không tìm thấy file", "File nguồn không còn ở đường dẫn ban đầu.")
            return
        if not self._source_is_current(self.result):
            messagebox.showwarning(
                "File đã thay đổi",
                "File nguồn đã thay đổi sau lần phân tích. Hãy bấm Phân tích lại trước khi xuất.",
            )
            return
        suggested = suggested_output_path(source)
        destination = filedialog.asksaveasfilename(
            title="Xuất bản Excel đã kiểm tra",
            initialdir=str(suggested.parent),
            initialfile=suggested.name,
            defaultextension=source.suffix,
            filetypes=[(f"Excel {source.suffix.upper()}", f"*{source.suffix}")],
        )
        if not destination:
            return
        if Path(destination).resolve() == source.resolve():
            messagebox.showerror("Không thể ghi đè", "Hãy chọn tên file khác file nguồn.")
            return
        self._set_busy(True, "Đang tạo bản Excel mới và ghi sheet Tong_all...")
        threading.Thread(
            target=self._export_worker,
            args=(str(source), destination, self.result),
            daemon=True,
        ).start()

    def _export_worker(self, source: str, destination: str, result: dict[str, Any]) -> None:
        try:
            output = export_result(source, destination, result)
        except (ExportError, SourceChangedError, OSError, ValueError) as exc:
            self.after(0, lambda: self._export_failed(exc))
            return
        self.after(0, lambda: self._export_succeeded(output))

    def _export_failed(self, error: Exception) -> None:
        self._set_busy(False, "Xuất kết quả thất bại")
        messagebox.showerror("Không thể xuất kết quả", str(error))

    def _export_succeeded(self, output: Path) -> None:
        self._set_busy(False, f"Đã xuất: {output}")
        messagebox.showinfo(
            "Xuất kết quả thành công",
            f"Đã tạo file mới:\n{output}\n\nFile Excel gốc không bị thay đổi.",
        )

    def _clear_session(self) -> None:
        if not self.result and not self.store.session_path.exists():
            return
        if not messagebox.askyesno(
            "Xóa phiên gần nhất",
            "Xóa kết quả phân tích đã lưu trong app? File Excel không bị ảnh hưởng.",
        ):
            return
        self.store.clear()
        self.result = None
        self.selected_path.set("")
        for variable in self.card_vars.values():
            variable.set("—")
        for tree in (
            self.ke_tree,
            self.viet_tree,
            self.dang_tree,
            self.reconciliation_tree,
            self.issue_tree,
            self.recovery_tree,
            self.publish_error_tree,
            self.publish_today_tree,
        ):
            tree.delete(*tree.get_children())
        self.category_combo.configure(values=["Tất cả"])
        self.filter_category.set("Tất cả")
        self.filter_level.set("Tất cả")
        self.filter_text.set("")
        self.issue_count_label.configure(text="0 dòng")
        self.recovery_count_text.set("0 dòng sẵn sàng sao chép")
        self.publish_review = {"errors": [], "posted_today": [], "retry_rows": []}
        self.publish_id_updates.clear()
        self.publish_review_count.set("Bấm Phân tích để đọc lỗi đăng bài")
        for variable in self.publish_detail_vars.values():
            variable.set("")
        self.publish_new_id.set("")
        self.health_title.set("CHƯA PHÂN TÍCH")
        self.health_detail.set("Chọn file Excel và bấm Phân tích")
        self.pipeline_text.set("—")
        self._render_health({"status": "CHƯA PHÂN TÍCH", "detail": "Chọn file Excel và bấm Phân tích", "level": "pending"})
        self.export_button.configure(state="disabled")
        self.recover_button.configure(state="disabled")
        self._update_flow_buttons()
        self.status_text.set("Đã xóa phiên phân tích trong app")

    def _on_close(self) -> None:
        if self._flow_process is not None:
            flow_name = "flow hiện tại"
            if self._flow_running_key:
                flow_name = flow_by_key(self._flow_running_key).name
            messagebox.showwarning(
                "Flow đang chạy",
                f"{flow_name} vẫn đang chạy.\n\n"
                "Hãy chờ flow kết thúc rồi đóng app để tránh mất nhật ký và trạng thái theo dõi.",
            )
            return
        self.destroy()


def run_app() -> None:
    app = ExcelAuditApp()
    app.mainloop()
