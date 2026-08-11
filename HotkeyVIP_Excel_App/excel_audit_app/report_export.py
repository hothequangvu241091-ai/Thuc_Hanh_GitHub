from __future__ import annotations

import json
import os
import shutil
import subprocess
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any

from .excel_io import OpenXmlWorkbook, file_fingerprint, same_fingerprint


class SourceChangedError(RuntimeError):
    pass


class ExportError(RuntimeError):
    pass


def suggested_output_path(source_path: str | Path) -> Path:
    source = Path(source_path).resolve()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return source.with_name(f"{source.stem}_da_kiem_tra_{timestamp}{source.suffix}")


def suggested_recovery_path(source_path: str | Path) -> Path:
    source = Path(source_path).resolve()
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return source.with_name(f"{source.stem}_khoi_phuc_DANG_BAI_{timestamp}{source.suffix}")


class _ReportPayloadBuilder:
    def __init__(self, width: int = 11):
        self.width = width
        self.rows: list[list[Any]] = []
        self.merges: list[str] = []
        self.section_rows: list[int] = []
        self.header_rows: list[int] = []
        self.total_rows: list[int] = []
        self.status_cells: list[str] = []
        self.center_ranges: list[str] = []
        self.filter_ranges: list[str] = []
        self.detail_header_row = 0
        self.detail_last_row = 0

    @property
    def row_number(self) -> int:
        return len(self.rows) + 1

    def _append(self, values: list[Any]) -> int:
        row_number = self.row_number
        row = list(values[: self.width])
        row.extend([None] * (self.width - len(row)))
        self.rows.append(row)
        return row_number

    def blank(self) -> None:
        self._append([])

    def title(self, value: str) -> None:
        row = self._append([value])
        self.merges.append(f"A{row}:{_column_letter(self.width)}{row}")

    def metadata(self, label: str, value: Any) -> None:
        row = self._append([label, value])
        self.merges.append(f"B{row}:{_column_letter(self.width)}{row}")

    def section(
        self,
        title: str,
        headers: list[str],
        data_rows: list[list[Any]],
        total_row: list[Any] | None = None,
        *,
        status_column: int | None = None,
        detail: bool = False,
        center_columns: list[int] | None = None,
    ) -> None:
        section_row = self._append([title])
        self.section_rows.append(section_row)
        self.merges.append(f"A{section_row}:{_column_letter(len(headers))}{section_row}")
        header_row = self._append(headers)
        self.header_rows.append(header_row)
        if detail:
            self.detail_header_row = header_row
        data_start_row = self.row_number
        for values in data_rows:
            row = self._append(values)
            if status_column is not None:
                self.status_cells.append(f"{_column_letter(status_column + 1)}{row}")
        if total_row is not None:
            row = self._append(total_row)
            self.total_rows.append(row)
            if status_column is not None:
                self.status_cells.append(f"{_column_letter(status_column + 1)}{row}")
        data_end_row = self.row_number - 1
        if data_end_row >= data_start_row:
            if center_columns is not None:
                for column in center_columns:
                    letter = _column_letter(column)
                    self.center_ranges.append(
                        f"{letter}{data_start_row}:{letter}{data_end_row}"
                    )
            elif detail:
                self.center_ranges.append(f"C{data_start_row}:C{data_end_row}")
            elif len(headers) >= 2:
                last_column = _column_letter(len(headers))
                self.center_ranges.append(
                    f"B{data_start_row}:{last_column}{data_end_row}"
                )
        if detail:
            self.detail_last_row = data_end_row
            last_column = _column_letter(len(headers))
            self.filter_ranges.append(
                f"A{header_row}:{last_column}{max(header_row, data_end_row)}"
            )
        self.blank()


def _column_letter(index: int) -> str:
    result: list[str] = []
    while index:
        index, remainder = divmod(index - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def _build_report_payload(result: dict[str, Any]) -> dict[str, Any]:
    summaries = result["summaries"]
    overall = result.get("overall", {})
    builder = _ReportPayloadBuilder(width=18)
    builder.title("BÁO CÁO ĐỐI SOÁT NỘI DUNG EXCEL")
    builder.metadata("File nguồn", result.get("source_name", ""))
    builder.metadata("Thời gian phân tích", result.get("analyzed_at", ""))
    builder.metadata("Tình trạng tổng thể", overall.get("status", ""))
    builder.metadata("Lỗi dữ liệu cần xử lý", overall.get("error_count", 0))
    builder.metadata("Cần khôi phục vào DANG_BAI", overall.get("recovery_count", 0))
    builder.metadata("Chưa chuyển sang DANG_BAI", overall.get("pending_count", 0))
    builder.blank()

    ke_headers = [
        "Tên miền", "Tổng bài", "Combo 4 đủ", "Combo 4 thiếu", "URL hợp lệ",
        "Đã viết", "URL trống", "URL sai/khác", "Dữ liệu vấn đề",
        "Nhóm trùng", "Dòng trùng", "KE có - VIET thiếu",
    ]
    ke_keys = [
        "domain", "total_rows", "combo4_complete", "combo4_missing", "url_valid",
        "url_written", "url_blank", "url_other", "problem_rows",
        "duplicate_groups", "duplicate_rows", "missing_in_viet",
    ]
    builder.section(
        "1. PHÂN TÍCH KE_HOACH",
        ke_headers,
        [[row[key] for key in ke_keys] for row in summaries["ke_hoach"]["rows"]],
        [summaries["ke_hoach"]["total"][key] for key in ke_keys],
    )

    viet_headers = [
        "Tên miền", "Tổng dòng", "Combo 4 đủ", "Combo 4 thiếu", "Hoàn tất OK",
        "OK + đủ Word + 2 ảnh", "Đã đăng, đã xóa tài nguyên", "Cần khôi phục DANG",
        "Thiếu tài nguyên bất thường", "Chưa hoàn tất", "Nhóm combo 4 trùng",
        "Dòng trùng", "VIET có - KE thiếu",
    ]
    viet_keys = [
        "domain", "total_rows", "combo4_complete", "combo4_missing", "completed_ok",
        "completed_with_assets", "archived_posted_no_assets", "recovery_no_assets",
        "unexplained_no_assets", "not_completed", "duplicate_groups", "duplicate_rows",
        "missing_in_ke",
    ]
    builder.section(
        "2. PHÂN TÍCH VIET_BAI",
        viet_headers,
        [[row[key] for key in viet_keys] for row in summaries["viet_bai"]["rows"]],
        [summaries["viet_bai"]["total"][key] for key in viet_keys],
    )

    dang_headers = [
        "Tên miền", "Tổng dòng", "Combo 4 đủ", "Combo 4 thiếu", "Đã đăng",
        "Có trong VIET", "Có URL, chưa đăng, đủ tài nguyên", "DANG có - VIET thiếu",
        "VIET chưa có trong DANG", "Cần khôi phục DANG", "Chênh lệch phân loại",
    ]
    dang_keys = [
        "domain", "total_rows", "combo4_complete", "combo4_missing", "posted",
        "in_viet", "url_not_posted_full_assets", "dang_missing_viet", "viet_missing_dang",
        "ke_url_missing_dang", "classification_difference",
    ]
    builder.section(
        "3. PHÂN TÍCH DANG_BAI",
        dang_headers,
        [[row[key] for key in dang_keys] for row in summaries["dang_bai"]["rows"]],
        [summaries["dang_bai"]["total"][key] for key in dang_keys],
    )

    rec_headers = [
        "Tên miền", "Tổng KE", "Tổng VIET", "KE có - VIET thiếu",
        "VIET có - KE thiếu", "Đã có trong DANG", "Cần khôi phục DANG",
        "Chưa chuyển DANG", "DANG có - VIET thiếu", "VIET thiếu Combo 4",
        "Chênh lệch", "Trạng thái",
    ]
    rec_keys = [
        "domain", "ke_total", "viet_total", "ke_missing_viet", "viet_missing_ke",
        "in_dang", "recovery_dang", "pending_dang", "dang_missing_viet",
        "viet_combo4_missing", "difference", "status",
    ]
    builder.section(
        "4. ĐỐI SOÁT BA SHEET",
        rec_headers,
        [[row[key] for key in rec_keys] for row in summaries["reconciliation"]["rows"]],
        [summaries["reconciliation"]["total"][key] for key in rec_keys],
        status_column=11,
    )

    detail_headers = [
        "Nhóm", "Loại chi tiết", "Sheet nguồn", "Dòng nguồn", "Sheet đối chiếu",
        "Dòng đối chiếu", "Tên miền", "Tiêu đề SEO", "H1", "Từ khóa/Tiêu đề",
        "Chi tiết",
    ]
    level_labels = {
        "error": "Lỗi dữ liệu",
        "recovery": "Cần khôi phục",
        "pending": "Chưa chuyển",
        "info": "Đã đăng",
    }
    detail_rows = [
        [
            level_labels.get(item.get("level", "error"), item.get("level", "")),
            item["category"], item["sheet"], item["row"], item.get("target_sheet", ""),
            item.get("target_row", ""), item["domain"], item["title"], item["h1"],
            item["keyword"], item["detail"],
        ]
        for item in result.get("details", result.get("issues", []))
    ]
    builder.section(
        "5. DANH SÁCH CHI TIẾT ĐỐI SOÁT",
        detail_headers,
        detail_rows,
        detail=True,
        center_columns=[4, 6],
    )

    recovery = result.get("recovery", {})
    recovery_headers = [
        header or f"Cột {index}"
        for index, header in enumerate(recovery.get("headers", []), start=1)
    ]
    recovery_rows = [item.get("values", []) for item in recovery.get("rows", [])]
    if recovery_headers:
        builder.section(
            "6. DỮ LIỆU KHÔI PHỤC - DÁN TRỰC TIẾP VÀO DANG_BAI",
            recovery_headers,
            recovery_rows,
            detail=False,
            center_columns=[],
        )

    status_updates = result.get("status_updates", {})
    non_empty_updates = [
        {"row": int(row), "status": status}
        for row, status in status_updates.items()
        if status
    ]
    last_data_row = max((int(row) for row in status_updates), default=1)
    ke_info = result["sheet_info"]["ke_hoach"]
    return {
        "report": {
            "rows": builder.rows,
            "row_count": len(builder.rows),
            "column_count": builder.width,
            "merges": builder.merges,
            "section_rows": builder.section_rows,
            "header_rows": builder.header_rows,
            "total_rows": builder.total_rows,
            "status_cells": builder.status_cells,
            "center_ranges": builder.center_ranges,
            "filter_ranges": builder.filter_ranges,
            "ok_status": "KHỚP",
            "detail_header_row": builder.detail_header_row,
            "detail_last_row": builder.detail_last_row,
            "last_column": _column_letter(builder.width),
            "column_widths": [26, 34, 22, 18, 26, 22, 26, 34, 15, 38, 30, 30, 30, 16, 18, 22, 16, 24],
        },
        "ke_hoach": {
            "sheet_name": result["resolved_sheets"]["ke_hoach"],
            "header_row": int(ke_info.get("header_row", 1)),
            "status_column": ke_info.get("status_column"),
            "max_column": int(ke_info.get("max_column", 0)),
            "last_data_row": last_data_row,
            "status_header": ke_info.get("status_column_header", "Trạng thái nguồn"),
            "owned_statuses": ["Bài viết trùng", "Dữ liệu có vấn đề"],
            "updates": non_empty_updates,
        },
    }


def _powershell_executable() -> str:
    system_root = os.environ.get("SystemRoot", r"C:\Windows")
    candidate = Path(system_root) / "System32" / "WindowsPowerShell" / "v1.0" / "powershell.exe"
    return str(candidate if candidate.exists() else Path("powershell.exe"))


def _run_excel_export(workbook_path: Path, payload_path: Path) -> None:
    script = Path(__file__).with_name("export_with_excel.ps1")
    if not script.exists():
        raise ExportError("Thiếu thành phần xuất Excel của app")
    command = [
        _powershell_executable(),
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-WorkbookPath",
        str(workbook_path),
        "-PayloadPath",
        str(payload_path),
    ]
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=240,
            creationflags=creation_flags,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise ExportError("Excel mất quá nhiều thời gian khi xuất; file tạm đã bị hủy") from exc
    if completed.returncode != 0:
        message = (completed.stderr or completed.stdout or "Excel không thể ghi file").strip()
        raise ExportError(message)


def _run_excel_recovery(workbook_path: Path, payload_path: Path) -> None:
    script = Path(__file__).with_name("recover_dang_bai_with_excel.ps1")
    if not script.exists():
        raise ExportError("Thiếu thành phần khôi phục DANG_BAI của app")
    command = [
        _powershell_executable(),
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        str(script),
        "-WorkbookPath",
        str(workbook_path),
        "-PayloadPath",
        str(payload_path),
    ]
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=240,
            creationflags=creation_flags,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        raise ExportError(
            "Excel mất quá nhiều thời gian khi khôi phục DANG_BAI; file tạm đã bị hủy"
        ) from exc
    if completed.returncode != 0:
        message = (completed.stderr or completed.stdout or "Excel không thể khôi phục DANG_BAI").strip()
        raise ExportError(message)


def _semantic_snapshot(
    workbook: OpenXmlWorkbook,
    requested_sheet: str,
    excluded_column: int | None = None,
) -> tuple[Any, ...]:
    sheet = workbook.find_sheet(requested_sheet)
    if sheet is None:
        raise ExportError(f"File sau khi xuất bị thiếu sheet {requested_sheet}")
    return workbook.semantic_sheet_snapshot(
        sheet,
        excluded_columns={excluded_column} if excluded_column else set(),
    )


def _verify_export(source: Path, output: Path, result: dict[str, Any]) -> None:
    if not zipfile.is_zipfile(output):
        raise ExportError("File xuất không còn là workbook Excel hợp lệ")
    with zipfile.ZipFile(output, "r") as archive:
        bad_part = archive.testzip()
        if bad_part:
            raise ExportError(f"File xuất có thành phần ZIP bị lỗi: {bad_part}")

    source_book = OpenXmlWorkbook(source)
    output_book = OpenXmlWorkbook(output)
    if output_book.find_sheet("Tong_all") is None:
        raise ExportError("File xuất không có sheet Tong_all")

    ke_info = result["sheet_info"]["ke_hoach"]
    status_column = ke_info.get("status_column") or (int(ke_info.get("max_column", 0)) + 1)
    for sheet_name in ("VIET_BAI", "DANG_BAI", "KE_HOACH"):
        excluded = int(status_column) if sheet_name == "KE_HOACH" else None
        before = _semantic_snapshot(source_book, sheet_name, excluded)
        after = _semantic_snapshot(output_book, sheet_name, excluded)
        if before != after:
            raise ExportError(
                f"Kiểm tra an toàn thất bại: dữ liệu sheet {sheet_name} đã thay đổi ngoài phạm vi cho phép"
            )

    with zipfile.ZipFile(source, "r") as source_zip, zipfile.ZipFile(output, "r") as output_zip:
        vba_path = "xl/vbaProject.bin"
        if vba_path in source_zip.namelist():
            if vba_path not in output_zip.namelist():
                raise ExportError("File xuất bị mất VBA/macro")
            if len(output_zip.read(vba_path)) == 0:
                raise ExportError("Thành phần VBA/macro trong file xuất bị rỗng")


def _build_recovery_payload(source: Path, result: dict[str, Any]) -> dict[str, Any]:
    recovery = result.get("recovery", {})
    headers = list(recovery.get("headers", []))
    rows = [list(item.get("values", [])) for item in recovery.get("rows", [])]
    if not rows:
        raise ExportError("Không có dòng nào cần khôi phục vào DANG_BAI")
    if int(result.get("overall", {}).get("error_count", 0)) != 0:
        raise ExportError(
            "File đang có lỗi dữ liệu. Hãy xử lý lỗi dữ liệu trước khi tự động khôi phục DANG_BAI"
        )
    if not headers:
        raise ExportError("Không đọc được cấu trúc cột của DANG_BAI")
    column_count = len(headers)
    if any(len(row) != column_count for row in rows):
        raise ExportError("Dữ liệu khôi phục không khớp số cột của DANG_BAI")

    workbook = OpenXmlWorkbook(source)
    requested_name = result.get("resolved_sheets", {}).get("dang_bai", "DANG_BAI")
    sheet = workbook.find_sheet(str(requested_name))
    if sheet is None:
        raise ExportError("File nguồn không còn sheet DANG_BAI")
    table = workbook.read_sheet(sheet)
    snapshot = workbook.semantic_sheet_snapshot(sheet)
    last_data_row = max((int(item[0]) for item in snapshot), default=table.header_row)
    return {
        "dang_bai": {
            "sheet_name": sheet.name,
            "header_row": table.header_row,
            "last_data_row": last_data_row,
            "column_count": column_count,
            "row_count": len(rows),
            "rows": rows,
        }
    }


def _verify_recovery(
    source: Path,
    output: Path,
    result: dict[str, Any],
    recovery_count: int,
) -> None:
    if not zipfile.is_zipfile(output):
        raise ExportError("File khôi phục không còn là workbook Excel hợp lệ")
    with zipfile.ZipFile(output, "r") as archive:
        bad_part = archive.testzip()
        if bad_part:
            raise ExportError(f"File khôi phục có thành phần ZIP bị lỗi: {bad_part}")

    source_book = OpenXmlWorkbook(source)
    output_book = OpenXmlWorkbook(output)
    for sheet_name in ("KE_HOACH", "VIET_BAI"):
        before = _semantic_snapshot(source_book, sheet_name)
        after = _semantic_snapshot(output_book, sheet_name)
        if before != after:
            raise ExportError(
                f"Kiểm tra an toàn thất bại: sheet {sheet_name} đã bị thay đổi"
            )

    dang_before = _semantic_snapshot(source_book, "DANG_BAI")
    dang_after = _semantic_snapshot(output_book, "DANG_BAI")
    if dang_after[: len(dang_before)] != dang_before:
        raise ExportError(
            "Kiểm tra an toàn thất bại: dữ liệu DANG_BAI cũ đã bị thay đổi"
        )
    if len(dang_after) != len(dang_before) + recovery_count:
        raise ExportError(
            "Kiểm tra an toàn thất bại: số dòng được thêm vào DANG_BAI không đúng"
        )

    with zipfile.ZipFile(source, "r") as source_zip, zipfile.ZipFile(output, "r") as output_zip:
        vba_path = "xl/vbaProject.bin"
        if vba_path in source_zip.namelist():
            if vba_path not in output_zip.namelist():
                raise ExportError("File khôi phục bị mất VBA/macro")
            if len(output_zip.read(vba_path)) == 0:
                raise ExportError("Thành phần VBA/macro trong file khôi phục bị rỗng")

    from .analysis import analyze_workbook

    checked = analyze_workbook(output)
    if int(checked.get("overall", {}).get("recovery_count", -1)) != 0:
        raise ExportError(
            "Kiểm tra sau khôi phục thất bại: vẫn còn bài cần khôi phục trong DANG_BAI"
        )
    before_total = int(result["summaries"]["dang_bai"]["total"]["total_rows"])
    after_total = int(checked["summaries"]["dang_bai"]["total"]["total_rows"])
    if after_total != before_total + recovery_count:
        raise ExportError(
            "Kiểm tra sau khôi phục thất bại: tổng số dòng DANG_BAI không tăng đúng"
        )


def export_result(
    source_path: str | Path,
    destination_path: str | Path,
    result: dict[str, Any],
) -> Path:
    source = Path(source_path).resolve()
    destination = Path(destination_path).resolve()
    if source == destination:
        raise ExportError("File xuất phải khác file nguồn")
    if destination.suffix.casefold() != source.suffix.casefold():
        raise ExportError("File xuất phải giữ nguyên phần mở rộng của file nguồn")
    current_fingerprint = file_fingerprint(source, include_hash=True)
    if not same_fingerprint(current_fingerprint, result.get("source_fingerprint", {})):
        raise SourceChangedError("File nguồn đã thay đổi từ sau lần phân tích gần nhất")

    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook_handle, workbook_name = tempfile.mkstemp(
        prefix=f".{destination.stem}_working_",
        suffix=destination.suffix,
        dir=destination.parent,
    )
    os.close(workbook_handle)
    working_copy = Path(workbook_name)
    payload_handle, payload_name = tempfile.mkstemp(
        prefix=f".{destination.stem}_payload_",
        suffix=".json",
        dir=destination.parent,
    )
    os.close(payload_handle)
    payload_path = Path(payload_name)
    try:
        shutil.copy2(source, working_copy)
        payload = _build_report_payload(result)
        payload_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8-sig")
        _run_excel_export(working_copy, payload_path)
        _verify_export(source, working_copy, result)
        os.replace(working_copy, destination)
    finally:
        payload_path.unlink(missing_ok=True)
        working_copy.unlink(missing_ok=True)
    return destination


def recover_dang_bai(
    source_path: str | Path,
    destination_path: str | Path,
    result: dict[str, Any],
) -> Path:
    source = Path(source_path).resolve()
    destination = Path(destination_path).resolve()
    if source == destination:
        raise ExportError("File khôi phục phải khác file nguồn")
    if destination.suffix.casefold() != source.suffix.casefold():
        raise ExportError("File khôi phục phải giữ nguyên phần mở rộng của file nguồn")
    fingerprint_before = file_fingerprint(source, include_hash=True)
    if not same_fingerprint(fingerprint_before, result.get("source_fingerprint", {})):
        raise SourceChangedError("File nguồn đã thay đổi từ sau lần phân tích gần nhất")

    payload = _build_recovery_payload(source, result)
    recovery_count = int(payload["dang_bai"]["row_count"])
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook_handle, workbook_name = tempfile.mkstemp(
        prefix=f".{destination.stem}_working_",
        suffix=destination.suffix,
        dir=destination.parent,
    )
    os.close(workbook_handle)
    working_copy = Path(workbook_name)
    payload_handle, payload_name = tempfile.mkstemp(
        prefix=f".{destination.stem}_recovery_",
        suffix=".json",
        dir=destination.parent,
    )
    os.close(payload_handle)
    payload_path = Path(payload_name)
    try:
        shutil.copy2(source, working_copy)
        payload_path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8-sig")
        _run_excel_recovery(working_copy, payload_path)
        _verify_recovery(source, working_copy, result, recovery_count)
        if not same_fingerprint(
            file_fingerprint(source, include_hash=True), fingerprint_before
        ):
            raise ExportError("File nguồn đã thay đổi trong lúc tạo bản khôi phục")
        os.replace(working_copy, destination)
    finally:
        payload_path.unlink(missing_ok=True)
        working_copy.unlink(missing_ok=True)
    return destination
