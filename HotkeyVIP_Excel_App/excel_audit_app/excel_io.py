from __future__ import annotations

import hashlib
import posixpath
import re
import unicodedata
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable
from xml.etree import ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"


def qn(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def normalize_spaces(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_text(value: Any) -> str:
    """Chuẩn hóa để so sánh; giữ nguyên dấu tiếng Việt."""
    return normalize_spaces(value).casefold()


def normalize_header(value: Any) -> str:
    """Chuẩn hóa tiêu đề cột theo kiểu không dấu, bỏ ký tự phân cách."""
    text = unicodedata.normalize("NFD", normalize_text(value))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9]+", "", text)


def normalize_sheet_name(value: Any) -> str:
    return normalize_header(value)


def is_blank(value: Any) -> bool:
    return normalize_spaces(value) == ""


def is_valid_url(value: Any) -> bool:
    return bool(re.match(r"^https?://", normalize_spaces(value), flags=re.IGNORECASE))


def column_index_from_ref(cell_ref: str) -> int:
    letters = re.match(r"[A-Za-z]+", cell_ref)
    if not letters:
        raise ValueError(f"Địa chỉ ô không hợp lệ: {cell_ref}")
    result = 0
    for char in letters.group(0).upper():
        result = result * 26 + (ord(char) - 64)
    return result


def column_letter(index: int) -> str:
    if index < 1:
        raise ValueError("Chỉ số cột phải lớn hơn hoặc bằng 1")
    result = []
    while index:
        index, remainder = divmod(index - 1, 26)
        result.append(chr(65 + remainder))
    return "".join(reversed(result))


def file_fingerprint(path: str | Path, include_hash: bool = True) -> dict[str, Any]:
    source = Path(path).resolve()
    stat = source.stat()
    result: dict[str, Any] = {
        "path": str(source),
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
    }
    if include_hash:
        digest = hashlib.sha256()
        with source.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
        result["sha256"] = digest.hexdigest()
    return result


def same_fingerprint(left: dict[str, Any], right: dict[str, Any]) -> bool:
    required = ("size", "mtime_ns")
    if any(left.get(key) != right.get(key) for key in required):
        return False
    if left.get("sha256") and right.get("sha256"):
        return left["sha256"] == right["sha256"]
    return True


@dataclass(slots=True)
class SheetInfo:
    name: str
    sheet_id: str
    relationship_id: str
    xml_path: str


@dataclass(slots=True)
class SheetTable:
    name: str
    xml_path: str
    header_row: int
    headers: dict[int, str]
    rows: list[tuple[int, dict[int, Any]]]
    max_column: int

    def find_column(self, aliases: Iterable[str], required: bool = True) -> int | None:
        normalized = {normalize_header(alias) for alias in aliases}
        for index, header in self.headers.items():
            if normalize_header(header) in normalized:
                return index
        if required:
            alias_text = ", ".join(aliases)
            raise WorkbookStructureError(
                f"Sheet {self.name} thiếu cột bắt buộc: {alias_text}"
            )
        return None

    def value(self, row: dict[int, Any], column: int | None) -> Any:
        if column is None:
            return ""
        return row.get(column, "")


class WorkbookStructureError(RuntimeError):
    pass


class OpenXmlWorkbook:
    """Bộ đọc XLSX/XLSM chỉ dùng thư viện chuẩn, không khởi động Excel."""

    def __init__(self, path: str | Path):
        self.path = Path(path).resolve()
        if self.path.suffix.casefold() not in {".xlsx", ".xlsm"}:
            raise WorkbookStructureError("App chỉ hỗ trợ file .xlsx và .xlsm")
        if not self.path.exists():
            raise WorkbookStructureError("Không tìm thấy file Excel đã chọn")
        if not zipfile.is_zipfile(self.path):
            raise WorkbookStructureError("File đã chọn không phải định dạng Excel hợp lệ")

        self.shared_strings: list[str] = []
        self.sheets: list[SheetInfo] = []
        self._load_index()

    def _load_index(self) -> None:
        with zipfile.ZipFile(self.path, "r") as archive:
            names = set(archive.namelist())
            if "xl/workbook.xml" not in names:
                raise WorkbookStructureError("File Excel thiếu xl/workbook.xml")

            if "xl/sharedStrings.xml" in names:
                root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
                for item in root.findall(qn(NS_MAIN, "si")):
                    text = "".join(node.text or "" for node in item.iter(qn(NS_MAIN, "t")))
                    self.shared_strings.append(text)

            relationships: dict[str, str] = {}
            rel_path = "xl/_rels/workbook.xml.rels"
            if rel_path not in names:
                raise WorkbookStructureError("File Excel thiếu quan hệ workbook")
            rel_root = ET.fromstring(archive.read(rel_path))
            for rel in rel_root.findall(qn(NS_REL_PKG, "Relationship")):
                rel_id = rel.attrib.get("Id", "")
                target = rel.attrib.get("Target", "")
                if target.startswith("/"):
                    full_path = target.lstrip("/")
                else:
                    full_path = posixpath.normpath(posixpath.join("xl", target))
                relationships[rel_id] = full_path

            workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
            sheets_node = workbook_root.find(qn(NS_MAIN, "sheets"))
            if sheets_node is None:
                raise WorkbookStructureError("Workbook không có danh sách sheet")
            for sheet in sheets_node.findall(qn(NS_MAIN, "sheet")):
                rel_id = sheet.attrib.get(qn(NS_REL_DOC, "id"), "")
                xml_path = relationships.get(rel_id)
                if not xml_path:
                    continue
                self.sheets.append(
                    SheetInfo(
                        name=sheet.attrib.get("name", ""),
                        sheet_id=sheet.attrib.get("sheetId", ""),
                        relationship_id=rel_id,
                        xml_path=xml_path,
                    )
                )

    @property
    def sheet_names(self) -> list[str]:
        return [sheet.name for sheet in self.sheets]

    def find_sheet(self, requested_name: str) -> SheetInfo | None:
        target = normalize_sheet_name(requested_name)
        for sheet in self.sheets:
            if normalize_sheet_name(sheet.name) == target:
                return sheet
        return None

    def require_sheets(self, requested_names: Iterable[str]) -> dict[str, SheetInfo]:
        found: dict[str, SheetInfo] = {}
        missing: list[str] = []
        for requested in requested_names:
            sheet = self.find_sheet(requested)
            if sheet is None:
                missing.append(requested)
            else:
                found[requested] = sheet
        if missing:
            raise WorkbookStructureError(
                "File Excel không có sheet theo yêu cầu: " + ", ".join(missing)
            )
        return found

    def read_sheet(self, sheet: SheetInfo) -> SheetTable:
        with zipfile.ZipFile(self.path, "r") as archive:
            if sheet.xml_path not in archive.namelist():
                raise WorkbookStructureError(f"Không đọc được dữ liệu sheet {sheet.name}")
            root = ET.fromstring(archive.read(sheet.xml_path))

        rows: list[tuple[int, dict[int, Any]]] = []
        max_column = 0
        sheet_data = root.find(qn(NS_MAIN, "sheetData"))
        if sheet_data is None:
            return SheetTable(sheet.name, sheet.xml_path, 1, {}, [], 0)

        for row_element in sheet_data.findall(qn(NS_MAIN, "row")):
            row_number = int(row_element.attrib.get("r", len(rows) + 1))
            values: dict[int, Any] = {}
            for cell in row_element.findall(qn(NS_MAIN, "c")):
                reference = cell.attrib.get("r", "")
                if not reference:
                    continue
                column = column_index_from_ref(reference)
                values[column] = self._read_cell(cell)
                max_column = max(max_column, column)
            rows.append((row_number, values))

        header_row, headers = self._detect_header(rows)
        data_rows = [(number, values) for number, values in rows if number > header_row]
        return SheetTable(
            name=sheet.name,
            xml_path=sheet.xml_path,
            header_row=header_row,
            headers=headers,
            rows=data_rows,
            max_column=max_column,
        )

    def semantic_sheet_snapshot(
        self,
        sheet: SheetInfo,
        excluded_columns: set[int] | None = None,
    ) -> tuple[Any, ...]:
        """Ảnh chụp dữ liệu/công thức để so sánh trước và sau khi xuất.

        Bỏ qua định dạng và các ô trống chỉ mang style; giữ số dòng, cột,
        giá trị và nội dung công thức. Nhờ đó Excel có thể tự lưu lại cấu
        trúc XML mà việc kiểm tra vẫn phát hiện mọi thay đổi dữ liệu thực.
        """
        excluded = excluded_columns or set()
        with zipfile.ZipFile(self.path, "r") as archive:
            if sheet.xml_path not in archive.namelist():
                raise WorkbookStructureError(f"Không đọc được dữ liệu sheet {sheet.name}")
            root = ET.fromstring(archive.read(sheet.xml_path))
        sheet_data = root.find(qn(NS_MAIN, "sheetData"))
        if sheet_data is None:
            return tuple()

        snapshot: list[Any] = []
        for row_element in sheet_data.findall(qn(NS_MAIN, "row")):
            row_number = int(row_element.attrib.get("r", "0"))
            cells: list[Any] = []
            for cell in row_element.findall(qn(NS_MAIN, "c")):
                reference = cell.attrib.get("r", "")
                if not reference:
                    continue
                column = column_index_from_ref(reference)
                if column in excluded:
                    continue
                formula = cell.find(qn(NS_MAIN, "f"))
                if formula is not None:
                    cells.append((column, "formula", formula.text or ""))
                    continue
                value = self._read_cell(cell)
                if is_blank(value):
                    continue
                cells.append((column, "value", value))
            if cells:
                snapshot.append((row_number, tuple(cells)))
        return tuple(snapshot)

    def _read_cell(self, cell: ET.Element) -> Any:
        cell_type = cell.attrib.get("t", "")
        if cell_type == "inlineStr":
            return "".join(node.text or "" for node in cell.iter(qn(NS_MAIN, "t")))

        value_node = cell.find(qn(NS_MAIN, "v"))
        value = "" if value_node is None or value_node.text is None else value_node.text
        if cell_type == "s":
            try:
                return self.shared_strings[int(value)]
            except (ValueError, IndexError):
                return ""
        if cell_type in {"str", "e", "d"}:
            return value
        if cell_type == "b":
            return value == "1"
        if value == "":
            return ""
        try:
            number = float(value)
            return int(number) if number.is_integer() else number
        except ValueError:
            return value

    @staticmethod
    def _detect_header(
        rows: list[tuple[int, dict[int, Any]]]
    ) -> tuple[int, dict[int, str]]:
        candidates = [(number, values) for number, values in rows[:20] if values]
        if not candidates:
            return 1, {}
        number, values = max(
            candidates,
            key=lambda item: sum(1 for value in item[1].values() if not is_blank(value)),
        )
        headers = {
            column: normalize_spaces(value)
            for column, value in values.items()
            if not is_blank(value)
        }
        return number, headers
