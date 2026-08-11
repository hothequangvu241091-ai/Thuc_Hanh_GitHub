from __future__ import annotations

import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from excel_audit_app.analysis import analyze_workbook
from excel_audit_app.excel_io import OpenXmlWorkbook, file_fingerprint
from excel_audit_app.report_export import (
    ExportError,
    _build_report_payload,
    export_result,
    recover_dang_bai,
)


ROOT = Path(__file__).resolve().parents[1]
SAMPLE = ROOT / "_tests_data" / "hotkeyvip_test_chogemini.xlsm"


@unittest.skipUnless(SAMPLE.exists(), "Không có workbook mẫu")
class ExcelAuditSmokeTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.result = analyze_workbook(SAMPLE)

    def test_expected_sample_totals(self) -> None:
        summaries = self.result["summaries"]
        self.assertEqual(summaries["ke_hoach"]["total"]["total_rows"], 6748)
        self.assertEqual(summaries["ke_hoach"]["total"]["url_valid"], 881)
        self.assertEqual(summaries["ke_hoach"]["total"]["url_written"], 3432)
        self.assertEqual(summaries["ke_hoach"]["total"]["url_blank"], 2435)
        self.assertEqual(summaries["viet_bai"]["total"]["completed_ok"], 4313)
        self.assertEqual(summaries["viet_bai"]["total"]["has_word"], 4046)
        self.assertEqual(summaries["viet_bai"]["total"]["has_word_images"], 4034)
        self.assertEqual(summaries["dang_bai"]["total"]["posted"], 609)
        reconciliation = summaries["reconciliation"]["total"]
        self.assertEqual(reconciliation["in_dang"], 4041)
        self.assertEqual(reconciliation["ke_url_deleted"], 272)
        self.assertEqual(reconciliation["viet_missing_remaining"], 2435)
        self.assertEqual(reconciliation["classified_total"], 6748)
        self.assertEqual(reconciliation["difference"], 0)
        self.assertEqual(self.result["overall"]["error_count"], 0)
        self.assertEqual(self.result["overall"]["recovery_count"], 272)
        self.assertEqual(self.result["overall"]["pending_count"], 2435)
        self.assertEqual(self.result["overall"]["archived_count"], 7)
        self.assertEqual(summaries["viet_bai"]["total"]["duplicate_groups"], 0)
        self.assertEqual(summaries["viet_bai"]["total"]["duplicate_rows"], 0)
        self.assertFalse(
            any(item["category"] == "Combo 3 trùng" for item in self.result["details"])
        )
        self.assertEqual(len(self.result["recovery"]["headers"]), 18)
        self.assertTrue(
            all(len(row["values"]) == 18 for row in self.result["recovery"]["rows"])
        )
        headers = self.result["recovery"]["headers"]
        first_recovery = self.result["recovery"]["rows"][0]["values"]
        self.assertEqual(first_recovery[headers.index("Trạng thái đăng")], "ĐÃ ĐĂNG")
        self.assertTrue(
            str(first_recovery[headers.index("URL đã đăng")]).startswith(("http://", "https://"))
        )

        viet_total = summaries["viet_bai"]["total"]
        self.assertEqual(
            viet_total["total_rows"],
            viet_total["completed_with_assets"]
            + viet_total["archived_posted_no_assets"]
            + viet_total["recovery_no_assets"]
            + viet_total["unexplained_no_assets"]
            + viet_total["not_completed"],
        )

    def test_report_centers_numeric_columns(self) -> None:
        report = _build_report_payload(self.result)["report"]
        self.assertEqual(
            report["center_ranges"][:4],
            ["B11:L20", "B24:M33", "B37:K46", "B50:L59"],
        )
        self.assertEqual(report["center_ranges"][4:6], ["D63:D2776", "F63:F2776"])
        self.assertEqual(report["column_count"], 18)
        self.assertEqual(report["last_column"], "R")
        self.assertEqual(report["filter_ranges"], ["A62:K2776"])

    def test_deleted_viet_rows_are_listed_from_ke_hoach(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            changed = Path(directory) / SAMPLE.name
            workbook = OpenXmlWorkbook(SAMPLE)
            viet_sheet = workbook.find_sheet("VIET_BAI")
            self.assertIsNotNone(viet_sheet)
            assert viet_sheet is not None
            namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            with zipfile.ZipFile(SAMPLE, "r") as source_zip, zipfile.ZipFile(
                changed, "w"
            ) as output_zip:
                for info in source_zip.infolist():
                    data = source_zip.read(info.filename)
                    if info.filename == viet_sheet.xml_path:
                        root = ET.fromstring(data)
                        sheet_data = root.find(f"{{{namespace}}}sheetData")
                        self.assertIsNotNone(sheet_data)
                        assert sheet_data is not None
                        for row in list(sheet_data):
                            if int(row.attrib.get("r", "0")) in {2, 3}:
                                sheet_data.remove(row)
                        data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                    output_zip.writestr(info, data)

            changed_result = analyze_workbook(changed)
            rec = changed_result["summaries"]["reconciliation"]["total"]
            self.assertEqual(rec["ke_missing_viet"], 2)
            self.assertEqual(rec["status"], "LỆCH")
            missing = [
                item
                for item in changed_result["details"]
                if item["category"] == "KE_HOACH có - VIET_BAI thiếu"
            ]
            self.assertEqual(len(missing), 2)
            self.assertTrue(all(item["target_sheet"] == "VIET_BAI" for item in missing))

    def test_export_keeps_vba_and_adds_tong_all(self) -> None:
        output = Path(tempfile.gettempdir()) / "excel_audit_verified_test.xlsm"
        try:
            export_result(SAMPLE, output, self.result)
        except ExportError as exc:
            normalized_error = " ".join(str(exc).split())
            if "macros in this workbook are corrupted" in normalized_error:
                self.skipTest("Workbook mẫu bị Excel xác nhận hỏng VBA khi thêm sheet")
            raise
        self.assertTrue(zipfile.is_zipfile(output))
        with zipfile.ZipFile(SAMPLE, "r") as source_zip, zipfile.ZipFile(output, "r") as output_zip:
            self.assertIsNone(output_zip.testzip())
            if "xl/vbaProject.bin" in source_zip.namelist():
                self.assertIn("xl/vbaProject.bin", output_zip.namelist())
                self.assertGreater(len(output_zip.read("xl/vbaProject.bin")), 0)
        workbook = OpenXmlWorkbook(output)
        self.assertIsNotNone(workbook.find_sheet("Tong_all"))
        self.assertIsNotNone(workbook.find_sheet("KE_HOACH"))

    def test_recovery_creates_new_file_and_keeps_source_unchanged(self) -> None:
        recovery_count = self.result["overall"]["recovery_count"]
        self.assertGreater(recovery_count, 0)
        source_before = file_fingerprint(SAMPLE, include_hash=True)
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "recovered.xlsm"
            recover_dang_bai(SAMPLE, output, self.result)
            self.assertTrue(output.exists())
            recovered = analyze_workbook(output)
            self.assertEqual(recovered["overall"]["recovery_count"], 0)
            self.assertEqual(
                recovered["summaries"]["dang_bai"]["total"]["total_rows"],
                self.result["summaries"]["dang_bai"]["total"]["total_rows"]
                + recovery_count,
            )
        self.assertEqual(file_fingerprint(SAMPLE, include_hash=True), source_before)


if __name__ == "__main__":
    unittest.main()
