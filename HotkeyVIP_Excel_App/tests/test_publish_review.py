from __future__ import annotations

import unittest
from datetime import date

from excel_audit_app.excel_io import SheetTable
from excel_audit_app.publish_review import (
    build_publish_review_from_tables,
    build_retry_publish_plan,
)


def table(name: str, headers: list[str], rows: list[list[object]]) -> SheetTable:
    return SheetTable(
        name=name,
        xml_path="",
        header_row=1,
        headers={index: value for index, value in enumerate(headers, start=1)},
        rows=[
            (row_number, {index: value for index, value in enumerate(values, start=1)})
            for row_number, values in enumerate(rows, start=2)
        ],
        max_column=len(headers),
    )


class PublishReviewTest(unittest.TestCase):
    def test_error_retry_urls_and_today(self) -> None:
        viet_headers = [
            "Tên miền", "Tiêu đề SEO", "H1", "Từ khóa", "URL GPT gốc", "URL ChatGPT"
        ]
        viet = table(
            "VIET_BAI",
            viet_headers,
            [["a.vn", "SEO A", "H1 A", "KW A", "https://gpt.example/a", "https://chat.example/a"]],
        )
        dang_headers = [
            "Tiêu đề", "Tiêu đề SEO", "H1", "Trạng thái đăng", "Tên miền", "Danh mục",
            "ID CMS", "URL đã đăng", "Đường dẫn Word", "Đường dẫn ảnh 1", "Đường dẫn ảnh 2",
            "Thời gian đăng", "Lỗi đăng",
        ]
        dang = table(
            "DANG_BAI",
            dang_headers,
            [
                ["KW A", "SEO A", "H1 A", "LỖI KIỂM TRA", "a.vn", "Cat", "", "", "a.docx", "", "", "11/08/2026 09:00:00", "Word lỗi"],
                ["KW B", "SEO B", "H1 B", "ĐÃ ĐĂNG", "b.vn", "Cat", "123", "", "b.docx", "", "", "11/08/2026 10:00:00", ""],
                ["KW C", "SEO C", "H1 C", "ĐÃ ĐĂNG", "c.vn", "Cat", "LỖI ID", "", "c.docx", "", "", "11/08/2026 11:00:00", ""],
            ],
        )
        result = build_publish_review_from_tables(dang, viet, date(2026, 8, 11))
        self.assertEqual([item["row"] for item in result["errors"]], [2, 4])
        self.assertEqual([item["row"] for item in result["retry_rows"]], [2])
        self.assertEqual([item["row"] for item in result["posted_today"]], [3])
        self.assertEqual(result["errors"][0]["chat_url"], "https://chat.example/a")
        plan = build_retry_publish_plan(result)
        self.assertEqual(plan["mode"], "explicit_error_rows")
        self.assertEqual([item["row"] for item in plan["selected_rows"]], [2])


if __name__ == "__main__":
    unittest.main()
