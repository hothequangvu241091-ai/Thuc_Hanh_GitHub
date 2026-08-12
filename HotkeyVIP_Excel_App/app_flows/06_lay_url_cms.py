import asyncio
import aiohttp
import ssl
import re
import os
import sys
import win32com.client as win32

APP_WORKBOOK = None

PROJECT_ROOT = os.environ.get("HOTKEYVIP_RUNTIME_ROOT", r"D:\CodexProjects\Hotkeyvip")
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
from hotkeyvip_config import EXCEL_FILE, PUBLISH_HEADERS, SHEET_CATEGORY_ID, SHEET_PUBLISH

# ================== CẤU HÌNH ==================
SHEET_DANGBAI = SHEET_PUBLISH
SHEET_TACGIA = SHEET_CATEGORY_ID
EXCEL_PATH = os.path.abspath(
    os.environ.get("HOTKEYVIP_SELECTED_EXCEL", str(EXCEL_FILE))
)

ROW_START = 2
ROW_END = 50000

TIMEOUT_S = 10
CONCURRENCY = 20


def clean_text(value):
    return str(value or "").strip()


def show_error(message):
    try:
        import tkinter as tk
        from tkinter import messagebox

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        messagebox.showerror("Xuất URL từ ID gặp lỗi", message, parent=root)
        root.destroy()
    except Exception:
        pass


def get_target_excel_workbook():
    """Chỉ trả về đúng hotkeyvip_test.xlsm đang mở; không tự mở Excel ẩn."""
    if APP_WORKBOOK is not None:
        workbook_path = os.path.abspath(clean_text(APP_WORKBOOK.FullName))
        if os.path.normcase(workbook_path) != os.path.normcase(EXCEL_PATH):
            raise RuntimeError(
                f"App truyền sai workbook. Cần: {EXCEL_PATH}; nhận: {workbook_path}"
            )
        return APP_WORKBOOK.Application, APP_WORKBOOK

    try:
        excel = win32.GetActiveObject("Excel.Application")
    except Exception as exc:
        raise RuntimeError(
            "Không kết nối được với Excel đang mở. "
            "Hãy mở hotkeyvip_test.xlsm rồi chạy lại."
        ) from exc

    try:
        for index in range(1, excel.Workbooks.Count + 1):
            workbook = excel.Workbooks(index)
            workbook_path = os.path.abspath(clean_text(workbook.FullName))
            if os.path.normcase(workbook_path) == os.path.normcase(EXCEL_PATH):
                return excel, workbook
    except Exception as exc:
        raise RuntimeError(
            "Đã thấy Excel nhưng không đọc được danh sách workbook đang mở."
        ) from exc

    raise RuntimeError(
        "Excel đang mở nhưng chưa mở đúng file hotkeyvip_test.xlsm.\n"
        f"Hãy mở file này rồi chạy lại:\n{EXCEL_PATH}"
    )


def normalize_header(value):
    text = clean_text(value).lower()
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def find_col(ws, header_name, header_row=1):
    target = normalize_header(header_name)

    last_col = ws.Cells(
        header_row,
        ws.Columns.Count
    ).End(-4159).Column  # xlToLeft

    for col in range(1, last_col + 1):
        current = normalize_header(ws.Cells(header_row, col).Value)

        if current == target:
            return col

    raise Exception(
        f"❌ Không tìm thấy cột '{header_name}' "
        f"trong sheet '{ws.Name}'"
    )


def normalize_domain(value):
    domain = clean_text(value).lower()
    domain = re.sub(r"^https?://", "", domain)
    domain = domain.strip("/")
    domain = re.sub(r"^www\.", "", domain)
    return domain


def normalize_category(value):
    text = clean_text(value).lower()
    text = text.replace("\n", " ").replace("\r", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def parse_int(value):
    try:
        return int(float(str(value).strip()))
    except (ValueError, TypeError):
        return None


def build_fake_url(domain, cat_id, post_id):
    domain = normalize_domain(domain)

    return f"https://{domain}/linkrutgon-{cat_id}-{post_id}.html"


async def fetch_final(session, url):
    try:
        async with session.head(
            url,
            allow_redirects=True
        ) as resp:
            return str(resp.url)

    except Exception:
        try:
            async with session.get(
                url,
                allow_redirects=True
            ) as resp:
                return str(resp.url)

        except Exception:
            return url


async def resolve_all(urls):
    ssl_ctx = ssl.create_default_context()
    ssl_ctx.check_hostname = False
    ssl_ctx.verify_mode = ssl.CERT_NONE

    timeout = aiohttp.ClientTimeout(
        total=None,
        sock_connect=TIMEOUT_S,
        sock_read=TIMEOUT_S
    )

    connector = aiohttp.TCPConnector(
        limit=CONCURRENCY,
        ssl=ssl_ctx
    )

    async with aiohttp.ClientSession(
        timeout=timeout,
        connector=connector
    ) as session:
        tasks = [
            fetch_final(session, url)
            for url in urls
        ]

        return await asyncio.gather(*tasks)


def main():
    # ================== KẾT NỐI EXCEL ĐANG MỞ ==================
    try:
        excel, wb = get_target_excel_workbook()
    except Exception as exc:
        print(f"Lỗi Excel: {exc}")
        show_error(str(exc))
        return

    try:
        ws = wb.Sheets(SHEET_DANGBAI)
    except Exception:
        print(f"❌ Không tìm thấy sheet '{SHEET_DANGBAI}'.")
        return

    try:
        ws_tg = wb.Sheets(SHEET_TACGIA)
    except Exception:
        print(f"❌ Không tìm thấy sheet '{SHEET_TACGIA}'.")
        return

    # ================== TÌM CỘT THEO TÊN ==================
    try:
        # Sheet dangbai
        COL_DOMAIN = find_col(ws, PUBLISH_HEADERS["domain"])
        COL_CATEGORY = find_col(ws, PUBLISH_HEADERS["category"])
        COL_POST_ID = find_col(ws, PUBLISH_HEADERS["cms_id"])
        COL_OUTPUT_URL = find_col(ws, PUBLISH_HEADERS["published_url"])

        # Sheet ID_tacgia
        COL_TG_DOMAIN = find_col(ws_tg, "Tên miền URL")
        COL_TG_CAT_ID = find_col(ws_tg, "ID")
        COL_TG_CATEGORY = find_col(ws_tg, "Danh mục")

    except Exception as e:
        print(e)
        return

    print("✅ Đã nhận diện cột theo tên:")
    print(
        f"- dangbai: "
        f"Tên miền={COL_DOMAIN}, "
        f"Danh mục={COL_CATEGORY}, "
            f"ID CMS={COL_POST_ID}, "
            f"URL đã đăng={COL_OUTPUT_URL}"
    )
    print(
        f"- ID_tacgia: "
        f"Tên miền URL={COL_TG_DOMAIN}, "
        f"ID={COL_TG_CAT_ID}, "
        f"Danh mục={COL_TG_CATEGORY}"
    )

    # ================== TẠO MAP DANH MỤC ==================
    category_map = {}

    last_tg_row = ws_tg.Cells(
        ws_tg.Rows.Count,
        COL_TG_DOMAIN
    ).End(-4162).Row  # xlUp

    def read_map_column(col):
        if last_tg_row < 2:
            return []
        raw = ws_tg.Range(
            ws_tg.Cells(2, col),
            ws_tg.Cells(last_tg_row, col),
        ).Value2
        if isinstance(raw, tuple):
            return [
                item[0] if isinstance(item, tuple) else item
                for item in raw
            ]
        return [raw]

    map_domains = read_map_column(COL_TG_DOMAIN)
    map_categories = read_map_column(COL_TG_CATEGORY)
    map_ids = read_map_column(COL_TG_CAT_ID)

    for domain_raw, category_raw, cat_id_raw in zip(
        map_domains,
        map_categories,
        map_ids,
    ):
        domain = normalize_domain(domain_raw)
        category = normalize_category(category_raw)
        cat_id = parse_int(cat_id_raw)

        if domain and category and cat_id:
            category_map[(domain, category)] = cat_id

    print(
        f"✅ Đã tạo map cho "
        f"{len(category_map)} tên miền/danh mục."
    )

    # ================== XỬ LÝ SHEET DANGBAI ==================
    rows = []
    fake_urls = []

    skipped_no_id = 0
    skipped_has_url = 0
    skipped_empty = 0
    no_domain = 0
    no_category = 0
    no_category_id = 0

    # Chỉ đọc Excel theo 4 vùng lớn, không gọi COM riêng cho từng ô/dòng.
    # Dòng cuối được xác định từ bốn cột liên quan và vẫn giới hạn bởi ROW_END.
    last_data_row = max(
        int(ws.Cells(ws.Rows.Count, col).End(-4162).Row)
        for col in (
            COL_DOMAIN,
            COL_CATEGORY,
            COL_POST_ID,
            COL_OUTPUT_URL,
        )
    )
    last_data_row = min(ROW_END, last_data_row)

    def read_column_values(col):
        if last_data_row < ROW_START:
            return []

        raw = ws.Range(
            ws.Cells(ROW_START, col),
            ws.Cells(last_data_row, col),
        ).Value2

        if isinstance(raw, tuple):
            return [
                item[0] if isinstance(item, tuple) else item
                for item in raw
            ]

        return [raw]

    domain_values = read_column_values(COL_DOMAIN)
    category_values = read_column_values(COL_CATEGORY)
    post_id_values = read_column_values(COL_POST_ID)
    output_url_values = read_column_values(COL_OUTPUT_URL)

    print(
        f"✅ Đã đọc 4 vùng Excel một lần, "
        f"chỉ kiểm tra dòng {ROW_START}–{last_data_row}."
    )

    for row, domain_raw, category_raw, post_id_raw, output_url_raw in zip(
        range(ROW_START, last_data_row + 1),
        domain_values,
        category_values,
        post_id_values,
        output_url_values,
    ):

        domain = normalize_domain(domain_raw)
        category = normalize_category(category_raw)
        post_id = parse_int(post_id_raw)
        current_url = clean_text(output_url_raw)

        # Dòng hoàn toàn trống thì bỏ qua
        if not domain and not category and not post_id and not current_url:
            skipped_empty += 1
            continue

        # Không có ID bài viết thì không chạy
        if not post_id:
            skipped_no_id += 1
            continue

        # Cột URL đã có dữ liệu thì không chạy lại
        if current_url:
            skipped_has_url += 1
            continue

        if not domain:
            ws.Cells(row, COL_OUTPUT_URL).Value = "NO_DOMAIN"
            no_domain += 1
            continue

        if not category:
            ws.Cells(row, COL_OUTPUT_URL).Value = "NO_CATEGORY"
            no_category += 1
            continue

        cat_id = category_map.get((domain, category))

        if not cat_id:
            ws.Cells(
                row,
                COL_OUTPUT_URL
            ).Value = "NO_CATEGORY_ID"

            no_category_id += 1

            print(
                f"⚠️ Dòng {row}: "
                f"không tìm thấy ID danh mục cho "
                f"{domain} | {category}"
            )

            continue

        fake_url = build_fake_url(
            domain,
            cat_id,
            post_id
        )

        rows.append(row)
        fake_urls.append(fake_url)

    # ================== LẤY URL THẬT ==================
    if fake_urls:
        print(
            f"🔄 Đang lấy URL thật cho "
            f"{len(fake_urls)} dòng..."
        )

        final_urls = asyncio.run(
            resolve_all(fake_urls)
        )

        for row, final_url in zip(rows, final_urls):
            ws.Cells(
                row,
                COL_OUTPUT_URL
            ).Value = final_url

            print(
                f"✅ Dòng {row}: {final_url}"
            )
    else:
        print("ℹ️ Không có dòng mới nào cần lấy URL.")

    # ================== LƯU FILE ==================
    try:
        wb.Save()
    except Exception as e:
        print(f"❌ Không thể lưu file Excel: {e}")
        return

    # ================== KẾT QUẢ ==================
    print("")
    print("========== KẾT QUẢ ==========")
    print(f"✅ Đã xử lý mới: {len(fake_urls)} URL")
    print(f"⏭️ Đã bỏ qua vì URL đã có dữ liệu: {skipped_has_url}")
    print(f"⏭️ Đã bỏ qua vì không có ID: {skipped_no_id}")
    print(f"⏭️ Dòng trống: {skipped_empty}")
    print(f"⚠️ Thiếu tên miền: {no_domain}")
    print(f"⚠️ Thiếu danh mục: {no_category}")
    print(f"⚠️ Không tìm thấy ID danh mục: {no_category_id}")
    print("✅ Đã lưu kết quả vào cột URL của sheet dangbai.")


if __name__ == "__main__":
    main()
