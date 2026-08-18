import asyncio
import aiohttp
import ssl
import re
import os
import sys
import unicodedata
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
GENERATED_URL_FILL_COLOR = 10284031  # RGB(255, 235, 156): vàng nhạt
GENERATED_URL_NOTE = "URL tự tạo từ H1 do HEAD không chuyển hướng; cần kiểm tra lại."
VERSION = "06_lay_url_cms (engine V1.7)"

TEMP_URL_PATTERN = re.compile(
    r"^https?://[^/]+/linkrutgon-\d+-\d+\.html(?:[?#].*)?$",
    re.IGNORECASE,
)


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


def slugify_h1(value):
    text = clean_text(value).casefold().replace("đ", "d")
    text = unicodedata.normalize("NFD", text)
    text = "".join(char for char in text if unicodedata.category(char) != "Mn")
    return re.sub(r"-+", "-", re.sub(r"[^a-z0-9]+", "-", text)).strip("-")


def build_real_url(domain, h1, cat_id, post_id):
    slug = slugify_h1(h1)
    if not slug:
        return ""
    return f"https://{normalize_domain(domain)}/{slug}-{cat_id}-{post_id}.html"


def mark_generated_url_cell(cell):
    """Chỉ đánh dấu trực quan; không thay đổi giá trị dùng bởi các flow khác."""
    cell.Interior.Color = GENERATED_URL_FILL_COLOR
    try:
        comment = cell.Comment
        if comment is None:
            cell.AddComment(GENERATED_URL_NOTE)
            return
        current_note = clean_text(comment.Text())
        if GENERATED_URL_NOTE not in current_note:
            comment.Text(
                f"{current_note}\n{GENERATED_URL_NOTE}".strip()
            )
    except Exception as exc:
        print(f"⚠️ Không thêm được Note vào ô URL: {exc}")


def is_temporary_url(value):
    return bool(TEMP_URL_PATTERN.match(clean_text(value)))


async def fetch_final(session, url, fallback_url):
    try:
        async with session.head(
            url,
            allow_redirects=True
        ) as resp:
            final_url = str(resp.url)
            if final_url != url:
                return final_url, False
    except Exception:
        pass

    # HEAD không chuyển được thì tạo URL theo quy luật slug H1-cat_id-post_id.
    # Nếu thiếu H1 thì giữ URL tạm để lần chạy sau xử lý lại.
    return fallback_url or url, bool(fallback_url)


async def resolve_all(urls, fallback_urls):
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
            fetch_final(session, url, fallback_url)
            for url, fallback_url in zip(urls, fallback_urls)
        ]

        return await asyncio.gather(*tasks)


def main():
    print(f"PHIÊN BẢN: {VERSION}")
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
        COL_H1 = find_col(ws, "H1")
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
            f"H1={COL_H1}, "
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
    fallback_urls = []

    skipped_no_id = 0
    skipped_has_url = 0
    retried_temporary_url = 0
    skipped_empty = 0
    no_domain = 0
    no_category = 0
    no_category_id = 0

    # Chỉ đọc Excel theo 5 vùng lớn, không gọi COM riêng cho từng ô/dòng.
    # Dòng cuối được xác định từ năm cột liên quan và vẫn giới hạn bởi ROW_END.
    last_data_row = max(
        int(ws.Cells(ws.Rows.Count, col).End(-4162).Row)
        for col in (
            COL_DOMAIN,
            COL_CATEGORY,
            COL_H1,
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
    h1_values = read_column_values(COL_H1)
    post_id_values = read_column_values(COL_POST_ID)
    output_url_values = read_column_values(COL_OUTPUT_URL)

    print(
        f"✅ Đã đọc 5 vùng Excel một lần, "
        f"chỉ kiểm tra dòng {ROW_START}–{last_data_row}."
    )

    for row, domain_raw, category_raw, h1_raw, post_id_raw, output_url_raw in zip(
        range(ROW_START, last_data_row + 1),
        domain_values,
        category_values,
        h1_values,
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

        # URL thật đã có thì bỏ qua. URL linkrutgon là URL tạm nên phải thử lại.
        if current_url and not is_temporary_url(current_url):
            skipped_has_url += 1
            continue
        if is_temporary_url(current_url):
            retried_temporary_url += 1

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
        fallback_urls.append(build_real_url(domain, h1_raw, cat_id, post_id))

    # ================== LẤY URL THẬT ==================
    if fake_urls:
        print(
            f"🔄 Đang lấy URL thật cho "
            f"{len(fake_urls)} dòng..."
        )

        final_urls = asyncio.run(
            resolve_all(fake_urls, fallback_urls)
        )

        resolved_urls = 0
        unresolved_urls = 0
        generated_from_h1 = 0
        for row, result in zip(rows, final_urls):
            final_url, used_h1_fallback = result
            ws.Cells(
                row,
                COL_OUTPUT_URL
            ).Value = final_url

            if is_temporary_url(final_url):
                unresolved_urls += 1
                print(f"⚠️ Dòng {row}: vẫn là URL tạm: {final_url}")
            else:
                resolved_urls += 1
                if used_h1_fallback:
                    generated_from_h1 += 1
                    mark_generated_url_cell(ws.Cells(row, COL_OUTPUT_URL))
                    print(f"🧩 Dòng {row}: URL tạo từ H1: {final_url}")
                else:
                    print(f"✅ Dòng {row}: {final_url}")
    else:
        resolved_urls = 0
        unresolved_urls = 0
        generated_from_h1 = 0
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
    print(f"🔎 Tổng URL đã thử: {len(fake_urls)}")
    print(f"✅ Lấy được URL thật: {resolved_urls}")
    print(f"🧩 URL được tự tạo từ H1, cần kiểm tra lại: {generated_from_h1}")
    print(f"⚠️ Vẫn còn URL tạm: {unresolved_urls}")
    print(f"🔁 URL tạm được xử lý lại: {retried_temporary_url}")
    print(f"⏭️ Đã bỏ qua vì URL đã có dữ liệu: {skipped_has_url}")
    print(f"⏭️ Đã bỏ qua vì không có ID: {skipped_no_id}")
    print(f"⏭️ Dòng trống: {skipped_empty}")
    print(f"⚠️ Thiếu tên miền: {no_domain}")
    print(f"⚠️ Thiếu danh mục: {no_category}")
    print(f"⚠️ Không tìm thấy ID danh mục: {no_category_id}")
    print("✅ Đã lưu kết quả vào cột URL của sheet dangbai.")


if __name__ == "__main__":
    main()
