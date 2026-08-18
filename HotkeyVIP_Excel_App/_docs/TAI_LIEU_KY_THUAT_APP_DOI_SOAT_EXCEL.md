# Tài liệu kỹ thuật app đối soát Excel

Phiên bản tài liệu: **1.0**  
Phiên bản app đang mô tả: **1.5.1**

Tài liệu này dùng cho việc bảo trì và phát triển app về sau. File hướng dẫn sử dụng nhanh cho người dùng là `HUONG_DAN_APP_DOI_SOAT_EXCEL.md`.

## 1. Mục đích dự án

App đọc một workbook Excel, phân tích ba sheet nghiệp vụ và trình bày kết quả trên giao diện Windows:

- `KE_HOACH`: dữ liệu kế hoạch và URL Page.
- `VIET_BAI`: trạng thái viết bài, file Word và hai ảnh.
- `DANG_BAI`: dữ liệu chuẩn bị đăng hoặc đã đăng.

App dùng để:

- Đếm số lượng theo từng tên miền và tổng toàn bộ.
- Đối chiếu dữ liệu giữa ba sheet bằng Combo 4.
- Phân biệt lỗi dữ liệu với công việc chưa hoàn thành.
- Chỉ ra sheet nguồn, dòng nguồn và sheet đang bị thiếu dữ liệu.
- Tạo danh sách có thể sao chép trực tiếp vào `DANG_BAI` để khôi phục bài đã có URL nhưng bị mất khỏi sheet này.
- Xuất một bản workbook mới có sheet `Tong_all` mà không ghi đè file gốc.

## 2. Những nguyên tắc không được thay đổi tùy tiện

1. Tên sheet và tiêu đề cột được nhận diện không phân biệt chữ hoa/thường và dấu phân cách.
2. Nội dung dùng để tạo Combo 4 được chuẩn hóa khoảng trắng và chữ hoa/thường nhưng vẫn giữ dấu tiếng Việt.
3. Combo 4 là:

   `Tên miền + Tiêu đề SEO + H1 + Từ khóa/Tiêu đề`

4. Trong `DANG_BAI`, cột `Tiêu đề` đóng vai trò Từ khóa.
5. Kiểm tra trùng trong `VIET_BAI` phải dùng Combo 4. Trùng Tiêu đề SEO + H1 + Từ khóa nhưng khác tên miền không phải lỗi.
6. URL hợp lệ chỉ khi bắt đầu bằng `http://` hoặc `https://`.
7. `Chưa chuyển sang DANG_BAI` là tiến độ bình thường, không được tính vào `Lỗi dữ liệu`.
8. Khi `KE_HOACH` vừa trùng Combo 4 vừa thiếu dữ liệu, trạng thái nguồn ưu tiên `Bài viết trùng`.
9. Phân tích không được thay đổi workbook.
10. Xuất kết quả luôn tạo file mới, không được ghi đè file nguồn.
11. Không được viết lại toàn bộ XML của `KE_HOACH`, `VIET_BAI` hoặc `DANG_BAI`. Cơ chế xuất hiện tại dùng Excel COM trên bản sao tạm và kiểm tra lại dữ liệu trước khi công bố file.

## 3. Cấu trúc thư mục

```text
CHAY_APP.cmd
HUONG_DAN_APP_DOI_SOAT_EXCEL.md
TAI_LIEU_KY_THUAT_APP_DOI_SOAT_EXCEL.md
excel_audit_app/
  __init__.py
  main.py
  ui.py
  analysis.py
  excel_io.py
  flow_catalog.py
  flow_host.py
  session.py
  report_export.py
  export_with_excel.ps1
  recover_dang_bai_with_excel.ps1
tests/
  test_excel_audit_app.py
```

Vai trò từng file:

| File | Nhiệm vụ |
|---|---|
| `CHAY_APP.cmd` | File người dùng nhấp đúp để mở app bằng `pyw.exe` hoặc `pythonw.exe`. |
| `main.py` | Điểm khởi động module Python. |
| `ui.py` | Giao diện Tkinter, bảng số liệu, bộ lọc, clipboard và luồng chạy nền. |
| `analysis.py` | Toàn bộ quy tắc nghiệp vụ, Combo 4, phân loại và đối chiếu. |
| `excel_io.py` | Đọc `.xlsx/.xlsm` trực tiếp từ Open XML, không mở Excel. |
| `flow_catalog.py` | Danh mục flow đang dùng, mô tả, cảnh báo và script được gọi. |
| `flow_host.py` | Mở workbook đang chọn bằng Excel ẩn rồi chạy flow trong tiến trình con. |
| `session.py` | Lưu và khôi phục kết quả phân tích gần nhất. |
| `report_export.py` | Chuẩn bị dữ liệu `Tong_all`, tạo bản sao tạm và kiểm tra an toàn sau xuất. |
| `export_with_excel.ps1` | Điều khiển Excel COM để cập nhật bản sao và lưu workbook. |
| `recover_dang_bai_with_excel.ps1` | Nối các dòng khôi phục vào `DANG_BAI` trên bản sao, không sửa file nguồn. |
| `test_excel_audit_app.py` | Kiểm thử số liệu, tình huống xóa dòng, dữ liệu khôi phục và an toàn xuất file. |

## 4. Yêu cầu môi trường

- Windows.
- Python 3 có Tkinter.
- App phân tích được `.xlsx` và `.xlsm`.
- Phân tích không cần Microsoft Excel.
- Chức năng `Xuất kết quả` cần Microsoft Excel vì dùng COM.
- `.xls` kiểu cũ hiện không được hỗ trợ.

## 5. Cách khởi động

Người dùng mở:

```text
CHAY_APP.cmd
```

File CMD thử lần lượt:

1. `pyw.exe -3`
2. `pythonw.exe`
3. Python đi kèm runtime Codex nếu tồn tại

Lệnh chạy module thực tế:

```powershell
python -m excel_audit_app.main
```

## 6. Sheet và cột bắt buộc

Các alias được khai báo tại `ALIASES` trong `analysis.py`. Nếu đổi tên cột Excel, nên bổ sung alias thay vì hard-code vị trí cột.

### 6.1. KE_HOACH

| Trường logic | Tên cột thường gặp |
|---|---|
| Tiêu đề SEO | `Title [SEO]`, `Title SEO`, `Tiêu đề SEO` |
| URL | `URL Page` |
| Tên miền | `Tên Miền`, `Tên miền` |
| Từ khóa | `Main Keyword`, `Từ khóa` |
| H1 | `Article Name [H1]`, `H1` |
| Danh mục | `CATE [POST]`, `Danh mục` |
| Trạng thái nguồn | `Trạng thái nguồn` |

`Danh mục` và `Trạng thái nguồn` là cột tùy chọn. Các cột còn lại là bắt buộc.

### 6.2. VIET_BAI

| Trường logic | Tên cột thường gặp |
|---|---|
| Tiêu đề SEO | `Tiêu đề SEO`, `Title [SEO]`, `Title SEO` |
| H1 | `H1` |
| Tên miền | `Tên Miền`, `Tên miền` |
| Từ khóa | `Từ khóa`, `Main Keyword` |
| Word | `Đường dẫn Word` |
| Ảnh 1 | `Đường dẫn ảnh 1` |
| Ảnh 2 | `Đường dẫn ảnh 2` |
| Hoàn tất | `Trạng thái hoàn tất` |

### 6.3. DANG_BAI

| Trường logic | Tên cột thường gặp |
|---|---|
| Từ khóa | `Tiêu đề` |
| Tiêu đề SEO | `Tiêu đề SEO` |
| Trạng thái đăng | `Trạng thái đăng` |
| Tên miền | `Tên Miền`, `Tên miền` |
| H1 | `H1` |
| URL | `URL đã đăng`, `URL` |
| Word | `Đường dẫn Word` |
| Ảnh 1 | `Đường dẫn ảnh 1` |
| Ảnh 2 | `Đường dẫn ảnh 2` |
| Danh mục | `Danh mục`, `CATE [POST]` |

Danh mục là cột tùy chọn.

## 7. Chuẩn hóa dữ liệu

`excel_io.py` có ba mức chuẩn hóa:

- `normalize_spaces`: đổi nhiều khoảng trắng thành một khoảng trắng và bỏ khoảng trắng hai đầu.
- `normalize_text`: chuẩn hóa khoảng trắng và dùng `casefold()` để so sánh không phân biệt hoa/thường; vẫn giữ dấu tiếng Việt.
- `normalize_header`: bỏ dấu tiếng Việt và ký tự phân cách để nhận diện tên sheet/cột linh hoạt.

Không nên dùng `normalize_header` để tạo Combo 4 vì việc bỏ dấu có thể làm hai nội dung khác nhau trở thành giống nhau.

## 8. Phân loại URL Page trong KE_HOACH

Mỗi dòng thuộc đúng một nhóm:

- `url_valid`: bắt đầu bằng `http://` hoặc `https://`.
- `url_written`: nội dung sau chuẩn hóa là `Đã viết`.
- `url_blank`: ô trống.
- `url_other`: nội dung khác ba nhóm trên; đây là dữ liệu cần kiểm tra.

Kiểm tra tổng:

```text
Tổng KE_HOACH = URL hợp lệ + Đã viết + URL trống + URL sai/khác
```

## 9. Các phép đối chiếu chính

Ký hiệu:

- `K`: tập Combo 4 hợp lệ trong `KE_HOACH`.
- `V`: tập Combo 4 hợp lệ trong `VIET_BAI`.
- `D`: tập Combo 4 hợp lệ trong `DANG_BAI`.
- `K_url`: Combo 4 trong `KE_HOACH` có URL hợp lệ.

### 9.1. KE_HOACH với VIET_BAI

- `KE có - VIET thiếu`: Combo 4 thuộc `K` nhưng không thuộc `V`.
- `VIET có - KE thiếu`: Combo 4 thuộc `V` nhưng không thuộc `K`.

Nếu người dùng xóa hai dòng trong `VIET_BAI`, app không thể biết số dòng VIET cũ nếu không có lịch sử trước khi xóa. App sẽ báo chính xác:

- Tên miền.
- Combo 4.
- Dòng nguồn còn tồn tại trong `KE_HOACH`.
- Sheet đích đang thiếu là `VIET_BAI`.

### 9.2. VIET_BAI với DANG_BAI

Mỗi dòng `VIET_BAI` có Combo 4 hợp lệ thuộc đúng một nhóm:

- `Đã có trong DANG`: Combo 4 thuộc `D`.
- `Cần khôi phục DANG`: chưa thuộc `D` nhưng thuộc `K_url`.
- `Chưa chuyển DANG`: chưa thuộc `D` và không thuộc `K_url`.
- `Thiếu Combo 4`: không thể đối chiếu.

Phép kiểm tra bắt buộc:

```text
Tổng VIET_BAI
= Đã có trong DANG
+ Cần khôi phục DANG
+ Chưa chuyển DANG
+ VIET thiếu Combo 4
```

`difference` phải bằng 0.

Không dùng trực tiếp `Tổng DANG_BAI + VIET chưa chuyển` vì `DANG_BAI` có thể chứa dòng không tồn tại trong `VIET_BAI`, làm phép cộng có vẻ đúng nhưng thực tế sai.

### 9.3. Kiểm tra chiều ngược DANG_BAI với VIET_BAI

- `DANG có - VIET thiếu`: Combo 4 có trong `DANG_BAI` nhưng không có trong `VIET_BAI`; đây là lỗi dữ liệu.
- Dòng `DANG_BAI` thiếu Combo 4 cũng là lỗi dữ liệu.

Phép kiểm tra nội bộ:

```text
Tổng DANG_BAI
= Có trong VIET_BAI
+ DANG có - VIET thiếu
+ DANG thiếu Combo 4
```

## 10. Phân loại tài nguyên của VIET_BAI

Mỗi dòng `VIET_BAI` thuộc đúng một nhóm:

- `completed_with_assets`: trạng thái hoàn tất là `OK` và đủ Word + ảnh 1 + ảnh 2.
- `archived_posted_no_assets`: đã `OK`, thiếu tài nguyên nhưng tìm thấy dòng `ĐÃ ĐĂNG` trong `DANG_BAI`; đây là trường hợp hợp lệ.
- `recovery_no_assets`: đã `OK`, thiếu tài nguyên, có URL trong `KE_HOACH` nhưng không còn trong `DANG_BAI`; cần khôi phục.
- `unexplained_no_assets`: đã `OK`, thiếu tài nguyên và không có bằng chứng đã đăng; đây là lỗi.
- `not_completed`: chưa có trạng thái hoàn tất `OK`.

Phép kiểm tra:

```text
Tổng VIET_BAI
= OK + đủ tài nguyên
+ Đã đăng, đã xóa tài nguyên
+ Cần khôi phục DANG
+ Thiếu tài nguyên bất thường
+ Chưa hoàn tất
```

## 11. Mức độ chi tiết

Mỗi dòng trong `details` có một `level`:

| Level | Hiển thị | Ý nghĩa |
|---|---|---|
| `error` | Lỗi dữ liệu | Cần sửa dữ liệu hoặc kiểm tra chênh lệch. |
| `recovery` | Cần khôi phục | Đã có URL nhưng thiếu trong `DANG_BAI`. |
| `pending` | Chưa chuyển | Tiến độ bình thường, không phải lỗi. |
| `info` | Đã đăng | Thiếu tài nguyên trong VIET nhưng đã xác nhận đăng. |

`issue_count` chỉ đếm `level == error`. Không được dùng tổng số dòng `details` làm số lỗi.

## 12. Trạng thái tổng thể

Thứ tự ưu tiên:

1. Có lỗi dữ liệu: `CẦN KIỂM TRA`, màu đỏ.
2. Không có lỗi nhưng có bài cần khôi phục: `CẦN KHÔI PHỤC`, màu cam.
3. Không có lỗi/khôi phục nhưng còn bài chưa chuyển: `DỮ LIỆU KHỚP`, màu xanh dương.
4. Không còn lỗi, khôi phục hoặc tiến độ: `ỔN`, màu xanh lá.

## 13. Danh sách khôi phục DANG_BAI

Điều kiện tạo dòng khôi phục:

```text
KE_HOACH có Combo 4 hợp lệ
AND URL Page hợp lệ
AND Combo 4 không tồn tại trong DANG_BAI
```

App dựng dữ liệu theo đúng số lượng và thứ tự cột hiện có trong sheet `DANG_BAI`.

Ánh xạ chính:

| Cột DANG_BAI | Nguồn |
|---|---|
| `Tiêu đề` | Từ khóa từ `KE_HOACH` |
| `Tiêu đề SEO` | Tiêu đề SEO từ `KE_HOACH` |
| `Trạng thái đăng` | Gán `ĐÃ ĐĂNG` |
| `Tên miền` | `KE_HOACH` |
| `Danh mục` | `CATE [POST]` từ `KE_HOACH` nếu có |
| `H1` | `KE_HOACH` |
| `URL đã đăng` | `URL Page` từ `KE_HOACH` |
| Word, ảnh 1, ảnh 2 | Lấy từ dòng Combo 4 tương ứng trong `VIET_BAI` nếu có |

Các cột không lấy được dữ liệu để trống. Nút sao chép không kèm dòng tiêu đề để người dùng dán vào dòng trống đầu tiên của `DANG_BAI`.

Nút `Tạo file mới + khôi phục` chỉ được bật khi có dòng khôi phục, file nguồn chưa thay đổi và `error_count = 0`. Luồng xử lý:

1. Người dùng chọn tên file mới, không được trùng file nguồn.
2. App sao chép nguyên workbook sang file tạm.
3. `recover_dang_bai_with_excel.ps1` mở bản tạm bằng Excel COM ở chế độ ẩn và nối các dòng vào cuối vùng dữ liệu `DANG_BAI`.
4. App xác nhận dữ liệu cũ của `DANG_BAI` không đổi, số dòng tăng đúng, `KE_HOACH` và `VIET_BAI` không đổi, VBA vẫn còn và số khôi phục sau phân tích bằng 0.
5. Chỉ khi tất cả kiểm tra đạt, bản tạm mới được đổi thành tên file người dùng chọn.
6. UI chuyển sang bản mới và hiển thị kết quả phân tích mới. File nguồn cũ không bị sửa.

## 14. Giao diện

Các tab:

- `Tổng quan`: trạng thái tổng thể, thẻ số liệu và đối soát theo tên miền.
- `Kế hoạch`: số liệu `KE_HOACH` theo tên miền.
- `Viết bài`: số liệu tài nguyên và hoàn tất.
- `Đăng bài`: số liệu `DANG_BAI` và chiều đối chiếu ngược.
- `Chi tiết đối soát`: bộ lọc level, loại chi tiết, tìm kiếm và dòng nguồn/đích.
- `Khôi phục DANG_BAI`: danh sách xem trước, nút sao chép và nút tạo bản mới để khôi phục tự động.

Nhấp vào thẻ số hoặc nhấp đúp ô chênh lệch trong bảng tổng quan sẽ mở bộ lọc chi tiết tương ứng.

## 15. Luồng xử lý khi bấm Phân tích

```text
Chọn file
  -> kiểm tra định dạng ZIP/XLSX/XLSM
  -> tìm ba sheet không phân biệt hoa/thường
  -> phát hiện hàng tiêu đề và cột bằng alias
  -> đọc dữ liệu Open XML
  -> tạo Combo 4 và số liệu theo tên miền
  -> tạo details/recovery/overall
  -> tính SHA-256 và dấu vân tay file
  -> lưu session JSON
  -> render giao diện
```

Phân tích chạy trong thread nền để cửa sổ không bị đứng khi đọc file. Việc đưa hàng nghìn dòng vào `Treeview` vẫn đang thực hiện trên UI thread.

## 16. Session và cache

Mặc định lưu tại:

```text
%LOCALAPPDATA%\ExcelAuditApp\session.json
```

Có thể đổi thư mục bằng biến môi trường:

```text
EXCEL_AUDIT_APPDATA
```

Session chứa kết quả phân tích, không chứa bản sao workbook.

- Nếu dấu vân tay file chưa đổi, có thể xem lại và xuất.
- Nếu file nguồn đã đổi, app tự phân tích lại ở thread nền khi khởi động.
- Nếu `APP_VERSION` khác phiên bản trong session, app không render dữ liệu cũ nhưng giữ đường dẫn và tự phân tích lại nếu file còn tồn tại.
- `Xóa phiên` chỉ xóa `session.json`, không xóa file Excel.

Khi tăng phiên bản có thay đổi cấu trúc dữ liệu, phải cập nhật đồng thời:

- `APP_VERSION` trong `analysis.py`.
- `__version__` trong `__init__.py`.
- Phiên bản trong hai file Markdown.

## 17. Luồng xuất file an toàn

Khi người dùng bấm `Xuất kết quả`:

1. Kiểm tra file nguồn vẫn có cùng kích thước, thời gian sửa và SHA-256 như lúc phân tích.
2. Tạo bản sao tạm trong thư mục đích.
3. Tạo payload JSON chứa `Tong_all` và cập nhật trạng thái nguồn.
4. Chạy `export_with_excel.ps1` để mở bản sao bằng Excel COM ở chế độ ẩn.
5. Chỉ cập nhật:

   - Cột `Trạng thái nguồn` trong `KE_HOACH`.
   - Sheet `Tong_all`.

6. Đóng và lưu bản sao.
7. Kiểm tra workbook là ZIP hợp lệ.
8. So sánh snapshot ngữ nghĩa của `KE_HOACH`, `VIET_BAI`, `DANG_BAI` trước và sau xuất. Riêng cột trạng thái nguồn được phép khác.
9. Kiểm tra `vbaProject.bin` vẫn tồn tại và không rỗng nếu file nguồn có macro.
10. Chỉ sau khi tất cả kiểm tra đạt mới đổi tên bản tạm thành file người dùng chọn.

Nếu bất kỳ bước nào lỗi, file tạm bị hủy và file nguồn không thay đổi.

## 18. Sheet Tong_all

`Tong_all` hiện chứa:

1. Thông tin file và trạng thái tổng thể.
2. Phân tích `KE_HOACH`.
3. Phân tích `VIET_BAI`.
4. Phân tích `DANG_BAI`.
5. Đối soát ba sheet theo tên miền.
6. Danh sách chi tiết đối soát.
7. Dữ liệu khôi phục theo đúng cấu trúc `DANG_BAI`.

Khi sửa số cột hoặc thứ tự section trong `report_export.py`, phải kiểm tra lại:

- `column_count` và `last_column`.
- `column_widths` có đúng số phần tử.
- `center_ranges`.
- `filter_ranges`.
- Các địa chỉ mong đợi trong unit test.

## 19. Kiểm thử

Chạy toàn bộ:

```powershell
python -m unittest discover -s tests -v
```

Kiểm thử xuất file cần Microsoft Excel. Các kiểm thử chính hiện có:

- Kiểm tra tổng số liệu mẫu.
- Kiểm tra phân loại lỗi, khôi phục, tiến độ và đã đăng.
- Kiểm tra mỗi dòng khôi phục có đúng 18 cột và URL hợp lệ.
- Kiểm tra phép cộng phân loại `VIET_BAI` bằng tổng dòng.
- Giả lập xóa hai dòng `VIET_BAI` và xác nhận app liệt kê hai dòng nguồn `KE_HOACH`.
- Kiểm tra `Tong_all`, VBA và ba sheet nguồn sau khi xuất.

Sau khi sửa logic, tối thiểu phải kiểm tra:

```text
KE tổng = tổng bốn nhóm URL
VIET tổng = tổng năm nhóm tài nguyên
VIET tổng = in_dang + recovery + pending + combo4_missing
DANG tổng = in_viet + dang_missing_viet + combo4_missing
difference = 0 khi dữ liệu hợp lệ
```

## 20. Điểm hiệu năng cần lưu ý

Hiện tại `ui.py` gọi `_render_issues()` ngay sau phân tích hoặc khôi phục session. Hàm này đưa toàn bộ `details` vào `ttk.Treeview`.

Với file gần nhất, tab chi tiết có thể chứa hơn 2.600 dòng, mỗi dòng có Tiêu đề SEO, H1, từ khóa và mô tả dài. Tkinter phải vẽ lại bảng lớn khi di chuyển hoặc thay đổi kích thước cửa sổ, gây cảm giác delay dù CPU ở trạng thái chờ bằng 0.

Hướng tối ưu ưu tiên cho phiên bản sau:

1. Không render `pending` cho đến khi người dùng bấm xem.
2. Phân trang 100-200 dòng.
3. Chỉ render lỗi và khôi phục ở lần đầu.
4. Hiện cột nội dung dài trong panel chi tiết khi chọn dòng, thay vì đặt toàn bộ vào Treeview.
5. Khi lọc, debounce thao tác và không xóa/chèn lại toàn bộ bảng nếu không cần.

Không nên tối ưu bằng cách bỏ dữ liệu khỏi kết quả phân tích; chỉ cần lazy-render ở lớp UI.

## 21. Hạn chế hiện tại

- Không biết số dòng cũ của một dòng đã bị xóa khỏi sheet nếu không có snapshot lịch sử. App chỉ ra dòng nguồn ở sheet còn tồn tại.
- Không hỗ trợ `.xls`.
- Phân tích Open XML không tự tính lại công thức Excel; nó đọc giá trị đã được lưu gần nhất trong workbook.
- Xuất file phụ thuộc Microsoft Excel trên Windows.
- `Treeview` chưa phân trang nên có thể lag với vài nghìn dòng chi tiết.
- Session chỉ giữ lần phân tích gần nhất, chưa có lịch sử nhiều phiên.

## 22. Quy trình sửa code an toàn

### Thêm alias cột

1. Sửa `ALIASES` trong `analysis.py`.
2. Xác định cột bắt buộc hay tùy chọn trong `OPTIONAL_COLUMNS`.
3. Thử file có chữ hoa/thường và dấu khác nhau.

### Thêm một chỉ số

1. Thêm khóa vào factory metrics tương ứng trong `analysis.py`.
2. Cập nhật logic tăng số đếm.
3. Cập nhật `_summary_rows` nếu tổng cần cách tính đặc biệt.
4. Thêm cột trong `ui.py` và đúng thứ tự key tại `_render_result`.
5. Thêm cột trong `report_export.py`.
6. Bổ sung assertion trong test.

### Thêm một loại chi tiết

1. Dùng `_issue()` với `level` chính xác.
2. Điền `sheet`, `row`, `target_sheet`, `target_row` nếu biết.
3. Không dùng `error` cho tiến độ bình thường.
4. Nếu muốn nhấp từ bảng tổng quan, cập nhật `category_map` trong `_open_reconciliation_detail()`.

### Sửa dữ liệu khôi phục

1. Sửa `build_recovery_values()` trong `analysis.py`.
2. Luôn dùng vị trí cột phát hiện từ header `DANG_BAI`, không hard-code A/B/C.
3. Trường không có nguồn phải để trống.
4. Không thêm dòng tiêu đề vào nội dung clipboard.
5. Chạy test độ dài hàng khôi phục và kiểm tra URL/trạng thái đăng.

### Sửa xuất Excel

1. Chỉ thao tác trên bản sao tạm.
2. Không tắt `_verify_export()`.
3. Không nới lỏng kiểm tra VBA.
4. Không cho phép đích trùng file nguồn.
5. Kiểm tra lại workbook mẫu `.xlsm` sau thay đổi.

## 23. Checklist phát hành phiên bản mới

- [ ] Cập nhật version trong `analysis.py` và `__init__.py`.
- [ ] Cập nhật tài liệu phiên bản.
- [ ] `python -m py_compile` không có lỗi.
- [ ] PowerShell parser không báo lỗi `export_with_excel.ps1`.
- [ ] Unit test logic đạt.
- [ ] Test xóa hai dòng đạt.
- [ ] Test xuất file đạt.
- [ ] `Tong_all` có đủ section.
- [ ] `KE_HOACH`, `VIET_BAI`, `DANG_BAI` không thay đổi ngoài phạm vi cho phép.
- [ ] VBA còn nguyên.
- [ ] Mở app và kiểm tra Tổng quan, Chi tiết và Khôi phục.
- [ ] Đóng/mở lại app để kiểm tra session và version cache.

## 24. Số liệu mẫu dùng để kiểm tra nhanh

Số liệu phụ thuộc workbook được chọn. Với file `hotkeyvip_test - Copy (2).xlsm` đã kiểm tra ở phiên bản 1.2.0 trước khi khôi phục:

```text
Tổng KE_HOACH:              6.748
Tổng VIET_BAI:              6.748
Đã có trong DANG_BAI:       4.342
Chưa chuyển DANG_BAI:       2.405
Cần khôi phục DANG_BAI:         1
VIET thiếu Combo 4:              0
Chênh lệch:                       0
Lỗi dữ liệu:                      0
Combo 4 trùng trong VIET_BAI:    0
Đã đăng, đã xóa tài nguyên:    279
```

Phép kiểm tra:

```text
6.748 = 4.342 + 1 + 2.405 + 0
```

Nếu workbook mẫu thay đổi có chủ ý, cần cập nhật cả số liệu mẫu và unit test liên quan.

## 25. Tích hợp các flow nghiệp vụ

Tab `Công việc` là điểm chạy chung cho các flow. UI không import trực tiếp các file Selenium lớn; mỗi flow chạy trong tiến trình riêng để lỗi của flow không làm sập app.

```text
File đang chọn trên app
  -> flow_host.py mở workbook bằng Excel COM ẩn
  -> đặt HOTKEYVIP_SELECTED_EXCEL
  -> chạy script flow trong tiến trình con
  -> script dùng đúng workbook đang chọn
  -> lưu và đóng workbook
  -> app phân tích lại bằng Open XML
  -> cập nhật dashboard và session
```

Quy tắc:

- Mỗi lần chỉ chạy một flow.
- Người dùng không cần mở Excel, nhưng máy phải cài Microsoft Excel.
- `HOTKEYVIP_SELECTED_EXCEL` ghi đè đường dẫn Excel mặc định của các flow có cấu hình cứng.
- `HOTKEYVIP_APP_RUN=1` cho flow biết nó đang được app điều phối; hiện dùng để bỏ bước chờ Enter và không dọn nhầm Excel ẩn của host.
- Các flow tác động CMS, ChatGPT/Gemini hoặc file công ty phải có xác nhận rõ trước khi chạy.
- App vẫn cho xem các tab số liệu trong lúc flow chạy nhưng khóa đổi file, phân tích, xuất và chạy flow thứ hai.
- Kết thúc flow phải phân tích lại workbook, kể cả flow trả mã lỗi, vì có thể đã có dữ liệu được lưu từng phần.

Flow app đang gọi:

| Thứ tự | Script | Vai trò |
|---|---|---|
| 1 | `app_flows/01_nhap_ke_hoach.py` | Nhập Article nguồn vào `KE_HOACH`. |
| 2 | `app_flows/02_chuan_bi_viet_bai.py` | Chuẩn bị và đối chiếu `VIET_BAI`. |
| 3 | `app_flows/03_viet_bai_tao_anh.py` | Viết bài, tạo Word và ảnh. |
| 4 | `app_flows/04_chuan_bi_dang_bai.py` | Chuẩn bị `DANG_BAI`. |
| 5 | `app_flows/05_dang_bai_cms.py` | Đăng bài CMS đa luồng. |
| 6 | `app_flows/06_lay_url_cms.py` | Lấy URL thật từ ID CMS. |
| 7 | `app_flows/07_bai_viet_lien_quan.py` | Cập nhật bài viết liên quan. |
| 8 | `app_flows/08_dong_bo_url.py` | Đồng bộ URL về file công ty. |

Danh sách file cũ/độc lập và quyết định giữ nguyên nằm trong `NHAT_KY_TICH_HOP_FLOW.md`.
