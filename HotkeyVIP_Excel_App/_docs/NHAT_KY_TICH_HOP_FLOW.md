# Nhật ký tích hợp flow vào app đối soát Excel

Ngày cập nhật: **2026-08-16**  
Phiên bản app: **1.5.1**

- Phiên bản 1.5.1 nâng Flow 07 lên engine V2.4: tự phục hồi Edge/bài lỗi,
  ghi Excel qua hàng đợi riêng, lưu theo lô, đếm worker theo domain và thu hồi
  bài khi worker tắt bất thường. Thiết lập Git cục bộ và `CHANGELOG.md` để tạo
  mốc phục hồi rõ ràng cho các lần sửa sau.

- Phiên bản 1.5.0 thêm tab `Theo dõi đăng bài`: xem lỗi, mở Word/URL ChatGPT, nhập ID CMS,
  đăng lại đúng các dòng `LỖI KIỂM TRA` bằng 1 worker và xem các bài đã đăng hôm nay.
- Thêm nút `Mở Excel`; nút tự khóa trong lúc phân tích hoặc chạy flow.
- Sau khi flow hoàn tất, app tự phân tích lại để làm mới task đăng bài.

## 1. Mục tiêu

- Chọn workbook một lần trên app.
- Xem số liệu đối soát và chạy công việc trong cùng một giao diện.
- Không yêu cầu người dùng mở Excel hoặc chạy từng file Python.
- Giữ các flow lớn ở tiến trình riêng để app không bị sập theo flow.
- Sau mỗi flow, tự phân tích lại workbook và cập nhật dashboard.

Lưu ý: Microsoft Excel vẫn chạy ẩn ở nền vì các flow hiện dùng COM/xlwings và cần giữ macro, công thức, bảo vệ sheet cùng hành vi giống thao tác Excel thật.

## 2. File mới được tạo

| File | Nội dung |
|---|---|
| `excel_audit_app/flow_catalog.py` | Danh mục 8 flow đang dùng, thứ tự, mô tả, cảnh báo và tham số chạy. |
| `excel_audit_app/flow_host.py` | Mở file Excel đang chọn bằng Excel ẩn, đặt biến môi trường và chạy flow con. |
| `excel_audit_app/recover_dang_bai_with_excel.ps1` | Ghi các dòng khôi phục vào `DANG_BAI` trên một bản sao mới. |
| `NHAT_KY_TICH_HOP_FLOW.md` | Nhật ký file sửa, file giữ nguyên và lý do. |

## 3. File đã sửa

| File | Thay đổi |
|---|---|
| `excel_audit_app/ui.py` | Thêm nút tạo file mới để khôi phục `DANG_BAI`; thêm tab `Công việc`, nút chạy 8 flow, xác nhận tác động, khung log và tự phân tích lại. |
| `excel_audit_app/report_export.py` | Thêm đường dẫn gợi ý, tạo bản tạm, ghi và kiểm tra an toàn cho khôi phục `DANG_BAI`. |
| `excel_audit_app/analysis.py` | Tăng phiên bản app lên `1.2.0`. Logic Combo 4 và số liệu không thay đổi. |
| `excel_audit_app/__init__.py` | Đồng bộ phiên bản `1.2.0`. |
| `app_flows/05_dang_bai_cms.py` | `EXCEL_PATH` ưu tiên `HOTKEYVIP_SELECTED_EXCEL`, mặc định cũ vẫn giữ. |
| `app_flows/07_bai_viet_lien_quan.py` | `EXCEL_PATH` ưu tiên file do app chọn; chạy từ app mặc định truyền 3 worker. |
| `app_flows/06_lay_url_cms.py` | Đường dẫn workbook ưu tiên file do app chọn. |
| `app_flows/08_dong_bo_url.py` | File nguồn ưu tiên file app chọn; khi chạy từ app không dọn Excel ẩn của host và không chờ Enter khi kết thúc. |
| `tests/test_excel_audit_app.py` | Thêm kiểm thử khôi phục tạo file mới, tăng đúng số dòng và không sửa file nguồn. |
| `HUONG_DAN_APP_DOI_SOAT_EXCEL.md` | Thêm hướng dẫn khôi phục tự động và tab Công việc. |
| `TAI_LIEU_KY_THUAT_APP_DOI_SOAT_EXCEL.md` | Thêm kiến trúc flow host, quy tắc chạy và danh mục flow. |

## 4. Flow được app sử dụng

### 1. Nhập KE_HOACH

File: `app_flows/01_nhap_ke_hoach.py`

- Đọc sheet `Article` từ file nguồn.
- Chống trùng và cập nhật `KE_HOACH`.
- Có giao diện chọn chế độ/file/thư mục riêng của flow.
- Không sửa code nội bộ; `flow_host.py` bảo đảm workbook app chọn là workbook đích đang hoạt động.

### 2. Chuẩn bị VIET_BAI

File: `app_flows/02_chuan_bi_viet_bai.py`

- Đối chiếu `KE_HOACH` với `VIET_BAI`.
- Thêm bài còn thiếu theo điều kiện của flow.
- Cập nhật trạng thái kiểm tra và giữ các cột nội dung khác.
- Không sửa code nội bộ; dùng workbook đang hoạt động do host mở.

### 3. Viết bài và tạo ảnh

File: `app_flows/03_viet_bai_tao_anh.py`

- Bản mới nhất trong nhóm viết bài.
- Điều phối worker ChatGPT, Word và Gemini.
- Ghi tiến độ, Word, ảnh và lỗi về `VIET_BAI`.
- Không sửa code nội bộ; xlwings nhận workbook ẩn đang hoạt động.

### 4. Chuẩn bị DANG_BAI

File: `app_flows/04_chuan_bi_dang_bai.py`

- Đối chiếu Combo 4 giữa ba sheet.
- Đưa bài đủ điều kiện sang `DANG_BAI` và tránh nhân dòng.
- Không sửa code nội bộ; dùng workbook đang hoạt động do host mở.

### 5. Đăng bài CMS

File: `app_flows/05_dang_bai_cms.py`

- Bản mới nhất trong nhóm đăng bài.
- Chạy Selenium/Word đa luồng và ghi URL, ID, trạng thái về `DANG_BAI`.
- Đã sửa để đường dẫn workbook lấy từ app.
- Đây là flow có tác động website thật nên luôn yêu cầu xác nhận.

### 6. Lấy URL từ ID CMS

File: `app_flows/06_lay_url_cms.py`

- Tạo URL tạm từ domain, category ID và post ID.
- Gửi yêu cầu mạng để lấy URL cuối sau chuyển hướng.
- Ghi URL về `DANG_BAI` của file app chọn.

### 7. Bài viết liên quan

File: `app_flows/07_bai_viet_lien_quan.py`

- Bản mới nhất trong nhóm bài viết liên quan.
- App gọi mặc định với 3 worker để không chờ nhập console.
- Chỉnh sửa CMS và ghi trạng thái về `DANG_BAI`.

### 8. Đồng bộ URL về file công ty

File: `app_flows/08_dong_bo_url.py`

- Lấy nguồn từ `DANG_BAI` của file app chọn.
- Backup rồi cập nhật các file domain bên ngoài.
- Không dọn tiến trình Excel ẩn khi chạy dưới app để tránh đóng nhầm host.

## 5. File được giữ nguyên và không gọi trực tiếp từ app

| File | Quyết định |
|---|---|
| `03_V1.12_vietbai_3cap_baohiem_anh_khong_ghi_de.py` | Giữ làm bản viết bài cũ; app dùng `V2.22`. |
| `V2.21_CODEX_skip_loi_chay_lai.py` | Giữ làm bản trước `V2.22`; không gọi để tránh hai flow cùng nhiệm vụ. |
| `05_V1.8_VIP_tudongdangbai_pipeline_kiemtra_noidung.py` | Giữ bản đăng bài cũ; app dùng `V2.10`. |
| `V1.0_bai_viet_lien_quan.py` | Giữ bản cũ; app dùng `V2.3`. |
| `VIPPPPPP_baivietxoa_xoa_nhu_lam_tay_COM.V2.2.py` | Giữ độc lập vì xóa hàng trên nhiều file domain từ file điều khiển, không dùng workbook đang chọn. Không đưa vào app để tránh nhầm phạm vi xóa. |
| `mo_profile_worker.py` | Giữ nguyên; đây là tiện ích profile, không phải flow Excel hoàn chỉnh. |

Không file nào trong danh sách trên bị xóa.

## 6. Cơ chế an toàn

- Chỉ chạy flow khi kết quả phân tích thuộc đúng file hiện tại.
- Mỗi lần chỉ chạy một flow.
- Flow tác động hệ thống bên ngoài có cảnh báo riêng.
- Excel chạy ẩn và workbook được đóng/lưu ở cuối host.
- App tự phân tích lại kể cả flow trả mã lỗi, vì flow có thể đã lưu một phần.
- Khôi phục `DANG_BAI` luôn tạo file mới, không sửa file gốc.
- Khôi phục chỉ được bật khi không có lỗi dữ liệu màu đỏ.
- Bản khôi phục được kiểm tra số dòng, snapshot ba sheet, ZIP và VBA trước khi công bố.

## 7. Kiểm thử đã thực hiện

- Biên dịch cú pháp các module app và các flow đã sửa.
- Chạy khôi phục trên workbook mẫu: số khôi phục về 0, `DANG_BAI` tăng đúng số dòng, file nguồn giữ nguyên.
- Chạy `Chuẩn bị VIET_BAI` qua `flow_host.py` trên bản sao workbook mẫu.
- Chạy `Chuẩn bị DANG_BAI` qua `flow_host.py` trên cùng bản sao.
- Phân tích lại bản sao: workbook còn hợp lệ, ba sheet đọc được và tổng KE/VIET vẫn khớp.
- Xác nhận trên đúng `hotkeyvip_test - Copy (2).xlsm`: xuất `Tong_all` thành công trên bản sao; file nguồn không đổi.
- Xác nhận khôi phục trên đúng `hotkeyvip_test - Copy (2).xlsm`: `Khôi phục` giảm từ 1 xuống 0 và tổng `DANG_BAI` tăng từ 4.342 lên 4.343.

Workbook `_tests_data/hotkeyvip_test_chogemini.xlsm` có dự án VBA bị chính Microsoft Excel báo hỏng khi thử thêm sheet mới. App dừng xuất và hủy file tạm đúng thiết kế. Kiểm thử thêm sheet trên workbook mẫu này được đánh dấu bỏ qua; kiểm thử xuất được xác nhận riêng trên file người dùng đang sử dụng nêu trên.

Không chạy thử thật các flow ChatGPT, Gemini, đăng CMS, bài viết liên quan hoặc đồng bộ file công ty vì chúng có tác động ra tài khoản/website/file bên ngoài.

## 8. Sửa lỗi Flow 2 không nhận workbook ẩn

- Ngày 09/08/2026: sửa `Chuẩn bị VIET_BAI` để nhận trực tiếp workbook do app mở, thay vì tìm `ActiveWorkbook` từ một tiến trình Python khác.
- `flow_host.py` chạy riêng Flow 2 trong cùng tiến trình Excel COM; các flow khác giữ nguyên cơ chế cũ.
- Kiểm thử trên bản sao `hotkeyvip_test.xlsm`: Flow 2 hoàn thành mã 0, ghi 2.405 dòng `[TRỐNG]`, 3.462 dòng `Đã viết`, 881 dòng `Đã đăng`.
- File Excel gốc không được dùng để kiểm thử và không bị thay đổi.

## 9. Ưu tiên hàng chờ viết bài và bảo hiểm đúng dòng

- Phiên bản 1.2.2 thêm hộp thiết lập trước khi chạy Flow 3.
- Người dùng có thể ưu tiên chạy lại bài lỗi, chọn một tên miền và số bài ưu tiên, sau đó tiếp tục luồng bình thường từ sau mốc `OK OK`.
- Flow không di chuyển hay sắp xếp dòng Excel; nó chỉ xếp lại số dòng trong hàng chờ RAM theo thứ tự `lỗi -> tên miền ưu tiên -> bình thường` và tự loại trùng.
- Nhóm lỗi và tên miền ưu tiên có thể lấy từ toàn bộ `VIET_BAI`; phần chạy bình thường vẫn bắt đầu sau mốc `OK OK` cuối cùng.
- Trước khi xử lý mỗi dòng, Excel Writer xác nhận lại `Tên miền + Từ khóa` tại dòng thật vẫn giống snapshot RAM. Nếu không khớp, toàn bộ flow dừng an toàn và không ghi vào dòng bị thay đổi.
- Nếu app không truyền kế hoạch, V2.22 giữ nguyên cơ chế cũ để vẫn có thể chạy độc lập.

## 10. Preview hàng chờ viết bài không cần mở Excel

- Phiên bản 1.2.3 thêm `excel_audit_app/write_plan.py` để đọc trực tiếp sheet `VIET_BAI` bằng Open XML, không khởi động Microsoft Excel.
- Danh sách tên miền hiển thị số bài chưa hoàn tất và số bài lỗi của từng tên miền.
- Hộp thiết lập hiển thị dòng chứa `OK OK` cuối cùng, dòng bắt đầu luồng thường, tổng hàng chờ và các bài chạy đầu tiên.
- Nếu có bài lỗi, preview chỉ rõ bài nào sẽ chạy ngay sau khi xử lý hết nhóm lỗi.

## 11. Batch đăng bài theo tên miền và một danh mục

- Phiên bản 1.2.4 thêm `excel_audit_app/publish_plan.py` để đọc và lập batch từ `DANG_BAI` mà không mở Excel.
- Mỗi tên miền chọn đúng một danh mục có nhiều bài hợp lệ nhất; nếu nhiều danh mục bằng nhau thì chọn danh mục xuất hiện trước trong sheet.
- Mỗi tên miền lấy tối đa số bài người dùng yêu cầu. Nếu chỉ có 5/7 bài thì đăng 5 và không ghép danh mục khác.
- UX hiển thị bảng xem trước domain, danh mục, số có thể đăng và số sẽ đăng; người dùng xác nhận trước khi chạy CMS thật.
- Khi chạy từ app, V2.10 dùng batch đã chốt và không hỏi lại tổng số bài. Chạy độc lập vẫn giữ hộp hỏi số lượng cũ.
- Trước khi nạp batch và ngay trước khi giao Worker, V2.10 xác nhận lại `Tên miền + Danh mục + Tiêu đề + Tiêu đề SEO + H1`; khác snapshot thì dừng, không đăng.
- Trạng thái `LỖI ĐĂNG` được đưa vào nhóm có thể chạy lại cùng `LỖI KIỂM TRA`.

## 12. Tự khôi phục và nâng phiên phân tích

- Phiên bản 1.2.5 mở tức thì từ session nếu app cùng phiên bản và file nguồn chưa thay đổi.
- Nếu app được nâng phiên bản nhưng file cũ vẫn còn, app nhớ đường dẫn và tự phân tích lại ở thread nền để tạo session mới tương thích.
- Nếu file nguồn đã thay đổi từ lần mở trước, app cũng tự phân tích lại thay vì yêu cầu người dùng bấm nút.
- Nếu file không còn ở đường dẫn cũ, app chỉ thông báo và không tự chọn file khác.

## 13. Giữ nút xác nhận batch đăng luôn hiển thị

- Phiên bản 1.2.6 đổi hộp batch đăng sang bố cục co giãn theo màn hình và Windows scaling.
- Bảng domain co giãn ở giữa; phần tổng kết và nút hành động được giữ ở các hàng riêng phía dưới.
- Nút `Tiếp tục` được đổi thành `Xác nhận batch và đăng` để thể hiện rõ bước tiếp theo.

## 14. Tự lưu workbook ẩn trước Flow 8

- Phiên bản 1.2.7 sửa lỗi Flow 8 báo `có thay đổi chưa Save` khi workbook đang được app mở ẩn.
- `flow_host.py` gọi `Save()` trên workbook ẩn trước khi khởi chạy script đồng bộ URL.
- Flow 8 gọi lại `Save()` và xác nhận trạng thái `Saved` khi nhận `HOTKEYVIP_APP_RUN=1`.
- Chạy Flow 8 độc lập vẫn giữ kiểm tra yêu cầu người dùng Save để không tự ghi một workbook thủ công đang mở ngoài app.

## 15. Chặn mở nhiều app

- Phiên bản 1.2.8 dùng named mutex Windows để mỗi phiên đăng nhập chỉ có một app đối soát.
- Phiên bản 1.2.9 đưa bước kiểm tra dòng trước khi đăng vào cùng hàng đợi Excel; khi Excel bận, app chờ và thử lại thay vì làm hỏng cả batch. Kết quả đăng cũng phải lưu Excel xong mới hiện `[EXCEL SAVED]` và tính bài hoàn thành.
- Phiên bản 1.2.10 dùng cùng một màu cho toàn bộ nút `Chạy`; nút chỉ chuyển xám khi đang bị khóa.
- Phiên bản 1.2.11 thêm nút `Lưu nhật ký`, cho phép lưu toàn bộ nội dung đang hiển thị trong khung log thành file `.log` khi người dùng yêu cầu.
- Phiên bản 1.2.12 tô riêng nút Flow 3 `Viết bài + tạo ảnh` màu xanh dương và Flow 5 `Đăng bài CMS` màu cam; các flow còn lại giữ màu xanh ngọc.
- Phiên bản 1.3.0 đóng gói app theo cấu trúc di động: 8 flow nằm trong `app_flows` với tên cố định, ngắn và không chứa số phiên bản triển khai.
- Phiên bản 1.3.1 thêm nút `Chuyển URL ngày đăng mới nhất`: chỉ nối URL công khai của các bài `ĐÃ ĐĂNG` trong ngày mới nhất vào hàng chờ Submit với trạng thái `PENDING`, chống trùng và giữ nguyên toàn bộ URL cũ.
- Phiên bản 1.3.2 đổi nút chuyển URL sang chế độ thay danh sách: sao lưu rồi xóa URL cũ, nạp URL ngày đăng mới nhất và hỏi có tự mở app Submit sau khi chuyển hay chỉ chuyển dữ liệu.
- Phiên bản 1.3.3 bỏ câu hỏi mở app sau khi chuyển và thêm nút `Mở Submit` riêng ngay cạnh `Chuyển URL`.
- Phiên bản 1.3.4 chạy Flow 4 cùng tiến trình `flow_host` và truyền trực tiếp workbook Excel ẩn, sửa lỗi `Không tìm thấy workbook Excel đang mở`; chạy Flow 4 độc lập vẫn dùng workbook active như cũ.
- Phiên bản 1.3.5 chuẩn hóa nhận đúng workbook theo đường dẫn cho Flow 1, 2, 4, 6 và 8. Flow 7 tự mở đúng đường dẫn do dùng multiprocessing. Nếu Excel ẩn do app tạo mở lỗi, app chỉ đóng phiên của chính nó rồi thử mở lại; không đóng Excel khác của người dùng.
- Phiên bản 1.3.6 thêm đối chiếu trực tiếp `VIẾT_BÀI - Hoàn tất OK` với `ĐĂNG_BÀI - Tổng dòng`. Tổng quan đổi đỏ và báo số dòng thiếu/dư khi hai số lệch; ô `Tiến độ - Chưa hoàn tất` trùng lặp được thay bằng thẻ đối chiếu này.
- Phiên bản 1.3.7 thêm đối chiếu `KẾ_HOẠCH - URL hợp lệ` với `ĐĂNG_BÀI - Đã đăng`. Ô `Hợp lệ - Đã đăng, đã xóa tài nguyên` được thay bằng thẻ so sánh trực tiếp; trạng thái Tổng quan báo riêng từng cặp hoặc báo cả hai cặp cùng lệch.
- Phiên bản 1.3.8 truyền động đường dẫn gốc app qua `PYTHONPATH` cho các flow chạy tiến trình con, sửa lỗi Flow 5 không import được `excel_audit_app.publish_plan` sau khi đóng gói hoặc di chuyển thư mục app.
- Phiên bản 1.3.9 cho phép chọn thủ công một danh mục riêng cho từng tên miền trong cửa sổ batch đăng bài. Số bài có thể đăng và số sẽ đăng cập nhật ngay; lựa chọn được truyền vào Flow 5 để dựng đúng batch thực tế.
- Phiên bản 1.4.0 tách 8 flow khỏi dữ liệu Tổng quan: flow chỉ dùng đường dẫn Excel đang chọn, không bắt phân tích lại và app không tự phân tích sau flow hoặc khi mở lại file đã thay đổi. Bộ điều phối bám đúng workbook đang mở theo đường dẫn đầy đủ; nếu file chưa mở thì tạo Excel ẩn riêng, chỉ đóng phiên do app tạo. Flow 3 và Flow 7 cũng được chuẩn hóa tìm đúng workbook theo đường dẫn. Flow 5 và Flow 7 luôn chạy 5 worker, không hiện hộp hỏi số luồng.
- Phiên bản 1.4.1 giữ nút `Mở Submit` luôn hoạt động trong khi flow đang chạy; nút `Chuyển URL` vẫn khóa để tránh đọc Excel giữa lúc dữ liệu đang được ghi.
- Nếu người dùng mở lần hai, tiến trình mới tìm cửa sổ có sẵn, khôi phục cửa sổ nếu đang thu nhỏ và đưa nó lên trước rồi thoát.
- Nếu cửa sổ đầu tiên đang khởi tạo và chưa tìm thấy HWND, app hiện thông báo đang có một phiên chạy.
- Khóa chỉ áp dụng cho app giao diện; các tiến trình `flow_host` và Worker con không bị chặn.
