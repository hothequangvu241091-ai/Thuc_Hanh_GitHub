# App đối soát nội dung Excel

Phiên bản hiện tại: **1.5.0**.

## Theo dõi đăng bài

Sau khi bấm **Phân tích**, tab **Theo dõi đăng bài** hiển thị các dòng lỗi trong `DANG_BAI`
và các bài đã đăng trong ngày hiện tại. Chọn một dòng để mở Word hoặc URL ChatGPT.
Có thể nhập ID CMS còn thiếu, sau đó bấm **Cập nhật ID & đăng lại LỖI KIỂM TRA**.
App chỉ đăng lại đúng các dòng `LỖI KIỂM TRA` bằng một worker; các lỗi khác không được đăng lại.

Tài liệu kỹ thuật dành cho bảo trì và phát triển: `TAI_LIEU_KY_THUAT_APP_DOI_SOAT_EXCEL.md`.

## Mở app

Nhấp đúp `CHAY_APP.cmd`.

App chạy local trên Windows và hỗ trợ `.xlsx` và `.xlsm`. Phần phân tích không cần mở Microsoft Excel. Khi bấm **Xuất kết quả**, máy cần có Microsoft Excel để app ghi an toàn trên một bản sao tạm.

Mỗi tài khoản Windows chỉ mở một cửa sổ app. Nếu nhấp mở lần nữa, app đưa cửa sổ đang chạy lên trước và không tạo tiến trình thứ hai, tránh hai flow cùng tranh workbook, Edge/Word hoặc session.

## Cách dùng

1. Bấm **Chọn file** và chọn workbook cần kiểm tra.
2. Bấm **Phân tích**.
3. Xem số liệu ở các tab **Tổng quan**, **Kế hoạch**, **Viết bài**, **Đăng bài**, **Chi tiết đối soát** và **Khôi phục DANG_BAI**.
4. Nếu cần ghi kết quả, bấm **Xuất kết quả** và chọn tên file mới.

Trước khi bấm **Xuất kết quả**, app không sửa file Excel. Khi xuất, app tạo một workbook mới, thêm hoặc cập nhật sheet `Tong_all` và cập nhật `Trạng thái nguồn` trong `KE_HOACH` của bản mới.

Trước khi công bố file mới, app tự so sánh dữ liệu của `KE_HOACH`, `VIET_BAI`, `DANG_BAI` với file nguồn và kiểm tra VBA. Nếu phát hiện thay đổi ngoài phạm vi cho phép, app hủy file tạm và báo lỗi.

## Phiên làm việc

Sau khi phân tích, app tự lưu kết quả gần nhất trong thư mục dữ liệu local của Windows. Khi mở lại app, kết quả cũ được khôi phục.

- Nếu file nguồn không thay đổi, có thể tiếp tục xem hoặc xuất.
- Nếu file nguồn đã thay đổi, app tự phân tích lại ở nền khi mở.
- Nếu app vừa nâng phiên bản, app vẫn nhớ đường dẫn file cũ và tự tạo phiên phân tích tương thích; không cần bấm **Phân tích**.
- Nếu file nguồn bị di chuyển hoặc xóa, vẫn xem được kết quả cũ nhưng không thể xuất.
- Nút **Xóa phiên** chỉ xóa dữ liệu phiên của app, không xóa hoặc sửa file Excel.

## Quy tắc chính

- Tên sheet được so sánh không phân biệt chữ hoa/thường.
- URL hợp lệ phải bắt đầu bằng `http://` hoặc `https://`.
- Combo 4: Tên miền + Tiêu đề SEO + H1 + Từ khóa/Tiêu đề.
- Kiểm tra trùng trong `VIET_BAI` dùng Combo 4: Tên miền + Tiêu đề SEO + H1 + Từ khóa.
- Trùng Tiêu đề SEO + H1 + Từ khóa nhưng khác tên miền không tính là lỗi.
- Trong `DANG_BAI`, cột `Tiêu đề` được dùng như Từ khóa.
- Nếu một dòng `KE_HOACH` vừa trùng vừa thiếu dữ liệu, trạng thái ưu tiên là `Bài viết trùng`.

## Cách đọc kết quả phiên bản 1.4.1

- **Lỗi dữ liệu**: trùng, thiếu Combo 4 hoặc dữ liệu tồn tại ở sheet sau nhưng bị thiếu ở sheet nguồn.
- **Cần khôi phục DANG_BAI**: `KE_HOACH` có URL hợp lệ nhưng Combo 4 không còn trong `DANG_BAI`.
- **Chưa chuyển sang DANG_BAI**: tiến độ bình thường, không tính là lỗi dữ liệu.
- **Đã đăng, đã xóa tài nguyên**: `VIET_BAI` đã OK, thiếu Word/ảnh nhưng tìm thấy bài `ĐÃ ĐĂNG` trong `DANG_BAI`.

Nhấp vào các thẻ màu ở **Tổng quan**, hoặc nhấp đúp vào ô chênh lệch trong bảng tên miền, để mở danh sách dòng liên quan.

Trong tab **Khôi phục DANG_BAI**, nút **Sao chép toàn bộ để dán vào DANG_BAI** sao chép dữ liệu theo đúng thứ tự cột của sheet `DANG_BAI`, không kèm dòng tiêu đề. Các trường không lấy được từ `KE_HOACH` hoặc `VIET_BAI` được để trống.

Nếu không muốn dán thủ công, bấm **Tạo file mới + khôi phục**. App yêu cầu chọn tên file mới, nối toàn bộ dòng cần khôi phục vào `DANG_BAI`, giữ nguyên file gốc và tự phân tích lại bản mới. Nút này chỉ hoạt động khi kết quả không có lỗi dữ liệu màu đỏ.

## Chạy các flow trong app

Tab **Công việc** chứa các flow theo đúng thứ tự vận hành. Tất cả đều dùng file Excel đang được chọn trên app; không cần tự mở Excel hoặc chạy từng file `.py`.

1. Phân tích file trước.
2. Mở tab **Công việc**.
3. Bấm **Chạy** tại flow cần dùng và đọc kỹ hộp xác nhận.
4. Theo dõi nội dung tại khung nhật ký bên phải.
5. Khi flow kết thúc, app tự đọc lại workbook và cập nhật toàn bộ số liệu.

Mỗi lần chỉ chạy một flow. App mở Microsoft Excel ở chế độ ẩn để các flow cũ vẫn giữ được công thức, macro và cách ghi COM; máy vẫn phải cài Excel nhưng người dùng không phải mở cửa sổ Excel. Nếu file đang được mở thủ công trong Excel và bị khóa, hãy đóng file đó trước khi chạy flow.

Các flow có thể đăng bài, sửa CMS hoặc cập nhật file công ty luôn có hộp xác nhận riêng. Bấm **Phân tích** không tự chạy bất kỳ flow nào.

Khi chạy **8. Đồng bộ URL về file công ty**, app tự lưu workbook Excel ẩn trước khi đọc và đồng bộ. Người dùng không cần mở file Excel để bấm Save. Khi chạy script Flow 8 độc lập ngoài app, kiểm tra yêu cầu Save thủ công vẫn được giữ.

### Ưu tiên viết bài

Khi bấm **3. Viết bài + tạo ảnh**, app tự đọc `VIET_BAI` và hiển thị số bài chưa hoàn tất của từng tên miền, số bài lỗi, vị trí `OK OK` cuối cùng, dòng bắt đầu chạy thường và các bài đầu tiên trong hàng chờ. Người dùng không cần mở Excel để tìm mốc. Sau đó có thể chọn một tên miền và số bài cần ưu tiên. Hàng chờ được xếp theo thứ tự: bài đang lỗi, bài của tên miền ưu tiên, rồi các bài bình thường sau mốc `OK OK`. Các dòng Excel không bị di chuyển và một dòng không được xếp trùng hai lần.

Trước khi một Worker bắt đầu xử lý, app xác nhận lại `Tên miền + Từ khóa` vẫn nằm ở đúng dòng đã nạp vào RAM. Nếu dòng đã bị sắp xếp, chèn, xóa hoặc thay đổi, flow dừng an toàn và không ghi kết quả vào dòng đó.

### Đăng bài cân bằng theo tên miền và danh mục

Khi bấm **5. Đăng bài CMS**, app tự đọc `DANG_BAI` và cho nhập số bài tối đa mỗi tên miền. Với từng tên miền, app chọn đúng một danh mục có nhiều bài hợp lệ nhất và lấy `min(số yêu cầu, số bài hiện có)`. Ví dụ yêu cầu 7 nhưng danh mục chỉ có 5 bài thì đăng 5; app không ghép bài từ danh mục khác.

Bảng xem trước hiển thị tên miền, danh mục được chọn, số bài có thể đăng và số bài thực tế trong batch. V2.10 dùng đúng batch này, không hỏi lại tổng số bài và dừng sau khi xử lý hết batch. Ngay trước khi giao Worker, flow kiểm tra lại `Tên miền + Danh mục + Tiêu đề + Tiêu đề SEO + H1`; nếu dữ liệu đã đổi thì dừng trước khi đăng.

Sau khi kiểm tra bảng, bấm **Xác nhận batch và đăng** ở góc dưới bên phải. App hiện thêm hộp xác nhận tác động CMS; đồng ý lần nữa thì V2.10 mới bắt đầu chạy.
