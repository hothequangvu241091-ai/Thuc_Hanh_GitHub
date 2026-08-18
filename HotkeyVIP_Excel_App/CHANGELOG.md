# Lịch sử thay đổi HotkeyVIP Excel App

Mỗi đợt sửa phải thêm một mục tại đây và tạo Git tag cùng số phiên bản.
Không dùng tên file kiểu `final`, `final2`, `moi_nhat` để thay cho lịch sử.

## [1.5.12] - 2026-08-17

### Flow 03 - Viết bài + tạo ảnh (engine V2.24)

- Trước đây CMD đã xử lý bài nhưng nhãn Worker có thể vẫn hiện `Đang chờ`, do
  cửa sổ chỉ phụ thuộc vào các message tuần tự trong hàng đợi giao diện.
- Lưu thêm trạng thái mới nhất của từng Worker trong RAM và đồng bộ cửa sổ mỗi
  200 ms; trạng thái đầu phiên đổi thành đang mở Edge/chuẩn bị phiên làm việc.
- Kiểm tra bằng biên dịch Python và test trạng thái giao diện Flow 03.

## [1.5.11] - 2026-08-17

### Flow 06 - Lấy URL từ ID CMS (engine V1.7)

- Ô `URL đã đăng` được tạo theo quy luật H1 nay được tô vàng nhạt và thêm Note
  `URL tự tạo từ H1 do HEAD không chuyển hướng; cần kiểm tra lại.`
- Dấu màu và Note chỉ là thông tin trực quan, không thay đổi giá trị URL và không
  ảnh hưởng dữ liệu đầu vào của các flow khác.

## [1.5.10] - 2026-08-17

### Flow 06 - Lấy URL từ ID CMS (engine V1.6)

- Vẫn ưu tiên HEAD song song tối đa 20 URL và không dùng GET.
- Nếu HEAD lỗi hoặc URL không chuyển hướng, tự tạo URL thật theo quy luật
  `slug từ H1-ID danh mục-ID CMS.html`; chỉ giữ URL tạm khi H1 trống.
- Nhật ký phân biệt URL do server chuyển hướng và URL được tạo từ H1.

## [1.5.9] - 2026-08-17

### Flow 06 - Lấy URL từ ID CMS (engine V1.5)

- Bỏ hoàn toàn fallback GET và retry vì website có thể giữ kết nối đến timeout,
  làm flow chờ lâu nhưng vẫn không lấy được đủ URL.
- Khôi phục cơ chế chỉ HEAD song song tối đa 20 URL; URL chưa chuyển hướng vẫn
  được giữ lại để lần chạy sau thử HEAD lại.

## [1.5.8] - 2026-08-17

### Flow 06 - Lấy URL từ ID CMS (engine V1.4)

- Giảm fallback GET từ 2 xuống 1 kết nối đồng thời cho mỗi tên miền để tránh
  website yếu từ chối hoặc làm timeout nhiều kết nối cùng lúc.
- HEAD vẫn chạy song song tối đa 20 URL; các tên miền khác nhau vẫn xử lý độc lập.

## [1.5.7] - 2026-08-17

### Flow 06 - Lấy URL từ ID CMS (engine V1.3)

- HEAD vẫn chạy song song tối đa 20 URL như trước.
- Chỉ URL phải fallback GET mới vào hàng chờ riêng; mỗi tên miền tối đa 2 GET
  đồng thời, tên miền khác vẫn chạy song song.
- GET chưa lấy được URL thật được thử tối đa 3 lần, nghỉ tăng dần giữa các lần.
- Ghi chi tiết lỗi cuối nếu cả 3 lượt GET đều thất bại.
- Kiểm tra bằng biên dịch Python, kiểm tra giới hạn GET và test liên quan.

## [1.5.6] - 2026-08-17

### Kết thúc flow nhanh và dọn Excel ẩn

- Trước đây sau Flow 1–8, app tự phân tích lại toàn bộ workbook trước khi mở
  lại các nút Chạy, gây chờ lâu với file nhiều nghìn dòng.
- Bỏ quét tự động sau flow; file vẫn được lưu và các nút được mở lại ngay. Báo
  cáo được đánh dấu chưa cập nhật và chỉ quét khi người dùng bấm `Phân tích`.
- Giữ khóa workbook và Save cuối để bảo vệ dữ liệu.
- Sau khi đóng workbook và gọi `Excel.Quit()`, chỉ chờ 1 giây; Excel ẩn do app
  sở hữu chưa thoát sẽ được dọn ngay thay vì để lần chạy kế tiếp xử lý.
- Kiểm tra bằng biên dịch Python và test liên quan.

## [1.5.5] - 2026-08-17

### Flow 06 - Lấy URL từ ID CMS (engine V1.2)

- Thử HEAD trước như cũ; nếu HEAD lỗi, timeout hoặc không đổi URL tạm thì mới
  fallback GET cho riêng URL đó để nhận chuyển hướng như trình duyệt.
- URL thật vẫn được bỏ qua ở lần chạy sau; URL chưa giải quyết vẫn được giữ để
  tự thử lại.
- Nhật ký tách rõ tổng URL đã thử, số URL thật lấy được và số URL vẫn còn tạm.
- Kiểm tra bằng biên dịch Python, kiểm tra nhánh HEAD/GET và test liên quan.

## [1.5.4] - 2026-08-17

### Flow 06 - Lấy URL từ ID CMS (engine V1.1)

- Trước đây URL tạm `linkrutgon-...` đã ghi vào Excel bị coi là URL có sẵn
  và bị bỏ qua ở mọi lần chạy sau.
- Nhận diện URL tạm để tự đưa vào hàng xử lý lại; URL thật vẫn được bỏ qua.
- Chỉ dùng yêu cầu HEAD để kiểm tra chuyển hướng. Nếu không lấy được URL thật,
  giữ URL tạm cho lần chạy sau; không dùng GET để tải nội dung trang.
- Hiển thị số URL tạm được xử lý lại trong phần kết quả.
- Kiểm tra bằng biên dịch Python và kiểm tra nhận diện URL tạm.

## [1.5.3] - 2026-08-17

### Flow 07 - Bài viết liên quan (engine V2.5)

- Trước đây cả 5 worker dùng chung một khóa đăng nhập, khiến worker ở các tên
  miền khác nhau vẫn phải chờ nhau khi phiên đăng nhập hết hạn.
- Đổi sang khóa đăng nhập riêng cho từng tên miền, tương tự khóa Save: worker
  cùng tên miền vẫn tuần tự, worker khác tên miền được đăng nhập song song.
- Bổ sung tên miền vào thông báo chờ khóa để dễ theo dõi trong nhật ký.
- Kiểm tra bằng biên dịch Python và bộ test `test_publish_review`.

## [1.5.2] - 2026-08-17

### Theo dõi đăng bài

- Trước đây nút bước 2 chỉ chuyển một dòng đang chọn từ `LỖI ĐĂNG` sang
  `LỖI KIỂM TRA`, nên phải thao tác lặp lại cho từng bài.
- Đổi nút bước 2 thành chuyển toàn bộ dòng `LỖI ĐĂNG` trong danh sách hiện tại
  sang `LỖI KIỂM TRA` bằng một lần xác nhận; bước này không chạy đăng bài.
- Nút bước 3 tiếp tục đăng lại hàng loạt toàn bộ dòng `LỖI KIỂM TRA` như trước.
- Kiểm tra bằng biên dịch Python và bộ test `test_publish_review`.

## [1.5.1] - 2026-08-16

### Flow 07 - Bài viết liên quan (engine V2.4)

- Sửa Dừng an toàn: đóng multiprocessing queue và thoát hẳn flow.
- Luôn đóng Edge khi worker lỗi, bỏ `detach=True` để tránh khóa profile.
- Ghi đầy đủ traceback vào log worker.
- Edge khởi động lỗi được thử lại tối đa 3 lần.
- Bài lỗi được đóng/mở Edge và thử lại tối đa 3 lượt trong cùng phiên.
- Kích hoạt `ExcelWriterQueue`; điều phối không còn `workbook.Save()` sau từng bài.
- Lưu Excel theo lô 10 kết quả và bắt buộc lưu lần cuối khi kết thúc.
- Quản lý domain bằng bộ đếm thay vì `set`.
- Phát hiện worker chết bất thường, thu hồi và giao bài cho worker khác một lần.

### Sắp xếp dự án

- Gom tool phục hồi `WORD_ERROR` dùng một lần vào
  `_archive/one_off_tools/word_error_20260816`.
- Giữ mapping và log sinh ra trong `outputs` nhưng không đưa chúng vào Git.
- Thiết lập Git cục bộ, `.gitignore` và quy trình tạo mốc phiên bản.
- Thêm `AGENTS.md` để Codex tự kiểm tra lịch sử, cập nhật changelog, kiểm thử và
  tạo mốc Git sau mỗi đợt sửa mà người dùng không phải nhắc lại.

## [1.5.0] - 2026-08-09

- Thêm tab Theo dõi đăng bài.
- Cho phép xem lỗi, mở Word/URL ChatGPT, nhập ID CMS và đăng lại dòng lỗi.
- Thêm nút Mở Excel và tự làm mới dữ liệu theo dõi sau flow.

Chi tiết các phiên bản cũ hơn nằm tại `_docs/NHAT_KY_TICH_HOP_FLOW.md`.
