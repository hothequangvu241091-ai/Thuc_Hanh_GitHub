# Kế hoạch hoàn thiện hệ thống Submit URL tự động

Cập nhật: 30/07/2026

## 1. Phạm vi bắt buộc

Chỉ làm việc trong:

```text
D:\CodexProjects\Hotkeyvip\06_du_lieu_chay\submit_edge_profiles
```

Không đưa file phụ vào `HotkeyVIP-Studio`, không sửa dự án khác và không lưu
mật khẩu Gmail.

Phải đọc thêm:

```text
README_HE_THONG_SUBMIT.md
```

## 2. Trạng thái hiện tại

File GUI chính:

```text
_he_thong\manage_submit_edge_profiles.ps1
```

File mở GUI:

```text
QUAN_LY_PROFILE_SUBMIT.bat
```

File cấu hình Gmail, tên miền và profile:

```text
_he_thong\submit_edge_profiles.json
```

File lưu danh sách URL lâu dài:

```text
_he_thong\submit_url_history.json
```

Các chức năng đang có:

- Quản lý Gmail, tên miền và ánh xạ Edge profile.
- Tìm profile/Gmail theo tên miền hoặc URL.
- Mở đúng property Google Search Console.
- Submit thủ công một URL bằng nút `SUBMIT`.
- Điền URL Inspection bằng DOM/CDP.
- Click `Yêu cầu lập chỉ mục`.
- Ô ngoài màn hình chính hiển thị danh sách URL đã lưu.
- Nút `NẠP THÊM URL` mở cửa sổ riêng để dán URL, mỗi dòng một URL.
- URL mới được thêm với trạng thái mặc định `PENDING`.
- URL trùng không được thêm lại.
- Nút `XEM / SỬA` mở bảng danh sách chi tiết.
- Có thể sửa URL, trạng thái, ghi chú hoặc xóa URL.
- Có thể chọn nhiều dòng và áp dụng chung một trạng thái.
- Có bộ lọc trạng thái và sắp xếp cột A-Z/Z-A.

Chưa có:

- Luồng tự động chưa được chạy thử bằng URL thật sau khi triển khai.
- Thông báo quota thực tế vẫn chưa được quan sát để bổ sung câu chữ chính xác.
- Chưa xác nhận trực tiếp mọi biến thể giao diện GSC khi URL đã index.

Đã triển khai nhưng cần test tuần tự:

- Nút `SUBMIT TỰ ĐỘNG`.
- Hộp hỏi số URL tối đa.
- Nút `DỪNG SAU URL HIỆN TẠI`.
- Vòng lặp xử lý nhiều URL tuần tự.
- Tự đọc kết quả và ghi trạng thái sau mỗi URL.
- Bộ đếm lỗi theo tên miền/Gmail.
- Chuyển Gmail/profile tự động.
- Xác minh property trước khi điền URL.
- Log JSONL, nội dung DOM và ảnh khi lỗi.

File bộ điều phối:

```text
_he_thong\auto_submit_queue.py
```

## 3. Các trạng thái URL đã chốt

Trong file có thể dùng mã ổn định; trên GUI phải hiện tiếng Việt:

| Mã nội bộ | Hiển thị |
|---|---|
| `PENDING` | `CHƯA SUBMIT` |
| `RUNNING` | `ĐANG SUBMIT` |
| `SUBMITTED` | `HOÀN THÀNH` |
| `ERROR` | `LỖI` |
| `QUOTA` | `VƯỢT HẠN NGẠCH` |
| `SKIPPED` | `BỎ QUA` |
| `UNMAPPED` | `KHÔNG TÌM THẤY PROFILE` |

Khi nạp URL mới:

```text
PENDING / CHƯA SUBMIT
```

## 4. Nút SUBMIT TỰ ĐỘNG

Khi người dùng bấm nút:

1. Hỏi muốn xử lý tối đa bao nhiêu URL.
2. Trong giai đoạn test nên mặc định là `3`.
3. Chỉ lấy URL có trạng thái `PENDING`.
4. Giới hạn tính theo số URL riêng biệt được lấy khỏi hàng đợi.
5. Việc thử lại cùng URL bằng Gmail phụ không làm tăng giới hạn.
6. Chạy tuần tự từng URL, chưa chạy nhiều luồng.
7. Lưu kết quả ngay sau từng bước quan trọng và sau từng URL.

Phải có nút:

```text
DỪNG SAU URL HIỆN TẠI
```

Không dừng giữa lúc Google đang xử lý một URL.

## 5. Cách chọn Gmail/profile

Với mỗi URL:

1. Tách tên miền.
2. Tìm tất cả Gmail chứa tên miền đó trong `accounts`.
3. Từ Gmail tìm profile đang được gán.
4. Không suy đoán Gmail/profile chưa có trong cấu hình.
5. Ưu tiên profile có số nhỏ nhất, ví dụ `submit_01` trước `submit_07`.
6. Profile số lớn hơn là phương án phụ.

Trước khi nhập URL phải xác minh profile thực sự có quyền với tên miền:

1. Đọc ô có `aria-label` bắt đầu bằng `Kiểm tra mọi URL trong`.
2. Tên miền trong `aria-label` phải trùng property dự kiến.
3. Nếu Google báo không có quyền, property khác hoặc profile chưa đăng nhập:
   - Không click submit.
   - Không tính vào bộ đếm lỗi URL/tên miền.
   - Ghi log `PROFILE_MAPPING_REJECTED`.
   - Loại cặp tên miền + Gmail đó khỏi phiên hiện tại.
   - Thử Gmail/profile phụ nếu có.
4. Không tự xóa ánh xạ khỏi file cấu hình; người dùng sửa thủ công sau khi
   xem log.

Nếu không có Gmail/profile hợp lệ:

```text
KHÔNG TÌM THẤY PROFILE
```

Sau đó tiếp tục URL kế tiếp.

## 6. Luồng xử lý một URL

```text
CHƯA SUBMIT
→ ĐANG SUBMIT
→ mở đúng Edge profile
→ mở đúng property GSC
→ điền URL Inspection
→ Enter
→ chờ trang kết quả
→ nếu URL đã index: BỎ QUA
→ nếu chưa index: click Yêu cầu lập chỉ mục
→ chờ hộp kết quả
→ phân loại
→ ghi file URL và log
→ URL kế tiếp
```

Nếu phát hiện URL đã nằm trên Google:

```text
BỎ QUA
Ghi chú: Đã index, không dùng lượt submit
```

## 7. Nhận diện kết quả Google

### Hoàn thành

Có một trong các nội dung:

```text
Đã yêu cầu lập chỉ mục
URL đã được thêm vào hàng đợi ưu tiên
```

Kết quả:

```text
HOÀN THÀNH
```

### Lỗi thông thường

Thông báo đã quan sát thực tế:

```text
Rất tiếc! Đã xảy ra sự cố
Chúng tôi gặp sự cố khi gửi yêu cầu lập chỉ mục của bạn.
Vui lòng thử lại sau.
```

Kết quả:

```text
LỖI
```

Thông báo này không phải vượt hạn ngạch.

### Vượt hạn ngạch

Chỉ phân loại là `VƯỢT HẠN NGẠCH` khi nội dung Google nói rõ về giới hạn
hoặc hạn ngạch. Không dùng quy tắc “không thành công thì mặc định là quota”.

Thông báo quota thực tế chưa được quan sát đầy đủ. Khi gặp lần đầu phải lưu
nguyên văn DOM và ảnh để bổ sung bộ nhận diện.

### Lỗi hệ thống

Các trường hợp:

- Edge bị đóng.
- Mất kết nối CDP.
- Profile mất đăng nhập.
- Không có quyền property.
- Không tìm thấy ô nhập hoặc nút do giao diện Google thay đổi.
- Script bị lỗi.

Đây không phải lỗi tên miền/Gmail. Phải ghi log và dừng toàn phiên để kiểm tra,
không được tự chuyển Gmail.

### Timeout

Trong giai đoạn test:

- Chờ tối đa đề xuất `180 giây`.
- Nếu không xác định được kết quả: ghi `TIMEOUT`.
- Dừng toàn phiên.
- Không tự chuyển Gmail để tránh submit trùng khi Google có thể đã nhận yêu cầu.

## 8. Quy tắc lỗi theo tên miền

Bộ đếm lỗi phải tách riêng theo:

```text
Tên miền + Gmail
```

Quy tắc:

1. Một URL lỗi thông thường: đánh dấu URL là `LỖI`, tiếp tục URL khác.
2. Hai URL liên tiếp của cùng tên miền cùng lỗi trên cùng Gmail:
   - Ngừng dùng Gmail đó cho riêng tên miền này trong phiên.
   - Thử lại URL gây ra lỗi lần thứ hai bằng Gmail/profile phụ.
3. Gmail phụ có bộ đếm riêng bắt đầu từ `0`.
4. Gmail phụ cũng lỗi hai lần liên tiếp trên tên miền:
   - Dừng tên miền đó trong phiên.
   - Các URL chưa chạy của tên miền vẫn giữ `CHƯA SUBMIT`.
   - Ghi chú lý do tạm dừng.
5. Nếu một URL thành công trên cặp tên miền + Gmail thì bộ đếm lỗi liên tiếp
   của cặp đó trở về `0`.

## 9. Quy tắc hai tên miền lỗi liên tiếp

Mỗi Gmail còn có bộ đếm số tên miền khác nhau bị lỗi liên tiếp.

Một tên miền được tính là “tên miền lỗi” khi nó đã chạm mức:

```text
2 URL lỗi liên tiếp trên Gmail đó
```

Nếu hai tên miền khác nhau liên tiếp cùng chạm mức này trên một Gmail:

```text
Dừng toàn bộ Gmail đó trong phiên
```

Sau đó:

- Các tên miền của Gmail đó chuyển sang Gmail/profile phụ nếu có.
- Gmail phụ dùng bộ đếm riêng.
- Nếu có một tên miền submit thành công xen giữa, bộ đếm tên miền lỗi liên
  tiếp của Gmail trở về `0`.

## 10. Quy tắc vượt hạn ngạch

Quota áp dụng cho toàn Gmail trong phiên:

```text
Gmail gặp quota
→ dừng Gmail đó ngay
→ tất cả tên miền thuộc Gmail đó chuyển sang Gmail/profile phụ nếu có
```

Nếu Gmail phụ cũng quota:

- Dừng các tên miền không còn Gmail hợp lệ.
- URL chưa chạy vẫn giữ `CHƯA SUBMIT`, kèm ghi chú tạm dừng.
- URL của Gmail/tên miền khác vẫn tiếp tục.

Khóa Gmail và các bộ đếm chỉ tồn tại trong phiên `SUBMIT TỰ ĐỘNG` hiện tại.
Khi người dùng chạy lại thủ công, tất cả khóa và bộ đếm bắt đầu lại từ đầu.

## 11. Log bắt buộc

Không dùng `submit_url_history.json` làm log duy nhất.

Đề xuất:

```text
_he_thong\logs\YYYY-MM-DD\run_YYYYMMDD_HHMMSS.jsonl
```

File log phải ghi nối tiếp, không ghi đè sự kiện trước.

Mỗi sự kiện cần có:

```text
timestamp
runId
step
url
domain
gmail
profile
accountPriority
attempt
result
rawGoogleMessage
elapsedSeconds
domainGmailConsecutiveErrors
gmailConsecutiveFailedDomains
action
```

Ví dụ `step`:

```text
RUN_STARTED
URL_SELECTED
PROFILE_SELECTED
GSC_OPENED
URL_ENTERED
INSPECTION_READY
INDEX_REQUEST_CLICKED
RESULT_DETECTED
PROFILE_FALLBACK
DOMAIN_PAUSED
GMAIL_PAUSED
SYSTEM_ERROR
RUN_STOPPED
RUN_FINISHED
```

Phải flush/ghi file ngay sau mỗi sự kiện để app tắt giữa chừng vẫn còn log.

## 12. Hiển thị log trong app

Thêm:

```text
XEM LOG PHIÊN CHẠY
```

Và một dòng tiến trình trực tiếp, ví dụ:

```text
[12:45:10] kythuatmarketing.com → submit_01 → ĐANG SUBMIT
[12:46:02] HOÀN THÀNH
[12:46:05] Đang lấy URL tiếp theo...
```

Khi có `ERROR`, `QUOTA`, `TIMEOUT` hoặc lỗi chưa nhận diện:

- Lưu nguyên văn nội dung DOM liên quan.
- Trong giai đoạn test, lưu thêm ảnh chụp trang lỗi vào thư mục log.
- Không lưu mật khẩu.

## 13. Báo cáo cuối phiên

Khi kết thúc hoặc người dùng yêu cầu dừng, hiển thị:

```text
Tổng URL đã lấy
Hoàn thành
Bỏ qua vì đã index
Lỗi
Vượt hạn ngạch
Không tìm thấy profile
Tên miền bị tạm dừng
Gmail bị tạm dừng
Thời gian chạy
Đường dẫn file log
```

Không dùng hộp thoại xác nhận cho từng URL. Chỉ có một báo cáo cuối phiên.

## 14. Thứ tự code đề xuất

1. Sao lưu file GUI hiện tại.
2. Tạo module ghi log JSONL.
3. Tạo trạng thái phiên chạy trong bộ nhớ.
4. Tạo hộp hỏi số URL tối đa.
5. Thêm nút `SUBMIT TỰ ĐỘNG`.
6. Thêm nút `DỪNG SAU URL HIỆN TẠI`.
7. Tách hàm tìm các profile ứng viên theo thứ tự số.
8. Nối `cdp_watch_indexing_result.py` vào luồng submit.
9. Thêm nhận diện URL đã index.
10. Thêm xử lý `SUCCESS`, `ERROR`, `QUOTA`, `TIMEOUT`, lỗi hệ thống.
11. Thêm bộ đếm lỗi theo tên miền + Gmail.
12. Thêm bộ đếm hai tên miền lỗi liên tiếp của Gmail.
13. Thêm chuyển Gmail phụ.
14. Ghi trạng thái URL ngay sau mỗi kết quả.
15. Thêm cửa sổ log và báo cáo cuối phiên.
16. Test trước với 1 URL, sau đó 3 URL.
17. Chỉ tăng số lượng khi log xác nhận đúng.

## 15. Việc chưa làm ở giai đoạn này

- Chưa chạy nhiều profile song song.
- Chưa đọc/ghi Excel.
- Chưa chạy theo lịch.
- Chưa tự mở lại Gmail bị khóa.
- Chưa tự động chạy lại URL lỗi từ phiên trước.

Các chức năng này chỉ làm sau khi luồng tuần tự đã ổn định.

## 16. Chỉ dẫn cho task/tài khoản tiếp theo

```text
Đọc README_HE_THONG_SUBMIT.md và KE_HOACH_SUBMIT_TU_DONG.md trước.
Chỉ làm trong submit_edge_profiles.
Không suy đoán Gmail/profile.
Không tự click URL thật ngoài số lượng người dùng đã xác nhận.
Triển khai luồng tuần tự trước, chưa chạy nhiều luồng.
Giữ nguyên quy tắc lỗi, quota và log đã chốt trong tài liệu.
```
