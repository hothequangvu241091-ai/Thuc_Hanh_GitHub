# Hệ thống Submit URL bằng 14 Edge Profile

Cập nhật gần nhất: 30/07/2026

## 1. Phạm vi dự án

Toàn bộ hệ thống nằm trong:

```text
D:\CodexProjects\Hotkeyvip\06_du_lieu_chay\submit_edge_profiles
```

Không đưa file phụ vào HotkeyVIP-Studio và không sửa dự án khác nếu chưa được yêu cầu.

Nguyên tắc:

- Không lưu mật khẩu Gmail.
- Mỗi thư mục `submit_01` đến `submit_14` giữ một phiên Edge riêng.
- Không suy đoán Gmail đang đăng nhập trong profile.
- Chỉ lưu ánh xạ sau khi đã xác định đúng.
- Tên miền thuộc Gmail, không thuộc trực tiếp profile.
- Profile chỉ được gán Gmail; khi đổi profile, toàn bộ tên miền đi theo Gmail.

## 2. Cấu trúc thư mục hiện tại

```text
submit_edge_profiles\
├── submit_01\ ... submit_14\       Dữ liệu người dùng của từng Edge profile
├── _he_thong\
│   ├── assets\
│   │   └── kiem_tra_url.png
│   ├── cdp_fill_search_console_url.py
│   ├── cdp_watch_indexing_result.py
│   ├── click_image_and_paste.py
│   ├── manage_submit_edge_profiles.ps1
│   ├── open_submit_edge_profile.ps1
│   ├── setup_submit_edge_profiles.ps1
│   ├── submit_edge_profiles.json
│   └── submit_edge_profiles.json.bak
├── DANG_NHAP_LAN_LUOT_14_PROFILE.bat
├── MO_PROFILE_SUBMIT.bat
└── QUAN_LY_PROFILE_SUBMIT.bat
```

File mở GUI:

```text
QUAN_LY_PROFILE_SUBMIT.bat
```

## 3. Mô hình dữ liệu

### 3.1. Gmail và tên miền

Nguồn dữ liệu chính:

```json
{
  "accounts": [
    {
      "email": "example@gmail.com",
      "domains": [
        "example.com"
      ]
    }
  ]
}
```

Mỗi Gmail giữ danh sách tên miền riêng.

### 3.2. Profile và Gmail

```json
{
  "profiles": [
    {
      "id": "submit_01",
      "email": "example@gmail.com"
    }
  ]
}
```

Trường `domains` cũ bên trong từng profile là dữ liệu kế thừa. Code mới phải lấy tên miền từ `accounts`, không lấy từ `profiles[].domains`.

### 3.3. Ánh xạ hiện tại

| Profile | Gmail |
|---|---|
| `submit_01` | `tiepthisaigon.com.vn@gmail.com` |
| `submit_02` | `kimnhung1080@gmail.com` |
| `submit_03` | `giavan1080@gmail.com` |
| `submit_04` | `myphammisa@gmail.com` |
| `submit_05` | `systemga4@gmail.com` |
| `submit_06` | `hoathanhxuan.ads@gmail.com` |
| `submit_07` | `seomasterclub@gmail.com` |
| `submit_08` | `thuyha.mark@gmail.com` |
| `submit_09` | `deponlinevietnam@gmail.com` |
| `submit_10` | `phucdinh1088@gmail.com` |
| `submit_11` | Chưa gán |
| `submit_12` | Chưa gán |
| `submit_13` | Chưa gán |
| `submit_14` | Chưa gán |

Hiện có 14 Gmail và 40 tên miền, chia thành 7 nhóm. Mỗi nhóm tên miền có hai Gmail có quyền.

## 4. Chức năng GUI đã có

### 4.1. Quản lý Gmail

- Một ô tìm/chọn Gmail có autocomplete.
- Nút `OK` để nạp Gmail đang nhập.
- Thêm Gmail mới.
- Sửa Gmail.
- Xóa Gmail.
- Hiển thị và sửa danh sách tên miền, mỗi tên miền một dòng.
- Lưu Gmail và tên miền vào JSON.

### 4.2. Quản lý profile

- Gán Gmail đang chọn vào một profile.
- Để Gmail ở trạng thái chưa gán.
- Xem ngược từ profile để biết Gmail đang được gán.
- Gán nhanh Gmail đang chọn vào profile đang xem.
- Mở trực tiếp Edge profile.
- Gmail cũ của profile bị thay thế sẽ trở thành chưa gán; tên miền của Gmail cũ vẫn được giữ.

### 4.3. Tìm URL hoặc tên miền

- Nhập tên miền hoặc URL đầy đủ.
- Tách hostname khỏi URL.
- Tìm tất cả Gmail chứa tên miền.
- Hiển thị từng kết quả trên một dòng:

```text
Gmail: example@gmail.com | Profile: submit_01
```

- Kết quả nằm trong ô văn bản có thể chọn và copy.
- Nếu tên miền có nhiều Gmail/profile, hiển thị danh sách để chọn.

### 4.4. Hai nút xử lý GSC

`MỞ GSC`:

- Mở đúng Edge profile.
- Mở đúng property Google Search Console.
- Không điền URL và không yêu cầu lập chỉ mục.

`SUBMIT`:

- Tìm Gmail/profile theo URL.
- Nếu có một profile thì chạy luôn.
- Nếu có nhiều profile thì yêu cầu chọn.
- Mở đúng property Search Console.
- Điền URL bài viết bằng DOM/CDP.
- Nhấn Enter.
- Chờ trang Kiểm tra URL.
- Tìm và click nút `Yêu cầu lập chỉ mục`.
- Hiện chưa tự ghi kết quả vào lịch sử.

## 5. Cách tự động hóa Google Search Console

### 5.1. Cổng DevTools

Mỗi profile có một cổng riêng:

| Profile | Port |
|---|---:|
| `submit_01` | 9301 |
| `submit_02` | 9302 |
| ... | ... |
| `submit_14` | 9314 |

Edge được mở với:

```text
--remote-debugging-address=127.0.0.1
--remote-debugging-port=93XX
```

Nếu profile đang chạy từ trước mà không có cổng DevTools, cần đóng profile đó và mở lại bằng hệ thống.

### 5.2. Ô nhập URL

Selector DOM đã thử thành công:

```css
input[role="combobox"][aria-label^="Kiểm tra mọi URL trong"]:not([disabled])
```

Không dùng class CSS của Google vì class có thể thay đổi.

Luồng:

```text
Tìm input đang hoạt động
→ focus
→ Input.insertText(URL)
→ gửi phím Enter bằng CDP
```

### 5.3. Nút yêu cầu lập chỉ mục

DOM tìm phần tử có:

```text
role="button"
```

và nội dung hoặc `aria-label` bắt đầu bằng:

```text
Yêu cầu lập chỉ mục
```

Sau khi tìm thấy, gọi `click()` và dừng chờ kết quả.

### 5.4. Nhận diện kết quả

File:

```text
_he_thong\cdp_watch_indexing_result.py
```

Theo dõi các thành phần DOM:

```css
[role="dialog"]
[role="alertdialog"]
[aria-live="assertive"]
[aria-live="polite"]
```

Sau đó đọc toàn bộ chữ bên trong.

Quy tắc hiện tại:

```text
Có “Đã yêu cầu lập chỉ mục”
hoặc “URL đã được thêm vào hàng đợi ưu tiên”
→ SUCCESS

Có “vượt quá hạn mức”
hoặc “quota exceeded”
→ QUOTA

Nội dung khác
→ UNKNOWN/ERROR

Không có kết quả sau thời gian chờ
→ TIMEOUT
```

Trường hợp `QUOTA` chưa được quan sát thực tế. Không được tự suy đoán câu chữ khác là quota.

## 6. Các thử nghiệm đã xác nhận

### 6.1. Điền URL và Enter bằng DOM

URL:

```text
https://vanhoavadulich.com/tieu-chi-chon-xe-day-phuc-vu-hoi-nghi-va-su-kien-9-139.html
```

Profile:

```text
submit_03 — giavan1080@gmail.com
```

Kết quả:

- Kết nối được Edge qua CDP.
- Tìm đúng input DOM.
- Điền URL.
- Enter.
- Trang chuyển sang `/search-console/inspect`.

### 6.2. Click yêu cầu lập chỉ mục và đọc kết quả

URL:

```text
https://xuctiendoanhnghiep.com/cach-tim-va-phat-trien-he-thong-dai-ly-phan-phoi-banh-xe-day-1-3639.html
```

Profile:

```text
submit_02 — kimnhung1080@gmail.com
```

Kết quả DOM:

```text
SUCCESS

Đã yêu cầu lập chỉ mục

URL đã được thêm vào hàng đợi ưu tiên thu thập dữ liệu.
```

Đây là chuỗi thành công đã được xác nhận thực tế.

## 7. Nhận diện ảnh

Ảnh mẫu đã lưu:

```text
_he_thong\assets\kiem_tra_url.png
```

Helper:

```text
_he_thong\click_image_and_paste.py
```

Nhận diện ảnh đã thử hoạt động, nhưng hiện không phải phương pháp chính. DOM/CDP chính xác và ổn định hơn.

Giữ helper ảnh làm phương án dự phòng; không tự động dùng fallback nếu chưa có quy tắc rõ ràng để tránh click sai.

## 8. Phần chưa hoàn thiện

- Chưa có hàng đợi nhiều URL.
- Chưa có file lịch sử URL lâu dài.
- Chưa nối `cdp_watch_indexing_result.py` vào luồng `SUBMIT` của GUI.
- GUI chưa tự hiển thị `SUCCESS`, `QUOTA`, `ERROR` sau khi click.
- Chưa có quy tắc dừng toàn bộ URL của Gmail khi gặp quota/lỗi.
- Chưa có cơ chế chuyển URL sang Gmail/profile thứ hai.
- Chưa đọc URL từ Excel.
- Chưa ghi trạng thái trở lại Excel.
- Chưa chạy theo giờ.
- Chưa có bộ điều phối chạy nhiều profile song song.
- Chưa có khóa chống hai profile submit trùng một URL.
- Chưa kiểm tra thực tế thông báo vượt hạn mức.
- Chưa xác định quota của hai Gmail cùng property là độc lập hay dùng chung.

## 9. Quy tắc hàng đợi dự kiến

### 9.1. Trạng thái URL

```text
PENDING
RUNNING
SUBMITTED
ALREADY_INDEXED
QUOTA
ERROR
PAUSED_BY_GMAIL
UNMAPPED
```

Ý nghĩa:

- `PENDING`: chờ xử lý.
- `RUNNING`: đang xử lý.
- `SUBMITTED`: Google đã nhận yêu cầu lập chỉ mục.
- `ALREADY_INDEXED`: URL đã nằm trên Google, không dùng lượt submit.
- `QUOTA`: URL gặp thông báo vượt hạn mức.
- `ERROR`: lỗi khác.
- `PAUSED_BY_GMAIL`: chưa chạy vì Gmail đã bị dừng.
- `UNMAPPED`: chưa tìm được Gmail/profile.

### 9.2. Trạng thái Gmail

```text
ACTIVE
STOPPED
```

Quy tắc tạm thời:

```text
SUCCESS
→ ghi SUBMITTED
→ chạy URL tiếp theo

QUOTA hoặc lỗi cần dừng
→ ghi trạng thái URL hiện tại
→ Gmail = STOPPED
→ toàn bộ URL PENDING của Gmail = PAUSED_BY_GMAIL
→ không submit thêm bằng Gmail đó trong phiên chạy
```

Ví dụ Gmail quản lý 5 tên miền, có 100 URL:

```text
URL 1–9    SUBMITTED
URL 10     QUOTA hoặc ERROR
URL 11–100 PAUSED_BY_GMAIL
```

Chưa chốt: khi Gmail số nhỏ dừng, các URL còn lại có tự chuyển sang Gmail thứ hai hay phải chờ người dùng quyết định.

## 10. Lọc URL trước khi submit

Lọc hai tầng:

### Tầng 1: Lịch sử nội bộ

```text
URL đã SUBMITTED hoặc ALREADY_INDEXED
→ bỏ qua hoàn toàn
```

### Tầng 2: Trạng thái Google

```text
URL chưa có lịch sử
→ mở Kiểm tra URL

Nếu URL đã nằm trên Google
→ ALREADY_INDEXED
→ không click Yêu cầu lập chỉ mục

Nếu URL chưa nằm trên Google
→ mới click Yêu cầu lập chỉ mục
```

Phải phân biệt:

- `SUBMITTED`: Google đã nhận yêu cầu; chưa chắc đã index.
- `ALREADY_INDEXED`: URL đã nằm trên Google.

Cả hai đều không submit lại.

## 11. Nguồn URL

### Giai đoạn test

Thêm một ô nhiều dòng trong GUI:

```text
DANH SÁCH URL CẦN SUBMIT

https://example.com/bai-1
https://example.com/bai-2

[THÊM VÀO HÀNG ĐỢI]
[BẮT ĐẦU CHẠY]
```

Mỗi URL một dòng.

### Giai đoạn chính thức

Nguồn chính dự kiến là Excel:

```text
Đọc URL từ Excel
→ đưa vào hàng đợi chung
→ chạy bộ máy submit
→ ghi trạng thái trở lại Excel
```

Chỉ thay nguồn đầu vào; không viết lại bộ máy submit.

## 12. Chạy tự động theo giờ

Dự kiến dùng Windows Task Scheduler:

```text
Đến giờ
→ đọc Excel
→ lấy URL mới/chưa xử lý
→ lọc lịch sử
→ chia URL theo Gmail/profile
→ chạy
→ ghi kết quả
→ dừng Gmail gặp giới hạn
→ kết thúc
```

Không submit lại URL đã `SUBMITTED` hoặc `ALREADY_INDEXED`.

## 13. Chạy nhiều profile song song

Có thể chạy một worker cho mỗi profile nhờ cổng CDP riêng:

```text
Profile 01 → queue Gmail 01
Profile 02 → queue Gmail 02
...
```

Nếu một Gmail dừng, các worker khác vẫn chạy.

Yêu cầu kỹ thuật:

- Một bộ điều phối trung tâm.
- Khóa URL chống submit trùng.
- Chỉ bộ điều phối ghi trạng thái để tránh hỏng JSON/Excel.
- Mỗi worker xử lý tuần tự URL của Gmail mình.
- Thử 2–3 profile trước khi tăng lên 10.

## 14. Thứ tự phát triển đề xuất

1. Nối bộ đọc kết quả DOM vào nút `SUBMIT`.
2. Ghi lịch sử một URL sau mỗi lần chạy.
3. Thêm ô dán nhiều URL và hàng đợi trong GUI.
4. Thêm trạng thái Gmail và quy tắc dừng.
5. Lọc `SUBMITTED` và `ALREADY_INDEXED`.
6. Thử nghiệm thực tế trường hợp `QUOTA`.
7. Chốt quy tắc chuyển sang Gmail thứ hai.
8. Đọc/ghi Excel.
9. Chạy tự động theo giờ.
10. Chạy nhiều profile song song.

## 15. Lưu ý khi tiếp tục ở task mới

Khi tạo task mới, cung cấp file này và yêu cầu:

```text
Đọc README_HE_THONG_SUBMIT.md trước.
Chỉ làm trong submit_edge_profiles.
Không suy đoán Gmail/profile.
Không tự động mở rộng chức năng ngoài yêu cầu.
```

Trước khi sửa:

1. Đọc `submit_edge_profiles.json` mới nhất.
2. Kiểm tra profile nào đang chạy.
3. Không đóng Edge profile nếu chưa được phép.
4. Không thử click `Yêu cầu lập chỉ mục` bằng URL thật nếu chưa xác định profile.
5. Thực hiện từng thay đổi nhỏ và báo rõ trạng thái.

## 16. Quy tắc xử lý kết quả đã chốt

Cập nhật ngày 30/07/2026:

```text
“Đã yêu cầu lập chỉ mục”
→ SUBMITTED
→ tiếp tục URL kế tiếp

“Rất tiếc! Đã xảy ra sự cố”
“Chúng tôi gặp sự cố khi gửi yêu cầu lập chỉ mục của bạn. Vui lòng thử lại sau.”
→ ERROR
→ chỉ bỏ qua URL đang lỗi
→ tiếp tục xử lý URL kế tiếp
→ KHÔNG coi là vượt hạn ngạch
→ KHÔNG dừng Gmail/profile

Chỉ khi thông báo có nội dung rõ ràng về vượt giới hạn hoặc vượt hạn ngạch
→ QUOTA
→ dừng các URL còn lại thuộc Gmail/profile đó
```

Không được dùng quy tắc “không thành công thì mặc định là vượt hạn ngạch”.
