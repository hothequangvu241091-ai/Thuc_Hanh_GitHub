# HotkeyVIP Studio

Trung tâm quan sát và tổ chức kho HotkeyVIP.

Tài liệu bàn giao đầy đủ cho task mới: `BAN_GIAO_DU_AN.md`.

## Mở app

Nhấp đúp `MỞ_HOTKEYVIP_STUDIO.bat`. Studio mở trong cửa sổ ứng dụng riêng,
không có tab hay thanh địa chỉ. Nếu cần kiểm tra bằng trình duyệt, dùng
`MỞ_TRONG_TRÌNH_DUYỆT.bat`.

Có thể mở nhanh bằng shortcut `HotkeyVIP Studio` ngoài Desktop. Shortcut dùng
icon chữ H riêng và gọi trực tiếp `studio_window.py`, không hiện cửa sổ CMD.

## Giới hạn an toàn của bản đầu

- Chỉ đọc kho `D:\CodexProjects\Hotkeyvip`.
- Bỏ qua hoàn toàn `07_ket_qua`; không quét file bên trong kho bài viết này.
- Không sửa, đổi tên, sao chép, di chuyển hoặc xóa file.
- Khu Cập nhật chỉ đọc tên và kích thước file được chọn, không tải hoặc lưu nội dung.
- Khu Chuyển máy chỉ tính gói xem trước.
- Khu Dọn kho chỉ phân loại sơ bộ; không có thao tác xóa.

## Chạy với kho khác

Có thể truyền đường dẫn kho làm tham số:

```text
python app.py "E:\HotkeyVIP"
```

App sử dụng thư viện có sẵn của Python, không yêu cầu cài thêm gói.

## Trình chạy

- Tự thêm nút bằng tên hiển thị và đường dẫn file `.py`/`.pyw`.
- Có thể sửa đường dẫn hoặc xóa nút; việc xóa nút không xóa file Python.
- Nút Chạy chỉ hoạt động khi file tồn tại.
- Có thể chọn hiện CMD để quan sát lỗi hoặc chạy ẩn.
- Danh sách nút được lưu trong `launchers.json` cạnh app.

## Log Python tập trung

- Mọi file `.py`/`.pyw` chạy từ Studio đều tạo một log trong thư mục `logs`.
- Log được tách theo từng nút và từng lượt chạy nên nhiều Python chạy đồng thời không bị trộn.
- Mỗi nút giữ tối đa 10 log gần nhất; mỗi log tối đa 5 MB.
- Mỗi dòng chương trình in ra trong log được gắn mốc giờ `[HH:MM:SS]`, giúp theo dõi thời gian giữa các bước lặp.
- File Python vẫn có thể tiếp tục ghi log riêng như trước.
- Nút `Lịch sử` trên mỗi hàng mở danh sách 10 lượt gần nhất và tự hiển thị nội dung log mới nhất.
- Có thể chọn từng lượt, tải lại log đang chạy, sao chép nội dung hoặc mở file log bằng ứng dụng Windows.

## Điều khiển giọng nói

Mục `Điều khiển giọng nói` khởi động `CHAY_THO_MAY.bat` theo đường dẫn lưu trong
`voice_control.json`, chạy Voice Control ở cổng `8766` và hiển thị giao diện
ngay trong Studio. Voice Control không tự chạy khi chỉ mở Studio.
