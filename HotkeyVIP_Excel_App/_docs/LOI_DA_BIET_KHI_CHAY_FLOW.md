# Lỗi đã biết khi chạy flow

## Word báo `File In Use` / `.docx is locked for editing by 'Admin'`

### Bối cảnh đã gặp

- Flow: `3. Viết bài + tạo ảnh` (`03_viet_bai_tao_anh.py`).
- Word hiện hộp thoại `File In Use` đối với một file như `11222.docx`.
- Các lựa chọn gồm mở Read Only, tạo local copy hoặc chờ nhận thông báo.
- Trong lần ghi nhận ngày 2026-08-10, flow vẫn tiếp tục; bảng trạng thái sau đó báo:
  - `WORD + VBA`: đã hoàn thành dòng hiện tại.
  - `EXCEL WRITER`: sống, ghi bình thường, hàng chờ 0.

### Ý nghĩa

File `.docx` đang bị một tiến trình Word giữ khóa chỉnh sửa tại thời điểm Word Worker hoặc VBA muốn mở nó. Đây là xung đột khóa Word, không phải lỗi Excel.

Nguyên nhân có thể gồm:

- Người dùng đang mở file kết quả bằng Word.
- Word ẩn của lượt trước chưa đóng kịp.
- Word Worker vừa tạo/lưu file nhưng chưa giải phóng document trước khi VBA mở lại.
- File khóa tạm `~$*.docx` còn tồn tại sau lần Word đóng không sạch.

### Xử lý ngay khi đang chạy

1. Bấm `Cancel`; không chọn Read Only và không tạo local copy.
2. Không mở thủ công file Word trong thư mục kết quả khi flow đang chạy.
3. Quan sát bảng trạng thái:
   - Nếu `WORD + VBA` tiếp tục hoàn thành và `EXCEL WRITER` bình thường thì để flow chạy.
   - Nếu Word đứng lâu hoặc hàng chờ tăng liên tục thì mới tạm dừng để kiểm tra tiến trình Word treo.

### Khi nào cần sửa code

Chỉ coi là lỗi cần sửa nếu hộp thoại lặp lại hoặc làm Word Worker đứng. Hướng sửa cần kiểm tra:

- Đóng document và giải phóng COM hoàn toàn trước bước VBA.
- Chờ file hết khóa trong thời gian ngắn trước khi mở lại.
- Retry có giới hạn khi file còn khóa.
- Tắt cảnh báo Word để hộp thoại không chặn worker.
- Duy trì duy nhất một luồng Word/VBA thao tác file tại một thời điểm.

### Từ khóa tra cứu

`11222.docx`, `File In Use`, `locked for editing by Admin`, `Word Worker`, `VBA`, `~$docx`
