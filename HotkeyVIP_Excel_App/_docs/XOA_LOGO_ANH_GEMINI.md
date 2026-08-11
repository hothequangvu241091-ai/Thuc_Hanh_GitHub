# Xóa logo ảnh Gemini trong Flow 3

Flow **3. Viết bài + tạo ảnh** tự xử lý logo ở vùng cố định gần góc phải-dưới
của ảnh Gemini 4:3. Luồng xử lý là: lưu ảnh gốc → xóa logo bằng LaMa GPU →
resize cạnh dài tối đa 800 px → ghi đường dẫn ảnh vào Excel.

## Thành phần đi kèm app

- Code: `app_flows/logo_cleanup.py`.
- Model: `_runtime/logo_cleanup/big-lama.pt` (khoảng 206 MB).
- Môi trường Python chạy Flow 3 cần: `torch` CUDA, `opencv-python`, `numpy`,
  `Pillow`, `simple_lama_inpainting`, cùng driver NVIDIA/CUDA hoạt động.

Flow 3 kiểm tra các thành phần này một lần khi khởi động. Model chỉ được nạp
lên GPU khi ảnh Gemini đầu tiên cần xử lý, sau đó được dùng chung cho cả phiên
chạy. Hai worker dùng chung một khóa GPU để không nạp/chạy hai model đồng thời.

## Khi không xử lý được logo

Xóa logo là bước không chặn Flow 3. Nếu model, GPU hoặc AI gặp lỗi, app vẫn giữ
ảnh Gemini, resize về 800 px và hoàn thành bài như bình thường.

- Cột `U/W` luôn chỉ chứa đường dẫn file hợp lệ.
- Cột trạng thái ảnh `T/V` sẽ ghi `CHƯA XÓA LOGO`.
- Cột `AD` lưu `LOGO_REMOVE_FAILED` cùng mô tả lỗi ngắn để có thể lọc trong
  Excel.

Tool dùng tọa độ theo tỷ lệ ảnh nên không phụ thuộc độ phân giải, nhưng yêu cầu
logo vẫn ở vị trí tương ứng sát góc phải-dưới của ảnh Gemini. Mask hình thoi
hiện tại che logo lấp lánh có tâm quanh 95% chiều rộng và 93,4% chiều cao,
được đo trực tiếp trên ảnh Gemini thật. Nếu Gemini đổi kiểu hoặc vị trí logo,
cần hiệu chỉnh `LOGO_POLY` trong `logo_cleanup.py`.

Khi sao chép app sang máy khác phải sao chép nguyên thư mục dự án, đặc biệt
không xóa `_runtime/logo_cleanup/big-lama.pt`. Nếu môi trường Python/GPU thiếu
thành phần, Flow 3 không dừng: ảnh được giữ nguyên và Excel ghi cảnh báo như trên.
