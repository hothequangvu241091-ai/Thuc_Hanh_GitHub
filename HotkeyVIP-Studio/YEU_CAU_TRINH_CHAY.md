# Yêu cầu — Trình chạy HotkeyVIP

Cập nhật: 2026-07-29

## Mục tiêu

Tạo một bảng nút cá nhân để người dùng chạy công cụ theo tên dễ nhớ mà không
phải nhớ tên file hoặc sửa đường dẫn trong AHK/Python.

## Trải nghiệm chính

1. Các nhóm hiển thị thành tab ngang giống GUI AHK:
   - bấm tab nào chỉ hiện danh sách của nhóm đó;
   - tên nhóm có thể chủ động đổi;
   - nhóm mới được tạo khi thêm công việc vào tên nhóm mới;
   - số lượng công việc được hiển thị trên tab.
   - có nút `Thêm nhóm` riêng;
   - nhóm trống vẫn được lưu và hiển thị.
2. Màn hình chính của Trình chạy chỉ hiển thị:
   - thứ tự;
   - tên công việc;
   - trạng thái file;
   - nút Chạy;
   - nút Sửa.
3. Không hiện đường dẫn hoặc thao tác xóa trên danh sách chính.
4. Bấm Sửa mới hiện:
   - tên hiển thị;
   - đường dẫn file;
   - mô tả;
   - nhóm;
   - hiện/ẩn CMD;
   - xóa nút.
   - cạnh đường dẫn có nút mở hộp chọn file Windows, không bắt nhập tay.
5. Người dùng có thể thêm nút mới bằng tên và đường dẫn.
6. Người dùng có thể kéo nút lên/xuống để đổi trình tự.
7. Có thể chuyển nhóm trong cửa sổ Sửa.
8. Nhóm là tên tùy chỉnh, ví dụ:
   - Tự động;
   - Cứu hộ thủ công;
   - Công cụ;
   - Sau khi đăng.
9. Thứ tự và nhóm phải được lưu lại sau khi đóng app.
10. Xóa nút chỉ xóa cấu hình GUI, tuyệt đối không xóa file thật.
11. Chỉ cho chạy khi đường dẫn tồn tại và đúng loại được hỗ trợ.

## Chạy chương trình

Phiên bản hiện tại hỗ trợ `.py` và `.pyw`.

- Hiện CMD: chạy Python trong cửa sổ lệnh riêng.
- Ẩn CMD: yêu cầu Windows mở file giống thao tác nhấp đúp, để giữ đúng liên
  kết Python và ngữ cảnh Desktop của người dùng.
- App không sửa nội dung file được chạy.
- App không được tự suy đoán hoặc đổi đường dẫn.

## Hướng mở rộng

Sau khi danh sách và thứ tự ổn định:

- chạy lần lượt cả nhóm;
- dừng khi một bước lỗi;
- mở rộng `.ahk`, `.bat`, `.ps1`, `.exe`, tài liệu, thư mục và URL;
- lưu log tập trung;
- chọn chế độ kiểm tra lỗi;
- khai báo file phụ bắt buộc.

## Nguyên tắc giao diện

- Chữ lớn, tương phản rõ, không gây mỏi mắt.
- Một công việc trên một hàng.
- Tên công việc là thông tin nổi bật nhất.
- Nhóm hiển thị bằng tab ngang, không xếp thành nhiều khối dọc.
- Chi tiết kỹ thuật được giấu trong cửa sổ Sửa.
- Không dùng nhiều thẻ lớn.
- Không đưa chức năng hiếm dùng lên màn hình chính.

## Giới hạn an toàn

- Không đọc nội dung Excel.
- Không quét `07_ket_qua`.
- Không tự sửa, di chuyển hoặc xóa file trong kho HotkeyVIP.
- Việc chạy file chỉ xảy ra khi người dùng bấm nút Chạy.
