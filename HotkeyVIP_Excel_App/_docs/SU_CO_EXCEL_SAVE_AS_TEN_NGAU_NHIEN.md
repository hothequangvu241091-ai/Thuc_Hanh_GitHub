# Sự cố Excel Save As với tên ngẫu nhiên

## Lần ghi nhận

- Ngày: 11/08/2026.
- Flow đang chạy: Flow 3 - Viết bài + tạo ảnh.
- Workbook chính: `D:\CodexProjects\Hotkeyvip\04_excel\hotkeyvip_test.xlsm`.
- Excel bất ngờ hiện hộp thoại `Save As` với tên ngẫu nhiên `040C7000.xlsm`.
- File `040C7000.xlsm` đã được tạo như một bản sao gần bằng dung lượng workbook chính.
- Người dùng cho biết hiện tượng tương tự đã xảy ra một lần vào ngày trước đó và có nguy cơ mất dữ liệu/chạy lại dữ liệu.

## Phân loại

Đây là lỗi quản lý phiên Excel/workbook, tách biệt với:

- lỗi Gemini không tạo được ảnh;
- cơ chế ưu tiên bài lỗi của Flow 3;
- checkpoint Word, brief và ảnh.

Flow 3 không có lệnh Excel `SaveAs` để tạo file `.xlsm`; các lệnh `SaveAs2` trong Flow 3 chỉ dùng cho Microsoft Word lưu `.docx`.

## Dấu hiệu đã quan sát

- Có nhiều tiến trình Excel cùng tồn tại: một Excel từ phiên trước và một Excel ẩn do app mở cho flow mới.
- Excel bị chặn bởi hộp thoại Save As nên các lệnh COM kiểm tra/lưu có thể đứng chờ.
- Tên file Save As là chuỗi ngẫu nhiên tám ký tự, không phải tên workbook người dùng đã chọn.
- Không được coi file tên ngẫu nhiên là workbook chính.

## Nguyên nhân có khả năng cao

Excel ẩn của flow trước không thoát hoàn toàn hoặc bị kẹt bởi hộp thoại, trong khi flow mới tiếp tục mở/bám workbook. Hai phiên thao tác liên quan cùng workbook có thể khiến Excel chuyển sang trạng thái lưu bản sao thay vì lưu đúng workbook chính.

Chưa khẳng định tuyệt đối nguyên nhân cuối cùng nằm trong Excel, macro hay add-in vì các tiến trình đã phải dừng cưỡng bức trước khi lấy được stack/log đầy đủ.

## Thay đổi đã thực hiện

Đã sửa `excel_audit_app/flow_host.py` ngày 11/08/2026:

1. Thêm lock riêng theo đường dẫn workbook, không cho hai `flow_host` cùng thao tác một workbook.
2. Trước khi mở flow mới, tự dọn các Excel ẩn còn sót nhưng chỉ với PID đã được app đăng ký sở hữu.
3. Khi gọi `Excel.Quit()`, chỉ xóa PID khỏi sổ theo dõi sau khi tiến trình Excel thực sự thoát.
4. Nếu Excel chưa thoát vì bị kẹt, giữ lại dấu vết để lần chạy sau nhận diện và dọn đúng tiến trình.
5. Đã kiểm tra `py_compile` và thử cơ chế tạo/gỡ workbook lock thành công.

## Nếu sự cố tái diễn

1. Không bấm `Save` vào tên ngẫu nhiên và không ghi đè workbook chính.
2. Không khởi chạy thêm flow thứ hai.
3. Chụp hộp thoại Save As, cửa sổ monitor flow và thời điểm xảy ra.
4. Ghi lại PID của `EXCEL.EXE`, `python.exe` và lệnh chạy trước khi đóng tiến trình.
5. Dừng đúng Worker/flow đang chạy; không đóng Excel thủ công của người dùng nếu chưa xác định PID.
6. Kiểm tra workbook chính và file tên ngẫu nhiên theo kích thước, thời gian sửa và dữ liệu `VIET_BAI` trước khi quyết định giữ file nào.

## Giới hạn của bản sửa

Bản sửa ngăn xung đột giữa các flow do app điều phối và tự dọn Excel ẩn do app sở hữu. Nó không thể bảo đảm hộp Save As không xuất hiện nếu nguyên nhân đến từ macro, add-in Excel hoặc người dùng mở cùng workbook bằng một phiên Excel khác. Nếu tái diễn, cần giữ tiến trình sống đủ lâu để kiểm tra chính xác cửa sổ Save As thuộc PID nào.

## Bảo hiểm Excel Writer bổ sung ngày 11/08/2026

- Giới hạn tổng thời gian retry một lệnh Excel còn 30 giây; không retry vô hạn.
- Mỗi vòng retry kiểm tra lệnh dừng.
- Quá 30 giây sẽ đặt trạng thái lỗi nghiêm trọng, dừng giao dòng mới và đánh thức Worker đang tạm dừng để thoát tại checkpoint.
- Hàng đợi đang chờ được giải phóng để tiến trình không mắc ở `Queue.join()`.
- Trước khi ghi ô và trước khi Save, Excel Writer xác nhận `workbook.FullName` vẫn đúng `HOTKEYVIP_SELECTED_EXCEL`.
- Nếu người dùng hoặc Excel Save As sang tên ngẫu nhiên, flow dừng ngay và không tiếp tục ghi vào bản sao.
- Đã kiểm tra cú pháp, kiểm thử timeout giả lập và kiểm thử chặn đường dẫn `040C7000.xlsm` thành công.
