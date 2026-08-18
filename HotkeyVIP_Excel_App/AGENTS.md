# Quy tắc làm việc với HotkeyVIP Excel App

Áp dụng cho mọi Codex/ChatGPT sửa file trong thư mục này.

## Trước khi sửa

1. Chạy `git status --short` và `git log -1 --oneline --decorate`.
2. Đọc mục phiên bản mới nhất trong `CHANGELOG.md`; không cần đọc toàn bộ lịch sử.
3. Xem thay đổi chưa commit trước khi đụng vào cùng file. Không ghi đè hoặc xóa
   thay đổi của người dùng và không dùng `git reset --hard`.

## Khi sửa

1. Chỉ sửa đúng phạm vi người dùng yêu cầu.
2. Giữ `app_flows` cho các flow chính; tool dùng một lần đặt trong
   `_archive/one_off_tools/<ten>_<YYYYMMDD>` hoặc thư mục phù hợp dưới `_tools`.
3. Nếu thay đổi hành vi app, tăng version app trong `excel_audit_app/analysis.py`
   và `excel_audit_app/__init__.py`.
4. Nếu thay đổi hành vi một flow, tăng engine version của flow và cập nhật mô tả
   tương ứng trong `excel_audit_app/flow_catalog.py` nếu có.

## Trước khi bàn giao

1. Chạy kiểm tra cú pháp/test phù hợp với phạm vi sửa.
2. Ghi một mục ngắn trong `CHANGELOG.md`: ngày, lỗi cũ, thay đổi và cách kiểm tra.
3. Không đưa `_runtime`, `outputs`, log, cache hoặc dữ liệu người dùng vào Git.
4. Tạo commit Git có mô tả rõ sau khi thay đổi đã được kiểm tra. Tạo tag khi có
   số phiên bản mới; không tự di chuyển hoặc ghi đè tag cũ.
5. Báo cho người dùng version/commit/tag mới và các kiểm tra đã chạy.

Chi tiết thao tác xem hoặc phục hồi nằm trong
`_docs/QUY_TRINH_QUAN_LY_PHIEN_BAN.md`.
