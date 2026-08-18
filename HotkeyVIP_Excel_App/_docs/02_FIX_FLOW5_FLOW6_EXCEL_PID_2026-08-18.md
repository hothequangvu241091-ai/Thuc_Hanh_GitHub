# FIX FLOW 5 + FLOW 6 — EXCEL / PID

Ngày sửa: 18/08/2026

Commit code: `7c909ab0752831f429c869f2cb030b7f9dec2cdb`

## Flow 5 — Đăng bài CMS

Vấn đề:

- Flow 5 dùng nhiều Python worker bằng `multiprocessing`.
- `stop_workers()` chỉ gửi STOP rồi `join(timeout=8)`.
- Nếu worker còn sống sau 8 giây thì process Flow 5 có thể bị giữ lại lúc Python shutdown.
- Khi Flow 5 chưa thoát, `flow_host` vẫn giữ workbook/Excel mà app đang dùng.

Cách sửa tại `excel_audit_app/flow_host.py`:

- Khi Flow 5 đã chạy xong phần điều phối và in dòng `Thời gian kết thúc: ...`, `flow_host` kiểm tra cây process con của đúng PID Flow 5.
- Chỉ nhắm tới Python child process nằm trong cây process Flow 5.
- Gửi `terminate()`.
- Chờ 3 giây.
- Process nào vẫn còn sống thì `kill()` và chờ thêm.
- Không quét/kill Python khác trên máy.

Mục tiêu:

```text
Flow 5 hoàn tất / lỗi
→ coordinator chạy finally
→ worker bình thường tự thoát
→ worker nào còn kẹt được flow_host dọn
→ process Flow 5 thoát
→ flow_host Save/Close/Quit Excel theo cơ chế chung
```

## Flow 6 — Lấy URL CMS

Vấn đề:

- Flow 6 chạy in-process và được truyền `APP_WORKBOOK`.
- Nhưng module Flow 6 còn tính `EXCEL_PATH` từ biến `HOTKEYVIP_SELECTED_EXCEL` ngay lúc import.
- Trước đây `flow_host` chỉ set biến này cho flow subprocess.
- Vì vậy có khả năng:

```text
APP_WORKBOOK = file app đang chọn
EXCEL_PATH = EXCEL_FILE mặc định
```

- Flow 6 sau đó tự kiểm tra hai đường dẫn và có thể báo `App truyền sai workbook`.

Cách sửa trung tâm tại `flow_host.py`:

Trước khi mở/chạy bất kỳ flow nào, luôn set:

```text
HOTKEYVIP_SELECTED_EXCEL = workbook_path của app
HOTKEYVIP_APP_RUN = 1
```

Áp dụng cho cả flow in-process và subprocess.

Sau sửa, ba nguồn phải thống nhất:

```text
--workbook
APP_WORKBOOK.FullName
HOTKEYVIP_SELECTED_EXCEL
```

đều là file người dùng đang chọn trong app.

## Ảnh hưởng phụ có lợi tới Flow 8 — Đồng bộ URL

Không sửa file `08_dong_bo_url.py` trong đợt này.

Tuy nhiên việc `flow_host` luôn set `HOTKEYVIP_APP_RUN=1` cũng làm Flow 8 nhận đúng ngữ cảnh chạy từ app. Vì Flow 8 chỉ gọi `close_orphan_hidden_excel()` khi `HOTKEYVIP_APP_RUN != 1`, nên khi chạy từ app nó sẽ không tự quét/kill Excel ẩn theo nhánh standalone.

Đây là bảo hiểm phụ từ bản sửa Flow 6, không thay đổi nghiệp vụ đồng bộ URL của Flow 8.

## Chưa sửa trong đợt này

- Flow 3 watchdog/thread timeout.
- Registry lâu dài cho toàn bộ PID flow qua các lần app crash.
- Chính sách Save/rollback khi flow lỗi.
- Logic nghiệp vụ các Flow 1/2/3/4/5/6/7/8.
