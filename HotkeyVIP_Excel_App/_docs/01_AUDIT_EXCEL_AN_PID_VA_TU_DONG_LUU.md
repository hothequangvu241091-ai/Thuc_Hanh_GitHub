# AUDIT EXCEL ẨN, PID TREO VÀ CƠ CHẾ TỰ LƯU

Ngày kiểm tra: 18/08/2026

Mục tiêu của tài liệu này:

- Xác định vì sao lâu lâu app báo Excel đang bị mở ẩn / file bị khóa.
- Xác định flow nào đã có cleanup và flow nào còn sót.
- Xác định “Excel gốc” thực tế là file nào.
- Xác định flow nào tự Save.
- Ghi lại các lỗ hổng PID / subprocess có thể làm lần chạy sau bị kẹt.

> Đây là tài liệu audit hiện trạng. Chưa sửa code trong lần audit này.

---

# 1. KẾT LUẬN NGẮN

Có 4 điểm cần ưu tiên.

## P0 — Flow 8 có thể tự giết chính Excel ẩn mà app vừa mở

`flow_host.py` đang xếp `08_dong_bo_url.py` vào nhóm chạy **in-process**.

Khi app chọn một file Excel chưa mở, `flow_host` tạo một Excel ẩn bằng `DispatchEx`, đăng ký PID và truyền workbook trực tiếp qua `APP_WORKBOOK`.

Nhưng đầu `08_dong_bo_url.py` có:

```python
if os.environ.get("HOTKEYVIP_APP_RUN") != "1":
    close_orphan_hidden_excel()
```

Hàm `close_orphan_hidden_excel()` dùng PowerShell để **kill toàn bộ EXCEL.EXE không có cửa sổ**.

Vấn đề: `HOTKEYVIP_APP_RUN=1` hiện chỉ được `flow_host` gán cho các flow chạy **subprocess**. Flow 8 chạy in-process nên biến này không được gán.

Kết quả có thể xảy ra:

```text
App mở Excel nguồn ẩn
→ Flow 8 bắt đầu
→ Flow 8 nghĩ mình đang chạy standalone
→ close_orphan_hidden_excel()
→ kill luôn Excel ẩn do app đang sở hữu
→ COM chết / workbook mất / flow lỗi
```

Đây là ứng viên rất mạnh cho lỗi “lâu lâu Excel đang mở ẩn / Excel chết”.

Tính “lâu lâu” có thể đến từ việc:

- nếu file Excel đã được người dùng mở và có cửa sổ → không bị hàm này kill;
- nếu app tự mở file bằng Excel ẩn → có nguy cơ bị kill.

### Hướng sửa

Một trong hai cách:

```text
A. flow_host đặt HOTKEYVIP_APP_RUN=1 cho cả in-process flow.
```

hoặc an toàn hơn:

```python
if APP_WORKBOOK is None and os.environ.get("HOTKEYVIP_APP_RUN") != "1":
    close_orphan_hidden_excel()
```

Không được kill Excel ẩn kiểu “tất cả process không có cửa sổ” khi app đã truyền `APP_WORKBOOK`.

---

## P0 — Flow 6 và Flow 8 có lệch giữa “file app chọn” và “file mặc định”

### Quy tắc đúng của app

Excel nguồn/gốc của Flow 1-8 phải là:

```text
file đang nằm trong ô “File Excel” của app
```

UI truyền đường dẫn đó cho `flow_host` bằng:

```text
--workbook <đường dẫn file đang chọn>
```

`flow_host` sau đó mở/bám đúng file đó.

### Flow subprocess 3 / 5 / 7

`flow_host` có gán:

```text
HOTKEYVIP_SELECTED_EXCEL=<file app chọn>
HOTKEYVIP_APP_RUN=1
```

nên các flow này nhận đúng đường dẫn.

### Flow in-process 1 / 2 / 4 / 6 / 8

`flow_host` truyền `APP_WORKBOOK` trực tiếp, nhưng hiện **không gán hai biến môi trường trên trước khi import/chạy flow**.

Flow 1, 2, 4 chủ yếu dùng thẳng `APP_WORKBOOK`, nên không bị lệch đường dẫn.

Nhưng Flow 6 có:

```python
EXCEL_PATH = os.path.abspath(
    os.environ.get("HOTKEYVIP_SELECTED_EXCEL", str(EXCEL_FILE))
)
```

sau đó khi có `APP_WORKBOOK` lại kiểm tra:

```text
APP_WORKBOOK.FullName phải bằng EXCEL_PATH
```

Nếu app đang chọn file khác `EXCEL_FILE` mặc định thì Flow 6 có thể báo “App truyền sai workbook”.

Flow 8 tương tự:

```python
SOURCE_FILE = Path(
    os.environ.get(
        "HOTKEYVIP_SELECTED_EXCEL",
        "D:\\CodexProjects\\Hotkeyvip\\04_excel\\hotkeyvip_test.xlsm",
    )
)
```

và `get_source_workbook()` cũng bắt `APP_WORKBOOK.FullName == SOURCE_FILE`.

### Hướng sửa

Ngay khi `flow_host.run_flow()` đã nhận `workbook_path`, cần đảm bảo toàn bộ flow nhìn thấy cùng một nguồn sự thật:

```text
HOTKEYVIP_SELECTED_EXCEL = workbook_path
HOTKEYVIP_APP_RUN = 1
```

cho **cả in-process và subprocess**.

Sau đó:

```text
APP_WORKBOOK
HOTKEYVIP_SELECTED_EXCEL
--workbook
```

đều phải chỉ cùng một file.

---

## P1 — Flow 5 còn khả năng để lại Python worker PID

Flow 5 chạy multi-process.

Khi kết thúc, hàm `stop_workers()` hiện làm:

```python
for process in workers.values():
    process.join(timeout=8)
```

nhưng **không có bước `terminate()` / `kill()` nếu worker vẫn còn sống sau 8 giây**.

Flow 7 đã làm tốt hơn:

```text
join timeout 8 giây
→ nếu còn sống: terminate()
→ join thêm 3 giây
```

Flow 5 chưa có lớp bảo hiểm tương đương.

Hậu quả có thể xảy ra:

```text
Worker CMS kẹt Selenium / Word / network
→ parent Flow 5 chờ hoặc thoát không sạch
→ python worker còn PID
→ flow_host vẫn chưa hoàn tất
→ Excel do flow_host mở vẫn còn khóa
→ lần chạy sau thấy Excel ẩn / file locked
```

### Hướng sửa

Cho Flow 5 dùng đúng chiến lược Flow 7:

```text
STOP sentinel
→ join 8s
→ terminate worker còn sống
→ join 3s
→ nếu vẫn sống: kill
```

---

## P1 — App chỉ theo dõi PID Excel, chưa theo dõi PID flow/python con

`flow_host.py` đã có registry:

```text
_runtime/app_owned_excel_processes.json
```

Registry lưu:

```text
PID Excel + create_time
```

Điểm tốt:

- tránh kill nhầm PID được hệ điều hành tái sử dụng;
- chỉ dọn Excel do app từng đăng ký;
- chỉ kill nếu Excel đó không có cửa sổ.

Nhưng registry hiện chỉ theo dõi **EXCEL.EXE**.

App không có persistent registry cho:

```text
flow_host python.exe
Flow 3 python.exe
Flow 5 python.exe
Flow 5 worker python.exe
Flow 7 python.exe / worker
```

UI chỉ giữ `self._flow_process` trong RAM.

Khi người dùng đóng app bình thường, UI không cho đóng nếu flow còn chạy. Đây là bảo hiểm tốt.

Nhưng nếu:

```text
Task Manager kill app
Windows restart/crash
pythonw app chết bất thường
IDE/Codex kill process cha
```

thì các child process có thể còn chạy mà app mới không biết PID của chúng.

Lần chạy kế tiếp chỉ có cơ chế dọn Excel đã tracking, chưa có cơ chế dọn “flow session” cũ.

### Hướng sửa

Tạo registry phiên chạy, ví dụ:

```text
_runtime/app_flow_processes.json
```

lưu:

```text
run_id
flow_key
flow_host_pid
child_pid(s)
started_at
workbook_path
```

Khi mở app / trước khi chạy flow mới:

```text
kiểm tra PID cũ
→ đối chiếu create_time / command line / project path
→ nếu đúng process HotkeyVIP cũ và không còn phiên UI sở hữu
→ terminate cây process
→ sau đó mới dọn Excel PID cũ
```

Không kill chung tất cả `python.exe`.

---

# 2. “CMD ẨN” THỰC TẾ LÀ GÌ?

`CHAY_APP.cmd` chỉ dùng `start` để mở `pyw.exe/pythonw.exe`, sau đó CMD thoát ngay.

Trong UI, flow được chạy bằng `subprocess.Popen(... CREATE_NO_WINDOW ...)`.

Trong `flow_host`, Flow 3/5/7 cũng chạy bằng `CREATE_NO_WINDOW`.

Vì vậy phần lớn trường hợp người dùng gọi là “CMD ẩn” thực tế có thể là:

```text
pythonw.exe của app
python.exe của flow_host
python.exe của Flow 3 / 5 / 7
python.exe worker của Flow 5 / 7
```

chứ không nhất thiết là `cmd.exe`.

Đây là lý do chỉ tìm/kill `cmd.exe` sẽ không giải quyết gốc lỗi.

---

# 3. FLOW_HOST HIỆN ĐÃ XỬ LÝ EXCEL ẨN ĐẾN ĐÂU?

## Đã có

Khi app cần tự mở workbook:

```text
DispatchEx("Excel.Application")
→ Visible=False
→ đăng ký PID + create_time
→ Open đúng workbook path
```

Khi flow xong và Excel thuộc app:

```text
workbook.Close(...)
→ excel.Quit()
→ xóa PID khỏi registry
```

Khi file bị khóa ở lần chạy sau:

```text
_cleanup_tracked_hidden_excel()
```

chỉ dọn các Excel:

- có trong registry;
- đúng PID + create_time;
- đúng process EXCEL.EXE;
- không có cửa sổ hiển thị.

Đây là cơ chế tốt và nên giữ.

## Còn thiếu

### Trường hợp Excel ẩn từ phiên bản cũ không có registry

Nếu file bị lock và không có Excel visible, `flow_host` hiện báo:

```text
File đang bị một Excel ẩn từ phiên bản cũ giữ khóa...
```

và yêu cầu người dùng tự đóng một lần.

Tức là chưa có cơ chế tự xác minh ownership cho Excel cũ không nằm trong registry.

Không nên sửa bằng cách kill tất cả Excel ẩn, vì có thể kill Excel của flow khác / app khác.

### Trường hợp Excel ẩn do Flow 8 tạo để xử lý file đích

Flow 8 còn tạo thêm:

```python
target_excel = win32.DispatchEx("Excel.Application")
```

để xử lý các file domain trên G:\.

Nó có `target_excel.Quit()` trong `finally`, nên exception Python bình thường sẽ dọn được.

Nhưng Excel này **không được đăng ký vào registry trung tâm của flow_host**.

Nếu process bị force-kill / máy crash đúng lúc đang xử lý file đích, instance này có thể trở thành Excel ẩn không được tracking.

Đây là một nguồn Excel mồ côi tiềm năng khác.

---

# 4. KIỂM TRA TỪNG FLOW

| Flow | Cách lấy Excel khi chạy từ app | Có tự mở Excel riêng? | Có Save? | Cleanup đáng chú ý |
|---|---|---:|---:|---|
| 1 Nhập KE_HOACH | `APP_WORKBOOK` | Không cho workbook chính | Có | Host quản lý Excel chính |
| 2 Chuẩn bị VIET_BAI | `APP_WORKBOOK` | Không | Có | Host quản lý |
| 3 Viết bài + ảnh | subprocess, tìm workbook theo `HOTKEYVIP_SELECTED_EXCEL` | Không mở workbook mới theo logic chính | Có, rất nhiều lần | Worker thread `join()` không timeout |
| 4 Chuẩn bị DANG_BAI | `APP_WORKBOOK` | Không | Có | Host quản lý |
| 5 Đăng CMS | subprocess, bám `EXCEL_PATH` theo env | Không chủ động tạo Excel mới; bám file host mở | Có, nhiều lần | Worker process còn thiếu terminate bảo hiểm |
| 6 Lấy URL CMS | `APP_WORKBOOK` | Không | Có | Có lỗi lệch `EXCEL_PATH` nếu env chưa được set |
| 7 Bài liên quan | subprocess, bám path; nếu standalone có thể tự mở Excel ẩn | Có fallback standalone | Có mỗi bài thành công | Có terminate worker sau timeout, khá tốt |
| 8 Đồng bộ URL | `APP_WORKBOOK` cho nguồn + Excel riêng cho file đích | **Có target Excel ẩn riêng** | Có | Có lỗi `HOTKEYVIP_APP_RUN` + target Excel chưa registry |

---

# 5. FLOW 3 CÓ THỂ GIỮ EXCEL LÂU NẾU WORKER TREO

Flow 3 dùng thread:

```text
ExcelWriter
WordWorker
Worker-1
Worker-2
```

Cuối flow hiện có các đoạn kiểu:

```python
for worker in workers:
    worker.join()

WORD_QUEUE.join()
word_worker.join()

RESULT_QUEUE.join()
writer.join()
```

Không có timeout ở các `join()` chính.

Nếu một worker Selenium / Gemini / Word bị mắc ở trạng thái không thoát thì Flow 3 có thể không return.

Khi Flow 3 không return:

```text
flow_host vẫn đợi process
→ Excel host vẫn mở
→ workbook vẫn lock
```

Đây không nhất thiết tạo Excel mồ côi; nó có thể là Excel **vẫn đang thuộc một flow bị treo**.

### Hướng sửa sau này

- Có watchdog tổng cho Flow 3.
- Mỗi worker có heartbeat.
- Join có timeout.
- Nếu worker mất heartbeat quá lâu: đóng Edge/Word của worker và kết thúc worker sạch.

---

# 6. APP CÓ TỰ ĐỘNG SAVE EXCEL KHÔNG?

**Có. Rất nhiều lớp đang Save.**

## Flow 1

`target_book.Save()` sau khi nhập KE_HOACH.

## Flow 2

`workbook.Save()` sau khi cập nhật VIET_BAI.

## Flow 3

Nhiều hàm gọi `wb.save()` ngay sau khi ghi trạng thái, Word, lỗi, ảnh...

Nghĩa là Flow 3 đang checkpoint liên tục.

## Flow 4

`workbook.Save()` sau khi chuẩn bị DANG_BAI.

## Flow 5

Nhiều chỗ `workbook.Save()` khi đổi trạng thái, ghi lỗi, ID, thời gian đăng...

## Flow 6

`wb.Save()` cuối bước lấy URL.

## Flow 7

`workbook.Save()` sau từng bài viết liên quan lưu thành công.

## Flow 8

Có nhiều Save:

- Save workbook nguồn trước đồng bộ trong chế độ app;
- SaveCopyAs backup;
- Save sau khi thêm cột trạng thái;
- Save cuối sau khi ghi `Đã chuyển`.

## flow_host còn Save thêm một lần

Sau flow in-process:

```python
return_code = _run_flow_in_process(...)
workbook.Save()
```

Sau flow subprocess:

```python
return_code = _run_flow_subprocess(...)
workbook.Save()
```

Tức là nhiều flow đang Save 2 lớp:

```text
flow tự Save
+
flow_host Save lần cuối
```

---

# 7. ĐIỂM QUAN TRỌNG: HIỆN CÓ THỂ SAVE CẢ KHI FLOW LỖI

Đây là hành vi cần hiểu rõ.

## Flow subprocess 3 / 5 / 7

`_run_flow_subprocess()` trả về `return_code`.

Dù return code khác 0, `flow_host` vẫn đi tiếp tới:

```python
workbook.Save()
```

Nghĩa là flow có thể lỗi nhưng các thay đổi đã ghi trước đó vẫn được Save.

Điều này có thể là chủ ý để giữ checkpoint, nhưng cần ghi rõ.

## Flow in-process 1 / 2 / 4 / 6 / 8

Nếu flow ném exception trước khi dòng Save cuối của host chạy thì:

```text
workbook_saved = False
```

Trong `finally`, host đang gọi:

```python
workbook.Close(SaveChanges=not workbook_saved)
```

Khi `workbook_saved=False` thì thực tế thành:

```python
workbook.Close(SaveChanges=True)
```

Tức là nếu Excel do app sở hữu, **exception giữa flow vẫn có thể làm host Save thay đổi dở dang khi đóng workbook**.

### Đây không hẳn là bug

Có hai triết lý:

1. **Checkpoint-first**: lỗi vẫn giữ mọi thay đổi đã hoàn thành trước đó.
2. **Transaction-first**: flow lỗi thì không Save phần chưa xác nhận.

Hiện app nghiêng mạnh về **checkpoint-first**.

Cần quyết định rõ từng flow nào được phép checkpoint và flow nào phải all-or-nothing.

---

# 8. “EXCEL GỐC” LÀ FILE NÀO?

Trong kiến trúc app hiện tại, định nghĩa nên là:

```text
EXCEL GỐC / EXCEL NGUỒN CHÍNH
= file người dùng đang chọn trong ô File Excel của app
```

Đường dẫn này đi theo chuỗi:

```text
ui.selected_path
→ --workbook
→ flow_host(workbook_path)
→ workbook được mở/bám đúng FullName
```

Sau đó:

### Flow in-process

Nên dùng:

```text
APP_WORKBOOK
```

làm nguồn sự thật chính.

### Flow subprocess

Nên dùng:

```text
HOTKEYVIP_SELECTED_EXCEL
```

và bắt buộc so đường dẫn với workbook thực tế trước khi ghi.

### Không nên có tình trạng

```text
APP_WORKBOOK = file A
nhưng EXCEL_PATH/SOURCE_FILE fallback = file B hardcode
```

Flow 6 và 8 hiện đang có nguy cơ này do env chỉ được set cho subprocess.

---

# 9. CÁC BƯỚC NGOÀI FLOW 1-8

## Phân tích

`analysis.py` / `excel_io.py` đọc Open XML, không cần mở Excel COM.

Không phải nguồn tạo Excel ẩn.

## Xuất kết quả

`report_export.py` gọi `export_with_excel.ps1` bằng PowerShell ẩn.

PowerShell script:

```text
New-Object -ComObject Excel.Application
→ Visible=False
→ mở bản working copy
→ Save
→ finally Close workbook
→ Excel.Quit
→ FinalReleaseComObject
→ GC
```

Cleanup hiện khá đầy đủ.

## Khôi phục DANG_BAI

`recover_dang_bai_with_excel.ps1` cũng có `finally`:

```text
Close workbook
→ Excel.Quit
→ release COM
→ GC
```

Cleanup cũng khá đầy đủ.

Rủi ro chủ yếu chỉ còn nếu PowerShell bị force-kill / máy chết giữa chừng.

---

# 10. THỨ TỰ NÊN SỬA

## Sửa 1 — bắt buộc

Centralize runtime context trong `flow_host`:

```python
HOTKEYVIP_SELECTED_EXCEL = workbook_path
HOTKEYVIP_APP_RUN = 1
```

phải tồn tại trước khi chạy **mọi flow**, kể cả in-process.

Mục tiêu:

```text
--workbook
APP_WORKBOOK.FullName
HOTKEYVIP_SELECTED_EXCEL
```

luôn trùng nhau.

## Sửa 2 — bắt buộc

Flow 8 không được chạy `close_orphan_hidden_excel()` nếu `APP_WORKBOOK` đã được app truyền vào.

## Sửa 3 — rất nên làm

Flow 5 `stop_workers()`:

```text
join
→ terminate
→ join
→ kill nếu cần
```

như Flow 7.

## Sửa 4 — nên làm

Tạo registry PID cho flow/python process tree, không chỉ registry Excel.

## Sửa 5 — nên làm

Flow 8 đăng ký `target_excel` vào cơ chế ownership chung hoặc có registry Excel phụ.

## Sửa 6 — cân nhắc

Xác định chính sách Save khi flow lỗi.

Ví dụ:

```text
Flow 3 / 5: checkpoint được phép
Flow 1 / 2 / 4: có thể muốn transaction rõ hơn
Flow 8: backup + checkpoint rõ ràng
```

Không nên để `Close(SaveChanges=True)` xảy ra chỉ vì `workbook_saved=False` mà không phân biệt flow lỗi hay flow thành công.

## Sửa 7 — sau cùng

Watchdog cho Flow 3 để worker/thread không giữ flow_host + Excel vô thời hạn.

---

# 11. CHECKLIST KHI GẶP LỖI “EXCEL ĐANG MỞ ẨN”

Đừng chỉ kill Excel ngay.

Kiểm tra theo thứ tự:

```text
1. App còn flow đang chạy không?
2. Có python.exe / pythonw.exe HotkeyVIP cũ còn sống không?
3. _runtime/app_owned_excel_processes.json có PID nào không?
4. PID Excel đó còn tồn tại không?
5. Excel đó có cửa sổ không?
6. File bị lock có đúng file app đang chọn không?
7. Flow vừa chạy là Flow 3, 5 hay 8?
```

Nếu Flow 8 vừa chạy và file trước đó do app mở ẩn, ưu tiên nghi lỗi P0 ở mục 1.

Nếu Flow 5 vừa chạy và có Python worker cũ, ưu tiên nghi `stop_workers()` chưa terminate.

Nếu Flow 3 vẫn có process sống, có thể worker thread đang treo và Excel chưa phải “mồ côi” — nó vẫn đang bị flow giữ.

---

# 12. MÔ HÌNH MỤC TIÊU SAU KHI SỬA

```text
UI chọn workbook A
        ↓
flow_host tạo RUN_ID
        ↓
set HOTKEYVIP_SELECTED_EXCEL=A
set HOTKEYVIP_APP_RUN=1
        ↓
đăng ký PID flow_host
        ↓
Nếu cần: mở Excel A ẩn
→ đăng ký PID Excel A
        ↓
chạy Flow
→ đăng ký child PID
→ flow dùng đúng workbook A
        ↓
Flow kết thúc / lỗi / user stop
        ↓
stop child process tree có kiểm soát
        ↓
Save theo policy của flow
        ↓
Close workbook
        ↓
Quit Excel app-owned
        ↓
xóa PID registry
```

Lần chạy mới:

```text
kiểm tra RUN_ID/PID cũ
→ dọn child HotkeyVIP mồ côi
→ dọn Excel app-owned mồ côi
→ xác nhận file hết lock
→ mới chạy flow mới
```

Đây là hướng bền vững hơn việc mỗi flow tự viết một cách kill Excel riêng.
