# BÀN GIAO DỰ ÁN HOTKEYVIP STUDIO

> Cập nhật gần nhất: 2026-07-29  
> Đọc file này trước khi tiếp tục sửa HotkeyVIP Studio trong một task Codex mới.

## 1. Mục tiêu

HotkeyVIP Studio là bảng điều khiển cá nhân để:

- Không phải nhớ file nằm ở đâu.
- Chạy nhanh Python/AHK và mở Excel, Word, INI.
- Phân nhóm công việc theo flow thực tế.
- Đánh dấu những nút đang cần dùng.
- Tập trung log Python về một chỗ.
- Sau này có thể mở rộng chạy tuần tự, cứu hộ flow, chuyển máy và quản lý đường dẫn gốc.

Không biến Studio thành một hệ thống khổng lồ ngay. Hiện tại chỉ ưu tiên **Trình chạy**.

## 2. Vị trí

- App: `D:\CodexProjects\HotkeyVIP-Studio`
- Kho code được gọi: `D:\CodexProjects\Hotkeyvip`
- Kho bài viết được bảo vệ: `D:\CodexProjects\Hotkeyvip\07_ket_qua`

### Quy tắc đặc biệt

- Không quét hoặc kiểm tra nội dung `07_ket_qua`.
- Không tự xóa, di chuyển hoặc sửa nội dung file trong kho Hotkeyvip.
- Việc xóa nút/nhóm trong Studio chỉ thay đổi cấu hình GUI, không xóa file thật.
- Các chức năng có tác động thật phải có xác nhận rõ ràng.

## 3. Cách mở

- `MỞ_HOTKEYVIP_STUDIO.bat`: mở trong cửa sổ Edge dạng ứng dụng, không có tab hay thanh địa chỉ.
- Shortcut `HotkeyVIP Studio` ngoài Desktop: cách mở được khuyên dùng, có icon chữ H riêng và không hiện CMD.
- `MỞ_TRONG_TRÌNH_DUYỆT.bat`: phương án dự phòng để mở bằng trình duyệt.
- `ĐÓNG_HOTKEYVIP_STUDIO.bat`: tắt máy chủ local ở cổng `8765`.
- Địa chỉ local: `http://127.0.0.1:8765/`

Chế độ cửa sổ riêng chỉ là lớp bọc. Code HTML/CSS/JS/Python vẫn sửa nhanh như trước và chưa đóng gói EXE.

## 4. Kiến trúc

- `app.py`: máy chủ local, API, lưu cấu hình, gọi file, mở thư mục và mở log.
- `launcher_worker.py`: bộ chạy Python trung gian và ghi log tập trung.
- `studio_window.py`: khởi động máy chủ rồi mở Edge ở chế độ app.
- `static/index.html`: cấu trúc giao diện.
- `static/styles.css`: toàn bộ kiểu dáng.
- `static/app.js`: tương tác giao diện, kéo thả, tìm kiếm, sắp xếp và gọi API.
- `launchers.json`: danh sách nút chạy.
- `launcher_groups.json`: thứ tự và tên nhóm.
- `logs/`: log Python do Studio tạo khi chạy thực tế.

Không cần gom tất cả vào một file Python. Khi đóng gói EXE sau này, công cụ đóng gói có thể gom các file này vào một ứng dụng.

## 5. Trạng thái giao diện

Giao diện hiện hiển thị:

- **Trình chạy**.
- **Điều khiển giọng nói**: nhúng giao diện THỢ MÁY Voice Hub vào Studio nhưng vẫn chạy server riêng.

Các mục Tổng quan, Công việc, Tra cứu kho, Cập nhật, Chuyển máy và Dọn kho vẫn còn code nhưng đang bị ẩn. Chúng không được tải dữ liệu hoặc quét kho lúc khởi động. Không xóa vội; sau này có thể bật lại từng mục.

Nút **Quét lại** cũng đang ẩn vì không liên quan trực tiếp tới Trình chạy.

## 6. Chức năng Trình chạy đã hoàn thành

### Nút và nhóm

- Thêm, sửa và xóa nút GUI.
- Thêm, đổi tên và xóa nhóm.
- Khi xóa nhóm, nút được chuyển về `Chưa phân nhóm`; file thật giữ nguyên.
- Thêm nút khi đang đứng ở nhóm nào thì mặc định vào nhóm đó.
- Đổi nhóm nhanh bằng menu trên từng hàng.
- Kéo nút lên/xuống trong nhóm.
- Kéo nút lên tab nhóm khác để chuyển nhóm.
- Kéo tab nhóm trái/phải để đổi độ ưu tiên nhóm.
- Thứ tự được lưu vào JSON.

### Nhóm đặc biệt

- `★ Cần dùng` là nhóm ảo.
- Bấm ngôi sao để một nút xuất hiện thêm trong Cần dùng.
- Nút vẫn giữ nguyên nhóm gốc.
- Trong Cần dùng chỉ sắp xếp lên/xuống, không làm đổi nhóm gốc.

### Số thứ tự

- Ô số bên trái có thể nhập trực tiếp.
- Ví dụ đang số 1, nhập 4 rồi Enter/bấm ra ngoài thì nút chuyển đến vị trí 4.
- Nếu có 11 nút mà nhập 99, nút xuống cuối và tự thành số 11.
- Nhập 0 hoặc số âm thì về vị trí 1.
- Khi đang tìm kiếm xuyên nhiều nhóm, ô thứ tự bị khóa.

### Tìm kiếm

- Tìm theo tên nút, mô tả, nhóm, tên file và đường dẫn.
- Tìm xuyên tất cả nhóm.
- `Ctrl + K`: tập trung vào ô tìm kiếm.
- `Escape`: xóa tìm kiếm và quay về nhóm đang xem.

### Loại file hỗ trợ

- `.py`, `.pyw`: Chạy.
- `.ahk`: Chạy bằng liên kết AutoHotkey của Windows.
- `.ini`: Mở bằng ứng dụng mặc định.
- `.xlsx`, `.xlsm`, `.xls`: Mở bằng Excel.
- `.docx`, `.doc`: Mở bằng Word.

Trong cửa sổ sửa nút:

- `Chọn file…`: mở hộp chọn file Windows.
- `Mở thư mục`: mở File Explorer và chọn sẵn file đang cấu hình.

## 7. Log Python tập trung

Mọi `.py/.pyw` chạy từ Studio đều đi qua `launcher_worker.py`.

Hành vi:

- Thành công hay lỗi đều tạo log.
- Ghi nội dung stdout và stderr/traceback.
- File Python vẫn có thể tiếp tục ghi log riêng như trước.
- Nhiều Python chạy đồng thời tạo các log độc lập, không trộn và không ghi đè.
- Mỗi nút giữ tối đa 10 log gần nhất.
- Mỗi log tối đa 5 MB.
- Khi đạt giới hạn, chương trình vẫn chạy nhưng Studio ngừng ghi thêm.
- Nếu cửa sổ CMD đang bật và tiến trình trả mã lỗi, CMD được giữ lại để người dùng đọc.
- Nút `Lịch sử` mở danh sách 10 lượt chạy gần nhất.
- Chọn từng lượt để đọc toàn bộ log ngay trong Studio, tải lại log đang chạy, sao chép hoặc mở file ngoài.

Đã kiểm tra:

- Python thành công có log.
- Python lỗi có traceback và mã thoát.
- Năm Python chạy đồng thời tạo đúng năm log riêng.
- Chạy 12 lượt chỉ giữ lại đúng 10 log.

Các file và log thử nghiệm đã được xóa.

## 7A. Voice Control nhúng trong Studio

- Mục `Điều khiển giọng nói` nằm ở thanh bên trái.
- Studio đọc đường dẫn file chạy từ `voice_control.json`.
- Đường dẫn hiện tại trỏ tới `D:\CodexProjects\VoiceControlV3_HoanChinh\CODEX_THO_MAY\CHAY_THO_MAY.bat`.
- BAT vẫn chạy như cũ khi mở trực tiếp.
- Khi Studio gọi BAT với `--studio`, Voice Control chạy tại `127.0.0.1:8766`, không tự mở Edge.
- Studio vẫn chạy tại `127.0.0.1:8765`.
- Giao diện Voice Control được nhúng bằng iframe có quyền microphone.
- Mở mục `Điều khiển giọng nói` chỉ hiển thị trạng thái; không tự khởi động.
- Voice Control chỉ chạy khi người dùng bấm nút `Khởi động`.
- Có nút Khởi động và `Tắt toàn bộ`.
- Có nút `Ẩn bảng nổi` và `Hiện bảng nổi`; chỉ đổi trạng thái cửa sổ Tkinter, không dừng server hoặc microphone.
- Trang Voice có ô đổi đường dẫn file chạy, ô đổi URL giao diện, nút `Chọn file chạy…` và `Lưu cấu hình`; người dùng không cần sửa JSON bằng tay.
- File `.bat/.cmd` được gọi qua CMD với `--studio`; `.py/.pyw` được gọi bằng Python với `--port 8766 --no-browser`; `.exe` được chạy trực tiếp.
- `Tắt toàn bộ` tìm PID đang giữ cổng `8766` rồi tắt cây tiến trình, nên vẫn hoạt động sau khi Studio tự cập nhật; Python server, CMD và bảng nổi Voice Control cùng đóng.
- Hai server độc lập; Voice Control lỗi không làm chết Trình chạy.

Đã kiểm tra server `8766` trả HTTP 200 và toàn bộ giao diện THỢ MÁY Voice Hub hiển thị trong Studio. Chưa tự bấm `Bắt đầu nghe` vì thao tác đó xin microphone và có thể thực thi lệnh thật.

## 8. Nhóm flow hiện tại

Danh sách thực tế nằm trong `launcher_groups.json` và có thể đã được người dùng đổi tên/sắp xếp tiếp bằng GUI. Các nhóm ban đầu:

- Chuẩn bị kế hoạch.
- Viết bài.
- Đăng bài.
- Phụ trợ đăng bài.
- Sau khi đăng.
- Chưa cần dùng.
- Kiểm tra.
- Các nhóm do người dùng tự thêm, ví dụ mở file Excel.

Không ghi đè `launchers.json` hoặc `launcher_groups.json` bằng danh sách mẫu khi người dùng đã chỉnh trong GUI.

## 9. Các file flow quan trọng đã đưa vào Studio

### Chuẩn bị kế hoạch

- `04_excel\nhap_du_lieu_ke_hoach_tu_thu_muc.py`

### Viết bài

- `02_viet_bai\chuan_bi_du_lieu_viet_bai_ten_mien.py`
- `02_viet_bai\code_tong_v3_no_mouse_CHAY_HIDE.py`
- `02_viet_bai\TUDONG_3_LUONG_WORD_BRIEF_VA_ANH.py`
- `02_viet_bai\TUDONG-CHAY_2_HOAC_3_LUONG_NO_MOUSE.py`
- `02_viet_bai\viet_lai_word_dot_2.py`

Lưu ý: file chính chạy 3 luồng là `TUDONG_3_LUONG_WORD_BRIEF_VA_ANH.py`; nó liên quan core `TUDONG-CHAY_2_HOAC_3_LUONG_NO_MOUSE.py`.

### Đăng bài

- `03_dang_bai\chuan_bi_du_lieu_dang_bai_moi_cate_post.py`
- `03_dang_bai\VIP_tudongdangbai_khong_ini_TEST.py`
- `03_dang_bai\VIP_tudongdangbai_3_5_luong_TEST.py`

### Phụ trợ

- `03_dang_bai\phu_tro\mo_word_tu_excel.py`
- `03_dang_bai\phu_tro\mo_url_dang_bai.py`
- `03_dang_bai\phu_tro\tong_hop_va_xu_ly_anh.py`
- `03_dang_bai\phu_tro\xu_ly_toan_bo_file_word.py`
- `03_dang_bai\phu_tro\bai_viet_lien_quan.py`
- `03_dang_bai\phu_tro\Vip_baivietlienquan_3_luong.py`
- `03_dang_bai\phu_tro\xuat_url_tu_id_cms.py`

Thư mục `03_dang_bai\phu_tro\chua can sai` được đưa vào nhóm Chưa cần dùng, không trộn với flow chính.

## 10. Sáu file phụ trợ đã đổi tên

Đã đổi tên file thật và cập nhật các tham chiếu liên quan:

- `mofileword_trongexcel.py` → `mo_word_tu_excel.py`
- `mo_url.py` → `mo_url_dang_bai.py`
- `tonghopanhv3sire900y.py` → `tong_hop_va_xu_ly_anh.py`
- `chinhdungchuanv2.py` → `xu_ly_toan_bo_file_word.py`
- `Vip_baivietlienquan.py` → `bai_viet_lien_quan.py`
- `xuatID_URL.py` → `xuat_url_tu_id_cms.py`

Đã cập nhật các lời gọi tương ứng trong:

- `01_hotkey\dangbaitag.ahk`
- `03_dang_bai\VIP_CHAY_1_BAI_DUNG_TRUOC_LUU.py`
- `03_dang_bai\VIP_tudongdangbai_khong_ini_TEST.py`

Không đổi các file chạy nhiều luồng không tham chiếu sáu tên này.

## 11. Điều chưa làm

- Chưa có chức năng xóa log theo từng nút trên giao diện.
- Chưa có nút mở code bằng VS Code/Notepad.
- Chưa chạy nhiều bước theo chuỗi.
- Chưa dừng flow tại bước lỗi và tiếp tục từ bước đó.
- Chưa có đường dẫn gốc linh hoạt để chuyển ổ/máy.
- Chưa đóng gói EXE hoặc tạo icon ứng dụng riêng.
- Chưa bật lại các khu chức năng đang ẩn.

## 12. Hướng làm tiếp được ưu tiên

1. Dùng thực tế hệ thống log Python và sửa lỗi nếu có.
2. Thêm menu cho mỗi nút: Chạy/Mở, Sửa file bằng VS Code, Mở thư mục, Sửa nút.
3. Bổ sung thao tác xóa log có xác nhận nếu cần.
4. Chạy các nút lần lượt theo thứ tự trong nhóm.
5. Thêm đường dẫn gốc để chuyển thư mục/máy không phải sửa từng nút.

Không thêm game, phim, gỡ phần mềm hoặc nhiều chức năng không liên quan trước khi flow công việc chính ổn định; tránh Studio phình to và khó dùng.

## 13. Cách bắt đầu một task Codex mới

Nói với Codex:

> Đọc toàn bộ `D:\CodexProjects\HotkeyVIP-Studio\BAN_GIAO_DU_AN.md`, kiểm tra trạng thái file hiện tại, không đụng `D:\CodexProjects\Hotkeyvip\07_ket_qua`, rồi tiếp tục yêu cầu của tôi.

Không yêu cầu Codex quét lại toàn bộ kho nếu công việc chỉ liên quan HotkeyVIP Studio.
