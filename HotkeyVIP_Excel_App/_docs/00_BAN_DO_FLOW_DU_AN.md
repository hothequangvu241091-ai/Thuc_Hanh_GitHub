# BẢN ĐỒ FLOW DỰ ÁN HOTKEYVIP EXCEL APP

> Mục đích của file này: khi dự án lớn và bắt đầu bị ngợp, **mở file này trước** để biết app đang chạy gì, lỗi thuộc flow nào và cần sửa file nào.
>
> Tài liệu này chỉ tập trung vào **FLOW CÔNG VIỆC**, không cố giải thích toàn bộ code.

---

## 0. BẢN ĐỒ 30 GIÂY

App hiện tại có 2 phần hoàn toàn khác nhau:

### A. Phần xem/đối soát dữ liệu

```text
CHAY_APP.cmd
  ↓
excel_audit_app/main.py
  ↓
excel_audit_app/ui.py
  ↓
Nút Phân tích
  ↓
excel_audit_app/analysis.py
  ↓
Đọc KE_HOACH + VIET_BAI + DANG_BAI
  ↓
Hiển thị Tổng quan / lỗi / chênh lệch
```

**Quan trọng:** `analysis.py` chủ yếu phục vụ màn hình đối soát. Nếu Flow 1-8 ghi dữ liệu sai thì thường **không sửa `analysis.py`**.

### B. Phần chạy công việc thật

```text
CHAY_APP.cmd
  ↓
excel_audit_app/main.py
  ↓
excel_audit_app/ui.py
  ↓
Tab Công việc
  ↓
excel_audit_app/flow_catalog.py
  ↓
excel_audit_app/flow_host.py
  ↓
app_flows/01 → 08
```

Nói ngắn gọn:

- `ui.py` = giao diện, nút bấm, hộp xác nhận.
- `flow_catalog.py` = danh sách flow nào đang tồn tại và file nào được gọi.
- `flow_host.py` = cách app mở/bám Excel và chạy flow.
- `app_flows/*.py` = code nghiệp vụ thật.

Nếu chỉ nhớ 4 dòng trên là đã đủ định hướng phần lớn dự án.

---

# 1. FLOW CÔNG VIỆC HIỆN TẠI

## Flow 1 — Nhập KE_HOACH

File:

```text
app_flows/01_nhap_ke_hoach.py
```

Luồng:

```text
Các file Article nguồn
  ↓
Đọc dữ liệu
  ↓
Chống trùng / đồng bộ dòng đã có
  ↓
KE_HOACH
```

Flow này chịu trách nhiệm:

- tìm file nguồn;
- đọc sheet Article;
- lấy Title / Search Question / Main Keyword / H1 / danh mục / prompt;
- ghi `File nguồn`, `Vị trí nguồn`, `Trạng thái nguồn`;
- xử lý trùng và đồng bộ lại dòng đã có.

### Muốn sửa gì thì vào đây?

- đổi thư mục chứa file nguồn;
- đổi cách xác định domain từ file;
- đổi cột nào được nhập vào KE_HOACH;
- đổi quy tắc trùng;
- đổi cách cập nhật dòng đã tồn tại.

Không sửa UI nếu vấn đề nằm ở dữ liệu nhập KE_HOACH.

---

## Flow 2 — Chuẩn bị VIET_BAI

File:

```text
app_flows/02_chuan_bi_viet_bai.py
```

Luồng:

```text
KE_HOACH
  ↓
Đối chiếu bài nào cần viết / đã viết / đã đăng / đã xóa
  ↓
Tạo hoặc cập nhật VIET_BAI
```

Flow này chịu trách nhiệm:

- đối chiếu KE_HOACH với VIET_BAI;
- thêm bài còn thiếu;
- cập nhật Prompt viết bài;
- cập nhật loại GPT / URL GPT;
- đánh dấu các trạng thái như đã viết, đã đăng, bài đã xóa, không còn trong kế hoạch;
- tránh tạo thêm dòng vô lý khi chạy lại.

### Muốn sửa gì thì vào đây?

- quy tắc khi nào thêm bài vào VIET_BAI;
- quy tắc `Bài đã xóa`;
- quy tắc `Đã viết` / `Đã đăng`;
- cách xác định bài không còn trong KE_HOACH;
- cột Prompt / GPT được lấy từ KE_HOACH.

---

## Flow 3 — Viết bài + tạo ảnh

File chính:

```text
app_flows/03_viet_bai_tao_anh.py
```

Luồng lớn:

```text
VIET_BAI
  ↓
Chọn bài cần chạy
  ↓
ChatGPT viết bài
  ↓
Lưu Word
  ↓
Xin brief ảnh
  ↓
Gemini tạo ảnh 1 + ảnh 2
  ↓
Ghi Word / ảnh / trạng thái / lỗi về VIET_BAI
```

Đây là một trong hai file phức tạp nhất dự án.

Nó đang chứa cùng lúc:

- Selenium;
- Edge profile;
- ChatGPT;
- Word;
- Gemini;
- multi-worker;
- retry;
- checkpoint;
- kiểm tra số từ;
- đường dẫn output;
- ghi trạng thái Excel.

### Phần chọn thứ tự bài trước khi chạy

Không nằm hoàn toàn trong Flow 3.

App dùng:

```text
excel_audit_app/write_plan.py
```

để xem trước hàng chờ:

```text
Bài lỗi trước
  ↓
Tên miền ưu tiên
  ↓
Các bài bình thường sau mốc OK OK
```

Giao diện chọn ưu tiên nằm trong:

```text
excel_audit_app/ui.py
```

### Muốn sửa gì thì vào đâu?

**Đổi thứ tự / ưu tiên bài:**

```text
excel_audit_app/write_plan.py
```

**Đổi giao diện chọn tên miền / số bài ưu tiên:**

```text
excel_audit_app/ui.py
```

**Đổi cách ChatGPT viết / retry / Word / Gemini / worker:**

```text
app_flows/03_viet_bai_tao_anh.py
```

**Đổi thư mục Word / profile worker / download:**

```text
app_flows/03_viet_bai_tao_anh.py
```

---

## Flow 4 — Chuẩn bị DANG_BAI

File:

```text
app_flows/04_chuan_bi_dang_bai.py
```

Luồng:

```text
VIET_BAI có Trạng thái hoàn tất = OK
  ↓
Đối chiếu Combo 4
  ↓
Lấy Word + ảnh + danh mục
  ↓
DANG_BAI
```

Flow này chịu trách nhiệm đưa dữ liệu đủ điều kiện từ VIET_BAI sang DANG_BAI mà không nhân dòng.

### Muốn sửa gì thì vào đây?

- điều kiện bài nào được chuyển sang DANG_BAI;
- Combo 4;
- cách lấy Word cũ / Word mới;
- điều kiện dùng bài viết mới;
- cách lấy ảnh;
- cách lấy danh mục;
- kiểm tra trùng trong DANG_BAI.

---

## Flow 5 — Đăng bài CMS

File chính:

```text
app_flows/05_dang_bai_cms.py
```

Luồng:

```text
DANG_BAI
  ↓
Chọn batch cần đăng
  ↓
Mở Word nền
  ↓
Lấy nội dung + ảnh
  ↓
Selenium mở CMS
  ↓
Điền tiêu đề / nội dung / danh mục / tác giả
  ↓
Lưu CMS
  ↓
Ghi trạng thái / ID / thời gian về DANG_BAI
```

Đây là file phức tạp thứ hai của dự án.

### Phần chọn batch trước khi đăng

App dùng:

```text
excel_audit_app/publish_plan.py
```

Luật hiện tại:

```text
Mỗi tên miền
  ↓
Chọn 1 danh mục
  ↓
Mặc định danh mục có nhiều bài hợp lệ nhất
  ↓
Lấy tối đa N bài / tên miền
```

Giao diện chọn batch nằm trong:

```text
excel_audit_app/ui.py
```

### Muốn sửa gì thì vào đâu?

**Đổi luật chọn bài / chọn danh mục / số bài:**

```text
excel_audit_app/publish_plan.py
```

**Đổi mặc định trên cửa sổ chọn batch:**

```text
excel_audit_app/ui.py
```

**Đổi cách thao tác CMS / Selenium / Word / ảnh / login / lưu bài:**

```text
app_flows/05_dang_bai_cms.py
```

**Đổi cấu hình website/profile/login dùng chung:**

Flow 5 còn phụ thuộc dữ liệu runtime bên ngoài repo tại:

```text
D:\CodexProjects\Hotkeyvip
```

và module cấu hình `hotkeyvip_config` ở môi trường runtime.

---

## Flow 6 — Lấy URL thật từ ID CMS

File:

```text
app_flows/06_lay_url_cms.py
```

Luồng:

```text
DANG_BAI có ID CMS
  ↓
Tạo URL tạm
  ↓
HEAD/GET website
  ↓
Theo redirect
  ↓
URL thật
  ↓
Ghi về DANG_BAI
```

### Muốn sửa gì thì vào đây?

- công thức tạo URL tạm;
- timeout request;
- số request chạy đồng thời;
- cột ID / URL;
- cách xử lý redirect.

Các cấu hình dễ thấy:

```text
TIMEOUT_S
CONCURRENCY
```

---

## Flow 7 — Bài viết liên quan

File:

```text
app_flows/07_bai_viet_lien_quan.py
```

Luồng:

```text
DANG_BAI đã có ID
  ↓
Mở trang sửa bài CMS
  ↓
Chọn bài liên quan
  ↓
Lưu lại CMS
  ↓
Ghi trạng thái về Excel
```

### Muốn sửa gì thì vào đây?

- số bài liên quan;
- cách chọn bài gần nhất;
- cách chọn bài cùng danh mục;
- cách mở URL sửa CMS;
- cách lưu bài;
- cách ghi trạng thái về Excel.

Hiện có các hằng số chính:

```text
NEAREST_RELATED_COUNT
SAME_CATEGORY_RELATED_COUNT
TARGET_RELATED_COUNT
```

---

## Flow 8 — Đồng bộ URL về file công ty

File:

```text
app_flows/08_dong_bo_url.py
```

Luồng:

```text
DANG_BAI
  ↓
Lấy URL đã đăng
  ↓
Tìm file domain tương ứng bên Google Drive G:\
  ↓
Đối chiếu Title / Keyword / H1
  ↓
Backup
  ↓
Cập nhật URL Page ở file công ty
```

### Muốn sửa gì thì vào đây?

- thư mục file công ty;
- quy tắc đối chiếu;
- cách xử lý `Bài đã xóa`;
- cách backup;
- số backup giữ lại;
- cách ghi URL về file đích.

Các vùng cần nhìn đầu tiên:

```text
TARGET_DIRECTORY
BACKUP_ROOT
MAX_BACKUPS
SOURCE_REQUIRED_HEADERS
TARGET_HEADERS_FULL
TARGET_HEADERS_FALLBACK
```

---

# 2. FLOW PHỤ KHÔNG NẰM TRONG 8 FLOW CHÍNH

## Chuyển URL ngày đăng mới nhất sang app Submit

File:

```text
excel_audit_app/submit_transfer.py
```

Luồng:

```text
DANG_BAI
  ↓
Tìm ngày đăng mới nhất
  ↓
Lấy URL ĐÃ ĐĂNG trong ngày đó
  ↓
Loại URL trùng / URL admin / URL sai
  ↓
Backup danh sách Submit cũ
  ↓
Ghi danh sách mới cho app Submit
```

Nếu nút **Chuyển URL** có vấn đề thì xem:

```text
excel_audit_app/submit_transfer.py
```

Nếu muốn sửa giao diện/nút xác nhận thì xem:

```text
excel_audit_app/ui.py
```

---

# 3. APP THỰC SỰ CHẠY FLOW NHƯ THẾ NÀO?

Khi bấm một nút Flow trong tab `Công việc`:

```text
ui.py
  ↓
flow_by_key()
  ↓
flow_catalog.py
  ↓
Lấy đường dẫn script tương ứng
  ↓
ui.py gọi
python -m excel_audit_app.flow_host
  ↓
flow_host.py
  ↓
Mở hoặc bám đúng workbook Excel
  ↓
Chạy script trong app_flows
```

### Flow chạy trực tiếp trên workbook Excel do app quản lý

Hiện gồm:

```text
01_nhap_ke_hoach.py
02_chuan_bi_viet_bai.py
04_chuan_bi_dang_bai.py
06_lay_url_cms.py
08_dong_bo_url.py
```

### Flow chạy process riêng

Hiện gồm:

```text
03_viet_bai_tao_anh.py
05_dang_bai_cms.py
07_bai_viet_lien_quan.py
```

Lý do dễ hiểu:

- Flow Excel thuần có thể làm trực tiếp với workbook.
- Flow có Selenium/Edge/Word/CMS nặng được tách process để app chính đỡ bị treo.

### Nếu lỗi kiểu:

- không bám được file Excel;
- Excel ẩn bị kẹt;
- workbook bị khóa;
- flow không khởi động;
- process không trả log;

thì **đừng sửa Flow 1-8 trước**.

Hãy xem:

```text
excel_audit_app/flow_host.py
```

---

# 4. MUỐN SỬA CÁI GÌ → MỞ FILE NÀO?

| Muốn sửa | File nên mở đầu tiên |
|---|---|
| App không mở | `CHAY_APP.cmd`, sau đó `excel_audit_app/main.py` |
| Tên app / cửa sổ / tab / nút | `excel_audit_app/ui.py` |
| Thêm/bớt/đổi tên Flow | `excel_audit_app/flow_catalog.py` |
| Flow gọi sai script | `excel_audit_app/flow_catalog.py` |
| Excel ẩn / workbook bị khóa / flow không chạy | `excel_audit_app/flow_host.py` |
| Dashboard đếm sai / Combo 4 đối soát sai | `excel_audit_app/analysis.py` |
| App đọc tên sheet/cột Excel sai | `excel_audit_app/excel_io.py` |
| Thứ tự ưu tiên viết bài | `excel_audit_app/write_plan.py` |
| Giao diện chọn ưu tiên viết | `excel_audit_app/ui.py` |
| Luật chọn batch đăng bài | `excel_audit_app/publish_plan.py` |
| Giao diện chọn batch đăng | `excel_audit_app/ui.py` |
| Nhập dữ liệu KE_HOACH | `app_flows/01_nhap_ke_hoach.py` |
| Chuẩn bị VIET_BAI | `app_flows/02_chuan_bi_viet_bai.py` |
| ChatGPT / Word / Gemini / ảnh | `app_flows/03_viet_bai_tao_anh.py` |
| Chuẩn bị DANG_BAI | `app_flows/04_chuan_bi_dang_bai.py` |
| Đăng CMS | `app_flows/05_dang_bai_cms.py` |
| Lấy URL từ ID | `app_flows/06_lay_url_cms.py` |
| Bài viết liên quan | `app_flows/07_bai_viet_lien_quan.py` |
| Đồng bộ URL file công ty | `app_flows/08_dong_bo_url.py` |
| Chuyển URL sang Submit | `excel_audit_app/submit_transfer.py` |
| Xuất file đối soát / khôi phục DANG | `excel_audit_app/report_export.py` + các file `.ps1` liên quan |
| App nhớ file / nhớ phiên trước | `excel_audit_app/session.py` |

---

# 5. VÍ DỤ THỰC TẾ: TAO MUỐN SỬA MỘT CÁI THÌ ĐI ĐÂU?

## Ví dụ 1 — Muốn đổi mặc định mỗi domain đăng 7 bài thành 10 bài

Đầu tiên mở:

```text
excel_audit_app/ui.py
```

Tìm phần thiết lập batch đăng bài và giá trị mặc định hiện tại.

Nếu muốn đổi luôn **logic chọn batch**, mở thêm:

```text
excel_audit_app/publish_plan.py
```

Không cần sửa `05_dang_bai_cms.py` nếu CMS vẫn đăng đúng các dòng được đưa vào batch.

---

## Ví dụ 2 — Muốn đổi luật ưu tiên bài viết

Ví dụ muốn:

```text
bài lỗi
→ domain A
→ domain B
→ bài bình thường
```

Mở:

```text
excel_audit_app/write_plan.py
```

Nếu cần thêm ô chọn domain B trên giao diện thì sửa thêm:

```text
excel_audit_app/ui.py
```

Không nên nhét luật xếp hàng mới trực tiếp vào `03_viet_bai_tao_anh.py`.

---

## Ví dụ 3 — ChatGPT viết xong nhưng Word sai

Mở:

```text
app_flows/03_viet_bai_tao_anh.py
```

Tìm phần Word.

Không sửa:

```text
analysis.py
publish_plan.py
05_dang_bai_cms.py
```

vì chúng không tạo file Word.

---

## Ví dụ 4 — DANG_BAI bị thêm sai dòng

Luồng cần kiểm tra trước:

```text
VIET_BAI
  ↓
Flow 4
  ↓
DANG_BAI
```

Mở:

```text
app_flows/04_chuan_bi_dang_bai.py
```

Nếu dữ liệu đã sai từ VIET_BAI thì quay ngược lại Flow 2 hoặc Flow 3.

---

## Ví dụ 5 — CMS điền sai nội dung / sai danh mục / sai tác giả

Mở:

```text
app_flows/05_dang_bai_cms.py
```

Nếu **batch chọn sai bài trước khi CMS chạy**, mở:

```text
excel_audit_app/publish_plan.py
```

Phân biệt rõ hai lỗi này để tránh sửa nhầm chỗ.

---

## Ví dụ 6 — URL đã đăng đúng nhưng file công ty không được cập nhật

Kiểm tra:

```text
app_flows/08_dong_bo_url.py
```

Không cần sửa Flow 5 nếu DANG_BAI đã có URL đúng.

---

# 6. THƯ MỤC NÀO THỪA / CÓ THỂ DỌN?

## Nhóm A — Có thể xóa khỏi repo ngay

### `__pycache__`

Đang có ít nhất ở:

```text
app_flows/__pycache__/
excel_audit_app/__pycache__/
```

Đây là file cache Python tự sinh.

**Không phải source code.**

Có thể xóa và nên thêm `.gitignore` sau này:

```gitignore
__pycache__/
*.pyc
```

---

## Nhóm B — Không tham gia flow đang chạy, có thể đưa ra khỏi repo chính

### `_archive/code_cu/`

Hiện chứa các bản cũ lớn như:

```text
03_V1.12_...
V2.21_...
05_V1.8_...
V1.0_bai_viet_lien_quan.py
```

Flow đang chạy đã có bản mới trong `app_flows/`.

Đề xuất:

```text
Nếu Git history đã giữ đủ phiên bản cũ
→ xóa _archive khỏi nhánh làm việc chính.
```

Hoặc nếu vẫn muốn giữ để đối chiếu:

```text
đưa sang repo/archive riêng
```

**Lưu ý:** `flow_catalog.py` đang có danh sách `LEGACY_OR_STANDALONE` nhắc tới một số file cũ. Nếu xóa `_archive`, nên dọn luôn các dòng metadata đó để tài liệu không trỏ vào file đã xóa.

---

### `_tests_data/`

Hiện chứa các file `.xlsm` test lớn.

Chúng không cần để app chạy hằng ngày.

Đề xuất tốt hơn:

```text
repo chính: chỉ giữ 1 file test nhỏ tối thiểu
file test Excel lớn: chuyển ra thư mục test ngoài repo / Google Drive
```

Không nên để file Excel test vài MB tăng dần trong Git vì Git lưu lịch sử các phiên bản file nhị phân rất nặng.

---

## Nhóm C — Công cụ độc lập, không thuộc flow app

### `_tools/`

Hiện có các tiện ích như:

```text
VIPPPPPP_baivietxoa_xoa_nhu_lam_tay_COM.V2.2.py
mo_profile_worker.py
```

Theo cấu trúc hiện tại chúng **không phải Flow 1-8**.

Nếu vẫn dùng thường xuyên: giữ.

Nếu ít dùng: chuyển sang repo/tool riêng để app chính nhìn gọn hơn.

---

## Nhóm D — Nên giữ

```text
CHAY_APP.cmd
README_APP.txt
excel_audit_app/
app_flows/
tests/
_docs/
```

### `tests/`

Không nên xóa chỉ vì không chạy khi mở app.

Test giúp bảo vệ các quy tắc đối soát khi sau này sửa code.

### `_docs/`

Nên giữ, nhưng có thể gom lại để dễ đọc.

File này nên được xem là **file vào cửa**.

Các tài liệu kỹ thuật dài chỉ mở khi thật sự cần chi tiết.

---

# 7. CÓ THỂ GOM LẠI THÀNH NHỮNG “SKILL” NÀO?

Mục tiêu không phải viết lại toàn bộ dự án ngay.

Chỉ cần từ giờ xem code theo 6 skill lớn.

## Skill 1 — APP SHELL / ĐIỀU PHỐI

Gom tư duy các file:

```text
main.py
ui.py
flow_catalog.py
flow_host.py
session.py
```

Nhiệm vụ:

```text
Mở app
→ chọn file
→ hiển thị nút
→ hỏi xác nhận
→ chạy flow
→ hiện log
```

Đây là **vỏ app**, không phải nghiệp vụ SEO/CMS.

---

## Skill 2 — EXCEL CORE / ĐỐI SOÁT

Gom tư duy:

```text
analysis.py
excel_io.py
report_export.py
export_with_excel.ps1
recover_dang_bai_with_excel.ps1
```

Nhiệm vụ:

```text
đọc Excel
→ chuẩn hóa heading
→ Combo 4
→ kiểm tra dữ liệu
→ xuất báo cáo / khôi phục
```

Sau này nên cố gắng dùng chung skill này thay vì mỗi Flow tự viết lại hàm `normalize`, `headers`, `require_columns` riêng.

---

## Skill 3 — CHUẨN BỊ DỮ LIỆU CONTENT

Gồm:

```text
01_nhap_ke_hoach.py
02_chuan_bi_viet_bai.py
write_plan.py
```

Luồng:

```text
File nguồn
→ KE_HOACH
→ VIET_BAI
→ xếp hàng viết
```

---

## Skill 4 — SẢN XUẤT CONTENT

Hiện gần như nằm hết trong:

```text
03_viet_bai_tao_anh.py
```

Nhưng về sau nên tách nhỏ theo chức năng:

```text
content_engine/
  chatgpt.py
  word.py
  brief.py
  gemini.py
  worker.py
  retry.py
```

Không cần tách ngay một lần.

Khi lần tới sửa phần nào lớn thì tách phần đó ra trước.

---

## Skill 5 — ĐĂNG BÀI CMS

Gồm tư duy:

```text
04_chuan_bi_dang_bai.py
publish_plan.py
05_dang_bai_cms.py
06_lay_url_cms.py
07_bai_viet_lien_quan.py
```

Luồng:

```text
VIET_BAI OK
→ DANG_BAI
→ chọn batch
→ đăng CMS
→ lấy URL
→ bổ sung bài liên quan
```

Đây nên được xem là **một cụm nghiệp vụ**, không phải 5 dự án khác nhau.

---

## Skill 6 — ĐỒNG BỘ SAU ĐĂNG

Gồm:

```text
08_dong_bo_url.py
submit_transfer.py
```

Luồng:

```text
DANG_BAI hoàn tất
→ đồng bộ URL về file công ty
→ chuyển URL mới sang hệ thống Submit
```

---

# 8. ĐỀ XUẤT CẤU TRÚC GỌN HƠN TRONG TƯƠNG LAI

Không cần đổi ngay.

Đây chỉ là hướng gom khi dự án tiếp tục lớn:

```text
HotkeyVIP_Excel_App/
│
├─ CHAY_APP.cmd
├─ README_APP.txt
│
├─ excel_audit_app/
│  ├─ shell/
│  │  ├─ main.py
│  │  ├─ ui.py
│  │  ├─ flow_catalog.py
│  │  └─ flow_host.py
│  │
│  ├─ core/
│  │  ├─ excel_io.py
│  │  ├─ matching.py
│  │  ├─ analysis.py
│  │  └─ config.py
│  │
│  └─ skills/
│     ├─ planning/
│     ├─ content/
│     ├─ publishing/
│     └─ sync/
│
├─ tests/
└─ _docs/
```

Nhưng **không nên refactor toàn bộ ngay** vì rủi ro làm hỏng app đang chạy.

Ưu tiên thực tế:

```text
1. Giữ nguyên Flow 1-8 đang chạy.
2. Dọn cache + code cũ.
3. Tách dần Flow 3.
4. Tách dần Flow 5.
5. Gom hàm Excel dùng chung sau cùng.
```

---

# 9. QUY TẮC SỬA CODE ĐỂ KHÔNG BỊ NGỢP

Mỗi lần cần sửa, chỉ làm 5 bước này.

## Bước 1 — Nói lỗi thuộc giai đoạn nào

```text
Nguồn
KE_HOACH
VIET_BAI
Viết bài/ảnh
DANG_BAI
CMS
URL
Đồng bộ
```

## Bước 2 — Chọn đúng Flow

```text
Nguồn → KE_HOACH       = Flow 1
KE_HOACH → VIET_BAI    = Flow 2
VIET_BAI → Word/Ảnh    = Flow 3
VIET_BAI → DANG_BAI    = Flow 4
DANG_BAI → CMS         = Flow 5
CMS ID → URL           = Flow 6
Bài liên quan          = Flow 7
URL → file công ty     = Flow 8
```

## Bước 3 — Chỉ mở file đó trước

Đừng mở cả project rồi tìm lung tung.

Ví dụ lỗi Flow 4:

```text
chỉ mở 04_chuan_bi_dang_bai.py trước
```

## Bước 4 — Chỉ mở file điều phối nếu lỗi nằm trước khi flow chạy

Nếu:

- nút sai;
- batch sai;
- flow không mở;
- Excel bị khóa;
- app gọi sai script;

thì mới quay về:

```text
ui.py
flow_catalog.py
flow_host.py
```

## Bước 5 — Test đúng một đoạn flow

Không cần chạy hết 1 → 8 sau mỗi thay đổi.

Ví dụ sửa Flow 2:

```text
chuẩn bị 1 file test
→ chạy Flow 2
→ kiểm tra VIET_BAI
```

Xong mới commit.

---

# 10. NHỮNG FILE KHÔNG NÊN SỬA NHẦM

## Đừng sửa `_archive` để chữa app hiện tại

Code trong `_archive` là bản cũ.

App hiện tại chạy file trong:

```text
app_flows/
```

## Đừng sửa `analysis.py` khi flow ghi Excel sai

`analysis.py` chủ yếu đọc và đối soát.

Nếu dữ liệu Excel bị Flow 2 ghi sai thì sửa Flow 2.

## Đừng sửa `05_dang_bai_cms.py` nếu chỉ sai luật chọn batch

Luật chọn batch nằm ở:

```text
publish_plan.py
```

## Đừng sửa `03_viet_bai_tao_anh.py` nếu chỉ sai thứ tự hàng chờ

Thứ tự hàng chờ nằm ở:

```text
write_plan.py
```

Đây là các điểm dễ sửa nhầm nhất.

---

# 11. NẾU MUỐN THÊM FLOW 9 SAU NÀY

Quy trình nên là:

```text
1. Tạo file app_flows/09_ten_flow.py
2. Đảm bảo file có main()
3. Thêm FlowDefinition vào excel_audit_app/flow_catalog.py
4. Nếu flow chỉ cần nút Chạy bình thường → UI sẽ tự sinh nút từ ACTIVE_FLOWS
5. Nếu flow cần màn hình chọn riêng → thêm phần preview vào ui.py
6. Nếu cần cách mở Excel đặc biệt → mới sửa flow_host.py
```

Không cần nhét logic Flow 9 trực tiếp vào `ui.py`.

---

# 12. BẢN GHI NHỚ NGẮN NHẤT

Nếu mai mở dự án lên mà lại quên hết, chỉ nhớ:

```text
CHAY_APP.cmd
   ↓
main.py
   ↓
ui.py
   ↓
flow_catalog.py
   ↓
flow_host.py
   ↓
app_flows/01 → 08
```

Và luồng nghiệp vụ:

```text
FILE NGUỒN
   ↓ Flow 1
KE_HOACH
   ↓ Flow 2
VIET_BAI
   ↓ Flow 3
WORD + ẢNH
   ↓ Flow 4
DANG_BAI
   ↓ Flow 5
CMS
   ↓ Flow 6
URL THẬT
   ↓ Flow 7
BÀI LIÊN QUAN
   ↓ Flow 8
FILE CÔNG TY
```

**Khi có lỗi: xác định nó đang đứng ở mũi tên nào → sửa đúng Flow ở mũi tên đó.**

Đó là cách đơn giản nhất để quản lý dự án này mà không cần hiểu toàn bộ code cùng lúc.
