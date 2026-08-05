# TÀI LIỆU HƯỚNG DẪN QUY TRÌNH TỰ ĐỘNG HÓA BÀI VIẾT SEO & ĐĂNG BÀI CMS

Tài liệu này tổng hợp toàn bộ quy trình làm việc (Workflow) từ 11 file Python trong thư mục `D:\Thuc_Hanh_GitHub`. Hệ thống phục vụ việc tự động hóa lập kế hoạch bài viết, sinh nội dung AI (ChatGPT + Gemini), chuẩn bị dữ liệu đăng bài, cập nhật URL CMS và chèn bài viết liên quan (Internal Linking).

---

## 🗺️ Sơ Đồ Tổng Quan Quy Trình (Workflow)

```
[ GIAI ĐOẠN 1: LẬP KẾ HOẠCH & CHUẨN BỊ ]
  ├── 01_nhap_du_lieu_ke_hoach_phan_loai_trung.V2.1.py
  ├── 02_chuan_bi_du_lieu_viet_bai_ten_mien.V1.1.py
  └── 03_chuan_bi_du_lieu_viet_bai_ten_mien.V1.1_copy.py

[ GIAI ĐOẠN 2: VIẾT BÀI AI & TẠO ẢNH ]
  ├── 04_V1.11_vietbai_3cap_baohiem_anh_khong_ghi_de.py (Bản V1.11 - Đơn luồng)
  ├── 05_V2.14_CODEX_daluong_word_queue_anh_khong_ghi_de.py (Bản V2.14 - Đa luồng)
  └── 06_V2.17_CODEX_skip_loi_chay_lai.py (Bản V2.17 - Mới nhất, tự phục hồi lỗi)

[ GIAI ĐOẠN 3: TỔNG HỢP CHUẨN BỊ ĐĂNG BÀI ]
  └── 07_chuan_bi_DANG_BAI_tu_KE_HOACH_va_VIET_BAI.py

[ GIAI ĐOẠN 4: ĐĂNG BÀI CMS & ĐỒNG BỘ URL ]
  ├── 08_xuat_url_tu_id_cms.py
  └── 09_dong_bo_URL_bai_da_viet_sang_file_cong_ty_V1.0.py

[ GIAI ĐOẠN 5: CHÈN BÀI VIẾT LIÊN QUAN (INTERNAL LINKS) ]
  ├── 10_V1.0_bai_viet_lien_quan.py
  └── 11_V2.3_dieu_phoi_bai_viet_lien_quan.py
```

---

## 📑 Bảng Đánh Số và Đổi Tên File Theo Thứ Tự Thực Thi

| STT | Tên File Mới | Tên File Gốc | Chức Năng Chính & Thứ Tự Chạy |
|---|---|---|---|
| **01** | `01_nhap_du_lieu_ke_hoach_phan_loai_trung.V2.1.py` | `01_nhap_du_lieu_ke_hoach_phan_loai_trung.V2.1.py` | Import dữ liệu thô từ Excel nguồn vào sheet `KE_HOACH`, lọc bài viết bị trùng. |
| **02** | `02_chuan_bi_du_lieu_viet_bai_ten_mien.V1.1.py` | `02_chuan_bi_du_lieu_viet_bai_ten_mien.V1.1.py` | Đọc dữ liệu từ sheet `KE_HOACH` và đẩy sang sheet `VIET_BAI`, tạo Prompt và gán URL GPT. |
| **03** | `03_chuan_bi_du_lieu_viet_bai_ten_mien.V1.1_copy.py` | `chuan_bi_du_lieu_viet_bai_ten_mien.V1.1.py` | Bản duplicate dự phòng của bước 02. |
| **04** | `04_V1.11_vietbai_3cap_baohiem_anh_khong_ghi_de.py` | `V1.11_vietbai_3cap_baohiem_anh_khong_ghi_de.py` | Viết bài ChatGPT -> Lưu Word -> Tạo ảnh Gemini (Bản V1.11 đơn luồng). |
| **05** | `05_V2.14_CODEX_daluong_word_queue_anh_khong_ghi_de.py` | `V2.14_CODEX_daluong_word_queue_anh_khong_ghi_de.py` | Viết bài AI (Bản V2.14 đa luồng + Hàng đợi Word Queue). |
| **06** | `06_V2.17_CODEX_skip_loi_chay_lai.py` | `V2.17_CODEX_skip_loi_chay_lai.py` | Viết bài AI (Bản V2.17 mới nhất - Tối ưu skip lỗi, retry tự động, đa luồng). |
| **07** | `07_chuan_bi_DANG_BAI_tu_KE_HOACH_va_VIET_BAI.py` | `02_chuan_bi_DANG_BAI_tu_KE_HOACH_va_VIET_BAI.py` | Tổng hợp thông tin từ `KE_HOACH` và `VIET_BAI` sang sheet `DANG_BAI` (chuẩn bị cho khâu đăng CMS). |
| **08** | `08_xuat_url_tu_id_cms.py` | `xuat_url_tu_id_cms.py` | Lấy URL bài viết thật từ CMS dựa trên ID bài viết sau khi đăng bài. |
| **09** | `09_dong_bo_URL_bai_da_viet_sang_file_cong_ty_V1.0.py` | `03_dong_bo_URL_bai_da_viet_sang_file_cong_ty_V1.0.py` | Đồng bộ URL từ DANG_BAI quay lại sheet KE_HOACH công ty. |
| **10** | `10_V1.0_bai_viet_lien_quan.py` | `V1.0_bai_viet_lien_quan.py` | Tự động chèn bài viết liên quan (Internal Link) trên CMS bằng Selenium (Bản V1.0). |
| **11** | `11_V2.3_dieu_phoi_bai_viet_lien_quan.py` | `V2.3_dieu_phoi_bai_viet_lien_quan.py` | Điều phối chèn bài viết liên quan nâng cao (Bản V2.3 đa tính năng). |

---

## 🔍 Mô Tả Chi Tiết Từng Bước Trong Quy Trình

### 🟢 Giai đoạn 1: Khởi Tạo & Lập Kế Hoạch Bài Viết
- **Bước 01 (`01_nhap_du_lieu_ke_hoach_phan_loai_trung.V2.1.py`)**:
  - Quét các file Excel kế hoạch trong thư mục làm việc.
  - Tổng hợp thông tin tiêu đề, từ khóa, tên miền vào sheet `KE_HOACH`.
  - Phân loại và đánh dấu các tiêu đề/H1 bị trùng lặp.
- **Bước 02 & 03 (`02_chuan_bi_du_lieu_viet_bai_ten_mien.V1.1.py` & `03_..._copy.py`)**:
  - Lọc ra các bài chưa viết từ sheet `KE_HOACH`.
  - Ghép cấu hình Prompt và URL ChatGPT tương ứng cho từng chuyên mục.
  - Điền dữ liệu chuẩn bị sẵn sàng vào sheet `VIET_BAI`.

---

### 🟡 Giai đoạn 2: Tự Động Viết Bài & Sinh Ảnh Bằng AI
*(Người dùng chọn 1 trong 3 phiên bản để thực thi tùy nhu cầu)*
- **Bước 04 (`04_V1.11_...py`)**: Chạy đơn luồng với 3 cấp bảo hiểm phục hồi lỗi.
- **Bước 05 (`05_V2.14_...py`)**: Nâng cấp đa luồng + xử lý hàng đợi file Word không lo nghẽn.
- **Bước 06 (`06_V2.17_...py`)**: **[Khuyên dùng]** Bản hoàn thiện nhất, skip dòng lỗi tự động, cho phép chạy lại nhanh chóng mà không ghi đè dữ liệu ảnh cũ.

---

### 🟠 Giai đoạn 3: Tổng Hợp Dữ Liệu Đăng Bài
- **Bước 07 (`07_chuan_bi_DANG_BAI_tu_KE_HOACH_va_VIET_BAI.py`)**:
  - Ghép thông tin kế hoạch bài viết + đường dẫn file Word nội dung bài + đường dẫn ảnh 1, ảnh 2 đã tạo ở Giai đoạn 2.
  - Đẩy toàn bộ dữ liệu hoàn chỉnh vào sheet `DANG_BAI`.

---

### 🔵 Giai đoạn 4: Đăng Bài & Đồng Bộ URL CMS
- **Bước 08 (`08_xuat_url_tu_id_cms.py`)**:
  - Sau khi các bài viết được đăng lên CMS (thu được ID bài viết trên CMS), script này gửi request bất đồng bộ để tra cứu URL thực tế của bài viết.
- **Bước 09 (`09_dong_bo_URL_bai_da_viet_sang_file_cong_ty_V1.0.py`)**:
  - Đồng bộ URL thực tế thu được ở Bước 08 từ sheet `DANG_BAI` quay lại cột `URL Page` trong sheet tổng `KE_HOACH`, đồng thời đánh dấu trạng thái "Đã viết" / "OK".

---

### 🟣 Giai đoạn 5: Chèn Liên Kết Nội Bộ (Internal Links)
- **Bước 10 (`10_V1.0_bai_viet_lien_quan.py`)**:
  - Sử dụng Selenium điều khiển trình duyệt Edge truy cập CMS và chèn bài viết liên quan tuần tự.
- **Bước 11 (`11_V2.3_dieu_phoi_bai_viet_lien_quan.py`)**:
  - Phiên bản điều phối nâng cấp giúp tối ưu tốc độ chèn link bài viết liên quan, xử lý đa tiến trình và bỏ qua bài viết lỗi.

---

## 🛠️ Hướng Dẫn Vận Hành Hệ Thống

1. **Chuẩn bị**: Mở file Excel tổng (`hotkeyvip_test.xlsm`).
2. **Chạy lần lượt**:
   - Chạy script `01_...` để nạp kế hoạch.
   - Chạy script `02_...` để chuẩn bị bảng viết bài.
   - Chạy script `06_...` (Bản V2.17) để sinh nội dung & ảnh tự động.
   - Chạy script `07_...` để chuẩn bị bảng đăng bài.
   - Tiến hành đăng bài lên CMS, sau đó chạy `08_...` và `09_...` để hoàn tất đồng bộ URL.
   - Chạy script `11_...` để chèn bài viết liên quan tự động.
