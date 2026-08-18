from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True, slots=True)
class FlowDefinition:
    key: str
    name: str
    description: str
    script: str
    confirmation: str
    script_args: tuple[str, ...] = ()
    external_effects: bool = False

    def script_path(self, project_root: Path) -> Path:
        return project_root / self.script


ACTIVE_FLOWS: tuple[FlowDefinition, ...] = (
    FlowDefinition(
        key="import_ke_hoach",
        name="1. Nhập KE_HOACH",
        description="Đọc các file Article nguồn và cập nhật KE_HOACH của file đang chọn.",
        script="app_flows/01_nhap_ke_hoach.py",
        confirmation="Flow sẽ cập nhật sheet KE_HOACH trong file đang chọn.",
    ),
    FlowDefinition(
        key="prepare_viet_bai",
        name="2. Chuẩn bị VIET_BAI",
        description="Đối chiếu KE_HOACH, thêm bài còn thiếu và cập nhật trạng thái VIET_BAI.",
        script="app_flows/02_chuan_bi_viet_bai.py",
        confirmation="Flow sẽ cập nhật dữ liệu và trạng thái trong VIET_BAI.",
    ),
    FlowDefinition(
        key="write_articles",
        name="3. Viết bài + tạo ảnh",
        description=(
            "Chạy engine V2.24 với trạng thái Worker bám tiến độ ChatGPT, Word và Gemini; "
            "ghi kết quả vào VIET_BAI."
        ),
        script="app_flows/03_viet_bai_tao_anh.py",
        confirmation=(
            "Flow sẽ mở trình duyệt, sử dụng tài khoản ChatGPT/Gemini, tạo file Word/ảnh "
            "và cập nhật VIET_BAI."
        ),
        external_effects=True,
    ),
    FlowDefinition(
        key="prepare_dang_bai",
        name="4. Chuẩn bị DANG_BAI",
        description="Đưa các bài VIET_BAI đủ điều kiện sang DANG_BAI theo Combo 4.",
        script="app_flows/04_chuan_bi_dang_bai.py",
        confirmation="Flow sẽ thêm hoặc cập nhật các dòng trong DANG_BAI.",
    ),
    FlowDefinition(
        key="publish_articles",
        name="5. Đăng bài CMS",
        description="Xem trước batch theo tên miền + một danh mục, rồi chạy V2.10 và ghi kết quả về DANG_BAI.",
        script="app_flows/05_dang_bai_cms.py",
        confirmation=(
            "Flow có thể tạo và lưu bài thật trên các website CMS, đồng thời cập nhật DANG_BAI."
        ),
        external_effects=True,
    ),
    FlowDefinition(
        key="export_urls",
        name="6. Lấy URL từ ID CMS",
        description=(
            "Chạy engine V1.7: nếu HEAD không chuyển được thì tạo URL từ H1; "
            "ô URL tự tạo được tô vàng và thêm Note để kiểm tra."
        ),
        script="app_flows/06_lay_url_cms.py",
        confirmation="Flow sẽ gửi yêu cầu mạng tới website và cập nhật URL trong DANG_BAI.",
        external_effects=True,
    ),
    FlowDefinition(
        key="related_articles",
        name="7. Bài viết liên quan",
        description=(
            "Chạy engine V2.5 với 5 worker, khóa đăng nhập/lưu riêng theo tên miền, "
            "tự retry và ghi Excel qua hàng đợi riêng."
        ),
        script="app_flows/07_bai_viet_lien_quan.py",
        confirmation=(
            "Flow sẽ chỉnh sửa các bài đã đăng trên website CMS và ghi trạng thái về DANG_BAI."
        ),
        script_args=("3",),
        external_effects=True,
    ),
    FlowDefinition(
        key="sync_company_urls",
        name="8. Đồng bộ URL về file công ty",
        description="Đồng bộ URL từ DANG_BAI sang các file domain bên ngoài và tạo backup.",
        script="app_flows/08_dong_bo_url.py",
        confirmation=(
            "Flow sẽ cập nhật nhiều file Excel công ty bên ngoài thư mục dự án và cập nhật file đang chọn."
        ),
        external_effects=True,
    ),
)


LEGACY_OR_STANDALONE: tuple[tuple[str, str], ...] = (
    ("_archive/code_cu/03_V1.12_vietbai_3cap_baohiem_anh_khong_ghi_de.py", "Bản viết bài cũ; thay bằng V2.22."),
    ("_archive/code_cu/V2.21_CODEX_skip_loi_chay_lai.py", "Bản viết bài cũ; thay bằng V2.22."),
    ("_archive/code_cu/05_V1.8_VIP_tudongdangbai_pipeline_kiemtra_noidung.py", "Bản đăng bài cũ; thay bằng V2.10."),
    ("_archive/code_cu/V1.0_bai_viet_lien_quan.py", "Bản bài viết liên quan cũ; thay bằng V2.3."),
    (
        "_tools/VIPPPPPP_baivietxoa_xoa_nhu_lam_tay_COM.V2.2.py",
        "Công cụ xóa hàng trên nhiều file domain; không dùng file Excel đang chọn nên giữ chạy độc lập.",
    ),
    ("_tools/mo_profile_worker.py", "Tiện ích mở profile worker; không phải flow Excel độc lập."),
)


def flow_by_key(key: str) -> FlowDefinition:
    for flow in ACTIVE_FLOWS:
        if flow.key == key:
            return flow
    raise KeyError(key)
