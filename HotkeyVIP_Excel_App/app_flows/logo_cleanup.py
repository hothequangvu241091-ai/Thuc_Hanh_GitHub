"""Xóa logo cố định khỏi ảnh Gemini bằng LaMa GPU.

Module này không có giao diện và không quét thư mục. Flow 3 gọi trực tiếp với
đường dẫn của từng ảnh vừa lưu. Mọi lỗi được trả về dạng kết quả để không làm
gián đoạn luồng tạo bài/ảnh.
"""

from __future__ import annotations

import os
import sys
import threading
from dataclasses import dataclass
from pathlib import Path


# Flow 3 có thể được chạy từ cmd/PowerShell Windows dùng bảng mã cũ.
# Chuẩn hóa log UTF-8 để thông báo tiếng Việt không biến thành lỗi xử lý ảnh.
if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")


# Tọa độ theo tỷ lệ ảnh (x, y), nên vẫn đúng khi Gemini thay đổi độ phân giải.
# Logo Gemini hiện tại là biểu tượng lấp lánh sát góc phải-dưới, tâm quanh 95%
# chiều rộng và 93.4% chiều cao. Tọa độ dưới đây được đo trực tiếp trên ảnh
# Gemini 800x597 và dùng mask hình thoi nhỏ để không tạo mảng vá hình chữ nhật.
# Vùng cũ (~89%/~86%) là vị trí của mẫu thử cũ, không đúng ảnh Gemini thực tế.
LOGO_POLY = (
    (0.950, 0.912),
    (0.968, 0.934),
    (0.950, 0.957),
    (0.932, 0.934),
)
MASK_PADDING_PX = 3
AI_CONTEXT_PX = 96
FEATHER_PX = 1.2
MODEL_PATH = (
    Path(__file__).resolve().parent.parent
    / "_runtime"
    / "logo_cleanup"
    / "big-lama.pt"
)


@dataclass(frozen=True)
class LogoCleanupResult:
    success: bool
    error: str = ""


_INIT_LOCK = threading.RLock()
_GPU_LOCK = threading.RLock()
_LAMA = None
_INIT_ERROR = None
_INITIALIZED = False
_PREFLIGHT_ERROR = None
_PREFLIGHT_DONE = False


def _short_error(exc: Exception | str) -> str:
    return " ".join(str(exc).split())[:180] or type(exc).__name__


def _load_lama_once() -> tuple[object | None, str]:
    """Nạp model đúng một lần cho cả phiên Flow 3; không lỗi cứng khi thiếu."""
    global _LAMA, _INIT_ERROR, _INITIALIZED
    with _INIT_LOCK:
        if _INITIALIZED:
            return _LAMA, _INIT_ERROR or ""
        _INITIALIZED = True
        try:
            preflight = prepare_logo_cleanup()
            if not preflight.success:
                _INIT_ERROR = preflight.error
                return None, _INIT_ERROR
            import torch
            from simple_lama_inpainting import SimpleLama
            os.environ["LAMA_MODEL"] = str(MODEL_PATH)
            _LAMA = SimpleLama(device=torch.device("cuda"))
            print(f"-> Đã nạp AI xóa logo trên GPU: {torch.cuda.get_device_name(0)}")
        except Exception as exc:
            _INIT_ERROR = _short_error(exc)
            print(f"⚠️ Xóa logo không sẵn sàng, sẽ giữ ảnh gốc: {_INIT_ERROR}")
        return _LAMA, _INIT_ERROR or ""


def prepare_logo_cleanup() -> LogoCleanupResult:
    """Kiểm tra một lần khi Flow 3 khởi động, nhưng chưa nạp model lên GPU."""
    global _PREFLIGHT_DONE, _PREFLIGHT_ERROR
    with _INIT_LOCK:
        if _PREFLIGHT_DONE:
            return LogoCleanupResult(not bool(_PREFLIGHT_ERROR), _PREFLIGHT_ERROR or "")
        _PREFLIGHT_DONE = True
        try:
            if not MODEL_PATH.is_file():
                raise FileNotFoundError(f"Thiếu model xóa logo: {MODEL_PATH}")
            # Import muộn để Flow 3 vẫn tạo/lưu ảnh bình thường nếu máy thiếu
            # bộ AI này. Lỗi sẽ được ghi vào Excel thay vì làm dừng cả flow.
            import cv2  # noqa: F401
            import numpy  # noqa: F401
            import torch
            from simple_lama_inpainting import SimpleLama  # noqa: F401

            if not torch.cuda.is_available():
                raise RuntimeError("Không nhận được GPU NVIDIA/CUDA")
            print("-> Bộ xóa logo đã sẵn sàng; sẽ nạp model khi có ảnh Gemini đầu tiên.")
        except Exception as exc:
            _PREFLIGHT_ERROR = _short_error(exc)
            print(f"⚠️ Xóa logo không sẵn sàng, sẽ giữ ảnh gốc: {_PREFLIGHT_ERROR}")
        return LogoCleanupResult(not bool(_PREFLIGHT_ERROR), _PREFLIGHT_ERROR or "")


def _build_mask(height, width, np, cv2):
    points = np.array(
        [[round(x * width), round(y * height)] for x, y in LOGO_POLY],
        dtype=np.int32,
    )
    mask = np.zeros((height, width), dtype=np.uint8)
    cv2.fillPoly(mask, [points], 255)
    if MASK_PADDING_PX:
        size = MASK_PADDING_PX * 2 + 1
        mask = cv2.dilate(mask, np.ones((size, size), np.uint8))
    return mask


def _remove_logo(image, lama, np, cv2, image_class):
    height, width = image.shape[:2]
    mask = _build_mask(height, width, np, cv2)
    ys, xs = np.where(mask > 0)
    x1 = max(0, int(xs.min()) - AI_CONTEXT_PX)
    y1 = max(0, int(ys.min()) - AI_CONTEXT_PX)
    x2 = min(width, int(xs.max()) + AI_CONTEXT_PX + 1)
    y2 = min(height, int(ys.max()) + AI_CONTEXT_PX + 1)

    crop_bgr = image[y1:y2, x1:x2]
    crop_rgb = cv2.cvtColor(crop_bgr, cv2.COLOR_BGR2RGB)
    crop_mask = mask[y1:y2, x1:x2]
    ai_result = np.asarray(
        lama(image_class.fromarray(crop_rgb), image_class.fromarray(crop_mask)),
        dtype=np.uint8,
    )
    ai_result = ai_result[: crop_rgb.shape[0], : crop_rgb.shape[1]]
    ai_bgr = cv2.cvtColor(ai_result, cv2.COLOR_RGB2BGR)

    alpha = crop_mask.astype(np.float32) / 255.0
    if FEATHER_PX > 0:
        alpha = cv2.GaussianBlur(alpha, (0, 0), FEATHER_PX)
    alpha = np.clip(alpha[..., None], 0.0, 1.0)
    merged = ai_bgr.astype(np.float32) * alpha + crop_bgr.astype(np.float32) * (1 - alpha)
    result = image.copy()
    result[y1:y2, x1:x2] = np.clip(merged, 0, 255).astype(np.uint8)
    return result


def remove_logo_from_file(image_path: str) -> LogoCleanupResult:
    """Ghi đè an toàn ảnh đã lưu. Chỉ vùng logo bị thay đổi; lỗi thì giữ ảnh cũ."""
    lama, init_error = _load_lama_once()
    if lama is None:
        return LogoCleanupResult(False, init_error)

    try:
        import cv2
        import numpy as np
        from PIL import Image

        path = Path(image_path)
        raw = np.fromfile(str(path), dtype=np.uint8)
        source = cv2.imdecode(raw, cv2.IMREAD_COLOR)
        if source is None:
            raise RuntimeError("Không đọc được file ảnh")

        # Hai worker có thể cùng hoàn tất ảnh. Khóa này bảo vệ VRAM và dùng
        # chung model thay vì nạp hai model GPU độc lập.
        with _GPU_LOCK:
            result = _remove_logo(source, lama, np, cv2, Image)

        ok, encoded = cv2.imencode(path.suffix.lower() or ".png", result)
        if not ok:
            raise RuntimeError("Không mã hóa được ảnh sau khi xóa logo")
        temp_path = path.with_name(f"{path.stem}.logo-tmp{path.suffix}")
        try:
            encoded.tofile(str(temp_path))
            os.replace(temp_path, path)
        finally:
            if temp_path.exists():
                temp_path.unlink(missing_ok=True)
        print(f"-> Đã xóa logo bằng AI: {path.name}")
        return LogoCleanupResult(True)
    except Exception as exc:
        error = _short_error(exc)
        print(f"⚠️ Không xóa được logo, giữ ảnh gốc: {error}")
        return LogoCleanupResult(False, error)
