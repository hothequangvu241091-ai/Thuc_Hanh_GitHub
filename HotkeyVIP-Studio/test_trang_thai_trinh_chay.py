from __future__ import annotations

import time
from datetime import datetime


TOTAL_SECONDS = 20


def main() -> None:
    print("=" * 58, flush=True)
    print("TEST TRẠNG THÁI HOTKEYVIP STUDIO", flush=True)
    print(f"Bắt đầu: {datetime.now():%H:%M:%S}", flush=True)
    print("File này chỉ chờ và in tiến độ; không sửa dữ liệu.", flush=True)
    print("=" * 58, flush=True)

    for step in range(1, TOTAL_SECONDS + 1):
        percent = round(step / TOTAL_SECONDS * 100)
        print(
            f"[TIẾN ĐỘ] {step:02d}/{TOTAL_SECONDS} giây - {percent:3d}%",
            flush=True,
        )
        time.sleep(1)

    print("=" * 58, flush=True)
    print("TEST HOÀN THÀNH THÀNH CÔNG", flush=True)
    print(f"Kết thúc: {datetime.now():%H:%M:%S}", flush=True)


if __name__ == "__main__":
    main()
