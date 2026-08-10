import os
import subprocess


EDGE_PATHS = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]

edge_path = next((path for path in EDGE_PATHS if os.path.exists(path)), None)
if not edge_path:
    raise FileNotFoundError("Không tìm thấy Microsoft Edge.")

while True:
    try:
        worker_id = int(input("Chọn profile Worker (1-5): "))
        if 1 <= worker_id <= 5:
            break
    except ValueError:
        pass
    print("Chỉ được chọn số từ 1 đến 5.")

profile_path = (
    rf"D:\CodexProjects\Hotkeyvip\02_viet_bai"
    rf"\du_lieu_3_workers\profiles\worker_{worker_id}"
)

os.makedirs(profile_path, exist_ok=True)

subprocess.Popen([
    edge_path,
    f"--user-data-dir={profile_path}",
    "https://gemini.google.com/app",
])

print(f"Đã mở profile Worker {worker_id}:")
print(profile_path)
