import ctypes


ctypes.windll.user32.MessageBoxW(
    None,
    "Nút Chạy đã mở file Python thành công.",
    "TEST THÀNH CÔNG",
    0x40,
)
