# Loi GUI bi tre khi keo cua so

## Ten loi de tra cuu

**Window dragging lag/jank khi bat "Show window contents while dragging"**.

Ten ky thuat gan dung: **DWM/GDI repaint bottleneck** (Windows Desktop
Window Manager phai ve lai toan bo noi dung cua so o moi buoc di chuyen).

## Dau hieu

- App van xu ly binh thuong; chi thao tac keo thanh tieu de/cua so bi delay.
- Notepad++ co the bi tuong tu; cac app dung pipeline GPU khac co the khong bi.
- Tat tuy chon `Show window contents while dragging` thi khung cua so keo
  muot ngay lap tuc.
- Restart app, xoa session hoac mo lai file khong giai quyet; restart Windows
  co the tam thoi xoa trang thai DWM/driver.

## Giai thich

Khi tuy chon nay bat, Windows phai repaint cua so Win32/GDI va ghep lai qua
DWM o moi frame keo chuot. Tkinter va Notepad++ dung thanh phan giao dien
Win32/GDI nen co the cung bi anh huong. Day khong phai dau hieu flow hay luong
Excel bi treo.

## Xu ly an toan

Giu tuy chon nay o trang thai **tat**. Khi keo, Windows chi hien khung vien;
khi tha chuot, cua so duoc ve lai binh thuong. Khong anh huong du lieu, chuc
nang app, luu file hay do on dinh.

Neu can dieu tra them, thu tat NVIDIA Overlay/Discord Overlay, kiem tra
refresh rate/DPI giua cac man hinh, va cap nhat/cai sach driver NVIDIA. Khong
can viet lai GUI chi vi loi keo cua so nay.

## Tu khoa

`window dragging lag`, `dragging jank`, `Show window contents while dragging`,
`DWM`, `GDI`, `repaint`, `Notepad++`, `Tkinter`, `Desktop Window Manager`
