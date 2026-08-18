# Quy trình quản lý phiên bản

## Trước khi sửa

1. Chạy `git status --short` để biết code đang sạch hay còn thay đổi chưa lưu.
2. Nếu đang sạch, bắt đầu sửa.
3. Nếu đang có thay đổi, tạo mốc trước hoặc xác định rõ thay đổi thuộc đợt nào.

## Sau mỗi đợt sửa chạy được

1. Kiểm tra cú pháp và test liên quan.
2. Tăng version app khi hành vi người dùng hoặc cấu trúc chung thay đổi.
3. Tăng engine version của flow nếu thay đổi nằm trong một flow.
4. Ghi `CHANGELOG.md`: ngày, file/flow, lỗi cũ, cách sửa và cách kiểm tra.
5. Chạy `TAO_MOC_CODE.cmd`, nhập mô tả ngắn và tag phiên bản.

Ví dụ tag:

```text
app-v1.5.1
flow07-v2.4
```

## Xem và phục hồi

Xem các mốc:

```powershell
git log --oneline --decorate --all
git tag --list
```

Xem một file đã thay đổi gì:

```powershell
git diff app-v1.5.1 -- app_flows/07_bai_viet_lien_quan.py
```

Không tự dùng `git reset --hard`. Khi cần quay lại, hãy tạo nhánh phục hồi hoặc
nhờ Codex so sánh tag rồi phục hồi đúng file để không làm mất thay đổi mới.

## Phạm vi Git

Git theo dõi source code, CMD, tài liệu và các bản code lưu trữ. Git không theo
dõi `_runtime`, `outputs`, log và cache Python vì đó là dữ liệu sinh ra khi chạy.

