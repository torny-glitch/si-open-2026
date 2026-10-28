# MSB Pickleball SI Open 2026 — website

Website giải đấu (Vite + React). Hai file chính:

- `src/main.jsx` — nội dung, dữ liệu và các thành phần giao diện
- `src/style.css` — toàn bộ giao diện (bảng màu lấy từ poster)

## Cấu trúc thư mục (bắt buộc đúng)

```
si-open-2026/
├─ index.html
├─ package.json
├─ vite.config.js
└─ src/
   ├─ main.jsx
   ├─ style.css
   └─ assets/qr.png      ← mã QR đăng ký (thay file này nếu đổi form)
```

## Chạy thử

1. Cài Node.js (bản LTS) từ https://nodejs.org/
2. `npm install` (bắt buộc chạy lại vì có thêm gói phông chữ)
3. `npm run dev` rồi mở địa chỉ hiện trong terminal
4. `npm run build` để kiểm tra bản chính thức

## Cập nhật thông tin giải

Mở `src/main.jsx`, sửa khối `EVENT` ở đầu file: ngày, giờ, địa điểm, link đăng ký (`registerUrl`),
hạn chót (`registerDeadline`, để trống thì không hiển thị), số VĐV, số sân, số trận.
Điền `startISO` để bật đồng hồ đếm ngược (để `''` nếu muốn tắt).

Danh sách VĐV: sửa `DEFAULT_PLAYERS` (Trình A–D, Họ tên, Phòng ban, Vai trò, Đội). Đội 5 là đội khách mời (8 người, không thuộc 4 trình).
Số suất mỗi trình: `LEVEL_SLOTS` (10 = 8 nội bộ + 2 khách mời). Khách mời có Đội = `Đội 5`, phòng ban ghi “Khách mời”.

Lịch thi đấu: `SCHEDULE` gồm 40 trận, mỗi dòng là `[lượt, sân, đội, đội, cặp]`. Mỗi lượt 5 trận trên 5 sân,
đội nào cũng có 2 cặp thi đấu, Sân 1 luôn là cặp 1 (hạng 1–2). Mã trận M01–M40 = (lượt − 1) × 5 + sân.
Giờ, tỉ số, trạng thái nhập qua **Khu vực BTC** (liên kết ở chân trang) bằng file Excel hai sheet `Lịch thi đấu` và `Vận động viên`
(nút “Tải file mẫu” xuất sẵn dữ liệu hiện tại).
Ngoài 40 trận vòng tròn còn 8 trận tranh hạng (TB1–TB4 tranh giải Ba, CK1–CK4 chung kết, mỗi trận đủ 4 cặp trên Sân 1–4).
Bảng xếp hạng, bảng chéo và nhánh đấu tự tính từ cột `Tỉ số` (dạng `11-8`): hai đội gặp nhau 4 trận, thắng nhiều trận hơn thì thắng đối đầu,
nếu 2–2 thì xét tổng điểm ghi được; xếp hạng theo số lần thắng đối đầu, rồi số trận thắng, rồi hiệu số điểm.
Khi vòng tròn kết thúc, tên đội ở các trận tranh hạng tự điền theo bảng xếp hạng.

Dữ liệu nạp qua Khu vực BTC chỉ lưu trong trình duyệt của người nạp, không hiển thị cho khách khác.
Muốn công khai: sửa dữ liệu trong mã nguồn rồi triển khai lại, hoặc nối cơ sở dữ liệu.

## Triển khai Vercel

Đưa mã lên GitHub → Vercel → Add New Project → chọn repo. Framework: Vite, build `npm run build`, output `dist`.
