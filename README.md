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

Danh sách VĐV: sửa `DEFAULT_PLAYERS` (cột Trình A–E, Họ tên, Phòng ban, Đội).
Lịch thi đấu: mã trận M01–M50 = (lượt − 1) × 5 + sân. Điền tên đội, giờ, tỉ số vào `DEFAULT_MATCHES`
hoặc dùng **Khu vực BTC** (liên kết ở chân trang) để nạp file Excel hai sheet `Lịch thi đấu` và `Vận động viên`.
Bảng xếp hạng, bảng chéo và nhánh đấu tự tính từ cột `Tỉ số` (dạng `11-8`).

Dữ liệu nạp qua Khu vực BTC chỉ lưu trong trình duyệt của người nạp, không hiển thị cho khách khác.
Muốn công khai: sửa dữ liệu trong mã nguồn rồi triển khai lại, hoặc nối cơ sở dữ liệu.

## Triển khai Vercel

Đưa mã lên GitHub → Vercel → Add New Project → chọn repo. Framework: Vite, build `npm run build`, output `dist`.
