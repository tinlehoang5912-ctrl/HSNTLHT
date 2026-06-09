# Hồ sơ nghiệm thu — Newtecons / Bộ phận MEP

Phần mềm lập **Phiếu yêu cầu & Biên bản nghiệm thu vật liệu đầu vào**. Chạy hoàn toàn trên trình duyệt, không cần backend.

## Tính năng
- Landing page + tải logo riêng.
- Upload **form Word (.docx)** cho từng gói thầu; app tự dò trường biến đổi (chữ đỏ + `{{placeholder}}`) và tự nhận diện **bảng khối lượng**.
- Trình nhập bảng khối lượng dạng spreadsheet giống form: dán dải ô từ Excel, xuất/nhập file Excel mẫu.
- Xuất Word giữ nguyên 100% định dạng form gốc, chỉ điền phần biến đổi.

## Chạy thử trên máy
```bash
npm install
npm run dev
```
Mở địa chỉ hiện ra (mặc định http://localhost:5173).

## Build production
```bash
npm run build      # tạo thư mục dist/
npm run preview    # xem thử bản build
```

## Deploy

### Cách 1 — Vercel (khuyến nghị, giống các app trước)
1. Đẩy thư mục này lên một repo GitHub.
2. Vào vercel.com → New Project → chọn repo.
3. Framework preset: **Vite** (Vercel tự nhận). Build command `npm run build`, Output directory `dist`. Bấm Deploy.

> Không cần file `vercel.json`. Vite là SPA một trang nên Vercel xử lý tự động.

Hoặc dùng CLI:
```bash
npm i -g vercel
vercel        # deploy preview
vercel --prod # deploy chính thức
```

### Cách 2 — Netlify
- Build command: `npm run build` · Publish directory: `dist`.

### Cách 3 — Hosting tĩnh bất kỳ
Chạy `npm run build` rồi tải toàn bộ nội dung thư mục `dist/` lên hosting (GitHub Pages, Cloudflare Pages, hoặc server nội bộ).

## Ghi chú
- App nạp 3 thư viện qua CDN khi cần: JSZip (đọc/ghi .docx), SheetJS (Excel), docx. Cần có internet ở máy người dùng lần đầu để tải các thư viện này.
- Dữ liệu nhập hiện lưu trong phiên làm việc của trình duyệt (chưa đồng bộ cloud).
