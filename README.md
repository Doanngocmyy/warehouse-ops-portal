# Warehouse Ops Portal — SOP + Pre-inbound + Inbound + Outbound

Site tĩnh, chạy 100% client-side (không backend, không gửi data ra ngoài). Mở `index.html` bằng trình duyệt hoặc publish qua GitHub Pages.

## Cấu trúc

```
wh-ops-portal/
├── index.html          # toàn bộ site (nav 4 tab: Overview / Pre-inbound / Inbound / Outbound)
├── css/style.css
└── js/
    ├── utils.js           # clean text, Excel parsing, QR matrix + vẽ QR
    ├── pdf-engine.js       # pdf-lib helpers (font Helvetica-Bold có sẵn, không cần font CJK)
    ├── label-templates.js  # 2 layout: per-EAN label & mix carton-summary label
    ├── asn-single.js       # port "Single carton_Duplicate_New logic.ipynb"
    ├── asn-mix.js          # port "Label Automation - Mix Case.ipynb" + shrink/grow
    ├── flowcharts.js       # click-to-detail + xuất ảnh PNG cho 2 flowchart
    └── app.js              # nav + wiring upload/generate
```

## 1. Overview

Hai flowchart, click vào từng ô để xem mô tả, nút "Xuất ảnh" tải PNG:
- **SOP Kho — Inbound & Outbound**: dựng lại từ ảnh bạn gửi (bảng swimlane HTML/CSS).
- **SOP Express**: dựng từ file HTML bạn upload, giữ nguyên toạ độ SVG gốc, bổ sung thêm phần SLA/Control points/Abbreviations cho đồng bộ với bản kia (bản gốc chưa có).

## 2. Pre-inbound — ASN Label Generator

Sau khi so khớp với 3 file mẫu thật bạn gửi (`ASN_MIX_.pdf`, `ASN_NO NEED.pdf`, `test mix label.pdf`), site tách thành **2 công cụ** vì đây thực chất là 2 loại label khác nhau:

### ASN Single (theo từng EAN)
Port từ `Single carton_Duplicate_New logic.ipynb`. Sheet **"Single"**, cột cố định B=SKU, C=EAN, H=Qty/box, I=Carton qty, J=Carton start, K=Carton end, L=MIX marker. MIX chỉ gom carton no cho các dòng liền kề — mỗi dòng vẫn ra 1 label riêng (đúng logic gốc + đúng 2 file mẫu bạn gửi).

### ASN Mix (label tổng hợp theo Carton)
Port từ `Label Automation - Mix Case.ipynb`. Sheet **"Mix"**, cột cố định C=SKU/EAN, H=Qty, I=group marker (1 hoặc "MIX"), J=Carton start, K=Carton end, P=Carton ID base. 1 carton = 1 label liệt kê tất cả SKU bên trong; Carton ID in ra = cột P + "-" + tổng Qty.

**Auto shrink/grow (tính năng mới, theo yêu cầu):**
- Số dòng SKU ≤ ngưỡng (mặc định 5, chỉnh được trong UI): vẫn giữ trang 6x4in như mẫu, tự động shrink cỡ chữ vừa đủ để không tràn label.
- Số dòng SKU > ngưỡng: tự động chuyển sang trang A4 (cỡ chữ bình thường, không shrink) để vẫn đọc được rõ.
- Đã test với 2/4/8 dòng SKU — cả 3 case đều render đúng, không tràn/chồng chữ (xem log tạo label trong site để biết mode nào được áp dụng cho từng carton).

**Giả định cần bạn xác nhận lại:** ASN Mix hiện tạo CẢ label tổng hợp theo carton này. Nếu carton MIX cũng cần label riêng cho từng EAN (như ASN Single) để dán từng kiện hàng bên trong, báo lại để mình bổ sung thêm — hiện tại site chưa tự động sinh thêm bộ label per-EAN cho carton MIX.

## 3. Inbound

Để trống theo yêu cầu — placeholder chờ bạn viết hướng dẫn dùng FOMS.

## 4. Outbound

Link thẳng sang site VAS Label Generator đã publish trước đó (`https://doanngocmyy.github.io/vas-label-generator/`) — tách repo riêng theo đề xuất, không nhúng lại code để tránh trùng lặp / lệch phiên bản.

## Font & QR

Toàn bộ text trên label ASN là ASCII (SKU/EAN/Carton ID) nên dùng font Helvetica-Bold có sẵn trong pdf-lib — không cần nhúng font CJK như site VAS. QR code vẽ trực tiếp bằng thư viện `qrcode-generator` (client-side, không gọi API ngoài như Labelary).

## Việc còn lại trước khi publish

1. Xác nhận lại giả định ASN Mix ở trên.
2. Test với file Excel Single/Mix thật của bạn (site hiện mới test bằng data mẫu tự tạo).
3. Xác nhận có publish GitHub Pages ngay không, và tên repo muốn dùng.
