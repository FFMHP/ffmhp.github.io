# FFM Chuyển Đơn

Công cụ đổi đơn đặt hàng của khách (Kmart, Morgant, BRG/Fujimart, AEON) sang form **FORM_2_HP** của Hyperion.

**Link cho team:** https://ffmhp.github.io/

## Cập nhật bảng dữ liệu (không cần gửi file cho team)

Các bảng nằm trong thư mục [`data`](data):

| File | Dùng để |
|---|---|
| `DMKH.xlsx` | Dò mã khách hàng (cột **Mã KH** + **Tên nội bộ**) |
| `DMSP.xlsx` | Barcode → product code, kiểm tra barcode |
| `Code chuyển đổi.xlsx` | AEON: barcode trên đơn → product code điền form |
| `Form final.xlsx` | Khuôn form FORM_2_HP |

Cách thay bảng mới:
1. Mở thư mục **data** → bấm **Add file → Upload files**.
2. Kéo file Excel mới vào. **Tên file phải giống hệt tên cũ** (ví dụ `DMKH.xlsx`) để ghi đè.
3. Bấm **Commit changes**. Khoảng 1–2 phút sau, team mở lại link là dùng bảng mới.

Dưới tiêu đề trang web có dòng "Bảng dữ liệu · DMKH: cập nhật …" để biết đang dùng bảng ngày nào.
Nếu bảng lỗi (đổi tên cột…), trang tự dùng bảng gắn sẵn và ghi "lỗi, dùng bảng gắn sẵn".
