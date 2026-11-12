# FFM Chuyển Đơn

Công cụ đổi đơn đặt hàng của khách (Kmart, Morgant, BRG/Fujimart, AEON, Sói Biển, An Nhiên) sang form **FORM_2_HP** của Hyperion.

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

## Thống kê sử dụng

Mỗi lần chuyển đổi / tải file, web ghi 1 dòng (người làm, account, số file, số đơn, đơn tải được / còn lỗi) vào Google Sheet **"FFM Chuyển Đơn - Thống kê"** trên Drive của anh Tài (sheet `Log`). Không gửi nội dung đơn hàng.

## Code chuyển đổi AEON do team tự thêm

Khi đơn AEON có barcode chưa có trong bảng, người làm chọn hoặc nhập product code là **tự động lưu** vào bảng chung.
Mã được lưu vào sheet **`Code AEON`** trong Google Sheet "FFM Chuyển Đơn - Thống kê" (có tên người thêm và thời gian), cả team dùng được từ lần mở web sau.
- Chỉ lưu được product code có trong DMSP.
- File `data/Code chuyển đổi.xlsx` luôn được ưu tiên hơn bảng team thêm.
- Muốn sửa hoặc xoá một mã team thêm sai: sửa hoặc xoá dòng đó trong sheet `Code AEON`.

## Mã khách hàng do team tự thêm

- Ở ô **Mã khách hàng (C10)** của một đơn, chọn tay cửa hàng khác với web dò ra thì hiện nút **"Ghi nhớ: ... là ..."**; bấm nút mới lưu (tên cửa hàng trên đơn → mã KH). Lần sau gặp đúng tên đó web tự chọn.
- Hoặc mở **"+ Thêm mã khách hàng"** dưới ô Account, nhập Mã KH + tên cửa hàng như ghi trên đơn (AEON: tên bắt đầu bằng mã cửa hàng, vd "1006 - AEON Hải Phòng").
- Lưu ở sheet **`Mã KH`** trong Google Sheet "FFM Chuyển Đơn - Thống kê". Sửa / xoá dòng sai ngay trong sheet.
