================================================================================
QUẢN LÝ QUÉT XE QR - TỔNG HỢP
================================================================================

📦 CÓ 3 FILE:

1. Quan_Ly_Quet_Xe.xlsx
   → File Excel template (cấu trúc sẵn)
   → Mở bằng Excel trên Windows

2. VBA_Code_QuetXe.bas
   → Toàn bộ code VBA macro
   → Copy-paste vào Excel (xem hướng dẫn)

3. HƯỚNG_DẪN_SỬ_DỤNG.txt
   → Hướng dẫn chi tiết từng bước


⚡ QUICK START (3 BƯỚC):
================================================================================

BƯỚC 1: Trên Windows, mở Excel
   • File → Open → Quan_Ly_Quet_Xe.xlsx

BƯỚC 2: Nhập code VBA
   • Alt+F11 → Insert Module
   • Mở VBA_Code_QuetXe.bas (Notepad) → Copy toàn bộ → Paste vào Module
   • Đóng VBA Editor (Alt+F4)

BƯỚC 3: Lưu file
   • File → Save As
   • Chọn định dạng: Excel Macro-Enabled Workbook (*.xlsm)
   • Bấm Save

✅ XONG! Dùng file .xlsm từ bây giờ


🔧 TÍNH NĂNG CHÍNH:
================================================================================

✓ Thêm mã xe ở cột A
   → Tự động ghi thời gian quét (cột B)
   → Tự động tính thời gian quay vòng (cột C, D)
   → Tô vàng nếu quay vòng >= 16 giờ

✓ Nhập "bẩn/ban/ko sạch/không sạch" ở cột E
   → Chạy macro CapNhatTongHop_Ngay (Alt+F8)
   → Tự động tóm tắt lịch sử bẩn vào Sheet TongHop

✓ Nhập "X" ở cột F
   → Macro hỏi xác nhận
   → Xóa dòng A:F + border

✓ Auto Save mỗi 30 phút
   → Tạo folder Backup_60Phut
   → Backup file tự động


📋 SHEET QuetXe:
================================================================================

| A      | B                | C            | D       | E    | F   |
|--------|------------------|--------------|---------|------|-----|
| Mã Xe  | Thời Gian Quét   | Quay Vòng     | Giờ     | Ghi  | Xóa |
| A001   | 29/09/2026 08:00 | 34H 00M      | 34.0 H  | bẩn  | X   |
| B002   | 29/09/2026 08:15 | 2H 30M       | 2.5 H   |      |     |


📊 SHEET TongHop (Tự tạo bằng macro CapNhatTongHop_Ngay):
================================================================================

| A    | B         | C          | D          | E        | F         |
|------|-----------|------------|------------|----------|-----------|
| Xe   | Lần Bẩn   | Ngày BĐ    | Ngày Vệ S  | Số Ngày  | Trung B   |
| A001 | 1         | 25/09/2026 | 29/09/2026 | 4 ngày   | 4.0 ngày  |
| A001 | 2         | 29/09/2026 | 03/10/2026 | 4 ngày   |           |


❓ CÂU HỎI THƯỜNG GẶP:
================================================================================

Q: Macro không chạy?
A: 1. Bật Enable Macros khi mở file
   2. Kiểm tra code đã paste đầy đủ chưa
   3. Lưu file dưới dạng .xlsm

Q: Thời gian quay vòng không tính?
A: 1. Kiểm tra cột B có định dạng thời gian không
   2. Quét lại cùng mã xe (phải có 2 lần quét)

Q: Muốn tính auto-save khác 30 phút?
A: Mở VBA → Tìm "OnTime TimerHandle, , False" → Sửa "00:30:00" thành giá trị khác

Q: Backup được tạo ở đâu?
A: Thư mục Backup_60Phut (cạnh file Excel)


📞 SUPPORT:
================================================================================

Đọc file: HƯỚNG_DẪN_SỬ_DỤNG.txt (chi tiết)

Vấn đề nào?
- Lỗi macro → kiểm tra Enable Macros
- Tính toán sai → re-scan mã xe
- Auto Save → kiểm tra .xlsm + Enable Macros

================================================================================
Version: 1.0 | Date: 29/09/2026
================================================================================
