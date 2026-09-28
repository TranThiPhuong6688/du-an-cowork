# HANDOFF – Dự án 原価計算 (Pallet xuất kho → VC/HQ)

## Quy tắc làm việc
- Chat chỉ dùng tiếng Việt. File: Times New Roman 12, hạn chế màu, không in nghiêng.
- HỎI TRƯỚC KHI LÀM; có nhiều phương án thì so sánh ưu/nhược; chưa chắc thì nói "chưa chắc".
- Luôn mở bản MỚI NHẤT của file trước khi sửa (đã từng mất phần user sửa). Chỉ sửa sheet cần thiết.
- GIỮ NGUYÊN tên file (link giữa các file): "Charge 成形・組立-V02.xlsx", "Tap hop chi phi ke toan-V03.xlsx".
- Macro VBA: chỉ dùng ký tự ASCII (không dấu) trong code. Import .bas từ đường dẫn có chữ Nhật/Việt bị lỗi → dùng file công cụ .xlsm hoặc copy code dán tay.
- Sửa Excel bằng openpyxl làm HỎNG external link: externalLink có 2 đường dẫn (tuyệt đối + tương đối) → Excel "repair" thành [RecoveredExternalLinkN]. Sau khi lưu, BẮT BUỘC chép lại nguyên xl/externalLinks/* (xml + _rels) từ file gốc; kiểm tra mọi r:id đều có trong rels.

## Đường dẫn
- Gốc: D:\D\管理データ\6.原価活動\1. 原価計算ファイル\
  - Data nguồn\ : Charge 成形・組立-V02.xlsx, Tap hop chi phi ke toan-V03.xlsx, Master sheet.xlsb, Gia thanh ma duc-V03.xlsm, Gia thanh ma lap rap-V05.xlsm, Cong cu Pallet.xlsm
  - Data hàng tháng\ : 完成品出庫xuat kho thanh pham.xlsx (sheet Data_T1…Data_T12), DEBIT TRUCKING KYOWA.xlsx

## File xuất kho (sheet Data_Tn)
- Dòng 1: ngày ở ô đầu mỗi khối 8 cột (D, L, T…); dòng 4: tên nhóm; dữ liệu từ dòng 5. Cột B = Part No, C = mã KYOWA.
- Mỗi ngày 8 cột: Tổng xuất kho, FBHP, KDTVN, OKI, DAIKIN, TOYOTA, EVA, KHÁC.
- Cuối sheet có dòng tổng / chênh lệch (B trống hoặc B là số âm) → bỏ qua.
- File tháng 6 có thêm khối ngày 1/7 → chỉ tính những ngày thuộc tháng.

## Quyết định đã chốt
- Pallet theo NGÀY, theo từng khách: tổng (SL ÷ SL/pallet) các mã trong ngày → phần nguyên + MIN(1, phần lẻ ÷ 0,7) (hệ số lấy ở VC_HQ!C5). Pallet tháng = Σ các ngày. Chỉ lưu kết quả THÁNG.
- Gán nhóm → khách: FBHP→FBHP, KDTVN→KDTVN, OKI→OKI DENKI, DAIKIN→DAIKIN, TOYOTA→TMV, EVA→EVA HP (CHƯA CHẮC: cột EVA có cả mã KDTVN/FBHP), KHÁC→"(Theo Master)" = tách theo khách của mã ở SL_Pallet cột B.
- SL/pallet = SL/thùng × Thùng/pallet (Master: sheet "Master sheet", dữ liệu từ dòng 7; cột J mã KYOWA, H khách, Y SL/thùng, AF thùng/pallet, AR quy cách SX).
- HẢI QUAN xuất: chia theo pallet xuất thực tế (HQ_Form cột D = XK_Pallet cột N, BQ/tháng; H = D).
- VẬN CHUYỂN: giữ theo pallet VC_DEBIT (VC_HQ không đổi).

## Charge 成形・組立-V02.xlsx (bản đã giao)
- Thêm 2 sheet sau HQ_Form:
  - XK_Pallet:
    - A. Tham số: C5 thư mục "..\Data hàng tháng", C6 tiền tố "Data_T", C7 tên sheet, C8 = 1 (tìm cả thư mục con), C9 = VC_HQ!C5, C10 = log.
    - B. Nhóm→khách: A14:B20.
    - C. Pallet: dòng 24–56 (khách = VC_HQ!A14:A46), 57 = chưa gán, 58 = tổng; B–M = T1–T12; N = BQ theo VC_HQ C7–C8; O = tổng.
    - D. SL: dòng 62–96, cùng bố cục.
    - E. Kiểm tra: dòng 100–107 (107 = OK/LỆCH).
    - F. Mã thiếu SL/pallet: từ dòng 111.
  - SL_Pallet: A:K theo cấu trúc Master_pallet, dữ liệu từ dòng 5; N:O = SL/pallet nhập tay (ưu tiên). Hiện TẠM lấy từ Gia thanh ma duc (5.764 mã, chưa có mã lắp ráp).
- HQ_Form: D31:D63 = INDEX(XK_Pallet!N24:N56, MATCH theo tên khách); H = N(D); cập nhật ghi chú A29.
- HD dòng 35–40: hướng dẫn.
- Tham khảo T6 (chỉ tính mã đã có quy cách): tổng SL 1.947.794; pallet 2.878,1 (KDTVN 1.047,0; FBHP 659,5; DAIKIN 630,4; OKI 315,2; EVA HP 197,7). Còn thiếu 188 mã / 292.577 cái, phần lớn là mã lắp ráp.

## Macro (Cong cu Pallet.xlsm có sẵn 2 module + 2 nút; Charge vẫn là .xlsx)
- CapNhatXuatKhoPallet (modXuatKhoPallet):
  - Tìm dữ liệu tháng m theo thứ tự: (1) sheet Data_Tm trong file đang mở; (2) file/thư mục tên Data_Tm trong C5; (3) sheet Data_Tm trong file có chữ "xuat kho" trong C5.
  - Không tìm thấy tháng nào → dừng, KHÔNG xóa số cũ.
  - Ghi kết quả vào XK_Pallet mục C–F; dòng 100 ghi "file | sheet" đã đọc.
- XuatMasterPallet (modXuatMasterPallet): chọn Master sheet.xlsb → xuất Master_pallet.xlsx (mọi mã, cả mã lắp ráp) → user dán vào SL_Pallet!A5.
- Trạng thái:
  - Mới kiểm tra biên dịch; CHƯA chạy thật trên Excel với dữ liệu thật.
  - Lần chạy đầu (bản cũ) không tìm thấy file và đã xóa T6 → cần chạy lại bằng bản mới.

## Việc tiếp theo
1. Chạy CapNhatXuatKhoPallet trên máy thật với sheet Data_T6 → kiểm tra dòng 107 = OK và dòng 100.
2. Chạy XuatMasterPallet → dán Master_pallet vào SL_Pallet → chạy lại (bổ sung mã lắp ráp thiếu).
3. Chốt cách gán nhóm EVA (EVA HP hay "(Theo Master)").
4. Kiểm tra VC_HQ / HQ_Form sau khi có pallet thực tế (USD/pallet theo khách).
