' ============================================================================
' QUẢN LÝ QUÉT XE - VBA MACRO
' ============================================================================
' Cách sử dụng:
' 1. Mở file Excel Quan_Ly_Quet_Xe.xlsx
' 2. Bấm Alt+F11 mở VBA Editor
' 3. Chọn Insert → Module
' 4. Copy toàn bộ code này vào Module
' 5. Đóng VBA Editor, lưu file dưới dạng .xlsm
' 6. Kích hoạt macro khi mở file
' ============================================================================

' ===== BIẾN TOÀN CỤC =====
Dim TimerHandle As Long
Dim LastAutoSaveTime As Date

' ===== KÍCH HOẠT MACRO KHI MỞ FILE =====
Private Sub Workbook_Open()
    ' Cấu hình bảo vệ sheet
    Call SetupSheetProtection

    ' Khởi động Auto Save
    Call StartAutoSave

    MsgBox "✓ Quản Lý Quét Xe đã sẵn sàng!" & vbCrLf & vbCrLf & _
            "- Nhập mã xe ở cột A" & vbCrLf & _
            "- Cột B tự ghi thời gian" & vbCrLf & _
            "- Nhập 'bẩn'/'ban'/'ko sạch'/'không sạch' ở cột E" & vbCrLf & _
            "- Nhập 'X' ở cột F để xóa dòng" & vbCrLf & _
            "- Dùng Alt+F8 → CapNhatTongHop_Ngay để tổng hợp", vbInformation, "Thông Báo"
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ' Hủy Auto Save khi đóng file
    On Error Resume Next
    Application.OnTime TimerHandle, , False
    On Error GoTo 0

    ' Lưu file trước khi đóng
    ThisWorkbook.Save
End Sub

' ===== SETUP BẢO VỆ SHEET =====
Private Sub SetupSheetProtection()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("QuetXe")

    ' Bỏ bảo vệ sheet để cài đặt
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0

    ' Khóa cột A, B, D (chỉ cho phép chỉnh sửa C, E, F)
    ' Tạm thời bỏ bảo vệ để dễ làm việc - có thể bật lại sau
End Sub

' ===== EVENT CHANGE - XỬ LÝ NHẬP DỮ LIỆU =====
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Dim targetRow As Long
    Dim maCo As String
    Dim colA As Range

    ' Không xử lý dòng tiêu đề
    If Target.Row = 1 Then Exit Sub

    ' Không xử lý nếu nhập nhiều ô cùng lúc
    If Target.Cells.Count > 1 Then Exit Sub

    Set ws = Target.Worksheet
    targetRow = Target.Row

    ' Tắt event để tránh vòng lặp
    Application.EnableEvents = False
    On Error GoTo ErrorHandler

    Select Case Target.Column
        Case 1 ' CỘT A - MÃ XE
            Call XuLyCotA(ws, targetRow, Target.Value)

        Case 2 ' CỘT B - THỜI GIAN QUÉT
            ' Không xử lý - do user không được sửa

        Case 3 ' CỘT C - THỜI GIAN QUAY VÒNG (tính tự động)
            ' Do macro tính tự động

        Case 5 ' CỘT E - GHI CHÚ
            ' Không cần xử lý đặc biệt

        Case 6 ' CỘT F - XÓA DÒNG
            Call XuLyCotF(ws, targetRow, Target.Value)
    End Select

ErrorHandler:
    Application.EnableEvents = True
End Sub

' ===== XỬ LÝ CỘT A (MÃ XE) =====
Private Sub XuLyCotA(ws As Worksheet, row As Long, maCo As String)
    Dim colBCell As Range
    Dim colCCell As Range
    Dim colDCell As Range
    Dim colECell As Range
    Dim colFCell As Range
    Dim prevRow As Long
    Dim prevTime As Date
    Dim timeDiff As Double
    Dim hours As Long
    Dim minutes As Long
    Dim borderStyle As Border
    Dim borderSide As XlBordersIndex
    Dim thin As XlLineStyle

    ' Nếu ô A trống → thoát
    If Trim(maCo) = "" Then Exit Sub

    ' Lấy reference ô B, C, D, E, F cùng dòng
    Set colBCell = ws.Cells(row, 2)
    Set colCCell = ws.Cells(row, 3)
    Set colDCell = ws.Cells(row, 4)
    Set colECell = ws.Cells(row, 5)
    Set colFCell = ws.Cells(row, 6)

    ' 1. THÊM THỜI GIAN QUÉT VÀO CỘT B (tự động bây giờ)
    If colBCell.Value = "" Then
        colBCell.Value = Now()
        colBCell.NumberFormat = "dd/mm/yyyy hh:mm:ss"
    End If

    ' 2. TÌM LẦN QUÉT TRƯỚC CỦA CHIẾC XE NÀY
    prevRow = FindPreviousScan(ws, row, maCo)

    If prevRow > 0 Then
        ' Có lần quét trước
        prevTime = ws.Cells(prevRow, 2).Value

        ' Tính thời gian quay vòng
        timeDiff = (colBCell.Value - prevTime) * 24 ' Chuyển sang giờ
        hours = Int(timeDiff)
        minutes = Int((timeDiff - hours) * 60)

        ' Ghi vào cột C (dạng giờ:phút)
        colCCell.Value = hours & "H " & Format(minutes, "00") & "M"

        ' Ghi vào cột D (dạng số giờ)
        colDCell.Value = Format(timeDiff, "0.0") & " H"
        colDCell.NumberFormat = "@"

        ' Nếu quay vòng >= 16 giờ → tô vàng
        If timeDiff >= 16 Then
            colCCell.Interior.Color = RGB(255, 255, 0) ' Vàng
            colDCell.Interior.Color = RGB(255, 255, 0)
        Else
            colCCell.Interior.Color = RGB(255, 255, 255) ' Trắng
            colDCell.Interior.Color = RGB(255, 255, 255)
        End If
    End If

    ' 3. TẠO BORDER CHO DÒNG A:F
    Call AddBorderToRow(ws, row, 1, 6)

    ' 4. CHUYỂN VỀ ÔC A TRỐNG TIẾP THEO (TÙY CHỌN)
    ' Bỏ comment nếu muốn chức năng này
    ' Call MoveToNextEmptyA(ws, row)
End Sub

' ===== TÌM LẦN QUÉT TRƯỚC =====
Private Function FindPreviousScan(ws As Worksheet, currentRow As Long, maCo As String) As Long
    Dim row As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Tìm từ dòng trên xuống (từ dòng currentRow - 1 về dòng 2)
    For row = currentRow - 1 To 2 Step -1
        If Trim(ws.Cells(row, 1).Value) = Trim(maCo) Then
            ' Tìm thấy lần quét trước
            FindPreviousScan = row
            Exit Function
        End If
    Next row

    FindPreviousScan = 0 ' Không tìm thấy
End Function

' ===== XỬ LÝ CỘT F (XÓA DÒNG) =====
Private Sub XuLyCotF(ws As Worksheet, row As Long, value As String)
    Dim deleteConfirm As VbMsgBoxResult

    ' Nếu nhập X (hoa hoặc thường)
    If Trim(UCase(value)) = "X" Then
        ' Xác nhận xóa
        deleteConfirm = MsgBox("Xóa dòng " & row & "? (Mã xe: " & ws.Cells(row, 1).Value & ")", _
                                vbYesNo + vbQuestion, "Xác Nhận Xóa")

        If deleteConfirm = vbYes Then
            ' Xóa dữ liệu A:F
            ws.Range("A" & row & ":F" & row).Delete xlShiftUp
        Else
            ' Xóa ô F nếu user không đồng ý
            ws.Cells(row, 6).Clear
        End If
    End If
End Sub

' ===== THÊM BORDER CHO DÒNG =====
Private Sub AddBorderToRow(ws As Worksheet, row As Long, startCol As Long, endCol As Long)
    Dim rng As Range
    Dim i As Long
    Dim borderStyle As Border

    For i = startCol To endCol
        Set rng = ws.Cells(row, i)
        With rng.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = 0
        End With
    Next i
End Sub

' ===== MACRO TỔNG HỢP LỊCH SỬ BẨN =====
Sub CapNhatTongHop_Ngay()
    Dim wsQuet As Worksheet
    Dim wsTongHop As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim maCo As String
    Dim ghiChu As String
    Dim ngayBatDau As Date
    Dim ngayVeSinh As Date
    Dim soNgay As Long
    Dim lanBan As Long
    Dim tongNgayBan As Long
    Dim tongLanBan As Long
    Dim trungBinhNgay As Double

    Dim dictXe As Object ' Dùng để lưu trữ số lần bẩn mỗi xe
    Set dictXe = CreateObject("Scripting.Dictionary")

    Set wsQuet = ThisWorkbook.Sheets("QuetXe")

    ' Tạo sheet TongHop nếu chưa tồn tại
    On Error Resume Next
    Set wsTongHop = ThisWorkbook.Sheets("TongHop")
    If wsTongHop Is Nothing Then
        Set wsTongHop = ThisWorkbook.Sheets.Add
        wsTongHop.Name = "TongHop"
        ' Thêm tiêu đề
        With wsTongHop
            .Range("A1:F1").Value = Array("Số Xe", "Lần Bẩn", "Ngày Bắt Đầu", "Ngày Bẩn Vệ Sinh", "Số Ngày", "Trung Bình")
            .Range("A1:F1").Font.Bold = True
            .Range("A1:F1").Interior.Color = RGB(54, 96, 146)
            .Range("A1:F1").Font.Color = RGB(255, 255, 255)
        End With
    End If
    On Error GoTo 0

    ' Xóa dữ liệu cũ trong TongHop (giữ tiêu đề)
    wsTongHop.Range("A2:F" & wsTongHop.Cells(wsTongHop.Rows.Count, 1).End(xlUp).Row).Delete

    ' Duyệt qua tất cả dòng trong sheet QuetXe
    lastRow = wsQuet.Cells(wsQuet.Rows.Count, 1).End(xlUp).Row

    Dim outputRow As Long
    outputRow = 2

    For row = 2 To lastRow
        maCo = Trim(wsQuet.Cells(row, 1).Value)
        ghiChu = Trim(wsQuet.Cells(row, 5).Value)

        ' Kiểm tra từ khóa bẩn
        If IsKeywordBan(ghiChu) Then
            ' Nếu xe chưa có trong từ điển, thêm vào
            If Not dictXe.Exists(maCo) Then
                dictXe.Add maCo, 0
            End If

            ' Tăng số lần bẩn
            dictXe(maCo) = dictXe(maCo) + 1
            lanBan = dictXe(maCo)

            ' Lấy ngày bắt đầu (lần bẩn trước)
            ngayBatDau = FindPreviousDirtyDate(wsQuet, row, maCo)
            If ngayBatDau = 0 Then
                ngayBatDau = wsQuet.Cells(row, 2).Value ' Lần đầu bẩn
            End If

            ' Ngày bẩn vệ sinh (ngày hiện tại)
            ngayVeSinh = wsQuet.Cells(row, 2).Value

            ' Tính số ngày quay vòng
            soNgay = DateDiff("d", ngayBatDau, ngayVeSinh)

            ' Thêm vào TongHop
            With wsTongHop
                .Cells(outputRow, 1).Value = maCo
                .Cells(outputRow, 2).Value = lanBan
                .Cells(outputRow, 3).Value = Format(ngayBatDau, "dd/mm/yyyy")
                .Cells(outputRow, 4).Value = Format(ngayVeSinh, "dd/mm/yyyy")
                .Cells(outputRow, 5).Value = soNgay & " ngày"

                ' Chỉ hiển thị trung bình ở dòng đầu của mỗi xe
                If lanBan = 1 Then
                    ' Tính trung bình: tổng ngày / tổng lần bẩn
                    trungBinhNgay = CalculateAverageForCar(wsTongHop, maCo)
                    .Cells(outputRow, 6).Value = Format(trungBinhNgay, "0.0") & " ngày"
                End If

                ' Thêm border
                .Range("A" & outputRow & ":F" & outputRow).Borders.LineStyle = xlContinuous
            End With

            outputRow = outputRow + 1
        End If
    Next row

    ' Chuyển sang sheet TongHop
    wsTongHop.Activate

    MsgBox "✓ Cập nhật tổng hợp thành công! " & vbCrLf & _
            "Tổng số xe bẩn: " & dictXe.Count & " chiếc", vbInformation, "Hoàn Thành"
End Sub

' ===== KIỂM TRA TỪ KHÓA BẨN =====
Private Function IsKeywordBan(ghiChu As String) As Boolean
    Dim keywords As Variant
    Dim i As Long
    Dim lowerGhiChu As String

    keywords = Array("bẩn", "ban", "ko sạch", "không sạch", "không sạch", "ko sạch")
    lowerGhiChu = LCase(Trim(ghiChu))

    For i = LBound(keywords) To UBound(keywords)
        If InStr(1, lowerGhiChu, LCase(keywords(i)), vbTextCompare) > 0 Then
            IsKeywordBan = True
            Exit Function
        End If
    Next i

    IsKeywordBan = False
End Function

' ===== TÌM NGÀY BẨN TRƯỚC =====
Private Function FindPreviousDirtyDate(ws As Worksheet, currentRow As Long, maCo As String) As Date
    Dim row As Long
    Dim ghiChu As String

    For row = currentRow - 1 To 2 Step -1
        If Trim(ws.Cells(row, 1).Value) = Trim(maCo) Then
            ghiChu = Trim(ws.Cells(row, 5).Value)
            If IsKeywordBan(ghiChu) Then
                FindPreviousDirtyDate = ws.Cells(row, 2).Value
                Exit Function
            End If
        End If
    Next row

    FindPreviousDirtyDate = 0
End Function

' ===== TÍNH TRUNG BÌNH NGÀY BẨN =====
Private Function CalculateAverageForCar(ws As Worksheet, maCo As String) As Double
    Dim row As Long
    Dim lastRow As Long
    Dim totalDays As Long
    Dim totalCount As Long
    Dim lastRowCar As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    totalDays = 0
    totalCount = 0
    lastRowCar = 0

    For row = 2 To lastRow
        If Trim(ws.Cells(row, 1).Value) = Trim(maCo) Then
            totalCount = totalCount + 1
            lastRowCar = row

            ' Lấy số ngày từ cột E
            Dim ngayStr As String
            ngayStr = Trim(ws.Cells(row, 5).Value)
            Dim soNgay As Long
            soNgay = Val(Left(ngayStr, InStr(ngayStr, " ") - 1))
            totalDays = totalDays + soNgay
        End If
    Next row

    If totalCount > 0 Then
        CalculateAverageForCar = totalDays / totalCount
    Else
        CalculateAverageForCar = 0
    End If
End Function

' ===== AUTO SAVE (MỖI 30 PHÚT) =====
Private Sub StartAutoSave()
    Dim backupFolder As String
    Dim filePath As String
    Dim fileName As String
    Dim backupPath As String

    ' Tạo Backup folder
    filePath = ThisWorkbook.FullName
    fileName = Dir(filePath)
    fileName = Left(fileName, InStr(fileName, ".") - 1)

    backupFolder = ThisWorkbook.Path & "\Backup_60Phut"
    If Dir(backupFolder, vbDirectory) = "" Then
        MkDir backupFolder
    End If

    ' Lưu backup
    backupPath = backupFolder & "\" & fileName & "_BACKUP.xlsm"
    ThisWorkbook.SaveCopyAs backupPath

    ' Hiển thị thời gian lưu tiếp theo
    Application.StatusBar = "Auto Save: Lần tiếp theo lúc " & Format(Now() + TimeValue("00:30:00"), "hh:mm:ss")

    ' Thiết lập auto save mỗi 30 phút
    TimerHandle = Now() + TimeValue("00:30:00")
    Application.OnTime TimerHandle, "AutoSaveBackup"
End Sub

Private Sub AutoSaveBackup()
    Dim backupFolder As String
    Dim fileName As String
    Dim backupPath As String

    On Error Resume Next

    fileName = Left(Dir(ThisWorkbook.FullName), InStr(Dir(ThisWorkbook.FullName), ".") - 1)
    backupFolder = ThisWorkbook.Path & "\Backup_60Phut"

    If Dir(backupFolder, vbDirectory) = "" Then
        MkDir backupFolder
    End If

    backupPath = backupFolder & "\" & fileName & "_BACKUP.xlsm"

    ' Xóa backup cũ nếu tồn tại
    If Dir(backupPath) <> "" Then
        Kill backupPath
    End If

    ' Tạo backup mới
    ThisWorkbook.SaveCopyAs backupPath

    ' Lưu file chính
    ThisWorkbook.Save

    ' Cập nhật status bar
    Application.StatusBar = "✓ Backup: " & Format(Now(), "hh:mm:ss") & " | Lần tiếp theo: " & Format(Now() + TimeValue("00:30:00"), "hh:mm:ss")

    ' Thiết lập auto save tiếp theo
    TimerHandle = Now() + TimeValue("00:30:00")
    Application.OnTime TimerHandle, "AutoSaveBackup"

    On Error GoTo 0
End Sub

' ===== CHUYỂN SANG ÔA TRỐNG TIẾP THEO (TÙY CHỌN) =====
Private Sub MoveToNextEmptyA(ws As Worksheet, currentRow As Long)
    Dim row As Long
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Tìm ô A trống tiếp theo
    For row = currentRow + 1 To lastRow + 1
        If Trim(ws.Cells(row, 1).Value) = "" Then
            ws.Cells(row, 1).Select
            Exit Sub
        End If
    Next row

    ' Nếu không tìm thấy, chọn ô A tiếp theo
    ws.Cells(lastRow + 1, 1).Select
End Sub
