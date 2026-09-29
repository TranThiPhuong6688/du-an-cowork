' ============================================================================
' QUẢN LÝ QUÉT XE - VBA MACRO V2 (ĐƠN GIẢN - CÓ BUTTON)
' ============================================================================
' Cách sử dụng:
' 1. Mở file Excel Quan_Ly_Quet_Xe.xlsx
' 2. Bấm Alt+F11 → Insert → Module
' 3. Copy toàn bộ code này vào Module
' 4. Đóng VBA Editor
' 5. Tạo Button "Thêm Xe" (Insert → Button)
' 6. Gán macro "ThemXe" cho button
' 7. Save file dưới dạng .xlsm
' ============================================================================

' ===== MACRO THÊM XE / QUÉT XE =====
Sub ThemXe()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim maCo As String
    Dim colBCell As Range
    Dim colCCell As Range
    Dim colDCell As Range
    Dim prevRow As Long
    Dim prevTime As Date
    Dim timeDiff As Double
    Dim hours As Long
    Dim minutes As Long

    Set ws = ThisWorkbook.Sheets("QuetXe")

    ' Tìm dòng cuối cùng có dữ liệu
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Nếu dòng cuối là tiêu đề (row 1), bắt đầu từ dòng 2
    If lastRow = 1 Then
        currentRow = 2
    Else
        ' Tìm dòng trống tiếp theo
        currentRow = lastRow + 1
    End If

    ' Yêu cầu nhập mã xe
    maCo = InputBox("Nhập mã xe (VD: A001, B002):", "Thêm Xe Mới", "")

    ' Nếu không nhập, thoát
    If Trim(maCo) = "" Then
        MsgBox "Hủy thêm xe", vbExclamation
        Exit Sub
    End If

    ' Kiểm tra mã xe đã tồn tại chưa (để tìm lần quét trước)
    prevRow = FindPreviousScan(ws, currentRow, maCo)

    ' Ghi mã xe vào cột A
    ws.Cells(currentRow, 1).Value = maCo
    ws.Cells(currentRow, 1).NumberFormat = "@"

    ' Ghi thời gian quét hiện tại vào cột B
    Set colBCell = ws.Cells(currentRow, 2)
    colBCell.Value = Now()
    colBCell.NumberFormat = "dd/mm/yyyy hh:mm:ss"

    ' Nếu có lần quét trước, tính thời gian quay vòng
    Set colCCell = ws.Cells(currentRow, 3)
    Set colDCell = ws.Cells(currentRow, 4)

    If prevRow > 0 Then
        ' Lấy thời gian quét trước
        prevTime = ws.Cells(prevRow, 2).Value

        ' Tính thời gian quay vòng (giờ)
        timeDiff = (colBCell.Value - prevTime) * 24

        ' Tách giờ và phút
        hours = Int(timeDiff)
        minutes = Int((timeDiff - hours) * 60)

        ' Ghi vào cột C (dạng giờ:phút)
        colCCell.Value = hours & "H " & Format(minutes, "00") & "M"
        colCCell.NumberFormat = "@"

        ' Ghi vào cột D (dạng số giờ)
        colDCell.Value = Format(timeDiff, "0.0") & " H"
        colDCell.NumberFormat = "@"

        ' Tô vàng nếu quay vòng >= 16 giờ
        If timeDiff >= 16 Then
            colCCell.Interior.Color = RGB(255, 255, 0) ' Vàng
            colDCell.Interior.Color = RGB(255, 255, 0)
        End If
    End If

    ' Tạo border cho dòng
    Call AddBorderToRow(ws, currentRow, 1, 6)

    ' Thông báo thành công
    MsgBox "✓ Thêm xe " & maCo & " thành công!" & vbCrLf & _
            "Dòng: " & currentRow, vbInformation, "Thành Công"
End Sub

' ===== TÌM LẦN QUÉT TRƯỚC =====
Private Function FindPreviousScan(ws As Worksheet, currentRow As Long, maCo As String) As Long
    Dim row As Long

    ' Tìm từ dòng trên xuống (từ dòng currentRow - 1 về dòng 2)
    For row = currentRow - 1 To 2 Step -1
        If Trim(ws.Cells(row, 1).Value) = Trim(maCo) Then
            FindPreviousScan = row
            Exit Function
        End If
    Next row

    FindPreviousScan = 0
End Function

' ===== THÊM BORDER CHO DÒNG =====
Private Sub AddBorderToRow(ws As Worksheet, row As Long, startCol As Long, endCol As Long)
    Dim i As Long
    Dim rng As Range

    For i = startCol To endCol
        Set rng = ws.Cells(row, i)
        With rng.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
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
    Dim outputRow As Long
    Dim lanBan As Long
    Dim dictXe As Object

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
            .Range("A1").Value = "Số Xe"
            .Range("B1").Value = "Lần Bẩn"
            .Range("C1").Value = "Ngày Bắt Đầu"
            .Range("D1").Value = "Ngày Bẩn Vệ Sinh"
            .Range("E1").Value = "Số Ngày"
            .Range("F1").Value = "Trung Bình"
            .Range("A1:F1").Font.Bold = True
            .Range("A1:F1").Interior.Color = RGB(54, 96, 146)
            .Range("A1:F1").Font.Color = RGB(255, 255, 255)
        End With
    End If
    On Error GoTo 0

    ' Xóa dữ liệu cũ
    wsTongHop.Range("A2:F" & wsTongHop.Cells(wsTongHop.Rows.Count, 1).End(xlUp).Row).Delete

    ' Duyệt qua tất cả dòng trong sheet QuetXe
    lastRow = wsQuet.Cells(wsQuet.Rows.Count, 1).End(xlUp).Row
    outputRow = 2

    For row = 2 To lastRow
        maCo = Trim(wsQuet.Cells(row, 1).Value)
        ghiChu = Trim(LCase(wsQuet.Cells(row, 5).Value))

        ' Kiểm tra từ khóa bẩn
        If IsKeywordBan(ghiChu) Then
            ' Tăng số lần bẩn
            If Not dictXe.Exists(maCo) Then
                dictXe.Add maCo, 0
            End If
            dictXe(maCo) = dictXe(maCo) + 1
            lanBan = dictXe(maCo)

            ' Ghi vào TongHop
            With wsTongHop
                .Cells(outputRow, 1).Value = maCo
                .Cells(outputRow, 2).Value = lanBan
                .Cells(outputRow, 3).Value = Format(wsQuet.Cells(row, 2).Value, "dd/mm/yyyy")
                .Cells(outputRow, 4).Value = Format(wsQuet.Cells(row, 2).Value, "dd/mm/yyyy")
                .Cells(outputRow, 5).Value = "Cần vệ sinh"

                ' Border
                .Range("A" & outputRow & ":F" & outputRow).Borders.LineStyle = xlContinuous
            End With

            outputRow = outputRow + 1
        End If
    Next row

    wsTongHop.Activate
    MsgBox "✓ Cập nhật tổng hợp thành công! Tổng xe bẩn: " & dictXe.Count, vbInformation
End Sub

' ===== KIỂM TRA TỪ KHÓA BẨN =====
Private Function IsKeywordBan(ghiChu As String) As Boolean
    Dim keywords As Variant
    Dim i As Long

    keywords = Array("ban", "banh", "ko sach", "khong sach", "sach", "binh")

    For i = LBound(keywords) To UBound(keywords)
        If InStr(1, ghiChu, keywords(i), vbTextCompare) > 0 Then
            IsKeywordBan = True
            Exit Function
        End If
    Next i

    IsKeywordBan = False
End Function

' ===== MACRO XÓA DÒ HIỆN TẠI =====
Sub XoaDongHienTai()
    Dim ws As Worksheet
    Dim deleteConfirm As VbMsgBoxResult
    Dim selectedRow As Long

    Set ws = ThisWorkbook.ActiveSheet
    selectedRow = ActiveCell.Row

    ' Không xóa tiêu đề
    If selectedRow = 1 Then
        MsgBox "Không thể xóa tiêu đề!", vbExclamation
        Exit Sub
    End If

    ' Xác nhận
    deleteConfirm = MsgBox("Xóa dòng " & selectedRow & "?", vbYesNo + vbQuestion)

    If deleteConfirm = vbYes Then
        ws.Range("A" & selectedRow & ":F" & selectedRow).Delete xlShiftUp
        MsgBox "✓ Xóa thành công", vbInformation
    End If
End Sub
