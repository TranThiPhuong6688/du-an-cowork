Attribute VB_Name = "modTongHop"
Option Explicit
' DAN VAO: Insert > Module (hoac File > Import File)
' Chay: Alt+F8 > CapNhatTongHop_Ngay > Run

' 1 = Ngay bat dau la lan quet TRUOC do cua xe (mac dinh)
' 2 = Ngay bat dau la lan BAN TRUOC do cua xe (lan dau: lan quet dau tien)
Private Const START_MODE As Long = 1

Public Function T(ByVal s As String) As String
    Dim p As Long, q As Long
    Do
        p = InStr(s, "{")
        If p = 0 Then Exit Do
        q = InStr(p, s, "}")
        If q = 0 Then Exit Do
        s = Left$(s, p - 1) & ChrW(CLng("&H" & Mid$(s, p + 1, q - p - 1))) & Mid$(s, q + 1)
    Loop
    T = s
End Function

Private Function IsDirty(ByVal s As String) As Boolean
    Dim kws As Variant, sep As Variant, i As Long
    s = LCase$(s)
    For Each sep In Array(",", ";", ".", "/", "-", "(", ")", "_")
        s = Replace(s, sep, " ")
    Next sep
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    s = " " & Trim$(s) & " "
    kws = Array(T("b{1EA9}n"), T("b{1EE7}n"), "ban", "bun", _
                T("kh{F4}ng s{1EA1}ch"), "khong sach", _
                T("ko s{1EA1}ch"), "ko sach", _
                T("ch{1B0}a s{1EA1}ch"), "chua sach")
    For i = LBound(kws) To UBound(kws)
        If InStr(s, " " & kws(i) & " ") > 0 Then IsDirty = True: Exit Function
    Next i
End Function

Public Sub CapNhatTongHop_Ngay()
    Dim src As Worksheet, wsT As Worksheet
    Dim last As Long, v As Variant, i As Long
    Dim dCars As Object, dNames As Object, key As String, code As String
    Dim colls As Object, x As Variant

    Set src = ActiveSheet
    If src.Name = "TongHop" Then
        On Error Resume Next
        Set src = ThisWorkbook.Worksheets("QuetXe")
        On Error GoTo 0
        If src.Name = "TongHop" Then
            MsgBox "Hay dung o sheet quet xe roi chay lai.", vbExclamation
            Exit Sub
        End If
    End If

    last = src.Cells(src.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then MsgBox T("Kh{F4}ng c{F3} d{1EEF} li{1EC7}u."), vbExclamation: Exit Sub
    v = src.Range("A2:E" & last).Value

    On Error Resume Next
    Set wsT = ThisWorkbook.Worksheets("TongHop")
    On Error GoTo 0
    If wsT Is Nothing Then
        Set wsT = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsT.Name = "TongHop"
    End If

    Set dCars = CreateObject("Scripting.Dictionary")
    Set dNames = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(v, 1)
        code = Trim$(v(i, 1) & "")
        If Len(code) > 0 Then
            If IsDate(v(i, 2)) Or IsNumeric(v(i, 2)) Then
                If Not IsEmpty(v(i, 2)) Then
                    key = LCase$(code)
                    If Not dCars.Exists(key) Then
                        Set colls = New Collection
                        dCars.Add key, colls
                        dNames.Add key, code
                    End If
                    dCars(key).Add i
                End If
            End If
        End If
    Next i

    Dim outv() As Variant, outN As Long
    ReDim outv(1 To UBound(v, 1), 1 To 6)

    Dim n As Long, idx() As Long, tm() As Double, k As Long, j As Long
    Dim ti As Long, td As Double
    Dim prevScan As Double, prevDirty As Double, st As Double, tt As Double
    Dim cnt As Long, sumD As Double, firstOut As Long, days As Long

    For Each x In dCars.Keys
        n = dCars(x).Count
        ReDim idx(1 To n): ReDim tm(1 To n)
        For k = 1 To n
            idx(k) = dCars(x)(k)
            If IsDate(v(idx(k), 2)) Then tm(k) = CDbl(CDate(v(idx(k), 2))) Else tm(k) = CDbl(v(idx(k), 2))
        Next k
        For k = 2 To n
            ti = idx(k): td = tm(k): j = k - 1
            Do While j >= 1
                If tm(j) <= td Then Exit Do
                idx(j + 1) = idx(j): tm(j + 1) = tm(j): j = j - 1
            Loop
            idx(j + 1) = ti: tm(j + 1) = td
        Next k

        prevScan = 0: prevDirty = 0: cnt = 0: sumD = 0: firstOut = 0
        For k = 1 To n
            tt = tm(k)
            If IsDirty(v(idx(k), 5) & "") Then
                cnt = cnt + 1
                If START_MODE = 2 Then
                    If prevDirty > 0 Then st = prevDirty Else st = tm(1)
                Else
                    If prevScan > 0 Then st = prevScan Else st = tt
                End If
                days = CLng(Int(tt) - Int(st))
                sumD = sumD + days
                outN = outN + 1
                outv(outN, 1) = dNames(x)
                outv(outN, 2) = cnt
                outv(outN, 3) = Int(st)
                outv(outN, 4) = Int(tt)
                outv(outN, 5) = days
                If cnt = 1 Then firstOut = outN
                prevDirty = tt
            End If
            prevScan = tt
        Next k
        If cnt > 0 Then outv(firstOut, 6) = sumD / cnt
    Next x

    wsT.Range("A1:F" & wsT.Rows.Count).Clear
    wsT.Cells.Font.Name = "Times New Roman"
    wsT.Cells.Font.Size = 12
    wsT.Range("A1:F1").Value = Array(T("S{1ED1} xe"), T("L{1EA7}n b{1EA9}n c{1EA7}n v{1EC7} sinh"), _
        T("Ng{E0}y b{1EAF}t {111}{1EA7}u"), T("Ng{E0}y b{1EA9}n v{1EC7} sinh"), _
        T("S{1ED1} ng{E0}y quay v{F2}ng b{1EA9}n"), T("Trung b{EC}nh s{1ED1} ng{E0}y b{1EA9}n quay v{F2}ng"))
    wsT.Range("A1:F1").Font.Bold = True
    wsT.Range("A1:F1").WrapText = True
    wsT.Range("A1:F1").HorizontalAlignment = xlCenter

    If outN = 0 Then
        wsT.Activate
        MsgBox T("Kh{F4}ng c{F3} d{F2}ng b{1EA9}n n{E0}o."), vbInformation
        Exit Sub
    End If

    wsT.Range("A2:A" & outN + 1).NumberFormat = "@"
    wsT.Range("C2:D" & outN + 1).NumberFormat = "dd/mm/yyyy"
    wsT.Range("E2:E" & outN + 1).NumberFormat = T("0"" ng{E0}y""")
    wsT.Range("F2:F" & outN + 1).NumberFormat = T("0.0"" ng{E0}y""")
    wsT.Range("A2").Resize(outN, 6).Value = outv
    wsT.Range("A1:F" & outN + 1).Borders.LineStyle = xlContinuous
    wsT.Range("B2:F" & outN + 1).HorizontalAlignment = xlCenter
    wsT.Columns("A:F").ColumnWidth = 20
    wsT.Activate
    MsgBox T("{110}{E3} c{1EAD}p nh{1EAD}t TongHop: ") & outN & " dong.", vbInformation
End Sub
