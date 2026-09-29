Attribute VB_Name = "modAutoSave"
Option Explicit
' DAN VAO: Insert > Module (hoac File > Import File)

Private Const INTERVAL_MIN As Long = 30
Private mNext As Date

Private Function ProcName() As String
    ProcName = "'" & ThisWorkbook.Name & "'!AutoSaveNow"
End Function

Public Sub StartAutoSave()
    StopAutoSave
    If Len(ThisWorkbook.Path) = 0 Then Exit Sub
    mNext = Now + TimeSerial(0, INTERVAL_MIN, 0)
    Application.OnTime mNext, ProcName()
    Application.StatusBar = modTongHop.T("Auto Save: l{1EA7}n l{1B0}u ti{1EBF}p theo l{FA}c ") & Format$(mNext, "hh:mm:ss")
End Sub

Public Sub StopAutoSave()
    If mNext <> 0 Then
        On Error Resume Next
        Application.OnTime mNext, ProcName(), , False
        On Error GoTo 0
        mNext = 0
    End If
    Application.StatusBar = False
End Sub

Public Sub AutoSaveNow()
    Dim folder As String, nm As String, p As Long
    Dim sep As String
    sep = Application.PathSeparator

    On Error Resume Next
    ThisWorkbook.Save

    folder = ThisWorkbook.Path & sep & "Backup_60Phut"
    If Len(Dir(folder, vbDirectory)) = 0 Then MkDir folder
    nm = ThisWorkbook.Name
    p = InStrRev(nm, ".")
    If p > 0 Then nm = Left$(nm, p - 1)
    ThisWorkbook.SaveCopyAs folder & sep & nm & "_BACKUP.xlsm"
    On Error GoTo 0

    StartAutoSave
End Sub
