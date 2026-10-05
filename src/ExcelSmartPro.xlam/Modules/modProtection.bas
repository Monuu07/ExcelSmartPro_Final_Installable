Attribute VB_Name = "modProtection"
Option Explicit

Public Sub BackupFile()
    Dim p As String, backupName As String
    p = ActiveWorkbook.Path
    If p = "" Then
        MsgBox "Please save the workbook first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    backupName = p & Application.PathSeparator & Replace(ActiveWorkbook.Name, ".", "_backup_") & Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ActiveWorkbook.SaveCopyAs backupName
    MsgBox "Backup created:" & vbCrLf & backupName, vbInformation, APP_NAME
End Sub

Public Sub ProtectSheets()
    Dim pwd As String, ws As Worksheet
    pwd = PromptText("Protect Sheets", "Password (leave blank for none)")
    SafeOff
    For Each ws In ActiveWorkbook.Worksheets
        ws.Protect Password:=pwd, DrawingObjects:=True, Contents:=True, Scenarios:=True
    Next ws
    SafeOn
    MsgBox "All sheets protected.", vbInformation, APP_NAME
End Sub

Public Sub UnprotectSheets()
    Dim pwd As String, ws As Worksheet
    pwd = PromptText("Unprotect Sheets", "Password (leave blank for none)")
    SafeOff
    For Each ws In ActiveWorkbook.Worksheets
        On Error Resume Next
        ws.Unprotect Password:=pwd
        On Error GoTo 0
    Next ws
    SafeOn
    MsgBox "Done. Sheets that accepted the password are unprotected.", vbInformation, APP_NAME
End Sub

Public Sub ProtectWorkbook()
    Dim pwd As String
    pwd = PromptText("Protect Workbook", "Password (leave blank for none)")
    ActiveWorkbook.Protect Password:=pwd, Structure:=True, Windows:=False
    MsgBox "Workbook structure protected.", vbInformation, APP_NAME
End Sub

Public Sub UnprotectWorkbook()
    Dim pwd As String
    pwd = PromptText("Unprotect Workbook", "Password (leave blank for none)")
    On Error Resume Next
    ActiveWorkbook.Unprotect Password:=pwd
    On Error GoTo 0
    MsgBox "Done.", vbInformation, APP_NAME
End Sub

Public Sub LockFormulas()
    On Error Resume Next
    TargetRangeOrSelection.SpecialCells(xlCellTypeFormulas).Locked = True
    On Error GoTo 0
End Sub

Public Sub UnlockFormulas()
    On Error Resume Next
    TargetRangeOrSelection.SpecialCells(xlCellTypeFormulas).Locked = False
    On Error GoTo 0
End Sub

Public Sub LockSelection()
    TargetRangeOrSelection.Locked = True
End Sub

Public Sub UnlockSelection()
    TargetRangeOrSelection.Locked = False
End Sub
