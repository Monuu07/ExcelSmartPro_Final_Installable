Attribute VB_Name = "modValidation"
Option Explicit

' NOTE: All formulas below use the Q() helper (returns a double-quote) to
' avoid nested-quote VBA compile errors that broke the original build.

Private Function FirstCellRef() As String
    FirstCellRef = TargetRangeOrSelection.Cells(1, 1).Address(False, False)
End Function

Public Sub EmailValidation()
    Dim f As String
    f = "=AND(ISNUMBER(SEARCH(" & Q() & "@" & Q() & "," & FirstCellRef() & "))," & _
        "ISNUMBER(SEARCH(" & Q() & "." & Q() & "," & FirstCellRef() & ")))"
    SetValidationFormula TargetRangeOrSelection, f, "Email", "Invalid Email"
End Sub

Public Sub PhoneValidation()
    Dim digits As String
    digits = "SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" & FirstCellRef() & "," & Q() & "-" & Q() & "," & Q() & Q() & ")," & Q() & " " & Q() & "," & Q() & Q() & ")," & Q() & "+" & Q() & "," & Q() & Q() & ")"
    Dim f As String
    f = "=AND(ISNUMBER(" & digits & "),LEN(" & digits & ")>=10,LEN(" & digits & ")<=15)"
    SetValidationFormula TargetRangeOrSelection, f, "Phone", "Invalid Phone"
End Sub

Public Sub NumericOnly()
    On Error Resume Next
    TargetRangeOrSelection.Validation.Delete
    On Error GoTo 0
    TargetRangeOrSelection.Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="-1E+307", Formula2:="1E+307"
End Sub

Public Sub TextOnly()
    SetTextLengthValidation TargetRangeOrSelection, 1, 255
End Sub

Public Sub URLValidation()
    Dim f As String
    f = "=OR(LEFT(" & FirstCellRef() & ",7)=" & Q() & "http://" & Q() & "," & _
        "LEFT(" & FirstCellRef() & ",8)=" & Q() & "https://" & Q() & ")"
    SetValidationFormula TargetRangeOrSelection, f, "URL", "Invalid URL"
End Sub

Public Sub DateValidation()
    On Error Resume Next
    TargetRangeOrSelection.Validation.Delete
    On Error GoTo 0
    TargetRangeOrSelection.Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=DATE(2000,1,1)", Formula2:="=DATE(2099,12,31)"
End Sub

Public Sub RemoveValidation()
    On Error Resume Next
    TargetRangeOrSelection.Validation.Delete
    On Error GoTo 0
End Sub

' =====================================================================
'  VALIDATE EXISTING DATA (marks invalid cells, XLBooster-style)
' =====================================================================

Public Sub ValidateEmailData()
    ValidateColumn "email"
End Sub

Public Sub ValidatePhoneData()
    ValidateColumn "phone"
End Sub

Private Sub ValidateColumn(ByVal validateType As String)
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the data cells to validate.", vbExclamation, APP_NAME
        Exit Sub
    End If
    Dim cell As Range
    Dim pattern As String
    Dim okCount As Long, badCount As Long

    Select Case validateType
        Case "email": pattern = "^[\w\.\-]+@[\w\-\.]+\.[A-Za-z]{2,}$"
        Case "phone": pattern = "^[\+\s\d\-\(\)]{10,15}$"
    End Select

    Application.ScreenUpdating = False
    For Each cell In Selection.Cells
        If Len(Trim$(CStr(cell.Value))) > 0 Then
            If IsRegexMatch(CStr(cell.Value), pattern) Then
                okCount = okCount + 1
            Else
                badCount = badCount + 1
                cell.Interior.Color = RGB(255, 199, 206)
                cell.Font.Color = RGB(156, 0, 6)
            End If
        End If
    Next cell
    Application.ScreenUpdating = True

    MsgBox "Valid: " & okCount & "   Invalid (marked red): " & badCount, vbInformation, APP_NAME
End Sub
