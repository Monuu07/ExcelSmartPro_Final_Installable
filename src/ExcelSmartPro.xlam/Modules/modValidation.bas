Attribute VB_Name = "modValidation"
Option Explicit

Private Function FirstCellRef() As String
    FirstCellRef = TargetRangeOrSelection.Cells(1, 1).Address(False, False)
End Function

Public Sub EmailValidation()
    AddValidationWithFormula TargetRangeOrSelection, _
        "=AND(ISNUMBER(SEARCH("@"," & FirstCellRef() & ")),ISNUMBER(SEARCH("."," & FirstCellRef() & ")))", _
        "Email", "Invalid Email"
End Sub

Public Sub PhoneValidation()
    AddValidationWithFormula TargetRangeOrSelection, _
        "=AND(LEN(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" & FirstCellRef() & ","-","")," ",""),"+",""))>=10,LEN(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(" & FirstCellRef() & ","-","")," ",""),"+",""))<=15)", _
        "Phone", "Invalid Phone"
End Sub

Public Sub NumericOnly()
    On Error Resume Next
    TargetRangeOrSelection.Validation.Delete
    On Error GoTo 0
    TargetRangeOrSelection.Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="-1E+307", Formula2:="1E+307"
End Sub

Public Sub TextOnly()
    AddTextValidation TargetRangeOrSelection, 1, 255
End Sub

Public Sub URLValidation()
    AddValidationWithFormula TargetRangeOrSelection, _
        "=OR(LEFT(" & FirstCellRef() & ",7)="http://",LEFT(" & FirstCellRef() & ",8)="https://")", _
        "URL", "Invalid URL"
End Sub

Public Sub DateValidation()
    On Error Resume Next
    TargetRangeOrSelection.Validation.Delete
    On Error GoTo 0
    TargetRangeOrSelection.Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=DATE(2000,1,1)", Formula2:="=DATE(2099,12,31)"
End Sub

Public Sub RemoveValidation()
    On Error Resume Next
    TargetRangeOrSelection.Validation.Delete
    On Error GoTo 0
End Sub
