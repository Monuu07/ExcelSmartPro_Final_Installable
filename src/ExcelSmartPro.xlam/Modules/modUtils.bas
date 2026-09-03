Attribute VB_Name = "modUtils"
Option Explicit

Public Function TargetRangeOrSelection() As Range
    If TypeName(Selection) = "Range" Then
        Set TargetRangeOrSelection = Selection
    Else
        Set TargetRangeOrSelection = ActiveCell
    End If
End Function

Public Function PromptText(ByVal title As String, ByVal prompt As String, Optional ByVal defaultValue As String = "") As String
    PromptText = InputBox(prompt, title, defaultValue)
End Function

Public Function AskYesNo(ByVal msg As String, Optional ByVal title As String = "ExcelSmart Pro") As Boolean
    AskYesNo = (MsgBox(msg, vbQuestion + vbYesNo, title) = vbYes)
End Function

Public Sub SafeOff()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
End Sub

Public Sub SafeOn()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
End Sub

Public Function LastUsedRow(ws As Worksheet) As Long
    On Error Resume Next
    LastUsedRow = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    If LastUsedRow = 0 Then LastUsedRow = 1
    On Error GoTo 0
End Function

Public Function LastUsedCol(ws As Worksheet) As Long
    On Error Resume Next
    LastUsedCol = ws.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    If LastUsedCol = 0 Then LastUsedCol = 1
    On Error GoTo 0
End Function

Public Function IsRegexMatch(ByVal textValue As String, ByVal pattern As String) As Boolean
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = True
    re.Global = False
    IsRegexMatch = re.Test(textValue)
End Function

Public Sub AddValidationWithFormula(ByVal target As Range, ByVal formula1 As String, Optional ByVal inputTitle As String = "Validation", Optional ByVal errorTitle As String = "Invalid Value")
    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0
    target.Validation.Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=formula1
    target.Validation.IgnoreBlank = True
    target.Validation.InCellDropdown = True
    target.Validation.InputTitle = inputTitle
    target.Validation.ErrorTitle = errorTitle
    target.Validation.InputMessage = "Enter a valid value."
    target.Validation.ErrorMessage = "The value does not match the required format."
    target.Validation.ShowInput = True
    target.Validation.ShowError = True
End Sub

Public Sub AddTextValidation(ByVal target As Range, ByVal minLen As Long, ByVal maxLen As Long)
    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0
    target.Validation.Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:=CStr(minLen), Formula2:=CStr(maxLen)
End Sub
