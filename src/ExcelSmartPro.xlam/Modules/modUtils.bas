Attribute VB_Name = "modUtils"
Option Explicit

Public Const APP_NAME As String = "ExcelSmart Pro"

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

Public Function AskYesNo(ByVal msg As String, Optional ByVal title As String = APP_NAME) As Boolean
    AskYesNo = (MsgBox(msg, vbQuestion + vbYesNo, title) = vbYes)
End Function

Public Sub SafeOff()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub SafeOn()
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
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

Public Function RegexExtractFirst(ByVal textValue As String, ByVal pattern As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.IgnoreCase = True
    If re.Test(textValue) Then
        Set m = re.Execute(textValue)(0)
        RegexExtractFirst = m.Value
    End If
End Function

' Double-quote helper: safer than nesting quotes in string literals.
Public Function Q() As String
    Q = Chr(34)
End Function

' Folder picker dialog. Returns "" if cancelled.
Public Function BrowseForFolder(ByVal prompt As String) As String
    Dim shellApp As Object, fld As Object
    Set shellApp = CreateObject("Shell.Application")
    Set fld = shellApp.BrowseForFolder(0, prompt, 0)
    If fld Is Nothing Then Exit Function
    On Error Resume Next
    BrowseForFolder = fld.Self.Path
    On Error GoTo 0
    If Len(BrowseForFolder) = 0 Then BrowseForFolder = ""
End Function

' Persistent settings stored per user in the registry.
Public Function GetSettingString(ByVal key As String, ByVal defaultValue As String) As String
    On Error Resume Next
    GetSettingString = GetSetting(APP_NAME, "Settings", key, defaultValue)
    On Error GoTo 0
    If Len(GetSettingString) = 0 Then GetSettingString = defaultValue
End Function

Public Sub SaveSettingString(ByVal key As String, ByVal value As String)
    On Error Resume Next
    SaveSetting APP_NAME, "Settings", key, value
    On Error GoTo 0
End Sub

Public Sub SetValidationFormula(ByVal target As Range, ByVal formula1 As String, _
    Optional ByVal inputTitle As String = "Validation", _
    Optional ByVal errorTitle As String = "Invalid Value")
    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0
    target.Validation.Add Type:=xlValidateCustom, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=formula1
    target.Validation.IgnoreBlank = True
    target.Validation.InputTitle = inputTitle
    target.Validation.ErrorTitle = errorTitle
    target.Validation.InputMessage = "Enter a valid value."
    target.Validation.ErrorMessage = "The value does not match the required format."
    target.Validation.ShowInput = True
    target.Validation.ShowError = True
End Sub

Public Sub SetTextLengthValidation(ByVal target As Range, ByVal minLen As Long, ByVal maxLen As Long)
    On Error Resume Next
    target.Validation.Delete
    On Error GoTo 0
    target.Validation.Add Type:=xlValidateTextLength, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:=CStr(minLen), Formula2:=CStr(maxLen)
End Sub

' Get a free (non-existent) sheet name based on a base name.
Public Function GetFreeSheetName(ByVal baseName As String) As String
    Dim ws As Worksheet, name As String, i As Long
    name = baseName
    i = 1
    Do
        Dim found As Boolean
        found = False
        For Each ws In ActiveWorkbook.Worksheets
            If StrComp(ws.Name, name, vbTextCompare) = 0 Then found = True
        Next ws
        If Not found Then Exit Do
        i = i + 1
        name = baseName & "_" & i
    Loop
    GetFreeSheetName = name
End Function
