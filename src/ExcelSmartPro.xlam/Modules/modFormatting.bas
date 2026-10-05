Attribute VB_Name = "modFormatting"
Option Explicit

' =====================================================================
'  COLOR / HIGHLIGHT TOOLS
' =====================================================================

' XLBooster-style tint/shade adjusters (HSL-based, 10% steps)
Public Sub IncreaseColour()
    AdjustFillColor True
End Sub

Public Sub DecreaseColour()
    AdjustFillColor False
End Sub

Private Sub AdjustFillColor(ByVal lighten As Boolean)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Dim h As Double, s As Double, l As Double
    Dim r As Double, g As Double, b As Double

    Application.ScreenUpdating = False
    For Each cell In Selection.Cells
        If cell.Interior.ColorIndex <> xlNone Then
            r = cell.Interior.Color Mod 256
            g = (cell.Interior.Color \ 256) Mod 256
            b = cell.Interior.Color \ 65536
            RGBtoHSL r, g, b, h, s, l
            If lighten Then l = l + 0.1 Else l = l - 0.1
            If l < 0 Then l = 0
            If l > 1 Then l = 1
            HSLtoRGB h, s, l, r, g, b
            cell.Interior.Color = RGB(CInt(r), CInt(g), CInt(b))
        End If
    Next cell
    Application.ScreenUpdating = True
End Sub

Private Sub RGBtoHSL(ByVal r As Double, ByVal g As Double, ByVal b As Double, _
    ByRef h As Double, ByRef s As Double, ByRef l As Double)
    r = r / 255: g = g / 255: b = b / 255
    Dim mx As Double, mn As Double, d As Double
    mx = Application.WorksheetFunction.Max(r, g, b)
    mn = Application.WorksheetFunction.Min(r, g, b)
    d = mx - mn
    l = (mx + mn) / 2
    If d = 0 Then
        h = 0: s = 0
    Else
        s = IIf(l > 0.5, d / (2 - mx - mn), d / (mx + mn))
        Select Case mx
            Case r: h = (g - b) / d + IIf(g < b, 6, 0)
            Case g: h = (b - r) / d + 2
            Case b: h = (r - g) / d + 4
        End Select
        h = h / 6
    End If
End Sub

Private Sub HSLtoRGB(ByVal h As Double, ByVal s As Double, ByVal l As Double, _
    ByRef r As Double, ByRef g As Double, ByRef b As Double)
    Dim q As Double, p As Double
    If s = 0 Then
        r = l: g = l: b = l
    Else
        q = IIf(l < 0.5, l * (1 + s), l + s - l * s)
        p = 2 * l - q
        r = ColorChannel(p, q, h + 1 / 3)
        g = ColorChannel(p, q, h)
        b = ColorChannel(p, q, h - 1 / 3)
    End If
    r = r * 255: g = g * 255: b = b * 255
End Sub

Private Function ColorChannel(ByVal p As Double, ByVal q As Double, ByVal tc As Double) As Double
    If tc < 0 Then tc = tc + 1
    If tc > 1 Then tc = tc - 1
    If tc < 1 / 6 Then
        ColorChannel = p + (q - p) * 6 * tc
    ElseIf tc < 1 / 2 Then
        ColorChannel = q
    ElseIf tc < 2 / 3 Then
        ColorChannel = p + (q - p) * (2 / 3 - tc) * 6
    Else
        ColorChannel = p
    End If
End Function

' =====================================================================
'  HIGHLIGHT TOOLS (fill + thick border, XLBooster-style)
' =====================================================================

Public Sub RedHighlight(): ApplyHighlight RGB(255, 199, 206), RGB(255, 40, 40): End Sub
Public Sub GreenHighlight(): ApplyHighlight RGB(198, 239, 206), RGB(40, 180, 60): End Sub
Public Sub OrangeHighlight(): ApplyHighlight RGB(255, 235, 156), RGB(255, 120, 20): End Sub
Public Sub BlueHighlight(): ApplyHighlight RGB(221, 235, 247), RGB(50, 140, 255): End Sub

Private Sub ApplyHighlight(ByVal fillColor As Long, ByVal borderColor As Long)
    If TypeName(Selection) <> "Range" Then Exit Sub
    With Selection
        .Interior.Color = fillColor
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThick
        .Borders.Color = borderColor
    End With
End Sub

Public Sub ClearHighlight()
    If TypeName(Selection) <> "Range" Then Exit Sub
    With Selection
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With
End Sub

Public Sub ChangeBorderColor()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim clr As Variant
    clr = Application.InputBox("Enter border color RGB as R,G,B (example: 0,112,192)", _
        "Border Color", "0,112,192", Type:=2)
    If VarType(clr) = vbBoolean Then Exit Sub
    Dim parts() As String
    parts = Split(CStr(clr), ",")
    If UBound(parts) <> 2 Then
        MsgBox "Please enter three numbers separated by commas.", vbExclamation, APP_NAME
        Exit Sub
    End If
    With Selection.Borders
        .Color = RGB(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Public Sub StatusTickmarks()
    ApplyTickToSelection 1
End Sub

' Apply tick by index into TickList (1-based)
Public Sub ApplyTickToSelection(ByVal tickIndex As Long)
    Dim list As Variant
    list = modTickmarks.TickList()
    If tickIndex < 1 Or tickIndex > UBound(list) - LBound(list) + 1 Then Exit Sub
    Dim entry As Variant
    entry = list(LBound(list) + tickIndex - 1)
    modTickmarks.InsertTickIntoSelection CLng(entry(1)), CLng(entry(2))
End Sub

' =====================================================================
'  FOCUS CELL / ROW / COLUMN (XLBooster-style visual focus)
' =====================================================================

Public Sub ToggleFocusCell()
    FocusMode "cell"
End Sub

Public Sub ToggleFocusRow()
    FocusMode "row"
End Sub

Public Sub ToggleFocusColumn()
    FocusMode "col"
End Sub

Private Sub FocusMode(ByVal targetMode As String)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cur As String
    cur = GetSettingString("FocusMode", "")
    If cur = targetMode Then
        SaveSettingString "FocusMode", ""
        RemoveFocusFormatting
        MsgBox "Focus mode turned OFF.", vbInformation, APP_NAME
    Else
        SaveSettingString "FocusMode", targetMode
        UpdateFocus
        MsgBox "Focus mode: " & Choose(InStr("cellrowcol", targetMode) \ 4 + 1, "Cell", "Row", "Column") & _
               " (active). Click it again to turn off.", vbInformation, APP_NAME
    End If
End Sub

' Called from the event handler on each selection change.
Public Sub UpdateFocus()
    On Error Resume Next
    Dim mode As String
    mode = GetSettingString("FocusMode", "")
    If Len(mode) = 0 Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub

    Dim ws As Worksheet: Set ws = ActiveSheet
    RemoveFocusFormatting
    Dim rng As Range: Set rng = Selection
    Dim lastR As Long, lastC As Long
    lastR = LastUsedRow(ws)
    lastC = LastUsedCol(ws)
    If lastR < 1 Or lastC < 1 Then Exit Sub

    Select Case mode
        Case "cell"
            ws.Range(ws.Cells(rng.Row, 1), ws.Cells(rng.Row, lastC)).Interior.Color = RGB(255, 255, 200)
            rng.Interior.Color = RGB(255, 242, 204)
        Case "row"
            ws.Range(ws.Cells(rng.Row, 1), ws.Cells(rng.Row, lastC)).Interior.Color = RGB(255, 255, 200)
        Case "col"
            ws.Range(ws.Cells(1, rng.Column), ws.Cells(lastR, rng.Column)).Interior.Color = RGB(255, 255, 200)
    End Select
End Sub

Private Sub RemoveFocusFormatting()
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim lastR As Long, lastC As Long
    lastR = LastUsedRow(ws)
    lastC = LastUsedCol(ws)
    If lastR < 1 Or lastC < 1 Then Exit Sub
    Dim cell As Range
    For Each cell In ws.Range(ws.Cells(1, 1), ws.Cells(lastR, lastC))
        If cell.Interior.Color = RGB(255, 255, 200) Or cell.Interior.Color = RGB(255, 242, 204) Then
            cell.Interior.Pattern = xlNone
        End If
    Next cell
End Sub
