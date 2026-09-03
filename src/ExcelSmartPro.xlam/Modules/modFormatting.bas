Attribute VB_Name = "modFormatting"
Option Explicit

Public Sub IncreaseColour()
    Dim c As Range
    For Each c In TargetRangeOrSelection.Cells
        c.Interior.TintAndShade = c.Interior.TintAndShade + 0.1
    Next c
End Sub

Public Sub DecreaseColour()
    Dim c As Range
    For Each c In TargetRangeOrSelection.Cells
        c.Interior.TintAndShade = c.Interior.TintAndShade - 0.1
    Next c
End Sub

Public Sub RedHighlight(): ApplyFillColor RGB(255, 199, 206): End Sub
Public Sub GreenHighlight(): ApplyFillColor RGB(198, 239, 206): End Sub
Public Sub OrangeHighlight(): ApplyFillColor RGB(255, 235, 156): End Sub
Public Sub BlueHighlight(): ApplyFillColor RGB(221, 235, 247): End Sub

Public Sub ClearHighlight()
    TargetRangeOrSelection.Interior.Pattern = xlNone
End Sub

Public Sub ChangeBorderColor()
    Dim clr As Variant
    clr = Application.InputBox("Enter border color RGB as R,G,B (example: 0,112,192)", "Border Color", "0,112,192", Type:=2)
    If clr = False Then Exit Sub
    Dim parts() As String
    parts = Split(CStr(clr), ",")
    If UBound(parts) <> 2 Then Exit Sub
    Dim rng As Range
    Set rng = TargetRangeOrSelection
    With rng.Borders
        .Color = RGB(CLng(parts(0)), CLng(parts(1)), CLng(parts(2)))
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
End Sub

Public Sub StatusTickmarks()
    Dim c As Range
    For Each c In TargetRangeOrSelection.Cells
        If Trim$(CStr(c.Value)) <> "" Then
            c.Value = ChrW(&H2713) & " " & c.Value
        End If
    Next c
End Sub

Private Sub ApplyFillColor(ByVal fillColor As Long)
    TargetRangeOrSelection.Interior.Color = fillColor
End Sub
