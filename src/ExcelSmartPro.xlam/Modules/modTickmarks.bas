Attribute VB_Name = "modTickmarks"
Option Explicit

' =====================================================================
'  TICKMARKS - 15 status symbols with colors, XLBooster-style
'  Each tick: name, Unicode codepoint, RGB color
' =====================================================================

Public Function TickList() As Variant
    TickList = Array( _
        Array("Green Tick", CLng(&H2714), RGB(0, 176, 80)), _
        Array("Red Cross", CLng(&H274C), RGB(255, 0, 0)), _
        Array("Question", CLng(&H2753), RGB(0, 112, 192)), _
        Array("Red Heart", CLng(&H2764), RGB(220, 0, 0)), _
        Array("Flag Red", CLng(&H2691), RGB(255, 0, 0)), _
        Array("Flag Green", CLng(&H2691), RGB(0, 176, 80)), _
        Array("Flag Yellow", CLng(&H2691), RGB(215, 193, 19)), _
        Array("Warning", CLng(&H26A0), RGB(255, 192, 0)), _
        Array("Clock", CLng(&H23F0), RGB(255, 128, 0)), _
        Array("Pending", CLng(&H23F3), RGB(255, 192, 0)), _
        Array("Info", CLng(&H24D8), RGB(0, 0, 0)), _
        Array("Star", CLng(&H2605), RGB(255, 192, 0)), _
        Array("Lock", CLng(&H26BF), RGB(0, 112, 192)), _
        Array("Blocked", CLng(&H26D4), RGB(192, 0, 0)), _
        Array("OK", CLng(&H2705), RGB(0, 176, 80)) _
    )
End Function

' Insert a tick into every cell of the current selection.
Public Sub InsertTickIntoSelection(ByVal codeHex As Long, ByVal tickColor As Long)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    Application.ScreenUpdating = False
    For Each cell In Selection.Cells
        InsertTickIntoCell cell, codeHex, tickColor
    Next cell
    Application.ScreenUpdating = True
    On Error Resume Next
    Selection.Select
    On Error GoTo 0
End Sub

Private Sub InsertTickIntoCell(ByVal targetCell As Range, ByVal codeHex As Long, ByVal tickColor As Long)
    If targetCell Is Nothing Then Exit Sub

    Dim rng As Range
    Set rng = targetCell
    If rng.MergeCells Then Set rng = rng.MergeArea.Cells(1, 1)

    Dim baseText As String
    baseText = CStr(rng.Value)

    Dim sym As String
    sym = ChrW(codeHex)

    Dim finalText As String
    If InStr(baseText, sym) > 0 Then
        ' symbol already present - remove it (toggle behaviour)
        finalText = Replace(baseText, sym & " ", "")
        finalText = Replace(finalText, sym, "")
    Else
        finalText = sym & " " & baseText
    End If

    rng.Value = finalText

    ' color the symbol
    If Len(finalText) > 0 Then
        On Error Resume Next
        rng.Characters(1, Len(sym)).Font.Color = tickColor
        rng.Characters(1, Len(sym)).Font.Bold = True
        On Error GoTo 0
    End If
End Sub

' Remove all ticks from selection
Public Sub RemoveAllTickmarks()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim list As Variant, i As Long
    Dim cell As Range
    list = TickList()
    Application.ScreenUpdating = False
    For Each cell In Selection.Cells
        Dim txt As String
        txt = CStr(cell.Value)
        For i = LBound(list) To UBound(list)
            txt = Replace(txt, ChrW(CLng(list(i)(1))) & " ", "")
            txt = Replace(txt, ChrW(CLng(list(i)(1))), "")
        Next i
        cell.Value = txt
    Next cell
    Application.ScreenUpdating = True
End Sub

' Ribbon callbacks for individual ticks
Public Sub Tick1(): ApplyTick 1: End Sub
Public Sub Tick2(): ApplyTick 2: End Sub
Public Sub Tick3(): ApplyTick 3: End Sub
Public Sub Tick4(): ApplyTick 4: End Sub
Public Sub Tick5(): ApplyTick 5: End Sub
Public Sub Tick6(): ApplyTick 6: End Sub
Public Sub Tick7(): ApplyTick 7: End Sub
Public Sub Tick8(): ApplyTick 8: End Sub
Public Sub Tick9(): ApplyTick 9: End Sub
Public Sub Tick10(): ApplyTick 10: End Sub
Public Sub Tick11(): ApplyTick 11: End Sub
Public Sub Tick12(): ApplyTick 12: End Sub
Public Sub Tick13(): ApplyTick 13: End Sub
Public Sub Tick14(): ApplyTick 14: End Sub
Public Sub Tick15(): ApplyTick 15: End Sub

Private Sub ApplyTick(ByVal idx As Long)
    Dim list As Variant, entry As Variant
    list = TickList()
    If idx < 1 Or idx > UBound(list) - LBound(list) + 1 Then Exit Sub
    entry = list(LBound(list) + idx - 1)
    InsertTickIntoSelection CLng(entry(1)), CLng(entry(2))
End Sub

' Settings: tick position (LeftInside / RightInside)
Public Function GetTickPosition() As String
    GetTickPosition = GetSettingString("TickPosition", "LeftInside")
End Function

Public Sub SaveTickPosition(ByVal pos As String)
    SaveSettingString "TickPosition", pos
End Sub

Public Sub ShowTickSettings()
    Dim pos As String
    pos = InputBox("Tick position - type LEFT or RIGHT:", APP_NAME, "LEFT")
    If Len(pos) = 0 Then Exit Sub
    Select Case UCase$(Left$(pos, 1))
        Case "L": SaveTickPosition "LeftInside"
        Case "R": SaveTickPosition "RightInside"
    End Select
    MsgBox "Tick position saved: " & GetTickPosition(), vbInformation, APP_NAME
End Sub
