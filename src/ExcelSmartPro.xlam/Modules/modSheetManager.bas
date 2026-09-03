Attribute VB_Name = "modSheetManager"
Option Explicit

Public Sub AddCharacters()
    Dim txt As String, side As String, c As Range
    txt = PromptText("Add Characters", "Characters to add")
    If txt = "" Then Exit Sub
    side = LCase$(PromptText("Add Characters", "Type LEFT or RIGHT", "RIGHT"))
    For Each c In TargetRangeOrSelection.Cells
        If side = "left" Then
            c.Value = txt & c.Value
        Else
            c.Value = c.Value & txt
        End If
    Next c
End Sub

Public Sub DeleteCharacters()
    Dim n As Long, side As String, c As Range, s As String
    n = CLng(Val(PromptText("Delete Characters", "How many characters", "1")))
    side = LCase$(PromptText("Delete Characters", "Type LEFT or RIGHT", "RIGHT"))
    For Each c In TargetRangeOrSelection.Cells
        s = CStr(c.Value)
        If side = "left" Then
            c.Value = Mid$(s, n + 1)
        Else
            c.Value = Left$(s, Application.Max(Len(s) - n, 0))
        End If
    Next c
End Sub

Public Sub CountCharacters()
    Dim c As Range
    For Each c In TargetRangeOrSelection.Cells
        c.Offset(0, 1).Value = Len(CStr(c.Value))
    Next c
End Sub

Public Sub SplitNames()
    Dim c As Range, parts() As String
    For Each c In TargetRangeOrSelection.Cells
        parts = Split(Trim$(CStr(c.Value)), " ")
        If UBound(parts) >= 0 Then c.Offset(0, 1).Value = parts(0)
        If UBound(parts) >= 1 Then c.Offset(0, 2).Value = parts(UBound(parts))
    Next c
End Sub

Public Sub FillBlankCells()
    On Error Resume Next
    TargetRangeOrSelection.SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
    TargetRangeOrSelection.Value = TargetRangeOrSelection.Value
    On Error GoTo 0
End Sub

Public Sub ColumnHide()
    TargetRangeOrSelection.EntireColumn.Hidden = True
End Sub

Public Sub ColumnUnhide()
    TargetRangeOrSelection.EntireColumn.Hidden = False
End Sub

Public Sub UPPERCASE()
    TransformSelectionText "U"
End Sub

Public Sub lowercase()
    TransformSelectionText "L"
End Sub

Public Sub ProperCase()
    TransformSelectionText "P"
End Sub

Private Sub TransformSelectionText(ByVal mode As String)
    Dim c As Range
    For Each c In TargetRangeOrSelection.Cells
        Select Case mode
            Case "U": c.Value = UCase$(CStr(c.Value))
            Case "L": c.Value = LCase$(CStr(c.Value))
            Case "P": c.Value = StrConv(CStr(c.Value), vbProperCase)
        End Select
    Next c
End Sub

Public Sub AutoFitSheets()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.EntireColumn.AutoFit
        ws.Cells.EntireRow.AutoFit
    Next ws
End Sub

Public Sub HideSheets()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> ActiveSheet.Name Then ws.Visible = xlSheetHidden
    Next ws
End Sub

Public Sub UnhideSheets()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
End Sub

Public Sub HideShapes()
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        shp.Visible = msoFalse
    Next shp
End Sub

Public Sub UnhideShapes()
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        shp.Visible = msoTrue
    Next shp
End Sub

Public Sub SelectShapes()
    On Error Resume Next
    ActiveSheet.Shapes.SelectAll
    On Error GoTo 0
End Sub

Public Sub ValuesToTextFormat()
    With TargetRangeOrSelection
        .NumberFormat = "@"
        .Value = .Value
    End With
End Sub

Public Sub DuplicateSheet()
    ActiveSheet.Copy After:=ActiveSheet
End Sub

Public Sub SortSheetsAZ()
    SortSheets True
End Sub

Public Sub SortSheetsZA()
    SortSheets False
End Sub

Private Sub SortSheets(ByVal ascending As Boolean)
    Dim i As Long, j As Long
    For i = 1 To Worksheets.Count - 1
        For j = i + 1 To Worksheets.Count
            If (ascending And UCase$(Worksheets(j).Name) < UCase$(Worksheets(i).Name)) Or _
               (Not ascending And UCase$(Worksheets(j).Name) > UCase$(Worksheets(i).Name)) Then
                Worksheets(j).Move Before:=Worksheets(i)
            End If
        Next j
    Next i
End Sub

Public Sub MonthlySheets()
    Dim arr, i As Long
    arr = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    For i = LBound(arr) To UBound(arr)
        If Not SheetExists(CStr(arr(i))) Then Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = CStr(arr(i))
    Next i
End Sub

Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    For Each ws In Worksheets
        If UCase$(ws.Name) = UCase$(sheetName) Then SheetExists = True: Exit Function
    Next ws
End Function

Public Sub RemoveBlankRows()
    Dim ws As Worksheet, r As Long, lastR As Long
    Set ws = ActiveSheet
    lastR = LastUsedRow(ws)
    For r = lastR To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then ws.Rows(r).Delete
    Next r
End Sub

Public Sub RemoveBlankColumns()
    Dim ws As Worksheet, c As Long, lastC As Long
    Set ws = ActiveSheet
    lastC = LastUsedCol(ws)
    For c = lastC To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Columns(c)) = 0 Then ws.Columns(c).Delete
    Next c
End Sub

Public Sub RemoveBlankSheets()
    Dim i As Long, ws As Worksheet
    For i = Worksheets.Count To 1 Step -1
        Set ws = Worksheets(i)
        If Application.WorksheetFunction.CountA(ws.Cells) = 0 And Worksheets.Count > 1 Then ws.Delete
    Next i
End Sub
