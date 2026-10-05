Attribute VB_Name = "modSheetManager"
Option Explicit

' =====================================================================
'  TEXT TOOLS
' =====================================================================

Public Sub AddCharacters()
    Dim txt As String, side As String, cell As Range
    txt = PromptText("Add Characters", "Characters to add")
    If txt = "" Then Exit Sub
    side = LCase$(PromptText("Add Characters", "Type LEFT or RIGHT", "RIGHT"))
    SafeOff
    For Each cell In TargetRangeOrSelection.Cells
        If side = "left" Then
            cell.Value = txt & cell.Value
        Else
            cell.Value = cell.Value & txt
        End If
    Next cell
    SafeOn
End Sub

Public Sub DeleteCharacters()
    Dim n As Long, side As String, cell As Range, s As String
    n = CLng(Val(PromptText("Delete Characters", "How many characters", "1")))
    side = LCase$(PromptText("Delete Characters", "Type LEFT or RIGHT", "RIGHT"))
    SafeOff
    For Each cell In TargetRangeOrSelection.Cells
        s = CStr(cell.Value)
        If side = "left" Then
            cell.Value = Mid$(s, n + 1)
        Else
            cell.Value = Left$(s, Application.Max(Len(s) - n, 0))
        End If
    Next cell
    SafeOn
End Sub

Public Sub CountCharacters()
    Dim cell As Range
    For Each cell In TargetRangeOrSelection.Cells
        cell.Offset(0, 1).Value = Len(CStr(cell.Value))
    Next cell
End Sub

Public Sub SplitNames()
    Dim cell As Range, parts() As String
    SafeOff
    For Each cell In TargetRangeOrSelection.Cells
        parts = Split(Trim$(CStr(cell.Value)), " ")
        If UBound(parts) >= 0 Then cell.Offset(0, 1).Value = parts(0)
        If UBound(parts) >= 1 Then cell.Offset(0, 2).Value = parts(UBound(parts))
    Next cell
    SafeOn
End Sub

Public Sub FillBlankCells()
    On Error Resume Next
    TargetRangeOrSelection.SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
    TargetRangeOrSelection.Value = TargetRangeOrSelection.Value
    On Error GoTo 0
End Sub

' Advanced trim: trims leading/trailing spaces AND collapses double spaces
Public Sub TrimSelectedText()
    Dim cell As Range
    SafeOff
    For Each cell In TargetRangeOrSelection.Cells
        If Not IsError(cell.Value) Then
            cell.Value = Application.WorksheetFunction.Trim(CStr(cell.Value))
        End If
    Next cell
    SafeOn
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
    Dim cell As Range
    SafeOff
    For Each cell In TargetRangeOrSelection.Cells
        Select Case mode
            Case "U": cell.Value = UCase$(CStr(cell.Value))
            Case "L": cell.Value = LCase$(CStr(cell.Value))
            Case "P": cell.Value = StrConv(CStr(cell.Value), vbProperCase)
        End Select
    Next cell
    SafeOn
End Sub

' =====================================================================
'  SHEET TOOLS
' =====================================================================

Public Sub AutoFitSheets()
    Dim ws As Worksheet
    Application.ScreenUpdating = False
    For Each ws In ActiveWorkbook.Worksheets
        ws.Cells.EntireColumn.AutoFit
    Next ws
    Application.ScreenUpdating = True
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
    Application.ScreenUpdating = False
    For i = 1 To Worksheets.Count - 1
        For j = i + 1 To Worksheets.Count
            If (ascending And UCase$(Worksheets(j).Name) < UCase$(Worksheets(i).Name)) Or _
               (Not ascending And UCase$(Worksheets(j).Name) > UCase$(Worksheets(i).Name)) Then
                Worksheets(j).Move Before:=Worksheets(i)
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
End Sub

Public Sub MonthlySheets()
    Dim arr As Variant, i As Long
    arr = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    For i = LBound(arr) To UBound(arr)
        If Not SheetExists(CStr(arr(i))) Then
            Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = CStr(arr(i))
        End If
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
    Application.ScreenUpdating = False
    For r = lastR To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Rows(r)) = 0 Then ws.Rows(r).Delete
    Next r
    Application.ScreenUpdating = True
End Sub

Public Sub RemoveBlankColumns()
    Dim ws As Worksheet, c As Long, lastC As Long
    Set ws = ActiveSheet
    lastC = LastUsedCol(ws)
    Application.ScreenUpdating = False
    For c = lastC To 1 Step -1
        If Application.WorksheetFunction.CountA(ws.Columns(c)) = 0 Then ws.Columns(c).Delete
    Next c
    Application.ScreenUpdating = True
End Sub

Public Sub RemoveBlankSheets()
    Dim i As Long, ws As Worksheet
    For i = Worksheets.Count To 1 Step -1
        Set ws = Worksheets(i)
        If Application.WorksheetFunction.CountA(ws.Cells) = 0 And Worksheets.Count > 1 Then ws.Delete
    Next i
End Sub

' =====================================================================
'  SHEET INDEX / DASHBOARD (XLBooster-style gallery index)
' =====================================================================

Public Sub CreateSheetIndex()
    Dim wsIndex As Worksheet
    Dim ws As Worksheet
    Dim r As Long
    Dim name As String
    Dim startTime As Double

    startTime = Timer
    name = GetFreeSheetName("Index")

    Application.ScreenUpdating = False
    Set wsIndex = Worksheets.Add(Before:=Worksheets(1))
    wsIndex.Name = name

    ' Title
    With wsIndex.Range("A1")
        .Value = "Sheet Index - " & Format(Now, "dd-mmm-yyyy hh:nn")
        .Font.Size = 14
        .Font.Bold = True
    End With
    wsIndex.Range("A1:G1").Merge

    ' Column headers
    With wsIndex.Range("A3:G3")
        .Value = Array("#", "Sheet Name", "Rows", "Columns", "Type", "Visibility", "Open Link")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
    End With

    ' Rows
    r = 4
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> wsIndex.Name Then
            wsIndex.Cells(r, 1).Value = r - 3
            wsIndex.Cells(r, 2).Value = ws.Name
            wsIndex.Cells(r, 3).Value = ws.UsedRange.Rows.Count
            wsIndex.Cells(r, 4).Value = ws.UsedRange.Columns.Count
            wsIndex.Cells(r, 5).Value = IIf(ws.ListObjects.Count > 0, "Table", "Worksheet")
            wsIndex.Cells(r, 6).Value = IIf(ws.Visible = xlSheetVisible, "Visible", "Hidden")
            wsIndex.Hyperlinks.Add Anchor:=wsIndex.Cells(r, 7), _
                Address:="", SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:="Open"
            r = r + 1
        End If
    Next ws

    ' Formatting
    wsIndex.Columns("A:G").AutoFit
    wsIndex.Columns(2).ColumnWidth = 25
    With wsIndex.Range("A3:G" & r - 1).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(180, 180, 180)
    End With

    ' Footer
    wsIndex.Cells(r + 1, 2).Value = "Total sheets: " & (r - 4) & "   |   Created in " & _
        Format(Timer - startTime, "0.0") & " seconds"
    wsIndex.Cells(r + 1, 2).Font.Italic = True

    Application.ScreenUpdating = True
    wsIndex.Activate
    wsIndex.Range("A1").Select
    MsgBox "Sheet index created on '" & name & "'." & vbCrLf & _
           "Tip: use 'Update Index' to refresh it anytime.", vbInformation, APP_NAME
End Sub

' Remove all back-button shapes / clean the index
Public Sub DeleteSheetFromIndex()
    Dim ws As Worksheet
    On Error Resume Next
    For Each ws In ActiveWorkbook.Worksheets
        If StrComp(Left$(ws.Name, 5), "Index", vbTextCompare) = 0 Then
            If AskYesNo("Delete index sheet '" & ws.Name & "'?") Then ws.Delete
        End If
    Next ws
    On Error GoTo 0
End Sub

' Merge (append) all sheets' data into one sheet
Public Sub MergeSheetsAppendData()
    Dim ws As Worksheet, wsOut As Worksheet
    Dim lastR As Long, lastC As Long
    Dim outR As Long
    Dim srcRange As Range
    Dim headersDone As Boolean

    If Not AskYesNo("This will copy data from every sheet into a new 'MergedData' sheet (headers kept once). Continue?") Then Exit Sub

    Application.ScreenUpdating = False
    Set wsOut = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    wsOut.Name = GetFreeSheetName("MergedData")
    outR = 1
    headersDone = False

    For Each ws In ActiveWorkbook.Worksheets
        If ws.Name <> wsOut.Name Then
            lastR = LastUsedRow(ws)
            lastC = LastUsedCol(ws)
            If lastR > 0 And lastC > 0 And Application.WorksheetFunction.CountA(ws.Cells) > 0 Then
                If Not headersDone Then
                    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastC)).Copy wsOut.Cells(outR, 1)
                    outR = outR + 1
                    headersDone = True
                End If
                If lastR >= 2 Then
                    ws.Range(ws.Cells(2, 1), ws.Cells(lastR, lastC)).Copy wsOut.Cells(outR, 1)
                    outR = wsOut.Cells(wsOut.Rows.Count, 1).End(xlUp).Row + 1
                End If
            End If
        End If
    Next ws

    wsOut.Columns.AutoFit
    Application.ScreenUpdating = True
    wsOut.Activate
    MsgBox "Merged data is ready on '" & wsOut.Name & "'.", vbInformation, APP_NAME
End Sub

' Unpivot a cross-tab table into flat list (XLBooster-style)
Public Sub UnpivotData()
    Dim src As Range
    Dim ws As Worksheet
    Dim vData As Variant
    Dim r As Long, c As Long
    Dim outR As Long
    Dim headerRow As Variant

    On Error Resume Next
    Set src = Application.InputBox("Select the whole table INCLUDING headers (row + column):", _
        "Unpivot Data", Selection.Address, Type:=8)
    On Error GoTo 0
    If src Is Nothing Then Exit Sub

    SafeOff
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = GetFreeSheetName("Unpivot")
    ws.Range("A1:C1").Value = Array("Row Label", "Column Label", "Value")
    ws.Range("A1:C1").Font.Bold = True
    outR = 2

    vData = src.Value
    For r = 2 To UBound(vData, 1)
        For c = 2 To UBound(vData, 2)
            If Len(CStr(vData(r, c))) > 0 Then
                ws.Cells(outR, 1).Value = vData(r, 1)
                ws.Cells(outR, 2).Value = vData(1, c)
                ws.Cells(outR, 3).Value = vData(r, c)
                outR = outR + 1
            End If
        Next c
    Next r

    ws.Columns.AutoFit
    SafeOn
    ws.Activate
    MsgBox "Unpivoted " & (outR - 2) & " rows on '" & ws.Name & "'.", vbInformation, APP_NAME
End Sub
