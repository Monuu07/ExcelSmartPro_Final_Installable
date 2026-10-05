Attribute VB_Name = "modFilters"
Option Explicit

' =====================================================================
'  SMART FILTER - XLBooster-style multi-column filter by selection
'  Select cells in one or more columns, run this, and every selected
'  value becomes a filter criterion for its column.
' =====================================================================

Public Sub ApplyFilterBySelection()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim filterRng As Range, intersectRng As Range
    Dim area As Range
    Dim colValues As Collection, uniqueCols As Collection
    Dim arr() As String
    Dim i As Long, colIndex As Long, c As Long, r As Long
    Dim currentTargetCol As Long, vCol As Variant, relCol As Long
    Dim validSelection As Boolean
    Dim vData As Variant
    Dim cellVal As String

    If TypeName(Selection) <> "Range" Then Exit Sub

    Set ws = ActiveSheet
    validSelection = False

    ' --- Scenario A: Excel Table (ListObject) ---
    Set tbl = ActiveCell.ListObject
    If Not tbl Is Nothing Then
        Set filterRng = tbl.Range
    Else
        ' --- Scenario B: Normal Range ---
        Set filterRng = ActiveCell.CurrentRegion
        If filterRng.Rows.Count <= 1 And filterRng.Columns.Count <= 1 Then
            MsgBox "Please select data cells so the filter can be applied properly.", vbExclamation, "Apply Filter"
            Exit Sub
        End If
        ' Turn on AutoFilter if it's off
        If Not ws.AutoFilterMode Then filterRng.AutoFilter
        Set filterRng = ws.AutoFilter.Range
    End If

    ' Ensure selection is inside the data table
    Set intersectRng = Intersect(Selection, filterRng)
    If intersectRng Is Nothing Then
        MsgBox "Please select data cells inside your table.", vbExclamation, "Apply Filter"
        Exit Sub
    End If

    SafeOff

    ' 1. Find all unique column indexes in the user's selection (handles Ctrl+Click areas)
    Set uniqueCols = New Collection
    On Error Resume Next
    For Each area In intersectRng.Areas
        For c = 1 To area.Columns.Count
            uniqueCols.Add area.Column + c - 1, CStr(area.Column + c - 1)
        Next c
    Next area
    On Error GoTo 0

    ' 2. Loop through each unique selected column
    For Each vCol In uniqueCols
        currentTargetCol = CLng(vCol)
        Set colValues = New Collection

        ' Gather unique values using memory arrays (fast)
        For Each area In intersectRng.Areas
            If currentTargetCol >= area.Column And currentTargetCol < area.Column + area.Columns.Count Then
                relCol = currentTargetCol - area.Column + 1

                If area.Cells.Count = 1 Then
                    If Not IsError(area.Value) Then
                        cellVal = Trim$(CStr(area.Value))
                        If Len(cellVal) > 0 Then
                            On Error Resume Next
                            colValues.Add cellVal, cellVal
                            On Error GoTo 0
                        End If
                    End If
                Else
                    vData = area.Value
                    On Error Resume Next ' Ignore duplicates
                    For r = 1 To UBound(vData, 1)
                        If Not IsError(vData(r, relCol)) Then
                            cellVal = Trim$(CStr(vData(r, relCol)))
                            If Len(cellVal) > 0 Then
                                colValues.Add cellVal, cellVal
                            End If
                        End If
                    Next r
                    On Error GoTo 0
                End If
            End If
        Next area

        ' 3. Apply filter to THIS specific column
        If colValues.Count > 0 Then
            validSelection = True

            ReDim arr(0 To colValues.Count - 1)
            For i = 1 To colValues.Count
                arr(i - 1) = colValues(i)
            Next i

            colIndex = currentTargetCol - filterRng.Column + 1

            On Error Resume Next
            If Not tbl Is Nothing Then
                tbl.Range.AutoFilter Field:=colIndex, Criteria1:=arr, Operator:=xlFilterValues
            Else
                ws.AutoFilter.Range.AutoFilter Field:=colIndex, Criteria1:=arr, Operator:=xlFilterValues
            End If
            On Error GoTo 0
        End If
    Next vCol

    SafeOn

    If Not validSelection Then
        MsgBox "Please select data cells so the filter can be applied properly.", vbExclamation, "Apply Filter"
    End If
End Sub

' =====================================================================
'  CLEAR ALL FILTERS
' =====================================================================

Public Sub ClearAllFilters()
    Dim ws As Worksheet
    Dim tbl As ListObject

    Set ws = ActiveSheet

    SafeOff
    On Error Resume Next

    ' 1. Clear normal range filters
    If ws.FilterMode Then
        ws.ShowAllData
    End If

    ' 2. Clear Excel table filters
    For Each tbl In ws.ListObjects
        If tbl.AutoFilter.FilterMode Then
            tbl.AutoFilter.ShowAllData
        End If
    Next tbl

    On Error GoTo 0
    SafeOn
End Sub

' =====================================================================
'  SMART HIGHLIGHTER - highlight rows matching values the user types
' =====================================================================

Public Sub StartSmartHighlighter()
    Dim searchText As String
    searchText = InputBox("Enter text to highlight matching cells:", "Smart Highlighter")
    If Len(searchText) = 0 Then Exit Sub

    Dim ws As Worksheet, rng As Range, cell As Range
    Set ws = ActiveSheet
    Set rng = ws.UsedRange

    SafeOff
    Dim foundCount As Long
    For Each cell In rng.Cells
        If Not IsError(cell.Value) Then
            If InStr(1, CStr(cell.Value), searchText, vbTextCompare) > 0 Then
                cell.Interior.Color = RGB(255, 235, 156)
                foundCount = foundCount + 1
            End If
        End If
    Next cell
    SafeOn

    MsgBox foundCount & " cell(s) highlighted containing '" & searchText & "'.", vbInformation, APP_NAME
End Sub

Public Sub ClearSmartHighlighter()
    If Not AskYesNo("Remove all yellow highlight from the active sheet's used range?") Then Exit Sub
    SafeOff
    On Error Resume Next
    Dim cell As Range
    For Each cell In ActiveSheet.UsedRange.Cells
        If cell.Interior.Color = RGB(255, 235, 156) Then cell.Interior.Pattern = xlNone
    Next cell
    On Error GoTo 0
    SafeOn
    MsgBox "Highlight removed.", vbInformation, APP_NAME
End Sub
