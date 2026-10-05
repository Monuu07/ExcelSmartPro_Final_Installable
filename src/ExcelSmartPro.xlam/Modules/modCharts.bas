Attribute VB_Name = "modCharts"
Option Explicit

' =====================================================================
'  SMART CHART - auto chart from selected range with type choice
' =====================================================================

Public Sub SmartChart()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select chart data range (include headers):", _
        "Smart Chart", Selection.Address, Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    Dim chartTypeChoice As String
    chartTypeChoice = InputBox("Chart type - enter a number:" & vbCrLf & vbCrLf & _
        "1 = Column   2 = Bar   3 = Line   4 = Pie   5 = Area", "Smart Chart", "1")
    If Len(chartTypeChoice) = 0 Then Exit Sub

    Dim ct As Long
    Select Case Val(chartTypeChoice)
        Case 1: ct = xlColumnClustered
        Case 2: ct = xlBarClustered
        Case 3: ct = xlLine
        Case 4: ct = xlPie
        Case 5: ct = xlArea
        Case Else: ct = xlColumnClustered
    End Select

    Dim co As ChartObject
    Set co = ActiveSheet.ChartObjects.Add(Left:=400, Top:=50, Width:=480, Height:=280)
    co.Chart.SetSourceData Source:=rng
    co.Chart.ChartType = ct
    co.Chart.HasTitle = True
    On Error Resume Next
    co.Chart.ChartTitle.Text = CStr(rng.Cells(1, 1).Value) & " - Chart"
    On Error GoTo 0
End Sub

' =====================================================================
'  POWER DASHBOARD - KPI cards + charts from a data table
' =====================================================================

Public Sub PowerDashboard()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select your data table (headers + one numeric column at least):", _
        "Power Dashboard", Selection.Address, Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    Dim ws As Worksheet
    Dim lastR As Long, lastC As Long
    Dim i As Long, c As Long
    Dim numericCol As Long
    Dim sumVal As Double, avgVal As Double, minVal As Double, maxVal As Double, countVal As Long

    lastR = rng.Rows.Count
    lastC = rng.Columns.Count

    ' Find first numeric column (excluding header)
    numericCol = 0
    For c = 2 To lastC
        If Application.WorksheetFunction.IsNumber(rng.Cells(2, c).Value) Then
            numericCol = c
            Exit For
        End If
    Next c
    If numericCol = 0 Then numericCol = lastC

    ' Compute stats
    On Error Resume Next
    sumVal = Application.WorksheetFunction.Sum(rng.Columns(numericCol))
    avgVal = Application.WorksheetFunction.Average(rng.Columns(numericCol))
    minVal = Application.WorksheetFunction.Min(rng.Columns(numericCol))
    maxVal = Application.WorksheetFunction.Max(rng.Columns(numericCol))
    countVal = rng.Rows.Count - 1
    On Error GoTo 0

    ' Build dashboard sheet
    SafeOff
    Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    ws.Name = GetFreeSheetName("Dashboard")

    ' KPI cards
    ws.Range("B2:G2").Merge
    With ws.Range("B2")
        .Value = "ExcelSmart Pro Dashboard"
        .Font.Size = 16
        .Font.Bold = True
    End With

    Dim cardLabels As Variant, cardValues As Variant
    cardLabels = Array("Total", "Average", "Min", "Max", "Rows")
    cardValues = Array(sumVal, avgVal, minVal, maxVal, countVal)
    For i = 0 To 4
        With ws.Cells(4, i * 2 + 2)
            .Value = cardLabels(i)
            .Font.Bold = True
            .Font.Size = 11
            .HorizontalAlignment = xlCenter
        End With
        With ws.Cells(5, i * 2 + 2).MergeArea.Cells(1, 1)
        End With
        With ws.Cells(5, i * 2 + 2)
            .Value = Format(cardValues(i), "#,##0.##")
            .Font.Size = 18
            .Font.Bold = True
            .Font.Color = RGB(0, 112, 192)
            .HorizontalAlignment = xlCenter
        End With
    Next i

    ' Card background
    ws.Range("B4:C5, D4:E5, F4:G5, H4:I5, J4:K5").Interior.Color = RGB(242, 242, 242)

    ' Chart 1: column chart of the numeric column
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=20, Top:=110, Width:=420, Height:=260)
    co.Chart.SetSourceData Source:=rng.Columns(numericCol)
    co.Chart.ChartType = xlColumnClustered
    co.Chart.HasTitle = True
    On Error Resume Next
    co.Chart.ChartTitle.Text = CStr(rng.Cells(1, numericCol).Value)
    On Error GoTo 0

    ' Chart 2: pie of the first text column (if it exists)
    If lastC >= 2 Then
        Dim co2 As ChartObject
        Set co2 = ws.ChartObjects.Add(Left:=450, Top:=110, Width:=420, Height:=260)
        co2.Chart.SetSourceData Source:=rng.Resize(lastR, numericCol)
        co2.Chart.ChartType = xlPie
        co2.Chart.HasTitle = True
        On Error Resume Next
        co2.Chart.ChartTitle.Text = "Distribution"
        On Error GoTo 0
    End If

    SafeOn
    ws.Activate
    MsgBox "Dashboard created on '" & ws.Name & "'.", vbInformation, APP_NAME
End Sub

' =====================================================================
'  EXPORT TOOLS (JSON / HTML / CSV)
' =====================================================================

Public Sub ExportAsJSON()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select data to export as JSON:", "Export JSON", Selection.Address, Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    Dim vData As Variant
    Dim r As Long, c As Long
    Dim json As String
    Dim headers() As String

    vData = rng.Value
    ReDim headers(1 To rng.Columns.Count)
    For c = 1 To UBound(vData, 2)
        headers(c) = CStr(vData(1, c))
    Next c

    json = "["
    For r = 2 To UBound(vData, 1)
        json = json & "{"
        For c = 1 To UBound(vData, 2)
            json = json & Q() & headers(c) & Q() & ":"
            If IsNumeric(vData(r, c)) And Len(CStr(vData(r, c))) > 0 Then
                json = json & CStr(vData(r, c))
            Else
                json = json & Q() & EscapeJson(CStr(vData(r, c))) & Q()
            End If
            If c < UBound(vData, 2) Then json = json & ","
        Next c
        json = json & "}"
        If r < UBound(vData, 1) Then json = json & ","
    Next r
    json = json & "]"

    ' Write to file
    Dim p As String
    p = Environ$("USERPROFILE") & "\Desktop\ExcelSmartPro_Export.json"
    Dim fno As Integer
    fno = FreeFile
    Open p For Output As #fno
    Print #fno, json
    Close #fno

    MsgBox "JSON saved to:" & vbCrLf & p, vbInformation, APP_NAME
End Sub

Public Sub ExportAsHTML()
    Dim rng As Range
    On Error Resume Next
    Set rng = Application.InputBox("Select data to export as HTML table:", "Export HTML", Selection.Address, Type:=8)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub

    Dim vData As Variant
    Dim r As Long, c As Long
    vData = rng.Value
    Dim body As String
    body = "<tr>"
    For c = 1 To UBound(vData, 2)
        body = body & "<th>" & CStr(vData(1, c)) & "</th>"
    Next c
    body = body & "</tr>"
    For r = 2 To UBound(vData, 1)
        body = body & "<tr>"
        For c = 1 To UBound(vData, 2)
            body = body & "<td>" & CStr(vData(r, c)) & "</td>"
        Next c
        body = body & "</tr>"
    Next r
    Dim html As String
    html = "<html><head><style>table{border-collapse:collapse;font-family:Segoe UI,Arial}" & _
        "th{background:#4472C4;color:#fff;padding:6px}td{border:1px solid #ddd;padding:6px}" & _
        "</style></head><body><table>" & body & "</table></body></html>"

    Dim p As String
    p = Environ$("USERPROFILE") & "\Desktop\ExcelSmartPro_Export.html"
    Dim fno As Integer
    fno = FreeFile
    Open p For Output As #fno
    Print #fno, html
    Close #fno

    MsgBox "HTML saved to:" & vbCrLf & p, vbInformation, APP_NAME
End Sub

Private Function EscapeJson(ByVal s As String) As String
    Dim out As String
    out = Replace(s, "\", "\\")
    out = Replace(out, Q(), "\" & Q())
    out = Replace(out, vbCrLf, "\n")
    out = Replace(out, vbLf, "\n")
    EscapeJson = out
End Function

' CSV split by delimiter into columns (Data > Text-to-Columns alternative)
Public Sub SplitByDelimiter()
    Dim delim As String
    delim = InputBox("Delimiter character (e.g. , or ; or | or tab):", "Split By Delimiter", ",")
    If Len(delim) = 0 Then Exit Sub
    If delim = "tab" Then delim = vbTab

    Dim rng As Range, cell As Range
    Set rng = TargetRangeOrSelection
    Dim parts() As String
    Dim i As Long
    SafeOff
    For Each cell In rng.Cells
        parts = Split(CStr(cell.Value), delim)
        For i = 0 To UBound(parts)
            cell.Offset(0, i + 1).Value = Trim$(parts(i))
        Next i
    Next cell
    SafeOn
End Sub
