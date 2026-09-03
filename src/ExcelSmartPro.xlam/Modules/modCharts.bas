Attribute VB_Name = "modCharts"
Option Explicit

Public Sub SmartChart()
    Dim rng As Range
    Set rng = Application.InputBox("Select chart data range", "Smart Chart", Selection.Address, Type:=8)
    If rng Is Nothing Then Exit Sub
    Dim co As ChartObject
    Set co = ActiveSheet.ChartObjects.Add(Left:=300, Top:=50, Width:=480, Height:=280)
    co.Chart.SetSourceData Source:=rng
    co.Chart.ChartType = xlColumnClustered
    co.Chart.HasTitle = True
    co.Chart.ChartTitle.Text = "ExcelSmart Pro Chart"
End Sub

Public Sub PowerDashboard()
    MsgBox "Dashboard wizard placeholder. Next build can generate KPI cards, charts, and slicer-ready summary sheets.", vbInformation, "ExcelSmart Pro"
End Sub

Public Sub AutoFileOrganizer()
    MsgBox "Automation placeholder. Planned: bulk file collection, workbook merge, export, and scheduled reporting.", vbInformation, "ExcelSmart Pro"
End Sub
