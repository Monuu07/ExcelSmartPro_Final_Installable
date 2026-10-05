Attribute VB_Name = "modFileTools"
Option Explicit

' =====================================================================
'  FILE & FOLDER AUTOMATION TOOLS (XLBooster-style)
' =====================================================================

Public Sub ExportSheetsToPDF()
    Dim ws As Worksheet
    Dim chosen As String
    Dim sheetNames As Collection
    Dim i As Long

    Set sheetNames = New Collection
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then sheetNames.Add ws.Name
    Next ws

    ' Build list string
    Dim listStr As String
    For i = 1 To sheetNames.Count
        listStr = listStr & i & " - " & sheetNames(i) & vbCrLf
    Next i
    chosen = InputBox("Enter sheet numbers to export, separated by commas (e.g. 1,3,5):" & vbCrLf & vbCrLf & listStr, _
        "Export Sheets To PDF")
    If Len(chosen) = 0 Then Exit Sub

    Dim parts() As String
    parts = Split(Replace(chosen, " ", ""), ",")
    Dim selected() As String
    ReDim selected(0 To UBound(parts))
    For i = 0 To UBound(parts)
        If CLng(Val(parts(i))) >= 1 And CLng(Val(parts(i))) <= sheetNames.Count Then
            selected(i) = sheetNames(CLng(Val(parts(i))))
        End If
    Next i

    Dim folderPath As String
    folderPath = ActiveWorkbook.Path
    If Len(folderPath) = 0 Then
        MsgBox "Please save the workbook first (PDF will use its folder).", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim pdfPath As String
    pdfPath = folderPath & Application.PathSeparator & _
        Left$(ActiveWorkbook.Name, InStrRev(ActiveWorkbook.Name, ".") - 1) & "_" & _
        Format(Now, "yyyymmdd_hhnnss") & ".pdf"

    SafeOff
    On Error GoTo ErrHandler
    ActiveWorkbook.Sheets(selected).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=pdfPath, Quality:=xlQualityStandard
    ActiveWorkbook.Worksheets(1).Select
    SafeOn

    MsgBox "PDF saved:" & vbCrLf & pdfPath, vbInformation, APP_NAME
    Exit Sub
ErrHandler:
    SafeOn
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, APP_NAME
End Sub

Public Sub ExportSheetsToWorkbooks()
    Dim ws As Worksheet
    Dim folderPath As String

    If Len(ActiveWorkbook.Path) = 0 Then
        MsgBox "Please save the workbook first.", vbExclamation, APP_NAME
        Exit Sub
    End If
    If Not AskYesNo("Every sheet will be saved as a separate .xlsx file in the same folder. Continue?") Then Exit Sub

    folderPath = ActiveWorkbook.Path
    SafeOff
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            ws.Copy
            ActiveWorkbook.SaveAs Filename:=folderPath & Application.PathSeparator & ws.Name & ".xlsx", _
                FileFormat:=xlWorkbookDefault
            ActiveWorkbook.Close SaveChanges:=False
        End If
    Next ws
    SafeOn
    MsgBox "All sheets exported to separate workbooks in:" & vbCrLf & folderPath, vbInformation, APP_NAME
End Sub

' Convert every .csv in a chosen folder to .xlsx
Public Sub ConvertAllCSVtoExcel()
    Dim folderPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim wb As Workbook

    folderPath = BrowseForFolder("Select folder containing CSV files")
    If Len(folderPath) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    SafeOff
    Dim converted As Long
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "csv" Then
            Set wb = Workbooks.Open(file.Path)
            wb.SaveAs Filename:=folderPath & Application.PathSeparator & _
                fso.GetBaseName(file.Name) & ".xlsx", FileFormat:=xlWorkbookDefault
            wb.Close SaveChanges:=False
            converted = converted + 1
        End If
    Next file
    SafeOn

    MsgBox converted & " CSV file(s) converted to .xlsx.", vbInformation, APP_NAME
End Sub

' Get file names from a folder into a new sheet
Public Sub GetFileNamesFromFolder()
    Dim folderPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim ws As Worksheet
    Dim r As Long

    folderPath = BrowseForFolder("Select folder to list files from")
    If Len(folderPath) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    Set ws = Worksheets.Add
    ws.Name = GetFreeSheetName("FileList")
    ws.Range("A1:D1").Value = Array("File Name", "Extension", "Size (KB)", "Modified")
    ws.Range("A1:D1").Font.Bold = True

    r = 2
    For Each file In folder.Files
        ws.Cells(r, 1).Value = file.Name
        ws.Cells(r, 2).Value = "." & fso.GetExtensionName(file.Name)
        ws.Cells(r, 3).Value = Round(file.Size / 1024, 1)
        ws.Cells(r, 4).Value = file.DateLastModified
        r = r + 1
    Next file

    ws.Columns.AutoFit
    MsgBox (r - 2) & " file(s) listed on '" & ws.Name & "'.", vbInformation, APP_NAME
    ws.Activate
End Sub

' Separate files into subfolders by extension type
Public Sub SeparateFilesByType()
    Dim folderPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim destPath As String
    Dim ext As String

    folderPath = BrowseForFolder("Select folder to organize")
    If Len(folderPath) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    Dim moved As Long
    For Each file In folder.Files
        ext = LCase$(fso.GetExtensionName(file.Name))
        If Len(ext) > 0 Then
            destPath = folderPath & Application.PathSeparator & UCase$(ext) & " Files"
            If Not fso.FolderExists(destPath) Then fso.CreateFolder destPath
            On Error Resume Next
            fso.MoveFile file.Path, destPath & Application.PathSeparator & file.Name
            If Err.Number = 0 Then moved = moved + 1
            On Error GoTo 0
        End If
    Next file

    MsgBox moved & " file(s) moved into type folders.", vbInformation, APP_NAME
End Sub

' Auto File Organizer: organize chosen folder by year/month
Public Sub AutoFileOrganizer()
    Dim folderPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim destPath As String
    Dim subName As String

    folderPath = BrowseForFolder("Select folder to auto-organize by year and month")
    If Len(folderPath) = 0 Then Exit Sub

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    Dim moved As Long
    For Each file In folder.Files
        subName = Year(file.DateLastModified) & "\" & Format(file.DateLastModified, "mmm")
        destPath = folderPath & Application.PathSeparator & subName
        If Not fso.FolderExists(destPath) Then fso.CreateFolder destPath
        On Error Resume Next
        fso.MoveFile file.Path, destPath & Application.PathSeparator & file.Name
        If Err.Number = 0 Then moved = moved + 1
        On Error GoTo 0
    Next file

    MsgBox moved & " file(s) organized into year/month folders.", vbInformation, APP_NAME
End Sub

' Word documents to PDF (batch) - uses Word if installed
Public Sub ConvertWordPagesToPDFs()
    Dim folderPath As String
    Dim fso As Object, folder As Object, file As Object
    Dim wordApp As Object
    Dim doc As Object

    folderPath = BrowseForFolder("Select folder containing Word documents")
    If Len(folderPath) = 0 Then Exit Sub

    On Error Resume Next
    Set wordApp = CreateObject("Word.Application")
    If wordApp Is Nothing Then
        MsgBox "Microsoft Word is not available on this computer.", vbExclamation, APP_NAME
        Exit Sub
    End If
    On Error GoTo 0
    wordApp.Visible = False

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)

    Dim converted As Long
    On Error Resume Next
    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Name)) = "docx" Or _
           LCase$(fso.GetExtensionName(file.Name)) = "doc" Then
            Set doc = wordApp.Documents.Open(file.Path, ReadOnly:=True)
            If Not doc Is Nothing Then
                doc.ExportAsFixedFormat OutputFileName:=fso.GetBaseName(file.Path) & ".pdf", _
                    ExportFormat:=17 ' wdExportFormatPDF
                doc.Close False
                converted = converted + 1
            End If
        End If
    Next file
    wordApp.Quit
    On Error GoTo 0

    MsgBox converted & " Word document(s) converted to PDF.", vbInformation, APP_NAME
End Sub
