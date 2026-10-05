Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

Public gRibbon As IRibbonUI

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Sub RunMacro(control As IRibbonControl)
    On Error GoTo ErrHandler
    Select Case control.Id

        Case "btnAIFormulaGenerator": modAI.AIFormulaGenerator
        Case "btnAIFormulaExplainer": modAI.AIFormulaExplainer
        Case "btnAIFormulaErrorDetector": modAI.AIFormulaErrorDetector
        Case "btnAIPythonScriptGenerator": modAI.AIPythonScriptGenerator
        Case "btnAICommentGenerator": modAI.AICommentGenerator
        Case "btnAISettings": modAI.AISettings

        Case "btnAddCharacters": AddCharacters
        Case "btnDeleteCharacters": DeleteCharacters
        Case "btnCountCharacters": CountCharacters
        Case "btnSplitNames": SplitNames
        Case "btnFillBlankCells": FillBlankCells
        Case "btnTrimText": TrimSelectedText
        Case "btnSplitByDelimiter": SplitByDelimiter
        Case "btnUpperCase": UPPERCASE
        Case "btnLowerCase": lowercase
        Case "btnProperCase": ProperCase

        Case "btnApplyFilterBySelection": ApplyFilterBySelection
        Case "btnClearAllFilters": ClearAllFilters
        Case "btnStartSmartHighlighter": StartSmartHighlighter
        Case "btnClearSmartHighlighter": ClearSmartHighlighter

        Case "btnIncreaseColour": IncreaseColour
        Case "btnDecreaseColour": DecreaseColour
        Case "btnRedHighlight": RedHighlight
        Case "btnGreenHighlight": GreenHighlight
        Case "btnOrangeHighlight": OrangeHighlight
        Case "btnBlueHighlight": BlueHighlight
        Case "btnClearHighlight": ClearHighlight
        Case "btnChangeBorderColor": ChangeBorderColor

        Case "btnCreateSheetIndex": CreateSheetIndex
        Case "btnMergeSheets": MergeSheetsAppendData
        Case "btnUnpivotData": UnpivotData
        Case "btnAutoFitSheets": AutoFitSheets
        Case "btnHideSheets": HideSheets
        Case "btnUnhideSheets": UnhideSheets
        Case "btnHideShapes": HideShapes
        Case "btnUnhideShapes": UnhideShapes
        Case "btnSelectShapes": SelectShapes
        Case "btnValuesToText": ValuesToTextFormat
        Case "btnDuplicateSheet": DuplicateSheet
        Case "btnSortAZ": SortSheetsAZ
        Case "btnSortZA": SortSheetsZA
        Case "btnMonthlySheets": MonthlySheets
        Case "btnRemoveBlankRows": RemoveBlankRows
        Case "btnRemoveBlankColumns": RemoveBlankColumns
        Case "btnRemoveBlankSheets": RemoveBlankSheets
        Case "btnColumnHide": ColumnHide
        Case "btnColumnUnhide": ColumnUnhide

        Case "btnStatusTickmarks": StatusTickmarks
        Case "btnTickRed": modTickmarks.Tick2
        Case "btnTickQuestion": modTickmarks.Tick3
        Case "btnTickWarning": modTickmarks.Tick8
        Case "btnTickStar": modTickmarks.Tick12
        Case "btnTickFlagRed": modTickmarks.Tick5
        Case "btnTickFlagGreen": modTickmarks.Tick6
        Case "btnTickFlagYellow": modTickmarks.Tick7
        Case "btnTickClock": modTickmarks.Tick9
        Case "btnTickPending": modTickmarks.Tick10
        Case "btnTickInfo": modTickmarks.Tick11
        Case "btnTickLock": modTickmarks.Tick13
        Case "btnTickBlocked": modTickmarks.Tick14
        Case "btnTick15": modTickmarks.Tick15
        Case "btnRemoveAllTicks": RemoveAllTickmarks
        Case "btnTickSettings": ShowTickSettings

        Case "btnBackupFile": BackupFile
        Case "btnProtectSheets": ProtectSheets
        Case "btnUnprotectSheets": UnprotectSheets
        Case "btnProtectWorkbook": ProtectWorkbook
        Case "btnUnprotectWorkbook": UnprotectWorkbook

        Case "btnEmailValidation": EmailValidation
        Case "btnPhoneValidation": PhoneValidation
        Case "btnNumericOnly": NumericOnly
        Case "btnTextOnly": TextOnly
        Case "btnURLValidation": URLValidation
        Case "btnDateValidation": DateValidation
        Case "btnRemoveValidation": RemoveValidation
        Case "btnValidateEmailData": ValidateEmailData
        Case "btnValidatePhoneData": ValidatePhoneData

        Case "btnSmartChart": SmartChart
        Case "btnPowerDashboard": PowerDashboard
        Case "btnExportAsJSON": ExportAsJSON
        Case "btnExportAsHTML": ExportAsHTML

        Case "btnExportSheetsToPDF": ExportSheetsToPDF
        Case "btnExportSheetsToWorkbooks": ExportSheetsToWorkbooks
        Case "btnConvertAllCSVtoExcel": ConvertAllCSVtoExcel
        Case "btnGetFileNamesFromFolder": GetFileNamesFromFolder
        Case "btnSeparateFilesByType": SeparateFilesByType
        Case "btnAutoFileOrganizer": AutoFileOrganizer
        Case "btnConvertWordPagesToPDFs": ConvertWordPagesToPDFs

        Case "btnMailAutomation": MailAutomation
        Case "btnAbout": ShowAbout
    End Select
    Exit Sub
ErrHandler:
    MsgBox "Error in " & control.Id & ": " & Err.Description, vbCritical, APP_NAME
End Sub

Public Sub ShowAbout()
    MsgBox "ExcelSmart Pro" & vbCrLf & vbCrLf & _
        "All-in-one Excel productivity add-in." & vbCrLf & _
        "AI tools, smart filters, tickmarks, sheet manager, file automation, mail automation, validation tools, and more." & vbCrLf & vbCrLf & _
        "Version 2.0 (local build)", vbInformation, APP_NAME
End Sub
