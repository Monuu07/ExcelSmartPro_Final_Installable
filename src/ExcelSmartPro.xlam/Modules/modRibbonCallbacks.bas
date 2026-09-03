Attribute VB_Name = "modRibbonCallbacks"
Option Explicit

Public gRibbon As IRibbonUI

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Sub RunMacro(control As IRibbonControl)
    Select Case control.Id
        Case "btnAIFormulaGenerator": AIFormulaGenerator
        Case "btnAIFormulaExplainer": AIFormulaExplainer
        Case "btnAIFormulaErrorDetector": AIFormulaErrorDetector
        Case "btnAIPythonScriptGenerator": AIPythonScriptGenerator
        Case "btnAICommentGenerator": AICommentGenerator
        Case "btnIncreaseColour": IncreaseColour
        Case "btnDecreaseColour": DecreaseColour
        Case "btnRedHighlight": RedHighlight
        Case "btnGreenHighlight": GreenHighlight
        Case "btnOrangeHighlight": OrangeHighlight
        Case "btnBlueHighlight": BlueHighlight
        Case "btnClearHighlight": ClearHighlight
        Case "btnChangeBorderColor": ChangeBorderColor
        Case "btnStatusTickmarks": StatusTickmarks
        Case "btnAddCharacters": AddCharacters
        Case "btnDeleteCharacters": DeleteCharacters
        Case "btnCountCharacters": CountCharacters
        Case "btnSplitNames": SplitNames
        Case "btnFillBlankCells": FillBlankCells
        Case "btnColumnHide": ColumnHide
        Case "btnColumnUnhide": ColumnUnhide
        Case "btnUpperCase": UPPERCASE
        Case "btnLowerCase": lowercase
        Case "btnProperCase": ProperCase
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
        Case "btnBackupFile": BackupFile
        Case "btnProtectSheets": ProtectSheets
        Case "btnUnprotectSheets": UnprotectSheets
        Case "btnProtectWorkbook": ProtectWorkbook
        Case "btnUnprotectWorkbook": UnprotectWorkbook
        Case "btnLockFormulas": LockFormulas
        Case "btnUnlockFormulas": UnlockFormulas
        Case "btnLockSelection": LockSelection
        Case "btnUnlockSelection": UnlockSelection
        Case "btnEmailValidation": EmailValidation
        Case "btnPhoneValidation": PhoneValidation
        Case "btnNumericOnly": NumericOnly
        Case "btnTextOnly": TextOnly
        Case "btnURLValidation": URLValidation
        Case "btnDateValidation": DateValidation
        Case "btnRemoveValidation": RemoveValidation
        Case "btnSmartChart": SmartChart
        Case "btnPowerDashboard": PowerDashboard
        Case "btnSmartDataExporter": SmartDataExporter
        Case "btnMailAutomation": MailAutomation
    End Select
End Sub
