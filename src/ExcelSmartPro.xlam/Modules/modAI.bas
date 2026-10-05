Attribute VB_Name = "modAI"
Option Explicit

' =====================================================================
'  AI TOOLS - HTTP API based (works with OpenAI-compatible endpoints)
'  User sets API key + endpoint once via AI Settings; stored in registry.
'  Without a key, tools show a friendly setup message.
' =====================================================================

Private Const AI_SETTINGS_KEY As String = "AI_API_Key"
Private Const AI_SETTINGS_URL As String = "AI_API_Url"
Private Const AI_SETTINGS_MODEL As String = "AI_Model"
Private Const DEFAULT_URL As String = "https://api.openai.com/v1/chat/completions"
Private Const DEFAULT_MODEL As String = "gpt-4o-mini"

#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
        Alias "URLDownloadToFileA" ( _
        ByVal pCaller As LongPtr, ByVal szURL As String, _
        ByVal szFileName As String, ByVal dwReserved As Long, _
        ByVal lpfnCB As LongPtr) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" _
        Alias "URLDownloadToFileA" ( _
        ByVal pCaller As Long, ByVal szURL As String, _
        ByVal szFileName As String, ByVal dwReserved As Long, _
        ByVal lpfnCB As Long) As Long
#End If

Public Sub AISettings()
    Dim key As String, url As String, model As String
    key = GetSettingString(AI_SETTINGS_KEY, "")
    url = GetSettingString(AI_SETTINGS_URL, DEFAULT_URL)
    model = GetSettingString(AI_SETTINGS_MODEL, DEFAULT_MODEL)

    key = InputBox("Enter your OpenAI-compatible API key (stored locally in your registry):", _
        "AI Settings", key)
    If StrPtr(key) = 0 Then Exit Sub ' cancelled
    SaveSettingString AI_SETTINGS_KEY, key

    url = InputBox("API endpoint URL:", "AI Settings", url)
    If StrPtr(url) = 0 Then Exit Sub
    SaveSettingString AI_SETTINGS_URL, url

    model = InputBox("Model name:", "AI Settings", model)
    If StrPtr(model) = 0 Then Exit Sub
    SaveSettingString AI_SETTINGS_MODEL, model

    MsgBox "AI settings saved.", vbInformation, APP_NAME
End Sub

Private Function HasAISetup() As Boolean
    HasAISetup = Len(GetSettingString(AI_SETTINGS_KEY, "")) > 0
End Function

Private Function BuildChatRequest(ByVal prompt As String, ByVal systemMsg As String) As String
    Dim body As String
    body = "{" & _
        Q() & "model" & Q() & ":" & Q() & GetSettingString(AI_SETTINGS_MODEL, DEFAULT_MODEL) & Q() & "," & _
        Q() & "messages" & Q() & ":[{" & _
        Q() & "role" & Q() & ":" & Q() & "system" & Q() & "," & _
        Q() & "content" & Q() & ":" & Q() & systemMsg & Q() & "},{" & _
        Q() & "role" & Q() & ":" & Q() & "user" & Q() & "," & _
        Q() & "content" & Q() & ":" & Q() & prompt & Q() & "}]," & _
        Q() & "temperature" & Q() & ":0.2}"
    BuildChatRequest = body
End Function

' POST JSON via MSXML2.XMLHTTP.60 (all modern Windows have it).
Private Function PostJson(ByVal url As String, ByVal body As String) As String
    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/json"
    http.setRequestHeader "Authorization", "Bearer " & GetSettingString(AI_SETTINGS_KEY, "")
    On Error Resume Next
    http.Send body
    If Err.Number <> 0 Then
        PostJson = ""
        Exit Function
    End If
    On Error GoTo 0
    If http.Status = 200 Then
        PostJson = http.responseText
    Else
        PostJson = ""
    End If
End Function

' Extract "content" from a chat completion JSON response (minimal parser).
Private Function ExtractContent(ByVal json As String) As String
    Dim marker As String
    marker = Q() & "content" & Q() & ":" & Q()
    Dim p As Long
    p = InStr(json, marker)
    If p = 0 Then Exit Function
    p = p + Len(marker)
    Dim out As String
    Dim ch As String
    Dim i As Long
    i = p
    Do While i <= Len(json)
        ch = Mid$(json, i, 1)
        If ch = Q() Then
            If Mid$(json, i, 2) = Q() & Q() Then ' escaped quote
                out = out & Q()
                i = i + 2
            Else
                Exit Do
            End If
        ElseIf ch = "\" And Mid$(json, i + 1, 1) = "n" Then
            out = out & vbLf
            i = i + 2
        ElseIf ch = "\" And Mid$(json, i + 1, 1) = "\" Then
            out = out & "\"
            i = i + 2
        Else
            out = out & ch
            i = i + 1
        End If
    Loop
    ExtractContent = out
End Function

Private Function AskAI(ByVal prompt As String, ByVal systemMsg As String) As String
    If Not HasAISetup() Then
        MsgBox "Please set your API key first (AI Tools > AI Settings).", vbExclamation, APP_NAME
        Exit Function
    End If
    Dim json As String
    json = PostJson(GetSettingString(AI_SETTINGS_URL, DEFAULT_URL), BuildChatRequest(prompt, systemMsg))
    If Len(json) = 0 Then
        MsgBox "The AI service did not respond. Check your internet connection, API key, and endpoint URL (AI Settings).", vbExclamation, APP_NAME
        Exit Function
    End If
    AskAI = ExtractContent(json)
End Function

' ---------------------------------------------------------------------
'  RIBBON FEATURES
' ---------------------------------------------------------------------

Public Sub AIFormulaGenerator()
    Dim desc As String
    desc = InputBox("Describe what you want the formula to do:", "AI Formula Generator")
    If Len(desc) = 0 Then Exit Sub
    Dim answer As String
    answer = AskAI(desc, "You are an Excel formula expert. Reply with the exact Excel formula only, no explanation, no markdown.")
    If Len(answer) = 0 Then Exit Sub
    ' Offer to insert directly into the active cell
    If AskYesNo("Formula:" & vbCrLf & vbCrLf & answer & vbCrLf & vbCrLf & "Insert into the active cell?") Then
        ActiveCell.Formula = answer
    End If
End Sub

Public Sub AIFormulaExplainer()
    Dim f As String
    On Error Resume Next
    f = CStr(ActiveCell.Formula)
    On Error GoTo 0
    If Len(f) = 0 Then
        f = InputBox("Paste or type the formula to explain:", "AI Formula Explainer")
        If Len(f) = 0 Then Exit Sub
    End If
    Dim answer As String
    answer = AskAI("Explain this Excel formula in simple steps: " & f, "You explain Excel formulas simply and clearly in short numbered steps.")
    If Len(answer) > 0 Then MsgBox answer, vbInformation, "Formula Explanation"
End Sub

Public Sub AIFormulaErrorDetector()
    Dim f As String
    On Error Resume Next
    f = CStr(ActiveCell.Formula)
    On Error GoTo 0
    If Len(f) = 0 Then
        MsgBox "The active cell does not contain a formula.", vbExclamation, APP_NAME
        Exit Sub
    End Sub
    Dim answer As String
    answer = AskAI("Find problems in this Excel formula and suggest the corrected version: " & f, "You are an Excel formula debugger. Reply with the issue and the corrected formula.")
    If Len(answer) > 0 Then MsgBox answer, vbInformation, "Formula Check"
End Sub

Public Sub AIPythonScriptGenerator()
    Dim desc As String
    desc = InputBox("Describe the Python script you need (pandas/openpyxl style for Excel tasks):", "AI Python Script Generator")
    If Len(desc) = 0 Then Exit Sub
    Dim answer As String
    answer = AskAI(desc, "You write short, clean Python scripts for Excel automation using pandas and openpyxl. Reply with code only.")
    If Len(answer) = 0 Then Exit Sub
    ' Write to a .py file next to the workbook
    Dim p As String
    p = Environ$("USERPROFILE") & "\Desktop\ExcelSmartPro_AI_Script.py"
    On Error GoTo ErrHandler
    Dim fno As Integer
    fno = FreeFile
    Open p For Output As #fno
    Print #fno, answer
    Close #fno
    MsgBox "Python script saved to:" & vbCrLf & p, vbInformation, APP_NAME
    Exit Sub
ErrHandler:
    MsgBox "Could not save the file: " & Err.Description, vbExclamation, APP_NAME
End Sub

Public Sub AICommentGenerator()
    ' Generate a cell comment from data context
    Dim context As String
    On Error Resume Next
    context = "Cell " & ActiveCell.Address(False, False) & " contains: " & CStr(ActiveCell.Value) & _
        ". Column header: " & CStr(ActiveCell.EntireColumn.Cells(1, 1).Value)
    On Error GoTo 0
    Dim answer As String
    answer = AskAI("Write a one-line Excel cell comment for this data: " & context, "You write very short, useful cell comments. One line maximum.")
    If Len(answer) = 0 Then Exit Sub
    On Error Resume Next
    ActiveCell.ClearComments
    ActiveCell.AddComment answer
    On Error GoTo 0
    MsgBox "Comment added: " & answer, vbInformation, APP_NAME
End Sub

Public Sub ShowAIForm()
    Dim choice As VbMsgBoxResult
    choice = MsgBox("AI Tools need an API key from an OpenAI-compatible service." & vbCrLf & vbCrLf & _
        "Open AI Settings now to enter your key?", vbQuestion + vbYesNo, APP_NAME)
    If choice = vbYes Then AISettings
End Sub

' ---------------------------------------------------------------------
'  UDFs (worksheet functions)
' ---------------------------------------------------------------------

Public Function XL_AskAI(ByVal prompt As String) As Variant
    If Not HasAISetup() Then
        XL_AskAI = "Set API key in AI Settings first"
        Exit Function
    End If
    Dim answer As String
    answer = AskAI(prompt, "Answer briefly.")
    If Len(answer) = 0 Then
        XL_AskAI = CVErr(xlErrValue)
    Else
        XL_AskAI = answer
    End If
End Function

Public Function XL_MagicFormula(ByVal prompt As String) As Variant
    Dim answer As String
    answer = AskAI(prompt, "Reply with the exact Excel formula only, no explanation.")
    If Len(answer) = 0 Then
        XL_MagicFormula = CVErr(xlErrValue)
    Else
        XL_MagicFormula = answer
    End If
End Function

Public Function XL_RegexExtract(ByVal text As String, ByVal pattern As String) As String
    Dim re As Object, m As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = pattern
    re.Global = False
    If re.Test(text) Then
        Set m = re.Execute(text)(0)
        XL_RegexExtract = m.Value
    End If
End Function

Public Function XL_Translate(ByVal text As String, Optional ByVal targetLanguage As String = "Hindi") As Variant
    Dim answer As String
    answer = AskAI("Translate to " & targetLanguage & ": " & text, "You are a translator. Reply with the translation only.")
    If Len(answer) = 0 Then
        XL_Translate = CVErr(xlErrValue)
    Else
        XL_Translate = answer
    End If
End Function
