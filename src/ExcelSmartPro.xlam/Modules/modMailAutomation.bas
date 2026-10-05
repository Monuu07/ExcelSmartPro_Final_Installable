Attribute VB_Name = "modMailAutomation"
Option Explicit

' =====================================================================
'  MAIL AUTOMATION - send email from Excel via Outlook or Gmail SMTP
'  Table layout required: To | Subject | Body  (headers in row 1)
' =====================================================================

Public Sub MailAutomation()
    Dim choice As VbMsgBoxResult
    choice = MsgBox("Send emails via:" & vbCrLf & vbCrLf & _
        "YES = Outlook (recommended if installed)" & vbCrLf & _
        "NO = Gmail SMTP (needs app password)", vbQuestion + vbYesNoCancel, APP_NAME)
    If choice = vbCancel Then Exit Sub

    If choice = vbYes Then
        SendBulkMailOutlook
    Else
        SendBulkMailGmail
    End If
End Sub

Private Function GetMailTable() As Range
    On Error Resume Next
    Dim tbl As ListObject
    Set tbl = ActiveCell.ListObject
    If Not tbl Is Nothing Then
        Set GetMailTable = tbl.Range
    Else
        Set GetMailTable = ActiveCell.CurrentRegion
    End If
    On Error GoTo 0
    If GetMailTable Is Nothing Then
        MsgBox "Select a cell inside your mail list table first." & vbCrLf & _
            "Required columns: To | Subject | Body", vbExclamation, APP_NAME
    End If
End Function

Public Sub SendBulkMailOutlook()
    Dim rng As Range
    Set rng = GetMailTable()
    If rng Is Nothing Then Exit Sub
    If rng.Rows.Count < 2 Then
        MsgBox "No data rows found in the selected table.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim outlookApp As Object
    On Error Resume Next
    Set outlookApp = CreateObject("Outlook.Application")
    If outlookApp Is Nothing Then
        MsgBox "Microsoft Outlook is not available.", vbExclamation, APP_NAME
        Exit Sub
    End If
    On Error GoTo 0

    Dim r As Long
    Dim sent As Long
    For r = 2 To rng.Rows.Count
        If Len(Trim$(CStr(rng.Cells(r, 1).Value))) > 0 Then
            Dim mail As Object
            Set mail = outlookApp.CreateItem(0)
            mail.To = CStr(rng.Cells(r, 1).Value)
            mail.Subject = CStr(rng.Cells(r, 2).Value)
            mail.Body = CStr(rng.Cells(r, 3).Value)
            On Error Resume Next
            mail.Display False ' show draft for review instead of blind-send
            On Error GoTo 0
            sent = sent + 1
            If sent >= 3 Then
                If Not AskYesNo("3 drafts created. Continue with the remaining " & (rng.Rows.Count - r) & " rows?") Then Exit Sub
                sent = 0
            End If
        End If
    Next r
    MsgBox "Draft emails created in Outlook. Review and press Send in Outlook.", vbInformation, APP_NAME
End Sub

' Gmail SMTP via CDO - needs a Google App Password
Public Sub SendBulkMailGmail()
    Dim rng As Range
    Set rng = GetMailTable()
    If rng Is Nothing Then Exit Sub
    If rng.Rows.Count < 2 Then
        MsgBox "No data rows found in the selected table.", vbExclamation, APP_NAME
        Exit Sub
    End If

    Dim gmailUser As String, gmailPass As String
    gmailUser = InputBox("Your Gmail address:", "Gmail SMTP", GetSettingString("GmailUser", ""))
    If StrPtr(gmailUser) = 0 Then Exit Sub
    gmailPass = InputBox("Your Gmail App Password (not your normal password - create one at myaccount.google.com > Security > App passwords):", "Gmail SMTP")
    If StrPtr(gmailPass) = 0 Then Exit Sub

    SaveSettingString "GmailUser", gmailUser

    Dim msg As Object, conf As Object
    Dim r As Long
    Dim sent As Long, failed As Long

    Set msg = CreateObject("CDO.Message")
    Set conf = msg.Configuration
    conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
    conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = gmailUser
    conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = gmailPass
    conf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    conf.Fields.Update

    For r = 2 To rng.Rows.Count
        If Len(Trim$(CStr(rng.Cells(r, 1).Value))) > 0 Then
            On Error Resume Next
            msg.To = CStr(rng.Cells(r, 1).Value)
            msg.From = gmailUser
            msg.Subject = CStr(rng.Cells(r, 2).Value)
            msg.TextBody = CStr(rng.Cells(r, 3).Value)
            msg.Send
            If Err.Number = 0 Then
                sent = sent + 1
            Else
                failed = failed + 1
                Debug.Print "Mail failed row " & r & ": " & Err.Description
            End If
            On Error GoTo 0
        End If
    Next r

    MsgBox "Sent: " & sent & IIf(failed > 0, "   Failed: " & failed, ""), vbInformation, APP_NAME
End Sub
