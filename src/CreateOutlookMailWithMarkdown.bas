Attribute VB_Name = "CreateOutlookMailWithMarkdown"
'==============================================================================================================
' Create Outlook Mail with Markdown
'==============================================================================================================
Option Explicit


'---------------------------------------------------------------------------------
' Display
'---------------------------------------------------------------------------------
Sub Display(sheetName As String, fontnameCell As String, fontsizeCell As String, toCell As String, ccCell As String, bccCell As String, subjectCell As String, greetingCell As String, bodyCell As String, signatureCell As String)
    Dim fontName As String
    Dim fontSize As String
    Dim addressTo As String
    Dim addressCc As String
    Dim addressBcc As String
    Dim subject As String
    Dim greeting As Variant
    Dim body As Variant
    Dim signature As Variant
    Dim htmlBody As Variant
    
    ' Get Font Name
    fontName = GetCellValue(sheetName, fontnameCell)
    
    ' Get Font Size
    fontSize = GetCellValue(sheetName, fontsizeCell)
    
    ' Get To
    addressTo = GetCellValue(sheetName, toCell)

    ' Get Cc
    addressCc = GetCellValue(sheetName, ccCell)
    
    ' Get Bcc
    addressBcc = GetCellValue(sheetName, bccCell)
    
    ' Get Subject
    subject = GetCellValue(sheetName, subjectCell)

    ' Get Gereeting
    greeting = Split(GetCellValue(sheetName, greetingCell), vbLf)
    
    ' Get Body
    body = Split(GetCellValue(sheetName, bodyCell), vbLf)
    
    ' Get Signature
    signature = Split(GetCellValue(sheetName, signatureCell), vbLf)
    
    ' Create HTML Body
    htmlBody = CreateHtmlBody(fontSize, fontName, greeting, body, signature)

    ' Display Mail
    Call DisplayMail(addressTo, addressCc, addressBcc, subject, htmlBody)
End Sub


'---------------------------------------------------------------------------------
' Create HTML Body with Markdown
'---------------------------------------------------------------------------------
Function CreateHtmlBody(fontSize As String, fontName As String, greeting As Variant, body As Variant, signature As Variant)
    CreateHtmlBody = ""
    CreateHtmlBody = CreateHtmlBody & "<font size=""" & fontSize & """ face=""" & fontName & """>"
    CreateHtmlBody = CreateHtmlBody & LineArrange.AddBrTag(greeting)
    CreateHtmlBody = CreateHtmlBody & Markdown.Parse(body)
    CreateHtmlBody = CreateHtmlBody & LineArrange.AddBrTag(signature)
    CreateHtmlBody = CreateHtmlBody & "</font>"
End Function


'---------------------------------------------------------------------------------
' Display Mail
'---------------------------------------------------------------------------------
Sub DisplayMail(addressTo As String, addressCc As String, addressBcc As String, subject As String, htmlBody As Variant)
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    On Error Resume Next
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    With OutlookMail
        .to = addressTo
        .CC = addressCc
        .BCC = addressBcc
        .subject = subject
        .htmlBody = htmlBody
        .Display
    End With
    
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
End Sub


'---------------------------------------------------------------------------------
' Get Cell Value
'---------------------------------------------------------------------------------
Function GetCellValue(sheetName As String, cellPosition As String)
    GetCellValue = Sheets(sheetName).Range(cellPosition).Value
End Function


'==============================================================================================================
' END
'==============================================================================================================
