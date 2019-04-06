Attribute VB_Name = "CreateOutlookMailWithMarkdown"
'==============================================================================================================
' Create Outlook Mail with Markdown
'==============================================================================================================
Option Explicit


'---------------------------------------------------------------------------------
' Display
'---------------------------------------------------------------------------------
Sub Display(fontName As String, fontSize As String, sheetName As String, toCell As String, ccCell As String, bccCell As String, subjectCell As String, greetingCell As String, bodyCell As String, signatureCell As String)
    Dim address_to As String
    Dim address_cc As String
    Dim address_bcc As String
    Dim subject As String
    Dim greeting As Variant
    Dim body As Variant
    Dim signature As Variant
    Dim htmlbody As Variant
    
    ' Get To
    address_to = GetCellValue(sheetName, toCell)
    
    ' Get Cc
    address_cc = GetCellValue(sheetName, ccCell)
    
    ' Get Bcc
    address_bcc = GetCellValue(sheetName, bccCell)
    
    ' Get Subject
    subject = GetCellValue(sheetName, subjectCell)

    ' Get Gereeting
    greeting = Split(GetCellValue(sheetName, greetingCell), vbLf)
    
    ' Get Body
    body = Split(GetCellValue(sheetName, bodyCell), vbLf)
    
    ' Get Signature
    signature = Split(GetCellValue(sheetName, signatureCell), vbLf)
    
    ' Create HTML Body
    htmlbody = CreateHtmlBody(fontSize, fontName, greeting, body, signature)

    ' Display Mail
    Call DisplayMail(address_to, address_cc, address_bcc, subject, htmlbody)
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
Sub DisplayMail(address_to As String, address_cc As String, address_bcc As String, subject As String, htmlbody As Variant)
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    On Error Resume Next
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    
    With OutlookMail
        .to = address_to
        .CC = address_cc
        .BCC = address_bcc
        .subject = subject
        .htmlbody = htmlbody
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
