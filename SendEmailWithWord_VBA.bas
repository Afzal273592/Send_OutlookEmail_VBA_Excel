Attribute VB_Name = "SendEmailWithWord_VBA"
Option Explicit

Sub CreateEmailWithWord()

    Dim ol As Outlook.Application
    Dim mi As Outlook.MailItem
    Dim doc As Word.Document
    Dim MsgText As String
    
    Set ol = New Outlook.Application
    Set mi = ol.CreateItem(olMailItem)
    
    mi.Display
    mi.Subject = "Movies"
    mi.To = "someone@somewhere.com"
    
    Set doc = mi.GetInspector.WordEditor
    
    MsgText = vbNewLine & vbNewLine & "Please reply with any questions."
    doc.Range(0, 0).InsertAfter Text:=MsgText
    
    Sheet1.ChartObjects("Chart 1").Chart.ChartArea.Copy
    doc.Range(0, 0).Paste

    MsgText = vbNewLine & vbNewLine & "Please see the chart below:" & vbNewLine & vbNewLine
    doc.Range(0, 0).InsertAfter Text:=MsgText

    Sheet1.Range("A1").CurrentRegion.Copy
    doc.Range(1, 1).Paste

    MsgText = _
        "Dear Someone," & vbNewLine & vbNewLine _
        & "Please see the table below:" & vbNewLine

    doc.Range(0, 0).InsertBefore Text:=MsgText
    
End Sub
