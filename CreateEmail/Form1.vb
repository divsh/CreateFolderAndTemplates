Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook

Public Class Form1
    Private Sub btnCreateEmail_Click(sender As Object, e As EventArgs) Handles btnCreateEmail.Click
        InsertBodyTextInOutlookWordEditor("C:\A1\work\autoDoc\259111 P2 - Linux machine disk space issue.docx")
    End Sub

    Sub CreateMail()
        Dim filePath As String
        filePath = """C:\\Users\\Me\\Desktop\\test.docx"""
        InsertBodyTextInOutlookWordEditor(filePath)
    End Sub

    Sub InsertBodyTextInOutlookWordEditor(filePath As String)
        Dim myMail As Outlook.MailItem
        Dim myInspector As Outlook.Inspector
        Dim wdDoc As Word.Document
        Dim wdRange As Word.Range
        Dim app As Outlook.Application = New Outlook.Application()

        wdDoc = (New Word.Application).Documents.Open(FileName:=filePath, [ReadOnly]:=True)
        On Error Resume Next
        myMail = app.CreateItem(OlItemType.olMailItem)
        myMail.Subject = "Here's the latest..."
        myMail.Display()
        myInspector = myMail.GetInspector
        wdDoc = myInspector.WordEditor
        If Not (wdDoc Is Nothing) Then
            wdRange = wdDoc.Content
            wdRange.Select()
            wdRange.Copy()

            ' myMail.HTMLBody = wdRange.Text
            wdRange.PasteSpecial()
            wdDoc.Close()

            'wdRange = wdDoc.Range(0, wdDoc.Characters.Count)
            'wdRange.Fields.Add(Range:=wdRange, Type:=Word.WdFieldType.wdFieldEmpty, Text:="INCLUDETEXT  " & filePath, PreserveFormatting:=True)
            'wdRange.Fields.Add(Range:=wdRange, Type:=Word.WdFieldType.wdFieldEmpty, Text:=filePath, PreserveFormatting:=True)
        End If

    End Sub



End Class
