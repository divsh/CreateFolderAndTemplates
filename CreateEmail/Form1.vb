Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Outlook
Imports System.IO

Public Class Form1
    Private Sub btnCreateEmail_Click(sender As Object, e As EventArgs) Handles btnCreateEmail.Click
        openEmailClient()
        'InsertBodyTextInOutlookWordEditor("C:\A1\work\autoDoc\259111 P2 - Linux machine disk space issue.docx")
        'insertWordInMail("C:\A1\work\autoDoc\259111 P2 - Linux machine disk space issue.docx")
    End Sub

    Sub CreateMail()
        Dim filePath As String
        filePath = """C:\\Users\\Me\\Desktop\\test.docx"""
        filePath = "C:\test.docx"
        'InsertBodyTextInOutlookWordEditor(filePath)
        insertWordInMail(filePath)
    End Sub

    Sub openEmailClient()
        Dim myMail As Outlook.MailItem
        myMail = (New Outlook.Application()).CreateItem(OlItemType.olMailItem)
        myMail.Subject = "Here's the latest..."
        setupEmailClient(myMail)
        myMail.Display()
    End Sub

    Sub setAppSettings()
        My.Settings.directoryname = Path.GetDirectoryName(Environment.CurrentDirectory)
        My.Settings.Save()
    End Sub

    Sub setupEmailClient(ByVal mail As MailItem)
        grabAttachments(Environment.CurrentDirectory).ForEach(Sub(x) mail.Attachments.Add(x))

        '  mail.sendus = My.Settings.EmailFrom
        mail.To = My.Settings.EmailTo
        mail.Subject = subtitutePlaceHolder(My.Settings.EmailSubject)
        mail.BCC = My.Settings.EmailBcc
    End Sub

    Function grabAttachments(sourcefolderPath As String) As List(Of String)
        Dim files As List(Of String) = New List(Of String)
        files.AddRange(Directory.GetFiles(sourcefolderPath, My.Settings.Attachment1))
        files.AddRange(Directory.GetFiles(sourcefolderPath, My.Settings.Attachment2))
        files.AddRange(Directory.GetFiles(sourcefolderPath, My.Settings.Attachment3))
        files.AddRange(Directory.GetFiles(sourcefolderPath, My.Settings.Attachment4))
        files.AddRange(Directory.GetFiles(sourcefolderPath, My.Settings.Attachment5))
        files.AddRange(Directory.GetFiles(sourcefolderPath, My.Settings.Attachment6))
        files.AddRange(Directory.GetFiles(sourcefolderPath, My.Settings.Attachment7))
        files.AddRange(Directory.GetFiles(sourcefolderPath, My.Settings.Attachment8))
        files.AddRange(Directory.GetFiles(sourcefolderPath, My.Settings.Attachment9))
        files.AddRange(Directory.GetFiles(sourcefolderPath, My.Settings.Attachment10))
        Return files
    End Function

    Function subtitutePlaceHolder(Text As String) As String
        If String.IsNullOrEmpty(Text) Then Return Text
        Return Text.Replace(Text, "@directoryname")
    End Function


    Sub InsertBodyTextInOutlookWordEditor(filePath As String)
        Dim myMail As Outlook.MailItem
        Dim myInspector As Outlook.Inspector
        Dim wdDoc As Word.Document
        Dim wdRange As Word.Range
        Dim app As Outlook.Application = New Outlook.Application()
        Dim wapp As Word.Application = New Word.Application

        wdDoc = wapp.Documents.Open(FileName:=filePath, [ReadOnly]:=True)

        On Error Resume Next
        myMail = app.CreateItem(OlItemType.olMailItem)
        myMail.Subject = "Here's the latest..."
        myMail.Display()


        myInspector = myMail.GetInspector
        wdDoc = myInspector.WordEditor
        If Not (wdDoc Is Nothing) Then
            wdRange = wdDoc.Selection.WholeStory


            wdRange.Copy()

            ' myMail.HTMLBody = wdRange.Text
            wdRange.PasteSpecial()
            AppActivate("Message")


            myMail.Body = wapp.Selection
            'wdRange = wdDoc.Range(0, wdDoc.Characters.Count)
            'wdRange.Fields.Add(Range:=wdRange, Type:=Word.WdFieldType.wdFieldEmpty, Text:="INCLUDETEXT  " & filePath, PreserveFormatting:=True)
            'wdRange.Fields.Add(Range:=wdRange, Type:=Word.WdFieldType.wdFieldEmpty, Text:=filePath, PreserveFormatting:=True)
        End If

    End Sub


    Sub insertWordInMail(filepath As String)
        Dim oWord As Word.Application
        Dim oDoc As Word.Document
        Dim oEmail As MailItem
        Dim oOutlook As Outlook.Application = New Outlook.Application()
        Dim strBody As String

        oWord = CreateObject("Word.Application")

        oDoc = oWord.Documents.Open(filepath, , True)

        'Get the text for the body of the Email
        'oWord.ActiveDocument.SaveAs(FileName:="HTMLDoc.html" FileFormat:=wdFormatHTML
        'strBody = oDoc.HTMLProject.HTMLProjectItems(1).Text
        'strBody = oDoc.Content.Copy()
        'Open Outlook and create new email
        oOutlook = New Outlook.Application
        oEmail = oOutlook.CreateItem(OlItemType.olMailItem)
        With oEmail
            .To = "recipient@email.com"
            .Subject = "Document.doc"
            '.HTMLBody = oWord.Selection


        End With
    End Sub


End Class
