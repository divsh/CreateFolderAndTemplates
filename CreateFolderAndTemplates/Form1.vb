Imports System.IO
Imports System.Collections.Generic
Imports Microsoft.Office.Tools.Word
Imports System.Net.Mail
Public Class Form1

    Private Const ClientFolderPath As String = "C:\A1\work\autoDoc\clients"
    Private Const TemplateFilesPath As String = "C:\A1\work\autoDoc\template"
    Private Const trelloBoardListEmailID As String = "a1x115+1ltukpnyj7ynskse8vmr@boards.trello.com"

    Private Sub btnCopyTemplates_Click(sender As Object, e As EventArgs) Handles btnCopyTemplates.Click
        If NoName() Then
            MessageBox.Show("Please provide at least one name")
            Return
        End If
        Dim clientFolderName As String
        clientFolderName = generateClientFolderName()
        copyFiles(Nothing, clientFolderName)
        createTrelloCard(clientFolderName)
    End Sub


    Private Sub createTrelloCard(clientFolderName As String)
        Dim t As Threading.Thread = New Threading.Thread(Sub() sendEmailToTrelloBoard(subject:=clientFolderName, trelloBoardListEmailID:=trelloBoardListEmailID))
        t.IsBackground = True
        t.Start()

        'sendEmailToTrelloBoard(subject:=clientFolderName, trelloBoardListEmailID:=trelloBoardListEmailID)
    End Sub

    Private Sub sendEmailToTrelloBoard(subject As String, trelloBoardListEmailID As String)
        Try
            txtName1.Invoke(Sub() txtName1.Text = "set by background thread")
            Dim mail As MailMessage = New MailMessage()
            Dim SmtpServer As SmtpClient = New SmtpClient("smtp.gmail.com")
            'Dim SmtpServer As SmtpClient = New SmtpClient("smtp.microsoft.com")
            '        mail.From = New MailAddress("divsh007@gmail.com")
            mail.From = New MailAddress("divsh@live.com")
            mail.To.Add(trelloBoardListEmailID)
            mail.Subject = subject
            mail.Body = "test body"
            SmtpServer.Port = 587
            SmtpServer.Credentials = New System.Net.NetworkCredential("divsh007@gmail.com", "Golmal^19")
            SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network
            SmtpServer.EnableSsl = True
            SmtpServer.Send(mail)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Function NoName() As Boolean
        Dim anyName As Boolean = False
        anyName = Not String.IsNullOrEmpty(txtName1.Text)
        anyName = Not String.IsNullOrEmpty(txtName2.Text)
        anyName = Not String.IsNullOrEmpty(txtName3.Text)
        anyName = Not String.IsNullOrEmpty(txtName4.Text)
        anyName = Not String.IsNullOrEmpty(txtName5.Text)
        anyName = Not String.IsNullOrEmpty(txtName6.Text)
        anyName = Not String.IsNullOrEmpty(txtName7.Text)
        anyName = Not String.IsNullOrEmpty(txtName8.Text)
        anyName = Not String.IsNullOrEmpty(txtName9.Text)
        anyName = False
        Return anyName
    End Function

    Function generateClientFolderName() As String
        Dim name(10) As String
        name(0) = txtName1.Text.Trim
        name(1) = txtName2.Text.Trim
        name(2) = txtName3.Text.Trim
        name(3) = txtName4.Text.Trim
        name(4) = txtName5.Text.Trim
        name(5) = txtName6.Text.Trim
        name(6) = txtName7.Text.Trim
        name(7) = txtName8.Text.Trim
        name(8) = txtName9.Text.Trim
        Return String.Join(" - ", name.ToList.Where(Function(x) Not String.IsNullOrEmpty(x)))
    End Function

    Function createFolder(name As String) As DirectoryInfo
        Return Directory.CreateDirectory(Path.Combine(ClientFolderPath, name))
    End Function
    Sub copyFiles(dinf As DirectoryInfo, clientFolderName As String)
        Dim clientFolderFullPathName As String
        clientFolderFullPathName = createFolder(clientFolderName).FullName
        For Each item As String In clbFile.CheckedItems
            File.Copy(Path.Combine(TemplateFilesPath, item), Path.Combine(clientFolderFullPathName, item))
        Next

    End Sub

    Sub loadFilesInCheckListCombo()
        Dim files() As String
        files = Directory.GetFiles(TemplateFilesPath)
        files.ToList.ForEach(Sub(x) clbFile.Items.Add(Path.GetFileName(x)))
        For i = 0 To clbFile.Items.Count - 1
            clbFile.SetItemCheckState(i, CheckState.Checked)
        Next
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        loadFilesInCheckListCombo()
    End Sub
End Class
