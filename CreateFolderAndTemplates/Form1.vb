Imports System.IO
Imports System.Collections.Generic
Imports Microsoft.Office.Tools.Word
Public Class Form1

    Private Const ClientFolderPath As String = "C:\A1\work\autoDoc\clients"
    Private Const TemplateFilesPath As String = "C:\A1\work\autoDoc\template"


    Private Sub btnCopyTemplates_Click(sender As Object, e As EventArgs) Handles btnCopyTemplates.Click
        If NoName() Then
            MessageBox.Show("Please provide at least one name")
            Return
        End If
        copyFiles(Nothing)
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
    Sub copyFiles(dinf As DirectoryInfo)
        Dim clientFolderName As String
        clientFolderName = createFolder(generateClientFolderName).FullName
        For Each item As String In clbFile.CheckedItems
            File.Copy(Path.Combine(TemplateFilesPath, item), Path.Combine(clientFolderName, item))
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
