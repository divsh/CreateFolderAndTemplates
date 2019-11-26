Imports System.Drawing.Printing
Imports System.IO
Imports Microsoft.Office.Interop.Word

Public Class Form1


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        loadAllDocs()
    End Sub
    Sub loadAllDocs()
        Directory.EnumerateFiles(Environment.CurrentDirectory).ToList.Where(Function(X) X.ToLower.EndsWith(".doc") OrElse X.ToLower.EndsWith(".docx")).ToList.ForEach(Sub(y) clbFile.Items.Add(Path.GetFileName(y)))
    End Sub

    Private Sub btnPrintPDF_Click(sender As Object, e As EventArgs) Handles btnPrintPDF.Click
        For Each item As Object In clbFile.CheckedItems
            printToFileDllWay(Path.Combine(Environment.CurrentDirectory, item.ToString))
        Next
    End Sub

    Sub printToFileDllWay(filename As String)
        Dim ap As Application = New Application()
        Dim d As Document = ap.Documents.Open(filename)

        d.ExportAsFixedFormat(OutputFileName:="C:\Users\devesh.sharma\Downloads\hello hello.pdf", ExportFormat:=WdExportFormat.wdExportFormatPDF, OpenAfterExport:=True, OptimizeFor:=
        WdExportOptimizeFor.wdExportOptimizeForPrint, Range:=WdExportRange.wdExportAllDocument, From:=1, To:=1,
        Item:=WdExportItem.wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True,
        CreateBookmarks:=WdExportCreateBookmarks.wdExportCreateNoBookmarks, DocStructureTags:=True,
        BitmapMissingFonts:=True, UseISO19005_1:=False)
    End Sub


    Sub printFile(filename As String)

        'initialize PrintDocument object
        Dim doc As PrintDocument = New PrintDocument()
        Dim PrinterSettings As PrinterSettings = New PrinterSettings()

        PrinterSettings.PrintToFile = True

        If File.Exists(Path.GetFileNameWithoutExtension(Path.Combine(Environment.CurrentDirectory, filename)) & ".pdf") Then
            Dim i As Integer = 1
            Do While File.Exists(Path.GetFileNameWithoutExtension(Path.Combine(Environment.CurrentDirectory, filename)) & "_" & i.ToString & ".pdf")
                i = i + 1
            Loop
            PrinterSettings.PrintFileName = Path.GetFileNameWithoutExtension(Path.Combine(Environment.CurrentDirectory, filename)) & "_" & (i).ToString & ".pdf"
        Else
            PrinterSettings.PrintFileName = Path.GetFileNameWithoutExtension(Path.Combine(Environment.CurrentDirectory, filename)) & ".pdf"
        End If
        doc.PrinterSettings = PrinterSettings
        ' set the filename to whatever you Like (full path)

        doc.Print()
    End Sub
End Class
