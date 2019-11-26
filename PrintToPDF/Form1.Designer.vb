<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.btnPrintPDF = New System.Windows.Forms.Button()
        Me.clbFile = New System.Windows.Forms.CheckedListBox()
        Me.SuspendLayout()
        '
        'btnPrintPDF
        '
        Me.btnPrintPDF.Location = New System.Drawing.Point(425, 283)
        Me.btnPrintPDF.Name = "btnPrintPDF"
        Me.btnPrintPDF.Size = New System.Drawing.Size(75, 23)
        Me.btnPrintPDF.TabIndex = 0
        Me.btnPrintPDF.Text = "Print PDF"
        Me.btnPrintPDF.UseVisualStyleBackColor = True
        '
        'clbFile
        '
        Me.clbFile.CheckOnClick = True
        Me.clbFile.FormattingEnabled = True
        Me.clbFile.Location = New System.Drawing.Point(12, 12)
        Me.clbFile.Name = "clbFile"
        Me.clbFile.Size = New System.Drawing.Size(488, 229)
        Me.clbFile.TabIndex = 10
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(512, 318)
        Me.Controls.Add(Me.clbFile)
        Me.Controls.Add(Me.btnPrintPDF)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents btnPrintPDF As Button
    Friend WithEvents clbFile As CheckedListBox
End Class
