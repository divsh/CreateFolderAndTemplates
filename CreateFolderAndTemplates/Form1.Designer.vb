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
        Me.txtName1 = New System.Windows.Forms.TextBox()
        Me.txtName2 = New System.Windows.Forms.TextBox()
        Me.txtName3 = New System.Windows.Forms.TextBox()
        Me.txtName4 = New System.Windows.Forms.TextBox()
        Me.txtName5 = New System.Windows.Forms.TextBox()
        Me.txtName6 = New System.Windows.Forms.TextBox()
        Me.txtName7 = New System.Windows.Forms.TextBox()
        Me.txtName8 = New System.Windows.Forms.TextBox()
        Me.txtName9 = New System.Windows.Forms.TextBox()
        Me.clbFile = New System.Windows.Forms.CheckedListBox()
        Me.btnCopyTemplates = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtName1
        '
        Me.txtName1.Location = New System.Drawing.Point(12, 9)
        Me.txtName1.Name = "txtName1"
        Me.txtName1.Size = New System.Drawing.Size(338, 20)
        Me.txtName1.TabIndex = 0
        '
        'txtName2
        '
        Me.txtName2.Location = New System.Drawing.Point(12, 35)
        Me.txtName2.Name = "txtName2"
        Me.txtName2.Size = New System.Drawing.Size(338, 20)
        Me.txtName2.TabIndex = 1
        '
        'txtName3
        '
        Me.txtName3.Location = New System.Drawing.Point(12, 61)
        Me.txtName3.Name = "txtName3"
        Me.txtName3.Size = New System.Drawing.Size(338, 20)
        Me.txtName3.TabIndex = 2
        '
        'txtName4
        '
        Me.txtName4.Location = New System.Drawing.Point(12, 87)
        Me.txtName4.Name = "txtName4"
        Me.txtName4.Size = New System.Drawing.Size(338, 20)
        Me.txtName4.TabIndex = 3
        '
        'txtName5
        '
        Me.txtName5.Location = New System.Drawing.Point(12, 113)
        Me.txtName5.Name = "txtName5"
        Me.txtName5.Size = New System.Drawing.Size(338, 20)
        Me.txtName5.TabIndex = 4
        '
        'txtName6
        '
        Me.txtName6.Location = New System.Drawing.Point(12, 139)
        Me.txtName6.Name = "txtName6"
        Me.txtName6.Size = New System.Drawing.Size(338, 20)
        Me.txtName6.TabIndex = 5
        '
        'txtName7
        '
        Me.txtName7.Location = New System.Drawing.Point(12, 165)
        Me.txtName7.Name = "txtName7"
        Me.txtName7.Size = New System.Drawing.Size(338, 20)
        Me.txtName7.TabIndex = 6
        '
        'txtName8
        '
        Me.txtName8.Location = New System.Drawing.Point(12, 191)
        Me.txtName8.Name = "txtName8"
        Me.txtName8.Size = New System.Drawing.Size(338, 20)
        Me.txtName8.TabIndex = 7
        '
        'txtName9
        '
        Me.txtName9.Location = New System.Drawing.Point(12, 217)
        Me.txtName9.Name = "txtName9"
        Me.txtName9.Size = New System.Drawing.Size(338, 20)
        Me.txtName9.TabIndex = 8
        '
        'clbFile
        '
        Me.clbFile.CheckOnClick = True
        Me.clbFile.FormattingEnabled = True
        Me.clbFile.Location = New System.Drawing.Point(392, 9)
        Me.clbFile.Name = "clbFile"
        Me.clbFile.Size = New System.Drawing.Size(396, 229)
        Me.clbFile.TabIndex = 9
        '
        'btnCopyTemplates
        '
        Me.btnCopyTemplates.Location = New System.Drawing.Point(680, 244)
        Me.btnCopyTemplates.Name = "btnCopyTemplates"
        Me.btnCopyTemplates.Size = New System.Drawing.Size(108, 23)
        Me.btnCopyTemplates.TabIndex = 10
        Me.btnCopyTemplates.Text = "Create Templates"
        Me.btnCopyTemplates.UseVisualStyleBackColor = True
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 450)
        Me.Controls.Add(Me.btnCopyTemplates)
        Me.Controls.Add(Me.clbFile)
        Me.Controls.Add(Me.txtName9)
        Me.Controls.Add(Me.txtName8)
        Me.Controls.Add(Me.txtName7)
        Me.Controls.Add(Me.txtName6)
        Me.Controls.Add(Me.txtName5)
        Me.Controls.Add(Me.txtName4)
        Me.Controls.Add(Me.txtName3)
        Me.Controls.Add(Me.txtName2)
        Me.Controls.Add(Me.txtName1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtName1 As TextBox
    Friend WithEvents txtName2 As TextBox
    Friend WithEvents txtName3 As TextBox
    Friend WithEvents txtName4 As TextBox
    Friend WithEvents txtName5 As TextBox
    Friend WithEvents txtName6 As TextBox
    Friend WithEvents txtName7 As TextBox
    Friend WithEvents txtName8 As TextBox
    Friend WithEvents txtName9 As TextBox
    Friend WithEvents clbFile As CheckedListBox
    Friend WithEvents btnCopyTemplates As Button
End Class
