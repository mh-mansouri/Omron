<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_main
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lst_read = New System.Windows.Forms.ListBox
        Me.cmd_PLCTest = New System.Windows.Forms.Button
        Me.cmd_PLCStatus = New System.Windows.Forms.Button
        Me.cmd_PLCread = New System.Windows.Forms.Button
        Me.txt_read = New System.Windows.Forms.TextBox
        Me.cmd_write = New System.Windows.Forms.Button
        Me.cmd_clear = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'lst_read
        '
        Me.lst_read.FormattingEnabled = True
        Me.lst_read.Location = New System.Drawing.Point(12, 137)
        Me.lst_read.Name = "lst_read"
        Me.lst_read.Size = New System.Drawing.Size(315, 238)
        Me.lst_read.TabIndex = 8
        '
        'cmd_PLCTest
        '
        Me.cmd_PLCTest.Location = New System.Drawing.Point(12, 11)
        Me.cmd_PLCTest.Name = "cmd_PLCTest"
        Me.cmd_PLCTest.Size = New System.Drawing.Size(315, 23)
        Me.cmd_PLCTest.TabIndex = 9
        Me.cmd_PLCTest.Text = "Test PLC"
        Me.cmd_PLCTest.UseVisualStyleBackColor = True
        '
        'cmd_PLCStatus
        '
        Me.cmd_PLCStatus.Enabled = False
        Me.cmd_PLCStatus.Location = New System.Drawing.Point(12, 40)
        Me.cmd_PLCStatus.Name = "cmd_PLCStatus"
        Me.cmd_PLCStatus.Size = New System.Drawing.Size(315, 23)
        Me.cmd_PLCStatus.TabIndex = 9
        Me.cmd_PLCStatus.Text = "PLC Error and Status"
        Me.cmd_PLCStatus.UseVisualStyleBackColor = True
        '
        'cmd_PLCread
        '
        Me.cmd_PLCread.Enabled = False
        Me.cmd_PLCread.Location = New System.Drawing.Point(12, 69)
        Me.cmd_PLCread.Name = "cmd_PLCread"
        Me.cmd_PLCread.Size = New System.Drawing.Size(221, 23)
        Me.cmd_PLCread.TabIndex = 9
        Me.cmd_PLCread.Text = "PLC Register Read"
        Me.cmd_PLCread.UseVisualStyleBackColor = True
        '
        'txt_read
        '
        Me.txt_read.Location = New System.Drawing.Point(333, 11)
        Me.txt_read.Multiline = True
        Me.txt_read.Name = "txt_read"
        Me.txt_read.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txt_read.Size = New System.Drawing.Size(315, 363)
        Me.txt_read.TabIndex = 10
        '
        'cmd_write
        '
        Me.cmd_write.Enabled = False
        Me.cmd_write.Location = New System.Drawing.Point(12, 98)
        Me.cmd_write.Name = "cmd_write"
        Me.cmd_write.Size = New System.Drawing.Size(315, 23)
        Me.cmd_write.TabIndex = 11
        Me.cmd_write.Text = "Write (0000) to D.101"
        Me.cmd_write.UseVisualStyleBackColor = True
        '
        'cmd_clear
        '
        Me.cmd_clear.Location = New System.Drawing.Point(239, 69)
        Me.cmd_clear.Name = "cmd_clear"
        Me.cmd_clear.Size = New System.Drawing.Size(88, 23)
        Me.cmd_clear.TabIndex = 12
        Me.cmd_clear.Text = "Clear Text"
        Me.cmd_clear.UseVisualStyleBackColor = True
        '
        'frm_main
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(655, 384)
        Me.Controls.Add(Me.cmd_clear)
        Me.Controls.Add(Me.cmd_write)
        Me.Controls.Add(Me.txt_read)
        Me.Controls.Add(Me.cmd_PLCread)
        Me.Controls.Add(Me.cmd_PLCStatus)
        Me.Controls.Add(Me.cmd_PLCTest)
        Me.Controls.Add(Me.lst_read)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frm_main"
        Me.Text = "Omron Serial Interface"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents lst_read As System.Windows.Forms.ListBox
    Private WithEvents cmd_PLCTest As System.Windows.Forms.Button
    Private WithEvents cmd_PLCStatus As System.Windows.Forms.Button
    Private WithEvents cmd_PLCread As System.Windows.Forms.Button
    Friend WithEvents txt_read As System.Windows.Forms.TextBox
    Friend WithEvents cmd_write As System.Windows.Forms.Button
    Friend WithEvents cmd_clear As System.Windows.Forms.Button

End Class
