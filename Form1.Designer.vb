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
        Me.ofdSourceFile = New System.Windows.Forms.OpenFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtSourcePath = New System.Windows.Forms.TextBox()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.pnlPayees = New System.Windows.Forms.Panel()
        Me.pnlBudget = New System.Windows.Forms.Panel()
        Me.lblBudgetTitle = New System.Windows.Forms.Label()
        Me.lblPayeeTitle = New System.Windows.Forms.Label()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.btnChangeSettings = New System.Windows.Forms.Button()
        Me.ofdSettingFile = New System.Windows.Forms.OpenFileDialog()
        Me.txtTotalBudget = New System.Windows.Forms.TextBox()
        Me.lblTotalBudget = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ofdSourceFile
        '
        Me.ofdSourceFile.FileName = "SourceFile"
        Me.ofdSourceFile.Filter = "CSV|*.csv"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(9, 7)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(146, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Choose transaction file (.csv):"
        '
        'txtSourcePath
        '
        Me.txtSourcePath.Location = New System.Drawing.Point(160, 5)
        Me.txtSourcePath.Margin = New System.Windows.Forms.Padding(2)
        Me.txtSourcePath.Name = "txtSourcePath"
        Me.txtSourcePath.ReadOnly = True
        Me.txtSourcePath.Size = New System.Drawing.Size(390, 20)
        Me.txtSourcePath.TabIndex = 1
        '
        'btnBrowse
        '
        Me.btnBrowse.Location = New System.Drawing.Point(554, 5)
        Me.btnBrowse.Margin = New System.Windows.Forms.Padding(2)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(56, 19)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'pnlPayees
        '
        Me.pnlPayees.AutoScroll = True
        Me.pnlPayees.Location = New System.Drawing.Point(351, 54)
        Me.pnlPayees.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlPayees.Name = "pnlPayees"
        Me.pnlPayees.Size = New System.Drawing.Size(373, 334)
        Me.pnlPayees.TabIndex = 3
        '
        'pnlBudget
        '
        Me.pnlBudget.AutoScroll = True
        Me.pnlBudget.Location = New System.Drawing.Point(25, 54)
        Me.pnlBudget.Margin = New System.Windows.Forms.Padding(2)
        Me.pnlBudget.Name = "pnlBudget"
        Me.pnlBudget.Size = New System.Drawing.Size(296, 333)
        Me.pnlBudget.TabIndex = 5
        '
        'lblBudgetTitle
        '
        Me.lblBudgetTitle.AutoSize = True
        Me.lblBudgetTitle.Font = New System.Drawing.Font("Baskerville Old Face", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblBudgetTitle.Location = New System.Drawing.Point(81, 29)
        Me.lblBudgetTitle.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblBudgetTitle.Name = "lblBudgetTitle"
        Me.lblBudgetTitle.Size = New System.Drawing.Size(71, 22)
        Me.lblBudgetTitle.TabIndex = 6
        Me.lblBudgetTitle.Text = "Budget"
        '
        'lblPayeeTitle
        '
        Me.lblPayeeTitle.AutoSize = True
        Me.lblPayeeTitle.Font = New System.Drawing.Font("Baskerville Old Face", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPayeeTitle.Location = New System.Drawing.Point(444, 29)
        Me.lblPayeeTitle.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblPayeeTitle.Name = "lblPayeeTitle"
        Me.lblPayeeTitle.Size = New System.Drawing.Size(124, 22)
        Me.lblPayeeTitle.TabIndex = 7
        Me.lblPayeeTitle.Text = "Payee Names"
        '
        'btnExit
        '
        Me.btnExit.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExit.Location = New System.Drawing.Point(642, 392)
        Me.btnExit.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(82, 24)
        Me.btnExit.TabIndex = 8
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'btnExport
        '
        Me.btnExport.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExport.Location = New System.Drawing.Point(523, 392)
        Me.btnExport.Margin = New System.Windows.Forms.Padding(2)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(115, 24)
        Me.btnExport.TabIndex = 9
        Me.btnExport.Text = "Export to Excel"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'btnChangeSettings
        '
        Me.btnChangeSettings.Location = New System.Drawing.Point(649, 5)
        Me.btnChangeSettings.Name = "btnChangeSettings"
        Me.btnChangeSettings.Size = New System.Drawing.Size(75, 20)
        Me.btnChangeSettings.TabIndex = 10
        Me.btnChangeSettings.Text = "Settings File"
        Me.btnChangeSettings.UseVisualStyleBackColor = True
        '
        'ofdSettingFile
        '
        Me.ofdSettingFile.FileName = "SettingFile"
        Me.ofdSettingFile.Filter = "XML Files|*.xml"
        '
        'txtTotalBudget
        '
        Me.txtTotalBudget.Location = New System.Drawing.Point(99, 392)
        Me.txtTotalBudget.Name = "txtTotalBudget"
        Me.txtTotalBudget.Size = New System.Drawing.Size(144, 20)
        Me.txtTotalBudget.TabIndex = 11
        '
        'lblTotalBudget
        '
        Me.lblTotalBudget.AutoSize = True
        Me.lblTotalBudget.Location = New System.Drawing.Point(22, 395)
        Me.lblTotalBudget.Name = "lblTotalBudget"
        Me.lblTotalBudget.Size = New System.Drawing.Size(71, 13)
        Me.lblTotalBudget.TabIndex = 12
        Me.lblTotalBudget.Text = "Total Budget:"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(733, 426)
        Me.Controls.Add(Me.lblTotalBudget)
        Me.Controls.Add(Me.txtTotalBudget)
        Me.Controls.Add(Me.btnChangeSettings)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.lblPayeeTitle)
        Me.Controls.Add(Me.lblBudgetTitle)
        Me.Controls.Add(Me.pnlBudget)
        Me.Controls.Add(Me.pnlPayees)
        Me.Controls.Add(Me.btnBrowse)
        Me.Controls.Add(Me.txtSourcePath)
        Me.Controls.Add(Me.Label1)
        Me.Margin = New System.Windows.Forms.Padding(2)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents ofdSourceFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtSourcePath As System.Windows.Forms.TextBox
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents pnlPayees As System.Windows.Forms.Panel
    Friend WithEvents pnlBudget As System.Windows.Forms.Panel
    Friend WithEvents lblBudgetTitle As System.Windows.Forms.Label
    Friend WithEvents lblPayeeTitle As System.Windows.Forms.Label
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents btnChangeSettings As System.Windows.Forms.Button
    Friend WithEvents ofdSettingFile As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtTotalBudget As System.Windows.Forms.TextBox
    Friend WithEvents lblTotalBudget As System.Windows.Forms.Label

End Class
