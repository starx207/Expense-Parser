<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class ExportDestination
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
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.FolderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
        Me.OpenFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.txtFile = New System.Windows.Forms.TextBox()
        Me.txtPath = New System.Windows.Forms.TextBox()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.btnFileBrowse = New System.Windows.Forms.Button()
        Me.btnPathBrowse = New System.Windows.Forms.Button()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.lblPath = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.rdoExisting = New System.Windows.Forms.RadioButton()
        Me.rdoNew = New System.Windows.Forms.RadioButton()
        Me.SuspendLayout
        '
        'btnSave
        '
        Me.btnSave.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btnSave.Location = New System.Drawing.Point(206, 110)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(68, 23)
        Me.btnSave.TabIndex = 8
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(278, 110)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(68, 23)
        Me.btnCancel.TabIndex = 9
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'OpenFileDialog
        '
        Me.OpenFileDialog.FileName = "OpenFileDialog1"
        Me.OpenFileDialog.Filter = "Excel Workbook (*.xlsx, *.xls) | *.xlsx; *.xls"
        '
        'txtFile
        '
        Me.txtFile.Location = New System.Drawing.Point(193, 15)
        Me.txtFile.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtFile.Name = "txtFile"
        Me.txtFile.ReadOnly = True
        Me.txtFile.Size = New System.Drawing.Size(279, 20)
        Me.txtFile.TabIndex = 10
        '
        'txtPath
        '
        Me.txtPath.Location = New System.Drawing.Point(193, 55)
        Me.txtPath.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.ReadOnly = True
        Me.txtPath.Size = New System.Drawing.Size(279, 20)
        Me.txtPath.TabIndex = 11
        '
        'txtName
        '
        Me.txtName.Location = New System.Drawing.Point(193, 78)
        Me.txtName.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(279, 20)
        Me.txtName.TabIndex = 12
        '
        'btnFileBrowse
        '
        Me.btnFileBrowse.Location = New System.Drawing.Point(476, 15)
        Me.btnFileBrowse.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnFileBrowse.Name = "btnFileBrowse"
        Me.btnFileBrowse.Size = New System.Drawing.Size(56, 19)
        Me.btnFileBrowse.TabIndex = 13
        Me.btnFileBrowse.Text = "Browse"
        Me.btnFileBrowse.UseVisualStyleBackColor = True
        '
        'btnPathBrowse
        '
        Me.btnPathBrowse.Location = New System.Drawing.Point(476, 55)
        Me.btnPathBrowse.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.btnPathBrowse.Name = "btnPathBrowse"
        Me.btnPathBrowse.Size = New System.Drawing.Size(56, 19)
        Me.btnPathBrowse.TabIndex = 14
        Me.btnPathBrowse.Text = "Browse"
        Me.btnPathBrowse.UseVisualStyleBackColor = True
        '
        'lblFile
        '
        Me.lblFile.AutoSize = True
        Me.lblFile.Location = New System.Drawing.Point(124, 18)
        Me.lblFile.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(65, 13)
        Me.lblFile.TabIndex = 16
        Me.lblFile.Text = "Choose File:"
        '
        'lblPath
        '
        Me.lblPath.AutoSize = True
        Me.lblPath.Location = New System.Drawing.Point(118, 58)
        Me.lblPath.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblPath.Name = "lblPath"
        Me.lblPath.Size = New System.Drawing.Size(71, 13)
        Me.lblPath.TabIndex = 17
        Me.lblPath.Text = "Choose Path:"
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(132, 80)
        Me.lblName.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(57, 13)
        Me.lblName.TabIndex = 18
        Me.lblName.Text = "File Name:"
        '
        'rdoExisting
        '
        Me.rdoExisting.AutoSize = True
        Me.rdoExisting.Location = New System.Drawing.Point(14, 15)
        Me.rdoExisting.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rdoExisting.Name = "rdoExisting"
        Me.rdoExisting.Size = New System.Drawing.Size(85, 17)
        Me.rdoExisting.TabIndex = 19
        Me.rdoExisting.TabStop = True
        Me.rdoExisting.Text = "Use Exisiting"
        Me.rdoExisting.UseVisualStyleBackColor = True
        '
        'rdoNew
        '
        Me.rdoNew.AutoSize = True
        Me.rdoNew.Location = New System.Drawing.Point(14, 54)
        Me.rdoNew.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.rdoNew.Name = "rdoNew"
        Me.rdoNew.Size = New System.Drawing.Size(81, 17)
        Me.rdoNew.TabIndex = 20
        Me.rdoNew.TabStop = True
        Me.rdoNew.Text = "Create New"
        Me.rdoNew.UseVisualStyleBackColor = True
        '
        'ExportDestination
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(552, 144)
        Me.Controls.Add(Me.rdoNew)
        Me.Controls.Add(Me.rdoExisting)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.lblPath)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.btnPathBrowse)
        Me.Controls.Add(Me.btnFileBrowse)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.txtPath)
        Me.Controls.Add(Me.txtFile)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.Margin = New System.Windows.Forms.Padding(2, 2, 2, 2)
        Me.Name = "ExportDestination"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Choose where to export"
        Me.ResumeLayout(false)
        Me.PerformLayout

End Sub
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents FolderBrowserDialog As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents OpenFileDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents txtFile As System.Windows.Forms.TextBox
    Friend WithEvents txtPath As System.Windows.Forms.TextBox
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents btnFileBrowse As System.Windows.Forms.Button
    Friend WithEvents btnPathBrowse As System.Windows.Forms.Button
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents lblPath As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents rdoExisting As System.Windows.Forms.RadioButton
    Friend WithEvents rdoNew As System.Windows.Forms.RadioButton
End Class
