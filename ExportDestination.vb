Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExportDestination
    Private IncomeTranactions As New CSVList
    Private ExpenseTransactions As New CSVList
    Private BudgetedCategories As New Budget
    Private UnbudgetedCategories As New Budget
    Private TotalBudget As Decimal
    Private xlApp As Excel.Application
    Private xlWorkBook As Excel.Workbook
    Private exporter As New ExcelExporter

    Sub New(ByVal transactions As CSVList, ByVal budget As Budget)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        exporter.PrepareExport(transactions, budget)

        If My.Settings.ExportFile = NoExportFile Then
            rdoExisting.Checked = False
            rdoNew.Checked = True
        Else
            rdoExisting.Checked = True
            rdoNew.Checked = False
            txtFile.Text = My.Settings.ExportFile
        End If

        switchRadio(rdoExisting.Checked)
    End Sub

    Private Sub enableSave()
        If rdoExisting.Checked And txtFile.Text <> "" Then
            btnSave.Enabled = True
        ElseIf rdoNew.Checked And txtName.Text <> "" And txtPath.Text <> "" Then
            btnSave.Enabled = True
        Else
            btnSave.Enabled = False
        End If
    End Sub

    Private Sub switchRadio(ByVal useExisting As Boolean)
        txtName.Enabled = Not useExisting
        txtPath.Enabled = Not useExisting
        btnPathBrowse.Enabled = Not useExisting
        lblPath.Enabled = Not useExisting
        lblName.Enabled = Not useExisting
        txtFile.Enabled = useExisting
        btnFileBrowse.Enabled = useExisting
        lblFile.Enabled = useExisting

        enableSave()
    End Sub

    Private Sub btnCancel_Click() Handles btnCancel.Click
        Close()
    End Sub

    Private Sub radio_CheckedChanged(sender As Object, e As EventArgs) Handles rdoExisting.CheckedChanged, rdoNew.CheckedChanged
        switchRadio(rdoExisting.Checked)
    End Sub

    Private Sub btnFileBrowse_Click() Handles btnFileBrowse.Click
        If OpenFileDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtFile.Text = OpenFileDialog.FileName
            enableSave()
        End If
    End Sub

    Private Sub btnPathBrowse_Click() Handles btnPathBrowse.Click
        If FolderBrowserDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtPath.Text = FolderBrowserDialog.SelectedPath
            txtName.Focus()
            enableSave()
        End If
    End Sub

    Private Sub txtName_TextChanged() Handles txtName.TextChanged
        enableSave()
    End Sub

    Private Sub txtName_Enter() Handles txtName.Enter
        txtName.SelectAll()
    End Sub

    Private Sub btnSave_Click() Handles btnSave.Click
        Dim fileName As String
        Try
            If rdoExisting.Checked Then
                fileName = txtFile.Text
            Else
                fileName = txtPath.Text & "\" & txtName.Text
            End If
            exporter.ExportMonth(fileName, rdoNew.Checked)
            My.Settings.ExportFile = fileName
            My.Settings.Save()
            MessageBox.Show("New Worksheet Created!", "Worksheet Saved")
            Close()
        Catch ex As Exception
            Dim err As String = ""
            While ex.InnerException IsNot Nothing
                err += ex.Message + vbCrLf
                ex = ex.InnerException
            End While
            err += ex.Message
            MessageBox.Show(err, "Export Failed")
        End Try
    End Sub
End Class