Imports Microsoft.VisualBasic.FileIO
Imports System.Xml

Public Class Form1
#Region "Variables"
    Private CSVItems As New CSVList
    Private budget As New Budget
    Private settingsFile As New XmlDocument
    Private fileLoaded As Boolean
    Private settingMngr As New SettingsManager
#End Region

#Region "Contructor/ Form Events"
    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        btnExport.Enabled = False
        lblBudgetTitle.Hide()
        lblPayeeTitle.Hide()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        settingsFile = settingMngr.LoadNewSettingsFile(My.Settings.BudgetSettingsFile)
        If settingsFile IsNot Nothing Then
            budget = settingMngr.LoadSettings(settingsFile)
            txtTotalBudget.Text = budget.TotalBudget.ToString(CurrencyFormat)
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not settingMngr.SaveSettingsFile(budget) Then
            MessageBox.Show("An error occurred while saving budget file" & vbCrLf & "Location: " & My.Settings.BudgetSettingsFile, "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
#End Region

#Region "Methods"

    Private Sub showPayees()
        pnlPayees.Controls.Clear()
        Dim payeeName As String
        For i As Integer = 0 To CSVItems.UniqueItems.Count - 1
            payeeName = CSVItems.UniqueItems(i)

            Dim newX As New Label
            newX.Location = New Point(13, 4 + (i * 25))
            newX.AutoSize = True
            newX.Name = "X" & payeeName
            newX.Text = "X"
            newX.TextAlign = ContentAlignment.MiddleCenter
            newX.ForeColor = DefaultBackColor()
            newX.Cursor = Cursors.Hand
            Dim XTT As New ToolTip
            XTT.SetToolTip(newX, "Remove")
            AddHandler newX.MouseEnter, AddressOf x_hover
            AddHandler newX.MouseLeave, AddressOf x_hover
            AddHandler newX.MouseClick, AddressOf x_click

            Dim newLabel As New Label
            newLabel.Location = New Point(176, 2 + (i * 25))
            newLabel.AutoSize = True
            newLabel.Name = "lbl" & payeeName
            If payeeName = "" Then
                newLabel.Text = "<No Listed Name>"
            Else
                newLabel.Text = payeeName
            End If

            Dim newCmbBox As New ComboBox
            newCmbBox.Location = New Point(30, i * 25)
            newCmbBox.Size = New Size(140, 22)
            newCmbBox.AutoCompleteMode = AutoCompleteMode.Suggest
            newCmbBox.DropDownStyle = ComboBoxStyle.DropDown
            newCmbBox.Items.AddRange(budget.AllCategoryNames().ToArray())
            newCmbBox.Name = "cmb" & payeeName
            newCmbBox.SelectedText = budget.PayeeCategory(payeeName)
            AddHandler newCmbBox.Leave, AddressOf type_changed

            ' Add new conrols to panel
            pnlPayees.Controls.Add(newX)
            pnlPayees.Controls.Add(newLabel)
            pnlPayees.Controls.Add(newCmbBox)
        Next
    End Sub

    Private Sub showBudget()
        pnlBudget.Controls.Clear()
        For i As Integer = 0 To budget.Categories.Count - 1
            Dim category As BudgetCategory = budget.Categories(i)
            Dim newX As New Label
            newX.Location = New Point(0, 4 + (i * 25))
            newX.AutoSize = True
            newX.Name = "budX" & category.Name
            newX.Text = "X"
            newX.TextAlign = ContentAlignment.MiddleCenter
            newX.ForeColor = DefaultBackColor()
            newX.Cursor = Cursors.Hand
            Dim XTT As New ToolTip
            XTT.SetToolTip(newX, "Remove")
            AddHandler newX.MouseEnter, AddressOf x_hover
            AddHandler newX.MouseLeave, AddressOf x_hover
            AddHandler newX.MouseClick, AddressOf budX_click

            Dim newIE As New Label
            newIE.Location = New Point(17, 2 + (i * 25))
            newIE.AutoSize = True
            newIE.Name = "ie" & category.Name
            newIE.Font = New Font("Times New Roman", 10)
            newIE.TextAlign = ContentAlignment.MiddleCenter
            newIE.Text = category.Type.GetEnumDescription()
            newIE.Cursor = Cursors.Hand
            Dim ieTT As New ToolTip
            ieTT.SetToolTip(newIE, "Toggle Expense/Income")
            setIEColor(newIE)
            AddHandler newIE.MouseClick, AddressOf ie_click

            Dim newLabel As New Label
            newLabel.Location = New Point(115 + 17, 2 + (i * 25))
            newLabel.AutoSize = True
            newLabel.Name = "lbl" & category.Name
            newLabel.Text = category.Name

            Dim newTxtBox As New TextBox
            newTxtBox.Location = New Point(15 + 17, i * 25)
            newTxtBox.Size = New Size(90, 22)
            newTxtBox.Name = "txt" & category.Name
            newTxtBox.Text = category.Budget.ToString(CurrencyFormat)
            AddHandler newTxtBox.Leave, AddressOf amount_changed

            pnlBudget.Controls.Add(newX)
            pnlBudget.Controls.Add(newIE)
            pnlBudget.Controls.Add(newLabel)
            pnlBudget.Controls.Add(newTxtBox)
        Next
    End Sub

    Private Sub setIEColor(ByRef label As Label)
        If label.Text = BudgetTypes.Expense.GetEnumDescription() Then
            label.ForeColor = Color.Red
        Else
            label.ForeColor = Color.Green
        End If
    End Sub

    Private Sub disableControls(ByVal payeeName As String)
        For Each cntrl As Control In pnlPayees.Controls
            If cntrl.Name.Contains(payeeName) Then
                cntrl.Enabled = False
            End If
        Next
    End Sub
#End Region
    
#Region "Event Handlers"
    Private Sub btnBrowse_Click() Handles btnBrowse.Click
        If ofdSourceFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
            txtSourcePath.Text = ofdSourceFile.FileName
            btnExport.Enabled = True
            CSVItems = New CSVList(New TextFieldParser(ofdSourceFile.FileName))
            showPayees()
            showBudget()
            lblBudgetTitle.Show()
            lblPayeeTitle.Show()
        End If
    End Sub

    Private Sub btnChangeSettings_Click(sender As Object, e As EventArgs) Handles btnChangeSettings.Click
        If fileLoaded Then
            If MessageBox.Show("A settings file as alreaady been loaded." & vbCrLf & "Are you sure you want to choose a different one?", "", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.No Then
                Exit Sub
            End If
        End If

        If ofdSettingFile.ShowDialog() = Windows.Forms.DialogResult.OK Then
            settingsFile = settingMngr.LoadNewSettingsFile(ofdSettingFile.FileName)
            If settingsFile IsNot Nothing Then
                My.Settings.BudgetSettingsFile = ofdSettingFile.FileName
                My.Settings.Save()
                budget = settingMngr.LoadSettings(settingsFile)
                txtTotalBudget.Text = budget.TotalBudget.ToString(CurrencyFormat)
            End If
        End If

        If CSVItems.Length > 0 Then
            showPayees()
            showBudget()
        End If
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Me.Close()
    End Sub

    Private Sub x_hover(sender As Object, e As EventArgs)
        Dim x = TryCast(sender, Label)
        If x IsNot Nothing Then
            If x.forecolor = Color.Red Then
                x.ForeColor = DefaultBackColor()
            Else
                x.ForeColor = Color.Red
            End If
        End If
    End Sub

    Private Sub x_click(sender As Object, e As EventArgs)
        Dim xLabel = TryCast(sender, Label)
        If xLabel IsNot Nothing Then
            Dim payeeToRemove As String = xLabel.Name.Substring(1)
            While CSVItems.IndexOf(xLabel.Name.Substring(1)) <> -1
                CSVItems.RemoveAt(CSVItems.IndexOf(xLabel.Name.Substring(1)))
            End While
            disableControls(payeeToRemove)
        End If
    End Sub

    Private Sub type_changed(sender As Object, e As EventArgs)
        Dim cmbBox = TryCast(sender, ComboBox)
        If cmbBox IsNot Nothing Then

            If cmbBox.Text = "" Then
                cmbBox.Text = Budget.UnassignedPayeeType
            End If

            If budget.ReassignPayee(cmbBox.Name.Substring(3), cmbBox.Text) Then
                ' A new category was added. Add it to all combo boxes and refresh budget panel
                For Each cmb As ComboBox In pnlPayees.Controls.OfType(Of ComboBox)()
                    cmb.Items.Clear()
                    cmb.Items.AddRange(budget.AllCategoryNames().ToArray())
                Next
                showBudget()
            End If
        End If
    End Sub

    Private Sub budX_click(sender As Object, e As EventArgs)
        Dim xLabel = TryCast(sender, Label)
        If xLabel IsNot Nothing Then
            budget.RemoveCategory(xLabel.Name.Substring(4))
            showPayees()
            showBudget()
        End If
    End Sub

    Private Sub ie_click(sender As Object, e As EventArgs)
        Dim ieLabel = TryCast(sender, Label)
        If ieLabel IsNot Nothing Then
            Dim categoryName As String = ieLabel.Name.Substring(2)
            If ieLabel.Text = BudgetTypes.Expense.GetEnumDescription() Then
                budget.GetCategoryByName(categoryName).Type = BudgetTypes.Income
                ieLabel.Text = BudgetTypes.Income.GetEnumDescription()
            Else
                budget.GetCategoryByName(categoryName).Type = BudgetTypes.Expense
                ieLabel.Text = BudgetTypes.Expense.GetEnumDescription()
            End If
            setIEColor(ieLabel)
        End If
    End Sub

    Private Sub amount_changed(sender As Object, e As EventArgs)
        Dim txtBox = TryCast(sender, TextBox)
        If txtBox IsNot Nothing Then
            Dim categoryName As String = txtBox.Name.Substring(3)
            txtBox.BackColor = color.White
            Dim amount As Decimal
            If Not Decimal.TryParse(txtBox.Text.Replace("$", ""), amount) Then
                txtBox.BackColor = Color.Red
                Exit Sub
            End If

            budget.GetCategoryByName(categoryName).Budget = Math.Round(amount, 2)

            If txtBox.Text.Substring(0, 1) <> "$" Then
                txtBox.Text = "$" & txtBox.Text
            End If

        End If
    End Sub

    Private Sub txtTotalBudget_Leave(sender As Object, e As EventArgs) Handles txtTotalBudget.Leave
        Dim amount As Decimal
        If Not Decimal.TryParse(txtTotalBudget.Text.Replace("$", ""), amount) Then
            txtTotalBudget.BackColor = Color.Red
            Exit Sub
        End If
        budget.TotalBudget = amount

        If txtTotalBudget.Text.Substring(0, 1) <> "$" Then
            txtTotalBudget.Text = "$" & txtTotalBudget.Text
        End If
    End Sub

    Private Sub btnExport_Click() Handles btnExport.Click
        If budget.TotalBudget = 0 Then
            MessageBox.Show("Cannot export until an amount > 0 is entered for Total Budget", "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Exit Sub
        End If
        budget.AssignCategoryUsedStatus(CSVItems)
        Dim newExportWin As New ExportDestination(CSVItems, budget)
        newExportWin.Show()
    End Sub
#End Region

End Class
