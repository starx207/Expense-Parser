Imports Microsoft.VisualBasic.FileIO
Imports System.Xml

Public Class Form1
#Region "Variables"
    Private CSVItems As New CSVList
    Private budget As New ListOfBudgetCategories
    Private settingsFile As New XmlDocument
    Private fileLoaded As Boolean
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
        If LoadNewSettingsFile(My.Settings.BudgetSettingsFile) Then
            LoadSettings()
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        If Not SaveSettingsFile() Then
            MessageBox.Show("An error occurred while saving budget file" & vbCrLf & "Location: " & My.Settings.BudgetSettingsFile, "", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End If
    End Sub
#End Region

#Region "Methods"
    Private Function SaveSettingsFile() As Boolean
        Try
            Dim settings As New XmlWriterSettings
            settings.Indent = True

            Using writer As XmlWriter = XmlWriter.Create(My.Settings.BudgetSettingsFile, settings)
                writer.WriteStartDocument()
                writer.WriteStartElement("MonthlyBudget")
                writer.WriteElementString("TotalBudget", txtTotalBudget.Text.Replace("$", ""))

                For Each category As BudgetCategory In budget.GenericList
                    writer.WriteStartElement("Category")
                    writer.WriteElementString("Name", category.Name)
                    writer.WriteElementString("Type", category.Type)
                    writer.WriteElementString("Budget", category.Budget.ToString())
                    writer.WriteStartElement("Payees")
                    For Each payeeName As String In category.Payees
                        writer.WriteElementString("Name", payeeName)
                    Next
                    writer.WriteEndElement()
                    writer.WriteEndElement()
                Next

                writer.WriteEndElement()
                writer.WriteEndDocument()
            End Using

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function LoadNewSettingsFile(ByVal fileName As String) As Boolean
        Try
            settingsFile.Load(fileName)
            fileLoaded = True
        Catch ex As Exception
            MessageBox.Show("Could not find Settings File. Please specify a new setting file", "", MessageBoxButtons.OK, MessageBoxIcon.Error)
            fileLoaded = False
        End Try

        btnBrowse.Enabled = fileLoaded
        Return fileLoaded
    End Function

    Private Sub LoadSettings()
        budget.GenericList.Clear()
        Dim newCategory As BudgetCategory
        If settingsFile.SelectSingleNode("MonthlyBudget/TotalBudget") IsNot Nothing Then
            txtTotalBudget.Text = "$" & settingsFile.SelectSingleNode("MonthlyBudget/TotalBudget").InnerText
        Else
            txtTotalBudget.Text = "$0.00"
        End If
        For Each budgetItem As XmlNode In settingsFile.SelectNodes("MonthlyBudget/Category")
            newCategory = New BudgetCategory
            If budgetItem("Name") IsNot Nothing Then
                newCategory.Name = budgetItem("Name").InnerText
            End If
            If budgetItem("Type") IsNot Nothing Then
                newCategory.Type = budgetItem("Type").InnerText
            End If
            If budgetItem("Budget") IsNot Nothing Then
                newCategory.Budget = CDbl(budgetItem("Budget").InnerText)
            End If
            If budgetItem("Payees") IsNot Nothing Then
                For Each payee As XmlNode In budgetItem.SelectNodes("Payees/Name")
                    newCategory.Payees.Add(payee.InnerText)
                Next
                newCategory.Payees.Sort()
            End If

            budget.GenericList.Add(newCategory)
        Next
    End Sub

    Private Sub ShowContent()
        ' Display Payee info
        showPayees()

        ' Display budget info
        showBudget()
    End Sub

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
        For i As Integer = 0 To budget.GenericList.Count - 1
            Dim category As BudgetCategory = budget.GenericList(i)
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
            newIE.Text = category.Type
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
            newTxtBox.Text = "$" & category.Budget.ToString()
            AddHandler newTxtBox.Leave, AddressOf amount_changed

            pnlBudget.Controls.Add(newX)
            pnlBudget.Controls.Add(newIE)
            pnlBudget.Controls.Add(newLabel)
            pnlBudget.Controls.Add(newTxtBox)
        Next
    End Sub

    Private Sub PopulateCSVItems(ByVal csv As TextFieldParser)
        CSVItems.Clear()
        Dim currentRow As String()
        Dim DateIndex As Integer
        Dim PayeeIndex As Integer
        Dim AmountIndex As Integer

        csv.TextFieldType = FieldType.Delimited
        csv.SetDelimiters(",")

        ' Set indexes for relavant columns
        currentRow = csv.ReadFields() ' Read header row
        If currentRow IsNot Nothing Then
            For i As Integer = 0 To currentRow.Length - 1
                Select Case Trim(currentRow.GetValue(i).ToString().ToUpper())
                    Case "DATE", "TRANSACTION DATE"
                        DateIndex = i
                    Case "PAYEE NAME", "MERCHANT"
                        PayeeIndex = i
                    Case "AMOUNT", "BILLING AMOUNT"
                        AmountIndex = i
                    Case Else
                        ' Do Nothing
                End Select
            Next
        End If

        While Not csv.EndOfData
            Try
                currentRow = csv.ReadFields()
                If currentRow IsNot Nothing Then
                    CSVItems.Add(currentRow.GetValue(DateIndex),
                                 currentRow.GetValue(PayeeIndex),
                                 currentRow.GetValue(AmountIndex))
                End If
            Catch ex As MalformedLineException
                MessageBox.Show("Line " & ex.Message & " is not valid and will be skipped")
            End Try
        End While

        csv.Close()

        ShowContent()
    End Sub

    Private Sub setIEColor(ByRef label As Label)
        If label.Text = "E" Then
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
            PopulateCSVItems(New TextFieldParser(ofdSourceFile.FileName))
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
            If LoadNewSettingsFile(ofdSettingFile.FileName) Then
                My.Settings.BudgetSettingsFile = ofdSettingFile.FileName
                My.Settings.Save()
                LoadSettings()
            End If
        End If

        If CSVItems.Length > 0 Then
            ShowContent()
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
                cmbBox.Text = "Misc"
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
            If ieLabel.Text = "E" Then
                budget.GetCategoryByName(categoryName).Type = "I"
                ieLabel.Text = "I"
            Else
                budget.GetCategoryByName(categoryName).Type = "E"
                ieLabel.Text = "E"
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

        If txtTotalBudget.Text.Substring(0, 1) <> "$" Then
            txtTotalBudget.Text = "$" & txtTotalBudget.Text
        End If
    End Sub

    Private Sub btnExport_Click() Handles btnExport.Click
        Dim amount As Decimal
        If Not Decimal.TryParse(txtTotalBudget.Text.Replace("$", ""), amount) Then
            MessageBox.Show("Cannot export until a valid amount is entered for Total Budget", "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Exit Sub
        End If
        If amount = 0 Then
            MessageBox.Show("Cannot export until an amount > 0 is entered for Total Budget", "", MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Exit Sub
        End If
        budget.AssignCategoryUsedStatus(CSVItems)
        Dim newExportWin As New ExportDestination
        newExportWin.prepareExport(CSVItems, budget, amount)
    End Sub
#End Region
    
End Class
