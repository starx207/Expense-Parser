Imports Excel = Microsoft.Office.Interop.Excel

Public Class ExportDestination
    Private IncomeTranactions As New CSVList
    Private ExpenseTransactions As New CSVList
    Private BudgetedCategories As New ListOfBudgetCategories
    Private UnbudgetedCategories As New ListOfBudgetCategories
    Private TotalBudget As Decimal
    Private xlApp As Excel.Application
    Private xlWorkBook As Excel.Workbook

    Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        If My.Settings.ExportFile = "None" Then
            rdoExisting.Checked = False
            rdoNew.Checked = True
        Else
            rdoExisting.Checked = True
            rdoNew.Checked = False
            txtFile.Text = My.Settings.ExportFile
        End If

        switchRadio()
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

    Private Sub switchRadio()
        If rdoExisting.Checked Then
            txtName.Enabled = False
            txtPath.Enabled = False
            btnPathBrowse.Enabled = False
            lblPath.Enabled = False
            lblName.Enabled = False
            txtFile.Enabled = True
            btnFileBrowse.Enabled = True
            lblFile.Enabled = True
        Else
            txtFile.Enabled = False
            btnFileBrowse.Enabled = False
            lblFile.Enabled = False
            txtName.Enabled = True
            txtPath.Enabled = True
            btnPathBrowse.Enabled = True
            lblPath.Enabled = True
            lblName.Enabled = True
        End If

        enableSave()
    End Sub

    Private Sub btnCancel_Click() Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub radio_CheckedChanged(sender As Object, e As EventArgs) Handles rdoExisting.CheckedChanged, rdoNew.CheckedChanged
        switchRadio()
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

    Private Function GetMonthName(ByVal checkDate As String) As String
        Dim indexOfEOM As Integer = checkDate.IndexOf("/")
        Dim monthNum As Integer = checkDate.Substring(0, indexOfEOM)
        Select Case monthNum
            Case "01", "1"
                Return "January"
            Case "02", "2"
                Return "February"
            Case "03", "3"
                Return "March"
            Case "04", "4"
                Return "April"
            Case "05", "5"
                Return "May"
            Case "06", "6"
                Return "June"
            Case "07", "7"
                Return "July"
            Case "08", "8"
                Return "August"
            Case "09", "9"
                Return "September"
            Case "10"
                Return "October"
            Case "11"
                Return "November"
            Case "12"
                Return "December"
            Case Else
                Return "Unknown"
        End Select
    End Function

    Public Sub exportToExcel()
        Dim sheetName As String = ""
        Dim monthName As String
        For Each transaction As CSVItem In ExpenseTransactions.GenericList
            monthName = GetMonthName(transaction.TransDate)
            If sheetName = "" Then
                sheetName = monthName
            ElseIf sheetName <> monthName Then
                MessageBox.Show("Source file has transactions for more than 1 month (MM/dd/yyyy)", "Cannot Export")
                Exit Sub
            End If
        Next
        For Each transaction As CSVItem In IncomeTranactions.GenericList
            monthName = GetMonthName(transaction.TransDate)
            If sheetName = "" Then
                sheetName = monthName
            ElseIf sheetName <> monthName Then
                MessageBox.Show("Source file has transactions for more than 1 month (MM/dd/yyyy)", "Cannot Export")
                Exit Sub
            End If
        Next

        ' Verify Excel is installed on computer
        xlApp = New Excel.Application
        If xlApp Is Nothing Then
            MessageBox.Show("Excel is not properly installed!", "Error")
            Exit Sub
        End If

        ' Disable Excel overwrite prompt
        xlApp.DisplayAlerts = False

        ' Prepare Workbook
        If rdoNew.Checked Then
            xlWorkBook = xlApp.Workbooks.Add(System.Reflection.Missing.Value)
        ElseIf rdoExisting.Checked Then
            Try
                xlWorkBook = xlApp.Workbooks.Open(txtFile.Text)
            Catch ex As Exception
                MessageBox.Show("Could not find file. Please check that it still exists", "Error")
                My.Settings.ExportFile = "None"
                My.Settings.Save()
                Exit Sub
            End Try
        End If



        ' Prepare Worksheet
        Dim xlWorkSheet = DirectCast(xlWorkBook.Worksheets.Add(), Excel.Worksheet)
        xlWorkSheet.Select()

        ' Prepare a name for worksheet
        ' Default name is Month
        ' If that name already exists, append a number to the end
        Dim sheetFound As Boolean
        Dim numSheetMatch As Integer = 2
        Do
            sheetFound = False
            For Each sheet As Excel.Worksheet In xlWorkBook.Worksheets
                If sheet.Name = sheetName Then
                    sheetFound = True
                    If numSheetMatch > 2 Then
                        sheetName = sheetName.Substring(0, sheetName.Length - numSheetMatch.ToString().Length)
                    End If
                    sheetName = sheetName & numSheetMatch.ToString()
                    numSheetMatch += 1
                    Exit For
                End If
            Next
        Loop While sheetFound

        xlWorkSheet.Name = sheetName

        ' Currency Format String
        Dim currencyFormat As String = "$ #,###,##0.00"
        Dim percentFormat As String = "##0.0%"
        Dim range As Excel.Range

        '-----------------------------------------------------------
        '          Create Expense Table
        '-----------------------------------------------------------
        ' Fill in Expense header cells
        range = xlWorkSheet.Range(xlWorkSheet.Cells(1, 14), xlWorkSheet.Cells(1, 17))
        xlWorkSheet.Cells(1, 14) = "Expenses"
        range.Font.Bold = True
        range.Interior.Color = Color.Red
        range.HorizontalAlignment = Excel.Constants.xlCenter
        range.Merge()

        xlWorkSheet.Range(xlWorkSheet.Cells(2, 14), xlWorkSheet.Cells(2, 17)).Interior.Color = Color.LightGray
        xlWorkSheet.Cells(2, 14) = "Date"
        xlWorkSheet.Cells(2, 15) = "Payee"
        xlWorkSheet.Cells(2, 16) = "Category"
        xlWorkSheet.Cells(2, 17) = "Amount"

        ' Fill in Expense value cells
        For i As Integer = 0 To ExpenseTransactions.Length - 1
            Dim transaction As CSVItem = ExpenseTransactions.Item(i)
            Dim category As String = BudgetedCategories.GetAssignedCategory(transaction.Payee)
            If category = "Unassigned" Then
                category = UnbudgetedCategories.GetAssignedCategory(transaction.Payee)
            End If
            xlWorkSheet.Range(xlWorkSheet.Cells(3 + i, 14), xlWorkSheet.Cells(3 + i, 17)).Interior.Color = Color.LightCoral
            xlWorkSheet.Cells(3 + i, 14) = transaction.TransDate
            xlWorkSheet.Cells(3 + i, 15) = transaction.Payee
            xlWorkSheet.Cells(3 + i, 16) = category
            xlWorkSheet.Cells(3 + i, 17) = "$" & transaction.Amount.Replace("-", "")
        Next
        ' Fill in Expense footer cells
        range = xlWorkSheet.Range(xlWorkSheet.Cells(3 + ExpenseTransactions.Length, 14), xlWorkSheet.Cells(3 + ExpenseTransactions.Length, 17))
        range.Font.Bold = True
        range.Interior.Color = Color.Red
        xlWorkSheet.Cells(3 + ExpenseTransactions.Length, 14) = "Total"
        xlWorkSheet.Cells(3 + ExpenseTransactions.Length, 17) = "=SUM(Q3:Q" & (2 + ExpenseTransactions.Length).ToString() & ")"
        ' Add Borders
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 14), xlWorkSheet.Cells(3 + ExpenseTransactions.Length, 17)).BorderAround2()
        ' Format Currency
        xlWorkSheet.Range(xlWorkSheet.Cells(3, 17), xlWorkSheet.Cells(3 + ExpenseTransactions.Length, 17)).NumberFormat = currencyFormat


        '-----------------------------------------------------
        '         Create Income Table
        '-----------------------------------------------------
        ' Fill in Income header cells
        xlWorkSheet.Cells(1, 10) = "Income"
        range = xlWorkSheet.Range(xlWorkSheet.Cells(1, 10), xlWorkSheet.Cells(1, 12))
        range.Font.Bold = True
        range.Interior.Color = Color.Olive
        range.HorizontalAlignment = Excel.Constants.xlCenter
        range.Merge()

        xlWorkSheet.Range(xlWorkSheet.Cells(2, 10), xlWorkSheet.Cells(2, 12)).Interior.Color = Color.LightGray
        xlWorkSheet.Cells(2, 10) = "Date"
        xlWorkSheet.Cells(2, 11) = "Source"
        xlWorkSheet.Cells(2, 12) = "Amount"

        ' Fill in Income value cells
        For i As Integer = 0 To IncomeTranactions.Length - 1
            Dim transaction As CSVItem = IncomeTranactions.Item(i)
            xlWorkSheet.Range(xlWorkSheet.Cells(3 + i, 10), xlWorkSheet.Cells(3 + i, 12)).Interior.Color = Color.PaleGreen
            xlWorkSheet.Cells(3 + i, 10) = transaction.TransDate
            xlWorkSheet.Cells(3 + i, 11) = transaction.Payee
            xlWorkSheet.Cells(3 + i, 12) = "$" & transaction.Amount.Replace("-", "")
        Next

        ' Fill in Income footer cells
        range = xlWorkSheet.Range(xlWorkSheet.Cells(3 + IncomeTranactions.Length, 10), xlWorkSheet.Cells(3 + IncomeTranactions.Length, 12))
        range.Font.Bold = True
        range.Interior.Color = Color.Olive
        xlWorkSheet.Cells(3 + IncomeTranactions.Length, 10) = "Total"
        xlWorkSheet.Cells(3 + IncomeTranactions.Length, 12) = "=SUM(L3:L" & (2 + IncomeTranactions.Length).ToString() & ")"
        ' Add Borders
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 10), xlWorkSheet.Cells(3 + IncomeTranactions.Length, 12)).BorderAround2()
        ' Format Currency
        xlWorkSheet.Range(xlWorkSheet.Cells(3, 12), xlWorkSheet.Cells(3 + IncomeTranactions.Length, 12)).NumberFormat = currencyFormat


        '----------------------------------------------------
        '         Create Unbudgeted Summary Table
        '----------------------------------------------------
        ' Fill in Unbudgeted header cells
        xlWorkSheet.Cells(1, 7) = "Unbudgeted Spending"
        range = xlWorkSheet.Range(xlWorkSheet.Cells(1, 7), xlWorkSheet.Cells(1, 8))
        range.Font.Bold = True
        range.Interior.Color = Color.AliceBlue
        range.HorizontalAlignment = Excel.Constants.xlCenter
        range.Merge()

        xlWorkSheet.Range(xlWorkSheet.Cells(2, 7), xlWorkSheet.Cells(2, 8)).Interior.Color = Color.LightGray
        xlWorkSheet.Cells(2, 7) = "Category"
        xlWorkSheet.Cells(2, 8) = "Amount"

        'Fill in Unbudgeted value cells
        For i As Integer = 0 To UnbudgetedCategories.GenericList.Count - 1
            xlWorkSheet.Range(xlWorkSheet.Cells(3 + i, 7), xlWorkSheet.Cells(3 + i, 7)).Interior.Color = Color.LightGray
            xlWorkSheet.Range(xlWorkSheet.Cells(3 + i, 8), xlWorkSheet.Cells(3 + i, 8)).Interior.Color = Color.LightCyan
            xlWorkSheet.Cells(3 + i, 7) = UnbudgetedCategories.GenericList.Item(i).Name
            xlWorkSheet.Cells(3 + i, 8) = "=SUMIF(P3:P" & (ExpenseTransactions.Length + 2).ToString() & ",G" & (3 + i).ToString() & ",Q3:Q" & (ExpenseTransactions.Length + 2).ToString() & ")"
        Next
        

        'Fill in Unbudgeted footer cells
        range = xlWorkSheet.Range(xlWorkSheet.Cells(3 + UnbudgetedCategories.GenericList.Count, 7), xlWorkSheet.Cells(3 + UnbudgetedCategories.GenericList.Count, 8))
        range.Font.Bold = True
        range.Interior.Color = Color.AliceBlue
        xlWorkSheet.Cells(3 + UnbudgetedCategories.GenericList.Count, 7) = "Total"
        xlWorkSheet.Cells(3 + UnbudgetedCategories.GenericList.Count, 8) = "=SUM(H3:H" & (UnbudgetedCategories.GenericList.Count + 2).ToString() & ")"
        ' Add Borders
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 7), xlWorkSheet.Cells(3 + UnbudgetedCategories.GenericList.Count, 8)).BorderAround2()
        ' Format Currency
        xlWorkSheet.Range(xlWorkSheet.Cells(3, 8), xlWorkSheet.Cells(3 + UnbudgetedCategories.GenericList.Count, 8)).NumberFormat = currencyFormat


        '-----------------------------------------------------
        '         Create Budgeted Summary Table
        '-----------------------------------------------------
        ' Fill in Monthly Budget header cells
        xlWorkSheet.Cells(1, 1) = "Monthly Budget"
        range = xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(1, 5))
        range.Font.Bold = True
        range.Interior.Color = Color.AliceBlue
        range.HorizontalAlignment = Excel.Constants.xlCenter
        range.Merge()

        xlWorkSheet.Range(xlWorkSheet.Cells(2, 1), xlWorkSheet.Cells(2, 5)).Interior.Color = Color.LightGray
        xlWorkSheet.Cells(2, 1) = "Category"
        xlWorkSheet.Cells(2, 2) = "Budget"
        xlWorkSheet.Cells(2, 3) = "Actual"
        xlWorkSheet.Cells(2, 4) = "Percent"
        xlWorkSheet.Cells(2, 5) = "Percent Over Budget"

        ' Fill in Monthly Budget value cells
        For i As Integer = 0 To BudgetedCategories.GenericList.Count - 1
            Dim category As BudgetCategory = BudgetedCategories.GenericList.Item(i)
            xlWorkSheet.Range(xlWorkSheet.Cells(3 + i, 1), xlWorkSheet.Cells(3 + i, 2)).Interior.Color = Color.LightGray
            xlWorkSheet.Range(xlWorkSheet.Cells(3 + i, 3), xlWorkSheet.Cells(3 + i, 5)).Interior.Color = Color.LightCyan
            xlWorkSheet.Cells(3 + i, 1) = category.Name
            xlWorkSheet.Cells(3 + i, 2) = "$" & category.Budget
            xlWorkSheet.Cells(3 + i, 3) = "=SUMIF(P3:P" & (ExpenseTransactions.Length + 2).ToString() & ",A" & (3 + i).ToString() & ",Q3:Q" & (ExpenseTransactions.Length + 2).ToString() & ")"
            xlWorkSheet.Cells(3 + i, 4) = "=C" & (3 + i).ToString() & "/Q" & (3 + ExpenseTransactions.Length).ToString()
            xlWorkSheet.Cells(3 + i, 5) = "=(C" & (3 + i).ToString() & "-B" & (3 + i).ToString() & ")/B" & (3 + i).ToString()
        Next
        xlWorkSheet.Range(xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 1), xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 2)).Interior.Color = Color.LightGray
        xlWorkSheet.Range(xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 3), xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 5)).Interior.Color = Color.LightCyan
        xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 1) = "Extra Spending"
        xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 2) = "=" & TotalBudget.ToString() & "-SUM(B3:B" & (2 + BudgetedCategories.GenericList.Count).ToString() & ")"
        xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 3) = "=H" & (3 + UnbudgetedCategories.GenericList.Count)
        xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 4) = "=C" & (3 + BudgetedCategories.GenericList.Count).ToString() & "/Q" & (3 + ExpenseTransactions.GenericList.Count).ToString()
        xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 5) = "=(C" & (3 + BudgetedCategories.GenericList.Count).ToString() & "-B" & (3 + BudgetedCategories.GenericList.Count).ToString() & ")/B" & (3 + BudgetedCategories.GenericList.Count).ToString()

        ' Add Borders
        xlWorkSheet.Range(xlWorkSheet.Cells(1, 1), xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 5)).BorderAround2()
        ' Format Currency & percent
        xlWorkSheet.Range(xlWorkSheet.Cells(3, 2), xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 3)).NumberFormat = currencyFormat
        xlWorkSheet.Range(xlWorkSheet.Cells(3, 4), xlWorkSheet.Cells(3 + BudgetedCategories.GenericList.Count, 5)).NumberFormat = percentFormat


        ' Fill in Monthly Budget footer cells
        range = xlWorkSheet.Range(xlWorkSheet.Cells(4 + BudgetedCategories.GenericList.Count, 1), xlWorkSheet.Cells(4 + BudgetedCategories.GenericList.Count, 3))
        range.Font.Bold = True
        range.Interior.Color = Color.AliceBlue
        range.BorderAround2()
        range.NumberFormat = currencyFormat
        xlWorkSheet.Cells(4 + BudgetedCategories.GenericList.Count, 1) = "Total"
        xlWorkSheet.Cells(4 + BudgetedCategories.GenericList.Count, 2) = "=SUM(B3:B" & (3 + BudgetedCategories.GenericList.Count).ToString() & ")"
        xlWorkSheet.Cells(4 + BudgetedCategories.GenericList.Count, 3) = "=SUM(C3:C" & (3 + BudgetedCategories.GenericList.Count).ToString() & ")"

        range = xlWorkSheet.Range(xlWorkSheet.Cells(6 + BudgetedCategories.GenericList.Count, 1), xlWorkSheet.Cells(6 + BudgetedCategories.GenericList.Count, 2))
        range.Font.Bold = True
        range.Interior.Color = Color.AliceBlue
        range.BorderAround2()
        range.NumberFormat = currencyFormat
        xlWorkSheet.Cells(6 + BudgetedCategories.GenericList.Count, 1) = "Amount Saved"
        xlWorkSheet.Cells(6 + BudgetedCategories.GenericList.Count, 2) = "=L" & (3 + IncomeTranactions.Length).ToString() & "-C" & (4 + BudgetedCategories.GenericList.Count).ToString()


        ' Adjust cell widths
        xlWorkSheet.Cells.EntireColumn.AutoFit()

        ' Add newly created worksheet to workbook
        'xlWorkBook.Worksheets.Add(xlWorkSheet)
        If rdoExisting.Checked Then
            xlWorkBook.Save()
        Else
            xlWorkBook.SaveAs(txtPath.Text & "\" & txtName.Text)
        End If
        xlWorkBook.Close()

        releaseObject(xlWorkSheet)
        releaseObject(xlWorkBook)
        releaseObject(xlApp)

        MessageBox.Show("New Worksheet Created!", "Worksheet Saved")
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub btnSave_Click() Handles btnSave.Click
        Try
            exportToExcel()
            If rdoExisting.Checked Then
                My.Settings.ExportFile = txtFile.Text
            Else
                My.Settings.ExportFile = txtPath.Text & "\" & txtName.Text
            End If
            My.Settings.Save()
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Could not export", "Error")
        End Try
    End Sub

    Public Sub prepareExport(ByVal list As CSVList, ByVal categories As ListOfBudgetCategories, ByVal total As Decimal)
        TotalBudget = total
        For Each category As BudgetCategory In categories.GenericList
            If category.Used And category.Budget > 0 And category.Type = "E" Then
                BudgetedCategories.GenericList.Add(category)
            ElseIf category.Used And category.Budget = 0 And category.Type = "E" Then
                UnbudgetedCategories.GenericList.Add(category)
            End If
        Next
        For Each csv As CSVItem In list.GenericList
            If categories.GetCategoryByName(categories.PayeeCategory(csv.Payee)).Type = "I" Then
                IncomeTranactions.Add(csv)
            Else
                ExpenseTransactions.Add(csv)
            End If
        Next

        Me.Show()
    End Sub
End Class