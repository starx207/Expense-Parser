Imports Microsoft.Office.Interop
Imports System.ComponentModel

Public Class ExcelExporter
    Private IncomeTransactions As New CSVList
    Private ExpenseTransactions As New CSVList
    Private BudgetedCategories As New Budget
    Private UnbudgetedCategories As New Budget
    Private TotalBudget As Double
    Private xlApp As Excel.Application
    Private xlWorkbook As Excel.Workbook
    Private xlWorksheet As Excel.Worksheet

    Private ReadOnly BudgetHeaderColor As Color = Color.AliceBlue
    Private ReadOnly BudgetSubheaderColor As Color = Color.LightGray
    Private ReadOnly BudgetValueColor As Color = Color.LightCyan
    Private Const IncomeTableName As String = "Income"
    Private Const ExpenseTableName As String = "Expense"
    Private Const PercentFormat As String = "##0.0%"
    Private Const ExpenseColumnStart As Integer = 14
    Private Const IncomeColumnStart As Integer = 10
    Private Const UnbudgetedColumnStart As Integer = 7
    Private Const BudgetedColumnStart As Integer = 1
    Private Const UnknownMonth As String = "Unknown"

    Public Enum TransactionColumns
        [Date]
        <Description("Source")>
        Payee
        Category
        Amount
    End Enum

    Private Enum MonthNames
        January = 1
        February
        March
        April
        May
        June
        July
        August
        September
        October
        November
        December
    End Enum

    Public Sub ExportMonth(ByVal fileName As String, Optional ByVal fromNew As Boolean = True)
        Dim sheetName As String
        Dim sheetFound As Boolean
        Dim numSheetMatch As Integer = 2
        Dim range As Excel.Range
        Dim expenseColumns As New List(Of TransactionColumns) From {TransactionColumns.Date, TransactionColumns.Payee, TransactionColumns.Category, TransactionColumns.Amount}
        Dim incomeColumns As New List(Of TransactionColumns) From {TransactionColumns.Date, TransactionColumns.Payee, TransactionColumns.Amount}
        Dim expenseAmountColumn As String = GetExcelColumn(expenseColumns.IndexOf(TransactionColumns.Amount) + ExpenseColumnStart)
        Dim expenseCategoryColumn As String = GetExcelColumn(expenseColumns.IndexOf(TransactionColumns.Category) + ExpenseColumnStart)
        Dim incomeAmountColumn As String = GetExcelColumn(incomeColumns.IndexOf(TransactionColumns.Amount) + IncomeColumnStart)

        sheetName = SetSheetName(New List(Of CSVList) From {IncomeTransactions, ExpenseTransactions})

        ' Verify Excel is installed on computer
        xlApp = New Excel.Application
        If xlApp Is Nothing Then
            Throw New NotSupportedException("Excel is not properly installed!")
        End If
        ' Disable Excel overwrite prompt
        xlApp.DisplayAlerts = False

        ' Prepare Workbook
        If fromNew Then
            xlWorkbook = xlApp.Workbooks.Add(Reflection.Missing.Value)
        Else
            Try
                xlWorkbook = xlApp.Workbooks.Open(fileName)
            Catch
                My.Settings.ExportFile = NoExportFile
                My.Settings.Save()
                Throw New InvalidOperationException("File not found. Please check that it still exists")
            End Try
        End If

        ' Prepare Worksheet
        xlWorksheet = DirectCast(xlWorkbook.Worksheets.Add(), Excel.Worksheet)
        xlWorksheet.Select()

        ' Prepare a name for worksheet
        ' Default name is Month
        ' If that name already exists, append a number to the end
        Do
            sheetFound = False
            For Each sheet As Excel.Worksheet In xlWorkbook.Worksheets
                If sheet.Name = sheetName Then
                    sheetFound = True
                    If numSheetMatch > 2 Then
                        sheetName = sheetName.Substring(0, sheetName.Length - numSheetMatch.ToString().Length)
                    End If
                    sheetName = sheetName + numSheetMatch.ToString()
                    numSheetMatch += 1
                    Exit For
                End If
            Next
        Loop While sheetFound

        xlWorksheet.Name = sheetName

        ' Create Expense Table
        CreateExcelTransactionTable(1, ExpenseColumnStart, ExpenseTransactions, Color.Red, Color.LightGray, Color.LightCoral, ExpenseTableName, expenseColumns)
        ' Create Income Table
        CreateExcelTransactionTable(1, IncomeColumnStart, IncomeTransactions, Color.Olive, Color.LightGray, Color.PaleGreen, IncomeTableName, incomeColumns)

        With xlWorksheet
            '--------------------------------------------
            '         Unbudgeted Summary Table
            '--------------------------------------------
            ' Fill in Unbudgeted header cells
            .Cells(1, UnbudgetedColumnStart) = "Unbudgeted Spending"
            range = .Range(.Cells(1, UnbudgetedColumnStart), .Cells(1, UnbudgetedColumnStart + 1))
            range.Font.Bold = True
            range.Interior.Color = BudgetHeaderColor
            range.HorizontalAlignment = Excel.Constants.xlCenter
            range.Merge()

            .Range(.Cells(2, UnbudgetedColumnStart), .Cells(2, UnbudgetedColumnStart + 1)).Interior.Color = BudgetSubheaderColor
            .Cells(2, UnbudgetedColumnStart) = "Category"
            .Cells(2, UnbudgetedColumnStart + 1) = "Amount"

            ' Fill in Unbudgeted value cells
            For i As Integer = 0 To UnbudgetedCategories.Categories.Count - 1
                .Range(.Cells(3 + i, UnbudgetedColumnStart), .Cells(3 + i, UnbudgetedColumnStart)).Interior.Color = BudgetSubheaderColor
                .Range(.Cells(3 + i, UnbudgetedColumnStart + 1), .Cells(3 + i, UnbudgetedColumnStart + 1)).Interior.Color = BudgetValueColor
                .Cells(3 + i, UnbudgetedColumnStart) = UnbudgetedCategories.Categories.Item(i).Name
                .Cells(3 + i, UnbudgetedColumnStart + 1) = "=SUMIF(" + expenseCategoryColumn + "3:" + expenseCategoryColumn + (ExpenseTransactions.Length + 2).ToString() +
                    "," + GetExcelColumn(UnbudgetedColumnStart) + (3 + i).ToString() +
                    "," + expenseAmountColumn + "3:" + expenseAmountColumn + (ExpenseTransactions.Length + 2).ToString() + ")"
            Next

            'Fill in Unbudgeted footer cells
            range = .Range(.Cells(3 + UnbudgetedCategories.Categories.Count, UnbudgetedColumnStart), .Cells(3 + UnbudgetedCategories.Categories.Count, UnbudgetedColumnStart + 1))
            range.Font.Bold = True
            range.Interior.Color = BudgetHeaderColor
            .Cells(3 + UnbudgetedCategories.Categories.Count, UnbudgetedColumnStart) = "Total"
            .Cells(3 + UnbudgetedCategories.Categories.Count, UnbudgetedColumnStart + 1) = "=SUM(" + GetExcelColumn(UnbudgetedColumnStart + 1) + "3:" + GetExcelColumn(UnbudgetedColumnStart + 1) + (UnbudgetedCategories.Categories.Count + 2).ToString() + ")"
            ' Add Borders
            .Range(.Cells(1, UnbudgetedColumnStart), .Cells(3 + UnbudgetedCategories.Categories.Count, UnbudgetedColumnStart + 1)).BorderAround2()
            ' Format Currency
            .Range(.Cells(3, UnbudgetedColumnStart + 1), .Cells(3 + UnbudgetedCategories.Categories.Count, UnbudgetedColumnStart + 1)).NumberFormat = CurrencyFormat


            '-----------------------------------------------------
            '         Create Budgeted Summary Table
            '-----------------------------------------------------
            ' Fill in Monthly Budget header cells
            .Cells(1, BudgetedColumnStart) = "Monthly Budget"
            range = .Range(.Cells(1, BudgetedColumnStart), .Cells(1, BudgetedColumnStart + 4))
            range.Font.Bold = True
            range.Interior.Color = BudgetHeaderColor
            range.HorizontalAlignment = Excel.Constants.xlCenter
            range.Merge()

            .Range(.Cells(2, BudgetedColumnStart), .Cells(2, BudgetedColumnStart + 4)).Interior.Color = BudgetSubheaderColor
            .Cells(2, BudgetedColumnStart) = "Category"
            .Cells(2, BudgetedColumnStart + 1) = "Budget"
            .Cells(2, BudgetedColumnStart + 2) = "Actual"
            .Cells(2, BudgetedColumnStart + 3) = "Percent"
            .Cells(2, BudgetedColumnStart + 4) = "Percent Over Budget"

            ' Fill in Monthly Budget value cells
            For i As Integer = 0 To BudgetedCategories.Categories.Count - 1
                Dim category As BudgetCategory = BudgetedCategories.Categories.Item(i)
                .Range(.Cells(3 + i, BudgetedColumnStart), .Cells(3 + i, BudgetedColumnStart + 1)).Interior.Color = BudgetSubheaderColor
                .Range(.Cells(3 + i, BudgetedColumnStart + 2), .Cells(3 + i, BudgetedColumnStart + 4)).Interior.Color = BudgetValueColor
                .Cells(3 + i, BudgetedColumnStart) = category.Name
                .Cells(3 + i, BudgetedColumnStart + 1) = category.Budget
                .Cells(3 + i, BudgetedColumnStart + 2) = "=SUMIF(" + expenseCategoryColumn + "3:" + expenseCategoryColumn + (ExpenseTransactions.Length + 2).ToString() +
                    "," + GetExcelColumn(BudgetedColumnStart) + (3 + i).ToString() +
                    "," + expenseAmountColumn + "3:" + expenseAmountColumn + (ExpenseTransactions.Length + 2).ToString() + ")"
                .Cells(3 + i, BudgetedColumnStart + 3) = "=" + GetExcelColumn(BudgetedColumnStart + 2) + (3 + i).ToString() + "/" + expenseAmountColumn + (3 + ExpenseTransactions.Length).ToString()
                .Cells(3 + i, BudgetedColumnStart + 4) = "=(" + GetExcelColumn(BudgetedColumnStart + 2) + (3 + i).ToString() + "-" + GetExcelColumn(BudgetedColumnStart + 1) + (3 + i).ToString() + ")/" + GetExcelColumn(BudgetedColumnStart + 1) + (3 + i).ToString()
            Next

            .Range(.Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart), .Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 1)).Interior.Color = BudgetSubheaderColor
            .Range(.Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 2), .Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 4)).Interior.Color = BudgetValueColor
            .Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart) = "Extra Spending"
            .Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 1) = "=" + TotalBudget.ToString() + "-SUM(" + GetExcelColumn(BudgetedColumnStart + 1) + "3:" + GetExcelColumn(BudgetedColumnStart + 1) + (2 + BudgetedCategories.Categories.Count).ToString() + ")"
            .Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 2) = "=" + GetExcelColumn(UnbudgetedColumnStart + 1) + (3 + UnbudgetedCategories.Categories.Count).ToString()
            .Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 3) = "=" + GetExcelColumn(BudgetedColumnStart + 2) + (3 + BudgetedCategories.Categories.Count).ToString() + "/" + expenseAmountColumn + (3 + ExpenseTransactions.GenericList.Count).ToString()
            .Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 4) = "=(" + GetExcelColumn(BudgetedColumnStart + 2) + (3 + BudgetedCategories.Categories.Count).ToString() + "-" + GetExcelColumn(BudgetedColumnStart + 1) + (3 + BudgetedCategories.Categories.Count).ToString() + ")/" + GetExcelColumn(BudgetedColumnStart + 1) + (3 + BudgetedCategories.Categories.Count).ToString()

            ' Add Borders
            .Range(.Cells(1, BudgetedColumnStart), .Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 4)).BorderAround2()
            ' Format Currency and percent
            .Range(.Cells(3, BudgetedColumnStart + 1), .Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 2)).NumberFormat = CurrencyFormat
            .Range(.Cells(3, BudgetedColumnStart + 3), .Cells(3 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 4)).NumberFormat = PercentFormat

            ' Fill in Monthly Budget footer cells
            range = .Range(.Cells(4 + BudgetedCategories.Categories.Count, BudgetedColumnStart), .Cells(4 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 2))
            range.Font.Bold = True
            range.Interior.Color = BudgetHeaderColor
            range.BorderAround2()
            range.NumberFormat = CurrencyFormat
            .Cells(4 + BudgetedCategories.Categories.Count, BudgetedColumnStart) = "Total"
            .Cells(4 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 1) = "=SUM(" + GetExcelColumn(BudgetedColumnStart + 1) + "3:" + GetExcelColumn(BudgetedColumnStart + 1) + (3 + BudgetedCategories.Categories.Count).ToString() + ")"
            .Cells(4 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 2) = "=SUM(" + GetExcelColumn(BudgetedColumnStart + 2) + "3:" + GetExcelColumn(BudgetedColumnStart + 2) + (3 + BudgetedCategories.Categories.Count).ToString() + ")"

            ' Add total saved line
            range = .Range(.Cells(6 + BudgetedCategories.Categories.Count, BudgetedColumnStart), .Cells(6 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 1))
            range.Font.Bold = True
            range.Interior.Color = BudgetHeaderColor
            range.BorderAround2()
            range.NumberFormat = CurrencyFormat
            .Cells(6 + BudgetedCategories.Categories.Count, BudgetedColumnStart) = "Amount Saved"
            .Cells(6 + BudgetedCategories.Categories.Count, BudgetedColumnStart + 1) = "=" + incomeAmountColumn + (3 + IncomeTransactions.Length).ToString() + "-" + GetExcelColumn(BudgetedColumnStart + 2) + (4 + BudgetedCategories.Categories.Count).ToString()

            ' Adjust cell widths
            .Cells.EntireColumn.AutoFit()
        End With

        If Not fromNew Then
            xlWorkbook.Save()
        Else
            xlWorkbook.SaveAs(fileName)
        End If
        xlWorkbook.Close()

        releaseObject(xlWorksheet)
        releaseObject(xlWorkbook)
        releaseObject(xlApp)
    End Sub

    Private Sub CreateExcelTransactionTable(ByVal startingRow As Integer,
                                            ByVal startingColumn As Integer,
                                            ByVal transactions As CSVList,
                                            ByVal headColor As Color,
                                            ByVal subHeadColor As Color,
                                            ByVal mainColor As Color,
                                            ByVal tableTitle As String,
                                            ByVal includedColumns As List(Of TransactionColumns))
        Dim range As Excel.Range
        Dim headerRow As Integer = startingRow
        Dim subHeaderRow As Integer = startingRow + 1
        Dim firstValueRow As Integer = startingRow + 2
        Dim footerRow As Integer = startingRow + 2 + transactions.Length
        Dim colOffset As Integer = includedColumns.Count - 1
        Dim amountCol As Integer = -1

        With xlWorksheet
            ' Fill in header cells
            range = .Range(.Cells(headerRow, startingColumn), .Cells(headerRow, startingColumn + colOffset))
            .Cells(headerRow, startingColumn) = tableTitle
            range.Font.Bold = True
            range.Interior.Color = headColor
            range.HorizontalAlignment = Excel.Constants.xlCenter
            range.Merge()

            .Range(.Cells(subHeaderRow, startingColumn), .Cells(subHeaderRow, startingColumn + colOffset)).Interior.Color = subHeadColor
            For i As Integer = 0 To colOffset
                ' TODO: Fix this. Not printing the sub header names
                Dim columnTitle As String
                If tableTitle = IncomeTableName And includedColumns.Item(i) = TransactionColumns.Payee Then
                    columnTitle = includedColumns.Item(i).GetEnumDescription()
                Else
                    columnTitle = includedColumns.Item(i).ToString()
                End If
                .Cells(subHeaderRow, startingColumn + i) = columnTitle
            Next

            ' Fill in value cells
            For i As Integer = 0 To transactions.Length - 1
                Dim transaction As CSVItem = transactions.Item(i)
                Dim category As String = BudgetedCategories.GetAssignedCategory(transaction.Payee)
                If category = UnassignedCategory Then
                    category = UnbudgetedCategories.GetAssignedCategory(transaction.Payee)
                End If
                .Range(.Cells(firstValueRow + i, startingColumn), .Cells(firstValueRow + i, startingColumn + colOffset)).Interior.Color = mainColor
                For j As Integer = 0 To colOffset
                    Dim cellValue As String = ""
                    Select Case includedColumns.Item(j)
                        Case TransactionColumns.Amount
                            cellValue = transaction.Amount.Replace("-", "")
                            amountCol = startingColumn + j
                        Case TransactionColumns.Category
                            cellValue = category
                        Case TransactionColumns.Date
                            cellValue = transaction.TransDate
                        Case TransactionColumns.Payee
                            cellValue = transaction.Payee
                    End Select
                    .Cells(firstValueRow + i, startingColumn + j) = cellValue
                Next
            Next

            ' Fill in Expense Footer cells
            If amountCol <> -1 Then
                range = .Range(.Cells(footerRow, startingColumn), .Cells(footerRow, startingColumn + colOffset))
                range.Font.Bold = True
                range.Interior.Color = headColor
                .Cells(footerRow, startingColumn) = "Total"
                .Cells(footerRow, startingColumn + colOffset) = "=SUM(" + GetExcelColumn(amountCol) + firstValueRow.ToString() + ":" + GetExcelColumn(amountCol) + (footerRow - 1).ToString() + ")"
                ' Add Borders
                .Range(.Cells(headerRow, startingColumn), .Cells(footerRow, startingColumn + colOffset)).BorderAround2()
                ' Format Currency
                .Range(.Cells(firstValueRow, amountCol), .Cells(footerRow, amountCol)).NumberFormat = CurrencyFormat
            End If
        End With
    End Sub

    Private Function SetSheetName(ByVal transactions As List(Of CSVList)) As String
        Dim monthName As String = ""
        Dim sheetName As String = ""
        For Each list As CSVList In transactions
            For Each transaction As CSVItem In list.GenericList
                monthName = GetMonthName(transaction.TransDate)
                If sheetName = "" Then
                    sheetName = monthName
                ElseIf sheetName <> monthName Then
                    Throw New ConstraintException("Source file has transactions for more than 1 month (MM/dd/yyyy)")
                End If
            Next
        Next
        Return monthName
    End Function

    Private Sub releaseObject(ByVal obj As Object)
        Try
            Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Public Sub PrepareExport(ByVal transactionList As CSVList, ByVal sourceBudget As Budget)
        TotalBudget = sourceBudget.TotalBudget
        BudgetedCategories = New Budget
        UnbudgetedCategories = New Budget
        IncomeTransactions = New CSVList
        ExpenseTransactions = New CSVList
        For Each category As BudgetCategory In sourceBudget.Categories
            If category.Used And category.Budget > 0 And category.Type = BudgetTypes.Expense Then
                BudgetedCategories.Categories.Add(category)
            ElseIf category.Used And category.Budget = 0 And category.Type = BudgetTypes.Expense Then
                UnbudgetedCategories.Categories.Add(category)
            End If
        Next

        For Each csv As CSVItem In transactionList.GenericList
            If sourceBudget.GetCategoryByName(sourceBudget.PayeeCategory(csv.Payee)).Type = BudgetTypes.Income Then
                IncomeTransactions.Add(csv)
            Else
                ExpenseTransactions.Add(csv)
            End If
        Next
    End Sub

    Private Function GetExcelColumn(ByVal colNum As Integer) As String
        Dim columnName As String = ""
        Const lettersInAlphabet As Integer = 26
        Dim quotient As Integer
        Dim remainder As Integer
        colNum -= 1

        remainder = colNum Mod lettersInAlphabet
        quotient = Math.Floor(colNum / lettersInAlphabet)

        If quotient > 0 Then
            columnName = GetExcelColumn(quotient)
        End If

        columnName += Char.ConvertFromUtf32(remainder + 65)

        Return columnName
    End Function

    Private Function GetMonthName(ByVal checkDate As String) As String
        Dim indexOfEOM As Integer = checkDate.IndexOf("/")
        Dim monthNum As Integer = checkDate.Substring(0, indexOfEOM)
        Dim monthName As String

        If [Enum].IsDefined(GetType(MonthNames), monthNum) Then
            monthName = DirectCast(monthNum, MonthNames).ToString()
        Else
            monthName = UnknownMonth
        End If

        Return monthName
    End Function
End Class
