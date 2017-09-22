Imports Expense_Parser

Public Class mock_CSVItem
    Implements ICSVItem

    Public Property Amount As String Implements ICSVItem.Amount

    Public Property Payee As String Implements ICSVItem.Payee

    Public Property TransDate As String Implements ICSVItem.TransDate
End Class
