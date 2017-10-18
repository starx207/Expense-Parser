Imports Expense_Parser
Imports System.Diagnostics.CodeAnalysis

<ExcludeFromCodeCoverage>
Public Class mock_CSVItem
    Implements ICSVItem

    Public Property Amount As String Implements ICSVItem.Amount

    Public Property Payee As String Implements ICSVItem.Payee

    Public Property TransDate As String Implements ICSVItem.TransDate
End Class
