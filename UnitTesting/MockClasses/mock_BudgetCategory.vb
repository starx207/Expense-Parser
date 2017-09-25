Imports Expense_Parser

Public Class mock_BudgetCategory
    Implements IBudgetCategory

    Public Property Budget As Double Implements IBudgetCategory.Budget

    Public Property Name As String Implements IBudgetCategory.Name

    Public Property Payees As List(Of String) Implements IBudgetCategory.Payees

    Public Property Type As BudgetTypes Implements IBudgetCategory.Type

    Public Property Used As Boolean Implements IBudgetCategory.Used
End Class
