Imports Expense_Parser
Imports System.Diagnostics.CodeAnalysis

<ExcludeFromCodeCoverage>
Public Class mock_BudgetCategory
    Implements IBudgetCategory

    Sub New()
        Payees = New List(Of String)
    End Sub

    Public Property Budget As Double Implements IBudgetCategory.Budget

    Public Property Name As String Implements IBudgetCategory.Name

    Public Property Payees As List(Of String) Implements IBudgetCategory.Payees

    Public Property Type As BudgetTypes Implements IBudgetCategory.Type

    Public Property Used As Boolean Implements IBudgetCategory.Used

    Public Shared Function ConvertFromIBudgetCategory(ByVal category As IBudgetCategory) As mock_BudgetCategory
        Dim mock As New mock_BudgetCategory
        mock.Name = category.Name
        mock.Payees = category.Payees
        mock.Type = category.Type
        mock.Budget = category.Budget
        mock.Used = category.Used

        Return mock
    End Function
End Class
