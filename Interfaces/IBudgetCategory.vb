Imports Expense_Parser

Public Interface IBudgetCategory
    Property Budget As Double
    Property Name As String
    Property Payees As List(Of String)
    Property Type As BudgetTypes
    Property Used As Boolean
End Interface
