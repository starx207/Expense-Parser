Imports System.ComponentModel

Module Constants
    Public Const RootNode As String = "MonthlyBudget"
    Public Const TotalBudgetNode As String = "TotalBudget"
    Public Const CategoryNode As String = "Category"
    Public Const NameNode As String = "Name"
    Public Const TypeNode As String = "Type"
    Public Const BudgetNode As String = "Budget"
    Public Const PayeesNode As String = "Payees"

    Public ReadOnly AccessTotalBudgetNode As String = RootNode + "/" + TotalBudgetNode
    Public ReadOnly AccessCategoryNode As String = RootNode + "/" + CategoryNode
    Public ReadOnly AccessPayeeNameNodes As String = PayeesNode + "/" + NameNode

    Public Enum BudgetType
        <Description("E")>
        Expense
        <Description("I")>
        Income
    End Enum
End Module
