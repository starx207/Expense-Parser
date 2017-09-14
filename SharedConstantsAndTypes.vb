Imports System.ComponentModel

Public Module SharedConstantsAndTypes
    Public Enum BudgetTypes
        <Description("E")>
        Expense = 1
        <Description("I")>
        Income
    End Enum

    Public Const UnassignedPayeeType As String = "Misc"
    Public Const NoExportFile As String = "None"
    Public Const UnassignedCategory As String = "Unassigned"
    Public Const CurrencyFormat As String = "$###,##0.00"

End Module
