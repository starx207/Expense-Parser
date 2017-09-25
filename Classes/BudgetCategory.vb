Public Class BudgetCategory
    Implements IBudgetCategory
    Public Property Name As String Implements IBudgetCategory.Name
    Public Property Type As BudgetTypes Implements IBudgetCategory.Type
    Public Property Budget As Double Implements IBudgetCategory.Budget
    Public Property Payees As List(Of String) Implements IBudgetCategory.Payees
    Private _inUse As Boolean
    Public Property Used() As Boolean Implements IBudgetCategory.Used
        Get
            If Budget > 0 Then
                Return True
            End If
            Return _inUse
        End Get
        Set(value As Boolean)
            If Budget <= 0 Then
                _inUse = value
            End If
        End Set
    End Property


    Sub New()
        Name = ""
        Type = BudgetTypes.Expense
        Budget = 0
        _inUse = False
        Payees = New List(Of String)
    End Sub

    Sub New(ByVal newName As String, ByVal newType As BudgetTypes, ByVal newBudget As Double)
        Name = newName
        Type = newType
        Budget = newBudget
        _inUse = True
        Payees = New List(Of String)
    End Sub
End Class
