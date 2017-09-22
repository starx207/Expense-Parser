Public Class BudgetCategory
    Public Property Name As String
    Public Property Type As BudgetTypes
    Public Property Budget As Double
    Public Property Payees As List(Of String)
    Private _inUse As Boolean
    Public Property Used() As Boolean
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
