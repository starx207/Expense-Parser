Public Class BudgetCategory
    Private _name As String
    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property
    Private _type As Char
    Public Property Type() As Char
        Get
            Return _type
        End Get
        Set(ByVal value As Char)
            _type = value
        End Set
    End Property
    Private _budget As Double
    Public Property Budget() As Double
        Get
            Return _budget
        End Get
        Set(ByVal value As Double)
            _budget = value
        End Set
    End Property
    Private _payees As List(Of String)
    Public Property Payees() As List(Of String)
        Get
            Return _payees
        End Get
        Set(ByVal value As List(Of String))
            _payees = value
        End Set
    End Property
    Private _inUse As Boolean
    Public Property Used() As Boolean
        Get
            If _budget > 0 Then
                Return True
            End If
            Return _inUse
        End Get
        Set(value As Boolean)
            If _budget = 0 Then
                _inUse = value
            End If
        End Set
    End Property


    Sub New()
        _name = ""
        _type = ""
        _budget = 0
        _inUse = False
        _payees = New List(Of String)
    End Sub

    Sub New(ByVal newName As String, ByVal newType As Char, ByVal newBudget As Double)
        _name = newName
        _type = newType
        _budget = newBudget
        _inUse = True
        _payees = New List(Of String)
    End Sub
End Class
