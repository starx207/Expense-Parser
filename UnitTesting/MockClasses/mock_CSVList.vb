Imports Expense_Parser

Public Class mock_CSVList
    Implements ICSVList

    Sub New()
        GenericList = New List(Of ICSVItem)
    End Sub

    Public Property GenericList As List(Of ICSVItem) Implements ICSVList.GenericList

    Public ReadOnly Property Item(i As Integer) As ICSVItem Implements ICSVList.Item
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property Length As Integer Implements ICSVList.Length
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property UniqueItems As List(Of String) Implements ICSVList.UniqueItems
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Sub Add(item As ICSVItem) Implements ICSVList.Add
        Throw New NotImplementedException()
    End Sub

    Public Sub Add(csvDate As String, csvName As String, csvAmt As String) Implements ICSVList.Add
        Throw New NotImplementedException()
    End Sub

    Public Sub RemoveAt(i As Integer) Implements ICSVList.RemoveAt
        Throw New NotImplementedException()
    End Sub

    Public Function IndexOf(payee As String) As Integer Implements ICSVList.IndexOf
        Dim i As Integer = 0
        For Each csv As ICSVItem In GenericList
            If csv.Payee = payee Then
                Return i
            End If
            i += 1
        Next
        Return -1
    End Function
End Class
