Imports Expense_Parser

Public Interface ICSVList
    Property GenericList As List(Of ICSVItem)
    ReadOnly Property Item(i As Integer) As ICSVItem
    ReadOnly Property Length As Integer
    ReadOnly Property UniqueItems As List(Of String)
    Sub Add(item As ICSVItem)
    Sub Add(csvDate As String, csvName As String, csvAmt As String)
    Sub RemoveAt(i As Integer)
    Function IndexOf(payee As String) As Integer
End Interface
