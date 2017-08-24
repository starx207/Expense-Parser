Public Class CSVList
    Private UniqueCSVs As New List(Of String)
    Private AllCSV As New List(Of CSVItem)
    Public ReadOnly Property Length() As Integer
        Get
            Return AllCSV.Count
        End Get
    End Property
    Public ReadOnly Property Item(ByVal i As Integer) As CSVItem
        Get
            Return AllCSV.Item(i)
        End Get
    End Property
    Public ReadOnly Property UniqueItems() As List(Of String)
        Get
            Return UniqueCSVs
        End Get
    End Property
    Public Property GenericList() As List(Of CSVItem)
        Get
            Return AllCSV
        End Get
        Set(value As List(Of CSVItem))
            AllCSV = value
        End Set
    End Property

    Public Function IndexOf(ByVal payee As String) As Integer
        Dim i As Integer = 0
        For Each csv As CSVItem In AllCSV
            If csv.Payee = payee Then
                Return i
            End If
            i += 1
        Next
        Return -1
    End Function

    Public Sub RemoveAt(ByVal i As Integer)
        Dim payeeRemoved As String = AllCSV(i).Payee
        AllCSV.RemoveAt(i)

        ' Check if the payee has other occurances
        For Each csv As CSVItem In AllCSV
            If csv.Payee = payeeRemoved Then
                Exit Sub
            End If
        Next
        ' If no occurances found, remove from unique list
        UniqueCSVs.Remove(payeeRemoved)
    End Sub

    Public Sub Clear()
        AllCSV.Clear()
        UniqueCSVs.Clear()
    End Sub

    Public Sub Add(ByVal csvDate As String, ByVal csvName As String, ByVal csvAmt As String)
        Dim newCSV As New CSVItem
        newCSV.TransDate = csvDate
        newCSV.Payee = csvName
        newCSV.Amount = csvAmt

        AllCSV.Add(newCSV)

        If UniqueCSVs.IndexOf(csvName) = -1 Then
            UniqueCSVs.Add(csvName)
            UniqueCSVs.Sort()
        End If
    End Sub

    Public Sub Add(ByVal item As CSVItem)
        Me.Add(item.TransDate, item.Payee, item.Amount)
    End Sub
End Class
