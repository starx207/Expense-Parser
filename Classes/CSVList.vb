Imports Microsoft.VisualBasic.FileIO

Public Class CSVList
    Private ReadOnly AllowableDateColumns As String() = {"DATE", "TRANSACTION DATE"}
    Private ReadOnly AllowablePayeeColumns As String() = {"PAYEE NAME", "MERCHANT"}
    Private ReadOnly AllowableAmountColumns As String() = {"AMOUNT", "BILLING AMOUNT"}

    Private UniqueCSVs As New List(Of String)
    Public Property GenericList As List(Of ICSVItem)
    Public ReadOnly Property Length() As Integer
        Get
            Return GenericList.Count
        End Get
    End Property
    Public ReadOnly Property Item(ByVal i As Integer) As ICSVItem
        Get
            Return GenericList.Item(i)
        End Get
    End Property
    Public ReadOnly Property UniqueItems() As List(Of String)
        Get
            Return UniqueCSVs
        End Get
    End Property

    Public Function IndexOf(ByVal payee As String) As Integer
        Dim i As Integer = 0
        For Each csv As ICSVItem In GenericList
            If csv.Payee = payee Then
                Return i
            End If
            i += 1
        Next
        Return -1
    End Function

    Public Sub RemoveAt(ByVal i As Integer)
        Dim payeeRemoved As String = GenericList(i).Payee
        GenericList.RemoveAt(i)

        ' Check if the payee has other occurances
        For Each csv As ICSVItem In GenericList
            If csv.Payee = payeeRemoved Then
                Exit Sub
            End If
        Next
        ' If no occurances found, remove from unique list
        UniqueCSVs.Remove(payeeRemoved)
    End Sub

    Public Sub Add(ByVal csvDate As String, ByVal csvName As String, ByVal csvAmt As String)
        Dim newCSV As New CSVItem
        newCSV.TransDate = csvDate
        newCSV.Payee = csvName
        newCSV.Amount = csvAmt

        GenericList.Add(newCSV)

        If UniqueCSVs.IndexOf(csvName) = -1 Then
            UniqueCSVs.Add(csvName)
            UniqueCSVs.Sort()
        End If
    End Sub

    Public Sub Add(ByVal item As ICSVItem)
        Add(item.TransDate, item.Payee, item.Amount)
    End Sub

    Sub New()
        GenericList = New List(Of ICSVItem)
    End Sub

    Sub New(ByVal csv As TextFieldParser)
        GenericList = New List(Of ICSVItem)
        Dim currentRow As String()
        Dim dateIndex As Integer
        Dim payeeIndex As Integer
        Dim amountIndex As Integer
        Dim columnName As String

        csv.TextFieldType = FieldType.Delimited
        csv.SetDelimiters(",")

        ' Set indexes for relavant columns
        currentRow = csv.ReadFields() ' Read header Row
        If currentRow IsNot Nothing Then
            For i As Integer = 0 To currentRow.Length - 1
                columnName = currentRow.GetValue(i).ToString().ToUpper()
                Select Case True
                    Case AllowableDateColumns.Contains(columnName)
                        dateIndex = i
                    Case AllowablePayeeColumns.Contains(columnName)
                        payeeIndex = i
                    Case AllowableAmountColumns.Contains(columnName)
                        amountIndex = i
                    Case Else
                        ' Do Nothing
                End Select
            Next
        End If

        While Not csv.EndOfData
            Try
                currentRow = csv.ReadFields()
                If currentRow IsNot Nothing Then
                    Add(currentRow.GetValue(dateIndex),
                        currentRow.GetValue(payeeIndex),
                        currentRow.GetValue(amountIndex))
                End If
            Catch
            End Try
        End While

        csv.Close()
    End Sub
End Class
