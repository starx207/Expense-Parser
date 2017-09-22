Public Class CSVItem
    Implements ICSVItem
    Public Property TransDate As String Implements ICSVItem.TransDate
    Public Property Payee As String Implements ICSVItem.Payee
    Private TransactionAmount As String
    Public Property Amount() As String Implements ICSVItem.Amount
        Get
            Return TransactionAmount
        End Get
        Set(ByVal value As String)
            While value.Substring(0, 1) = "$"
                value = value.Substring(1)
            End While
            TransactionAmount = value
        End Set
    End Property
End Class
