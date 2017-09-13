Public Class CSVItem
    Public Property TransDate As String
    Public Property Payee As String

    Private TransactionAmount As String
    Public Property Amount() As String
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
