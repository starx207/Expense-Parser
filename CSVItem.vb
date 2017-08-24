Public Class CSVItem
    Private TransactionDate As String
    Public Property TransDate() As String
        Get
            Return TransactionDate
        End Get
        Set(ByVal value As String)
            TransactionDate = value
        End Set
    End Property

    Private PayeeName As String
    Public Property Payee() As String
        Get
            Return PayeeName
        End Get
        Set(ByVal value As String)
            PayeeName = value
        End Set
    End Property

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
