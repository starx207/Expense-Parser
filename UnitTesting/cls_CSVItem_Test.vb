Imports System.Text
Imports Expense_Parser
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class cls_CSVItem_Test

    <TestMethod()>
    Public Sub CSVItem_ClassExists()
        Dim csvItem As Object
        Try
            csvItem = New CSVItem
        Catch ex As Exception
            Assert.Fail("Class ""CSVItem"" not implemented")
        End Try

        ' Test passes if it gets here
        Assert.IsTrue(True)
    End Sub

    <TestMethod>
    Public Sub CSVItem_TransDatePropertyReadWrite()
        Dim testItem As New CSVItem
        Dim input As String = "12/12/2017"
        Dim expectedOutput As String = input

        testItem.TransDate = input

        Assert.AreEqual(expectedOutput, testItem.TransDate) ', $"""TransDate"" Property did not save value ""{input}"" correctly")
    End Sub

    <TestMethod>
    Public Sub CSVItem_PayeePropertyReadWrite()
        Dim testItem As New CSVItem
        Dim input As String = "Payee1"
        Dim expectedOutput As String = input

        testItem.Payee = input

        Assert.AreEqual(expectedOutput, testItem.Payee) ', $"""Payee"" Property did not save the value ""{input}"" correctly")
    End Sub

    <DataTestMethod, DataRow("20.00", "20.00"), DataRow("$30.00", "30.00")>
    Public Sub CSVItem_AmountPropertyReadWrite(ByVal input As String, ByVal expectedOutput As String)
        Dim testItem As New CSVItem

        testItem.Amount = input

        Assert.AreEqual(expectedOutput, testItem.Amount) ', $"""Amount"" Property did not save the value ""{input}"" correctly")
    End Sub
End Class