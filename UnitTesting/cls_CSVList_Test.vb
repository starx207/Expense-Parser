Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Expense_Parser
Imports System.IO
Imports Microsoft.VisualBasic.FileIO

<TestClass()> Public Class cls_CSVList_Test

    Private Const TestResourcesFolder As String = "..\..\TestResources\"

    Sub New()
        If Not Directory.Exists(TestResourcesFolder) Then
            Directory.CreateDirectory(TestResourcesFolder)
        End If
    End Sub

    <TestMethod>
    Public Sub CSVList_ClassExists()
        Dim csvList As Object
        Try
            csvList = New CSVList
        Catch
            Assert.Fail("Class ""CSVList"" not implemented")
        End Try

        ' Test passes if it gets here
        Assert.IsTrue(True)
    End Sub

    <TestMethod>
    Public Sub CSVList_HasAGenericListProperty()
        Dim testList As New CSVList

        Dim input As New List(Of ICSVItem) From {New mock_CSVItem, New mock_CSVItem}

        testList.GenericList = input

        Assert.AreSame(input, testList.GenericList, "Object refernce not the same")
    End Sub

    <TestMethod>
    Public Sub CSVList_LengthPropertyReturnsGenericListLength()
        Dim testList As New CSVList

        testList.GenericList = New List(Of ICSVItem) From {New mock_CSVItem, New mock_CSVItem, New mock_CSVItem, New mock_CSVItem}

        Dim expectedCount As Integer = 4

        Assert.AreEqual(expectedCount, testList.Length, "The counts are not equal")
    End Sub

    <TestMethod>
    Public Sub CSVList_ItemPropertyReturnsCorrespondingItemInGenericList()
        Dim testList As New CSVList
        Dim input As New mock_CSVItem
        Dim index As Integer = 2

        testList.GenericList = New List(Of ICSVItem) From {New mock_CSVItem, New mock_CSVItem}

        testList.GenericList.Add(input)

        Assert.AreSame(input, testList.Item(index), "Object refernce not the same")
    End Sub

    <TestMethod>
    Public Sub CSVList_AddMethodAcceptsExplicitValuesForNewCSVItem()
        Dim addDate As String = "12/12/2017"
        Dim addName As String = "Payee Name"
        Dim addAmt As String = "$20.00"

        Dim testList As New CSVList

        testList.Add(addDate, addName, addAmt)

        Dim output As ICSVItem = testList.Item(0)

        Assert.AreEqual(addDate, output.TransDate, "Dates are not equal")
        Assert.AreEqual(addName, output.Payee, "Payee Names are not equal")
        Assert.AreEqual(addAmt.Substring(1), output.Amount, "Amounts are not equal")
    End Sub

    <TestMethod>
    Public Sub CSVList_AddMethodAcceptsACSVItemObject()
        Dim input As New mock_CSVItem
        input.TransDate = "12/12/2017"
        input.Payee = "Payee Name"
        input.Amount = "20.00"

        Dim testList As New CSVList

        testList.Add(input)

        Dim output As ICSVItem = testList.Item(0)

        Assert.AreEqual(input.Amount, output.Amount, "Amounts are not equal")
        Assert.AreEqual(input.Payee, output.Payee, "Payee Names are not equal")
        Assert.AreEqual(input.TransDate, output.TransDate, "Dates are not equal")
    End Sub

    <DataTestMethod, DataRow("Fred"), DataRow("Payee1"), DataRow("Payee3")>
    Public Sub CSVList_IndexOfMethodReturnsCorrectIndexOfCSVItemWithSpecifiedPayeeName(ByVal searchFor As String)
        Dim testList As New CSVList
        testList.Add("12/12/2017", "Payee1", "20.00")
        testList.Add("12/12/2017", "Payee2", "30.00")
        testList.Add("12/12/2017", "Payee3", "40.00")

        Dim expectedIndex As Integer
        Select Case searchFor
            Case "Payee1"
                expectedIndex = 0
            Case "Payee2"
                expectedIndex = 1
            Case "Payee3"
                expectedIndex = 2
            Case Else
                expectedIndex = -1
        End Select


        Dim actualIndex As Integer = testList.IndexOf(searchFor)

        Assert.AreEqual(expectedIndex, actualIndex, "IndexOf() returned the wrong value")
    End Sub

    <DataTestMethod, DataRow(-1), DataRow(0), DataRow(1), DataRow(2)>
    Public Sub CSVList_RemoveAtMethodRemovesCorrectCSVItemAtIndex(ByVal removeAtIndex As Integer)
        Dim testList As New CSVList()
        testList.Add("12/12/2017", "Payee1", "20.00")
        testList.Add("12/12/2017", "Payee2", "30.00")
        testList.Add("12/12/2017", "Payee3", "40.00")

        Dim expectedLength As Integer = 2
        Dim expectedPayeeName As String
        Select Case removeAtIndex
            Case 0
                expectedPayeeName = "Payee2"
            Case 1
                expectedPayeeName = "Payee3"
            Case Else
                expectedPayeeName = "Unknown"
        End Select


        If removeAtIndex < 0 Then
            Assert.ThrowsException(Of ArgumentOutOfRangeException)(Sub() testList.RemoveAt(removeAtIndex), "Negative indexes should throw an exception")
        Else
            testList.RemoveAt(removeAtIndex)

            Assert.AreEqual(expectedLength, testList.Length)

            If removeAtIndex > 1 Then
                Assert.ThrowsException(Of ArgumentOutOfRangeException)(Sub() testList.Item(removeAtIndex).Payee.StartsWith("test"))
            Else
                Assert.AreEqual(expectedPayeeName, testList.Item(removeAtIndex).Payee)
            End If
        End If
    End Sub

    <TestMethod>
    Public Sub CSVList_UniqueItemsPropertyReturnsStringListWithoutDuplicatePayeeNames()
        Dim testList As New CSVList
        testList.Add("12/12/2017", "Payee1", "20.00")
        testList.Add("12/12/2017", "Payee2", "30.00")
        testList.Add("12/12/2017", "Payee3", "40.00")
        testList.Add("12/12/2017", "Payee1", "25.00")
        testList.Add("12/12/2017", "Payee2", "35.00")
        testList.Add("12/12/2017", "Payee2", "70.00")

        Dim expectedCount As Integer = 3

        Dim uniqueList As List(Of String) = testList.UniqueItems

        Dim payee1count As Integer = uniqueList.Where(Function(n) n = "Payee1").Count()
        Dim payee2count As Integer = uniqueList.Where(Function(n) n = "Payee2").Count()
        Dim payee3count As Integer = uniqueList.Where(Function(n) n = "Payee3").Count()

        Assert.AreEqual(expectedCount, uniqueList.Count, "Number of unique payees different than expected")
        Assert.AreEqual(1, payee1count, """Payee1"" was listed multiple times")
        Assert.AreEqual(1, payee2count, """Payee2"" was listed multiple times")
        Assert.AreEqual(1, payee3count, """Payee3"" was listed multiple times")
    End Sub

    <TestMethod>
    Public Sub CSVList_RemovingUniquePayeeAlsoRemoveFromUniqueItemProperty()
        Dim testList As New CSVList
        testList.Add("12/12/2017", "Payee1", "20.00")
        testList.Add("12/12/2017", "Payee2", "30.00")
        testList.Add("12/12/2017", "Payee3", "40.00")
        testList.Add("12/12/2017", "Payee1", "25.00")
        testList.Add("12/12/2017", "Payee2", "35.00")
        testList.Add("12/12/2017", "Payee2", "70.00")

        Dim uniquePayee As String = "Payee3"
        Dim uniquePayeeIndex As Integer = testList.IndexOf(uniquePayee)

        testList.RemoveAt(uniquePayeeIndex)

        Assert.IsFalse(testList.UniqueItems.Contains(uniquePayee), """Payee3"" not removed from UniqueItems list")
    End Sub

    <TestMethod>
    Public Sub CSVList_OverloadedConstructorAcceptsTextFieldParserToLoad()
        Dim fileContent As String = "Date,Payee Name,Amount" + vbCrLf +
            "12/12/2017,Payee1,20.00" + vbCrLf +
            "12/12/2017,Payee2,$30.00" + vbCrLf +
            "12/12/2017,Payee3,40.00" + vbCrLf +
            "12/12/2017,Payee4,49.00"
        Dim testFileName As String = "CSVImportTestFile.csv"
        Dim expectedLength As Integer = 4
        Dim checkIndex As Integer = 1
        Dim expectedDate As String = "12/12/2017"
        Dim expectedPayee As String = "Payee2"
        Dim expectedAmount As String = "30.00"

        File.WriteAllText(TestResourcesFolder + testFileName, fileContent)
        Dim txtParser As New TextFieldParser(TestResourcesFolder + testFileName)

        Dim testList As New CSVList(txtParser)

        txtParser.Close()

        Assert.AreEqual(expectedLength, testList.Length, "CSV lines loaded are different than expected")
        Assert.AreEqual(expectedDate, testList.Item(checkIndex).TransDate, "Dates not loaded correctly")
        Assert.AreEqual(expectedPayee, testList.Item(checkIndex).Payee, "Payees not loaded correctly")
        Assert.AreEqual(expectedAmount, testList.Item(checkIndex).Amount, "Amounts not loaded correctly")
    End Sub

    <TestMethod>
    Public Sub CSVList_OverloadedConstructorAcceptsTextFieldParserToLoad_AlternateHeaderNamesAndOrder()
        Dim fileContent As String = "merchant,TRANSACTION date,   billing amount" + vbCrLf +
            "Payee1,12/12/2017,20.00" + vbCrLf +
            "Payee2,12/12/2017,$30.00" + vbCrLf +
            "Payee3,12/12/2017,40.00" + vbCrLf +
            "Payee4,12/12/2017,49.00"
        Dim testFileName As String = "CSVImportTestFile.csv"
        Dim expectedLength As Integer = 4
        Dim checkIndex As Integer = 1
        Dim expectedDate As String = "12/12/2017"
        Dim expectedPayee As String = "Payee2"
        Dim expectedAmount As String = "30.00"

        File.WriteAllText(TestResourcesFolder + testFileName, fileContent)
        Dim txtParser As New TextFieldParser(TestResourcesFolder + testFileName)

        Dim testList As New CSVList(txtParser)

        txtParser.Close()

        Assert.AreEqual(expectedLength, testList.Length, "CSV lines loaded are different than expected")
        Assert.AreEqual(expectedDate, testList.Item(checkIndex).TransDate, "Dates not loaded correctly")
        Assert.AreEqual(expectedPayee, testList.Item(checkIndex).Payee, "Payees not loaded correctly")
        Assert.AreEqual(expectedAmount, testList.Item(checkIndex).Amount, "Amounts not loaded correctly")
    End Sub

End Class