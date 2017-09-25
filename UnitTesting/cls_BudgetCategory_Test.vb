Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Expense_Parser

<TestClass()> Public Class cls_BudgetCategory_Test

    <TestMethod()>
    Public Sub BudgetCategory_ClassExists()
        Dim category As Object
        Try
            category = New BudgetCategory
        Catch ex As Exception
            Assert.Fail("Class ""BudgetCategory"" not implemented")
        End Try

        ' Test passes if it gets here
        Assert.IsTrue(True)
    End Sub

    <TestMethod>
    Public Sub BudgetCategory_NamePropertyReadWrite()
        Dim category As New BudgetCategory
        Dim input As String = "CategoryName"
        Dim expectedOutput As String = input

        category.Name = input

        Assert.AreEqual(expectedOutput, category.Name, $"""Name"" Property did not save value ""{input}"" correctly")
    End Sub

    <TestMethod>
    Public Sub BudgetCategory_TypePropertyReadWrite()
        Dim category As New BudgetCategory
        Dim input As BudgetTypes = BudgetTypes.Expense
        Dim expectedOutput As BudgetTypes = input

        category.Type = input

        Assert.AreEqual(expectedOutput, category.Type, $"""Type"" Property did not save value ""{input}"" correctly")
    End Sub

    <TestMethod>
    Public Sub BudgetCategory_BudgetPropertyReadWrite()
        Dim category As New BudgetCategory
        Dim input As Double = 100.5
        Dim expectedOutput As Double = input

        category.Budget = input

        Assert.AreEqual(expectedOutput, category.Budget, $"""Name"" Property did not save value ""{input}"" correctly")
    End Sub

    <TestMethod>
    Public Sub BudgetCategory_PayeesPropertyReadWrite()
        Dim category As New BudgetCategory
        Dim input As New List(Of String) From {"Payee1", "Payee2", "Payee3", "Payee4"}
        Dim expectedOutput As List(Of String) = input

        category.Payees = input

        Assert.AreSame(expectedOutput, category.Payees, """Payees"" Property did not save list correctly")
    End Sub

    <DataTestMethod,
        DataRow(0, True),
        DataRow(100, True),
        DataRow(-100, True),
        DataRow(0, False),
        DataRow(100, False),
        DataRow(-100, False)>
    Public Sub BudgetCategory_UsedPropertyReadWrite(ByVal budgetValue As Double, ByVal input As Boolean)
        Dim category As New BudgetCategory
        category.Budget = budgetValue

        Dim budgetGreaterThanZero As Boolean = budgetValue > 0
        Dim expectedOutput As Boolean = budgetGreaterThanZero Or input

        category.Used = input

        Assert.AreEqual(expectedOutput, category.Used, $"""Used"" returns wrong value for a budget of {budgetValue} and an input value of {input}")
    End Sub

    <TestMethod>
    Public Sub BudgetCategory_OverloadedConstructorAccepts_Name_Type_Amount()
        Dim name As String = "Name1"
        Dim type As BudgetTypes = BudgetTypes.Income
        Dim amount As Double = 250.25

        Dim category As New BudgetCategory(name, type, amount)

        Assert.AreEqual(name, category.Name)
        Assert.AreEqual(type, category.Type)
        Assert.AreEqual(amount, category.Budget)
    End Sub

End Class