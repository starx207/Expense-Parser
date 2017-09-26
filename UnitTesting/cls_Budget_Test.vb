Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Expense_Parser

<TestClass()> Public Class cls_Budget_Test

    <TestMethod()>
    Public Sub Budget_ClassExists()
        Dim testBudget As Object
        Try
            testBudget = New Budget
        Catch ex As Exception
            Assert.Fail("Class ""Budget"" not implemented")
        End Try

        ' Test passes if it gets here
        Assert.IsTrue(True)
    End Sub

    <TestMethod>
    Public Sub Budget_CategoriesProperty_ReadWrite()
        Dim input As New List(Of IBudgetCategory) From {
            New mock_BudgetCategory(),
            New mock_BudgetCategory(),
            New mock_BudgetCategory()
        }

        Dim testBudget As New Budget

        testBudget.Categories = input

        Assert.AreSame(input, testBudget.Categories, "Input list not the same as saved list")
    End Sub

    <DataTestMethod, DataRow(2000.0), DataRow(0.0), DataRow(-100.23), DataRow(1423.5)>
    Public Sub Budget_TotalBudgetProperty_ReadWrite(ByVal input As Double)
        Dim testBudget As New Budget

        testBudget.TotalBudget = input

        Assert.AreEqual(input, testBudget.TotalBudget, $"""Budget"" class did not save value of {input} correctly")
    End Sub

    <TestMethod>
    Public Sub Budget_GetCategoryByNameMethod_ReturnsCorrectBudgetCategory()
        Dim categoryName As String = "Category1"
        Dim category As New mock_BudgetCategory With {.Name = categoryName}
        Dim testBudget As New Budget
        testBudget.Categories.Add(category)
        testBudget.Categories.Add(New mock_BudgetCategory With {.Name = "Not this one"})
        testBudget.Categories.Add(New mock_BudgetCategory With {.Name = "Wrong one"})
        Dim output As mock_BudgetCategory

        output = testBudget.GetCategoryByName(categoryName)

        Assert.AreSame(category, output, "The wrong BudgetCategory was returned")
    End Sub

    <TestMethod>
    Public Sub Budget_GetCategoryByNameMethod_AddsNonExistingCategoryAndReturns()
        Dim categoryName As String = "Category That doesn't yet exist"
        Dim testBudget As New Budget
        testBudget.Categories.Add(New mock_BudgetCategory With {.Name = "Not this one"})
        testBudget.Categories.Add(New mock_BudgetCategory With {.Name = "Wrong one"})
        Dim output As mock_BudgetCategory

        output = mock_BudgetCategory.ConvertFromIBudgetCategory(testBudget.GetCategoryByName(categoryName))

        Assert.AreEqual(categoryName, output.Name, "The BudgetCategory returned has the wrong name")
    End Sub

    <DataTestMethod, DataRow("Payee1", "Category1"), DataRow("Payee2", "Category1"), DataRow("Payee3", "")>
    Public Sub Budget_GetAssignedCategoryMethod_ReturnsCategoryNameThatPayeeIsAssignedTo(ByVal payee As String, ByVal category As String)
        Dim testBudget As New Budget
        Dim testCategory As New mock_BudgetCategory
        Dim output As String

        If category = "" Then
            category = Budget.UnassignedPayeeType
        End If

        If category <> Budget.UnassignedPayeeType Then
            testCategory.Name = category
            testCategory.Payees.Add(payee)
            testBudget.Categories.Add(testCategory)
        End If

        testBudget.Categories.Add(New mock_BudgetCategory With {.Name = "WrongCategory1"})
        testBudget.Categories.Add(New mock_BudgetCategory With {.Name = "WrongCategory2"})


        output = testBudget.GetAssignedCategory(payee)

        Assert.AreEqual(category, output, $"Payee ""{payee}"" should be assigned to ""{category}"", not ""{output}""")
    End Sub

    <DataTestMethod, DataRow("Payee1", "Category1"), DataRow("Payee3", "")>
    Public Sub Budget_FetchOrAddPayeeCategoryMethod_ReturnsCategoryPayeeIsAssignedTo_OrAddsPayeeToMiscCategoryIfNotYetAssigned(ByVal payee As String, ByVal category As String)
        Dim testBudget As New Budget
        Dim testCategory As New mock_BudgetCategory
        Dim output As String

        testBudget.Categories.Add(New mock_BudgetCategory With {.Name = "WrongCategory1"})
        testBudget.Categories.Add(New mock_BudgetCategory With {.Name = "WrongCategory2"})

        If category = "" Then
            category = Budget.UnassignedPayeeType
        End If

        If category = Budget.UnassignedPayeeType Then
            testCategory = mock_BudgetCategory.ConvertFromIBudgetCategory(testBudget.GetCategoryByName(category))

            Assert.IsFalse(testCategory.Payees.Contains(payee), $"Test results invalid because category ""{category}"" already contains payee ""{payee}""")

            output = testBudget.FetchOrAddPayeeCategory(payee)
            testCategory = mock_BudgetCategory.ConvertFromIBudgetCategory(testBudget.GetCategoryByName(category))

            Assert.AreEqual(Budget.UnassignedPayeeType, output, $"An unassigned payee should be assigned to ""{Budget.UnassignedPayeeType}""")
            Assert.IsTrue(testCategory.Payees.Contains(payee), $"Payee ""{payee}"" was not added to ""{category}"" category")
        Else
            testCategory.Name = category
            testCategory.Payees.Add(payee)
            testBudget.Categories.Add(testCategory)

            output = testBudget.GetAssignedCategory(payee)

            Assert.AreEqual(category, output, $"Payee ""{payee}"" should be assigned to ""{category}"", not ""{output}""")
        End If
    End Sub

    <TestMethod>
    Public Sub Budget_AllCategoryNamesMethod_ReturnsAllCategoryNamesInSortedOrder()
        Dim unorderedCategories As New List(Of String) From {"D_Category", "A_Category", "C_Category", "B_Category"}
        Dim orderedCategories As List(Of String)
        Dim testBudget As New Budget
        Dim expectedValue As String
        Dim actualValue As String
        For Each name As String In unorderedCategories
            testBudget.Categories.Add(New mock_BudgetCategory() With {.Name = name})
        Next
        unorderedCategories.Sort()

        orderedCategories = testBudget.AllCategoryNames()

        Assert.AreEqual(unorderedCategories.Count, orderedCategories.Count, "Not all categories were returned")
        For i As Integer = 0 To orderedCategories.Count - 1
            expectedValue = unorderedCategories.Item(i)
            actualValue = orderedCategories.Item(i)

            Assert.AreEqual(expectedValue, actualValue, $"Categories are in the wrong order. Position {i} should have value of ""{expectedValue}"" instead of ""{actualValue}""")
        Next
    End Sub

    <TestMethod>
    Public Sub Budget_ReassignPayeeMethod_MovesAPayeeFromOneCategoryToAnother()
        Dim payeeToMove As String = "Test Payee"
        Dim currentCategory As String = "Category1"
        Dim newCategory As String = "Category2"
        Dim testBudget As New Budget
        testBudget.Categories = New List(Of IBudgetCategory) From {
            New mock_BudgetCategory With {.Name = currentCategory, .Payees = New List(Of String) From {payeeToMove, "Payee1"}},
            New mock_BudgetCategory With {.Name = newCategory, .Payees = New List(Of String) From {"Payee2", "Payee3", "Payee4"}}
        }

        testBudget.ReassignPayee(payeeToMove, newCategory)

        Assert.IsFalse(testBudget.GetCategoryByName(currentCategory).Payees.Contains(payeeToMove), $"""{payeeToMove}"" is still assigned to old category")
        Assert.IsTrue(testBudget.GetCategoryByName(newCategory).Payees.Contains(payeeToMove), $"""{payeeToMove}"" was not assigned to new category")
    End Sub

    <TestMethod>
    Public Sub Budget_ReassignPayeeMethod_ReturnsBooleanToIndicateIfACategoryWasAddedAsAResultOfTheMove()
        Dim payeeToMove As String = "Test Payee"
        Dim currentCategory As String = "Category1"
        Dim newCategory As String = "Category2"
        Dim testBudget As New Budget
        Dim categoryAdded As Boolean
        testBudget.Categories = New List(Of IBudgetCategory) From {
            New mock_BudgetCategory With {.Name = currentCategory, .Payees = New List(Of String) From {payeeToMove, "Payee1"}}
        }

        categoryAdded = testBudget.ReassignPayee(payeeToMove, newCategory)

        Assert.IsTrue(categoryAdded, "ReassignPayee() should return ""True"" when a category is added as a result of payee reassignment")
    End Sub

    <TestMethod>
    Public Sub Budget_AssignCategoryUsedStatusMethod_SetUnusedCategoriesUsedPropertyToFalseIfNoPayeesInCSVListInTheCategory()
        Dim list As New mock_CSVList
        list.GenericList.Add(New mock_CSVItem With {.Payee = "Payee1"})
        list.GenericList.Add(New mock_CSVItem With {.Payee = "Payee2"})

        Dim testBudget As New Budget
        testBudget.Categories = New List(Of IBudgetCategory) From {
            New mock_BudgetCategory With {.Name = "Category1", .Payees = New List(Of String) From {"Payee2", "Payee3", "Payee4"}},
            New mock_BudgetCategory With {.Name = "Category2", .Payees = New List(Of String) From {"Payee1"}},
            New mock_BudgetCategory With {.Name = "Category3", .Payees = New List(Of String) From {"Payee5", "Payee6"}}
        }

        testBudget.AssignCategoryUsedStatus(list)

        For Each category As IBudgetCategory In testBudget.Categories
            If category.Name = "Category3" Then
                Assert.IsFalse(category.Used, "Category3 should not be marked as used")
            Else
                Assert.IsTrue(category.Used, $"{category.Name} should be marked as used")
            End If
        Next
    End Sub

    <TestMethod>
    Public Sub Budget_RemoveCategoryMethod_RemovesCategoryAndReassignsAllPayeesToMisc()
        Dim categoryToRemove As String = "Category1"
        Dim payeesInCategory As New List(Of String) From {"Payee1", "Payee2", "Payee3"}
        Dim testBudget As New Budget
        testBudget.Categories.Add(
            New mock_BudgetCategory With {.Name = categoryToRemove, .Payees = payeesInCategory}
        )

        testBudget.RemoveCategory(categoryToRemove)

        Dim categoryExists As Boolean = False
        For Each category As IBudgetCategory In testBudget.Categories
            If category.Name = categoryToRemove Then
                categoryExists = True
            End If
        Next

        Assert.IsFalse(categoryExists, "Category was not removed")

        For Each payee As String In payeesInCategory
            Assert.IsTrue(testBudget.GetCategoryByName(Budget.UnassignedPayeeType).Payees.Contains(payee), $"""{payee}"" was not added to ""{Budget.UnassignedPayeeType}""")
        Next
    End Sub

End Class