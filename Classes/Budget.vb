Public Class Budget
    Public Property Categories As List(Of BudgetCategory)
    Public Property TotalBudget As Double

    Sub New()
        Categories = New List(Of BudgetCategory)
        TotalBudget = 0.00
    End Sub

    Public Function PayeeCategory(ByVal payeeName As String) As String
        For Each category As BudgetCategory In Categories
            If category.Payees.Contains(payeeName) Then
                Return category.Name
            End If
        Next
        ' Add the payee to UnassignedPayeeType
        GetCategoryByName(UnassignedPayeeType).Payees.Add(payeeName)

        Return UnassignedPayeeType
    End Function

    Public Function GetAssignedCategory(ByVal payeeName As String) As String
        For Each category As BudgetCategory In Categories
            If category.Payees.Contains(payeeName) Then
                Return category.Name
            End If
        Next

        Return UnassignedCategory
    End Function

    Public Function AllCategoryNames() As List(Of String)
        Dim categoryNames As New List(Of String)
        For Each category As BudgetCategory In Categories
            categoryNames.Add(category.Name)
        Next
        categoryNames.Sort()
        Return categoryNames
    End Function

    Public Function GetCategoryByName(ByVal categoryName As String) As BudgetCategory
        For Each category As BudgetCategory In Categories
            If category.Name = categoryName Then
                Return category
            End If
        Next
        ' Add the category and rerun function
        Categories.Add(New BudgetCategory(categoryName, BudgetTypes.Expense, 0))
        Return GetCategoryByName(categoryName)
    End Function

    Public Function ReassignPayee(ByVal payeeName As String, ByVal newCategory As String) As Boolean
        Dim currentCategory As String = PayeeCategory(payeeName)
        If currentCategory <> newCategory Then
            Dim preCount As Integer = Categories.Count
            GetCategoryByName(currentCategory).Payees.Remove(payeeName)
            GetCategoryByName(newCategory).Payees.Add(payeeName)

            Return preCount <> Categories.Count
        End If
        Return False
    End Function

    Public Sub AssignCategoryUsedStatus(ByRef payees As CSVList)
        Dim stillUsed As Boolean
        For Each category As BudgetCategory In Categories
            stillUsed = False
            For Each payee As String In category.Payees
                If payees.IndexOf(payee) <> -1 Then
                    stillUsed = True
                End If
            Next
            category.Used = stillUsed
        Next
    End Sub

    Public Sub RemoveCategory(ByVal categoryName As String)
        Dim temp As New List(Of String)
        For Each payeeName As String In GetCategoryByName(categoryName).Payees
            temp.Add(payeeName)
        Next

        For Each payeeName As String In temp
            ReassignPayee(payeeName, UnassignedPayeeType)
        Next

        Categories.Remove(GetCategoryByName(categoryName))
    End Sub
End Class
