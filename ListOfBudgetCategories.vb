Public Class ListOfBudgetCategories
    Private budgetList As List(Of BudgetCategory)
    Public Property GenericList() As List(Of BudgetCategory)
        Get
            Return budgetList
        End Get
        Set(ByVal value As List(Of BudgetCategory))
            budgetList = value
        End Set
    End Property

    Sub New()
        budgetList = New List(Of BudgetCategory)
    End Sub

    Public Function PayeeCategory(ByVal payeeName As String) As String
        For Each category As BudgetCategory In budgetList
            If category.Payees.Contains(payeeName) Then
                Return category.Name
            End If
        Next
        ' Add the payee to Misc
        GetCategoryByName("Misc").Payees.Add(payeeName)

        Return "Misc"
    End Function

    Public Function GetAssignedCategory(ByVal payeeName As String) As String
        For Each category As BudgetCategory In budgetList
            If category.Payees.Contains(payeeName) Then
                Return category.Name
            End If
        Next

        Return "Unassigned"
    End Function

    Public Function AllCategoryNames() As List(Of String)
        Dim categoryNames As New List(Of String)
        For Each category As BudgetCategory In budgetList
            categoryNames.Add(category.Name)
        Next
        categoryNames.Sort()
        Return categoryNames
    End Function

    Public Function GetCategoryByName(ByVal categoryName As String) As BudgetCategory
        For Each category As BudgetCategory In budgetList
            If category.Name = categoryName Then
                Return category
            End If
        Next
        ' Add the category and rerun function
        budgetList.Add(New BudgetCategory(categoryName, "E", 0))
        Return GetCategoryByName(categoryName)
    End Function

    Public Function ReassignPayee(ByVal payeeName As String, ByVal newCategory As String) As Boolean
        Dim currentCategory As String = PayeeCategory(payeeName)
        If currentCategory <> newCategory Then
            Dim preCount As Integer = budgetList.Count
            GetCategoryByName(currentCategory).Payees.Remove(payeeName)
            GetCategoryByName(newCategory).Payees.Add(payeeName)

            Return preCount <> budgetList.Count
        End If
        Return False
    End Function

    Public Sub AssignCategoryUsedStatus(ByRef payees As CSVList)
        Dim stillUsed As Boolean
        For Each category As BudgetCategory In budgetList
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
            ReassignPayee(payeeName, "Misc")
        Next

        budgetList.Remove(GetCategoryByName(categoryName))
    End Sub
End Class
