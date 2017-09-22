Imports System.Xml

Public Class SettingsManager
    Private Const RootNode As String = "MonthlyBudget"
    Private Const TotalBudgetNode As String = "TotalBudget"
    Private Const CategoryNode As String = "Category"
    Private Const NameNode As String = "Name"
    Private Const PayeesNode As String = "Payees"
    Private Const TypeNode As String = "Type"
    Private Const BudgetNode As String = "Budget"
    Private ReadOnly AccessTotalBudgetNode As String = RootNode + "/" + TotalBudgetNode
    Private ReadOnly AccessCategoryNode As String = RootNode + "/" + CategoryNode
    Private ReadOnly AccessPayeeNameNodes As String = PayeesNode + "/" + NameNode

    Public Function SaveSettingsFile(ByRef budget As Budget) As Boolean
        Try
            Dim settings As New XmlWriterSettings
            settings.Indent = True

            Using writer As XmlWriter = XmlWriter.Create(My.Settings.BudgetSettingsFile, settings)
                writer.WriteStartDocument()
                writer.WriteStartElement(RootNode)
                writer.WriteElementString(TotalBudgetNode, budget.TotalBudget)

                For Each category As BudgetCategory In budget.Categories
                    writer.WriteStartElement(CategoryNode)
                    writer.WriteElementString(NameNode, category.Name)
                    writer.WriteElementString(TypeNode, category.Type)
                    writer.WriteElementString(BudgetNode, category.Budget.ToString())
                    writer.WriteStartElement(PayeesNode)
                    For Each payeeName As String In category.Payees
                        writer.WriteElementString(NameNode, payeeName)
                    Next
                    writer.WriteEndElement()
                    writer.WriteEndElement()
                Next

                writer.WriteEndElement()
                writer.WriteEndDocument()
            End Using

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function LoadNewSettingsFile(ByVal fileName As String) As XmlDocument
        Dim settingsFile As New XmlDocument
        Try
            settingsFile.Load(fileName)
            Return settingsFile
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function LoadSettings(ByVal settingsFile As XmlDocument) As Budget
        Dim budget As New Budget
        Dim newCategory As BudgetCategory
        If settingsFile.SelectSingleNode(AccessTotalBudgetNode) IsNot Nothing Then
            budget.TotalBudget = CDbl(settingsFile.SelectSingleNode(AccessTotalBudgetNode).InnerText)
        Else
            budget.TotalBudget = 0.00
        End If

        For Each budgetItem As XmlNode In settingsFile.SelectNodes(AccessCategoryNode)
            newCategory = New BudgetCategory
            If budgetItem(NameNode) IsNot Nothing Then
                newCategory.Name = budgetItem(NameNode).InnerText
            End If
            If budgetItem(TypeNode) IsNot Nothing Then
                newCategory.Type = DirectCast(CInt(budgetItem(TypeNode).InnerText), BudgetTypes)
            End If
            If budgetItem(BudgetNode) IsNot Nothing Then
                newCategory.Budget = CDbl(budgetItem(BudgetNode).InnerText)
            End If
            If budgetItem(PayeesNode) IsNot Nothing Then
                For Each payee As XmlNode In budgetItem.SelectNodes(AccessPayeeNameNodes)
                    newCategory.Payees.Add(payee.InnerText)
                Next
                newCategory.Payees.Sort()
            End If

            budget.Categories.Add(newCategory)
        Next

        Return budget
    End Function
End Class
