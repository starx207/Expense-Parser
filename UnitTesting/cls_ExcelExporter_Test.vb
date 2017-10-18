Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports Expense_Parser
Imports System.Diagnostics.CodeAnalysis

<ExcludeFromCodeCoverage>
<TestClass()> Public Class cls_ExcelExporter_Test

    <TestMethod()>
    Public Sub ExcelExporter_ClassExists()
        Dim textExport As Object
        Try
            textExport = New ExcelExporter
        Catch ex As Exception
            Assert.Fail("Class ""ExcelExporter"" not implemented")
        End Try

        ' Test passes if it gets here
        Assert.IsTrue(True)
    End Sub

End Class