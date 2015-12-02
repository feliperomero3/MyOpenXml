Imports MyOpenXml
Imports System.IO

<TestClass()>
Public Class MySpreadsheetTest

    <TestMethod()>
    Public Sub CreateSpreadsheet()
        Dim filePath = "c:\Users\Public\Documents\Libro1.xlsx"
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
        Assert.IsTrue(File.Exists(filePath))
    End Sub

End Class
