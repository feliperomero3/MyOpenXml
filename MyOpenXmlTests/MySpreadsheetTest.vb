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

    <TestMethod()>
    <ExpectedException(GetType(DirectoryNotFoundException))>
    Public Sub InvalidFilePathShouldThrowException()
        Dim filePath = "Users\Public\Documents\Libro2.xlsx"
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
    End Sub

    <TestMethod()>
    Public Sub DontCreateIfFileExtensionIsInvalid()
        Dim filePath = "c:\Users\Public\Documents\Libro3"
        Dim fileName = Path.GetFileName(filePath)
        Dim spreadSheet = New MySpreadsheet()
        Try
            spreadSheet.Create(filePath)
        Catch ex As ArgumentException
            Assert.IsFalse(File.Exists(filePath), "fileName: " & fileName)
            Return
        End Try

        Assert.Fail("No exception was thrown")
    End Sub

    <TestMethod()>
    <ExpectedException(GetType(DirectoryNotFoundException))>
    Public Sub InvalidFileNameShouldThrowException()
        Dim filePath = "c:\Users\Public\Documents\"
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
    End Sub

    Public Sub EmptyTextInputShouldThrowArgumentException()

    End Sub

    <TestMethod()>
    Public Sub WriteTextToSpreadsheet()
        Dim filePath = "c:\Users\Public\Documents\WriteTextToSpreadsheet.xlsx"
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
        spreadSheet.Write("SampleText")

    End Sub

End Class
