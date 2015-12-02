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
    Public Sub DontCreateIfFilePathIsInvalid()
        Dim filePath = "c:\\\Users\Public\Documents\Libro2.xlsx"
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
        Assert.IsTrue(File.Exists(filePath))
        'Dim isValidFilePath = Directory.Exists(Path.GetDirectoryName(filePath))

    End Sub

    <TestMethod()>
    Public Sub DontCreateIfFileExtensionIsInvalid()
        Dim filePath = "c:\\Users\Public\Documents\Libro3"
        Dim fileName = Path.GetFileName(filePath)
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
        Assert.IsFalse(File.Exists(filePath), "fileName is:" & fileName)
    End Sub

    <TestMethod()>
    <ExpectedException(GetType(DirectoryNotFoundException))>
    Public Sub InvalidFileNameShouldThrowException()
        Dim filePath = "c:\\Users\Public\Documents\"
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
    End Sub

End Class
