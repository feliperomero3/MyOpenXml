Imports MyOpenXml
Imports System.IO
Imports Entities

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

    <TestMethod()>
    Public Sub WriteDataTableToSpreadsheet()
        Dim dataTable = New DataTable("SampleDataTable")
        dataTable.Columns.Add("Id")
        dataTable.Columns.Add("Artist")
        dataTable.Columns.Add("Album")
        dataTable.Columns.Add("Genre")
        dataTable.Columns.Add("Year")
        dataTable.Rows.Add(New [Object]() {"1", "Metallica", "Kill 'Em All", "Metal", "1983"})
        dataTable.Rows.Add(New [Object]() {"2", "Finch", "What It Is to Burn", "Alternative", "2002"})
        dataTable.Rows.Add(New [Object]() {"3", "A Perfect Circle", "Thirteenth Step", "Alternative", "2003"})
        dataTable.Rows.Add(New [Object]() {"4", "Coldplay", "Parachutes", "Alternative", "2000"})
        dataTable.Rows.Add(New [Object]() {"5", "Héroes del Silencio", "Avalancha", "Alternativo & Rock Latino", "1995"})
        dataTable.Rows.Add(New [Object]() {"6", "British India", "Controller", "Alternative", "2013"})
        Dim filePath = "c:\Users\Public\Documents\WriteDataTableToSpreadsheet.xlsx"
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
        spreadSheet.Write(dataTable)
    End Sub

    <TestMethod()>
    Public Sub WriteDatabaseDataToSpreadsheet()
        Dim dataTable = New DataTable("Products")
        Using db As New AdventureWorks2014()
            Dim result = db.Product.AsQueryable()
            For Each p As Product In result
                dataTable.Rows.Add(p)
            Next
        End Using
        Dim filePath = "c:\Users\Public\Documents\WriteDatabaseDataToSpreadsheet.xlsx"
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
        spreadSheet.Write(dataTable)
    End Sub

End Class
