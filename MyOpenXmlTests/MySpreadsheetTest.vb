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
    <ExpectedException(GetType(ArgumentException))>
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

        Dim type = GetType(Product)
        Dim propertiesInfo = type.GetProperties()
        'Debug.Print("propertiesInfo: " & propertiesInfo.Count())

        For Each propertyInfo In propertiesInfo
            dataTable.Columns.Add(propertyInfo.Name)
        Next

        Dim productAsArray As Object() = Nothing

        Using db As New AdventureWorks2014()
            Dim result = db.Product.Take(100).ToList()
            'Dim result = db.Product.SqlQuery("SELECT TOP 100 * FROM Production.Product")
            For Each p As Product In result
                productAsArray = GetProductValuesAsArray(p)
                For Each obj In productAsArray
                    Debug.Print("GetProductAsArray(p): " & _
                                If(obj IsNot Nothing, obj.ToString(), "NULL"))
                Next
                dataTable.Rows.Add(productAsArray)
            Next
        End Using

        Dim filePath = "c:\Users\Public\Documents\WriteDatabaseDataToSpreadsheet.xlsx"
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
        spreadSheet.Write(dataTable)
    End Sub

    <TestMethod()>
    Public Sub WriteDatabaseTableColumnsToSpreadsheet()
        Dim dataTable = New DataTable("Products")

        Dim type = GetType(Product)
        Dim propertiesInfo = type.GetProperties()
        Debug.Print("propertiesInfo: " & propertiesInfo.Count())

        For Each propertyInfo In propertiesInfo
            dataTable.Columns.Add(propertyInfo.Name)
        Next

        Dim filePath = "c:\Users\Public\Documents\WriteDataBaseTableColumnsToSpreadsheet.xlsx"
        Dim spreadSheet = New MySpreadsheet()
        spreadSheet.Create(filePath)
        spreadSheet.Write(dataTable)
    End Sub

    <TestMethod()>
    Public Sub GetColumnNameFromColumnNumber()
        Dim spreadsheet = New MySpreadsheet()
        Dim columnNumbers = {0, 1, 5, 27, 127}
        Dim columnName = String.Empty
        Dim message = String.Empty
        For Each number As Integer In columnNumbers
            columnName = spreadsheet.GetColumnName(number)
            message = "Column number {0} equals column name {1}"
            message = String.Format(message, number, columnName)
            Debug.Print(message)
        Next
        ' REFAC usar Assert.Equal para comparar con nombres de columna esperados.
    End Sub

    <TestMethod()>
    Public Sub GetColumnNumberFromColumnName()
        Dim spreadsheet = New MySpreadsheet()
        Dim columnNames = {"", "A", "E", "AA", "DW"}
        Dim columnNumber = 0
        Dim message = String.Empty
        For Each name As String In columnNames
            columnNumber = spreadsheet.GetColumnNumber(name)
            message = "Column name {0} equals column number {1}"
            message = String.Format(message, name, columnNumber)
            Debug.Print(message)
        Next
        ' REFAC usar Assert.Equal para comparar con los numeros de columna esperados.

    End Sub

    ' Probar métodos anteriores con cargas intensas como 25*100 (25 columnas * 100 filas)
    ' que sería el número de veces que sería llamado, esto con el objetivo de determinar
    ' cuanto peso tiene este método [GetColumnNameFromColumnNumber()] 
    ' en el tiempo que demora la exportación.

    ' Helper Methods

    Private Function GetProductPropertiesAsArray(product As Product)





        Return 0
    End Function

    Private Function GetProductValuesAsArray(product As Product) As Object()
        Dim type = product.GetType()
        Dim propertiesInfo = type.GetProperties()
        Dim array As Object() = New Object(propertiesInfo.GetUpperBound(0)) {}

        For index = 0 To array.GetUpperBound(0)
            array(index) = propertiesInfo(index).GetValue(product)
        Next

        Return array
    End Function

    Private Function GetProductAsKeyValuePairArray(product As Product) As Object()
        Dim type = product.GetType()
        Dim propertiesInfo = type.GetProperties()
        Dim array As Object() = New Object(propertiesInfo.GetUpperBound(0)) {}
        Dim keyValuePair = String.Empty

        For index = 0 To array.GetUpperBound(0)
            keyValuePair = "{0} -> {1}"
            array(index) = String.Format(keyValuePair, propertiesInfo(index).Name,
                                         propertiesInfo(index).GetValue(product))
        Next

        Return array
    End Function

    Private Function GetCustomObjectPropertiesValuesAsArray(customObject As Object)
        Dim type = GetType(Object)
        Dim array As Object() = New Object(10) {}


        Return 0
    End Function

End Class
