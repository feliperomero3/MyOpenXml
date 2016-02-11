Imports SpreadsheetLightSample
Imports Entities

<TestClass()>
Public Class SpreadsheetLightSampleTest
    <TestMethod()>
    Public Sub ExportProductsTableToExcelTest()
        Dim filePath = "c:\Users\Public\Documents\ProductsTableToExcelTest.xlsx"
        Dim dataTable = New DataTable("Products")

        Dim type = GetType(Product)
        Dim propertiesInfo = type.GetProperties()
        'Debug.Print("propertiesInfo: " & propertiesInfo.Count())

        For Each propertyInfo In propertiesInfo
            dataTable.Columns.Add(propertyInfo.Name)
        Next

        Dim productAsArray As Object() = Nothing

        Using db As New AdventureWorks2014()
            Dim result = db.Product.ToList()
            For Each p As Product In result
                productAsArray = GetProductValuesAsArray(p)
                'For Each obj In productAsArray
                '    Debug.Print("GetProductAsArray(p): " & _
                '                If(obj IsNot Nothing, obj.ToString(), "NULL"))
                'Next
                dataTable.Rows.Add(productAsArray)
            Next
        End Using

        MySpreadsheet.ExportDatatableToExcel(filePath, dataTable)
    End Sub

    <TestMethod()>
    Public Sub ExportSalesOrderDetailTableToExcelTest()
        Dim filePath = "c:\Users\Public\Documents\SalesOrderDetailTableToExcelTest.xlsx"
        Dim dataTable = New DataTable("SalesOrderDetail")

        Dim type = GetType(SalesOrderDetail)
        Dim propertiesInfo = type.GetProperties().
            Where(Function(x) Not x.PropertyType.IsClass _
                              Or x.PropertyType.IsClass _
                              AndAlso x.PropertyType = GetType(String)).
            ToArray()

        For Each propertyInfo In propertiesInfo
            dataTable.Columns.Add(propertyInfo.Name)
        Next

        Dim productAsArray As Object() = Nothing

        Using db As New AdventureWorks2014()
            Dim result = db.SalesOrderDetail.Include("SalesOrderHeader").ToList()
            'Dim result = db.SalesOrderDetail.ToList()
            For Each s As SalesOrderDetail In result
                productAsArray = GetSalesOrderDetailsValuesAsArray(s)
                dataTable.Rows.Add(productAsArray)
            Next
        End Using

        MySpreadsheet.ExportDatatableToExcel(filePath, dataTable)
    End Sub

    Private Function GetProductValuesAsArray(product As Product) As Object()
        Dim type = product.GetType()
        Dim propertiesInfo = type.GetProperties()
        Dim array As Object() = New Object(propertiesInfo.GetUpperBound(0)) {}

        For index = 0 To array.GetUpperBound(0)
            array(index) = propertiesInfo(index).GetValue(product)
        Next

        Return array
    End Function

    Private Function GetSalesOrderDetailsValuesAsArray(order As SalesOrderDetail) As Object()
        Dim type = order.GetType()
        Dim propertiesInfo = type.GetProperties().
            Where(Function(x) Not x.PropertyType.IsClass _
                              Or x.PropertyType.IsClass _
                              AndAlso x.PropertyType = GetType(String)).
            ToArray()
        Dim array As Object() = New Object(propertiesInfo.GetUpperBound(0)) {}

        For index = 0 To array.GetUpperBound(0)
            array(index) = propertiesInfo(index).GetValue(order)
        Next

        Return array
    End Function
End Class
