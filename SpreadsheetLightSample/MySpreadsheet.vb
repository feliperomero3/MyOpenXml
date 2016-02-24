Imports SpreadsheetLight

Public Class MySpreadsheet
    Public Shared Sub ExportDatatableToExcel(filePath As String, datatable As DataTable)
        Dim document = New SLDocument()
        document.ImportDataTable("A1", datatable, True)
        document.SaveAs(filePath)
    End Sub

    Public Function ImportExcelToDataTable(dataTable As DataTable)
        Dim document = New SLDocument()
        Dim defaultView = dataTable.DefaultView
        ' TODO document. 

        Return Nothing
    End Function
End Class
