Imports SpreadsheetLight

Public Class MySpreadsheet
    Public Shared Sub ExportDatatableToExcel(filePath As String, datatable As DataTable)
        Dim document = New SLDocument()
        document.ImportDataTable("A1", datatable, True)
        document.SaveAs(filePath)
    End Sub
End Class
