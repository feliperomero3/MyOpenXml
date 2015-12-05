Imports SpreadsheetLight

Public Class SpreadsheetLightSample
    Public Shared Sub ExportDatatableToExcel(datatable As DataTable)
        Dim document = New SLDocument()
        document.ImportDataTable("A1", datatable, True)
        document.SaveAs("c:\Users\Public\Documents\WriteDatabaseDataToSpreadsheetSL.xlsx")
    End Sub
End Class
