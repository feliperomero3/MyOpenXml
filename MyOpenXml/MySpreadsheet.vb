Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet

' https://msdn.microsoft.com/EN-US/library/office/ff478153.aspx?cs-save-lang=1&cs-lang=vb

Public Class MySpreadsheet
    Private _dataTable As DataTable
    Private _spreadSheetDocument As SpreadsheetDocument
    Private _workbookPart As WorkbookPart
    Private _worksheetPart As WorksheetPart
    Private _sheets As Sheets
    Private _sheet As Sheet

    Public Property FilePath As String

    Public Sub New()

    End Sub

    Private Sub Create()
        ' Create a spreadsheet document by supplying the filepath.
        ' By default, AutoSave = true, Editable = true, and Type = xlsx.
        _spreadSheetDocument = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook)

        ' Add a WorkbookPart to the document.
        _workbookPart = _spreadSheetDocument.AddWorkbookPart()

        ' Initialize Workbook
        _workbookPart.Workbook = New Workbook()

        ' Add a WorksheetPart to the WorkbookPart.
        _worksheetPart = _workbookPart.AddNewPart(Of WorksheetPart)()

        ' Initialize Worksheet
        _worksheetPart.Worksheet = New Worksheet(New SheetData())

        ' Add Sheets to the Workbook.
        _sheets = _spreadSheetDocument.WorkbookPart.Workbook.AppendChild(Of Sheets)(New Sheets())

        ' Append a new worksheet and associate it with the workbook.
        _sheet = New Sheet()
        _sheet.Id = _spreadSheetDocument.WorkbookPart.GetIdOfPart(_worksheetPart)
        _sheet.SheetId = 1
        _sheet.Name = "Hoja1"

        _sheets.Append(_sheet)

        _workbookPart.Workbook.Save()

        _spreadSheetDocument.Close()
    End Sub

    Public Sub Create(newFilePath As String)
        If Not String.IsNullOrWhiteSpace(newFilePath) Then
            FilePath = newFilePath
            Create()
        End If
    End Sub
End Class
