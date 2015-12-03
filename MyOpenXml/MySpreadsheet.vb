Imports System.Text.RegularExpressions
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Imports System.IO

' https://msdn.microsoft.com/EN-US/library/office/ff478153.aspx?cs-save-lang=1&cs-lang=vb

Public Class MySpreadsheet

    Private _dataTable As DataTable
    Private _spreadSheetDocument As SpreadsheetDocument
    Private _workbookPart As WorkbookPart
    Private _worksheetPart As WorksheetPart
    Private _sheets As Sheets
    Private _sheet As Sheet
    Private Const _sheetName As String = "Hoja1"

    Public Property FilePath As String

    Public Sub New()

    End Sub

    Public Sub Create(newFilePath As String)
        If Not String.IsNullOrWhiteSpace(newFilePath) Then
            Dim fileExtension = Path.GetExtension(newFilePath)
            Dim pattern = "^(.xls|.xlsx)$"
            If Regex.IsMatch(fileExtension, pattern) Then
                FilePath = newFilePath
                Create()
            Else
                Throw New ArgumentException("Invalid file extension: " & fileExtension)
            End If
        End If
    End Sub

    Private Sub Create()
        Try
            ' Create a spreadsheet document by supplying the filepath.
            ' By default, AutoSave = true, Editable = true, and Type = xlsx.
            _spreadSheetDocument = SpreadsheetDocument.Create(FilePath, SpreadsheetDocumentType.Workbook)
        Catch ex As DirectoryNotFoundException
            Throw
        End Try

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
        _sheet.Name = _sheetName

        _sheets.Append(_sheet)

        ' By default, AutoSave = true
        '_workbookPart.Workbook.Save()

        _spreadSheetDocument.Close()
    End Sub

    Public Sub Write(text As String)
        If Not String.IsNullOrWhiteSpace(text) Then
            Write(text, "A", 1)
        Else
            Throw New ArgumentException("Nothing to write.")
        End If
    End Sub

    Public Sub Write(dataTable As DataTable)
        If dataTable.Rows.Count > 0 Then
            Dim index = 0
            Dim spreadsheetColumns = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            For Each col As DataColumn In dataTable.Columns
                Write(col.ColumnName, spreadsheetColumns(index), 1)
                index += 1
            Next
            index = 0

            For i As Integer = 0 To dataTable.Rows.Count - 1 Step 1
                For j As Integer = 0 To dataTable.Columns.Count - 1 Step 1
                    Write(dataTable.Rows(i)(dataTable.Columns(j)),
                          spreadsheetColumns(index), i + 2)
                    index += 1
                Next
                index = 0
            Next
        Else
            Throw New ArgumentException("Nothing to write.")
        End If
    End Sub

    Private Sub Write(text As String, column As String, row As Integer)
        Try
            ' Open the document for editing.
            _spreadSheetDocument = SpreadsheetDocument.Open(FilePath, True)

        Catch ex As DirectoryNotFoundException
            Throw
        End Try

        ' Get the SharedStringTablePart. If it does not exist, create a new one.
        Dim shareStringPart As SharedStringTablePart

        If (_spreadSheetDocument.WorkbookPart.GetPartsOfType(Of SharedStringTablePart).Count() > 0) Then
            shareStringPart = _spreadSheetDocument.WorkbookPart.
                GetPartsOfType(Of SharedStringTablePart).First()
        Else
            shareStringPart = _spreadSheetDocument.WorkbookPart.
                AddNewPart(Of SharedStringTablePart)()
        End If

        ' Insert the text into the SharedStringTablePart.
        Dim index As Integer = InsertSharedStringItem(text, shareStringPart)

        ' Get a reference of the first worksheet
        _worksheetPart = _spreadSheetDocument.WorkbookPart.WorksheetParts.First()

        ' Insert cell in column & row specified into the worksheet.
        Dim cell As Cell = InsertCellInWorksheet(column, row, _worksheetPart)

        ' Set the value of cell A1.
        cell.CellValue = New CellValue(index.ToString)
        cell.DataType = New EnumValue(Of CellValues)(CellValues.SharedString)

        ' Save the new worksheet.
        _worksheetPart.Worksheet.Save()

        _spreadSheetDocument.Close()

    End Sub

    ' Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
    ' and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
    Private Function InsertSharedStringItem(text As String, shareStringPart As SharedStringTablePart) As Integer
        ' If the part does not contain a SharedStringTable, create one.
        If (shareStringPart.SharedStringTable Is Nothing) Then
            shareStringPart.SharedStringTable = New SharedStringTable
        End If

        Dim i As Integer = 0

        ' Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
        For Each item As SharedStringItem In shareStringPart.SharedStringTable.Elements(Of SharedStringItem)()
            If (item.InnerText = text) Then
                Return i
            End If
            i = (i + 1)
        Next

        ' The text does not exist in the part. Create the SharedStringItem and return its index.
        shareStringPart.SharedStringTable.AppendChild(New SharedStringItem(New DocumentFormat.OpenXml.Spreadsheet.Text(text)))
        shareStringPart.SharedStringTable.Save()

        Return i
    End Function


    ' Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
    ' If the cell already exists, return it. 
    Private Function InsertCellInWorksheet(columnName As String, rowIndex As Integer, ByRef worksheetPart As WorksheetPart) As Cell
        Dim worksheet As Worksheet = worksheetPart.Worksheet
        Dim sheetData As SheetData = worksheet.GetFirstChild(Of SheetData)()
        Dim cellReference As String = (columnName + rowIndex.ToString())

        ' If the worksheet does not contain a row with the specified row index, insert one.
        Dim row As Row
        If (sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).Count() <> 0) Then
            row = sheetData.Elements(Of Row).Where(Function(r) r.RowIndex.Value = rowIndex).First()
        Else
            row = New Row()
            row.RowIndex = rowIndex
            sheetData.Append(row)
        End If

        ' If there is not a cell with the specified column name, insert one.  
        If (row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = columnName + rowIndex.ToString()).Count() > 0) Then
            Return row.Elements(Of Cell).Where(Function(c) c.CellReference.Value = cellReference).First()
        Else
            ' Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Dim refCell As Cell = Nothing
            For Each cell As Cell In row.Elements(Of Cell)()
                If (String.Compare(cell.CellReference.Value, cellReference, True) > 0) Then
                    refCell = cell
                    Exit For
                End If
            Next

            Dim newCell As Cell = New Cell
            newCell.CellReference = cellReference

            row.InsertBefore(newCell, refCell)
            worksheet.Save()

            Return newCell
        End If
    End Function

End Class
