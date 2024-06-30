Imports System
Imports System.Data
Imports System.Configuration
Imports System.IO
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls

Imports NPOI.HSSF.UserModel
Imports NPOI.HPSF
Imports NPOI.POIFS.FileSystem
Imports NPOI.SS.UserModel

Public Class GridViewExportUtil



    Public Shared Sub Export(ByVal fileName As String, ByVal gv As GridView)
        HttpContext.Current.Response.Clear()
        HttpContext.Current.Response.AddHeader("content-disposition", String.Format("attachment; filename={0}", fileName))
        HttpContext.Current.Response.ContentType = "application/ms-excel"

        ' Create NPOI Workbook
        Dim HssfWb As HSSFWorkbook = New HSSFWorkbook()
        InitializeWorkbook(HssfWb)
        Dim sheet1 As Sheet = HssfWb.CreateSheet("Sheet1")

        Dim rowCount As Integer = 0
        Dim colCount As Integer = 0

        Dim headerFont As Font = HssfWb.CreateFont()
        headerFont.Boldweight = FontBoldWeight.BOLD

        Dim headerStyle As CellStyle = HssfWb.CreateCellStyle()
        headerStyle.BorderBottom = CellBorderType.MEDIUM
        headerStyle.BorderTop = CellBorderType.MEDIUM
        headerStyle.BorderLeft = CellBorderType.THIN
        headerStyle.BorderRight = CellBorderType.THIN
        headerStyle.FillPattern = FillPatternType.SOLID_FOREGROUND
        headerStyle.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.GREY_25_PERCENT.index
        headerStyle.SetFont(headerFont)

        Dim bodyStyle As CellStyle = HssfWb.CreateCellStyle()
        bodyStyle.BorderBottom = CellBorderType.THIN
        bodyStyle.BorderLeft = CellBorderType.THIN
        bodyStyle.BorderRight = CellBorderType.THIN

        '  add the header row to the table
        If (Not (gv.HeaderRow) Is Nothing) Then
            GridViewExportUtil.PrepareControlForExport(gv.HeaderRow)
            'table.Rows.Add(gv.HeaderRow)

            Dim sheetRow As Row = sheet1.CreateRow(rowCount)
            colCount = 0
            For Each cell As TableCell In gv.HeaderRow.Cells
                Dim sheetCell As Cell = sheetRow.CreateCell(colCount)
                sheetCell.SetCellValue(PrepareCellText(cell.Text))
                sheetCell.CellStyle = headerStyle
                colCount += 1
            Next
            rowCount += 1
        End If

        '  add each of the data rows to the table
        For Each row As GridViewRow In gv.Rows
            GridViewExportUtil.PrepareControlForExport(row)
            'table.Rows.Add(row)

            Dim sheetRow As Row = sheet1.CreateRow(rowCount)
            colCount = 0
            For Each cell As TableCell In row.Cells
                Dim sheetCell As Cell = sheetRow.CreateCell(colCount)
                sheetCell.SetCellValue(PrepareCellText(cell.Text))
                sheetCell.CellStyle = bodyStyle
                colCount += 1
            Next
            rowCount += 1
        Next

        '  add the footer row to the table
        If (Not (gv.FooterRow) Is Nothing) Then
            GridViewExportUtil.PrepareControlForExport(gv.FooterRow)
            'table.Rows.Add(gv.FooterRow)

            Dim sheetRow As Row = sheet1.CreateRow(rowCount)
            colCount = 0
            For Each cell As TableCell In gv.FooterRow.Cells
                Dim sheetCell As Cell = sheetRow.CreateCell(colCount)
                sheetCell.SetCellValue(PrepareCellText(cell.Text))
                sheetCell.CellStyle = headerStyle
                colCount += 1
            Next
            rowCount += 1

        End If

        For i As Integer = 0 To colCount - 1
            sheet1.AutoSizeColumn(i)
        Next

        HttpContext.Current.Response.BinaryWrite(WriteToStream(HssfWb).GetBuffer())
        HttpContext.Current.Response.End()
    End Sub

    Private Shared Function WriteToStream(ByRef Workbook As HSSFWorkbook) As MemoryStream
        'Write the stream data of workbook to the root directory
        Dim file As MemoryStream = New MemoryStream()
        Workbook.Write(file)
        Return file
    End Function

    Private Shared Sub InitializeWorkbook(ByRef Workbook As HSSFWorkbook)

        Workbook = New HSSFWorkbook()

        'create a entry of DocumentSummaryInformation
        Dim dsi As DocumentSummaryInformation = PropertySetFactory.CreateDocumentSummaryInformation()
        dsi.Company = "HP"
        Workbook.DocumentSummaryInformation = dsi

        'create a entry of SummaryInformation
        Dim si As SummaryInformation = PropertySetFactory.CreateSummaryInformation()
        si.Subject = "NPOI Export"
        Workbook.SummaryInformation = si

    End Sub

    Private Shared Function PrepareCellText(ByVal input As String) As String
        Dim output As String = input
        output = output.Replace("&nbsp;", " ")
        Return output
    End Function

    ' Replace any of the contained controls with literals
    Private Shared Sub PrepareControlForExport(ByVal control As Control)
        Dim i As Integer = 0
        Do While (i < control.Controls.Count)
            Dim current As Control = control.Controls(i)
            If (TypeOf current Is LinkButton) Then
                control.Controls.Remove(current)
                control.Controls.AddAt(i, New LiteralControl(CType(current, LinkButton).Text))
            ElseIf (TypeOf current Is ImageButton) Then
                control.Controls.Remove(current)
                control.Controls.AddAt(i, New LiteralControl(CType(current, ImageButton).AlternateText))
            ElseIf (TypeOf current Is System.Web.UI.WebControls.HyperLink) Then
                control.Controls.Remove(current)
                control.Controls.AddAt(i, New LiteralControl(CType(current, System.Web.UI.WebControls.HyperLink).Text))
            ElseIf (TypeOf current Is DropDownList) Then
                control.Controls.Remove(current)
                control.Controls.AddAt(i, New LiteralControl(CType(current, DropDownList).SelectedItem.Text))
            ElseIf (TypeOf current Is CheckBox) Then
                control.Controls.Remove(current)
                control.Controls.AddAt(i, New LiteralControl(CType(current, CheckBox).Checked))
                'TODO: Warning!!!, inline IF is not supported ?
            End If
            If current.HasControls Then
                GridViewExportUtil.PrepareControlForExport(current)
            End If
            i = (i + 1)
        Loop
    End Sub
End Class

