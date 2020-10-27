Imports System.IO
Imports ClosedXML.Excel

Public Class inicio
    Inherits System.Web.UI.Page

    Protected Sub ExportExcel(ByVal nameReport As String, ByVal wControl As GridView)
        Dim dt As DataTable = New DataTable("Hoja1")

        For i As Integer = 0 To wControl.Columns.Count - 1
            dt.Columns.Add(Trim(wControl.Columns(i).HeaderText))
        Next

        For i As Integer = 0 To wControl.Rows.Count - 1
            dt.Rows.Add()
            For ii As Integer = 0 To wControl.Columns.Count - 1
                dt.Rows(i)(ii) = HttpUtility.HtmlDecode(Trim(wControl.Rows(i).Cells(ii).Text))
            Next
        Next

        Using wb As XLWorkbook = New XLWorkbook()
            wb.Worksheets.Add(dt)
            Response.Clear()
            Response.Buffer = True
            'Response.Charset = ""
            'Response.Charset = "UTF-8"
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            Response.AddHeader("content-disposition", "attachment;filename=" & nameReport & ".xlsx")

            Using MyMemoryStream As MemoryStream = New MemoryStream()
                wb.SaveAs(MyMemoryStream)
                MyMemoryStream.WriteTo(Response.OutputStream)
                Response.Flush()
                Response.[End]()
            End Using
        End Using
    End Sub



    Protected Sub btnExportar_Click(sender As Object, e As EventArgs) Handles btnExportar.Click
        ExportExcel("Exportado", GridView1)
    End Sub
End Class