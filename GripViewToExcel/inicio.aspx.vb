Imports System.IO
Imports ClosedXML.Excel
Imports Microsoft.Build.Framework.XamlTypes

Public Class inicio
    Inherits System.Web.UI.Page

    Protected Sub DatatableToExcel(ByVal nameReport As String, ByVal tabla As DataTable)

        Using wb As XLWorkbook = New XLWorkbook()
            wb.Worksheets.Add(tabla)
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


    ''' <summary>
    ''' Este metodo exporta un GripView creado desde un DataSource a Excel.
    ''' En la definiciòn de las columnas dentro del GRIPVIEW
    ''' el SortExpression debe ser igual al DataField. 
    ''' </summary>
    ''' <param name="nameReport">Nobre para el archivo creado (colocar sin la extencion)</param>
    ''' <param name="wControl">GripView a exportar</param>
    ''' 
    ''' <remarks>Retorna un Archivo Excel para la Descarga</remarks>
    Protected Sub GripViewToExcel(ByVal nameReport As String, ByVal wControl As GridView)
        Dim i As Integer
        Dim ii As Integer

        Using wb As XLWorkbook = New XLWorkbook()

            Dim ws As IXLWorksheet = wb.Worksheets.Add("Hoja 1")
            Dim dato As String

            ''https://docs.microsoft.com/en-us/dotnet/api/system.data.datacolumn.datatype?view=netframework-4.5.2


            ' Crear Columnas 
            For i = 0 To wControl.Columns.Count - 1
                If wControl.Columns(i).Visible Then
                    ws.Cell(1, i + 1).Value = HttpUtility.HtmlDecode(wControl.Columns(i).HeaderText.Trim)
                End If
            Next

            For i = 1 To wControl.Rows.Count
                For ii = 0 To wControl.Columns.Count - 1
                    If wControl.Columns(ii).Visible Then
                        dato = Trim(wControl.Rows(i - 1).Cells(ii).Text)
                        If (dato <> "") Then
                            ws.Cell(i + 1, ii + 1).Value = HttpUtility.HtmlDecode(dato)
                        End If
                    End If
                Next
            Next

            ''agregar en caso de que tengan Footer como un row mas. 
            If (wControl.ShowFooter) Then
                For ii = 0 To wControl.Columns.Count - 1
                    If wControl.Columns(ii).Visible Then
                        ws.Cell(i + 2, ii + 1).Value = HttpUtility.HtmlDecode(wControl.Columns(ii).FooterText.Trim)
                    End If
                Next
            End If

            Response.Clear()
            Response.Buffer = True
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
        GripViewToExcel("reporte", GridView1)
    End Sub
End Class