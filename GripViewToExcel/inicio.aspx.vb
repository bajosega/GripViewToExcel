Imports System.IO
Imports ClosedXML.Excel
Imports Microsoft.Build.Framework.XamlTypes

Public Class inicio
    Inherits System.Web.UI.Page

    Protected Sub DatatableToExportExcel(ByVal nameReport As String, ByVal tabla As DataTable)

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
    ''' <param name="ds">Data Source del GripView a exportar</param>
    ''' 
    ''' <remarks>Retorna un Archivo Excel para la Descarga</remarks>
    Protected Sub GripViewToExcel(ByVal nameReport As String, ByVal wControl As GridView, ds As SqlDataSource)

        Dim dt As DataTable = New DataTable("Hoja1")
        Dim dt1 = New DataTable()
        Dim i As Integer
        Dim dato As String
        Dim dv = New DataView()

        'https://docs.microsoft.com/en-us/dotnet/api/system.data.datacolumn.datatype?view=netframework-4.5.2

        dv = ds.Select(DataSourceSelectArguments.Empty)
        dt1 = dv.ToTable() ' se usa la tabla para obtener el tipo de dato de la columna 

        ' Crear Columnas 
        For i = 0 To wControl.Columns.Count - 1
            dt.Columns.Add(HttpUtility.HtmlDecode(Trim(wControl.Columns(i).HeaderText))).DataType = dt1.Columns(wControl.Columns(i).SortExpression ).DataType
        Next

        ' Pasar los datos 
        For i = 0 To wControl.Rows.Count - 1
            dt.Rows.Add()
            For ii As Integer = 0 To wControl.Columns.Count - 1
                dato = Trim(wControl.Rows(i).Cells(ii).Text)
                If (dato <> "") Then
                    dt.Rows(i)(ii) = HttpUtility.HtmlDecode(dato)
                End If
            Next
        Next

        ''agregar en caso de que tengan Footer como un row mas. 
        'If (wControl.ShowFooter) Then
        '    dt.Rows.Add()
        '    For ii As Integer = 0 To wControl.Columns.Count - 1
        '        dt.Rows(i)(ii) = HttpUtility.HtmlDecode(Trim(wControl.Columns(ii).FooterText))
        '    Next
        'End If


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
        GripViewToExcel("reporte", GridView1, SqlDataSource1)
    End Sub
End Class