Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.Security
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
Imports System.Collections
Imports System.Drawing.Drawing2D
Imports System.Data.OleDb
Imports System.Net
Imports System.Configuration
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Style
Imports System.Drawing
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Imports Font = iTextSharp.text.Font
Imports System.Drawing.Image

Imports Microsoft.Office.Interop
Public Class Report_01
    Inherits System.Web.UI.Page
    Dim obj As New Negocio.NReclamaciones
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'verificacion de que la Cookie exista
        CargarActivacionGeneral()

        If Not Page.IsPostBack Then
            If Session("user") IsNot Nothing Then
                Dim user As String
                user = Session("user")
                lbluserad.Text = Session("user").ToString()
                'lblemail.Text = Session("email_usuario").ToString()


                'lbldatosusuario.Text = Session("Datos_usuario").ToString()
                'lbladusuario.Text = Session("user").ToString()

                Lista_Unidades_Acceso()
                lista_tipo()
                'ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)

                'Registro Auditoria
                registro_acceso_pagina(Session("ACTIVACION_GENERAL").ToString(), Session("SistemaAcceso").ToString(), Session("user").ToString())

            End If

        End If


    End Sub
    Public Sub lista_tipo()

        Dim table As New DataTable

        ' Create 3 typed columns in the DataTable.
        table.Columns.Add("cod", GetType(Integer))
        table.Columns.Add("nom", GetType(String))

        ' Add five rows with those columns filled in the DataTable.
        table.Rows.Add(0, "-- Seleccione --")
        table.Rows.Add(1, "Virtual")
        table.Rows.Add(2, "Físico")

        cbotipo.DataSource = table
        cbotipo.DataTextField = "nom"
        cbotipo.DataValueField = "cod"
        cbotipo.DataBind()


    End Sub
    Public Sub Lista_Unidades_Acceso()
        Dim ds As New DataSet
        ds = Session("Acceso_Unidades")

        cbounidad.DataSource = ds.Tables(1)
        cbounidad.DataTextField = "desc_unidad"
        cbounidad.DataValueField = "cod_unidad_usuario"
        cbounidad.DataBind()

        ds.Dispose()
    End Sub
    Protected Sub btnbuscar_Click(sender As Object, e As EventArgs) Handles btnbuscar.Click
        Dim ds_busqueda As New DataSet
        Dim Fecha_Emision_1 As String
        Dim año_fecha_emision As String
        Dim mes_fecha_emision As String
        Dim dia_fecha_emision As String
        Dim fecha_emision As String = ""
        Dim fecha_inicial As String
        Dim fecha_final As String
        Dim codigo_unidad As String
        Dim tipo As String
        Dim dni_cliente As String


        '--cbounidad.SelectedIndex = 0 Or
        If txtfecIni.Text = "" Or txtfecFin.Text = "" Then
            Return
        End If
        'If txtfecIni.Text = "" Or txtfecFin.Text = "" Then
        '    Return
        'End If

        Fecha_Emision_1 = txtfecIni.Text
        año_fecha_emision = Fecha_Emision_1.Substring(0, 4).ToString()
        mes_fecha_emision = Fecha_Emision_1.Substring(5, 2).ToString()
        dia_fecha_emision = Fecha_Emision_1.Substring(8, 2).ToString()
        fecha_emision = año_fecha_emision + mes_fecha_emision + dia_fecha_emision
        fecha_inicial = fecha_emision

        Fecha_Emision_1 = txtfecFin.Text
        año_fecha_emision = Fecha_Emision_1.Substring(0, 4).ToString()
        mes_fecha_emision = Fecha_Emision_1.Substring(5, 2).ToString()
        dia_fecha_emision = Fecha_Emision_1.Substring(8, 2).ToString()
        fecha_emision = año_fecha_emision + mes_fecha_emision + dia_fecha_emision
        fecha_final = fecha_emision

        codigo_unidad = cbounidad.SelectedValue.ToString()

        dni_cliente = txtdnicliente.Text.ToString()
        tipo = cbotipo.SelectedValue.ToString()

        ds_busqueda = obj.Busqueda_Report01(fecha_inicial, fecha_final, codigo_unidad, dni_cliente, tipo)

        grvResultado.DataSource = ds_busqueda.Tables(0)
        grvResultado.DataBind()

        Dim odt As New DataTable
        Dim odt2 As New DataTable
        Dim odt3 As New DataTable
        Dim odt4 As New DataTable
        Dim odt5 As New DataTable

        odt = ds_busqueda.Tables(0)
        odt2 = ds_busqueda.Tables(1)
        odt3 = ds_busqueda.Tables(2)
        odt4 = ds_busqueda.Tables(3)
        odt5 = ds_busqueda.Tables(4)


        Session("RESULTADO_PROCESADO") = odt
        Session("RESULTADO_PROCESADO2") = odt2
        Session("RESULTADO_PROCESADO3") = odt3
        Session("RESULTADO_PROCESADO4") = odt4
        Session("RESULTADO_PROCESADO5") = odt5


        lblcan.Text = odt.Rows.Count.ToString()
    End Sub
    Protected Sub btnexcel_Click(sender As Object, e As EventArgs) Handles btnexcel.Click
        Dim odt As New DataTable
        Dim can_registros As Integer
        odt = Session("RESULTADO_PROCESADO")

        can_registros = lblcan.Text.ToString()

        If can_registros = 0 Then
            Return
        End If

        ExportarExcel(odt)
    End Sub
    Public Sub ExportarExcel(ByVal odt As DataTable)
        Dim opcion As Boolean = True
        Dim Unidad As String
        Dim fase As String
        Dim ds As New DataSet
        Dim ds2 As New DataSet
        Dim ds3 As New DataSet
        Dim Datos As New DataTable

        Try

            Datos = odt

            'If Datos.Rows.Count = 0 Then
            '    mensajeError("No existe ninguna Evaluacion.")
            '    opcion = False
            '    Exit Sub
            'End If

            Using package = New ExcelPackage
                package.Workbook.Properties.Author = "Acurio Restaurantes"
                package.Workbook.Properties.Title = "Libro de Reclamaciones"
                Dim HojaExcel = package.Workbook.Worksheets.Add("Libro de Reclamaciones")

                HojaExcel.Name = "Libro de Reclamaciones"

                'Hoja 1 : EVALUACIONES

                Dim formatRango As ExcelRange
                formatRango = HojaExcel.Cells(1, 1, 1, 1)
                formatRango.Value = "Reclamos Virtuales Acurio Restaurantes"
                formatRango.Style.Font.Size = 12
                formatRango.Style.Font.Color.SetColor(Color.Gray)

                formatRango = HojaExcel.Cells(2, 1, 2, 20)
                formatRango.Style.Font.Bold = True
                formatRango.Style.Font.Size = 7
                formatRango.Style.Fill.PatternType = ExcelFillStyle.Solid
                formatRango.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 51, 102))
                formatRango.Style.Font.Color.SetColor(Color.White)

                HojaExcel.Column(2).Width = 15.86

                formatRango = HojaExcel.Cells(2, 1, 3, 1)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Nro Hoja"

                formatRango = HojaExcel.Cells(2, 2, 3, 2)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Fecha Registro"

                formatRango = HojaExcel.Cells(2, 3, 3, 3)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Natural/Juridica"


                formatRango = HojaExcel.Cells(2, 4, 3, 4)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Razón Social/Nombre"

                'HojaExcel.Column(4).Width = 16.14


                formatRango = HojaExcel.Cells(2, 5, 3, 5)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Domicilio"

                formatRango = HojaExcel.Cells(2, 6, 3, 6)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Tipo Doc."

                'HojaExcel.Column(7).Width = 17.86

                formatRango = HojaExcel.Cells(2, 7, 3, 7)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Nro."

                formatRango = HojaExcel.Cells(2, 8, 3, 8)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Teléfono"

                formatRango = HojaExcel.Cells(2, 9, 3, 9)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Value = "Email"

                formatRango = HojaExcel.Cells(2, 10, 3, 10)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Menor de Edad"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 11, 3, 11)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Apoderado"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 12, 3, 12)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Bien / Servicio"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 13, 3, 13)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Monto reclamado"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 14, 3, 14)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Descripción"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 15, 3, 15)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Reclamo/Queja"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 16, 3, 16)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Detalle del reclamo"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 17, 3, 17)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Pedido del consumidor"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 18, 3, 18)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Fecha Respuesta"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 19, 3, 19)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Respuesta"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 20, 3, 20)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Estado"
                formatRango.Style.Border.Bottom.Style = 4


                HojaExcel.Column(3).Width = 20
                HojaExcel.Column(4).Width = 50
                HojaExcel.Column(5).Width = 50
                HojaExcel.Column(6).Width = 20
                HojaExcel.Column(7).Width = 20
                HojaExcel.Column(8).Width = 20
                HojaExcel.Column(9).Width = 20
                HojaExcel.Column(12).Width = 20
                HojaExcel.Column(14).Width = 50
                HojaExcel.Column(15).Width = 20
                HojaExcel.Column(16).Width = 50
                HojaExcel.Column(17).Width = 50

                'HojaExcel.Column(2).Width = 40
                'HojaExcel.Column(3).Width = 15
                'HojaExcel.Column(13).Width = 30
                'HojaExcel.Column(14).Width = 30
                'HojaExcel.Column(15).Width = 30
                'HojaExcel.Column(16).Width = 30
                'HojaExcel.Column(17).Width = 30
                'HojaExcel.Column(18).Width = 10

                Dim y As Integer = 3

                formatRango = HojaExcel.Cells(y + 1, 1, 1 + (y + Datos.Rows.Count - 1), 20)
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Font.Size = 7
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4

                For Each x In Datos.Rows
                    y = y + 1
                    HojaExcel.Cells(y, 1).Value = CStr(x(0).ToString)
                    HojaExcel.Cells(y, 2).Value = CStr(x(1).ToString)
                    HojaExcel.Cells(y, 3).Value = CStr(x(2).ToString)
                    HojaExcel.Cells(y, 4).Value = CStr(x(3).ToString)
                    HojaExcel.Cells(y, 5).Value = CStr(x(4).ToString)
                    HojaExcel.Cells(y, 6).Value = CStr(x(5).ToString)
                    HojaExcel.Cells(y, 7).Value = CStr(x(6).ToString)
                    HojaExcel.Cells(y, 8).Value = CStr(x(7).ToString)
                    HojaExcel.Cells(y, 9).Value = CStr(x(8).ToString)
                    HojaExcel.Cells(y, 10).Value = CStr(x(9).ToString)
                    HojaExcel.Cells(y, 11).Value = CStr(x(10).ToString)
                    HojaExcel.Cells(y, 12).Value = CStr(x(11).ToString)
                    HojaExcel.Cells(y, 13).Value = CStr(x(12).ToString)
                    HojaExcel.Cells(y, 14).Value = CStr(x(13).ToString)
                    HojaExcel.Cells(y, 15).Value = CStr(x(14).ToString)
                    HojaExcel.Cells(y, 16).Value = CStr(x(15).ToString)
                    HojaExcel.Cells(y, 17).Value = CStr(x(16).ToString)
                    HojaExcel.Cells(y, 18).Value = CStr(x(17).ToString)
                    HojaExcel.Cells(y, 19).Value = CStr(x(18).ToString)
                    HojaExcel.Cells(y, 20).Value = CStr(x(19).ToString)


                Next


                'formatRango.Style.Font.Bold = True
                formatRango.Worksheet.View.ShowGridLines = False


                Response.BinaryWrite(package.GetAsByteArray())
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                Response.AddHeader("Content-Disposition", "attachment; filename=LibroReclamaciones_RGeneral.xlsx")
                'Response.AddHeader("Content-Disposition", "attachment; filename=LibroReclamaciones_RGeneral.csv")
            End Using
        Catch ex As Exception
            mensajeError(ex.Message & " - ExportarExcel")
        Finally
            If opcion Then
                Response.End()
                Response.Flush()
            End If
        End Try
    End Sub

    Public Sub ExportarExcel_Analisis(ByVal odt As DataTable, ByVal odt2 As DataTable, ByVal odt3 As DataTable, ByVal odt4 As DataTable)
        Dim opcion As Boolean = True
        Dim Unidad As String
        Dim fase As String
        Dim ds As New DataSet
        Dim ds2 As New DataSet
        Dim ds3 As New DataSet
        Dim Datos As New DataTable
        Dim Datos2 As New DataTable
        Dim Datos3 As New DataTable
        Dim Datos4 As New DataTable

        Try

            Datos = odt
            Datos2 = odt2
            Datos3 = odt3
            Datos4 = odt4

            'If Datos.Rows.Count = 0 Then
            '    mensajeError("No existe ninguna Evaluacion.")
            '    opcion = False
            '    Exit Sub
            'End If

            Using package = New ExcelPackage
                package.Workbook.Properties.Author = "Acurio Restaurantes"
                package.Workbook.Properties.Title = "Libro de Reclamaciones"
                Dim HojaExcel = package.Workbook.Worksheets.Add("Libro de Reclamaciones")

                HojaExcel.Name = "Reclamos por Cliente"

                'Hoja 1 : EVALUACIONES

                Dim formatRango As ExcelRange
                formatRango = HojaExcel.Cells(1, 1, 1, 1)
                formatRango.Value = "Reclamos Virtuales Acurio Restaurantes - Análisis"
                formatRango.Style.Font.Size = 12
                formatRango.Style.Font.Color.SetColor(Color.Gray)

                formatRango = HojaExcel.Cells(2, 1, 2, 20)
                formatRango.Style.Font.Bold = True
                formatRango.Style.Font.Size = 7
                formatRango.Style.Fill.PatternType = ExcelFillStyle.Solid
                formatRango.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 51, 102))
                formatRango.Style.Font.Color.SetColor(Color.White)

                HojaExcel.Column(2).Width = 15.86

                formatRango = HojaExcel.Cells(2, 1, 3, 1)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Nro. doc"

                formatRango = HojaExcel.Cells(2, 2, 3, 2)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Nombre Cliente"

                formatRango = HojaExcel.Cells(2, 3, 3, 3)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Doc. Cliente"


                formatRango = HojaExcel.Cells(2, 4, 3, 4)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Domicilio"

                'HojaExcel.Column(4).Width = 16.14


                formatRango = HojaExcel.Cells(2, 5, 3, 5)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Teléfono"

                formatRango = HojaExcel.Cells(2, 6, 3, 6)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Email"

                'HojaExcel.Column(7).Width = 17.86

                formatRango = HojaExcel.Cells(2, 7, 3, 7)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Fecha de Reclamo"

                formatRango = HojaExcel.Cells(2, 8, 3, 8)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Tipo de Reclamo"

                formatRango = HojaExcel.Cells(2, 9, 3, 9)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Value = "Tipo de Bien"

                formatRango = HojaExcel.Cells(2, 10, 3, 10)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Marca"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 11, 3, 11)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Tienda"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 12, 3, 12)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Canal"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 13, 3, 13)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Motivo de Reclamo"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 14, 3, 14)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Monto de Reclamo"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 15, 3, 15)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Fecha de Atención del Administrador"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 16, 3, 16)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Estado de Reclamo"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 17, 3, 17)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Fecha de Atención de Legal"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 18, 3, 18)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Administrador"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 19, 3, 19)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Gerente de Marca"
                formatRango.Style.Border.Bottom.Style = 4

                formatRango = HojaExcel.Cells(2, 20, 3, 20)
                formatRango.Merge = True
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4
                formatRango.Value = "Fecha de Atención del Gerente"
                formatRango.Style.Border.Bottom.Style = 4


                HojaExcel.Column(3).Width = 20
                HojaExcel.Column(4).Width = 20
                HojaExcel.Column(5).Width = 20
                HojaExcel.Column(6).Width = 20
                HojaExcel.Column(7).Width = 20
                HojaExcel.Column(8).Width = 20
                HojaExcel.Column(9).Width = 20
                HojaExcel.Column(12).Width = 20
                HojaExcel.Column(14).Width = 20
                HojaExcel.Column(15).Width = 20
                HojaExcel.Column(16).Width = 20
                HojaExcel.Column(17).Width = 20
                HojaExcel.Column(18).Width = 20
                HojaExcel.Column(19).Width = 20
                HojaExcel.Column(20).Width = 20

                'HojaExcel.Column(2).Width = 40
                'HojaExcel.Column(3).Width = 15
                'HojaExcel.Column(13).Width = 30
                'HojaExcel.Column(14).Width = 30
                'HojaExcel.Column(15).Width = 30
                'HojaExcel.Column(16).Width = 30
                'HojaExcel.Column(17).Width = 30
                'HojaExcel.Column(18).Width = 10

                Dim y As Integer = 3

                formatRango = HojaExcel.Cells(y + 1, 1, 1 + (y + Datos.Rows.Count - 1), 20)
                formatRango.Style.WrapText = True
                formatRango.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango.Style.Font.Size = 7
                formatRango.Style.Border.Top.Style = 4
                formatRango.Style.Border.Left.Style = 4
                formatRango.Style.Border.Right.Style = 4
                formatRango.Style.Border.Bottom.Style = 4

                For Each x In Datos.Rows
                    y = y + 1
                    HojaExcel.Cells(y, 1).Value = CStr(x(0).ToString)
                    HojaExcel.Cells(y, 2).Value = CStr(x(1).ToString)
                    HojaExcel.Cells(y, 3).Value = CStr(x(2).ToString)
                    HojaExcel.Cells(y, 4).Value = CStr(x(3).ToString)
                    HojaExcel.Cells(y, 5).Value = CStr(x(4).ToString)
                    HojaExcel.Cells(y, 6).Value = CStr(x(5).ToString)
                    HojaExcel.Cells(y, 7).Value = CStr(x(6).ToString)
                    HojaExcel.Cells(y, 8).Value = CStr(x(7).ToString)
                    HojaExcel.Cells(y, 9).Value = CStr(x(8).ToString)
                    HojaExcel.Cells(y, 10).Value = CStr(x(9).ToString)
                    HojaExcel.Cells(y, 11).Value = CStr(x(10).ToString)
                    HojaExcel.Cells(y, 12).Value = CStr(x(11).ToString)
                    HojaExcel.Cells(y, 13).Value = CStr(x(12).ToString)
                    HojaExcel.Cells(y, 14).Value = CStr(x(13).ToString)
                    HojaExcel.Cells(y, 15).Value = CStr(x(14).ToString)
                    HojaExcel.Cells(y, 16).Value = CStr(x(15).ToString)
                    HojaExcel.Cells(y, 17).Value = CStr(x(16).ToString)
                    HojaExcel.Cells(y, 18).Value = CStr(x(17).ToString)
                    HojaExcel.Cells(y, 19).Value = CStr(x(18).ToString)
                    HojaExcel.Cells(y, 20).Value = CStr(x(19).ToString)


                Next


                'formatRango.Style.Font.Bold = True
                formatRango.Worksheet.View.ShowGridLines = False




                'Hoja 2

                Dim HojaExcel2 = package.Workbook.Worksheets.Add("Libro de Reclamaciones")

                HojaExcel2.Name = "Soluciones por cliente"

                Dim formatRango2 As ExcelRange
                formatRango2 = HojaExcel2.Cells(1, 1, 1, 1)
                formatRango2.Value = "Reclamos Virtuales Acurio Restaurantes - Análisis"
                formatRango2.Style.Font.Size = 12
                formatRango2.Style.Font.Color.SetColor(Color.Gray)

                formatRango2 = HojaExcel2.Cells(2, 1, 2, 21)
                formatRango2.Style.Font.Bold = True
                formatRango2.Style.Font.Size = 7
                formatRango2.Style.Fill.PatternType = ExcelFillStyle.Solid
                formatRango2.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 51, 102))
                formatRango2.Style.Font.Color.SetColor(Color.White)

                HojaExcel2.Column(2).Width = 15.86

                formatRango2 = HojaExcel2.Cells(2, 1, 3, 1)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Nro. doc"

                formatRango2 = HojaExcel2.Cells(2, 2, 3, 2)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Nombre Cliente"

                formatRango2 = HojaExcel2.Cells(2, 3, 3, 3)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Doc. Cliente"


                formatRango2 = HojaExcel2.Cells(2, 4, 3, 4)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Domicilio"

                'HojaExcel.Column(4).Width = 16.14


                formatRango2 = HojaExcel2.Cells(2, 5, 3, 5)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Teléfono"

                formatRango2 = HojaExcel2.Cells(2, 6, 3, 6)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Email"

                'HojaExcel.Column(7).Width = 17.86

                formatRango2 = HojaExcel2.Cells(2, 7, 3, 7)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Fecha de Reclamo"

                formatRango2 = HojaExcel2.Cells(2, 8, 3, 8)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Tipo de Reclamo"

                formatRango2 = HojaExcel2.Cells(2, 9, 3, 9)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Value = "Tipo de Bien"

                formatRango2 = HojaExcel2.Cells(2, 10, 3, 10)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Marca"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 11, 3, 11)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Tienda"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 12, 3, 12)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Canal"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 13, 3, 13)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Acción Realizada"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 14, 3, 14)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Monto de Reclamo"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 15, 3, 15)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Importe de Cortesía"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 16, 3, 16)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Fecha de Atención del Administrador"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 17, 3, 17)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Estado de Reclamo"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 18, 3, 18)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Fecha de Atención de Legal"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 19, 3, 19)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Administrador"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 20, 3, 20)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Gerente de Marca"
                formatRango2.Style.Border.Bottom.Style = 4

                formatRango2 = HojaExcel2.Cells(2, 21, 3, 21)
                formatRango2.Merge = True
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4
                formatRango2.Value = "Fecha de Atención del Gerente"
                formatRango2.Style.Border.Bottom.Style = 4


                HojaExcel2.Column(3).Width = 20
                HojaExcel2.Column(4).Width = 20
                HojaExcel2.Column(5).Width = 20
                HojaExcel2.Column(6).Width = 20
                HojaExcel2.Column(7).Width = 20
                HojaExcel2.Column(8).Width = 20
                HojaExcel2.Column(9).Width = 20
                HojaExcel2.Column(12).Width = 20
                HojaExcel2.Column(14).Width = 20
                HojaExcel2.Column(15).Width = 20
                HojaExcel2.Column(16).Width = 20
                HojaExcel2.Column(17).Width = 20
                HojaExcel2.Column(18).Width = 20
                HojaExcel2.Column(19).Width = 20
                HojaExcel2.Column(20).Width = 20
                HojaExcel2.Column(21).Width = 20

                'HojaExcel.Column(2).Width = 40
                'HojaExcel.Column(3).Width = 15
                'HojaExcel.Column(13).Width = 30
                'HojaExcel.Column(14).Width = 30
                'HojaExcel.Column(15).Width = 30
                'HojaExcel.Column(16).Width = 30
                'HojaExcel.Column(17).Width = 30
                'HojaExcel.Column(18).Width = 10

                Dim y2 As Integer = 3

                formatRango2 = HojaExcel2.Cells(y2 + 1, 1, 1 + (y2 + Datos2.Rows.Count - 1), 21)
                formatRango2.Style.WrapText = True
                formatRango2.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango2.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango2.Style.Font.Size = 7
                formatRango2.Style.Border.Top.Style = 4
                formatRango2.Style.Border.Left.Style = 4
                formatRango2.Style.Border.Right.Style = 4
                formatRango2.Style.Border.Bottom.Style = 4

                For Each x2 In Datos2.Rows
                    y2 = y2 + 1
                    HojaExcel2.Cells(y2, 1).Value = CStr(x2(0).ToString)
                    HojaExcel2.Cells(y2, 2).Value = CStr(x2(1).ToString)
                    HojaExcel2.Cells(y2, 3).Value = CStr(x2(2).ToString)
                    HojaExcel2.Cells(y2, 4).Value = CStr(x2(3).ToString)
                    HojaExcel2.Cells(y2, 5).Value = CStr(x2(4).ToString)
                    HojaExcel2.Cells(y2, 6).Value = CStr(x2(5).ToString)
                    HojaExcel2.Cells(y2, 7).Value = CStr(x2(6).ToString)
                    HojaExcel2.Cells(y2, 8).Value = CStr(x2(7).ToString)
                    HojaExcel2.Cells(y2, 9).Value = CStr(x2(8).ToString)
                    HojaExcel2.Cells(y2, 10).Value = CStr(x2(9).ToString)
                    HojaExcel2.Cells(y2, 11).Value = CStr(x2(10).ToString)
                    HojaExcel2.Cells(y2, 12).Value = CStr(x2(11).ToString)
                    HojaExcel2.Cells(y2, 13).Value = CStr(x2(12).ToString)
                    HojaExcel2.Cells(y2, 14).Value = CStr(x2(13).ToString)
                    HojaExcel2.Cells(y2, 15).Value = CStr(x2(14).ToString)
                    HojaExcel2.Cells(y2, 16).Value = CStr(x2(15).ToString)
                    HojaExcel2.Cells(y2, 17).Value = CStr(x2(16).ToString)
                    HojaExcel2.Cells(y2, 18).Value = CStr(x2(17).ToString)
                    HojaExcel2.Cells(y2, 19).Value = CStr(x2(18).ToString)
                    HojaExcel2.Cells(y2, 20).Value = CStr(x2(19).ToString)
                    HojaExcel2.Cells(y2, 21).Value = CStr(x2(20).ToString)

                Next


                'formatRango.Style.Font.Bold = True
                formatRango2.Worksheet.View.ShowGridLines = False






                'Hoja 3

                Dim HojaExcel3 = package.Workbook.Worksheets.Add("Libro de Reclamaciones")

                HojaExcel3.Name = "Reclamos por Empresa"

                Dim formatRango3 As ExcelRange
                formatRango3 = HojaExcel3.Cells(1, 1, 1, 1)
                formatRango3.Value = "Reclamos Virtuales Acurio Restaurantes - Análisis"
                formatRango3.Style.Font.Size = 12
                formatRango3.Style.Font.Color.SetColor(Color.Gray)

                formatRango3 = HojaExcel3.Cells(2, 1, 2, 17)
                formatRango3.Style.Font.Bold = True
                formatRango3.Style.Font.Size = 7
                formatRango3.Style.Fill.PatternType = ExcelFillStyle.Solid
                formatRango3.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 51, 102))
                formatRango3.Style.Font.Color.SetColor(Color.White)

                HojaExcel3.Column(2).Width = 15.86

                formatRango3 = HojaExcel3.Cells(2, 1, 3, 1)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Nro. doc"

                formatRango3 = HojaExcel3.Cells(2, 2, 3, 2)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Empresa"

                formatRango3 = HojaExcel3.Cells(2, 3, 3, 3)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Fecha de Reclamo"


                formatRango3 = HojaExcel3.Cells(2, 4, 3, 4)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Tipo Reclamo"

                'HojaExcel.Column(4).Width = 16.14


                formatRango3 = HojaExcel3.Cells(2, 5, 3, 5)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Tipo de Bien"

                formatRango3 = HojaExcel3.Cells(2, 6, 3, 6)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Marca"

                'HojaExcel.Column(7).Width = 17.86

                formatRango3 = HojaExcel3.Cells(2, 7, 3, 7)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Tienda"

                formatRango3 = HojaExcel3.Cells(2, 8, 3, 8)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Canal"

                formatRango3 = HojaExcel3.Cells(2, 9, 3, 9)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Value = "Motivo de Reclamo"

                formatRango3 = HojaExcel3.Cells(2, 10, 3, 10)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Monto de Reclamo"
                formatRango3.Style.Border.Bottom.Style = 4

                formatRango3 = HojaExcel3.Cells(2, 11, 3, 11)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Importe de Cortesía"
                formatRango3.Style.Border.Bottom.Style = 4

                formatRango3 = HojaExcel3.Cells(2, 12, 3, 12)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Fecha de Atención del Administrador"
                formatRango3.Style.Border.Bottom.Style = 4

                formatRango3 = HojaExcel3.Cells(2, 13, 3, 13)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Estado de Reclamo"
                formatRango3.Style.Border.Bottom.Style = 4

                formatRango3 = HojaExcel3.Cells(2, 14, 3, 14)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Fecha de Atención de legal"
                formatRango3.Style.Border.Bottom.Style = 4

                formatRango3 = HojaExcel3.Cells(2, 15, 3, 15)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Administrador"
                formatRango3.Style.Border.Bottom.Style = 4

                formatRango3 = HojaExcel3.Cells(2, 16, 3, 16)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Gerente de Marca"
                formatRango3.Style.Border.Bottom.Style = 4

                formatRango3 = HojaExcel3.Cells(2, 17, 3, 17)
                formatRango3.Merge = True
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4
                formatRango3.Value = "Fecha de atención del Gerente"
                formatRango3.Style.Border.Bottom.Style = 4




                HojaExcel3.Column(3).Width = 20
                HojaExcel3.Column(4).Width = 20
                HojaExcel3.Column(5).Width = 20
                HojaExcel3.Column(6).Width = 20
                HojaExcel3.Column(7).Width = 20
                HojaExcel3.Column(8).Width = 20
                HojaExcel3.Column(9).Width = 20
                HojaExcel3.Column(12).Width = 20
                HojaExcel3.Column(14).Width = 20
                HojaExcel3.Column(15).Width = 20
                HojaExcel3.Column(16).Width = 20
                HojaExcel3.Column(17).Width = 20


                Dim y3 As Integer = 3

                formatRango3 = HojaExcel3.Cells(y3 + 1, 1, 1 + (y3 + Datos3.Rows.Count - 1), 17)
                formatRango3.Style.WrapText = True
                formatRango3.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango3.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango3.Style.Font.Size = 7
                formatRango3.Style.Border.Top.Style = 4
                formatRango3.Style.Border.Left.Style = 4
                formatRango3.Style.Border.Right.Style = 4
                formatRango3.Style.Border.Bottom.Style = 4

                For Each x3 In Datos3.Rows
                    y3 = y3 + 1

                    HojaExcel3.Cells(y3, 1).Value = CStr(x3(0).ToString)
                    HojaExcel3.Cells(y3, 2).Value = CStr(x3(1).ToString)
                    HojaExcel3.Cells(y3, 3).Value = CStr(x3(2).ToString)
                    HojaExcel3.Cells(y3, 4).Value = CStr(x3(3).ToString)
                    HojaExcel3.Cells(y3, 5).Value = CStr(x3(4).ToString)
                    HojaExcel3.Cells(y3, 6).Value = CStr(x3(5).ToString)
                    HojaExcel3.Cells(y3, 7).Value = CStr(x3(6).ToString)
                    HojaExcel3.Cells(y3, 8).Value = CStr(x3(7).ToString)
                    HojaExcel3.Cells(y3, 9).Value = CStr(x3(8).ToString)
                    HojaExcel3.Cells(y3, 10).Value = CStr(x3(9).ToString)
                    HojaExcel3.Cells(y3, 11).Value = CStr(x3(10).ToString)
                    HojaExcel3.Cells(y3, 12).Value = CStr(x3(11).ToString)
                    HojaExcel3.Cells(y3, 13).Value = CStr(x3(12).ToString)
                    HojaExcel3.Cells(y3, 14).Value = CStr(x3(13).ToString)
                    HojaExcel3.Cells(y3, 15).Value = CStr(x3(14).ToString)
                    HojaExcel3.Cells(y3, 16).Value = CStr(x3(15).ToString)
                    HojaExcel3.Cells(y3, 17).Value = CStr(x3(16).ToString)


                Next


                'formatRango.Style.Font.Bold = True
                formatRango3.Worksheet.View.ShowGridLines = False






                'Hoja 4

                Dim HojaExcel4 = package.Workbook.Worksheets.Add("Libro de Reclamaciones")

                HojaExcel4.Name = "Soluciones por Empresa"

                Dim formatRango4 As ExcelRange
                formatRango4 = HojaExcel4.Cells(1, 1, 1, 1)
                formatRango4.Value = "Reclamos Virtuales Acurio Restaurantes - Análisis"
                formatRango4.Style.Font.Size = 12
                formatRango4.Style.Font.Color.SetColor(Color.Gray)

                formatRango4 = HojaExcel4.Cells(2, 1, 2, 17)
                formatRango4.Style.Font.Bold = True
                formatRango4.Style.Font.Size = 7
                formatRango4.Style.Fill.PatternType = ExcelFillStyle.Solid
                formatRango4.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 51, 102))
                formatRango4.Style.Font.Color.SetColor(Color.White)

                HojaExcel4.Column(2).Width = 15.86

                formatRango4 = HojaExcel4.Cells(2, 1, 3, 1)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Nro. doc"

                formatRango4 = HojaExcel4.Cells(2, 2, 3, 2)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Empresa"

                formatRango4 = HojaExcel4.Cells(2, 3, 3, 3)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Fecha de Reclamo"


                formatRango4 = HojaExcel4.Cells(2, 4, 3, 4)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Tipo Reclamo"

                'HojaExcel.Column(4).Width = 16.14


                formatRango4 = HojaExcel4.Cells(2, 5, 3, 5)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Tipo de Bien"

                formatRango4 = HojaExcel4.Cells(2, 6, 3, 6)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Marca"

                'HojaExcel.Column(7).Width = 17.86

                formatRango4 = HojaExcel4.Cells(2, 7, 3, 7)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Tienda"

                formatRango4 = HojaExcel4.Cells(2, 8, 3, 8)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Canal"

                formatRango4 = HojaExcel4.Cells(2, 9, 3, 9)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Value = "Acción Realizada"

                formatRango4 = HojaExcel4.Cells(2, 10, 3, 10)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Monto de Reclamo"
                formatRango4.Style.Border.Bottom.Style = 4

                formatRango4 = HojaExcel4.Cells(2, 11, 3, 11)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Importe de Cortesía"
                formatRango4.Style.Border.Bottom.Style = 4

                formatRango4 = HojaExcel4.Cells(2, 12, 3, 12)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Fecha de Atención del Administrador"
                formatRango4.Style.Border.Bottom.Style = 4

                formatRango4 = HojaExcel4.Cells(2, 13, 3, 13)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Estado de Reclamo"
                formatRango4.Style.Border.Bottom.Style = 4

                formatRango4 = HojaExcel4.Cells(2, 14, 3, 14)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Fecha de Atención de legal"
                formatRango4.Style.Border.Bottom.Style = 4

                formatRango4 = HojaExcel4.Cells(2, 15, 3, 15)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Administrador"
                formatRango4.Style.Border.Bottom.Style = 4

                formatRango4 = HojaExcel4.Cells(2, 16, 3, 16)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Gerente de Marca"
                formatRango4.Style.Border.Bottom.Style = 4

                formatRango4 = HojaExcel4.Cells(2, 17, 3, 17)
                formatRango4.Merge = True
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4
                formatRango4.Value = "Fecha de atención del Gerente"
                formatRango4.Style.Border.Bottom.Style = 4


                HojaExcel4.Column(3).Width = 20
                HojaExcel4.Column(4).Width = 20
                HojaExcel4.Column(5).Width = 20
                HojaExcel4.Column(6).Width = 20
                HojaExcel4.Column(7).Width = 20
                HojaExcel4.Column(8).Width = 20
                HojaExcel4.Column(9).Width = 20
                HojaExcel4.Column(12).Width = 20
                HojaExcel4.Column(14).Width = 20
                HojaExcel4.Column(15).Width = 20
                HojaExcel4.Column(16).Width = 20
                HojaExcel4.Column(17).Width = 20


                Dim y4 As Integer = 3

                formatRango4 = HojaExcel4.Cells(y4 + 1, 1, 1 + (y4 + Datos4.Rows.Count - 1), 17)
                formatRango4.Style.WrapText = True
                formatRango4.Style.VerticalAlignment = ExcelVerticalAlignment.Center
                formatRango4.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                formatRango4.Style.Font.Size = 7
                formatRango4.Style.Border.Top.Style = 4
                formatRango4.Style.Border.Left.Style = 4
                formatRango4.Style.Border.Right.Style = 4
                formatRango4.Style.Border.Bottom.Style = 4


                For Each x4 In Datos4.Rows
                    y4 = y4 + 1
                    HojaExcel4.Cells(y4, 1).Value = CStr(x4(0).ToString)
                    HojaExcel4.Cells(y4, 2).Value = CStr(x4(1).ToString)
                    HojaExcel4.Cells(y4, 3).Value = CStr(x4(2).ToString)
                    HojaExcel4.Cells(y4, 4).Value = CStr(x4(3).ToString)
                    HojaExcel4.Cells(y4, 5).Value = CStr(x4(4).ToString)
                    HojaExcel4.Cells(y4, 6).Value = CStr(x4(5).ToString)
                    HojaExcel4.Cells(y4, 7).Value = CStr(x4(6).ToString)
                    HojaExcel4.Cells(y4, 8).Value = CStr(x4(7).ToString)
                    HojaExcel4.Cells(y4, 9).Value = CStr(x4(8).ToString)
                    HojaExcel4.Cells(y4, 10).Value = CStr(x4(9).ToString)
                    HojaExcel4.Cells(y4, 11).Value = CStr(x4(10).ToString)
                    HojaExcel4.Cells(y4, 12).Value = CStr(x4(11).ToString)
                    HojaExcel4.Cells(y4, 13).Value = CStr(x4(12).ToString)
                    HojaExcel4.Cells(y4, 14).Value = CStr(x4(13).ToString)
                    HojaExcel4.Cells(y4, 15).Value = CStr(x4(14).ToString)
                    HojaExcel4.Cells(y4, 16).Value = CStr(x4(15).ToString)
                    HojaExcel4.Cells(y4, 17).Value = CStr(x4(16).ToString)

                Next

                'formatRango.Style.Font.Bold = True
                formatRango4.Worksheet.View.ShowGridLines = False





                Response.BinaryWrite(package.GetAsByteArray())
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                Response.AddHeader("Content-Disposition", "attachment; filename=LibroReclamaciones_RGeneral.xlsx")
            End Using
        Catch ex As Exception
            mensajeError(ex.Message & " - ExportarExcel")
        Finally
            If opcion Then
                Response.End()
                Response.Flush()
            End If
        End Try
    End Sub
    Public Sub mensajeError(ByVal msj As String)
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "alertIns", "alert('" & msj & "');", True)
    End Sub
    Protected Sub btnCerrarSesion_Click(sender As Object, e As EventArgs) Handles btnCerrarSesion.Click
        CerrarSesion_ActivacionGeneral()

        Dim cogeCookie = Request.Cookies.Get("appNameAuth")
        If Not cogeCookie Is Nothing Then
            Request.Cookies.Remove("appNameAuth")
        End If

        FormsAuthentication.SignOut() 'ahora si cierras session!!!
        Session.Abandon()
        Session.Clear()
        Response.Redirect("admin_documentos.aspx")


    End Sub
    Protected Sub btnTxt_Click(sender As Object, e As EventArgs) Handles btnTxt.Click

        Dim odt As New DataTable
        odt = Session("RESULTADO_PROCESADO")

        If Session("RESULTADO_PROCESADO") Is Nothing Then
            Return
        End If
        Dim txt_principal As String = String.Empty

        Dim i As Integer = 0
        Dim col_1 As String
        Dim col_2 As String
        Dim valor As String

        For x = 0 To odt.Rows.Count - 1
            For i = 20 To 37 '5 Columnas
                valor = odt.Rows(x)(i).ToString()
                If i = 37 Then 'ultima columna
                    txt_principal += valor
                Else
                    txt_principal += valor + "|"
                End If

            Next

            txt_principal += vbCr & vbLf
        Next



        'Dim dt As New DataTable()
        'dt.Columns.AddRange(New DataColumn(2) {New DataColumn("Id"), New DataColumn("Name"), New DataColumn("Country")})
        'dt.Rows.Add(1, "John Hammond", "United States")
        'dt.Rows.Add(2, "Mudassar Khan", "India")
        'dt.Rows.Add(3, "Suzanne Mathews", "France")
        'dt.Rows.Add(4, "Robert Schidner", "Russia")
        'GridView1.DataSource = dt
        'GridView1.DataBind()

        ''Build the Text file data.
        'Dim txt As String = String.Empty

        'For Each cell As TableCell In GridView1.HeaderRow.Cells
        '    'Add the Header row for Text file.
        '    txt += cell.Text + vbTab & vbTab
        'Next

        ''Add new line.
        'txt += vbCr & vbLf

        'For Each row As GridViewRow In GridView1.Rows
        '    For Each cell As TableCell In row.Cells
        '        'Add the Data rows.
        '        txt += cell.Text + vbTab & vbTab
        '    Next

        '    'Add new line.
        '    txt += vbCr & vbLf
        'Next

        'Download the Text file.
        Response.Clear()
        Response.Buffer = True
        Response.AddHeader("content-disposition", "attachment;filename=IndecopiReclamaciones.txt")
        Response.Charset = ""
        Response.ContentType = "application/text"
        Response.Output.Write(txt_principal)
        Response.Flush()
        Response.End()

    End Sub
    Protected Sub btnexcel_analisis_Click(sender As Object, e As EventArgs) Handles btnexcel_analisis.Click
        Dim odt As New DataTable
        Dim odt2 As New DataTable
        Dim odt3 As New DataTable
        Dim odt4 As New DataTable

        Dim can_registros As Integer
        odt = Session("RESULTADO_PROCESADO2")
        odt2 = Session("RESULTADO_PROCESADO3")
        odt3 = Session("RESULTADO_PROCESADO4")
        odt4 = Session("RESULTADO_PROCESADO5")

        can_registros = lblcan.Text.ToString()

        If can_registros = 0 Then
            Return
        End If

        ExportarExcel_Analisis(odt, odt2, odt3, odt4)
    End Sub
    Private Sub btnPWBI_Click(sender As Object, e As EventArgs) Handles btnPWBI.Click
        Dim url As String
        url = ConfigurationManager.AppSettings("Url_PowerBI")
        Response.Write("<script>window.open('" + url + "','_blank');</script>")
    End Sub

    Public Sub CerrarSesion_ActivacionGeneral()
        obj.CerrarSesionGlobal(Session("ACTIVACION_GENERAL").ToString())
    End Sub

    Public Sub registro_acceso_pagina(ByVal TokenGlobal As String, ByVal sistema As String, ByVal user As String)
        obj.Registro_SesionSistema(TokenGlobal, user, sistema, "Report_01.aspx")
    End Sub

    Public Function CargarActivacionGeneral() As Boolean


        Dim token As String = String.Empty
        Dim ds As New DataSet
        If Session("ACTIVACION_GENERAL") IsNot Nothing Then
            token = Session("ACTIVACION_GENERAL").ToString()

            ds = obj.ExisteTokenActivo(token)
            If ds.Tables(0).Rows.Count > 0 Then
                Return True
            Else
                Response.Redirect("~/SesionExpirada.aspx", False)
                Return False
            End If
        Else
            Response.Redirect("~/SesionExpirada.aspx")
            Return False
        End If

    End Function
End Class