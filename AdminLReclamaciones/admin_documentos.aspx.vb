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
Public Class admin_documentos
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

                'Valores que pueden ser utilizados más adelante.
                'lblnombreusuario.Text = Session("user").ToString()
                'lblemail.Text = Session("email_usuario").ToString()
                'lbldatosusuario.Text = Session("Datos_usuario").ToString()
                'lbladusuario.Text = Session("user").ToString()


                Lista_Unidades_Acceso()
                estados()
                'ScriptManager.RegisterStartupScript(Page, Page.GetType(), "myModal", "$('#myModal').modal();", True)


                'Registro Auditoria
                registro_acceso_pagina(Session("ACTIVACION_GENERAL").ToString(), Session("SistemaAcceso").ToString(), Session("user").ToString())


            End If

        End If
        If (Session("user") Is Nothing) Then
                Response.Redirect("~/login.aspx")
            End If

    End Sub
    Public Sub estados()
        Dim ds As New DataSet
        ds = obj.Lista_Desplegable_estados()

        cboestados.DataSource = ds.Tables(0)
        cboestados.DataTextField = "desc_estado"
        cboestados.DataValueField = "id_estado"
        cboestados.DataBind()


        ds.Dispose()
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

    Protected Sub btnbuscar_Click(sender As Object, e As EventArgs) Handles btnbuscar.Click
        'Buscar del Popup
        Dim ds_busqueda As New DataSet
        Dim Fecha_Emision_1 As String
        Dim año_fecha_emision As String
        Dim mes_fecha_emision As String
        Dim dia_fecha_emision As String
        Dim fecha_emision As String = ""
        Dim fecha_inicial As String
        Dim fecha_final As String
        Dim codigo_unidad As String
        Dim dni_cliente As String
        Dim id_perfil As Integer
        Dim User As String
        Dim id_estado As Integer

        If txtfecIni.Text = "" And txtfecFin.Text <> "" Then 'Debe Ingresar  las 2 fechas no solo 1
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "myModal", "$('#myModal').modal();", True)
            Return
        End If
        If txtfecFin.Text = "" And txtfecIni.Text <> "" Then 'Debe Ingresar  las 2 fechas no solo 1
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "myModal", "$('#myModal').modal();", True)
            Return
        End If


        If txtfecIni.Text = "" And txtfecFin.Text = "" Then
            fecha_inicial = ""
            fecha_final = ""
        Else
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
        End If

        codigo_unidad = cbounidad.SelectedValue.ToString()

        dni_cliente = txtdnicliente.Text.ToString()
        id_perfil = Session("id_perfil")

        id_estado = Integer.Parse(cboestados.SelectedValue.ToString())

        'Por perfil carga documentos

        User = Session("user")
        'ds_busqueda = obj.Busqueda_Documentos_Perfil_estado(fecha_inicial, fecha_final, codigo_unidad, dni_cliente, id_perfil, User)
        ds_busqueda = obj.Busqueda_Documentos_Perfil_estado_main(fecha_inicial, fecha_final, codigo_unidad, dni_cliente, id_perfil, User, id_estado)



        'grvResultado.DataSource = ds_busqueda.Tables(0)
        'grvResultado.DataBind()

        Dim odt As New DataTable
        odt = ds_busqueda.Tables(0)
        Session("RESULTADO_PROCESADO") = odt
        'lblcan.Text = odt.Rows.Count.ToString()
        grvDocumentos.DataSource = ds_busqueda.Tables(0)
        grvDocumentos.DataBind()


        'cantidades
        lblcan1.Text = ds_busqueda.Tables(1).Rows(2)(1).ToString()
        lblcan2.Text = ds_busqueda.Tables(1).Rows(0)(1).ToString()
        lblcan3.Text = ds_busqueda.Tables(1).Rows(1)(1).ToString()
        lblcan4.Text = ds_busqueda.Tables(1).Rows(3)(1).ToString()


        ds_busqueda.Dispose()

        grilla_estado()

    End Sub
    Public Sub grilla_estado()
        For Each gvClaserow As GridViewRow In grvDocumentos.Rows
            Dim label_Estado As Label
            Dim Estado As String

            Dim Pendiente_Atencion As Label
            Dim Atendido_Administrador As Label
            Dim Aprobado_area_legal As Label
            Dim aprobado_gerente_marca As Label
            Dim rechazado_area_legal As Label
            Dim rechazado_gerente_marca As Label
            Dim anulado_fisico As Label

            label_Estado = gvClaserow.FindControl("lblid_estado")
            Estado = label_Estado.Text.ToString()


            Pendiente_Atencion = gvClaserow.FindControl("lblestado_pendiente")
            Atendido_Administrador = gvClaserow.FindControl("lblestado_adminUnidad")
            Aprobado_area_legal = gvClaserow.FindControl("lblestado_legal")
            aprobado_gerente_marca = gvClaserow.FindControl("lblestado_GerenteMarca")
            rechazado_area_legal = gvClaserow.FindControl("lblestado_Rlegal")
            rechazado_gerente_marca = gvClaserow.FindControl("lblestado_RGerenteMarca")
            anulado_fisico = gvClaserow.FindControl("lblestado_anuladofisico")


            If Estado = 1 Then
                Pendiente_Atencion.Visible = True
            End If
            If Estado = 2 Then
                Atendido_Administrador.Visible = True
            End If
            If Estado = 3 Then
                Aprobado_area_legal.Visible = True
            End If
            If Estado = 4 Then
                aprobado_gerente_marca.Visible = True
            End If

            If Estado = 5 Then
                rechazado_area_legal.Visible = True
            End If
            If Estado = 6 Then
                rechazado_gerente_marca.Visible = True
            End If
            If Estado = 7 Then
                anulado_fisico.Visible = True
            End If

        Next
    End Sub

    Protected Sub grvDocumentos_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles grvDocumentos.RowCommand
        Dim NroDocumento As String
        Dim odt As New DataTable
        Dim NroDocumento_aux As String

        If e.CommandName = "atender" Then
            Dim fisico As String = ""
            NroDocumento = (e.CommandArgument).ToString()
            'Response.Redirect("~/Atencion_Documento.aspx?NroExpediente=" + NroExpediente.ToString() + "&NroCaso=" + "")

            odt = Session("RESULTADO_PROCESADO")
            For i = 0 To odt.Rows.Count - 1
                NroDocumento_aux = odt.Rows(i)(0).ToString()
                If NroDocumento = NroDocumento_aux Then
                    'Mostrar detalle del documento
                    fisico = odt.Rows(i)(25).ToString()
                    Exit For
                End If
            Next

            If fisico = "1" Then 'Para documento Fisico
                Response.Redirect("~/Detalle_Admin_Documento_F.aspx?dato1=" + NroDocumento.ToString())

            Else  'Para documento Digital

                Response.Redirect("~/Detalle_Admin_Documento.aspx?dato1=" + NroDocumento.ToString())
            End If


            'Response.Redirect("~/Atencion_Documento.aspx?datos1=" + CodProductoSeleccionado.ToString() + "&datos2=" + DescProductoSeleccionado.ToString())
        End If
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
    Public Sub CerrarSesion_ActivacionGeneral()
        obj.CerrarSesionGlobal(Session("ACTIVACION_GENERAL").ToString())
    End Sub
    Public Sub registro_acceso_pagina(ByVal TokenGlobal As String, ByVal sistema As String, ByVal user As String)
        obj.Registro_SesionSistema(TokenGlobal, user, sistema, "admin_documentos.aspx")
    End Sub
End Class