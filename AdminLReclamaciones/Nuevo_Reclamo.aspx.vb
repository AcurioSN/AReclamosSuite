Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.Security
Imports Negocio
Imports System.IO
Imports System.Net.Mail
Imports System
Imports System.Collections
Imports System.Drawing.Drawing2D
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Net
Imports System.Configuration
Imports System.Web.Script.Services.ScriptMethodAttribute
Imports System.Web.Services.WebMethodAttribute
Imports System.Web.Services
Imports System.DirectoryServices


'Agregar Referencia PDF - IM
Imports iTextSharp.text
Imports iTextSharp.text.pdf
Public Class Nuevo_Reclamo
    Inherits System.Web.UI.Page
    Dim obj As New Negocio.NLibro_Reclamacion
    Dim objN As New Negocio.NReclamaciones
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'verificacion de que la Cookie exista
        CargarActivacionGeneral()


        If Not Page.IsPostBack Then
            If Session("user") IsNot Nothing Then
                Lista_Marcas() 'Llenar el combo desplegable de marcas
                Lista_Reclamo_Queja() 'Llena el combo desplegable de queja o reclamo
                Lista_Canal_Pedido() 'Llena el combo desplegable canal de pedido
                Lista_Distritos() 'Llena el combo desplegable de distritos
                Lista_Producto_Servicio()

                Dim user As String
                user = Session("user")
                lblnombreusuario.Text = Session("user").ToString()
                lblemail.Text = Session("email_usuario").ToString()


                lbldatosusuario.Text = Session("Datos_usuario").ToString()
                lbladusuario.Text = Session("user").ToString()

                'Limpia los datos de la unidad seleccionable
                'lblruc.Text = ""
                'lbldireccion.Text = ""
                'lblrazonsocial.Text = ""
                'lblnomunidad.Text = ""

                Dim id_perfil As String
                id_perfil = Session("id_perfil").ToString()
                If id_perfil <> 3 And id_perfil <> 1 Then 'Solo ingresa el administrador de la unidad
                    Response.Redirect("~/main.aspx")
                End If
                'Registro Auditoria
                registro_acceso_pagina(Session("ACTIVACION_GENERAL").ToString(), Session("SistemaAcceso").ToString(), Session("user").ToString())

            End If



        End If
    End Sub
    Public Sub Lista_Producto_Servicio()
        Dim ds As New DataSet
        ds = obj.Lista_Producto_Servicio()
        cboproducto_servicio.DataSource = ds.Tables(0)
        cboproducto_servicio.DataTextField = "descripcion"
        cboproducto_servicio.DataValueField = "id"
        cboproducto_servicio.DataBind()



        ds.Dispose()
    End Sub
    Public Sub Lista_Reclamo_Queja()
        Dim ds As New DataSet
        ds = obj.Lista_Reclamo_Queja()
        cboquejareclamo.DataSource = ds.Tables(0)
        cboquejareclamo.DataTextField = "descripcion"
        cboquejareclamo.DataValueField = "id"
        cboquejareclamo.DataBind()



        ds.Dispose()
    End Sub
    Public Sub Lista_Distritos()
        Dim ds As New DataSet
        ds = obj.Lista_Distritos()
        cbodistrito.DataSource = ds.Tables(0)
        cbodistrito.DataTextField = "desc_distrito"
        cbodistrito.DataValueField = "id_distrito"
        cbodistrito.DataBind()



        ds.Dispose()

    End Sub
    Public Sub Lista_Canal_Pedido()
        Dim ds As New DataSet
        ds = obj.Lista_Canal_Pedido()
        cbocanalPedido.DataSource = ds.Tables(0)
        cbocanalPedido.DataTextField = "desc_canal"
        cbocanalPedido.DataValueField = "id_canal"
        cbocanalPedido.DataBind()



        ds.Dispose()
    End Sub
    Public Sub Lista_Marcas()
        Dim ds As New DataSet
        ds = obj.Lista_Marcas()
        cbomarca.DataSource = ds.Tables(0)
        cbomarca.DataTextField = "desc_marca"
        cbomarca.DataValueField = "id_marca"
        cbomarca.DataBind()
        ds.Dispose()
    End Sub
    Public Sub Lista_Unidades(ByVal id_marca As Integer)
        Dim ds As New DataSet
        Dim user As String
        user = Session("user")

        ds = obj.Lista_Unidades_Marca_Fisico(id_marca, user)
        cbolocal.DataSource = ds.Tables(0)
        cbolocal.DataTextField = "desc_unidad"
        cbolocal.DataValueField = "id_unidad"
        cbolocal.DataBind()

        'lblruc.Text = ""
        'lbldireccion.Text = ""
        'lblrazonsocial.Text = ""
        'lblnomunidad.Text = ""

        ds.Dispose()
    End Sub
    Public Sub Lista_Detalle_Unidad(ByVal id_unidad As Integer)
        Dim ds As New DataSet
        'ds = obj.Lista_Detalle_Unidad(id_unidad)

        'lblruc.Text = ds.Tables(0).Rows(0)(2).ToString()
        'lbldireccion.Text = ds.Tables(0).Rows(0)(1).ToString()
        'lblrazonsocial.Text = ds.Tables(0).Rows(0)(3).ToString()
        'lblnomunidad.Text = ds.Tables(0).Rows(0)(0).ToString()
        ds.Dispose()

    End Sub
    Protected Sub btnregistrar_Click(sender As Object, e As EventArgs) Handles btnregistrar.Click
        Try

            If chkanulado_fisico.Checked = True Then

                Registro_Documento_Anulado()

            Else
                Registro_Documento()
            End If



        Catch ex As Exception

            'save(ex.Message.ToString())
            'mensajeError(ex.Message.ToString())
        End Try

    End Sub

    Public Sub Registro_Documento()

        'Proceso de Registro
        ''-------------------------------------------------------
        Dim NroDocFisico As String
        Dim nom_consumidor As String
        Dim domicilio_consumidor As String
        Dim dni_consumidor As String
        Dim telef_consumidor As String
        Dim email_consumidor As String
        Dim menor_edad As String
        Dim reclamo_queja As String
        Dim detalle_reclamo_queja As String
        Dim detalle_canal_pedido As String
        Dim id_canal As Integer
        Dim id_distrito As Integer
        Dim id_unidad As Integer
        Dim producto_servicio As String
        Dim monto_reclamado As String
        Dim desc_bien_contratado As String

        'Captura de Parametros
        NroDocFisico = txtNroFisico.Text.ToString()
        nom_consumidor = txtNombre.Text.ToString() + " " + txtApellido.Text.ToString()
        domicilio_consumidor = txtDomicilio.Text.ToString()
        dni_consumidor = txtDNI.Text.ToString()
        telef_consumidor = txtTelefono.Text.ToString()
        email_consumidor = txtEmail.Text.Trim.ToString()
        menor_edad = txtNombrePadre.Text.ToString()
        reclamo_queja = Integer.Parse(cboquejareclamo.SelectedValue.ToString())
        detalle_reclamo_queja = txtDetalle.Text.ToString()
        detalle_canal_pedido = txtPedido.Text.ToString
        id_canal = Integer.Parse(cbocanalPedido.SelectedValue.ToString())
        id_distrito = Integer.Parse(cbodistrito.SelectedValue.ToString())
        id_unidad = Integer.Parse(cbolocal.SelectedValue.ToString())
        producto_servicio = cboproducto_servicio.SelectedValue.ToString()
        monto_reclamado = txtMonto.Text
        desc_bien_contratado = txtDescripcion.Text.ToString()

        'System.Threading.Thread.Sleep(900)
        If Session("dt_adjuntos") Is Nothing Then
            lblmensaje_foto.Text = "Debe adjuntar la fotografía del reclamo."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion2", "$('#modal-validacion2').modal();", True)
            Return
        End If

        'If chkanulado_fisico.Checked = True Then
        '    If txtsustentoanulacion.Text = "" Then
        '        lblmensaje_foto.Text = "El documento al ser anulado, debe ingresar un sustento."
        '        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion2", "$('#modal-validacion2').modal();", True)
        '        Return
        '    End If
        'End If


        'Fecha del documento
        Dim Fecha_Emision_1 As String
        Dim fecha_inicial As String
        Dim año_fecha_emision As String
        Dim mes_fecha_emision As String
        Dim dia_fecha_emision As String
        Dim fecha_emision As String
        If txtFechaDocumento.Text = "" Then
            fecha_inicial = ""

        Else
            Fecha_Emision_1 = txtFechaDocumento.Text
            año_fecha_emision = Fecha_Emision_1.Substring(0, 4).ToString()
            mes_fecha_emision = Fecha_Emision_1.Substring(5, 2).ToString()
            dia_fecha_emision = Fecha_Emision_1.Substring(8, 2).ToString()
            fecha_emision = año_fecha_emision + mes_fecha_emision + dia_fecha_emision
            fecha_inicial = fecha_emision

        End If

        'Registramos el Nro de documento.

        Dim Nro_Documento As String
        Nro_Documento = obj.RegistrarDocumento(NroDocFisico, nom_consumidor, domicilio_consumidor, dni_consumidor, telef_consumidor, email_consumidor, menor_edad, reclamo_queja,
           detalle_reclamo_queja, detalle_canal_pedido, id_canal, id_distrito, id_unidad, producto_servicio, monto_reclamado, desc_bien_contratado, fecha_inicial)


        Dim j As Integer = 0
        Dim Adjuntos_Cliente As String
        Dim Adjuntos_Cliente_Temp As String
        'Ubicaciones
        Adjuntos_Cliente = ConfigurationManager.AppSettings("Adjuntos_Cliente")
        Adjuntos_Cliente_Temp = ConfigurationManager.AppSettings("Adjuntos_Cliente_Temp")

        Dim dt_adjuntos As New DataTable()
        dt_adjuntos = Session("dt_adjuntos")
        If dt_adjuntos IsNot Nothing Then
            For j = 0 To dt_adjuntos.Rows.Count - 1
                Dim nombre_archivo As String
                nombre_archivo = dt_adjuntos.Rows(j)(1).ToString()

                'Copio archivo a la carpeta principal
                My.Computer.FileSystem.CopyFile(Adjuntos_Cliente_Temp + "\" + nombre_archivo, Adjuntos_Cliente + "\" + nombre_archivo, True)
                'Eliminar archivo de la ubicacion origen
                My.Computer.FileSystem.DeleteFile(Adjuntos_Cliente_Temp + "\" + nombre_archivo)

                objN.Registra_Archivos_Adjuntos_Fisico(Nro_Documento, nombre_archivo)
            Next
        End If

        'If chkanulado_fisico.Checked = True Then
        '    objN.Anular_Documento_fisico(Nro_Documento, txtsustentoanulacion.Text)
        'End If


        'Popu de mensaje de confirmacion de guardado.
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "alerta", "<script>popup()</script>", False)






    End Sub

    Public Sub Registro_Documento_Anulado()


        'Proceso de Registro
        ''-------------------------------------------------------
        Dim NroDocFisico As String
        Dim nom_consumidor As String
        Dim domicilio_consumidor As String
        Dim dni_consumidor As String
        Dim telef_consumidor As String
        Dim email_consumidor As String
        Dim menor_edad As String
        Dim reclamo_queja As String
        Dim detalle_reclamo_queja As String
        Dim detalle_canal_pedido As String
        Dim id_canal As Integer
        Dim id_distrito As Integer
        Dim id_unidad As Integer
        Dim producto_servicio As String
        Dim monto_reclamado As String
        Dim desc_bien_contratado As String

        'Captura de Parametros
        NroDocFisico = txtNroFisico.Text.ToString()
        If NroDocFisico = "" Then
            lblmensaje_foto.Text = "Ingresar obligatoriamente el Nro. de Documento Físico del Reclamo."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion2", "$('#modal-validacion2').modal();", True)
            Return
        End If

        nom_consumidor = txtNombre.Text.ToString() + " " + txtApellido.Text.ToString()
        If nom_consumidor = " " Then
            nom_consumidor = "[VACÍO]"
        End If

        domicilio_consumidor = txtDomicilio.Text.ToString()
        If domicilio_consumidor = "" Then
            domicilio_consumidor = "[VACÍO]"
        End If

        dni_consumidor = txtDNI.Text.ToString()
        If dni_consumidor = "" Then
            dni_consumidor = "[VACÍO]"
        End If

        telef_consumidor = txtTelefono.Text.ToString()
        email_consumidor = txtEmail.Text.Trim.ToString()
        menor_edad = txtNombrePadre.Text.ToString()

        reclamo_queja = 1

        detalle_reclamo_queja = txtDetalle.Text.ToString()
        If detalle_reclamo_queja = "" Then
            detalle_reclamo_queja = "[VACÍO]"
        End If

        detalle_canal_pedido = txtPedido.Text.ToString
        If detalle_canal_pedido = "" Then
            detalle_canal_pedido = "[VACÍO]"
        End If

        id_canal = 1
        id_distrito = 0

        Dim id_marca As Integer
        id_marca = cbomarca.SelectedIndex.ToString()
        If id_marca = 0 Then
            lblmensaje_foto.Text = "Seleccione una Local para poder continuar."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion2", "$('#modal-validacion2').modal();", True)
            Return
        End If

        id_unidad = Integer.Parse(cbolocal.SelectedValue.ToString())
        If id_unidad = 0 Then
            lblmensaje_foto.Text = "Seleccione una Local para poder continuar."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion2", "$('#modal-validacion2').modal();", True)
            Return
        End If

        producto_servicio = "1"
        monto_reclamado = txtMonto.Text
        desc_bien_contratado = txtDescripcion.Text.ToString()

        If chkanulado_fisico.Checked = True Then
            If txtsustentoanulacion.Text = "" Then
                lblmensaje_foto.Text = "El documento al ser anulado, debe ingresar un sustento."
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion2", "$('#modal-validacion2').modal();", True)
                Return
            End If
        End If


        'Fecha del documento
        Dim Fecha_Emision_1 As String
        Dim fecha_inicial As String
        Dim año_fecha_emision As String
        Dim mes_fecha_emision As String
        Dim dia_fecha_emision As String
        Dim fecha_emision As String
        If txtFechaDocumento.Text = "" Then
            fecha_inicial = ""
            lblmensaje_foto.Text = "Ingresar obligatoriamente la fecha del Nro. de Documento Físico."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion2", "$('#modal-validacion2').modal();", True)
            Return
        Else
            Fecha_Emision_1 = txtFechaDocumento.Text
            año_fecha_emision = Fecha_Emision_1.Substring(0, 4).ToString()
            mes_fecha_emision = Fecha_Emision_1.Substring(5, 2).ToString()
            dia_fecha_emision = Fecha_Emision_1.Substring(8, 2).ToString()
            fecha_emision = año_fecha_emision + mes_fecha_emision + dia_fecha_emision
            fecha_inicial = fecha_emision

        End If

        'Registramos el Nro de documento.

        Dim Nro_Documento As String
        Nro_Documento = obj.RegistrarDocumento(NroDocFisico, nom_consumidor, domicilio_consumidor, dni_consumidor, telef_consumidor, email_consumidor, menor_edad, reclamo_queja,
           detalle_reclamo_queja, detalle_canal_pedido, id_canal, id_distrito, id_unidad, producto_servicio, monto_reclamado, desc_bien_contratado, fecha_inicial)


        If chkanulado_fisico.Checked = True Then
            objN.Anular_Documento_fisico(Nro_Documento, txtsustentoanulacion.Text)
        End If


        'Popu de mensaje de confirmacion de guardado.
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "alerta", "<script>popup()</script>", False)


    End Sub
    Public Sub mensajeError(ByVal msj As String)
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "alertIns", "alert('" & msj & "');", True)
    End Sub

    Public Shared Sub save(ByVal ex As String)
        Dim fecha As String = System.DateTime.Now.ToString("yyyyMMdd")
        Dim hora As String = System.DateTime.Now.ToString("HH:mm:ss")
        Dim path As String = HttpContext.Current.Request.MapPath("log/" & fecha & ".txt")
        Dim sw As StreamWriter = New StreamWriter(path, True)
        Dim stacktrace As StackTrace = New StackTrace()

        sw.WriteLine(stacktrace.GetFrame(1).GetMethod().Name & " - " + ex)
        sw.WriteLine("")
        sw.Flush()
        sw.Close()
    End Sub

    Protected Sub cbomarca_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbomarca.SelectedIndexChanged
        Try
            Dim id_marca As Integer

            If cbomarca.SelectedIndex <> 0 Then
                id_marca = Integer.Parse(cbomarca.SelectedValue.ToString())
                Lista_Unidades(id_marca)


                'If cbomarca.SelectedValue = "1" Or cbomarca.SelectedValue = "4" Or cbomarca.SelectedValue = "5" Or
                '    cbomarca.SelectedValue = "6" Or cbomarca.SelectedValue = "7" Or
                '    cbomarca.SelectedValue = "9" Then

                '    linkpoliticas.Visible = True
                '    linkpoliticas_anticuchos.Visible = False
                '    Linkpoliticas_papachos.Visible = False
                '    Linkpoliticas_ch_arequipa.Visible = False
                '    Linkpoliticas_ch_cusco.Visible = False
                'End If
                'If cbomarca.SelectedValue = "2" Then
                '    linkpoliticas_anticuchos.Visible = True
                '    linkpoliticas.Visible = False
                '    Linkpoliticas_papachos.Visible = False
                '    Linkpoliticas_ch_arequipa.Visible = False
                '    Linkpoliticas_ch_cusco.Visible = False
                'End If
                'If cbomarca.SelectedValue = "3" Then
                '    Linkpoliticas_papachos.Visible = True
                '    linkpoliticas.Visible = False
                '    linkpoliticas_anticuchos.Visible = False
                '    Linkpoliticas_ch_arequipa.Visible = False
                '    Linkpoliticas_ch_cusco.Visible = False
                'End If
                'If cbomarca.SelectedValue = "8" Then 'CHICHA AREQUIPA 15/09/2022
                '    Linkpoliticas_papachos.Visible = False
                '    linkpoliticas.Visible = False
                '    linkpoliticas_anticuchos.Visible = False
                '    Linkpoliticas_ch_arequipa.Visible = True
                '    Linkpoliticas_ch_cusco.Visible = False
                'End If
                'If cbomarca.SelectedValue = "10" Then 'CHICHA CUSCO 15/09/2022
                '    Linkpoliticas_papachos.Visible = False
                '    linkpoliticas.Visible = False
                '    linkpoliticas_anticuchos.Visible = False
                '    Linkpoliticas_ch_arequipa.Visible = False
                '    Linkpoliticas_ch_cusco.Visible = True
                'End If

                Return
            Else
                Lista_Unidades(0)
                'lblruc.Text = ""
                'lbldireccion.Text = ""
                'lblrazonsocial.Text = ""
                'lblnomunidad.Text = ""

            End If


        Catch ex As Exception

        End Try
    End Sub

    Protected Sub cbolocal_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbolocal.SelectedIndexChanged
        Try
            Dim id_unidad As Integer

            If cbolocal.SelectedIndex <> 0 Then
                id_unidad = Integer.Parse(cbolocal.SelectedValue.ToString())
                Lista_Detalle_Unidad(id_unidad)
                Return
            Else
                Lista_Unidades(0)


            End If


        Catch ex As Exception

        End Try

    End Sub

    Protected Sub btncerrar_Click(sender As Object, e As EventArgs) Handles btncerrar.Click
        'Cuando salga el mensaje de confirmacion de guardado. refrescamos nuevamente la pagina.
        Response.Redirect("Nuevo_Reclamo.aspx")
    End Sub

    Protected Sub btnAdjuntarP_Click(sender As Object, e As EventArgs) Handles btnAdjuntarP.Click
        Try
            Dim IPSERVER As String
            Dim itemAgregado As String
            Dim existe As String = "0"
            Dim ruta As String
            IPSERVER = ConfigurationManager.AppSettings("Adjuntos_Cliente_Temp")

            Dim dt_adjuntos As New DataTable()
            dt_adjuntos.Columns.AddRange(New DataColumn() {
                                                New DataColumn("nom_archivo_original", GetType(String)),
                                                New DataColumn("nom_archivo_server", GetType(String))
            })


            ' If FileUpload1.HasFile Then
            If inputGroupFile04.HasFile Then
                Dim año As String
                Dim mes As String
                Dim dia As String
                Dim Hor As String
                Dim min As String
                Dim seg As String

                año = Date.Now.Year.ToString()
                mes = Date.Now.Month.ToString()
                dia = Date.Now.Day.ToString()
                Hor = Date.Now.Hour.ToString()
                min = Date.Now.Minute.ToString()
                seg = Date.Now.Second.ToString()

                Dim ext As String
                ext = System.IO.Path.GetExtension(inputGroupFile04.FileName)

                If ext.ToString() <> ".jpg" And ext.ToString() <> ".jpeg" Then
                    Return
                End If


                Dim archivo_general As String
                Dim ad_usuario As String
                ad_usuario = Session("user").ToString()

                archivo_general = ad_usuario + "_" + año + mes + dia + Hor + min + seg + "_" + inputGroupFile04.FileName

                inputGroupFile04.SaveAs(IPSERVER + "\" + archivo_general) 'Guardamos en Temporal antes de la principal

                'If lblidestado.Text = "5" Then 'Si ya es rechazado por el area legal

                '    If Session("dt_adjuntos_estado_rechazado") Is Nothing Then
                '        dt_adjuntos.Rows.Add(inputGroupFile04.FileName.ToString(), archivo_general)
                '        Session("dt_adjuntos_estado_rechazado") = dt_adjuntos
                '    Else
                '        dt_adjuntos = Session("dt_adjuntos_estado_rechazado")
                '        dt_adjuntos.Rows.Add(inputGroupFile04.FileName.ToString(), archivo_general)
                '        Session("dt_adjuntos_estado_rechazado") = dt_adjuntos
                '    End If

                'Else 'caso nuevo
                If Session("dt_adjuntos") Is Nothing Then
                    dt_adjuntos.Rows.Add(inputGroupFile04.FileName.ToString(), archivo_general)
                    Session("dt_adjuntos") = dt_adjuntos
                Else
                    dt_adjuntos = Session("dt_adjuntos")
                    dt_adjuntos.Rows.Add(inputGroupFile04.FileName.ToString(), archivo_general)
                    Session("dt_adjuntos") = dt_adjuntos

                End If
                'End If




                'Almacena en datatable - grilla

                grvadjuntos.DataSource = dt_adjuntos
                grvadjuntos.DataBind()

                'txtuno.Text = dt_adjuntos.Rows.Count.ToString()


            End If
        Catch ex As Exception

        End Try

    End Sub
    Protected Sub grvadjuntos_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles grvadjuntos.RowCommand
        If e.CommandName = "EliminaAdjunto" Then
            Dim NombreArchivo_eliminar As String
            NombreArchivo_eliminar = (e.CommandArgument).ToString()

            Dim dt_adjuntos As New DataTable()
            dt_adjuntos.Columns.AddRange(New DataColumn() {
                                                New DataColumn("nom_archivo_original", GetType(String)),
                                                New DataColumn("nom_archivo_server", GetType(String))
            })

            dt_adjuntos = Session("dt_adjuntos")

            Dim i As Integer = 0
            Dim can As Integer = 0
            can = dt_adjuntos.Rows.Count

            For i = 0 To can - 1
                Dim nom_archivo As String
                nom_archivo = dt_adjuntos.Rows(i)(0).ToString

                If NombreArchivo_eliminar = nom_archivo Then
                    dt_adjuntos.Rows(i).Delete()
                    Exit For
                End If
            Next




            Session("dt_adjuntos") = dt_adjuntos
            grvadjuntos.DataSource = dt_adjuntos
            grvadjuntos.DataBind()

            'txtuno.Text = dt_adjuntos.Rows.Count.ToString()

        End If
    End Sub
    Protected Sub chkanulado_fisico_CheckedChanged(sender As Object, e As EventArgs) Handles chkanulado_fisico.CheckedChanged
        If chkanulado_fisico.Checked = True Then
            txtsustentoanulacion.Visible = True
        Else
            txtsustentoanulacion.Visible = False
        End If
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

    Public Sub CerrarSesion_ActivacionGeneral()
        objN.CerrarSesionGlobal(Session("ACTIVACION_GENERAL").ToString())
    End Sub

    Public Sub registro_acceso_pagina(ByVal TokenGlobal As String, ByVal sistema As String, ByVal user As String)
        objN.Registro_SesionSistema(TokenGlobal, user, sistema, "Nuevo_Reclamo.aspx")
    End Sub

    Public Function CargarActivacionGeneral() As Boolean

        Dim token As String = String.Empty
        Dim ds As New DataSet
        If Session("ACTIVACION_GENERAL") IsNot Nothing Then
            token = Session("ACTIVACION_GENERAL").ToString()

            ds = objN.ExisteTokenActivo(token)
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