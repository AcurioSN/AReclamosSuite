Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Web
Imports System.Web.UI
'Imports System.Web.UI.WebControls
Imports System.Web.Security
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net.Mail
'Imports EASendMail
Imports System.Collections
Imports System.Drawing.Drawing2D
Imports System.Data.OleDb
Imports System.Net
Imports System.Configuration
Imports OfficeOpenXml
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml.Style
Imports System.Drawing
Imports iTextSharp.text.pdf
Imports Font = iTextSharp.text.Font
Public Class Detalle_Admin_Documento
    Inherits System.Web.UI.Page
    Dim obj As New Negocio.NReclamaciones
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'verificacion de que la Cookie exista
        CargarActivacionGeneral()

        If Not Page.IsPostBack Then
            If Session("user") IsNot Nothing Then
                Dim user As String
                Dim ds As New DataSet
                Dim id_perfil As String

                user = Session("user")
                lbluserad.Text = Session("user").ToString()
                lblemail.Text = Session("email_usuario").ToString()
                id_perfil = Session("id_perfil")

                lbldatosusuario.Text = Session("Datos_usuario").ToString()
                lbladusuario.Text = Session("user").ToString()

                Dim nro_documento As String = Request.QueryString("dato1")
                lblnrodocumento.Text = nro_documento.ToString()
                ds = obj.lista_Detalle_Documento(nro_documento)

                lblreclamo_queja.Text = ds.Tables(0).Rows(0)(9).ToString()
                lblproducto_servicio.Text = ds.Tables(0).Rows(0)(15).ToString()
                lblmonto_reclamado.Text = ds.Tables(0).Rows(0)(16).ToString()
                lbldesc_biencontratado.Text = ds.Tables(0).Rows(0)(17).ToString()
                lbldetalle_reclamo_queja.Text = ds.Tables(0).Rows(0)(10).ToString()
                lbldetalle_canal_pedido.Text = ds.Tables(0).Rows(0)(11).ToString()
                lbldescunidad.Text = ds.Tables(0).Rows(0)(18).ToString()
                lblnomcliente.Text = ds.Tables(0).Rows(0)(3).ToString()
                lbldnicliente.Text = ds.Tables(0).Rows(0)(5).ToString()
                lbldireccioncliente.Text = ds.Tables(0).Rows(0)(4).ToString()

                Dim email As String
                email = ds.Tables(0).Rows(0)(7).ToString()
                lblemail_cliente.Text = email
                lbltelefono_email.Text = ds.Tables(0).Rows(0)(6).ToString() + "/" + email
                lblmenoredad.Text = ds.Tables(0).Rows(0)(8).ToString()
                lbldnihistoria.Text = ds.Tables(0).Rows(0)(5).ToString()

                Dim id_estado As Integer
                Dim desc_estado As String
                desc_estado = ds.Tables(0).Rows(0)(20).ToString()
                id_estado = ds.Tables(0).Rows(0)(19).ToString()
                lblidestado.Text = id_estado

                If id_estado = "5" Then
                    Session("dt_adjuntos_estado_rechazado") = Nothing
                End If

                If id_estado = 1 Then
                    lblestado_pendiente.Visible = True
                    lblestado_pendiente.Text = desc_estado
                End If
                If id_estado = 2 Then
                    lblestado_adminUnidad.Visible = True
                    lblestado_adminUnidad.Text = desc_estado
                End If
                If id_estado = 3 Then
                    lblestado_legal.Visible = True
                    lblestado_legal.Text = desc_estado
                End If
                If id_estado = 4 Then
                    lblestado_GerenteMarca.Visible = True
                    lblestado_GerenteMarca.Text = desc_estado
                End If
                If id_estado = 5 Then
                    lblestado_Rlegal.Visible = True
                    lblestado_Rlegal.Text = desc_estado
                End If
                If id_estado = 6 Then
                    lblestado_RGerenteMarca.Visible = True
                    lblestado_RGerenteMarca.Text = desc_estado
                End If

                If id_estado <> 1 Then
                    Dim atencion_manual As String
                    atencion_manual = ds.Tables(1).Rows(0)(10).ToString()
                    If atencion_manual = "1" Then
                        chkmanual.Checked = True
                    Else
                        chkmanual.Checked = False
                    End If
                End If

                lblfecharegistro.Text = ds.Tables(0).Rows(0)(2).ToString()

                If ds.Tables(1).Rows.Count = 0 Then
                    lblfecha_aprobacion_gerente.Text = "" '----04/10/2022
                Else

                    Dim id_estado_aux As Integer = ds.Tables(1).Rows(0)(11).ToString()
                    If id_estado_aux = 4 Then
                        lblfecha_aprobacion_gerente.Text = ds.Tables(1).Rows(0)(12).ToString() '----04/10/2022
                    Else
                        lblfecha_aprobacion_gerente.Text = "" '----04/10/2022
                    End If
                End If

                lbldiaretraso.Text = ds.Tables(0).Rows(0)(21).ToString()
                lblidunidad.Text = ds.Tables(0).Rows(0)(14).ToString()
                listas_desplegables()

                Data_Cortesia(ds)


                perfil_estado(id_perfil, id_estado, ds)
                Lista_archivos_adjuntos(ds)

                ds.Dispose()

                'Registro Auditoria
                registro_acceso_pagina(Session("ACTIVACION_GENERAL").ToString(), Session("SistemaAcceso").ToString(), Session("user").ToString())


            End If


        End If



    End Sub
    Public Sub Data_Cortesia(ByVal ds As DataSet)
        Dim si_no_contesta As String
        Dim importe_cortesia_si As String
        Dim modo_cortesia As String
        Dim can_pers_cortesia As String
        Dim monto_max_cortesia As String
        Dim valido_cortesia As String
        Dim coordinar_cortesia As String
        Dim email_cont_cortesia As String
        Dim lugar_cortesia As String

        If ds.Tables(7).Rows.Count > 0 Then
            pnlcortesias.Visible = True
            si_no_contesta = ds.Tables(7).Rows(0)(1).ToString()
            importe_cortesia_si = ds.Tables(7).Rows(0)(2).ToString()
            modo_cortesia = ds.Tables(7).Rows(0)(3).ToString()
            can_pers_cortesia = ds.Tables(7).Rows(0)(4).ToString()
            monto_max_cortesia = ds.Tables(7).Rows(0)(5).ToString()
            valido_cortesia = ds.Tables(7).Rows(0)(6).ToString()
            coordinar_cortesia = ds.Tables(7).Rows(0)(7).ToString()
            email_cont_cortesia = ds.Tables(7).Rows(0)(8).ToString()
            lugar_cortesia = ds.Tables(7).Rows(0)(9).ToString()

            If si_no_contesta = "SI" Then
                chkopcion_cortesia1.Checked = True
                chkopcion_cortesia2.Checked = False
                txtimportecortesia.Text = importe_cortesia_si.ToString()
            End If
            If si_no_contesta = "NO" Then
                chkopcion_cortesia2.Checked = True
                cbomodo.SelectedValue = modo_cortesia
                chkopcion_cortesia1.Checked = False
                txtcanpersonas.Text = can_pers_cortesia.ToString()
                txtmontomaximo.Text = monto_max_cortesia.ToString()
                txtvalidohasta.Text = valido_cortesia.ToString()
                txtcoodinarContacto.Text = coordinar_cortesia.ToString()
                txtcorreoCortesia.Text = email_cont_cortesia.ToString()
                txtlugarcortesia.Text = lugar_cortesia
            End If
        End If


    End Sub
    Public Sub Lista_archivos_adjuntos(ByVal ds As DataSet)
        grvarchivosadjuntos.DataSource = ds.Tables(6)
        grvarchivosadjuntos.DataBind()

        If ds.Tables(6).Rows.Count > 0 Then
            Dim dt_adjuntos As New DataTable()
            dt_adjuntos = ds.Tables(6)
            Session("dt_adjuntos") = dt_adjuntos
        Else
            Session("dt_adjuntos") = Nothing
        End If


        ds.Dispose()
    End Sub
    Public Sub listas_desplegables()
        Dim ds As New DataSet

        ds = obj.Lista_Desplegables()
        'cboaccion.DataSource = ds.Tables(0)
        'cboaccion.DataTextField = "desc_accion"
        'cboaccion.DataValueField = "id_accion"
        'cboaccion.DataBind()

        chklistAccion.DataSource = ds.Tables(0)
        chklistAccion.DataTextField = "desc_accion"
        chklistAccion.DataValueField = "id_accion"
        chklistAccion.DataBind()


        chklisturgencia.DataSource = ds.Tables(1)
        chklisturgencia.DataTextField = "desc_urgencia"
        chklisturgencia.DataValueField = "id_urgencia"
        chklisturgencia.DataBind()

        chklistempreclamo.DataSource = ds.Tables(2)
        chklistempreclamo.DataTextField = "desc_emp_reclamo"
        chklistempreclamo.DataValueField = "id_emp_reclamo"
        chklistempreclamo.DataBind()

        chklistProductoServicio.DataSource = ds.Tables(3)
        chklistProductoServicio.DataTextField = "desc_prod_serv"
        chklistProductoServicio.DataValueField = "id_prod_serv"
        chklistProductoServicio.DataBind()

        '07/09/2021
        Dim tab_acciones_etiquetas As New DataTable()
        tab_acciones_etiquetas = ds.Tables(4)
        Session("tab_acciones_etiquetas") = tab_acciones_etiquetas

        Dim tab_urgencias_etiquetas As New DataTable()
        tab_urgencias_etiquetas = ds.Tables(5)
        Session("tab_urgencias_etiquetas") = tab_urgencias_etiquetas



        ds.Dispose()

    End Sub

    Public Sub perfil_estado(ByVal id_perfil As String, ByVal id_estado As String, ByVal ds As DataSet)

        If id_perfil = "3" Then 'ADMINISTRADOR DE LA UNIDAD

            If id_estado = 1 Then 'Pendiente de atención
                txtobslegal.ReadOnly = True
                txtobsGerenteMarca.ReadOnly = True
                txtobslegal.BackColor = System.Drawing.ColorTranslator.FromHtml("White")
                txtobsGerenteMarca.BackColor = System.Drawing.ColorTranslator.FromHtml("White")
                btnrechazar.Visible = False
                btnCartaPreview.Visible = True '07/09/2021
                btnadjuntos.Visible = False

            End If

            If id_estado = 2 Then 'Atendido por Administrador de la Unidad

                bloquear_controles_carga_data(ds) 'Bloquea los controles y muestra la data ingresada
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 3 Then 'Aprobado por el area legal

                bloquear_controles_carga_data(ds) 'Bloquea los controles y muestra la data ingresada
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 4 Then 'Aprobado por el gerente de Marca

                bloquear_controles_carga_data(ds) 'Bloquea los controles y muestra la data ingresada
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 5 Then 'Rechazado por el area legal
                bloquear_controles_carga_data(ds)

                'Visualizar los combos de seleccion y ocultar los labels
                habilitar_controles()

                btnrechazar.Visible = False
                txtobslegal.ReadOnly = True
                txtobsGerenteMarca.ReadOnly = True
                txtobslegal.BackColor = System.Drawing.ColorTranslator.FromHtml("White")
                txtobsGerenteMarca.BackColor = System.Drawing.ColorTranslator.FromHtml("White")


                txtotraaccion.ReadOnly = False
                txtotromotivo.ReadOnly = False
                pnladjuntar.Visible = True
                pnlcortesias.Enabled = True
                chkmanual.Enabled = True
            End If
            If id_estado = 6 Then 'Rechazado por el Gerente de Marca
                bloquear_controles_carga_data(ds)
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
        End If


        If id_perfil = "4" Then 'Area Legal
            If id_estado = 1 Then 'Pendiente de atención
                bloquear_controles_carga_data(ds)
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                btnCartaPreview.Visible = False
                btnadjuntos.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 2 Then 'Atendido por Administrador de la Unidad - Listo para que el area legal atienda el caso

                bloquear_controles_carga_data(ds) 'Bloquea los controles y muestra la data ingresada
                txtobslegal.ReadOnly = False
                txtobsGerenteMarca.ReadOnly = True
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 3 Then 'Aprobado por el area legal
                bloquear_controles_carga_data(ds)
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 4 Then 'Aprobado por el gerente de Marca
                bloquear_controles_carga_data(ds)
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 5 Then 'Rechazado por el area legal
                bloquear_controles_carga_data(ds)

                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 6 Then 'Rechazado por el Gerente de Marca
                bloquear_controles_carga_data(ds)
                txtobslegal.ReadOnly = False
                txtobsGerenteMarca.ReadOnly = True
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
        End If

        If id_perfil = "2" Then 'GERENTE DE MARCA
            If id_estado = 1 Then 'Pendiente de atención
                bloquear_controles_carga_data(ds)
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                btnCartaPreview.Visible = False
                btnadjuntos.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 2 Then 'Atendido por Administrador de la Unidad
                bloquear_controles_carga_data(ds)
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 3 Then 'Aprobado por el area legal
                bloquear_controles_carga_data(ds)
                txtobslegal.ReadOnly = True
                txtobsGerenteMarca.ReadOnly = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 4 Then 'Aprobado por el area gerente de marca
                bloquear_controles_carga_data(ds)
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 5 Then 'Rechazado por el area legal
                bloquear_controles_carga_data(ds)
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
            If id_estado = 6 Then 'Rechazado por el Gerente de Marca
                bloquear_controles_carga_data(ds)
                btnaprobar.Visible = False
                btnrechazar.Visible = False
                pnlcortesias.Enabled = False
                chkmanual.Enabled = False
            End If
        End If


    End Sub
    Public Sub habilitar_controles()
        'Habilitar Controles
        'cboaccion.Visible = True
        'lbl_cboaccion.Visible = False
        'cboproducto_servicio.Visible = True
        'lbl_cboproducto_servicio.Visible = False

        chklistProductoServicio.Enabled = True
        chklistAccion.Enabled = True

        txtobslegal.ReadOnly = True
        txtobsGerenteMarca.ReadOnly = True
        chklistempreclamo.Enabled = True
        chklisturgencia.Enabled = True
        txtcosto_accion.ReadOnly = False
        txtotros_gastos.ReadOnly = False
        txtobsadmin_unidad.ReadOnly = False

    End Sub
    Public Sub bloquear_controles_carga_data(ByVal ds As DataSet)
        Dim ds_data As New DataSet
        ds_data = ds

        'Bloquear Controles
        txtobslegal.ReadOnly = True
        txtobsGerenteMarca.ReadOnly = True
        txtotraaccion.ReadOnly = True
        txtotromotivo.ReadOnly = True
        'cboaccion.Visible = False
        'cboproducto_servicio.Visible = False
        'lbl_cboaccion.Visible = True
        'lbl_cboproducto_servicio.Visible = True
        chklistAccion.Enabled = False


        chklistProductoServicio.Enabled = False
        chklistempreclamo.Enabled = False
        chklisturgencia.Enabled = False
        txtcosto_accion.ReadOnly = True
        txtotros_gastos.ReadOnly = True

        txtobsadmin_unidad.ReadOnly = True
        txtobslegal.BackColor = System.Drawing.ColorTranslator.FromHtml("White")
        txtobsGerenteMarca.BackColor = System.Drawing.ColorTranslator.FromHtml("White")
        txtobsadmin_unidad.BackColor = System.Drawing.ColorTranslator.FromHtml("White")
        txtcosto_accion.BackColor = System.Drawing.ColorTranslator.FromHtml("White")
        txtotros_gastos.BackColor = System.Drawing.ColorTranslator.FromHtml("White")
        txtotraaccion.BackColor = System.Drawing.ColorTranslator.FromHtml("White")
        txtotromotivo.BackColor = System.Drawing.ColorTranslator.FromHtml("White")
        pnladjuntar.Visible = False
        'carga data
        If ds.Tables(1).Rows.Count > 0 Then ' Si hay data

            ' Dim id_accion As Integer = Integer.Parse(ds.Tables(1).Rows(0)(1).ToString())
            'cboaccion.SelectedValue = id_accion
            'lbl_cboaccion.Text = ds.Tables(1).Rows(0)(2).ToString()

            'Dim id_producto_servicio As Integer = ds.Tables(1).Rows(0)(5).ToString()
            'cboproducto_servicio.SelectedValue = id_producto_servicio
            'lbl_cboproducto_servicio.Text = ds.Tables(1).Rows(0)(6).ToString()


            txtobsadmin_unidad.Text = ds.Tables(1).Rows(0)(3).ToString()
            txtcosto_accion.Text = ds.Tables(1).Rows(0)(4).ToString()
            txtotros_gastos.Text = ds.Tables(1).Rows(0)(9).ToString()
            txtobsGerenteMarca.Text = ds.Tables(1).Rows(0)(8).ToString()
            txtobslegal.Text = ds.Tables(1).Rows(0)(7).ToString()

            Dim id_urgencia As Integer
            Dim item As ListItem
            Dim i As Integer
            Dim id_urgencia_x As Integer
            Dim can As Integer = ds_data.Tables(2).Rows.Count
            For i = 0 To can - 1
                id_urgencia_x = Integer.Parse(ds_data.Tables(2).Rows(i)(0).ToString())

                For Each item In chklisturgencia.Items
                    id_urgencia = item.Value.ToString()

                    If id_urgencia = id_urgencia_x Then
                        item.Selected = True
                    End If

                Next


            Next




            Dim id_emp As Integer
            Dim item2 As ListItem
            Dim i2 As Integer
            Dim id_emp_x As Integer
            Dim can2 As Integer = ds_data.Tables(3).Rows.Count
            For i2 = 0 To can2 - 1
                id_emp_x = Integer.Parse(ds_data.Tables(3).Rows(i2)(0).ToString())

                For Each item2 In chklistempreclamo.Items
                    id_emp = item2.Value.ToString()

                    If id_emp = id_emp_x Then
                        item2.Selected = True
                    End If

                Next


            Next



            Dim id_acc As Integer
            Dim det_otra_accion As String
            Dim item3 As ListItem
            Dim i3 As Integer
            Dim id_acc_x As Integer
            Dim can3 As Integer = ds_data.Tables(4).Rows.Count
            For i3 = 0 To can3 - 1
                id_acc_x = Integer.Parse(ds_data.Tables(4).Rows(i3)(0).ToString())
                det_otra_accion = ds_data.Tables(4).Rows(i3)(1).ToString()

                For Each item3 In chklistAccion.Items
                    id_acc = item3.Value.ToString()

                    If id_acc = id_acc_x Then
                        item3.Selected = True

                        If id_acc = "9" Then
                            txtotraaccion.Text = det_otra_accion
                        End If
                    End If

                Next


            Next



            Dim id_prod As Integer
            Dim item4 As ListItem
            Dim i4 As Integer
            Dim id_prod_x As Integer
            Dim can4 As Integer = ds_data.Tables(5).Rows.Count
            For i4 = 0 To can4 - 1
                id_prod_x = Integer.Parse(ds_data.Tables(5).Rows(i4)(0).ToString())

                For Each item4 In chklistProductoServicio.Items
                    id_prod = item4.Value.ToString()

                    If id_prod = id_prod_x Then
                        item4.Selected = True
                    End If

                Next


            Next



            Dim id_mot As Integer
            Dim det_otra_motivo As String
            Dim item5 As ListItem
            i3 = 0
            Dim id_motiv_x As Integer
            Dim can5 As Integer = ds_data.Tables(2).Rows.Count
            For i3 = 0 To can5 - 1
                id_motiv_x = Integer.Parse(ds_data.Tables(2).Rows(i3)(0).ToString())
                det_otra_motivo = ds_data.Tables(2).Rows(i3)(1).ToString()

                For Each item5 In chklisturgencia.Items
                    id_mot = item5.Value.ToString()

                    If id_mot = id_motiv_x Then
                        item5.Selected = True

                        If id_mot = "35" Then
                            txtotromotivo.Text = det_otra_motivo
                        End If
                    End If

                Next


            Next


        End If

    End Sub

    Protected Sub btnaprobar_Click(sender As Object, e As EventArgs) Handles btnaprobar.Click

        Dim aprobado As Integer = 1
        Registro_Aprobar_Rechazar(aprobado)

    End Sub

    Public Sub Registro_Aprobar_Rechazar(ByVal aprobado As Integer)
        Dim nro_documento As String
        Dim user_ad As String
        Dim id_accion As Integer
        Dim detalle_accion As String
        Dim costo_accion As Decimal
        Dim otros_gastos As Decimal
        Dim producto_servicio As String
        Dim obs_legal As String
        Dim obs_gerente_marca As String
        Dim id_perfil As Integer



        nro_documento = lblnrodocumento.Text.ToString()
        user_ad = Session("user").ToString()
        id_perfil = Integer.Parse(Session("id_perfil").ToString())


        'prueba'
        Dim id_urgencia2 As Integer
        Dim item2x As ListItem
        Dim salud As Integer
        salud = 0
        For Each item2x In chklisturgencia.Items
            If item2x.Selected = True Then
                'codigo para agregar registro
                id_urgencia2 = item2x.Value.ToString()
                If id_urgencia2 = 3 Or id_urgencia2 = 14 Then 'Afectación a la salud , Elemento extraño en el pedido
                    salud = 1
                End If
            End If
        Next


        'If cboaccion.SelectedIndex <> 0 Then
        '    id_accion = cboaccion.SelectedValue.ToString
        'Else
        '    lblmensaje_validacion.Text = "Debe seleccionar una *Acción* que realizó para dar solución al caso."
        '    ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
        '    Return
        'End If

        detalle_accion = txtobsadmin_unidad.Text.ToString()
        'producto_servicio = cboproducto_servicio.SelectedValue.ToString()
        obs_legal = txtobslegal.Text.ToString()
        obs_gerente_marca = txtobsGerenteMarca.Text.ToString()


        Dim dt_adjuntos1 As New DataTable()
        Dim cantidad_archivos As Integer
        dt_adjuntos1 = Session("dt_adjuntos")

        Dim id_estado1 As Integer
        id_estado1 = Integer.Parse(lblidestado.Text.ToString())
        If dt_adjuntos1 Is Nothing And id_estado1 = 1 Then 'la validacion es cuando esta en estado pendiente

            If chkmanual.Checked = True Then
                lblmensaje_validacion.Text = "Debe ingresar la carta manual que enviará al cliente. Ya que el proceso de atención es manual."
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                Return
            End If
        End If


        'Validamos si es numerico el valor ingresado
        Dim x As Integer = 0
        Dim s As String = txtcosto_accion.Text.ToString()
        Dim result As Boolean = Decimal.TryParse(s, x)

        If result = False Then
            lblmensaje_validacion.Text = "Debe ingresar un valor numérico en el campo de costo."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
            Return
        Else
            costo_accion = Decimal.Parse(txtcosto_accion.Text.ToString())
        End If

        x = 0
        s = txtotros_gastos.Text.ToString()
        result = Decimal.TryParse(s, x)

        If result = False Then
            lblmensaje_validacion.Text = "Debe ingresar un valor numérico en el campo de otros gastos."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
            Return
        Else
            otros_gastos = Decimal.Parse(txtotros_gastos.Text.ToString())
        End If



        Dim can4 As Integer = 0
        Dim itemz As WebControls.ListItem
        For Each itemz In chklistProductoServicio.Items
            If itemz.Selected = True Then
                can4 = can4 + 1
            End If
        Next
        If can4 = 0 Then
            lblmensaje_validacion.Text = "Debe seleccionar alguna opción dentro de la selección de Producto o Servicio."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
            Return
        End If


        Dim can3 As Integer = 0
        Dim marco_otros As Integer = 0
        Dim itemy As WebControls.ListItem
        For Each itemy In chklistAccion.Items
            If itemy.Selected = True Then
                can3 = can3 + 1
            End If
        Next
        If can3 = 0 Then
            lblmensaje_validacion.Text = "Debe seleccionar alguna opción dentro del listado de acciones."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
            Return
        End If


        If Trim(txtotraaccion.Text) <> "" Then
            can3 = 0
            For Each itemy In chklistAccion.Items
                If itemy.Value.ToString = "9" Then
                    If itemy.Selected = False Then
                        can3 = can3 + 1
                        Exit For
                    End If

                End If
            Next

            If can3 = 1 Then
                lblmensaje_validacion.Text = "Al ingresar otra acción realizada, debe  marcar la casilla de acciones Otros."
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                Return
            End If
        Else
            For Each itemy In chklistAccion.Items
                If itemy.Value.ToString = "9" Then
                    If itemy.Selected = True Then
                        If Trim(txtotraaccion.Text) = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar el texto en la opción otra acción realizada."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        End If
                        Exit For
                    End If

                End If
            Next

        End If


        If Trim(txtotromotivo.Text) <> "" Then
            can3 = 0
            For Each itemy In chklisturgencia.Items
                If itemy.Value.ToString = "35" Then
                    If itemy.Selected = False Then
                        can3 = can3 + 1
                        Exit For
                    End If

                End If
            Next

            If can3 = 1 Then
                lblmensaje_validacion.Text = "Al ingresar otro motivo realizado, debe  marcar la casilla de acciones Otros."
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                Return
            End If
        Else
            For Each itemy In chklisturgencia.Items
                If itemy.Value.ToString = "35" Then
                    If itemy.Selected = True Then
                        If Trim(txtotromotivo.Text) = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar el texto en la opción otro motivo realizado."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        End If
                        Exit For
                    End If

                End If
            Next

        End If



        Dim can1 As Integer = 0
        Dim item1 As ListItem
        For Each item1 In chklisturgencia.Items
            If item1.Selected = True Then
                'codigo para agregar registro
                can1 = can1 + 1
            End If
        Next
        If can1 = 0 Then
            lblmensaje_validacion.Text = "Debe seleccionar alguna opción dentro del listado de Urgencias."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
            Return
        End If



        Dim can2 As Integer = 0
        Dim itemx As WebControls.ListItem
        For Each itemx In chklistempreclamo.Items
            If itemx.Selected = True Then
                can2 = can2 + 1
            End If
        Next
        If can2 = 0 Then
            lblmensaje_validacion.Text = "Debe seleccionar alguna opción dentro del listado de empresas."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
            Return
        End If

        'Validacion de acciones - Cortesias
        Dim registra_cortesias As String
        Dim id_accion_select_c As Integer
        Dim itemc As ListItem
        registra_cortesias = "0"
        'Si se le dio cortesia
        Dim importecortesia As Decimal
        'No se le dio cortesia
        Dim modo As String
        Dim canpersonas As Integer
        Dim montomaximo As Decimal
        Dim validohasta As DateTime
        Dim coodinarContacto As String
        Dim correoCortesia As String
        Dim si_no_cortesia As String
        Dim lugar_cortesia As String

        For Each itemc In chklistAccion.Items
            If itemc.Selected = True Then
                'codigo para agregar registro
                id_accion_select_c = itemc.Value.ToString()
                If id_accion_select_c = 5 Then 'Cortesías

                    If chkopcion_cortesia1.Checked = False And chkopcion_cortesia2.Checked = False Then
                        lblmensaje_validacion.Text = "Debe elegir una opción para el tema de cortesías. Si brindó cortesía o No."
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                        Return
                    End If

                    If chkopcion_cortesia1.Checked = True And chkopcion_cortesia2.Checked = True Then
                        lblmensaje_validacion.Text = "Solo debe elegir una opción para el tema de cortesías. Si brindó cortesía o No."
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                        Return
                    End If


                    If chkopcion_cortesia1.Checked = True Then
                        x = 0
                        s = txtimportecortesia.Text.ToString()
                        result = Decimal.TryParse(s, x)

                        If result = False Then
                            lblmensaje_validacion.Text = "Debe ingresar un valor numérico en el campo de importe de cortesía."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        Else
                            importecortesia = Decimal.Parse(txtimportecortesia.Text.ToString())
                        End If

                        modo = "0"
                        si_no_cortesia = "SI"
                        registra_cortesias = "1"
                        correoCortesia = ""
                        coodinarContacto = ""
                        lugar_cortesia = ""
                        validohasta = DateTime.Parse("01/01/1990")
                    End If

                    If chkopcion_cortesia2.Checked = True Then
                        If cbomodo.SelectedIndex = 0 Then
                            lblmensaje_validacion.Text = "Debe seleccionar el modo para la cortesía."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        Else
                            modo = cbomodo.SelectedValue.ToString()
                        End If

                        If txtlugarcortesia.Text = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar el lugar donde aplicará la cortesía."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        Else
                            lugar_cortesia = txtlugarcortesia.Text.ToString()
                        End If


                        If txtcanpersonas.Text = "" Or txtcanpersonas.Text = "0" Then
                            canpersonas = 0
                        Else
                            x = 0
                            s = txtcanpersonas.Text.ToString()
                            result = Integer.TryParse(s, x)

                            If result = False Then
                                lblmensaje_validacion.Text = "Debe ingresar un valor numérico válido en el campo de cantidad de personas para la cortesía."
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                                Return
                            Else
                                canpersonas = Integer.Parse(txtcanpersonas.Text)
                            End If
                        End If


                        If txtmontomaximo.Text = "" Or txtmontomaximo.Text = "0" Then
                            montomaximo = 0
                        Else
                            x = 0
                            s = txtmontomaximo.Text.ToString()
                            result = Decimal.TryParse(s, x)

                            If result = False Then
                                lblmensaje_validacion.Text = "Debe ingresar un monto máximo para el importe de cortesía."
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                                Return
                            Else
                                montomaximo = Decimal.Parse(txtmontomaximo.Text.ToString())
                            End If
                        End If


                        If txtvalidohasta.Text = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar hasta que fecha es válido la cortesía."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        Else
                            validohasta = DateTime.Parse(txtvalidohasta.Text)
                        End If

                        If txtcoodinarContacto.Text = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar el nombre de la persona de contácto."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        Else
                            coodinarContacto = txtcoodinarContacto.Text
                        End If

                        If txtcorreoCortesia.Text = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar el correo de la persona de contácto."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        Else
                            correoCortesia = txtcorreoCortesia.Text
                        End If

                        registra_cortesias = "1"
                        si_no_cortesia = "NO"
                    End If

                End If
            End If
        Next

        Dim atencion_manual As String
        If chkmanual.Checked = True Then
            atencion_manual = "1"
        Else
            atencion_manual = "0"
        End If

        Dim id_documento_flujo As Integer
        id_documento_flujo = obj.Registra_Documento_flujo(nro_documento, user_ad, detalle_accion,
                                                 costo_accion, obs_legal, obs_gerente_marca, id_perfil,
                                                  aprobado, otros_gastos, atencion_manual)
        'Llenar las otras 2 tablas

        Dim id_motivo_select As Integer
        Dim id_urgencia As Integer
        Dim item As ListItem
        For Each item In chklisturgencia.Items
            If item.Selected = True Then
                'codigo para agregar registro
                id_urgencia = item.Value.ToString()
                Dim otro_motivo As String = ""
                If id_urgencia = 35 Then 'Siempre sera el ID de OTROS
                    otro_motivo = txtotromotivo.Text.ToString()
                End If
                obj.Registra_Detalle_Urgencia_Flujo(id_documento_flujo, id_urgencia, otro_motivo)
            End If
        Next


        Dim id_emp_reclamo As Integer
        Dim item2 As ListItem
        For Each item2 In chklistempreclamo.Items
            If item2.Selected = True Then
                'codigo para agregar registro
                id_emp_reclamo = item2.Value.ToString()

                obj.Registra_Detalle_emp_reclamo_Flujo(id_documento_flujo, id_emp_reclamo)
            End If
        Next

        Dim id_prod_serv As Integer
        Dim item3 As ListItem
        For Each item3 In chklistProductoServicio.Items
            If item3.Selected = True Then
                'codigo para agregar registro
                id_prod_serv = item3.Value.ToString()

                obj.Registra_Detalle_prodserv_Flujo(id_documento_flujo, id_prod_serv)
            End If
        Next



        Dim id_accion_select As Integer
        Dim item4 As ListItem
        For Each item4 In chklistAccion.Items
            If item4.Selected = True Then
                'codigo para agregar registro
                id_accion_select = item4.Value.ToString()
                Dim otra_accion As String = ""
                If id_accion_select = 9 Then 'Siempre sera el ID de OTROS
                    otra_accion = txtotraaccion.Text.ToString()
                End If
                obj.Registra_Detalle_accion_Flujo(id_documento_flujo, id_accion_select, otra_accion)
            End If
        Next


        'If registra_cortesias = "1" Then
        'Registro datos de cortesía
        Dim fecha_bl As DateTime
        fecha_bl = DateTime.Parse("01/01/1990")
        If registra_cortesias = "0" Then
            validohasta = fecha_bl
            si_no_cortesia = ""
            modo = ""
            coodinarContacto = ""
            correoCortesia = ""
            lugar_cortesia = ""
        End If

        obj.Registra_Datos_Cortesia(nro_documento, si_no_cortesia, importecortesia, modo, canpersonas, montomaximo, validohasta, coodinarContacto, correoCortesia, registra_cortesias, lugar_cortesia)

        'End If


        'Guardar Archivos Adjuntos - Solo guardar cuando el estado es en PENDIENTE EN ATENCION : 1

        Dim id_estado As Integer
        id_estado = lblidestado.Text
        If id_estado = 1 Then 'Estado Innicial - Pendiente de atencion

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

                    obj.Registra_Archivos_Adjuntos_Flujo(id_documento_flujo, nombre_archivo)
                Next
            End If


        End If

        If id_estado = 5 Then 'Vuelve a adjuntar los archivos - Rechazado por el area legal

            Dim j As Integer = 0
            Dim Adjuntos_Cliente As String
            Dim Adjuntos_Cliente_Temp As String
            'Ubicaciones
            Adjuntos_Cliente = ConfigurationManager.AppSettings("Adjuntos_Cliente")
            Adjuntos_Cliente_Temp = ConfigurationManager.AppSettings("Adjuntos_Cliente_Temp")

            Dim dt_adjuntos As New DataTable()
            dt_adjuntos = Session("dt_adjuntos_estado_rechazado")
            'eliminar archivos del documento para ser reemplazados por los nuevos, eso pasa en estado: Rechazado por el area legal.
            obj.Eliminar_Todos_Archivos_Adjuntos_Flujo(nro_documento)
            If dt_adjuntos IsNot Nothing Then
                For j = 0 To dt_adjuntos.Rows.Count - 1
                    Dim nombre_archivo As String
                    nombre_archivo = dt_adjuntos.Rows(j)(1).ToString()

                    'Copio archivo a la carpeta principal
                    My.Computer.FileSystem.CopyFile(Adjuntos_Cliente_Temp + "\" + nombre_archivo, Adjuntos_Cliente + "\" + nombre_archivo, True)
                    'Eliminar archivo de la ubicacion origen
                    My.Computer.FileSystem.DeleteFile(Adjuntos_Cliente_Temp + "\" + nombre_archivo)

                    obj.Registra_Archivos_Adjuntos_Flujo(id_documento_flujo, nombre_archivo)
                Next
            End If


        End If




        'Generación del archivo PDF - carta al cliente
        'Se dejara la carta al cliente en la ubicacion indicada para que el servicio pueda leerlo por numero de documento y enviarlo al cliente

        If id_estado = 3 And atencion_manual <> "1" Then ' El gerente de marca aprueba y genera el archivo pdf carta al cliente
            Dim carta_principal As String
            carta_principal = "1"
            Generar_PDF(carta_principal)
        End If

        Enviar_Correo_Cambio_Estado()

        Response.Redirect("~/admin_documentos.aspx")

        Return
    End Sub
    Protected Sub btnrechazar_Click(sender As Object, e As EventArgs) Handles btnrechazar.Click
        Dim aprobado As Integer = 2
        Registro_Aprobar_Rechazar(aprobado)
    End Sub
    Protected Sub btnregresar_Click(sender As Object, e As EventArgs) Handles btnregresar.Click
        Response.Redirect("~/admin_documentos.aspx")
    End Sub

    Public Sub Generar_PDF(ByVal carta_cliente As String)
        Dim txt As New List(Of String)
        Dim sw As New StringWriter()
        Dim html As String = sw.ToString()
        Dim Doc As New iTextSharp.text.Document()
        Dim nro_documento As String
        Dim año1 As String
        Dim mes1 As String
        Dim dia1 As String
        Dim Hor1 As String
        Dim min1 As String
        Dim seg1 As String

        año1 = Date.Now.Year.ToString()
        mes1 = Date.Now.Month.ToString()
        dia1 = Date.Now.Day.ToString()
        Hor1 = Date.Now.Hour.ToString()
        min1 = Date.Now.Minute.ToString()
        seg1 = Date.Now.Second.ToString()

        Dim Ruta_Descargas As String = ConfigurationManager.AppSettings("Ruta_Documentos") 'PARA LA  CARTA PRINCIPAL DEL CLIENTE
        Dim Ruta_Descargas_temp As String = ConfigurationManager.AppSettings("Adjuntos_Cliente_Temp") ' CARTA DE VISTA PREELIMINAR

        If carta_cliente = "1" Then
            Ruta_Descargas = Ruta_Descargas 'CARTA  PRINCIPAL DEL CLIENTE
            nro_documento = lblnrodocumento.Text.ToString()
        Else
            Ruta_Descargas = Ruta_Descargas_temp ' CARTA DE VISTA PREELIMINAR
            nro_documento = año1 + mes1 + dia1 + Hor1 + min1 + seg1 + "_" + lblnrodocumento.Text.ToString()
        End If


        PdfWriter.GetInstance(Doc, New FileStream(Ruta_Descargas & "\" + nro_documento + ".pdf", FileMode.Create))
        Doc.Open()



        Dim table01 As New PdfPTable(1)
        table01.WidthPercentage = 100

        Dim table02 As New PdfPTable(1)
        table02.WidthPercentage = 100

        Dim table03 As New PdfPTable(1)
        table03.WidthPercentage = 100

        Dim table04 As New PdfPTable(1)
        table04.WidthPercentage = 100

        Dim table05 As New PdfPTable(1)
        table05.WidthPercentage = 100

        Dim table06 As New PdfPTable(1)
        table06.WidthPercentage = 100

        Dim table07 As New PdfPTable(1)
        table07.WidthPercentage = 100

        Dim table08 As New PdfPTable(1)
        table08.WidthPercentage = 100

        Dim table09 As New PdfPTable(1)
        table09.WidthPercentage = 100

        Dim table09_1 As New PdfPTable(1)
        table09_1.WidthPercentage = 100

        Dim table09_2 As New PdfPTable(1)
        table09_2.WidthPercentage = 100

        Dim table10 As New PdfPTable(1)
        table10.WidthPercentage = 100

        Dim table10_1 As New PdfPTable(1)
        table10_1.WidthPercentage = 100

        Dim table10_2 As New PdfPTable(1)
        table10_2.WidthPercentage = 100

        Dim table11 As New PdfPTable(1)
        table11.WidthPercentage = 100

        Dim table12 As New PdfPTable(1)
        table12.WidthPercentage = 100

        Dim table13 As New PdfPTable(1)
        table13.WidthPercentage = 100

        Dim table14 As New PdfPTable(1)
        table14.WidthPercentage = 100

        Dim table15 As New PdfPTable(1)
        table15.WidthPercentage = 100


        '---Fhecha de sistema---
        Dim año As String
        Dim mes As String
        Dim dia As String
        Dim fecha_sistema As String

        año = Now.Year.ToString()
        mes = Now.Month.ToString()
        dia = Now.Day.ToString()

        If dia.Length = 1 Then
            dia = "0" + dia
        End If
        If mes.Length = 1 Then
            mes = "0" + mes
        End If
        fecha_sistema = dia + "/" + mes + "/" + año

        If lblfecha_aprobacion_gerente.Text <> "" Then '---04/10/2022
            fecha_sistema = lblfecha_aprobacion_gerente.Text.ToString()
        End If

        Dim cell_tb01_01 As New PdfPCell(New iTextSharp.text.Phrase("Lima, " + fecha_sistema, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb01_01.Colspan = 1
        cell_tb01_01.HorizontalAlignment = 2
        cell_tb01_01.VerticalAlignment = 1
        cell_tb01_01.Border = 0

        table01.AddCell(cell_tb01_01)


        Dim cell_tb02_01 As New PdfPCell(New iTextSharp.text.Phrase("Sr(a).", New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb02_01.Colspan = 1
        cell_tb02_01.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb02_01.VerticalAlignment = 1
        cell_tb02_01.Border = 0

        table02.AddCell(cell_tb02_01)

        Dim cell_tb03_01 As New PdfPCell(New iTextSharp.text.Phrase(lblnomcliente.Text, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb03_01.Colspan = 1
        cell_tb03_01.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb03_01.VerticalAlignment = 1
        cell_tb03_01.Border = 0

        table03.AddCell(cell_tb03_01)

        Dim cell_tb04_01 As New PdfPCell(New iTextSharp.text.Phrase(lbldireccioncliente.Text, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb04_01.Colspan = 1
        cell_tb04_01.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb04_01.VerticalAlignment = 1
        cell_tb04_01.Border = 0

        Dim cell_tb04_02 As New PdfPCell(New iTextSharp.text.Phrase(lblemail_cliente.Text, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb04_02.Colspan = 1
        cell_tb04_02.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb04_02.VerticalAlignment = 1
        cell_tb04_02.Border = 0

        table04.AddCell(cell_tb04_01)
        table04.AddCell(cell_tb04_02)

        Dim cell_tb05_01 As New PdfPCell(New iTextSharp.text.Phrase("Presente:", New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.UNDERLINE)))
        cell_tb05_01.Colspan = 1
        cell_tb05_01.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb05_01.VerticalAlignment = 1
        cell_tb05_01.Border = 0

        table05.AddCell(cell_tb05_01)

        Dim cell_tb06_01 As New PdfPCell(New iTextSharp.text.Phrase("Referencia: Hoja  de Reclamación N° " + lblnrodocumento.Text, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.UNDERLINE)))
        cell_tb06_01.Colspan = 1
        cell_tb06_01.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb06_01.VerticalAlignment = 1
        cell_tb06_01.Border = 0

        table06.AddCell(cell_tb06_01)

        Dim cell_tb07_01 As New PdfPCell(New iTextSharp.text.Phrase("Estimado(a) Cliente,", New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb07_01.Colspan = 1
        cell_tb07_01.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb07_01.VerticalAlignment = 1
        cell_tb07_01.Border = 0

        table07.AddCell(cell_tb07_01)

        'Listado de Producto Servicio
        Dim lista_prod_serv As String = ""
        Dim valor2 As String = ""
        Dim can2 As Integer = 0
        Dim item2 As ListItem
        For Each item2 In chklistProductoServicio.Items
            If item2.Selected = True Then
                can2 = can2 + 1 'Cantidad de check
            End If
        Next
        For Each item2 In chklistProductoServicio.Items
            If item2.Selected = True Then
                If can2 = 1 Then
                    valor2 = item2.Text.ToString
                    lista_prod_serv = valor2
                    Exit For
                Else
                    valor2 = item2.Text.ToString
                    lista_prod_serv = valor2 + " y " + lista_prod_serv
                End If

            End If
        Next
        If can2 = 1 Then
            lista_prod_serv = " " + valor2
        Else
            lista_prod_serv = " " + lista_prod_serv.Substring(0, lista_prod_serv.Length - 3)
        End If


        Dim texto1 As String = "Reciba nuestro más cordial saludo, el motivo de la actual carta es darle respuesta al reclamo y/o queja registrado en el libro de reclamaciones, Nro. " _
                                + lblnrodocumento.Text + " sobre nuestro" + lista_prod_serv + "."

        Dim cell_tb08_01 As New PdfPCell(New iTextSharp.text.Phrase(texto1, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb08_01.Colspan = 1
        cell_tb08_01.HorizontalAlignment = 3 '0=Left, 1=Centre, 2=Right
        cell_tb08_01.VerticalAlignment = 1
        cell_tb08_01.Border = 0

        table08.AddCell(cell_tb08_01)


        Dim fechaConsumo As DateTime = DateTime.Parse(lblfecharegistro.Text.ToString())
        Dim fechaConsumo_txt As String
        año = fechaConsumo.Year.ToString()
        mes = fechaConsumo.Month.ToString()
        dia = fechaConsumo.Day.ToString()

        If dia.Length = 1 Then
            dia = "0" + dia
        End If
        If mes.Length = 1 Then
            mes = "0" + mes
        End If
        fechaConsumo_txt = dia + "/" + mes + "/" + año


        'Listado de Urgencias elegidas
        'Inicio 07/09/2021
        Dim lista_ugencias As String = ""
        Dim valor As String = ""
        Dim id_urgencia As String
        Dim can1 As Integer = 0
        Dim item1 As ListItem
        For Each item1 In chklisturgencia.Items
            If item1.Selected = True Then

                can1 = can1 + 1 'Cantidad de check
            End If
        Next
        Dim conta1 As Integer = 0
        For Each item1 In chklisturgencia.Items
            If item1.Selected = True Then
                If can1 = 1 Then
                    valor = item1.Text.ToString


                    '07/09/2021 busca su descripcion para el cliente nueva columna y reeemplazamos parametro valor
                    Dim dt1 As New DataTable
                    Dim dt2 As New DataTable
                    id_urgencia = item1.Value.ToString()
                    dt2 = Session("tab_urgencias_etiquetas")
                    Dim qlWhere As String = "id_urgencia = '" + id_urgencia + "'"
                    dt1 = dt2.Select(qlWhere).CopyToDataTable
                    valor = dt1.Rows(0)(1).ToString()
                    '07/09/2021
                    If valor = "Otros" Then
                        valor = txtotromotivo.Text.ToString()
                    End If

                    lista_ugencias = valor
                    Exit For
                Else
                    valor = item1.Text.ToString

                    '07/09/2021 busca su descripcion para el cliente nueva columna y reeemplazamos parametro valor
                    Dim dt1 As New DataTable
                    Dim dt2 As New DataTable
                    id_urgencia = item1.Value.ToString()
                    dt2 = Session("tab_urgencias_etiquetas")
                    Dim qlWhere As String = "id_urgencia = '" + id_urgencia + "'"
                    dt1 = dt2.Select(qlWhere).CopyToDataTable
                    valor = dt1.Rows(0)(1).ToString()

                    If valor = "Otros" Then
                        valor = txtotromotivo.Text.ToString()
                    End If

                    '07/09/2021
                    If conta1 = can1 - 1 Then
                        lista_ugencias = lista_ugencias.Substring(0, lista_ugencias.Length - 2) + " y " + valor

                    Else
                        lista_ugencias = valor + "; " + lista_ugencias
                    End If

                    conta1 = conta1 + 1
                    'lista_ugencias = valor + ", " + lista_ugencias

                End If

            End If
        Next
        ' lista_ugencias = lista_ugencias.Substring(0, lista_ugencias.Length - 1)


        If can1 = 1 Then
            lista_ugencias = " " + valor
        Else
            lista_ugencias = " " + lista_ugencias.Substring(0, lista_ugencias.Length)
        End If


        'Fin 07/09/21

        Dim texto2 As String = "Sobre el particular, lamentamos la experiencia poco placentera vivida el " _
                                + fechaConsumo_txt + " debido a que se habría presentado por lo siguiente:" + lista_ugencias + "." '_
        '+ " motivo por el cual decidió manifestar su reclamo."

        'Dim text2_1 As String = lista_ugencias
        'Dim text2_2 As String = "Motivo por el cual decidió manifestar su reclamo."

        Dim cell_tb09_01 As New PdfPCell(New iTextSharp.text.Phrase(texto2, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb09_01.Colspan = 1
        cell_tb09_01.HorizontalAlignment = 3 '0=Left, 1=Centre, 2=Right
        cell_tb09_01.VerticalAlignment = 1
        cell_tb09_01.Border = 0

        table09.AddCell(cell_tb09_01)


        'bucle inicio
        'Dim Valor_item As String
        'For Each item1 In chklisturgencia.Items
        '    If item1.Selected = True Then
        '        Valor_item = item1.Text.ToString
        '        Dim cell_tb09_01_1 As New PdfPCell(New iTextSharp.text.Phrase("- " + Valor_item, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        '        cell_tb09_01_1.Colspan = 1
        '        cell_tb09_01_1.VerticalAlignment = 1
        '        cell_tb09_01_1.Border = 0

        '        table09_1.AddCell(cell_tb09_01_1)
        '    End If
        'Next
        ''bucle fin



        'Dim cell_tb09_01_2 As New PdfPCell(New iTextSharp.text.Phrase(text2_2, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        'cell_tb09_01_2.Colspan = 1
        'cell_tb09_01_2.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        'cell_tb09_01_2.VerticalAlignment = 1
        'cell_tb09_01_2.Border = 0

        'table09_2.AddCell(cell_tb09_01_2)



        'listado de acciones - principal

        'Ini 07/09/2021
        Dim lista_acciones As String = ""
        Dim valor3 As String = ""
        Dim id_accion As String
        Dim can3 As Integer = 0
        Dim item3 As ListItem
        For Each item3 In chklistAccion.Items
            If item3.Selected = True Then
                can3 = can3 + 1 'Cantidad de check
            End If
        Next

        Dim conta As Integer = 0
        'Dim id_accionx As Integer

        For Each item3 In chklistAccion.Items
            If item3.Selected = True Then
                If can3 = 1 Then
                    valor3 = item3.Text.ToString

                    '07/09/2021 busca su descripcion para el cliente nueva columna y reeemplazamos parametro valor
                    Dim dt1 As New DataTable
                    Dim dt2 As New DataTable
                    id_accion = item3.Value.ToString()
                    dt2 = Session("tab_acciones_etiquetas")
                    Dim qlWhere As String = "id_accion = '" + id_accion + "'"
                    dt1 = dt2.Select(qlWhere).CopyToDataTable
                    valor3 = dt1.Rows(0)(1).ToString()
                    '07/09/2021
                    If valor3 = "Otros" Then
                        valor3 = txtotraaccion.Text.ToString()
                    Else '22/03/22022
                        If id_accion = 5 And chkopcion_cortesia2.Checked = True Then
                            valor3 = valor3 + "(1)"
                        End If
                    End If
                    lista_acciones = valor3
                    Exit For
                Else
                    valor3 = item3.Text.ToString

                    'Ordenar una accion adelante 23/06/2023

                    'Dim texto_id6 As String
                    'id_accionx = item3.Value.ToString()


                    '07/09/2021 busca su descripcion para el cliente nueva columna y reeemplazamos parametro valor
                    Dim dt1 As New DataTable
                    Dim dt2 As New DataTable
                    id_accion = item3.Value.ToString()
                    dt2 = Session("tab_acciones_etiquetas")
                    Dim qlWhere As String = "id_accion = '" + id_accion + "'"
                    dt1 = dt2.Select(qlWhere).CopyToDataTable
                    valor3 = dt1.Rows(0)(1).ToString()

                    '07/09/2021
                    If valor3 = "Otros" Then
                        valor3 = txtotraaccion.Text.ToString()
                    End If
                    '07/09/2021
                    'lista_acciones = valor3 + ", " + lista_acciones
                    If conta = can3 - 1 Then
                        'lista_acciones = lista_acciones.Substring(0, lista_acciones.Length - 2) + " y " + valor3
                        lista_acciones = lista_acciones.Substring(0, lista_acciones.Length) + " y " + valor3 '23062023
                    Else
                        'If id_accionx = 6 Then
                        '    texto_id6 = valor3
                        'Else
                        'lista_acciones = valor3 + "; " + lista_acciones
                        'End If
                        If conta = 0 Then
                            lista_acciones = valor3
                        Else
                            lista_acciones = lista_acciones + "; " + valor3
                        End If


                    End If

                    conta = conta + 1

                End If

            End If
        Next

        If can3 = 1 Then
            lista_acciones = " " + valor3
        Else
            'If id_accionx = 6 Then
            '    lista_acciones = valor3 + "; " + lista_acciones
            'Else 'Proceso normal
            lista_acciones = " " + lista_acciones.Substring(0, lista_acciones.Length)
            'End If

        End If

        'Fin 07/09/2021

        Dim texto3 As String = "Ante esta situación, nuestro personal, con la finalidad de subsanar la molestia generada realizó la(s) siguiente(s) accion(es):" _
                                + lista_acciones + "."

        Dim cell_tb10_01 As New PdfPCell(New iTextSharp.text.Phrase(texto3, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb10_01.Colspan = 1
        cell_tb10_01.HorizontalAlignment = 3 '0=Left, 1=Centre, 2=Right
        cell_tb10_01.VerticalAlignment = 1
        cell_tb10_01.Border = 0

        table10.AddCell(cell_tb10_01)


        'Dim texto3 As String = "Al respecto, ante esta situación el administrador del restaurante realizó la(s) siguiente(s) accion(es): "
        'Dim text3_2 As String = "Con la finalidad de subsanar la molestia ocasionada narrada en el párrafo precedente."


        'Dim cell_tb10_01 As New PdfPCell(New iTextSharp.text.Phrase(texto3, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        'cell_tb10_01.Colspan = 1
        'cell_tb10_01.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        'cell_tb10_01.VerticalAlignment = 1
        'cell_tb10_01.Border = 0

        'table10.AddCell(cell_tb10_01)


        'Dim Valor_item2 As String
        'For Each item1 In chklistAccion.Items
        '    If item1.Selected = True Then
        '        Valor_item2 = item1.Text.ToString
        '        Dim cell_tb10_01_1 As New PdfPCell(New iTextSharp.text.Phrase("- " + Valor_item2, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        '        cell_tb10_01_1.Colspan = 1
        '        cell_tb10_01_1.VerticalAlignment = 1
        '        cell_tb10_01_1.Border = 0

        '        table10_1.AddCell(cell_tb10_01_1)
        '    End If
        'Next



        'Dim cell_tb10_01_2 As New PdfPCell(New iTextSharp.text.Phrase(text3_2, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        'cell_tb10_01_2.Colspan = 1
        'cell_tb10_01_2.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        'cell_tb10_01_2.VerticalAlignment = 1
        'cell_tb10_01_2.Border = 0

        'table10_2.AddCell(cell_tb10_01_2)


        'Listado de aaciones

        Dim texto4 As String = "Por último agradecemos que nos haya hecho llegar su reclamo, puesto que ello nos ayuda a mejorar nuestro servicio" _
                                + " y cuidar los detalles, de igual forma reafirmamos nuestro compromiso de brindarle la mejor experiencia y productos de calidad."

        Dim cell_tb11_01 As New PdfPCell(New iTextSharp.text.Phrase(texto4, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb11_01.Colspan = 1
        cell_tb11_01.HorizontalAlignment = 3 '0=Left, 1=Centre, 2=Right
        cell_tb11_01.VerticalAlignment = 1
        cell_tb11_01.Border = 0

        table11.AddCell(cell_tb11_01)


        Dim texto5 As String = "Sin otro particular, quedamos de usted."

        Dim cell_tb12_01 As New PdfPCell(New iTextSharp.text.Phrase(texto5, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb12_01.Colspan = 1
        cell_tb12_01.HorizontalAlignment = 3 '0=Left, 1=Centre, 2=Right
        cell_tb12_01.VerticalAlignment = 1
        cell_tb12_01.Border = 0

        table12.AddCell(cell_tb12_01)

        Dim texto6 As String = "Atentamente."

        Dim cell_tb13_01 As New PdfPCell(New iTextSharp.text.Phrase(texto6, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb13_01.Colspan = 1
        cell_tb13_01.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb13_01.VerticalAlignment = 1
        cell_tb13_01.Border = 0

        table13.AddCell(cell_tb13_01)


        Dim gerente_marca As String
        If carta_cliente = "1" Then 'VISTA PREELIMINAR DE LA CARTA
            gerente_marca = Session("dato_completo").ToString()
        Else
            gerente_marca = "[Documento Pendiente]"
        End If

        Dim cell_tb14_01 As New PdfPCell(New iTextSharp.text.Phrase(gerente_marca, New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb14_01.Colspan = 1
        cell_tb14_01.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        cell_tb14_01.VerticalAlignment = 1
        cell_tb14_01.Border = 0

        table14.AddCell(cell_tb14_01)


        Dim cell_tb15_01 As New PdfPCell(New iTextSharp.text.Phrase("Gerente de Marca", New Font(Font.FontFamily.TIMES_ROMAN, 12.0F, Font.NORMAL)))
        cell_tb15_01.Colspan = 1
        cell_tb15_01.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        cell_tb15_01.VerticalAlignment = 1
        cell_tb15_01.Border = 0

        table15.AddCell(cell_tb15_01)

        '----PIE DE PAGINA-----

        'PIE DE PAGINA
        Dim id_accion_select_c As Integer
        Dim mensaje_pie_pagina As String = ""
        Dim modo As String
        Dim cantidad_pers_cortesia As String
        Dim monto_max_consumo As String
        Dim invitacion_valida As String
        Dim coordinar_con As String
        Dim correo_coordinar As String
        Dim lugar_cortesia As String
        Dim añox As String
        Dim mesx As String
        Dim diax As String
        For Each itemc In chklistAccion.Items
            If itemc.Selected = True Then
                'codigo para agregar registro
                id_accion_select_c = itemc.Value.ToString()
                If id_accion_select_c = 5 Then 'Cortesías
                    If chkopcion_cortesia2.Checked = True Then

                        modo = cbomodo.SelectedItem.Text.ToString()
                        cantidad_pers_cortesia = txtcanpersonas.Text.ToString()
                        monto_max_consumo = txtmontomaximo.Text.ToString()
                        invitacion_valida = txtvalidohasta.Text.ToString()
                        coordinar_con = txtcoodinarContacto.Text.ToString()
                        correo_coordinar = txtcorreoCortesia.Text.ToString()
                        lugar_cortesia = txtlugarcortesia.Text.ToString()

                        añox = invitacion_valida.Substring(0, 4)
                        mesx = invitacion_valida.Substring(5, 2)
                        diax = invitacion_valida.Substring(8, 2)
                        invitacion_valida = diax + "/" + mesx + "/" + añox

                        Dim texto_can_personas As String = ""
                        Dim texto_monto_maximo As String = ""

                        If cantidad_pers_cortesia <> "" And cantidad_pers_cortesia <> "0" Then
                            texto_can_personas = " / Cantidad de Personas: " + cantidad_pers_cortesia.ToString()
                        End If

                        If monto_max_consumo <> "" And monto_max_consumo <> "0" And monto_max_consumo <> "0.00" Then
                            texto_monto_maximo = " / Monto máximo de consumo S/." + monto_max_consumo.ToString()
                        End If

                        mensaje_pie_pagina = "(1) La invitación es válida para " + modo + texto_can_personas + " / Aplica: " + lugar_cortesia.ToString() + texto_monto_maximo.ToString() _
                                            + ". Cortesía válida hasta el " + invitacion_valida + ". Coordinar previamente con " + coordinar_con + " - " + correo_coordinar
                    End If
                End If
            End If
        Next


        Dim table16 As New PdfPTable(1)
        table16.WidthPercentage = 100
        Dim cell_tb16_01 As New PdfPCell(New iTextSharp.text.Phrase(mensaje_pie_pagina, New Font(Font.FontFamily.TIMES_ROMAN, 9.0F, Font.ITALIC)))
        cell_tb16_01.Colspan = 1
        cell_tb16_01.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb16_01.VerticalAlignment = 1
        cell_tb16_01.Border = 0
        table16.AddCell(cell_tb16_01)


        '------pie de pagina fin --------



        Dim c As New iTextSharp.text.Chunk("" & "" & vbLf & "", iTextSharp.text.FontFactory.GetFont("Verdana", 10))
        Dim p As New iTextSharp.text.Paragraph()
        p.Alignment = iTextSharp.text.Element.ALIGN_CENTER
        p.Add(c)

        ' Creamos la imagen y le ajustamos el tamaño
        Dim Logo As String = ConfigurationManager.AppSettings("Logo")
        Dim imagen As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(Logo.ToString())
        imagen.BorderWidth = 0
        imagen.Alignment = iTextSharp.text.Element.ALIGN_LEFT
        Dim percentage As Single = 0.0F
        percentage = 150 / imagen.Width
        imagen.ScalePercent(percentage * 50)

        Doc.Add(imagen)
        Doc.Add(table01)
        Doc.Add(p)
        Doc.Add(table02)
        Doc.Add(p)
        Doc.Add(table03)
        Doc.Add(table04)
        Doc.Add(p)
        Doc.Add(table05)
        Doc.Add(table06)
        Doc.Add(p)
        Doc.Add(p)
        Doc.Add(table07)
        Doc.Add(p)
        Doc.Add(table08)
        Doc.Add(p)
        Doc.Add(table09)
        Doc.Add(p)
        'Doc.Add(table09_1)
        'Doc.Add(p)
        'Doc.Add(table09_2)
        Doc.Add(table10)
        Doc.Add(p)
        'Doc.Add(table10_1)
        'Doc.Add(p)
        'Doc.Add(table10_2)
        Doc.Add(table11)
        Doc.Add(p)
        Doc.Add(table12)
        Doc.Add(p)
        Doc.Add(p)
        Doc.Add(p)
        Doc.Add(table13)
        Doc.Add(p)
        Doc.Add(p)
        Doc.Add(p)

        Doc.Add(table14)
        Doc.Add(table15)

        Doc.Add(p)
        Doc.Add(p)
        Doc.Add(p)
        Doc.Add(p)
        Doc.Add(p)
        Doc.Add(p)
        'PIE DE PAGINA
        Doc.Add(table16)


        Dim xmlReader As System.Xml.XmlTextReader = New System.Xml.XmlTextReader(New StringReader(html))
        Doc.Close()
        'Dim Path As String = Ruta_Descargas & "\Doc_" & nro_documento & ".pdf"

        Dim Path As String = Ruta_Descargas & "\" + nro_documento + ".pdf"
        txt.Add(Path)


        If carta_cliente = "0" Then 'VISTA PREELIMINAR DE LA CARTA
            ShowPdf(txt)
        End If

    End Sub

    Public Sub ShowPdf(ByVal strS As List(Of String))
        Try

            Dim nom_archivo As String = [String].Format("Pdf_{0}.pdf", DateTime.Now.ToString("yyyy-MMM-dd-HHmmss"))

            For Each row In strS
                Try
                    Response.ClearContent()
                    Response.ClearHeaders()
                    Response.ContentType = "application/pdf"
                    'Response.AddHeader("Content-Disposition", "attachment; filename=" & "ActaVisita.pdf")
                    Response.AddHeader("Content-Disposition", "attachment; filename=" & nom_archivo)
                    Response.TransmitFile(row)
                    Response.Flush()
                    Response.Clear()
                Catch ex As Exception
                Finally
                    Response.SuppressContent = True
                    HttpContext.Current.ApplicationInstance.CompleteRequest()
                End Try
            Next
        Catch ex As Exception
        Finally
        End Try
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
                Dim archivo_general As String
                Dim ad_usuario As String
                ad_usuario = Session("user").ToString()

                archivo_general = ad_usuario + "_" + año + mes + dia + Hor + min + seg + "_" + inputGroupFile04.FileName

                inputGroupFile04.SaveAs(IPSERVER + "\" + archivo_general) 'Guardamos en Temporal antes de la principal

                If lblidestado.Text = "5" Then 'Si ya es rechazado por el area legal

                    If Session("dt_adjuntos_estado_rechazado") Is Nothing Then
                        dt_adjuntos.Rows.Add(inputGroupFile04.FileName.ToString(), archivo_general)
                        Session("dt_adjuntos_estado_rechazado") = dt_adjuntos
                    Else
                        dt_adjuntos = Session("dt_adjuntos_estado_rechazado")
                        dt_adjuntos.Rows.Add(inputGroupFile04.FileName.ToString(), archivo_general)
                        Session("dt_adjuntos_estado_rechazado") = dt_adjuntos
                    End If

                Else 'caso nuevo
                    If Session("dt_adjuntos") Is Nothing Then
                        dt_adjuntos.Rows.Add(inputGroupFile04.FileName.ToString(), archivo_general)
                        Session("dt_adjuntos") = dt_adjuntos
                    Else
                        dt_adjuntos = Session("dt_adjuntos")
                        dt_adjuntos.Rows.Add(inputGroupFile04.FileName.ToString(), archivo_general)
                        Session("dt_adjuntos") = dt_adjuntos
                    End If
                End If




                'Almacena en datatable - grilla

                grvadjuntos.DataSource = dt_adjuntos
                grvadjuntos.DataBind()

                'txtuno.Text = dt_adjuntos.Rows.Count.ToString()


            End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub btnadjuntos_Click(sender As Object, e As EventArgs) Handles btnadjuntos.Click
        ScriptManager.RegisterStartupScript(Me, GetType(Page), "alerta", "<script>popup()</script>", False)
    End Sub

    Protected Sub grvarchivosadjuntos_RowCommand(sender As Object, e As GridViewCommandEventArgs) Handles grvarchivosadjuntos.RowCommand
        Dim ID As Integer

        If e.CommandName = "adjunto" Then
            ID = (e.CommandArgument).ToString()
            For Each gvClaserow As GridViewRow In grvarchivosadjuntos.Rows
                Dim label_ruta As Label
                Dim label_id As Label

                label_ruta = gvClaserow.FindControl("lblarchivoadjunto")
                label_id = gvClaserow.FindControl("lblid")

                If ID.ToString() = label_id.Text.ToString() Then

                    Dim IPSERVER As String = ConfigurationManager.AppSettings("Adjuntos_Cliente")
                    Dim ruta As String
                    Dim nro_ruc As String
                    Dim nombre_archivo As String
                    nombre_archivo = label_ruta.Text
                    'nro_ruc = lblruc_prov.Text.ToString()

                    ruta = IPSERVER + "\" + nombre_archivo.ToString()

                    Dim prueba As String
                    HttpContext.Current.Response.Clear()
                    'HttpContext.Current.Response.ContentType = "txt"
                    prueba = Path.GetFileName(ruta).ToString()
                    HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment; filename=" + prueba)
                    HttpContext.Current.Response.TransmitFile(ruta)
                    HttpContext.Current.Response.End()
                End If
            Next
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
    Protected Sub btnCartaPreview_Click(sender As Object, e As EventArgs) Handles btnCartaPreview.Click
        Dim carta_principal As String
        carta_principal = "0" 'Sera una vista preeliminar

        'Validar que tenga información los controles

        Dim can_chek_1 As Integer
        Dim can_chek_2 As Integer
        Dim can_chek_3 As Integer
        Dim can_chek_4 As Integer

        If chkmanual.Checked = True Then
            lblmensaje_validacion.Text = "Si el documento será atendido de manera manual no podrá tener la vista preeliminar de la carta al cliente."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
            Return
        End If

        For Each item1 In chklistProductoServicio.Items
            If item1.Selected = True Then
                can_chek_1 = can_chek_1 + 1 'Cantidad de check
            End If
        Next
        If can_chek_1 = 0 Then
            lblmensaje_validacion.Text = "Debe seleccionar alguna opción dentro de la selección de Producto o Servicio."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
            Return
        End If


        For Each item2 In chklistAccion.Items
            If item2.Selected = True Then
                can_chek_2 = can_chek_2 + 1 'Cantidad de check
            End If
        Next
        If can_chek_2 = 0 Then
            lblmensaje_validacion.Text = "Debe seleccionar alguna opción dentro del listado de acciones."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
            Return
        End If

        Dim can4 As Integer
        If Trim(txtotraaccion.Text) <> "" Then
            can4 = 0
            For Each itemy In chklistAccion.Items
                If itemy.Value.ToString = "9" Then
                    If itemy.Selected = False Then
                        can4 = can4 + 1
                        Exit For
                    End If

                End If
            Next

            If can4 = 1 Then
                lblmensaje_validacion.Text = "Al ingresar otra acción realizada, debe  marcar la casilla de acciones Otros."
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                Return
            End If
        Else
            For Each itemy In chklistAccion.Items
                If itemy.Value.ToString = "9" Then
                    If itemy.Selected = True Then
                        If Trim(txtotraaccion.Text) = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar el texto en la opción otra acción realizada."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        End If
                        Exit For
                    End If

                End If
            Next
        End If


        For Each item3 In chklisturgencia.Items
            If item3.Selected = True Then
                can_chek_3 = can_chek_3 + 1 'Cantidad de check
            End If
        Next
        If can_chek_3 = 0 Then
            lblmensaje_validacion.Text = "Debe seleccionar alguna opción dentro del listado de Urgencias."
            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
            Return
        End If


        can4 = 0
        If Trim(txtotromotivo.Text) <> "" Then
            can4 = 0
            For Each itemy In chklisturgencia.Items
                If itemy.Value.ToString = "35" Then
                    If itemy.Selected = False Then
                        can4 = can4 + 1
                        Exit For
                    End If

                End If
            Next

            If can4 = 1 Then
                lblmensaje_validacion.Text = "Al ingresar otro motivo realizado, debe  marcar la casilla de acciones Otros."
                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                Return
            End If
        Else
            For Each itemy In chklisturgencia.Items
                If itemy.Value.ToString = "35" Then
                    If itemy.Selected = True Then
                        If Trim(txtotromotivo.Text) = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar el texto en la opción otro motivo realizado."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        End If
                        Exit For
                    End If

                End If
            Next
        End If

        '-- 22/03/2022

        Dim id_accion_select_c As Integer
        Dim x As Integer
        Dim s As String
        Dim Result As Decimal


        For Each itemc In chklistAccion.Items
            If itemc.Selected = True Then
                'codigo para agregar registro
                id_accion_select_c = itemc.Value.ToString()
                If id_accion_select_c = 5 Then 'Cortesías

                    If chkopcion_cortesia1.Checked = False And chkopcion_cortesia2.Checked = False Then
                        lblmensaje_validacion.Text = "Debe elegir una opción para el tema de cortesías. Si brindó cortesía o No."
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                        Return
                    End If

                    If chkopcion_cortesia1.Checked = True And chkopcion_cortesia2.Checked = True Then
                        lblmensaje_validacion.Text = "Solo debe elegir una opción para el tema de cortesías. Si brindó cortesía o No."
                        ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                        Return
                    End If


                    If chkopcion_cortesia2.Checked = True Then
                        If cbomodo.SelectedIndex = 0 Then
                            lblmensaje_validacion.Text = "Debe seleccionar el modo para la cortesía."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        End If

                        If txtcanpersonas.Text <> "" Then
                            x = 0
                            s = txtcanpersonas.Text.ToString()
                            Result = Integer.TryParse(s, x)

                            If Result = False Then
                                lblmensaje_validacion.Text = "Debe ingresar un valor numérico en el campo de cantidad de personas para la cortesía."
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                                Return
                            End If
                        End If


                        If txtmontomaximo.Text <> "" Then
                            x = 0
                            s = txtmontomaximo.Text.ToString()
                            Result = Decimal.TryParse(s, x)

                            If Result = False Then
                                lblmensaje_validacion.Text = "Debe ingresar un valor válido para el monto máximo para el importe de cortesía."
                                ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                                Return
                            End If
                        End If


                        If txtvalidohasta.Text = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar hasta que fecha es válido la cortesía."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        End If

                        If txtcoodinarContacto.Text = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar el nombre de la persona de contácto."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return
                        End If

                        If txtcorreoCortesia.Text = "" Then
                            lblmensaje_validacion.Text = "Debe ingresar el correo de la persona de contácto."
                            ScriptManager.RegisterStartupScript(Page, Page.GetType(), "modal-validacion", "$('#modal-validacion').modal();", True)
                            Return

                        End If
                    End If

                End If
            End If
        Next


        Generar_PDF(carta_principal)
    End Sub
    Protected Sub btnPDFCliente_Click(sender As Object, e As EventArgs) Handles btnPDFCliente.Click
        Dim nro_documento As String
        Dim id_unidad As Integer
        nro_documento = lblnrodocumento.Text.ToString()
        id_unidad = Integer.Parse(lblidunidad.Text.ToString())

        Generar_PDF(nro_documento, id_unidad)

    End Sub

    Public Sub Generar_pdf(ByVal nro_documento As String, ByVal id_unidad As Integer)
        Dim Ruta_Descargas As String = ConfigurationManager.AppSettings("Adjuntos_Cliente_Temp")
        Dim Logo As String = ConfigurationManager.AppSettings("Logo")

        Dim txt As New List(Of String)
        Dim año As String
        Dim mes As String
        Dim dia As String
        Dim dato_empresa As String
        Dim nombre As String
        Dim domicilio As String
        Dim dni As String
        Dim telefono As String
        Dim email As String
        Dim menor_edad As String
        Dim producto_servicio As String
        Dim producto_activo As String = ""
        Dim servicio_activo As String = ""
        Dim monto_reclamado As String
        Dim desc_bien_reclamado As String
        Dim reclamo_queja As String
        Dim reclamo_activo As String
        Dim queja_activo As String
        Dim detalle_reclamo_queja As String
        Dim detalle_canal_pedido As String
        'Traer data del documento --Ini
        Dim ds_datos As New DataSet
        Dim año1 As String
        Dim mes1 As String
        Dim dia1 As String
        Dim Hor1 As String
        Dim min1 As String
        Dim seg1 As String

        año1 = Date.Now.Year.ToString()
        mes1 = Date.Now.Month.ToString()
        dia1 = Date.Now.Day.ToString()
        Hor1 = Date.Now.Hour.ToString()
        min1 = Date.Now.Minute.ToString()
        seg1 = Date.Now.Second.ToString()

        ds_datos = obj.obtener_datos_documento(nro_documento, id_unidad)

        año = ds_datos.Tables(0).Rows(0)(1).ToString()
        mes = ds_datos.Tables(0).Rows(0)(2).ToString()
        dia = ds_datos.Tables(0).Rows(0)(3).ToString()
        dato_empresa = ds_datos.Tables(0).Rows(0)(4).ToString()
        nombre = ds_datos.Tables(0).Rows(0)(5).ToString()
        domicilio = ds_datos.Tables(0).Rows(0)(6).ToString()
        dni = ds_datos.Tables(0).Rows(0)(7).ToString()
        telefono = ds_datos.Tables(0).Rows(0)(8).ToString()
        email = ds_datos.Tables(0).Rows(0)(9).ToString()
        menor_edad = ds_datos.Tables(0).Rows(0)(10).ToString()
        producto_servicio = ds_datos.Tables(0).Rows(0)(11).ToString()
        monto_reclamado = ds_datos.Tables(0).Rows(0)(12).ToString()
        desc_bien_reclamado = ds_datos.Tables(0).Rows(0)(13).ToString()
        reclamo_queja = ds_datos.Tables(0).Rows(0)(14).ToString()
        detalle_reclamo_queja = ds_datos.Tables(0).Rows(0)(15).ToString()
        detalle_canal_pedido = ds_datos.Tables(0).Rows(0)(16).ToString()

        If telefono <> "" Then
            email = telefono + " / " + email
        End If



        If producto_servicio = "1" Then 'Producto
            producto_activo = "X"
            servicio_activo = ""
        Else 'Servicio
            producto_activo = ""
            servicio_activo = "X"
        End If

        If reclamo_queja = "1" Then 'Reclamo
            reclamo_activo = "X"
            queja_activo = ""
        Else 'Queja
            reclamo_activo = ""
            queja_activo = "X"
        End If


        'Traer data del documento --Fin


        Dim sw As New StringWriter()
        Dim html As String = sw.ToString()
        Dim Doc As New iTextSharp.text.Document()

        'Ubicacion para crear el documento PDF
        'PdfWriter.GetInstance(Doc, New FileStream(Ruta_Descargas & "\" + id_unidad.ToString() + "_" + nro_documento + ".pdf", FileMode.Create))
        PdfWriter.GetInstance(Doc, New FileStream(Ruta_Descargas & "\" + año1.ToString() + mes1.ToString() + dia1.ToString() + Hor1.ToString() + min1.ToString() + seg1.ToString() + "_Documento_" + nro_documento + ".pdf", FileMode.Create))


        Doc.Open()


        'Parrafo del titulo
        Dim c As New iTextSharp.text.Chunk("Formato de Hoja de Reclamación del Libro de Reclamaciones" & "" & vbLf & "", iTextSharp.text.FontFactory.GetFont("Verdana", 15))
        Dim p As New iTextSharp.text.Paragraph()
        p.Alignment = iTextSharp.text.Element.ALIGN_CENTER
        p.Add(c)


        'tabla de Libro de reclamaciones

        Dim table01 As New PdfPTable(5)
        table01.WidthPercentage = 100

        table01.SetWidthPercentage(New Single() {60, 60, 60, 60, 355}, iTextSharp.text.PageSize.A4)


        Dim cell_tb01 As New PdfPCell(New iTextSharp.text.Phrase("LIBRO DE RECLAMACIONES", New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.BOLD)))
        cell_tb01.Colspan = 4
        cell_tb01.BorderWidthBottom = 1.0F
        cell_tb01.BorderWidthTop = 1.0F
        cell_tb01.PaddingBottom = 10.0F
        cell_tb01.PaddingLeft = 10.0F
        cell_tb01.PaddingTop = 4.0F
        cell_tb01.HorizontalAlignment = 0

        Dim cell_tb01_ As New PdfPCell(New iTextSharp.text.Phrase("HOJA DE RECLAMACIÓN " + nro_documento, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.BOLD)))
        cell_tb01_.Colspan = 1
        cell_tb01_.Rowspan = 2
        cell_tb01_.BorderWidthBottom = 1.0F
        cell_tb01_.BorderWidthTop = 1.0F
        cell_tb01_.PaddingBottom = 10.0F
        cell_tb01_.PaddingLeft = 10.0F
        cell_tb01_.PaddingTop = 4.0F
        cell_tb01_.HorizontalAlignment = 1
        cell_tb01_.VerticalAlignment = 1


        Dim cell_tb01_11 As New PdfPCell(New iTextSharp.text.Phrase("FECHA:", New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb01_11.Colspan = 1
        cell_tb01_11.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        cell_tb01_11.BorderWidthBottom = 1.0F
        cell_tb01_11.BorderWidthTop = 1.0F
        cell_tb01_11.PaddingBottom = 10.0F
        cell_tb01_11.PaddingLeft = 10.0F
        cell_tb01_11.PaddingTop = 4.0F

        Dim cell_tb01_12 As New PdfPCell(New iTextSharp.text.Phrase(dia, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb01_12.Colspan = 1
        cell_tb01_12.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        cell_tb01_12.BorderWidthBottom = 1.0F
        cell_tb01_12.BorderWidthTop = 1.0F
        cell_tb01_12.PaddingBottom = 10.0F
        cell_tb01_12.PaddingLeft = 10.0F
        cell_tb01_12.PaddingTop = 4.0F

        Dim cell_tb01_13 As New PdfPCell(New iTextSharp.text.Phrase(mes, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb01_13.Colspan = 1
        cell_tb01_13.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        cell_tb01_13.BorderWidthBottom = 1.0F
        cell_tb01_13.BorderWidthTop = 1.0F
        cell_tb01_13.PaddingBottom = 10.0F
        cell_tb01_13.PaddingLeft = 10.0F
        cell_tb01_13.PaddingTop = 4.0F

        Dim cell_tb01_14 As New PdfPCell(New iTextSharp.text.Phrase(año, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb01_14.Colspan = 1
        cell_tb01_14.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        cell_tb01_14.BorderWidthBottom = 1.0F
        cell_tb01_14.BorderWidthTop = 1.0F
        cell_tb01_14.PaddingBottom = 10.0F
        cell_tb01_14.PaddingLeft = 10.0F
        cell_tb01_14.PaddingTop = 4.0F

        Dim cell_tb01_21 As New PdfPCell(New iTextSharp.text.Phrase(dato_empresa, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb01_21.Colspan = 5
        cell_tb01_21.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb01_21.BorderWidthBottom = 1.0F
        cell_tb01_21.BorderWidthTop = 1.0F
        cell_tb01_21.PaddingBottom = 10.0F
        cell_tb01_21.PaddingLeft = 10.0F
        cell_tb01_21.PaddingTop = 4.0F




        table01.AddCell(cell_tb01)
        table01.AddCell(cell_tb01_)
        table01.AddCell(cell_tb01_11)
        table01.AddCell(cell_tb01_12)
        table01.AddCell(cell_tb01_13)
        table01.AddCell(cell_tb01_14)
        table01.AddCell(cell_tb01_21)


        'tabla de Libro de consumidor (pertenece a una misma tabla)

        Dim table02 As New PdfPTable(2)
        table02.WidthPercentage = 100
        table02.SetWidthPercentage(New Single() {150, 445}, iTextSharp.text.PageSize.A4)

        Dim cell_tb02 As New PdfPCell(New iTextSharp.text.Phrase("1. IDENTIFICACION DEL CONSUMIDOR RECLAMANTE", New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.BOLD)))
        cell_tb02.Colspan = 2
        cell_tb02.BorderWidthBottom = 1.0F
        cell_tb02.BorderWidthTop = 1.0F
        cell_tb02.PaddingBottom = 10.0F
        cell_tb02.PaddingLeft = 10.0F
        cell_tb02.PaddingTop = 4.0F
        cell_tb02.HorizontalAlignment = 0

        Dim cell_tb02_11 As New PdfPCell(New iTextSharp.text.Phrase("NOMBRE: " + nombre, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb02_11.Colspan = 2
        cell_tb02_11.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb02_11.BorderWidthBottom = 1.0F
        cell_tb02_11.BorderWidthTop = 1.0F
        cell_tb02_11.PaddingBottom = 10.0F
        cell_tb02_11.PaddingLeft = 10.0F
        cell_tb02_11.PaddingTop = 4.0F

        Dim cell_tb02_21 As New PdfPCell(New iTextSharp.text.Phrase("DOMICILIO: " + domicilio, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb02_21.Colspan = 2
        cell_tb02_21.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb02_21.BorderWidthBottom = 1.0F
        cell_tb02_21.BorderWidthTop = 1.0F
        cell_tb02_21.PaddingBottom = 10.0F
        cell_tb02_21.PaddingLeft = 10.0F
        cell_tb02_21.PaddingTop = 4.0F

        Dim cell_tb02_31 As New PdfPCell(New iTextSharp.text.Phrase("DNI / C.E: " + dni, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb02_31.Colspan = 1
        cell_tb02_31.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb02_31.BorderWidthBottom = 1.0F
        cell_tb02_31.BorderWidthTop = 1.0F
        cell_tb02_31.PaddingBottom = 10.0F
        cell_tb02_31.PaddingLeft = 10.0F
        cell_tb02_31.PaddingTop = 4.0F

        Dim cell_tb02_32 As New PdfPCell(New iTextSharp.text.Phrase("TELÉFONO / EMAIL: " + email, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb02_32.Colspan = 1
        cell_tb02_32.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb02_32.BorderWidthBottom = 1.0F
        cell_tb02_32.BorderWidthTop = 1.0F
        cell_tb02_32.PaddingBottom = 10.0F
        cell_tb02_32.PaddingLeft = 10.0F
        cell_tb02_32.PaddingTop = 4.0F

        Dim cell_tb02_41 As New PdfPCell(New iTextSharp.text.Phrase("PADRE O MADRE: " + menor_edad, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb02_41.Colspan = 2
        cell_tb02_41.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb02_41.BorderWidthBottom = 1.0F
        cell_tb02_41.BorderWidthTop = 1.0F
        cell_tb02_41.PaddingBottom = 10.0F
        cell_tb02_41.PaddingLeft = 10.0F
        cell_tb02_41.PaddingTop = 4.0F


        table02.AddCell(cell_tb02)
        table02.AddCell(cell_tb02_11)
        table02.AddCell(cell_tb02_21)
        table02.AddCell(cell_tb02_31)
        table02.AddCell(cell_tb02_32)
        table02.AddCell(cell_tb02_41)


        'tabla del bien contratado (pertenece a una misma tabla)

        Dim table03 As New PdfPTable(3)
        table03.WidthPercentage = 100
        table03.SetWidthPercentage(New Single() {75, 30, 490}, iTextSharp.text.PageSize.A4)

        Dim cell_tb03 As New PdfPCell(New iTextSharp.text.Phrase("2. IDENTIFICACION DEL BIEN CONTRATADO", New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.BOLD)))
        cell_tb03.Colspan = 5
        cell_tb03.BorderWidthBottom = 1.0F
        cell_tb03.BorderWidthTop = 1.0F
        cell_tb03.PaddingBottom = 10.0F
        cell_tb03.PaddingLeft = 10.0F
        cell_tb03.PaddingTop = 4.0F
        cell_tb03.HorizontalAlignment = 0

        Dim cell_tb03_11 As New PdfPCell(New iTextSharp.text.Phrase("PRODUCTO", New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb03_11.Colspan = 1
        cell_tb03_11.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb03_11.BorderWidthBottom = 1.0F
        cell_tb03_11.BorderWidthTop = 1.0F
        cell_tb03_11.PaddingBottom = 10.0F
        cell_tb03_11.PaddingLeft = 10.0F
        cell_tb03_11.PaddingTop = 4.0F

        Dim cell_tb03_13 As New PdfPCell(New iTextSharp.text.Phrase(producto_activo, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb03_13.Colspan = 1
        'cell_tb03_13.HorizontalAlignment = 1
        cell_tb03_13.BorderWidthBottom = 1.0F
        cell_tb03_13.BorderWidthTop = 1.0F
        cell_tb03_13.PaddingBottom = 10.0F
        cell_tb03_13.PaddingLeft = 10.0F
        cell_tb03_13.PaddingTop = 4.0F

        Dim cell_tb03_14 As New PdfPCell(New iTextSharp.text.Phrase("MONTO RECLAMADO: " + monto_reclamado, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb03_14.Colspan = 1
        cell_tb03_14.BorderWidthBottom = 1.0F
        cell_tb03_14.BorderWidthTop = 1.0F
        cell_tb03_14.PaddingBottom = 10.0F
        cell_tb03_14.PaddingLeft = 10.0F
        cell_tb03_14.PaddingTop = 4.0F

        Dim cell_tb03_21 As New PdfPCell(New iTextSharp.text.Phrase("SERVICIO", New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb03_21.Colspan = 1
        'cell_tb03_21.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb03_21.BorderWidthBottom = 1.0F
        cell_tb03_21.BorderWidthTop = 1.0F
        cell_tb03_21.PaddingBottom = 10.0F
        cell_tb03_21.PaddingLeft = 10.0F
        cell_tb03_21.PaddingTop = 4.0F

        Dim cell_tb03_23 As New PdfPCell(New iTextSharp.text.Phrase(servicio_activo, New Font(Font.FontFamily.TIMES_ROMAN, 9.0F, Font.NORMAL)))
        cell_tb03_23.Colspan = 1
        cell_tb03_23.HorizontalAlignment = 1
        cell_tb03_23.BorderWidthBottom = 1.0F
        cell_tb03_23.BorderWidthTop = 1.0F
        cell_tb03_23.PaddingBottom = 10.0F
        cell_tb03_23.PaddingLeft = 10.0F
        cell_tb03_23.PaddingTop = 4.0F

        Dim cell_tb03_24 As New PdfPCell(New iTextSharp.text.Phrase("DESCRIPCIÓN: " + desc_bien_reclamado, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb03_24.Colspan = 1
        cell_tb03_24.HorizontalAlignment = 0
        cell_tb03_24.BorderWidthBottom = 1.0F
        cell_tb03_24.BorderWidthTop = 1.0F
        cell_tb03_24.PaddingBottom = 10.0F
        cell_tb03_24.PaddingLeft = 10.0F
        cell_tb03_24.PaddingTop = 4.0F


        table03.AddCell(cell_tb03)
        table03.AddCell(cell_tb03_11)
        table03.AddCell(cell_tb03_13)
        table03.AddCell(cell_tb03_14)
        table03.AddCell(cell_tb03_21)
        table03.AddCell(cell_tb03_23)
        table03.AddCell(cell_tb03_24)




        Dim table04 As New PdfPTable(5)
        table04.WidthPercentage = 100
        table04.SetWidthPercentage(New Single() {395, 70, 30, 70, 30}, iTextSharp.text.PageSize.A4)

        Dim cell_tb04 As New PdfPCell(New iTextSharp.text.Phrase("3. DETALLE DE LA RECLAMACIÓN Y PEDIDO DEL CONSUMIDOR", New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.BOLD)))
        cell_tb04.Colspan = 1
        cell_tb04.BorderWidthBottom = 1.0F
        cell_tb04.BorderWidthTop = 1.0F
        cell_tb04.PaddingBottom = 10.0F
        cell_tb04.PaddingLeft = 10.0F
        cell_tb04.PaddingTop = 4.0F
        cell_tb04.HorizontalAlignment = 0

        Dim cell_tb04_12 As New PdfPCell(New iTextSharp.text.Phrase("RECLAMO", New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb04_12.Colspan = 1
        cell_tb04_12.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb04_12.BorderWidthBottom = 1.0F
        cell_tb04_12.BorderWidthTop = 1.0F
        cell_tb04_12.PaddingBottom = 10.0F
        cell_tb04_12.PaddingLeft = 10.0F
        cell_tb04_12.PaddingTop = 4.0F

        Dim cell_tb04_13 As New PdfPCell(New iTextSharp.text.Phrase(reclamo_activo, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb04_13.Colspan = 1
        'cell_tb04_13.HorizontalAlignment = 1
        cell_tb04_13.BorderWidthBottom = 1.0F
        cell_tb04_13.BorderWidthTop = 1.0F
        cell_tb04_13.PaddingBottom = 10.0F
        cell_tb04_13.PaddingLeft = 10.0F
        cell_tb04_13.PaddingTop = 4.0F

        Dim cell_tb04_14 As New PdfPCell(New iTextSharp.text.Phrase("QUEJA", New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb04_14.Colspan = 1
        cell_tb04_14.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb04_14.BorderWidthBottom = 1.0F
        cell_tb04_14.BorderWidthTop = 1.0F
        cell_tb04_14.PaddingBottom = 10.0F
        cell_tb04_14.PaddingLeft = 10.0F
        cell_tb04_14.PaddingTop = 4.0F


        Dim cell_tb04_15 As New PdfPCell(New iTextSharp.text.Phrase(queja_activo, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb04_15.Colspan = 1
        'cell_tb04_15.HorizontalAlignment = 1 '0=Left, 1=Centre, 2=Right
        cell_tb04_15.BorderWidthBottom = 1.0F
        cell_tb04_15.BorderWidthTop = 1.0F
        cell_tb04_15.PaddingBottom = 10.0F
        cell_tb04_15.PaddingLeft = 10.0F
        cell_tb04_15.PaddingTop = 4.0F

        Dim cell_tb04_21 As New PdfPCell(New iTextSharp.text.Phrase("DETALLE: " + detalle_reclamo_queja, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb04_21.Colspan = 5
        cell_tb04_21.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb04_21.BorderWidthBottom = 1.0F
        cell_tb04_21.BorderWidthTop = 1.0F
        cell_tb04_21.PaddingBottom = 10.0F
        cell_tb04_21.PaddingLeft = 10.0F
        cell_tb04_21.PaddingTop = 4.0F

        Dim cell_tb04_31 As New PdfPCell(New iTextSharp.text.Phrase("PEDIDO: " + detalle_canal_pedido, New Font(Font.FontFamily.TIMES_ROMAN, 10.0F, Font.NORMAL)))
        cell_tb04_31.Colspan = 5
        cell_tb04_31.HorizontalAlignment = 0 '0=Left, 1=Centre, 2=Right
        cell_tb04_31.BorderWidthBottom = 1.0F
        cell_tb04_31.BorderWidthTop = 1.0F
        cell_tb04_31.PaddingBottom = 10.0F
        cell_tb04_31.PaddingLeft = 10.0F
        cell_tb04_31.PaddingTop = 4.0F


        table04.AddCell(cell_tb04)
        table04.AddCell(cell_tb04_12)
        table04.AddCell(cell_tb04_13)
        table04.AddCell(cell_tb04_14)
        table04.AddCell(cell_tb04_15)
        table04.AddCell(cell_tb04_21)
        table04.AddCell(cell_tb04_31)




        ' Creamos la imagen y le ajustamos el tamaño
        'Dim imagen As iTextSharp.text.Image = iTextSharp.text.Image.GetInstance(Logo.ToString())
        'imagen.BorderWidth = 0
        'imagen.Alignment = Element.ALIGN_LEFT
        'Dim percentage As Single = 0.0F
        'percentage = 150 / imagen.Width
        'imagen.ScalePercent(percentage * 50)

        Dim firma As New iTextSharp.text.Chunk("Área de Control Interno", iTextSharp.text.FontFactory.GetFont("Verdana", 12))
        Dim p_firma As New iTextSharp.text.Paragraph()
        p_firma.Alignment = iTextSharp.text.Element.ALIGN_RIGHT
        p_firma.Add(firma)

        'Doc.Add(imagen)
        Doc.Add(p)
        Doc.Add(iTextSharp.text.Chunk.NEWLINE)
        Doc.Add(table01)
        Doc.Add(table02)
        Doc.Add(table03)
        Doc.Add(table04)

        'Doc.Add(p_firma)

        Dim xmlReader As System.Xml.XmlTextReader = New System.Xml.XmlTextReader(New StringReader(html))
        Doc.Close()
        'Dim Path As String = Ruta_Descargas & "\Doc_" & nro_documento & ".pdf"




        Dim Path As String = Ruta_Descargas & "\" + año1.ToString() + mes1.ToString() + dia1.ToString() + Hor1.ToString() + min1.ToString() + seg1.ToString() + "_Documento_" + nro_documento + ".pdf"
        txt.Add(Path)

        ShowPdf(txt)




    End Sub

    Public Sub Enviar_Correo_Cambio_Estado()
        Dim EmailFrom As String
        Dim PassFrom As String
        Dim ImagenCorreo1 As String
        Dim ImagenCorreo2 As String
        Dim EmailSubject As String

        'Lista Documentos por Enviar
        Dim ds_documentos As New DataSet
        Dim ds_contactos As New DataSet
        Dim ds_contactos_salud As New DataSet '22062023

        Dim ds_contactos_unidades As New DataSet
        Dim nombre_unidad As String
        Dim Nro_Documento As String
        Dim Nom_Cliente As String
        Dim email_responsable As String
        Dim id_unidad As Integer
        Dim dni As String
        Dim email_cliente As String
        Dim nombre_archivo_otros As String
        Dim can_documento As Integer = 0
        Dim obs_arealegal As String
        Dim obs_gerente_marca As String

        Dim j As Integer
        Dim k As Integer
        Dim i As Integer

        EmailFrom = ConfigurationManager.AppSettings("EmailFrom")
        PassFrom = ConfigurationManager.AppSettings("PassFrom")
        ImagenCorreo1 = ConfigurationManager.AppSettings("ImagenCorreo1")


        Nro_Documento = lblnrodocumento.Text.ToString()
        Nom_Cliente = lblnomcliente.Text.ToString()
        dni = lbldnicliente.Text.ToString()
        email_cliente = lblemail.Text.ToString()
        id_unidad = Integer.Parse(lblidunidad.Text.ToString())
        nombre_unidad = lbldescunidad.Text.ToString()

        If txtobslegal.Text = "" Then
            obs_arealegal = "No existe Observación."
        Else
            obs_arealegal = txtobslegal.Text.ToString()
        End If

        If txtobsGerenteMarca.Text = "" Then
            obs_gerente_marca = "No existe Observación."
        Else
            obs_gerente_marca = txtobsGerenteMarca.Text.ToString()
        End If


        'traer responsables de la unidad
        ds_contactos = obj.lista_responsables_unidad(id_unidad)
        ds_contactos_salud = obj.lista_responsables_unidad_salud(id_unidad) '22062023

        'Verificar estado, que imagen de alerta cargar 
        Dim estado As String
        Dim dias_transcurridos As Integer
        Dim can_nivel_urgencia As Integer
        Dim Nivel_Urgencia_texto As String = ""
        Dim ds_estado As New DataSet
        ds_estado = obj.estado_actual_documento(Nro_Documento)
        estado = ds_estado.Tables(0).Rows(0)(0).ToString() 'COnsultamos a la bd del documento en que estado esta
        dias_transcurridos = Integer.Parse(ds_estado.Tables(1).Rows(0)(0).ToString()) 'Dias transcurridos desde la creacion del documento a la fecha
        can_nivel_urgencia = Integer.Parse(ds_estado.Tables(2).Rows(0)(0).ToString()) 'Dias transcurridos desde la creacion del documento a la fecha

        If can_nivel_urgencia > 0 Then 'Hay uno de Nivel de Urgencia
            Nivel_Urgencia_texto = "(Urgente por Atender)"
        End If
        EmailSubject = ConfigurationManager.AppSettings("EmailSubject") + " / " + Nro_Documento + " " + Nivel_Urgencia_texto

        If estado = 2 Then
            ImagenCorreo2 = ConfigurationManager.AppSettings("ImagenCorreo_estado2")
        End If
        If estado = 3 Then
            ImagenCorreo2 = ConfigurationManager.AppSettings("ImagenCorreo_estado3")
        End If
        If estado = 4 Then
            ImagenCorreo2 = ConfigurationManager.AppSettings("ImagenCorreo_estado4")
        End If
        If estado = 5 Then
            ImagenCorreo2 = ConfigurationManager.AppSettings("ImagenCorreo_estado5")
        End If
        If estado = 6 Then
            ImagenCorreo2 = ConfigurationManager.AppSettings("ImagenCorreo_estado6")
        End If

        Dim fullPath_ImagenCorreo1 As String = ImagenCorreo1
        Dim fullPath_ImagenCorreo2 As String = ImagenCorreo2

        'Cuerpo del correo
        Dim htmlBody As String = "<table style=" + """" + "background-color: #f2edd9;" + """" +
                " border=" + """" + "0" + """" + " width=" + """" + "100%" + """" + " cellpadding=" + """" + "0" + """" + ">
                    <tbody>
                    <tr style=" + """" + "background-color: #a8123e;" + """" + ">
                    <td style=" + """" + "width 696px;" + """" + " colspan=" + """" + "2" + """" + ">
                    <p><img src=""cid:Pic1""></p>
                    </td>
                    </tr>
                    <tr style=" + """" + "background-color: #a8123e;" + """" + ">
                    <td style=" + """" + "width 696px;" + """" + " colspan=" + """" + "2" + """" + ">
                    <p><img src=""cid:Pic2""></p>
                    </td>
                    </tr>
                    <tr>
                    <td style=" + """" + "width 696px;" + """" + " colspan=" + """" + "2" + """" + ">
                    <p>&nbsp;</p>
                    <p style=" + """" + "padding-left: 30px;" + """" + "><strong>Estimados Usuarios</strong></p>
                    </td>
                    </tr>
                    <tr>
                    <td style=" + """" + "width 696px;" + """" + " colspan=" + """" + "2" + """" + ">
                    <p style=" + """" + "padding-left: 30px;" + """" + ">El documento Nro." + Nro_Documento.ToString() + " cambio de estado respecto a la atención de un cliente dentro del flujo de aprobación.
                    Se recomienda ingresar al sistema Administrador de Reclamaciones y verificar el documento para su revisión.</p>
                    <p>&nbsp;</p>
                    </td>
                    </tr>
                    <tr>
                    <td style=" + """" + "width 696px;" + """" + " colspan=" + """" + "2" + """" + ">
                    <p style=" + """" + "padding-left: 30px;" + """" + "><strong>Última Observación ingresada por el Área Legal: </strong>" + obs_arealegal.ToString() + ".</p>
                    <p>&nbsp;</p>
                    </td>
                    </tr>
                    <tr>
                    <td style=" + """" + "width 696px;" + """" + " colspan=" + """" + "2" + """" + ">
                    <p style=" + """" + "padding-left: 30px;" + """" + "><strong>Última Observación ingresada por el Gerente de Marca: </strong>" + obs_gerente_marca.ToString() + ".</p>
                    <p>&nbsp;</p>
                    </td>
                    </tr>

                    <tr>
                    <td style=" + """" + "width 696px;" + """" + " colspan=" + """" + "2" + """" + ">
                    <p style=" + """" + "padding-left: 30px;" + """" + "><strong>Días Transcurridos desde la creación del documento: " + dias_transcurridos.ToString() + " días.</strong></p>
                    </td>
                    </tr>

                    <tr>
                    <td style=" + """" + "width 696px;" + """" + " colspan=" + """" + "2" + """" + ">
                    <p style=" + """" + "padding-left: 30px;" + """" + "><strong><a href=" + """" + "https://argestionreclamos.acuriorestaurantes.net/areclamos2025/" + """" + ">Click para acceder al Sistema</a></strong></p>
                    </td>
                    </tr>

                    <tr>
                    <td style=" + """" + "width 696px;" + """" + " colspan=" + """" + "2" + """" + ">
                    <p style=" + """" + "padding-left: 30px;" + """" + "><strong>Saludos.</strong></p>
                    </td>
                    </tr>
                    <tr>
                    <td style=" + """" + "width 696px;" + """" + " colspan=" + """" + "2" + """" + ">
                    <p style=" + """" + "padding-left: 30px;" + """" + "><strong>" + nombre_unidad + "</strong></p>
                    </td>
                    </tr>
                    </tbody>
                    </table>"

        Dim avHtml As AlternateView = AlternateView.CreateAlternateViewFromString _
             (htmlBody, Nothing, System.Net.Mime.MediaTypeNames.Text.Html)

        Dim pic1 As LinkedResource = New LinkedResource(fullPath_ImagenCorreo1, System.Net.Mime.MediaTypeNames.Image.Jpeg)
        pic1.ContentId = "Pic1"
        Dim pic2 As LinkedResource = New LinkedResource(fullPath_ImagenCorreo2, System.Net.Mime.MediaTypeNames.Image.Jpeg)
        pic2.ContentId = "Pic2"
        avHtml.LinkedResources.Add(pic1)
        avHtml.LinkedResources.Add(pic2)

        Dim serverSmtp As String = "smtp.office365.com"
        Dim emailSmtp As String = EmailFrom
        System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12


        Dim aMailMessage As MailMessage = New MailMessage()

        j = 0
        For j = 0 To ds_contactos.Tables(0).Rows.Count - 1
            email_responsable = ds_contactos.Tables(0).Rows(j)(0).ToString()
            'email_responsable = "imartinez@acuriorestaurantes.com"
            aMailMessage.To.Add(email_responsable)
        Next



        '******Verificar tema Salud / nuevos correos - tabla responsable / columna "activo_unidad_responsable" '22062023
        Dim id_urgencia As Integer
        Dim item As ListItem
        Dim salud As Integer
        salud = 0
        For Each item In chklisturgencia.Items
            If item.Selected = True Then
                'codigo para agregar registro
                id_urgencia = item.Value.ToString()
                If id_urgencia = 3 Or id_urgencia = 14 Then 'Afectación a la salud , Elemento extraño en el pedido
                    salud = 1
                End If
            End If
        Next

        If salud = 1 And estado = 2 Then 'estado Atendido Adm. de Unidad
            j = 0
            For j = 0 To ds_contactos_salud.Tables(0).Rows.Count - 1
                email_responsable = ds_contactos_salud.Tables(0).Rows(j)(0).ToString()
                aMailMessage.To.Add(email_responsable)
            Next
        End If
        '******tema de salud - fin '22062023




        Dim SmtpServer As SmtpClient = New SmtpClient(serverSmtp)
        SmtpServer.EnableSsl = True
        SmtpServer.Port = 587
        SmtpServer.Timeout = 10000
        SmtpServer.DeliveryMethod = SmtpDeliveryMethod.Network
        SmtpServer.UseDefaultCredentials = False
        SmtpServer.Credentials = New System.Net.NetworkCredential(EmailFrom, PassFrom)

        aMailMessage.Subject = EmailSubject
        'aMailMessage.Body = htmlBody ' htmlBody 'body
        aMailMessage.IsBodyHtml = True
        aMailMessage.AlternateViews.Add(avHtml)

        aMailMessage.From = New MailAddress(emailSmtp)
        aMailMessage.BodyEncoding = UTF8Encoding.UTF8
        aMailMessage.DeliveryNotificationOptions = DeliveryNotificationOptions.OnFailure

        'Comentamos solo para las pruebas 17/02/205
        SmtpServer.Send(aMailMessage)


    End Sub

    Protected Sub chklistAccion_SelectedIndexChanged(sender As Object, e As EventArgs) Handles chklistAccion.SelectedIndexChanged

        Dim id_accion_select_c As Integer
        For Each itemc In chklistAccion.Items
            If itemc.Selected = True Then
                'codigo para agregar registro
                id_accion_select_c = itemc.Value.ToString()
                If id_accion_select_c = 5 Then
                    pnlcortesias.Visible = True
                    Return
                Else
                    pnlcortesias.Visible = False
                End If
            Else
                id_accion_select_c = itemc.Value.ToString()
                If id_accion_select_c = 5 Then
                    pnlcortesias.Visible = False
                End If
            End If
        Next





    End Sub

    Private Sub btnAdjuntar2_Click(sender As Object, e As EventArgs) Handles btnAdjuntar2.Click

    End Sub

    Public Sub CerrarSesion_ActivacionGeneral()
        obj.CerrarSesionGlobal(Session("ACTIVACION_GENERAL").ToString())
    End Sub
    Public Sub registro_acceso_pagina(ByVal TokenGlobal As String, ByVal sistema As String, ByVal user As String)
        obj.Registro_SesionSistema(TokenGlobal, user, sistema, "Detalle_Admin_Documento.aspx")
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