Public Class main
    Inherits System.Web.UI.Page
    Dim obj As New Negocio.NReclamaciones
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load


        '================ PROCESO COOKIE =================
        UtilLog.EscribirLog(
                "main.aspx: Abrio main.aspx - CargarActivacionGeneral()")

        CargarActivacionGeneral()

        If Not Page.IsPostBack Then
            If Session("user") IsNot Nothing Then
                Dim user As String
                user = Session("user")
                lbluserad.Text = Session("user").ToString()

                'Registro Auditoria
                registro_acceso_pagina(Session("ACTIVACION_GENERAL").ToString(), Session("SistemaAcceso").ToString(), Session("user").ToString())


            End If

        End If
        If (Session("user") Is Nothing) Then
                Response.Redirect("~/login.aspx")
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
        Response.Redirect("main.aspx")



    End Sub

    Public Sub CerrarSesion_ActivacionGeneral()
        obj.CerrarSesionGlobal(Session("ACTIVACION_GENERAL").ToString())


        '================ PROCESO COOKIE =================
        'Session.Abandon()

        'Dim ck As New HttpCookie("ACTIVACION_GENERAL")
        'ck.Expires = DateTime.Now.AddDays(-1)

        'Response.Cookies.Add(ck)
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
    Public Sub registro_acceso_pagina(ByVal TokenGlobal As String, ByVal sistema As String, ByVal user As String)
        obj.Registro_SesionSistema(TokenGlobal, user, sistema, "main.aspx")
    End Sub
End Class