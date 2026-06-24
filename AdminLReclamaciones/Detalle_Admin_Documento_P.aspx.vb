Public Class Detalle_Admin_Documento_P
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim nro_documento As String = Request.QueryString("dato1")
    End Sub

End Class