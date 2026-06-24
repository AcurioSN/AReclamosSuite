Imports System
Imports System.Data
Imports System.Data.Common
Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Collections.Specialized
Imports System.Web
Imports System.Text
Imports System.IO
Public Class Dmain
    Public Function grafica_main(ByVal id_unidad As Integer) As DataSet
        Dim odt As New DataTable

        Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
        Dim cn As New SqlConnection(strConnString)

        Dim cmd As New SqlCommand
        cmd.Connection = cn

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "sp_LReclamaciones_reclamos_año_unidad"

        cmd.Parameters.Add("@id_unidad", SqlDbType.Int)
        cmd.Parameters(0).Value = id_unidad


        cn.Open()
        cmd.ExecuteNonQuery()
        cn.Close()

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds)


        Return ds
    End Function
End Class
