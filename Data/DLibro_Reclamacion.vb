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
Public Class DLibro_Reclamacion
    Public Function Lista_Marcas() As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_Lista_Marcas"

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)

            Return ds

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Unidades_Marca_Fisico(ByVal id_marca As Integer, ByVal user As String) As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_Lista_Unidades_Fisico"
            cmd.Parameters.Add("@id_marca", SqlDbType.Int)
            cmd.Parameters(0).Value = id_marca
            cmd.Parameters.Add("@user", SqlDbType.VarChar)
            cmd.Parameters(1).Value = user

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)

            Return ds

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Detalle_Unidad(ByVal id_unidad As Integer) As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_Detalle_Unidad"
            cmd.Parameters.Add("@id_unidad", SqlDbType.Int)
            cmd.Parameters(0).Value = id_unidad

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)

            Return ds

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Reclamo_Queja() As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_tipo"

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)

            Return ds

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_canal_pedido() As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_canal_pedido"

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)

            Return ds

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Distritos() As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_distritos"

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)

            Return ds

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Producto_Servicio() As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_Bien_Contratado"

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)

            Return ds

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    'Agregar IM
    Public Function obtener_datos_documento(ByVal nro_documento As String, ByVal id_unidad As Integer) As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_datos_documento"
            cmd.Parameters.Add("@nro_documento", SqlDbType.VarChar)
            cmd.Parameters(0).Value = nro_documento
            cmd.Parameters.Add("@id_unidad", SqlDbType.Int)
            cmd.Parameters(1).Value = id_unidad

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)

            Return ds

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function



    'Agregar IM
    Public Function obtener_responsables_correo_imagen(ByVal id_unidad As Integer) As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_responsables_unidad"

            cmd.Parameters.Add("@id_unidad", SqlDbType.Int)
            cmd.Parameters(0).Value = id_unidad

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)

            Return ds

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function


    Public Function RegistrarDocumento(ByVal NroDocFisico As String, ByVal nom_consumidor As String, ByVal domicilio_consumidor As String, ByVal dni_consumidor As String,
                                                         ByVal telef_consumidor As String, ByVal email_consumidor As String, ByVal menor_edad As String,
                                                         ByVal reclamo_queja As String, ByVal detalle_reclamo_queja As String, ByVal detalle_canal_pedido As String,
                                                         ByVal id_canal As Integer, ByVal id_distrito As Integer, ByVal id_unidad As Integer, ByVal producto_servicio As String,
                                                         ByVal monto_reclamado As String, ByVal desc_bien_contratado As String, ByVal fec_documento As String) As String
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_RegistrarDocumento_Fisico"
            cmd.Parameters.Add("@NroDocFisico", SqlDbType.VarChar)
            cmd.Parameters(0).Value = NroDocFisico
            cmd.Parameters.Add("@nom_consumidor", SqlDbType.VarChar)
            cmd.Parameters(1).Value = nom_consumidor
            cmd.Parameters.Add("@domicilio_consumidor", SqlDbType.VarChar)
            cmd.Parameters(2).Value = domicilio_consumidor
            cmd.Parameters.Add("@dni_consumidor", SqlDbType.VarChar)
            cmd.Parameters(3).Value = dni_consumidor
            cmd.Parameters.Add("@telef_consumidor", SqlDbType.VarChar)
            cmd.Parameters(4).Value = telef_consumidor
            cmd.Parameters.Add("@email_consumidor", SqlDbType.VarChar)
            cmd.Parameters(5).Value = email_consumidor
            cmd.Parameters.Add("@menor_edad", SqlDbType.VarChar)
            cmd.Parameters(6).Value = menor_edad
            cmd.Parameters.Add("@reclamo_queja", SqlDbType.VarChar)
            cmd.Parameters(7).Value = reclamo_queja
            cmd.Parameters.Add("@detalle_reclamo_queja", SqlDbType.VarChar)
            cmd.Parameters(8).Value = detalle_reclamo_queja
            cmd.Parameters.Add("@detalle_canal_pedido", SqlDbType.VarChar)
            cmd.Parameters(9).Value = detalle_canal_pedido
            cmd.Parameters.Add("@id_canal", SqlDbType.Int)
            cmd.Parameters(10).Value = id_canal
            cmd.Parameters.Add("@id_distrito", SqlDbType.Int)
            cmd.Parameters(11).Value = id_distrito
            cmd.Parameters.Add("@id_unidad", SqlDbType.Int)
            cmd.Parameters(12).Value = id_unidad
            cmd.Parameters.Add("@producto_servicio", SqlDbType.VarChar)
            cmd.Parameters(13).Value = producto_servicio
            cmd.Parameters.Add("@monto_reclamado", SqlDbType.VarChar)
            cmd.Parameters(14).Value = monto_reclamado
            cmd.Parameters.Add("@desc_bien_contratado", SqlDbType.VarChar)
            cmd.Parameters(15).Value = desc_bien_contratado
            cmd.Parameters.Add("@fec_documento", SqlDbType.VarChar)
            cmd.Parameters(16).Value = fec_documento


            Dim outputParam As SqlParameter
            outputParam = cmd.Parameters.Add("@msj", SqlDbType.VarChar, 100)
            outputParam.Direction = ParameterDirection.Output

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim id As String = outputParam.Value
            Return id

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

End Class
