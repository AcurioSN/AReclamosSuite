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
Public Class DReclamaciones
    Public Function Reporte_Consolidado_Marcas(ByVal fec_ini As String, ByVal fec_fin As String) As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_consolidado_marca"

            cmd.Parameters.Add("@fec_ini", SqlDbType.Int)
            cmd.Parameters(0).Value = fec_ini
            cmd.Parameters.Add("@fec_fin", SqlDbType.Int)
            cmd.Parameters(1).Value = fec_fin

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim da As New SqlDataAdapter(cmd)
            Dim ds As New DataSet
            da.Fill(ds)

            Return ds

        Catch ex As Exception

        End Try
    End Function

    Public Function Busqueda_Report01(ByVal fecha_inicial As String, ByVal fecha_final As String, ByVal cod_unidad As String, ByVal dni_cliente As String, ByVal tipo As String) As DataSet
        Dim odt As New DataTable
        Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
        Dim cn As New SqlConnection(strConnString)

        Dim cmd As New SqlCommand
        cmd.Connection = cn

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "sp_LReclamaciones_Report01_2023_N"
        cmd.CommandTimeout = 0

        cmd.Parameters.Add("@FechaInicial", SqlDbType.VarChar)
        cmd.Parameters(0).Value = fecha_inicial
        cmd.Parameters.Add("@FechaFinal", SqlDbType.VarChar)
        cmd.Parameters(1).Value = fecha_final
        cmd.Parameters.Add("@cod_unidad", SqlDbType.VarChar)
        cmd.Parameters(2).Value = cod_unidad
        cmd.Parameters.Add("@dni_cliente", SqlDbType.VarChar)
        cmd.Parameters(3).Value = dni_cliente
        cmd.Parameters.Add("@tipo", SqlDbType.VarChar)
        cmd.Parameters(4).Value = tipo


        cn.Open()
        cmd.ExecuteNonQuery()
        cn.Close()

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds)

        Return ds
    End Function

    Public Function Busqueda_Documentos_Perfil_estado(ByVal fecha_inicial As String, ByVal fecha_final As String, ByVal cod_unidad As String, ByVal dni_cliente As String, ByVal id_perfil As Integer, ByVal User As String) As DataSet
        Dim odt As New DataTable
        Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
        Dim cn As New SqlConnection(strConnString)

        Dim cmd As New SqlCommand
        cmd.Connection = cn

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "sp_LReclamaciones_Documentos_Perfil_estado"
        cmd.CommandTimeout = 0

        cmd.Parameters.Add("@FechaInicial", SqlDbType.VarChar)
        cmd.Parameters(0).Value = fecha_inicial
        cmd.Parameters.Add("@FechaFinal", SqlDbType.VarChar)
        cmd.Parameters(1).Value = fecha_final
        cmd.Parameters.Add("@cod_unidad", SqlDbType.VarChar)
        cmd.Parameters(2).Value = cod_unidad
        cmd.Parameters.Add("@dni_cliente", SqlDbType.VarChar)
        cmd.Parameters(3).Value = dni_cliente
        cmd.Parameters.Add("@id_perfil", SqlDbType.VarChar)
        cmd.Parameters(4).Value = id_perfil
        cmd.Parameters.Add("@user", SqlDbType.VarChar)
        cmd.Parameters(5).Value = User


        cn.Open()
        cmd.ExecuteNonQuery()
        cn.Close()

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds)

        Return ds
    End Function


    Public Function lista_Detalle_Documento(ByVal nro_documento As String) As DataSet
        Dim odt As New DataTable
        Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
        Dim cn As New SqlConnection(strConnString)

        Dim cmd As New SqlCommand
        cmd.Connection = cn

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "sp_LReclamaciones_detalle_documento"
        cmd.CommandTimeout = 0

        cmd.Parameters.Add("@nro_documento", SqlDbType.VarChar)
        cmd.Parameters(0).Value = nro_documento


        cn.Open()
        cmd.ExecuteNonQuery()
        cn.Close()

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds)

        Return ds
    End Function


    Public Function Lista_Desplegables() As DataSet
        Dim odt As New DataTable
        Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
        Dim cn As New SqlConnection(strConnString)

        Dim cmd As New SqlCommand
        cmd.Connection = cn

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "sp_LReclamaciones_lista_desplegables"
        cmd.CommandTimeout = 0

        cn.Open()
        cmd.ExecuteNonQuery()
        cn.Close()

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds)

        Return ds
    End Function

    Public Function Lista_Desplegable_estados() As DataSet
        Dim odt As New DataTable
        Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
        Dim cn As New SqlConnection(strConnString)

        Dim cmd As New SqlCommand
        cmd.Connection = cn

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "sp_LReclamaciones_estados"
        cmd.CommandTimeout = 0

        cn.Open()
        cmd.ExecuteNonQuery()
        cn.Close()

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds)

        Return ds
    End Function

    Public Function Busqueda_Documentos_Perfil_estado_main(ByVal fecha_inicial As String, ByVal fecha_final As String, ByVal cod_unidad As String, ByVal dni_cliente As String, ByVal id_perfil As Integer, ByVal User As String, ByVal id_estado As Integer) As DataSet
        Dim odt As New DataTable
        Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
        Dim cn As New SqlConnection(strConnString)

        Dim cmd As New SqlCommand
        cmd.Connection = cn

        cmd.CommandType = CommandType.StoredProcedure
        cmd.CommandText = "sp_LReclamaciones_Documentos_Perfil_estado_main"
        cmd.CommandTimeout = 0

        cmd.Parameters.Add("@FechaInicial", SqlDbType.VarChar)
        cmd.Parameters(0).Value = fecha_inicial
        cmd.Parameters.Add("@FechaFinal", SqlDbType.VarChar)
        cmd.Parameters(1).Value = fecha_final
        cmd.Parameters.Add("@cod_unidad", SqlDbType.VarChar)
        cmd.Parameters(2).Value = cod_unidad
        cmd.Parameters.Add("@dni_cliente", SqlDbType.VarChar)
        cmd.Parameters(3).Value = dni_cliente
        cmd.Parameters.Add("@id_perfil", SqlDbType.VarChar)
        cmd.Parameters(4).Value = id_perfil
        cmd.Parameters.Add("@user", SqlDbType.VarChar)
        cmd.Parameters(5).Value = User
        cmd.Parameters.Add("@id_estado", SqlDbType.VarChar)
        cmd.Parameters(6).Value = id_estado


        cn.Open()
        cmd.ExecuteNonQuery()
        cn.Close()

        Dim da As New SqlDataAdapter(cmd)
        Dim ds As New DataSet
        da.Fill(ds)

        Return ds
    End Function


    Public Function Registra_Documento_flujo(ByVal nro_documento As String, ByVal user_ad As String, ByVal detalle_accion As String,
                                                 ByVal costo_accion As Decimal, ByVal obs_legal As String, ByVal obs_gerente_marca As String,
                                                 ByVal id_perfil As Integer, ByVal aprobado As Integer, ByVal otros_gastos As Decimal, ByVal atencion_manual As String) As Integer
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            'cmd.CommandText = "sp_LReclamaciones_registra_flujo_documento"
            cmd.CommandText = "sp_LReclamaciones_registra_flujo_documento_N" '--- 23032022

            cmd.Parameters.Add("@nro_documento", SqlDbType.VarChar)
            cmd.Parameters(0).Value = nro_documento
            cmd.Parameters.Add("@user_ad", SqlDbType.VarChar)
            cmd.Parameters(1).Value = user_ad
            cmd.Parameters.Add("@detalle_accion", SqlDbType.VarChar)
            cmd.Parameters(2).Value = detalle_accion
            cmd.Parameters.Add("@costo_accion", SqlDbType.Decimal)
            cmd.Parameters(3).Value = costo_accion
            cmd.Parameters.Add("@obs_legal", SqlDbType.VarChar)
            cmd.Parameters(4).Value = obs_legal
            cmd.Parameters.Add("@obs_gerente_marca", SqlDbType.VarChar)
            cmd.Parameters(5).Value = obs_gerente_marca
            cmd.Parameters.Add("@id_perfil", SqlDbType.Int)
            cmd.Parameters(6).Value = id_perfil
            cmd.Parameters.Add("@aprobado", SqlDbType.Int)
            cmd.Parameters(7).Value = aprobado
            cmd.Parameters.Add("@otros_gastos", SqlDbType.Decimal)
            cmd.Parameters(8).Value = otros_gastos
            cmd.Parameters.Add("@atencion_manual", SqlDbType.VarChar)
            cmd.Parameters(9).Value = atencion_manual

            Dim outputParam As SqlParameter
            outputParam = cmd.Parameters.Add("@id_documento_flujo", SqlDbType.Int)
            outputParam.Direction = ParameterDirection.Output



            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Dim id_documento_flujo As Integer = outputParam.Value
            Return id_documento_flujo

        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function Actualizar_Documento_campos_fisico(ByVal nro_documento As String, ByVal tipo_documento As Integer, ByVal producto_servicio As Integer, ByVal monto_reclamado As String,
                                                       ByVal desc_bien_contratado As String, ByVal detalle_reclamacion As String, ByVal Pedido_cliente As String,
                                                       ByVal nombre_cliente As String, ByVal dni As String, ByVal direccion As String,
                                                       ByVal telefono As String, ByVal email As String) As Boolean
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            'cmd.CommandText = "sp_LReclamaciones_registra_flujo_documento"
            cmd.CommandText = "sp_LReclamaciones_Actualizar_campos_fisico" '--- 23032022

            cmd.Parameters.Add("@nro_documento", SqlDbType.VarChar)
            cmd.Parameters(0).Value = nro_documento
            cmd.Parameters.Add("@tipo_documento", SqlDbType.Int)
            cmd.Parameters(1).Value = tipo_documento
            cmd.Parameters.Add("@producto_servicio", SqlDbType.Int)
            cmd.Parameters(2).Value = producto_servicio
            cmd.Parameters.Add("@monto_reclamado", SqlDbType.VarChar)
            cmd.Parameters(3).Value = monto_reclamado
            cmd.Parameters.Add("@desc_bien_contratado", SqlDbType.VarChar)
            cmd.Parameters(4).Value = desc_bien_contratado
            cmd.Parameters.Add("@detalle_reclamacion", SqlDbType.VarChar)
            cmd.Parameters(5).Value = detalle_reclamacion
            cmd.Parameters.Add("@Pedido_cliente", SqlDbType.VarChar)
            cmd.Parameters(6).Value = Pedido_cliente
            cmd.Parameters.Add("@nombre_cliente", SqlDbType.VarChar)
            cmd.Parameters(7).Value = nombre_cliente
            cmd.Parameters.Add("@dni", SqlDbType.VarChar)
            cmd.Parameters(8).Value = dni
            cmd.Parameters.Add("@direccion", SqlDbType.VarChar)
            cmd.Parameters(9).Value = direccion
            cmd.Parameters.Add("@telefono", SqlDbType.VarChar)
            cmd.Parameters(10).Value = telefono
            cmd.Parameters.Add("@email", SqlDbType.VarChar)
            cmd.Parameters(11).Value = email


            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()


            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function Anular_Documento_fisico(ByVal nro_documento As String, ByVal sustento As String) As Boolean
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            'cmd.CommandText = "sp_LReclamaciones_registra_flujo_documento"
            cmd.CommandText = "sp_LReclamaciones_Anular_doc_fisico" '--- 23032022

            cmd.Parameters.Add("@nro_documento", SqlDbType.VarChar)
            cmd.Parameters(0).Value = nro_documento
            cmd.Parameters.Add("@sustento", SqlDbType.VarChar)
            cmd.Parameters(1).Value = sustento

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()


            Return True

        Catch ex As Exception
            Return False
        End Try

    End Function


    Public Function Registra_Detalle_Urgencia_Flujo(ByVal id_doc_flujo As Integer, ByVal id_urgencia As Integer, ByVal otro_motivo As String) As Boolean
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_registra_urgencia"

            cmd.Parameters.Add("@id_doc_flujo", SqlDbType.Int)
            cmd.Parameters(0).Value = id_doc_flujo
            cmd.Parameters.Add("@id_urgencia", SqlDbType.Int)
            cmd.Parameters(1).Value = id_urgencia
            cmd.Parameters.Add("@otro_motivo", SqlDbType.VarChar)
            cmd.Parameters(2).Value = otro_motivo

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function Registra_Detalle_emp_reclamo_Flujo(ByVal id_doc_flujo As Integer, ByVal id_emp_reclamo As Integer) As Boolean
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_registra_emp_reclamo"

            cmd.Parameters.Add("@id_doc_flujo", SqlDbType.Int)
            cmd.Parameters(0).Value = id_doc_flujo
            cmd.Parameters.Add("@id_emp_reclamo", SqlDbType.Int)
            cmd.Parameters(1).Value = id_emp_reclamo


            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function Registra_Detalle_accion_Flujo(ByVal id_doc_flujo As Integer, ByVal id_accion As Integer, ByVal det_otro_accion As String) As Boolean
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_registra_accion_flujo"

            cmd.Parameters.Add("@id_doc_flujo", SqlDbType.Int)
            cmd.Parameters(0).Value = id_doc_flujo
            cmd.Parameters.Add("@id_accion", SqlDbType.Int)
            cmd.Parameters(1).Value = id_accion
            cmd.Parameters.Add("@det_otro_accion", SqlDbType.VarChar)
            cmd.Parameters(2).Value = det_otro_accion


            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function Registra_Detalle_prodserv_Flujo(ByVal id_doc_flujo As Integer, ByVal id_prod_serv As Integer) As Boolean
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_registra_prodserv_flujo"

            cmd.Parameters.Add("@id_doc_flujo", SqlDbType.Int)
            cmd.Parameters(0).Value = id_doc_flujo
            cmd.Parameters.Add("@id_prod_serv", SqlDbType.Int)
            cmd.Parameters(1).Value = id_prod_serv


            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function Registra_Archivos_Adjuntos_Flujo(ByVal id_doc_flujo As Integer, ByVal nom_archivo As String) As Boolean
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_registra_archivo_adjunto"

            cmd.Parameters.Add("@id_doc_flujo", SqlDbType.Int)
            cmd.Parameters(0).Value = id_doc_flujo
            cmd.Parameters.Add("@nom_archivo", SqlDbType.VarChar)
            cmd.Parameters(1).Value = nom_archivo


            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function Registra_Archivos_Adjuntos_Fisico(ByVal Nro_Documento As String, ByVal nom_archivo As String) As Boolean
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_registra_archivo_adjunto_Fisico"

            cmd.Parameters.Add("@nro_documento", SqlDbType.VarChar)
            cmd.Parameters(0).Value = Nro_Documento
            cmd.Parameters.Add("@nom_archivo", SqlDbType.VarChar)
            cmd.Parameters(1).Value = nom_archivo


            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

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

    Public Function lista_responsables_unidad(ByVal id_unidad As Integer) As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_responsables_unidad_enviar"
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

    Public Function lista_responsables_unidad_salud(ByVal id_unidad As Integer) As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_responsables_unidad_enviar_salud"
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

    Public Function estado_actual_documento(ByVal nro_documento As String) As DataSet
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_estado_actual_documento"
            cmd.Parameters.Add("@NroDocumento", SqlDbType.VarChar)
            cmd.Parameters(0).Value = nro_documento

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

    Public Function Registra_Datos_Cortesia(ByVal nro_documento As String, ByVal si_no_cortesia As String, ByVal importe_cortesia_si As Decimal, ByVal modo_cortesia As String,
                                            ByVal can_pers_cortesia As Integer, ByVal monto_max_cortesia As Decimal, ByVal valido_cortesia As DateTime, ByVal coordinar_cortesia As String,
                                            ByVal email_cont_cortesia As String, ByVal registra_cortesias As String, ByVal lugar_cortesia As String) As Boolean
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_Registra_cortesia"

            cmd.Parameters.Add("@nro_documento", SqlDbType.VarChar)
            cmd.Parameters(0).Value = nro_documento
            cmd.Parameters.Add("@si_no_cortesia", SqlDbType.VarChar)
            cmd.Parameters(1).Value = si_no_cortesia
            cmd.Parameters.Add("@importe_cortesia_si", SqlDbType.Decimal)
            cmd.Parameters(2).Value = importe_cortesia_si
            cmd.Parameters.Add("@modo_cortesia", SqlDbType.Char)
            cmd.Parameters(3).Value = modo_cortesia
            cmd.Parameters.Add("@can_pers_cortesia", SqlDbType.Int)
            cmd.Parameters(4).Value = can_pers_cortesia
            cmd.Parameters.Add("@monto_max_cortesia", SqlDbType.Decimal)
            cmd.Parameters(5).Value = monto_max_cortesia
            cmd.Parameters.Add("@valido_cortesia", SqlDbType.DateTime)
            cmd.Parameters(6).Value = valido_cortesia
            cmd.Parameters.Add("@coordinar_cortesia", SqlDbType.VarChar)
            cmd.Parameters(7).Value = coordinar_cortesia
            cmd.Parameters.Add("@email_cont_cortesia", SqlDbType.VarChar)
            cmd.Parameters(8).Value = email_cont_cortesia
            cmd.Parameters.Add("@registra_cortesias", SqlDbType.VarChar)
            cmd.Parameters(9).Value = registra_cortesias
            cmd.Parameters.Add("@lugar_cortesia", SqlDbType.VarChar)
            cmd.Parameters(10).Value = lugar_cortesia

            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function

    Public Function Eliminar_Todos_Archivos_Adjuntos_Flujo(ByVal Nro_Documento As String) As Boolean
        Try
            Dim odt As New DataTable
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion").ConnectionString
            Dim cn As New SqlConnection(strConnString)

            Dim cmd As New SqlCommand
            cmd.Connection = cn

            cmd.CommandType = CommandType.StoredProcedure
            cmd.CommandText = "sp_LReclamaciones_Eliminar_todos_archivo_adjunto"

            cmd.Parameters.Add("@Nro_Documento", SqlDbType.VarChar)
            cmd.Parameters(0).Value = Nro_Documento


            cn.Open()
            cmd.ExecuteNonQuery()
            cn.Close()

            Return True
        Catch ex As Exception
            Return False
        End Try

    End Function
    Public Function ExisteTokenActivo(ByVal token As String) As DataSet
        Dim ds As New DataSet()

        Try
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion_arsysusers").ConnectionString

            Using cn As New SqlConnection(strConnString)
                Using cmd As New SqlCommand("sp_arsuite_Validartoken", cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 0

                    ' Definir parámetro con tipo y tamaño
                    cmd.Parameters.Add("@token", SqlDbType.VarChar, 500).Value = token

                    ' Llenar el DataSet usando SqlDataAdapter
                    Using da As New SqlDataAdapter(cmd)
                        da.Fill(ds)
                    End Using
                End Using
            End Using

        Catch ex As Exception
            ' Registrar el error si es necesario
            ' LogError(ex.Message) 
            Throw ' Relanzar el error para ser manejado en un nivel superior
        End Try

        Return ds
    End Function

    Public Function CerrarSesionGlobal(ByVal tokenGlobal As String) As Boolean
        Try
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion_arsysusers").ConnectionString

            Using cn As New SqlConnection(strConnString)
                Using cmd As New SqlCommand("sp_arsuite_CerrarSesionGlobal", cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 0

                    ' Agregar parámetros con tipo y tamaño
                    cmd.Parameters.Add("@token", SqlDbType.VarChar, 500).Value = tokenGlobal

                    cn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            Return True
        Catch ex As Exception
            ' Aquí puedes registrar el error antes de relanzarlo
            ' LogError(ex.Message) 
            Throw ' Lanza el error para que sea manejado en niveles superiores
        End Try
    End Function

    Public Function Registro_SesionSistema(ByVal tokenGlobal As String, ByVal usuario As String, ByVal CodigoSistema As String, ByVal PaginaAcceso As String) As Boolean
        Try
            Dim strConnString As String = ConfigurationManager.ConnectionStrings("cadenaConexion_arsysusers").ConnectionString

            Using cn As New SqlConnection(strConnString)
                Using cmd As New SqlCommand("sp_arsuite_Registro_SesionSistema", cn)
                    cmd.CommandType = CommandType.StoredProcedure
                    cmd.CommandTimeout = 0

                    ' Agregar parámetros con tipo y tamaño
                    cmd.Parameters.Add("@tokenGlobal", SqlDbType.VarChar, 500).Value = tokenGlobal
                    cmd.Parameters.Add("@usuario", SqlDbType.VarChar, 50).Value = usuario
                    cmd.Parameters.Add("@CodigoSistema", SqlDbType.VarChar, 20).Value = CodigoSistema
                    cmd.Parameters.Add("@PaginaAcceso", SqlDbType.VarChar, 200).Value = PaginaAcceso

                    cn.Open()
                    cmd.ExecuteNonQuery()
                End Using
            End Using

            Return True
        Catch ex As Exception
            ' Aquí puedes registrar el error antes de relanzarlo
            ' LogError(ex.Message) 
            Throw ' Lanza el error para que sea manejado en niveles superiores
        End Try
    End Function
End Class
