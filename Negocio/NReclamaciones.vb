Imports System.Data
Imports Data.DReclamaciones
Public Class NReclamaciones
    Dim obj As New Data.DReclamaciones
    Public Function Reporte_Consolidado_Marcas(ByVal fec_ini As String, ByVal fec_fin As String) As DataSet
        Try
            Return obj.Reporte_Consolidado_Marcas(fec_ini, fec_fin)

        Catch ex As Exception

        End Try
    End Function

    Public Function Busqueda_Report01(ByVal fecha_inicial As String, ByVal fecha_final As String, ByVal cod_unidad As String, ByVal dni_cliente As String, ByVal tipo As String) As DataSet
        Try
            Return obj.Busqueda_Report01(fecha_inicial, fecha_final, cod_unidad, dni_cliente, tipo)

        Catch ex As Exception

        End Try
    End Function

    Public Function Busqueda_Documentos_Perfil_estado(ByVal fecha_inicial As String, ByVal fecha_final As String, ByVal cod_unidad As String, ByVal dni_cliente As String, ByVal id_perfil As Integer, ByVal User As String) As DataSet
        Try
            Return obj.Busqueda_Documentos_Perfil_estado(fecha_inicial, fecha_final, cod_unidad, dni_cliente, id_perfil, User)

        Catch ex As Exception

        End Try
    End Function


    Public Function lista_Detalle_Documento(ByVal nro_documento As String) As DataSet
        Try
            Return obj.lista_Detalle_Documento(nro_documento)

        Catch ex As Exception

        End Try
    End Function

    Public Function Lista_Desplegables() As DataSet
        Try
            Return obj.Lista_Desplegables()

        Catch ex As Exception

        End Try
    End Function

    Public Function Lista_Desplegable_estados() As DataSet
        Try
            Return obj.Lista_Desplegable_estados()

        Catch ex As Exception

        End Try
    End Function

    Public Function Busqueda_Documentos_Perfil_estado_main(ByVal fecha_inicial As String, ByVal fecha_final As String, ByVal cod_unidad As String, ByVal dni_cliente As String, ByVal id_perfil As Integer, ByVal User As String, ByVal id_estado As Integer) As DataSet
        Try
            Return obj.Busqueda_Documentos_Perfil_estado_main(fecha_inicial, fecha_final, cod_unidad, dni_cliente, id_perfil, User, id_estado)

        Catch ex As Exception

        End Try
    End Function

    Public Function Registra_Documento_flujo(ByVal nro_documento As String, ByVal user_ad As String, ByVal detalle_accion As String,
                                                 ByVal costo_accion As Decimal, ByVal obs_legal As String, ByVal obs_gerente_marca As String,
                                                 ByVal id_perfil As Integer, ByVal aprobado As Integer, ByVal otros_gastos As Decimal, ByVal atencion_manual As String) As Integer
        Dim id_documento_flujo As Integer
        id_documento_flujo = obj.Registra_Documento_flujo(nro_documento, user_ad, detalle_accion, costo_accion, obs_legal, obs_gerente_marca, id_perfil, aprobado, otros_gastos, atencion_manual)
        Return id_documento_flujo
    End Function

    Public Function Actualizar_Documento_campos_fisico(ByVal nro_documento As String, ByVal tipo_documento As Integer, ByVal producto_servicio As Integer, ByVal monto_reclamado As String,
                                                       ByVal desc_bien_contratado As String, ByVal detalle_reclamacion As String, ByVal Pedido_cliente As String,
                                                        ByVal nombre_cliente As String, ByVal dni As String, ByVal direccion As String,
                                                       ByVal telefono As String, ByVal email As String) As Boolean

        obj.Actualizar_Documento_campos_fisico(nro_documento, tipo_documento, producto_servicio, monto_reclamado, desc_bien_contratado, detalle_reclamacion, Pedido_cliente,
                                                nombre_cliente, dni, direccion, telefono, email)
        Return True
    End Function
    Public Function Anular_Documento_fisico(ByVal nro_documento As String, ByVal sustento As String) As Boolean

        obj.Anular_Documento_fisico(nro_documento, sustento)
        Return True
    End Function


    Public Function Registra_Detalle_Urgencia_Flujo(ByVal id_doc_flujo As Integer, ByVal id_urgencia As Integer, ByVal otro_motivo As String) As Boolean
        Dim resultado As Boolean
        resultado = obj.Registra_Detalle_Urgencia_Flujo(id_doc_flujo, id_urgencia, otro_motivo)

        Return resultado
    End Function

    Public Function Registra_Detalle_emp_reclamo_Flujo(ByVal id_doc_flujo As Integer, ByVal id_urgencia As Integer) As Boolean
        Dim resultado As Boolean
        resultado = obj.Registra_Detalle_emp_reclamo_Flujo(id_doc_flujo, id_urgencia)

        Return resultado
    End Function

    Public Function Registra_Detalle_accion_Flujo(ByVal id_doc_flujo As Integer, ByVal id_accion As Integer, ByVal det_otro_accion As String) As Boolean
        Dim resultado As Boolean
        resultado = obj.Registra_Detalle_accion_Flujo(id_doc_flujo, id_accion, det_otro_accion)

        Return resultado
    End Function

    Public Function Registra_Detalle_prodserv_Flujo(ByVal id_doc_flujo As Integer, ByVal id_prodserv As Integer) As Boolean
        Dim resultado As Boolean
        resultado = obj.Registra_Detalle_prodserv_Flujo(id_doc_flujo, id_prodserv)

        Return resultado
    End Function

    Public Function Registra_Archivos_Adjuntos_Flujo(ByVal id_doc_flujo As Integer, ByVal nom_archivo As String) As Boolean
        Dim resultado As Boolean
        resultado = obj.Registra_Archivos_Adjuntos_Flujo(id_doc_flujo, nom_archivo)

        Return resultado
    End Function
    Public Function Registra_Archivos_Adjuntos_Fisico(ByVal Nro_Documento As String, ByVal nom_archivo As String) As Boolean
        Dim resultado As Boolean
        resultado = obj.Registra_Archivos_Adjuntos_Fisico(Nro_Documento, nom_archivo)

        Return resultado
    End Function

    Public Function obtener_datos_documento(ByVal Nro_documento As String, ByVal id_unidad As Integer) As DataSet
        Try
            Return obj.obtener_datos_documento(Nro_documento, id_unidad)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function


    Public Function lista_responsables_unidad(ByVal id_unidad As Integer) As DataSet
        Try
            Return obj.lista_responsables_unidad(id_unidad)

        Catch ex As Exception

        End Try
    End Function
    Public Function lista_responsables_unidad_salud(ByVal id_unidad As Integer) As DataSet
        Try
            Return obj.lista_responsables_unidad_salud(id_unidad)

        Catch ex As Exception

        End Try
    End Function

    Public Function estado_actual_documento(ByVal nro_documento As String) As DataSet
        Try
            Return obj.estado_actual_documento(nro_documento)

        Catch ex As Exception

        End Try
    End Function

    Public Function Registra_Datos_Cortesia(ByVal nro_documento As String, ByVal si_no_cortesia As String, ByVal importe_cortesia_si As Decimal, ByVal modo_cortesia As String,
                                            ByVal can_pers_cortesia As Integer, ByVal monto_max_cortesia As Decimal, ByVal valido_cortesia As DateTime, ByVal coordinar_cortesia As String,
                                            ByVal email_cont_cortesia As String, ByVal registra_cortesias As String, ByVal lugar_cortesia As String) As Boolean
        Dim resultado As Boolean
        resultado = obj.Registra_Datos_Cortesia(nro_documento, si_no_cortesia, importe_cortesia_si, modo_cortesia, can_pers_cortesia, monto_max_cortesia, valido_cortesia, coordinar_cortesia, email_cont_cortesia, registra_cortesias, lugar_cortesia)

        Return resultado
    End Function

    Public Function Eliminar_Todos_Archivos_Adjuntos_Flujo(ByVal Nro_Documento As String) As Boolean
        Dim resultado As Boolean
        resultado = obj.Eliminar_Todos_Archivos_Adjuntos_Flujo(Nro_Documento)

        Return resultado
    End Function

    Public Function ExisteTokenActivo(ByVal token As String) As DataSet
        Dim ds As New DataSet
        ds = obj.ExisteTokenActivo(token)
        Return ds
    End Function

    Public Function CerrarSesionGlobal(ByVal tokenGlobal As String) As Boolean

        obj.CerrarSesionGlobal(tokenGlobal)
        Return True
    End Function

    Public Function Registro_SesionSistema(ByVal tokenGlobal As String, ByVal usuario As String, ByVal CodigoSistema As String, ByVal PaginaAcceso As String) As Boolean

        obj.Registro_SesionSistema(tokenGlobal, usuario, CodigoSistema, PaginaAcceso)
        Return True
    End Function
End Class
