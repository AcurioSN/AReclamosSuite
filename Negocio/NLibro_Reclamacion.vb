Imports Data.DLibro_Reclamacion
Public Class NLibro_Reclamacion
    Dim obj As New Data.DLibro_Reclamacion
    Public Function Lista_Marcas() As DataSet
        Try
            Return obj.Lista_Marcas()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Unidades_Marca_Fisico(ByVal id_marca As Integer, ByVal user As String) As DataSet
        Try
            Return obj.Lista_Unidades_Marca_Fisico(id_marca, user)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Detalle_Unidad(ByVal id_unidad As Integer) As DataSet
        Try
            Return obj.Lista_Detalle_Unidad(id_unidad)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Reclamo_Queja() As DataSet
        Try
            Return obj.Lista_Reclamo_Queja()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Canal_Pedido() As DataSet
        Try
            Return obj.Lista_canal_pedido()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Distritos() As DataSet
        Try
            Return obj.Lista_Distritos()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    Public Function Lista_Producto_Servicio() As DataSet
        Try
            Return obj.Lista_Producto_Servicio()

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    'Agregar IM
    Public Function obtener_datos_documento(ByVal Nro_documento As String, ByVal id_unidad As Integer) As DataSet
        Try
            Return obj.obtener_datos_documento(Nro_documento, id_unidad)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function

    'Agregar IM
    Public Function obtener_responsables_correo_imagen(ByVal id_unidad As Integer) As DataSet
        Try
            Return obj.obtener_responsables_correo_imagen(id_unidad)

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
            Return obj.RegistrarDocumento(NroDocFisico, nom_consumidor, domicilio_consumidor, dni_consumidor, telef_consumidor, email_consumidor, menor_edad, reclamo_queja,
                                        detalle_reclamo_queja, detalle_canal_pedido, id_canal, id_distrito, id_unidad, producto_servicio, monto_reclamado, desc_bien_contratado, fec_documento)

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Function
End Class
