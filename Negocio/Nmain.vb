Public Class Nmain
    Dim obj As New Data.Dmain
    Public Function grafica_main(ByVal id_unidad As Integer) As DataSet
        Dim ds As New DataSet
        ds = obj.grafica_main(id_unidad)
        Return ds
    End Function
End Class
