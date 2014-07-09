Public Class ctlCuadroDeTexto
    Private cls As New clsCuadroDeTexto

    Function Validar(ByVal ID As Integer, ByVal Campo As String, ByVal CampoValor As String, ByVal CampoID As String, ByVal Tabla As String) As Boolean
        cls._ID = ID
        cls._Campo = Campo
        cls._CampoValor = CampoValor
        cls._CampoID = CampoID
        cls._Tabla = Tabla

        If ID = 0 Then
            Return cls.Validar()
        Else
            If cls.EsMio Then
                Return False
            Else
                Return cls.Validar()
            End If
        End If
    End Function

End Class
