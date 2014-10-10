
Public Module BD

#Region "Atributos"
    Public Cnx As System.Data.SqlClient.SqlConnection
    Public cmd As System.Data.SqlClient.SqlCommand
    Public transaction As System.Data.SqlClient.SqlTransaction
#End Region



#Region "Propiedades Conexion"
    Public Sub Conectar()
        If transaction Is Nothing Then
            Cnx = New System.Data.SqlClient.SqlConnection

            If Config.connectionString = String.Empty Then
                DataBase.buildConnectionString(Config.Server)
            End If

            Cnx.ConnectionString = Config.connectionString

            Abrir()
        End If
    End Sub

    Public Sub Abrir()
        Try
            If Cnx.State = ConnectionState.Open Then Cnx.Close()
            Cnx.Open()
        Catch e As Exception
        End Try
    End Sub

    Public Sub Cerrar()
        If transaction Is Nothing Then
            Try
                If Cnx.State = ConnectionState.Open Then
                    Cnx.Close()
                    Cnx.Dispose()
                End If
            Catch e As Exception
            End Try
        End If
    End Sub

    Public Sub EmpezarTransaccion()
        Conectar()
        transaction = Cnx.BeginTransaction()
    End Sub

    Public Sub TerminarTransaccion()
        transaction.Commit()
        transaction = Nothing
        Cerrar()
    End Sub

    Public Sub CancelarTransaccion()
        transaction.Rollback()
        transaction = Nothing
        Cerrar()
    End Sub
#End Region

#Region "Consultas Alejandro"
    Private Linea As String

    Private Function RealizarConsulta(ByVal Cadena As String) As DataTable
        Dim dtsTabla As New DataTable
        If transaction Is Nothing Then Conectar()
        Try
            dtsTabla = EjecutarRealizarConsulta(Cadena)
            Return dtsTabla
        Catch ex As Exception
            Return Nothing
        Finally
            If transaction Is Nothing Then Cerrar()
        End Try
    End Function



    Private Function EjecutarRealizarConsulta(ByVal strRealizarConsulta As String) As DataTable
        Dim dtsTemp As New DataSet
        Dim cmd As System.Data.SqlClient.SqlCommand = Nothing

        Try
            cmd = New System.Data.SqlClient.SqlCommand(strRealizarConsulta, Cnx)
            If Not transaction Is Nothing Then cmd.Transaction = transaction
            cmd.CommandTimeout = 300
            Dim Ad As New System.Data.SqlClient.SqlDataAdapter(cmd)
            Ad.Fill(dtsTemp, "NuevaTabla")
            cmd.Dispose()
        Catch ex As Exception
            cmd.Dispose()
            MessageBox.Show("problema en BD.EjecutarRealizarConsulta Ad.Fill" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        Return dtsTemp.Tables(0)
    End Function

    Public Function realizarConsultaAlteraciones(ByVal strrealizarConsulta As String) As Integer
        Dim exito As Integer
        If transaction Is Nothing Then Conectar()
        Dim cmd As System.Data.SqlClient.SqlCommand
        cmd = New System.Data.SqlClient.SqlCommand(strrealizarConsulta, Cnx)
        If Not transaction Is Nothing Then cmd.Transaction = transaction
        Try
            cmd.ExecuteNonQuery()
            exito = 1
        Catch ex As Exception
            MessageBox.Show("Ha ocurrido un error. Detalles:" & Environment.NewLine & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            exito = 0
        End Try
        cmd.Dispose()
        If transaction Is Nothing Then Cerrar()
        Return exito
    End Function


    Public Function ConsultaInsertar(ByVal Datos As String, ByVal tabla As String) As Integer
        Linea = "insert into " + tabla + " values " + "( " + Datos + ",'" + Calendar.ArmarFecha((Today + " " + TimeOfDay)) + "'," + Config.User.ToString + ")"
        Return realizarConsultaAlteraciones(Linea)
    End Function

    Public Function ConsultaInsertarConcampos(ByVal campos As String, ByVal Datos As String, ByVal tabla As String) As Integer
        Linea = "insert into " + tabla + campos + " values " + "( " + Datos + ",'" + Calendar.ArmarFecha((Today + " " + TimeOfDay)) + "'," + Config.User.ToString + ")"
        Return realizarConsultaAlteraciones(Linea)
    End Function

    Public Function ConsultaInsertarSinDatosUsuario(ByVal Datos As String, ByVal tabla As String) As Integer
        Linea = "insert into " + tabla + " values " + "( " + Datos + ")"
        Return realizarConsultaAlteraciones(Linea)
    End Function

    Public Function ConsultaModificar(ByVal tabla As String, ByVal valor1 As String, ByVal restriccion As String) As Integer
        Linea = "UPDATE  " + tabla + " SET " + valor1 + ", FechaModificacion='" + Calendar.ArmarFecha(Today + " " + TimeOfDay) + "',UsuarioModificacion=" & Config.User.ToString + " WHERE " + restriccion
        Return realizarConsultaAlteraciones(Linea)
    End Function

    'Public Function ConsultaEliminar(ByVal tabla As String, ByVal restriccion As String) As Integer
    '    Linea = "delete from " + tabla + " where " + restriccion
    '    Return realizarConsultaAlteraciones(Linea)
    'End Function

    Public Function ConsultaVer(ByVal datos As String, ByVal tabla As String, ByVal restriccion As String, ByVal orderBy As String) As DataTable
        If restriccion.Length = 0 Then
            Linea = "select " + datos + " from " + tabla + " order by " + orderBy
        Else
            Linea = "select " + datos + " from " + tabla + " where " + restriccion + " order by " + orderBy
        End If
        Return RealizarConsulta(Linea)
    End Function

    Public Function ConsultaVer(ByVal datos As String, ByVal tabla As String, ByVal restriccion As String) As DataTable
        Linea = "select " + datos + " from " + tabla + " where " + restriccion
        Return RealizarConsulta(Linea)
    End Function

    Public Function ConsultaVer(ByVal top100 As Boolean, ByVal datos As String, ByVal tabla As String, ByVal restriccion As String) As DataTable
        If top100 Then
            Linea = "select top 100 " + datos + " from " + tabla + " where " + restriccion
        Else
            Linea = "select " + datos + " from " + tabla + " where " + restriccion
        End If
        Return RealizarConsulta(Linea)
    End Function

    Public Function ConsultaVer(ByVal datos As String, ByVal tabla As String) As DataTable
        Linea = "select " + datos + If(tabla <> "", " from " + tabla, "")
        Return RealizarConsulta(Linea)
    End Function
    Public Function ConsultaProcedAlmacenado(ByVal NombreProcedimiento As String, ByVal datos As String) As DataTable
        Try
            Linea = NombreProcedimiento + " " + datos
            Return RealizarConsulta(Linea)
        Catch ex As Exception
            MessageBox.Show(ex, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function


#End Region





End Module
