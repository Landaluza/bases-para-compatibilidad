
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

    Public Function RealizarConsulta(ByVal Cadena As String) As DataTable
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



    Public Function EjecutarRealizarConsulta(ByVal strRealizarConsulta As String) As DataTable
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


    


#End Region





End Module
