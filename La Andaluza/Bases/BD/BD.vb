
Public Module BD

    'Public Cnx As System.Data.SqlClient.SqlConnection
    'Public cmd As System.Data.SqlClient.SqlCommand
    'Public transaction As System.Data.SqlClient.SqlTransaction


    'Public Sub Conectar()
    '    If transaction Is Nothing Then
    '        Cnx = New System.Data.SqlClient.SqlConnection

    '        If Config.connectionString = String.Empty Then
    '            DataBase.buildConnectionString(Config.Server)
    '        End If

    '        Cnx.ConnectionString = Config.connectionString

    '        Abrir()
    '    End If
    'End Sub

    'Public Sub Abrir()
    '    Try
    '        If Cnx.State = ConnectionState.Open Then Cnx.Close()
    '        Cnx.Open()
    '    Catch e As Exception
    '    End Try
    'End Sub

    'Public Sub Cerrar()
    '    If transaction Is Nothing Then
    '        Try
    '            If Cnx.State = ConnectionState.Open Then
    '                Cnx.Close()
    '                Cnx.Dispose()
    '            End If
    '        Catch e As Exception
    '        End Try
    '    End If
    'End Sub

    'Public Sub EmpezarTransaccion()
    '    Conectar()
    '    transaction = Cnx.BeginTransaction()
    'End Sub

    'Public Sub TerminarTransaccion()
    '    transaction.Commit()
    '    transaction = Nothing
    '    Cerrar()
    'End Sub

    'Public Sub CancelarTransaccion()
    '    transaction.Rollback()
    '    transaction = Nothing
    '    Cerrar()
    'End Sub

 




End Module
