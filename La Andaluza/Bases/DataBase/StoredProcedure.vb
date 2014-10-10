

'.SqlConnection
'.SqlTransaction
'.SqlCommand
'.SqlDataReader


Public MustInherit Class StoredProcedure
    'Private connection As System.Data.SqlClient.SqlConnection
    Private selectProcedureName, insertProcedureName, updateProcedureName, deleteProcedureName, selectDgvAllProcedureName, selectDgvByProcedureName As String
    Private DgvProcedure As String
    Protected dtb As DataBase

    Public Property DataGridViewStoredProcedure As String
        Get
            Return Me.DgvProcedure
        End Get
        Set(ByVal value As String)
            Me.DgvProcedure = value
        End Set
    End Property

    Public ReadOnly Property DataGridViewStoredProcedureForSelect As String
        Get
            Return Me.selectDgvAllProcedureName
        End Get
    End Property

    Public ReadOnly Property DataGridViewStoredProcedureForFilteredSelect As String
        Get
            Return selectDgvByProcedureName
        End Get
    End Property

    Public Sub New(ByVal selectProcedureName As String, ByVal insertProcedureName As String, ByVal updateProcedureName As String, ByVal deleteProcedureName As String, ByVal selectDgvAllProcedureName As String, ByVal selectDgvByProcedureName As String)
        Me.selectProcedureName = selectProcedureName
        Me.insertProcedureName = insertProcedureName
        Me.updateProcedureName = updateProcedureName
        Me.deleteProcedureName = deleteProcedureName
        Me.selectDgvAllProcedureName = selectDgvAllProcedureName
        Me.selectDgvByProcedureName = selectDgvByProcedureName

        dtb = New DataBase(Config.Server)
    End Sub

    Public Sub Select_Record(ByRef dbo As DataBussines, Optional ByRef trans As System.Data.SqlClient.SqlTransaction = Nothing)
        dbo.searchKey.name = dbo.key.name
        dbo.searchKey.value = dbo.key.value
        Me.Select_proc(dbo, selectProcedureName, trans)
        'Dim reader As System.Data.SqlClient.SqlDataReader = Nothing
        'Dim selectCommand As System.Data.SqlClient.SqlCommand

        'Try
        '    If trans Is Nothing Then
        '        dtb.Conectar()            
        '    End If

        '    selectCommand = dtb.Comando(selectProcedureName)
        '    selectCommand.CommandType = CommandType.StoredProcedure

        '    If Not trans Is Nothing Then
        '        selectCommand.Transaction = dtb.Transaccion
        '    End If

        '    selectCommand.Parameters.AddWithValue(dbo.searchKey.name, dbo.searchKey.value)
        '    reader = selectCommand.ExecuteReader(CommandBehavior.SingleRow)

        '    If reader.Read Then
        '        Dim cont As Integer = 1

        '        While cont <= dbo.count
        '            dbo.item(cont).value = reader(dbo.item(cont).sqlName)
        '            cont += 1
        '        End While
        '    Else
        '        dbo = Nothing
        '    End If

        'Catch ex As Exception
        '    dbo = Nothing
        'Finally
        '    If Not reader Is Nothing Then If Not reader.IsClosed Then reader.Close()
        '    dtb.Desconectar()
        'End Try
    End Sub

    Protected Sub Select_proc(ByRef dbo As DataBussines, ByVal proc As String, Optional ByRef trans As System.Data.SqlClient.SqlTransaction = Nothing)
        Dim reader As System.Data.SqlClient.SqlDataReader = Nothing
        Dim selectCommand As System.Data.SqlClient.SqlCommand

        Try
            'If trans Is Nothing Then
            dtb.Conectar()
            'End If

            selectCommand = dtb.Comando(proc)
            selectCommand.CommandType = CommandType.StoredProcedure

            If Not trans Is Nothing Then
                selectCommand.Transaction = dtb.Transaccion
            End If

            selectCommand.Parameters.AddWithValue(dbo.searchKey.name, dbo.searchKey.value)
            reader = selectCommand.ExecuteReader(CommandBehavior.SingleRow)

            If reader.Read Then
                Dim cont As Integer = 1

                While cont <= dbo.count
                    dbo.item(cont).value = reader(dbo.item(cont).sqlName)
                    cont += 1
                End While
            Else
                dbo = Nothing
            End If

        Catch ex As Exception
            dbo = Nothing
        Finally
            If Not reader Is Nothing Then If Not reader.IsClosed Then reader.Close()
            dtb.Desconectar()
        End Try
    End Sub

    Protected Function InsertProcedure(ByRef dbo As DataBussines, Optional ByRef trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean
        Try
            If trans Is Nothing Then
                If BD.transaction Is Nothing Then Conectar()
            End If
            Dim connection As System.Data.SqlClient.SqlConnection = BD.Cnx
            Dim insertCommand As New System.Data.SqlClient.SqlCommand(insertProcedureName, connection)
            insertCommand.CommandType = CommandType.StoredProcedure

            If Not trans Is Nothing Then
                insertCommand.Transaction = trans
            Else
                If Not BD.transaction Is Nothing Then
                    insertCommand.Transaction = BD.transaction
                End If
            End If

            Dim obj As DataBussinesParameter
            Dim cont As Integer = 2 'saltamos el id

            While cont <= dbo.count
                obj = dbo.item(cont)
                If IsDBNull(obj.value) Then
                    insertCommand.Parameters.AddWithValue(obj.name, Convert.DBNull)
                Else
                    If obj.value = Nothing Then
                        If TypeOf obj.value Is Boolean Then
                            insertCommand.Parameters.AddWithValue(obj.name, False)
                        Else
                            insertCommand.Parameters.AddWithValue(obj.name, Convert.DBNull)
                        End If

                    Else
                        insertCommand.Parameters.AddWithValue(obj.name, obj.value)
                    End If
                End If
                'insertCommand.Parameters.AddWithValue(obj.name, if(if(IsDBNull(obj.value), Nothing, obj.value) = Nothing, Convert.DBNull, obj.value))
                cont += 1
            End While

            insertCommand.Parameters.AddWithValue("@UsuarioModificacion", dbo.UsuarioModificacion.value)
            insertCommand.Parameters.AddWithValue("@FechaModificacion", dbo.FechaModificacion.value)

            insertCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            dbo.resetKey()
            Return False
            'Throw
        Finally
            If trans Is Nothing Then
                If BD.transaction Is Nothing Then Cerrar()
            End If
        End Try
    End Function

    Protected Function UpdateProcedure(ByRef dbo As DataBussines, Optional ByRef trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean
        Try
            If trans Is Nothing Then
                If BD.transaction Is Nothing Then Conectar()
            End If
            Dim connection As System.Data.SqlClient.SqlConnection = BD.Cnx
            Dim updateCommand As New System.Data.SqlClient.SqlCommand(updateProcedureName, connection)
            updateCommand.CommandType = CommandType.StoredProcedure
            If Not trans Is Nothing Then
                updateCommand.Transaction = trans
            Else
                If Not BD.transaction Is Nothing Then
                    updateCommand.Transaction = BD.transaction
                End If
            End If

            Dim obj As DataBussinesParameter
            Dim cont As Integer = 1

            While cont <= dbo.count
                obj = dbo.item(cont)
                'updateCommand.Parameters.AddWithValue(obj.name, if(if(IsDBNull(obj.value), Nothing, obj.value) = Nothing, Convert.DBNull, obj.value))

                If IsDBNull(obj.value) Then
                    updateCommand.Parameters.AddWithValue(obj.name, Convert.DBNull)
                Else
                    If obj.value = Nothing Then
                        If TypeOf obj.value Is Boolean Then
                            updateCommand.Parameters.AddWithValue(obj.name, False)
                        Else
                            updateCommand.Parameters.AddWithValue(obj.name, Convert.DBNull)
                        End If

                    Else
                        updateCommand.Parameters.AddWithValue(obj.name, obj.value)
                    End If
                End If

                cont += 1
            End While

            updateCommand.Parameters.AddWithValue("@UsuarioModificacion", dbo.UsuarioModificacion.value)
            updateCommand.Parameters.AddWithValue("@FechaModificacion", dbo.FechaModificacion.value)
            updateCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Return False
            'Throw
        Finally
            If trans Is Nothing Then
                If BD.transaction Is Nothing Then Cerrar()
            End If
        End Try
    End Function

    Protected Function DeleteProcedure(ByRef dbo As DataBussines, Optional ByRef trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean
        Dim deleteCommand As System.Data.SqlClient.SqlCommand

        Try
            dtb.Conectar()

            deleteCommand = dtb.Comando(deleteProcedureName)
            deleteCommand.CommandType = CommandType.StoredProcedure

            If Not trans Is Nothing Then
                deleteCommand.Transaction = dtb.Transaccion
            End If

            deleteCommand.Parameters.AddWithValue(dbo.searchKey.name, dbo.searchKey.value)

            deleteCommand.ExecuteNonQuery()

            Return True
        Catch ex As Exception
            Return False
        Finally
            dtb.Desconectar()
        End Try
    End Function

    Public Overridable Function Grabar(ByRef dbo As DataBussines, Optional ByRef trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean
        If IsNothing(dbo.key.value) Or dbo.key.value = 0 Then
            Return InsertProcedure(dbo, trans)
        Else
            Return UpdateProcedure(dbo, trans)
        End If
    End Function

    Public Function select_Dgv() As DataTable
        dtb.PrepararConsulta(Me.selectDgvAllProcedureName)
        Return dtb.Consultar
        'Return dtb.Consultar(Me.selectDgvAllProcedureName, True)
    End Function

    Public Function select_DgvBy(ByVal searchTerm As String) As DataTable
        dtb.PrepararConsulta(Me.selectDgvByProcedureName & " '" & searchTerm & "'")
        Return dtb.Consultar
        'Return dtb.Consultar(Me.selectDgvByProcedureName & " '" & searchTerm & "'", True)
    End Function

    Public MustOverride Function Delete(ByVal id As Integer, Optional ByRef trans As System.Data.SqlClient.SqlTransaction = Nothing) As Boolean

End Class
