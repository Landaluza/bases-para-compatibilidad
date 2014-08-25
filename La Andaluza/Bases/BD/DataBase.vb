
Public Class DataBase    
    Public Const LOCAL As Integer = 0
    Public Const SERVIDOR As Integer = 1
    Public Const DEFAULT_TIMEOUT As Integer = 120
    Private Cnn As System.Data.SqlClient.SqlConnection
    Private transaction As System.Data.SqlClient.SqlTransaction
    Private time_out As Integer
    Private command As System.Data.SqlClient.SqlCommand

    Public Property TimeOut As Integer
        Get
            Return Me.time_out
        End Get
        Set(value As Integer)
            Me.time_out = value
        End Set
    End Property

    Public ReadOnly Property Conexion As System.Data.SqlClient.SqlConnection
        Get
            Return Me.Cnn
        End Get
    End Property

    Public ReadOnly Property Transaccion As System.Data.SqlClient.SqlTransaction
        Get
            Return Me.transaction
        End Get
    End Property

    Public Function Comando(ByVal procedure As String) As System.Data.SqlClient.SqlCommand
        Dim sc As System.Data.SqlClient.SqlCommand = New System.Data.SqlClient.SqlCommand(procedure, Me.Cnn)
        If Not Me.transaction Is Nothing Then sc.Transaction = Me.Transaccion
        sc.CommandTimeout = Me.time_out

        Return sc
    End Function

    Public Shared Sub buildConnectionString(ByVal server As Integer)
        Dim Name As String
        Dim bd As String

        bd = "LA"

        If server = SERVIDOR Then
            Name = "192.168.1.101\SQLEXPRESS,1608"
            'Name = "192.168.1.200"
            Config.connectionString = "User ID=ssa;Password=Trucha0122;Trusted_Connection=False;"
        Else

            If server = LOCAL Then
                Config.connectionString = "User ID=ssa;Password=Trucha0122;Trusted_Connection=False;"
                Name = My.Computer.Name & "\SQLEXPRESS" ' Name = "VMDESARROLLO\SQLEXPRESS"
            Else
                If My.Computer.Name = "MAMVAIO" Then
                    Config.connectionString = "User ID=mamvaio\mam;Trusted_Connection=True;"
                    Name = "MAMVAIO\SQLEXPRESS"
                Else
                    Name = "192.168.1.124\SQLEXPRESS"
                    Config.connectionString = "User ID=ssa;Password=Trucha0122;Trusted_Connection=False;"
                End If
            End If
        End If

        Config.connectionString = "workstation id=" & Name & ";packet size=4096;Connect Timeout = 200;" & Config.connectionString & "data source= " _
            & Name & ";persist security info=False;initial catalog=" & bd

        'My.Settings.Item("LAConnectionString1") = Config.connectionString

    End Sub

    Public Sub New(ByVal source As Integer)
        If Config.connectionString = String.Empty Then
            DataBase.buildConnectionString(source)
        End If

        Me.time_out = DEFAULT_TIMEOUT
    End Sub

    Public Sub New(ByVal source As Integer, ByRef Cnn As System.Data.SqlClient.SqlConnection, ByRef trans As System.Data.SqlClient.SqlTransaction)
        Me.New(source)

        Me.Cnn = Cnn
        Me.transaction = trans
        Me.time_out = DEFAULT_TIMEOUT
    End Sub

    Public Sub Conectar()
        If Me.transaction Is Nothing Then
            If Not BD.transaction Is Nothing Then
                Me.Cnn = BD.Cnx
                Me.transaction = BD.transaction
            Else
                If Not Me.Cnn Is Nothing Then Me.Desconectar()
                Me.Cnn = New System.Data.SqlClient.SqlConnection
                Me.Cnn.ConnectionString = Config.connectionString
                Me.Cnn.Open()
            End If
            'Else
            '    Me.Cnn = BD.Cnx
            '    Me.transaction = BD.transaction
        End If
    End Sub

    Public Sub Desconectar()
        If transaction Is Nothing Then
            If Not Me.Cnn Is Nothing Then
                If Cnn.State = System.Data.ConnectionState.Open Then
                    Cnn.Close()
                    Cnn.Dispose()
                End If
            End If
        End If
    End Sub

    Public Sub EmpezarTransaccion()
        Me.Conectar()
        Me.transaction = Me.Cnn.BeginTransaction()
    End Sub

    Public Sub TerminarTransaccion()
        Me.transaction.Commit()
        Me.transaction = Nothing
        Me.Desconectar()
    End Sub

    Public Sub CancelarTransaccion()
        Me.transaction.Rollback()
        Me.transaction = Nothing
        Me.Desconectar()
    End Sub

    Public Function Consultar(ByVal Cadena As String, Optional procedure As Boolean = True) As DataTable
        Dim dtsTabla As New DataTable

        If BD.transaction Is Nothing Then
            Me.Conectar()
        Else
            Me.Cnn = BD.Cnx
            Me.transaction = BD.transaction
        End If

        Dim dtsTemp As New DataSet
        Dim cmd As System.Data.SqlClient.SqlCommand '= Comando(Cadena)

        Try
            If procedure Then
                If Not Cadena.Contains("exec ") Then
                    Cadena = "exec " & Cadena
                End If

                cmd = Comando(Cadena)
            Else
                cmd = Comando(Cadena)
            End If

            'If procedure Then cmd.CommandType = CommandType.StoredProcedure

            Dim Ad As System.Data.SqlClient.SqlDataAdapter = New System.Data.SqlClient.SqlDataAdapter(cmd)
            Ad.Fill(dtsTemp, "NuevaTabla")
            dtsTabla = dtsTemp.Tables(0)
            'cmd.Dispose()
            Return dtsTabla
        Catch ex As Exception
            Return Nothing
        Finally

            If Me.transaction Is Nothing Then Me.Desconectar()
        End Try
    End Function

    Public Function ConsultaAlteraciones(ByVal strrealizarConsulta As String) As Boolean
        If Me.transaction Is Nothing Then
            If BD.transaction Is Nothing Then
                Me.Conectar()
            Else
                Me.Cnn = BD.Cnx
            End If
        End If

        Dim cmd As System.Data.SqlClient.SqlCommand
        cmd = New System.Data.SqlClient.SqlCommand(strrealizarConsulta, Me.Cnn)

        If Not Me.transaction Is Nothing Then
            cmd.Transaction = Me.transaction
        Else
            If Not BD.transaction Is Nothing Then
                cmd.Transaction = BD.transaction
            End If
        End If

        Try
            cmd.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
        Finally
            cmd.Dispose()
            If Me.transaction Is Nothing Then
                If BD.transaction Is Nothing Then Me.Desconectar()
            End If
        End Try
    End Function

    Public Sub PrepararConsulta(ByVal consulta As String)
        Me.Conectar()
        command = Comando(consulta)
    End Sub

    Public Sub AñadirParametroConsulta(ByVal nombre As String, ByVal valor As Object)
        command.Parameters.AddWithValue(nombre, valor)
    End Sub

    Public Function Consultar() As DataTable
        Dim dtsTabla As New DataTable
        Dim dtsTemp As New DataSet
        Dim Ad As System.Data.SqlClient.SqlDataAdapter

        Try
            Ad = New System.Data.SqlClient.SqlDataAdapter(command)
            Ad.Fill(dtsTemp, "NuevaTabla")
            dtsTabla = dtsTemp.Tables(0)
            'cmd.Dispose()
            Return dtsTabla
        Catch ex As Exception
            Return Nothing
        Finally
            Me.Desconectar()
        End Try
    End Function
End Class
