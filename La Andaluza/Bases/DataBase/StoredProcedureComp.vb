Public Class sp
    Private connection As System.Data.SqlClient.SqlConnection
    Protected selectProcedureName, insertProcedureName, updateProcedureName, deleteProcedureName, selectDgvAllProcedureName, selectDgvByProcedureName As String
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

    'Public Sub New()

    'End Sub
    Public Sub New(ByVal selectProcedureName As String, ByVal insertProcedureName As String, ByVal updateProcedureName As String, ByVal deleteProcedureName As String, ByVal selectDgvAllProcedureName As String, ByVal selectDgvByProcedureName As String)
        Me.selectProcedureName = selectProcedureName
        Me.insertProcedureName = insertProcedureName
        Me.updateProcedureName = updateProcedureName
        Me.deleteProcedureName = deleteProcedureName
        Me.selectDgvAllProcedureName = selectDgvAllProcedureName
        Me.selectDgvByProcedureName = selectDgvByProcedureName

        dtb = New DataBase(Config.Server)
    End Sub
End Class
