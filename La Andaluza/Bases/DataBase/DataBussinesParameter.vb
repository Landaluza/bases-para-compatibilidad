Public Class DataBussinesParameter
    Private m_value As Object
    Private m_name As String
    Private m_sqlName As String
    Private m_insertable As Boolean
    Private m_updatable As Boolean

    Public Sub New()

    End Sub

    Public Sub New(ByRef obj As Object, ByVal name As String, ByVal sqlName As String)
        Me.m_name = name
        Me.sqlName = if(sqlName = "", name.Substring(1), sqlName)
        Me.m_value = obj
    End Sub

    Public Sub New(ByVal name As String, ByVal sqlName As String)
        Me.m_name = name
        Me.sqlName = if(sqlName = "", name.Substring(1), sqlName)
        Me.m_value = Convert.DBNull
    End Sub

    Public Property name() As String
        Get
            Return m_name
        End Get
        Set(ByVal value As String)
            m_name = value
        End Set
    End Property

    Public Property sqlName() As String
        Get
            Return m_sqlName
        End Get
        Set(ByVal value As String)
            m_sqlName = value
        End Set
    End Property

    Public Property value() As Object
        Get
            Return m_value
        End Get
        Set(ByVal value As Object)
            m_value = value
        End Set
    End Property

    Public Sub reset()
        Me.m_value = Nothing
    End Sub
End Class
