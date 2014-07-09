Public Class DataBussines
    Private m_atributos As Collection
    Private m_FechaModificacion As DataBussinesParameter
    Private m_UsuarioModificacion As DataBussinesParameter
    Private m_primaryKey As DataBussinesParameter
    Private m_searchKey As DataBussinesParameter

    Public Sub New()
        m_atributos = New Collection
        m_FechaModificacion = New DataBussinesParameter("FechaModificacion", "@FechaModificacion")
        m_UsuarioModificacion = New DataBussinesParameter("UsuarioModificacion", "@UsuarioModificacion")
    End Sub

    Public ReadOnly Property count() As Integer
        Get
            Return m_atributos.Count
        End Get
    End Property

    Public ReadOnly Property item(ByVal index As Integer) As DataBussinesParameter
        Get
            Return CType(atributos.Item(index), DataBussinesParameter)
        End Get
    End Property

    Public ReadOnly Property item(ByVal indexName As String) As DataBussinesParameter
        Get
            Return CType(atributos.Item(indexName), DataBussinesParameter)
        End Get
    End Property

    Protected Property atributos() As Collection
        Get
            Return m_atributos
        End Get
        Set(ByVal value As Collection)
            m_atributos = value
        End Set
    End Property

    Protected Property primaryKey() As DataBussinesParameter
        Get
            Return m_primaryKey
        End Get
        Set(ByVal value As DataBussinesParameter)
            m_primaryKey = value
        End Set
    End Property

    Public Property searchKey() As DataBussinesParameter
        Get
            Return if(m_searchKey Is Nothing, primaryKey, m_searchKey)
        End Get
        Set(ByVal value As DataBussinesParameter)
            m_searchKey = value
        End Set
    End Property

    Public ReadOnly Property key() As DataBussinesParameter
        Get
            If IsDBNull(primaryKey.value) Then primaryKey.value = Nothing
            Return primaryKey
        End Get
    End Property

    Public ReadOnly Property FechaModificacion() As Object
        Get
            m_FechaModificacion.value = Today
            Return m_FechaModificacion
        End Get
    End Property

    Public ReadOnly Property UsuarioModificacion() As Object
        Get
            m_UsuarioModificacion.value = Config.User
            Return m_UsuarioModificacion
        End Get
    End Property

    Public Sub resetKey()
        For Each a As DataBussinesParameter In m_atributos
            If Me.m_primaryKey.name = a.name Then
                a.value = Nothing
                Exit For
            End If
        Next

        Me.m_primaryKey.value = Nothing
    End Sub
End Class
