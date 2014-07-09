Public Interface Savable
    Event afterSave(sender As Object, args As EventArgs) '
    Sub setValores()
    Function getValores() As Boolean
    Sub Guardar(Optional ByRef trans As SqlClient.SqlTransaction = Nothing)
End Interface
