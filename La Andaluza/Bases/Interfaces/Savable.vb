Public Interface Savable
    Event afterSave(sender As Object, args As EventArgs) '
    Sub setValores()
    Function getValores() As Boolean
    Sub Guardar(Optional ByRef dtb As DataBase = Nothing)
End Interface
