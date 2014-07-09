Public Class frmEspera

    Public Sub New(Optional ByVal mensaje As String = "")

        InitializeComponent()


        Me.lMensaje.Text = mensaje
        'Me.Width = lMensaje.Width + 100
        Me.TopMost = True
    End Sub
End Class