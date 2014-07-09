Public Class frmEntrada
    Private entradaUsuario As String

    Public ReadOnly Property Result As String
        Get
            Return Me.entradaUsuario
        End Get
    End Property

    Public Sub New(ByVal titulo As String, ByVal TextoEtiqueta As String, Optional ByVal textoCampo As String = "", Optional ancho As Integer = 155)
        InitializeComponent()


        Me.lentrada.Text = TextoEtiqueta
        Me.txtEntrada.Text = textoCampo
        Me.Width = ancho
    End Sub

    Private Sub btnAceptar_Click(sender As System.Object, e As System.EventArgs) Handles btnAceptar.Click
        Me.entradaUsuario = Me.txtEntrada.Text
        Me.DialogResult = Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
End Class