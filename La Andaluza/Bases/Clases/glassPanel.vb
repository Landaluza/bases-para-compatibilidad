Public Class glassPanel
    Private padre As Form
    Private cont As Integer
    Public Sub New(ByRef frm As Form)
        InitializeComponent()

        padre = frm
        cont = 0
    End Sub
    Public Sub recolocar(sender As Object, e As EventArgs)
        Me.Size = New Size(padre.Size.Width + 200, padre.Size.Height + 100)
        Me.Location = New Point(padre.Location.X - 100, padre.Location.Y - 50)
    End Sub
End Class