Public Class glassPanel
    Private padre As Form
    Public Sub New(ByRef frm As Form)
        InitializeComponent()

        padre = frm
    End Sub
    Public Sub recolocar(sender As Object, e As EventArgs)
        Me.Size = New Size(padre.Size.Width + 200, padre.Size.Height + 100)
        Me.Location = New Point(padre.Location.X - 100, padre.Location.Y - 50)
        Me.Visible = True
    End Sub
    Public Sub ocultar(sender As Object, e As EventArgs)
        Me.Visible = False
    End Sub

    Private Sub glassPanel_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim result As DialogResult = padre.ShowDialog
        Me.DialogResult = result
    End Sub
End Class