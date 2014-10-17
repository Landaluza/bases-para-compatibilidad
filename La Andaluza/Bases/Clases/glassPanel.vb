Public Class glassPanel
    Private padre As Form
    Public Sub New(ByRef frm As Form)
        InitializeComponent()

        padre = frm
    End Sub
    Public Sub recolocar(sender As Object, e As EventArgs)
        'Me.Size = New Size(padre.Size.Width + 200, padre.Size.Height + 100)
        'Me.Location = New Point(padre.Location.X - 100, padre.Location.Y - 50)
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle

        Me.Location = New Point(padre.Location.X, padre.Location.Y)
        
        Me.Visible = True
        Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        padre.Focus()
    End Sub
    Public Sub ocultar(sender As Object, e As EventArgs)
        Me.Visible = False        
        Me.WindowState = FormWindowState.Normal
        Me.Size = padre.Size
    End Sub

    Private Sub glassPanel_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim result As DialogResult = padre.ShowDialog
        padre.Focus()
        Me.DialogResult = result
    End Sub
End Class