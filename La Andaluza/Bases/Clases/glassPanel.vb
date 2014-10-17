Public Class glassPanel
    Private padre As Form
    Private pantalla As Pantalla
    Private izq As Screen
    Private der As Screen

    Public Sub New(ByRef frm As Form)
        InitializeComponent()

        padre = frm
        pantalla = New Pantalla

        If pantalla.isMultiScreen Then
            If Screen.AllScreens(0).Bounds.X = 0 Then
                izq = Screen.AllScreens(0)
                der = Screen.AllScreens(1)
            Else
                izq = Screen.AllScreens(1)
                der = Screen.AllScreens(0)
            End If
        End If
    End Sub
    Public Sub recolocar(sender As Object, e As EventArgs)
        'Me.Size = New Size(padre.Size.Width + 200, padre.Size.Height + 100)
        'Me.Location = New Point(padre.Location.X - 100, padre.Location.Y - 50)
        'Me.FormBorderStyle = Windows.Forms.FormBorderStyle.FixedSingle       
        Me.Visible = True

        'Me.Location = New Point(padre.Location.X - 100, padre.Location.Y - 75)

        'Me.Size = New Size(padre.Size.Width + 200, padre.Size.Height + 150)

        If pantalla.isMultiScreen Then
            If 0 <= padre.Location.X And padre.Location.X < der.Bounds.Location.X Then

                Me.Location = New Point(izq.WorkingArea.Location.X, izq.WorkingArea.Location.Y)
                Me.Size = izq.WorkingArea.Size
            Else
                Me.Location = New Point(der.WorkingArea.Location.X, der.WorkingArea.Location.Y)
                Me.Size = der.WorkingArea.Size
            End If
        Else

            Me.Location = New Point(My.Computer.Screen.WorkingArea.Location.X, My.Computer.Screen.WorkingArea.Location.Y)
            Me.Size = My.Computer.Screen.WorkingArea.Size
        End If

        'Pantalla.centerIn(Me.Panel1, Me)
        'Pantalla.centerIn(Me.Panel1, Me)

        ' Me.WindowState = FormWindowState.Maximized
        Me.FormBorderStyle = Windows.Forms.FormBorderStyle.None
        padre.Focus()
    End Sub
    Public Sub ocultar(sender As Object, e As EventArgs)
        Me.Visible = False
        ' Me.WindowState = FormWindowState.Normal
        ' Me.Size = New Size(300, 300)
    End Sub

    Private Sub glassPanel_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Dim result As DialogResult = padre.ShowDialog
        padre.Focus()
        Me.DialogResult = result
    End Sub
End Class