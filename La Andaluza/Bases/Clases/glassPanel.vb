Public Class glassPanel

    Public Sub recolocar(sender As Object, e As EventArgs)
        Dim frm As Form = sender
        Me.SetDesktopLocation(frm.Location.X, frm.Location.Y)
    End Sub
End Class