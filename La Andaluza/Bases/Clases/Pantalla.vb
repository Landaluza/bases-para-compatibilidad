Public Class Pantalla
    Public Const CENTRADA As Integer = 0
    Public Const MARGEN_IZQUIERDO As Integer = 1
    Public Const MARGERN_DERECHO As Integer = 2
    Public Const ESQUINA_SUPERIOR_IZQUIERDA As Integer = 3
    Public Const ESQUINA_SUPERIOR_DERECHA As Integer = 4
    Public Const ESQUINA_INFERIOR_IZQUIERDA As Integer = 5
    Public Const ESQUINA_INFERIOR_DERECHA As Integer = 6
    Public Const MAXIMIZADA As Integer = 9

    Private screens() As Screen


    Public Function isMultiScreen() As Boolean
        If Screen.AllScreens.Length > 1 Then
            Return True
        Else : Return False
        End If
    End Function

    Public Sub reubicar_formulario_en_pantalla_secundaria(ByRef form As Windows.Forms.Form, Optional ByVal posicion As Integer = CENTRADA)
        Dim scrn As New Pantalla()
        scrn.ubicar(1, form, posicion)
    End Sub

    Public Sub reubicar_formulario_en_pantalla_primaria(ByRef form As Windows.Forms.Form, Optional ByVal posicion As Integer = CENTRADA)
        Dim scrn As New Pantalla()
        scrn.ubicar(0, form, posicion)
    End Sub

    Public Function mostrar_formulario_en_pantalla_secundaria(ByRef form As Windows.Forms.Form, ByVal modal As Boolean, Optional ByVal posicion As Integer = CENTRADA) As Object
        Dim scrn As New Pantalla()
        Return scrn.mostrar_formulario(1, form, modal, posicion)
    End Function

    Public Function mostrar_formulario_en_pantalla_primaria(ByRef form As Windows.Forms.Form, ByVal modal As Boolean, Optional ByVal posicion As Integer = CENTRADA) As Object
        Dim scrn As New Pantalla()
        Return scrn.mostrar_formulario(0, form, modal, posicion)
    End Function

    Public Sub New()
        Me.screens = Screen.AllScreens()
    End Sub

    Private Function mostrar_formulario(ByVal pantalla As Integer, ByRef form As Windows.Forms.Form, ByVal modal As Boolean, ByVal posicion As Integer) As DialogResult
        Me.ubicar(pantalla, form, posicion)

        If modal Then
            Return form.ShowDialog()
        Else
            form.Show()
            Return DialogResult.OK
        End If
    End Function


    Private Sub ubicar(ByVal pantalla As Integer, ByRef form As Windows.Forms.Form, ByVal posicion As Integer)
        Dim screen As Screen
        ' mostramos el formualrio en la pantalla secundaria
        screen = Me.screens(pantalla)

        ' establecemos la posicion inicial en manual

        form.StartPosition = FormStartPosition.Manual

        ' colocamos el formualrio en la posicion 100, 100 en el monitor secundario
        Select Case posicion
            Case CENTRADA
                form.Location = New Point(screen.Bounds.Location.X + CType((screen.Bounds.Width - form.Width) / 2, Integer), screen.Bounds.Location.Y + CType((screen.Bounds.Height - form.Height) / 2, Integer))

            Case MARGERN_DERECHO
                form.Location = New Point(screen.Bounds.Location.X + screen.Bounds.Width - form.Width, screen.Bounds.Location.Y + CType((screen.Bounds.Height - form.Height) / 2, Integer))

            Case MARGEN_IZQUIERDO
                form.Location = New Point(screen.Bounds.Location.X, screen.Bounds.Location.Y + CType((screen.Bounds.Height - form.Height) / 2, Integer))

            Case ESQUINA_INFERIOR_DERECHA
                form.Location = New Point(screen.Bounds.Location.X + screen.Bounds.Width - form.Width, screen.Bounds.Location.Y + screen.Bounds.Height - form.Height)

            Case ESQUINA_INFERIOR_IZQUIERDA
                form.Location = New Point(screen.Bounds.Location.X, screen.Bounds.Location.Y + screen.Bounds.Height - form.Height)

            Case ESQUINA_SUPERIOR_DERECHA
                form.Location = New Point(screen.Bounds.Location.X + screen.Bounds.Width - form.Width, screen.Bounds.Location.Y)

            Case ESQUINA_SUPERIOR_IZQUIERDA
                form.Location = New Point(screen.Bounds.Location.X, screen.Bounds.Location.Y)

            Case MAXIMIZADA
                form.Location = screen.Bounds.Location
                form.Width = screen.WorkingArea.Width
                form.Height = screen.WorkingArea.Height
        End Select
    End Sub

    Public Sub fadeInPanel(ByRef panel1 As Panel)
        Dim Opacity2 As Double
        Dim FadeCount1 As Integer

        Dim h As Integer = panel1.Height
        Dim w As Integer = panel1.Width

        panel1.Width = 10
        panel1.Height = 10

        For FadeCount1 = 10 To 90 Step 5
            Opacity2 = FadeCount1 / 100
            panel1.BackColor = Color.FromArgb(CType(Opacity2, Integer))
            panel1.Width = CType(w / (100 - FadeCount1), Integer)
            panel1.Height = CType(h / (100 - FadeCount1), Integer)
            Threading.Thread.Sleep(25)
        Next

        panel1.Width = w
        panel1.Height = h
        panel1.Visible = True
    End Sub

    Public Sub fadeOutPanel(ByRef panel1 As Panel)
        Dim FadeCount As Integer
        Dim Opacity1 As Double

        For FadeCount = 90 To 10 Step -5
            Opacity1 = FadeCount / 100
            panel1.BackColor = Color.FromArgb(CType(Opacity1, Integer))
            Threading.Thread.Sleep(25)
            System.Windows.Forms.Application.DoEvents()
        Next

        panel1.Visible = False
    End Sub

    Public Sub fadePanelTo(ByRef panel1 As Panel, ByRef panel2 As Panel)
        Dim Opacity2 As Double
        Dim FadeCount1 As Integer
        Dim FadeCount As Integer
        Dim Opacity1 As Double

        For FadeCount = 90 To 10 Step -5
            Opacity1 = FadeCount / 100
            panel2.BackColor = Color.FromArgb(CType(Opacity1, Integer))
            panel1.Refresh()
            Threading.Thread.Sleep(25)
            System.Windows.Forms.Application.DoEvents()
        Next

        For FadeCount1 = 10 To 90 Step 5
            Opacity2 = FadeCount1 / 100
            panel1.BackColor = Color.FromArgb(CType(Opacity2, Integer))
            panel2.Refresh()
            Threading.Thread.Sleep(25)
        Next

        panel1.Visible = True
        panel2.Visible = False

    End Sub

    Public Function capturarPantalla(ByRef frm As Form) As Bitmap

        Dim gr As Graphics = frm.CreateGraphics
        ' Tamaño de lo que queremos copiar
        Dim fSize As Size = frm.Size
        ' Creamos el bitmap con el área que vamos a capturar
        ' En este caso, con el tamaño del formulario actual
        Dim bm As New Bitmap(fSize.Width, fSize.Height, gr)
        ' Un objeto Graphics a partir del bitmap
        Dim gr2 As Graphics = Graphics.FromImage(bm)
        ' Copiar el área de la pantalla que ocupa el formulario
        gr2.CopyFromScreen(frm.Location.X, frm.Location.Y, 0, 0, fSize)

        Return bm

    End Function

    Public Function capturarPantallaEntera(ByRef frm As Form) As Bitmap
        ' Capturar toda la pantalla

        Dim gr As Graphics = frm.CreateGraphics
        ' Tamaño de lo que queremos copiar
        ' En este caso el tamaño de la ventana principal
        Dim fSize As Size = Screen.PrimaryScreen.Bounds.Size
        ' Creamos el bitmap con el área que vamos a capturar
        Dim bm As New Bitmap(fSize.Width, fSize.Height, gr)
        ' Un objeto Graphics a partir del bitmap
        Dim gr2 As Graphics = Graphics.FromImage(bm)
        ' Copiar todo el área de la pantalla
        gr2.CopyFromScreen(0, 0, 0, 0, fSize)

        ' Asignamos la imagen al PictureBox
        Return bm
    End Function
End Class
