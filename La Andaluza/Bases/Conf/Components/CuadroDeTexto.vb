

' Specifies a type that indicates the assembly to search, and the name 
' of an image resource to look for.



' Visual Basic
' Specifies the bitmap associated with the Button type.
'<ToolboxBitmap(GetType(Button))> Class MyControl1
'    ' Specifies a bitmap file.
'End Class
'<ToolboxBitmap("C:\Documents and Settings\Administrador.ANDALUZA\Escritorio\2009.11.08 Generador\Imagenes\Generador.ico")> _
'   Class MyControl1
'End Class

'ToolboxBitmap(typeof(ClockControlLibrary.ClockControl), "images.ClockControl.ico"),

<ToolboxBitmap("CuadroDeTexto.bmp")> _
Public Class CuadroDeTexto
    Inherits System.Windows.Forms.TextBox

    Private Local_Modificado As Boolean
    Private Local_Minimo As Integer
    Private Local_Obligatorio As Boolean
    Private Local_EsUnicoCampo As String
    Private Local_EsUnicoCampoID As String
    Private Local_EsUnicoID As Integer
    Private Local_EsUnicoTabla As String

    Private Local_ParametroID As Integer
    Private Local_ValorMaximo As Double
    Private Local_ValorMinimo As Double


    Private Local_cantDoublees As Integer
    Private Local_Numerico_SeparadorMiles As Boolean
    Private Local_Numerico As Boolean

    Private bandera As Boolean = False
    Private bandera2 As Boolean = False
    Private band As Boolean = True
    '----------------------------------------------------------------------------------------
    Public Overrides Property Text As String
        Get
            If Me.Local_Numerico_SeparadorMiles Then
                Return MyBase.Text.Replace(".", "")
            Else
                Return MyBase.Text
            End If
        End Get
        Set(value As String)
            MyBase.Text = value
        End Set
    End Property

    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        If Local_Numerico Then
            If e.KeyChar = Environment.NewLine Then
                SendKeys.Send("{TAB}")
            Else
                Local_Modificado = True
            End If

            Dim Keystroke As Integer = Asc(e.KeyChar)

            If Keystroke = 46 Then
                SendKeys.Send(",")
            ElseIf Keystroke = 44 And Convert.ToBoolean(InStr(Me.Text, ",")) Then
                e.Handled = True
                Return

            ElseIf Keystroke = 45 And Len(Me.Text) > 0 Then
                e.Handled = True
                Return
            ElseIf (Keystroke < 48 Or Keystroke > 57) And _
              (Keystroke <> 45 And Keystroke <> 44 And Keystroke <> 8) Then
                e.Handled = True
                Return
            End If

        Else
            If e.KeyChar = Environment.NewLine Then
                SendKeys.Send("{TAB}")
            Else
                Local_Modificado = True
            End If
        End If

    End Sub

    Public Sub New()
        Local_Modificado = False
        Me.TextAlign = HorizontalAlignment.Left
        Local_Obligatorio = False
        Local_Minimo = 0
        Local_EsUnicoCampo = ""
        Local_EsUnicoCampoID = ""
        Local_EsUnicoID = 0
        Local_EsUnicoTabla = ""
        Local_ParametroID = 0
        Local_ValorMaximo = 0
        Local_ValorMinimo = 0

        'numerico
        Local_Numerico = False
        Local_cantDoublees = 0
    End Sub

    Public Sub EsModificado(ByVal valor As Boolean)
        Local_Modificado = valor
    End Sub

    Public Sub New(ByVal ParametroId As Integer, ByVal Obligatorio As Boolean, ByVal Minimo As Integer, ByVal EsUnicoCampo As String, ByVal EsUnicoCampoID As String, ByVal EsUnicoID As Integer, ByVal EsUnicoTabla As String, ByVal ValorMaximo As Double, ByVal ValorMinimo As Double, _
    ByVal cantDoublees As Integer, ByVal Numerico_SeparadorMiles As Boolean, ByVal oblig As Boolean, ByVal Numerico As Boolean)

        Local_Obligatorio = Obligatorio
        Local_Minimo = Minimo
        Local_EsUnicoCampo = EsUnicoCampo
        Local_EsUnicoCampoID = EsUnicoCampoID
        Local_EsUnicoID = EsUnicoID
        Local_EsUnicoTabla = EsUnicoTabla
        Local_ParametroID = ParametroId
        Local_ValorMaximo = ValorMaximo
        Local_ValorMinimo = ValorMinimo

        Local_cantDoublees = cantDoublees
        Local_Numerico_SeparadorMiles = Numerico_SeparadorMiles
        Obligatorio = oblig
        Local_Numerico = Numerico
        If Local_Numerico Then
            Me.TextAlign = HorizontalAlignment.Right
        Else
            Me.TextAlign = HorizontalAlignment.Left
        End If
    End Sub


    Protected Overrides Sub onTextChanged(ByVal e As System.EventArgs)
        If band Then
            band = False
            If Local_Numerico Then
                Dim aux As Double
                Dim nfi As System.Globalization.NumberFormatInfo = New System.Globalization.CultureInfo("es-ES", False).NumberFormat
                If Me.Text <> "" Then
                    Me.Text = Me.Text.Replace(".", "")
                    aux = Double.Parse(Me.Text)
                    Dim formatstr As String = "N" & Local_cantDoublees.ToString.Trim
                    If Local_Numerico_SeparadorMiles Then
                        Me.Text = aux.ToString(formatstr, nfi)
                    Else
                        'Me.Text = aux.ToString(formatstr)
                    End If
                End If
            End If
            MyBase.OnTextChanged(e)
            band = True
        Else
            MyBase.OnTextChanged(e)
        End If

    End Sub

    Protected Overrides Sub OnGotFocus(ByVal e As System.EventArgs)
        band = False
        Dim ColorObtenerFoco As String = System.Configuration.ConfigurationManager.AppSettings("ColorObtenerFoco")
        Me.BackColor = Color.FromName(ColorObtenerFoco)
        MyBase.OnGotFocus(e)
    End Sub

    Protected Overrides Sub OnLostFocus(ByVal e As System.EventArgs)
        Try
            If Local_Numerico Then
                Dim aux As Double
                Dim nfi As System.Globalization.NumberFormatInfo = New System.Globalization.CultureInfo("es-ES", False).NumberFormat
                If Me.Text <> "" Then
                    Me.Text = Me.Text.Replace(".", "")
                    aux = Double.Parse(Me.Text)
                    Dim formatstr As String = "N" & Local_cantDoublees.ToString.Trim
                    If Local_Numerico_SeparadorMiles Then
                        Me.Text = aux.ToString(formatstr, nfi)
                    Else
                        'Me.Text = aux.ToString(formatstr)
                    End If
                End If
            End If
            Dim ColorPerderFoco As String = System.Configuration.ConfigurationManager.AppSettings("ColorPerderFoco")
            Me.BackColor = Color.FromName(ColorPerderFoco)
            MyBase.OnLostFocus(e)
            band = True
        Catch ex As Exception
        End Try
    End Sub

    'Public Sub validarCodigoLote()
    '    Try
    '        If Local_EsUnicoCampo <> "" And Local_EsUnicoCampoID <> "" And Local_EsUnicoTabla <> "" Then
    '            Dim ctl As New ctlCuadroDeTexto
    '            If ctl.Validar(Local_EsUnicoID, Local_EsUnicoCampo, Me.Text, Local_EsUnicoCampoID, Local_EsUnicoTabla) Then
    '                Dim aux As String
    '                aux = Me.Text.Substring(0, 14)


    '                Dim c As Char = Me.Text.Substring(14, 1)
    '                Dim digit As Integer = Asc(c)
    '                If digit = 57 Then digit = 64
    '                digit += 1
    '                c = ChrW(digit)
    '                aux = aux & c


    '                If aux.Length > 15 Then
    '                    messagebox.show("Solo se permite 9 lotes por dia.", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    '                End If
    '                Me.Text = aux
    '                validarCodigoLote()
    '            End If
    '        End If
    '    Catch ex As Exception
    '    End Try

    'End Sub

    Public Sub validarCodigoBota()
        If Local_EsUnicoCampo <> "" And Local_EsUnicoCampoID <> "" And Local_EsUnicoTabla <> "" Then
            Dim ctl As New ctlCuadroDeTexto
            If ctl.Validar(Local_EsUnicoID, Local_EsUnicoCampo, Me.Text, Local_EsUnicoCampoID, Local_EsUnicoTabla) Then
                MessageBox.Show("Esa posicion de bota ya se encuentra ocupada", "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            End If
        End If
    End Sub



    Protected Overrides Sub OnValidating(ByVal e As System.ComponentModel.CancelEventArgs)
        bandera = False
        Dim Razon As String = ""
        If Local_Obligatorio Then
            If Me.Text = "" Then
                Razon = "No dejar vacio el campo es obligatorio."
                e.Cancel = True
            End If
        End If
        If Local_Minimo <> 0 Then
            If Me.TextLength < Minimo Then
                Razon = Razon & " El minimo de caracteres debe ser de " & Convert.ToString(Local_Minimo) & " caracteres."
                e.Cancel = True
            End If
        End If
        If Local_EsUnicoCampo <> "" And Local_EsUnicoCampoID <> "" And Local_EsUnicoTabla <> "" Then
            Dim ctl As New ctlCuadroDeTexto
            If ctl.Validar(Local_EsUnicoID, Local_EsUnicoCampo, Me.Text, Local_EsUnicoCampoID, Local_EsUnicoTabla) Then
                If Me.Name <> "txtCodigoLote" Then
                    Razon = Razon & " Este campo debe ser unico y ya se encuentra repetido en la bd."
                    e.Cancel = True
                End If
            End If
        End If

        If Local_ValorMinimo <> 0 Then
            If Me.Text = "" Then
                e.Cancel = True
                Razon = Razon & " Este campo admite un valor minimo de " & Local_ValorMinimo & ", no se puede dejar en blanco."
            Else
                If Me.Text < Local_ValorMinimo Then
                    e.Cancel = True
                    Razon = Razon & " Este campo admite un valor minimo de " & Local_ValorMinimo & "."
                End If
            End If
        End If

        If Local_ValorMaximo <> 0 Then
            If Me.Text = "" Then
                Me.Text = 0
                e.Cancel = True
            Else
                If Me.Text > Local_ValorMaximo Then
                    e.Cancel = True
                    Razon = Razon & " Este campo admite un valor maximo de " & Local_ValorMaximo & "."
                End If
            End If

        End If
        If Razon <> "" Then
            messagebox.show(Razon, "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            Me.Modified = Local_Modificado
        End If
        MyBase.OnValidating(e)
    End Sub

    Public Property Modificado() As Boolean
        Get
            Return Local_Modificado
        End Get
        Set(ByVal Value As Boolean)
            Local_Modificado = Value
        End Set
    End Property

    'Obligatorio. Indica si es obligatorio o no
    <System.ComponentModel.Description("Indica si es obligatorio o no"), System.ComponentModel.Category("MAM")> _
    Public Property Obligatorio() As Boolean
        Get
            Return Local_Obligatorio
        End Get
        Set(ByVal Value As Boolean)
            Local_Obligatorio = Value
        End Set
    End Property

    'Minimo. Indica si tiene minimo o no
    <System.ComponentModel.Description("Indica si tiene minimo o no"), System.ComponentModel.Category("MAM")> _
    Public Property Minimo() As Integer
        Get
            Return Local_Minimo
        End Get
        Set(ByVal Value As Integer)
            Local_Minimo = Value
        End Set
    End Property

    'EsUnicoCampo. Indica si en este campo se verificara si es unico 
    <System.ComponentModel.Description("Indica si en este campo se verificara si es unico o no"), System.ComponentModel.Category("MAM")> _
    Public Property EsUnicoCampo() As String
        Get
            Return Local_EsUnicoCampo
        End Get
        Set(ByVal Value As String)
            Local_EsUnicoCampo = Value
        End Set
    End Property

    'EsUnicoCampoID. Indica el nombre del Campo Identificador
    <System.ComponentModel.Description("Indica el nombre del Campo Identificador"), System.ComponentModel.Category("MAM")> _
    Public Property EsUnicoCampoID() As String
        Get
            Return Local_EsUnicoCampoID
        End Get
        Set(ByVal Value As String)
            Local_EsUnicoCampoID = Value
        End Set
    End Property

    'EsUnicoCampoID. Indica el ID
    <System.ComponentModel.Description("Indica el ID"), System.ComponentModel.Category("MAM")> _
    Public Property EsUnicoID() As Integer
        Get
            Return Local_EsUnicoID
        End Get
        Set(ByVal Value As Integer)
            Local_EsUnicoID = Value
        End Set
    End Property

    'EsUnicoTabla. Indica si en esta tabla se verificara el campo 
    <System.ComponentModel.Description("Indica si en esta tabla se verificara el campo"), System.ComponentModel.Category("MAM")> _
    Public Property EsUnicoTabla() As String
        Get
            Return Local_EsUnicoTabla
        End Get
        Set(ByVal Value As String)
            Local_EsUnicoTabla = Value
        End Set
    End Property

    'ValorMaximo. Indica si tendra un posible valor Maximo
    <System.ComponentModel.Description("Indica si tendra un posible valor Maximo"), System.ComponentModel.Category("MAM")> _
    Public Property ValorMaximo() As Double
        Get
            Return Local_ValorMaximo
        End Get
        Set(ByVal Value As Double)
            Local_ValorMaximo = Value
        End Set
    End Property

    'ValorMinimo. Indica si tendra un posible valor Minimo
    <System.ComponentModel.Description("Indica si tendra un posible valor Minimo"), System.ComponentModel.Category("MAM")> _
    Public Property ValorMinimo() As Double
        Get
            Return Local_ValorMinimo
        End Get
        Set(ByVal Value As Double)
            Local_ValorMinimo = Value
        End Set
    End Property

    'ParametroID. Indica si es un parametro, y cual es su ID
    <System.ComponentModel.Description("Indica si es un parametro, y cual es su ID"), System.ComponentModel.Category("MAM")> _
    Public Property ParametroID() As Integer
        Get
            Return Local_ParametroID
        End Get
        Set(ByVal Value As Integer)
            Local_ParametroID = Value
        End Set
    End Property

    'Numerico_NumeroDoublees. Selecciona numero de Doublees a mostrar en los campos numericos
    <System.ComponentModel.Description("Selecciona numero de Doublees a mostrar en los campos numericos"), System.ComponentModel.Category("MAM")> _
    Public Property Numerico_NumeroDoublees() As Integer
        Get
            Return Local_cantDoublees
        End Get
        Set(ByVal Value As Integer)
            Local_cantDoublees = Value
        End Set
    End Property

    'Numerico_SeparadorMiles. Indica si se pondra separador de miles o no
    <System.ComponentModel.Description("Indica si se pondra separador de miles o no"), System.ComponentModel.Category("MAM")> _
    Public Property Numerico_SeparadorMiles() As Boolean
        Get
            Return Local_Numerico_SeparadorMiles
        End Get
        Set(ByVal Value As Boolean)
            Local_Numerico_SeparadorMiles = Value
        End Set
    End Property

    'Numerico. Indica si el campo es numerico o no
    <System.ComponentModel.Description("Indica si el campo es numerico o no"), System.ComponentModel.Category("MAM")> _
    Public Property Numerico_EsNumerico() As Boolean
        Get
            Return Local_Numerico
        End Get
        Set(ByVal Value As Boolean)
            Local_Numerico = Value
            If Local_Numerico Then
                Me.TextAlign = HorizontalAlignment.Right
            Else
                Me.TextAlign = HorizontalAlignment.Left
            End If
        End Set
    End Property
End Class
