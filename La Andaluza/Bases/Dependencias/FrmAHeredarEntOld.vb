'@DEPRECATED
Public Class FrmAHeredarEntOld
    Public ParaGuardar As Boolean = False
    Public Posicion As Integer
    Public ModoDeApertura As Integer
    Public Const INSERCION As Integer = 1
    Public Const MODIFICACION As Integer = 2
    Public Const CERRADO As Integer = 3
    Public Const VISION As Integer = 99
    Protected dtb As DataBase

    Public Sub New(ByVal modo_de_apertura As Integer)
        Me.ModoDeApertura = modo_de_apertura
        InitializeComponent()
        dtb = New DataBase()
    End Sub

    Public Sub New(ByVal modo_de_apertura As String)
        Select Case modo_de_apertura
            Case FrmAheredarOld.ACCION_INSERTAR
                Me.ModoDeApertura = INSERCION
            Case FrmAheredarOld.ACCION_MODIFICAR
                Me.ModoDeApertura = MODIFICACION
            Case FrmAheredarOld.ACCION_VER
                Me.ModoDeApertura = VISION
        End Select

        InitializeComponent()
        dtb = New DataBase()

    End Sub

    Public Sub New()
        InitializeComponent()
        dtb = New DataBase()

    End Sub

    Sub ModificarBindingNavigator()

        If Me.Text.Substring(0, 3) = "Ver" Then
            Dim obj1 As Control
            Dim obj2 As Control
            For Each obj1 In Me.Controls
                If TypeName(obj1) = "CuadroDeTexto" Then
                    obj1.Enabled = False
                End If
                If TypeName(obj1) = "CuadroDeTexto" Then
                    obj1.Enabled = False
                End If
                If TypeName(obj1) = "ComboMAM" Then
                    obj1.Enabled = False
                End If
                If TypeName(obj1) = "TextBox" Then
                    obj1.Enabled = False
                End If
                If TypeName(obj1) = "ComboBox" Then
                    obj1.Enabled = False
                End If
                If TypeName(obj1) = "RichTextBox" Then
                    obj1.Enabled = False
                End If
                If TypeName(obj1) = "DateTimePicker" Then
                    obj1.Enabled = False
                End If
                If TypeName(obj1) = "CheckBox" Then
                    obj1.Enabled = False
                End If
                If TypeName(obj1) = "DomainUpDown" Then
                    obj1.Enabled = False
                End If
                If TypeName(obj1) = "Button" Then
                    obj1.Enabled = False
                End If

                If TypeName(obj1) = "GroupBox" Then
                    For Each obj2 In obj1.Controls
                        If TypeName(obj2) <> "Label" Then
                            obj2.Enabled = False
                        End If
                    Next
                End If
            Next

            BindingNavigatorMoveFirstItem.Visible = True
            BindingNavigatorMovePreviousItem.Visible = True
            BindingNavigatorPositionItem.Visible = True
            BindingNavigatorCountItem.Visible = True
            BindingNavigatorMoveNextItem.Visible = True
            BindingNavigatorMoveLastItem.Visible = True
            Separator2.Visible = True
            butGuardar.Visible = False
            butSalir.Visible = True
        Else
            BindingNavigatorMoveFirstItem.Visible = False
            BindingNavigatorMovePreviousItem.Visible = False
            BindingNavigatorPositionItem.Visible = False
            BindingNavigatorCountItem.Visible = False
            BindingNavigatorMoveNextItem.Visible = False
            BindingNavigatorMoveLastItem.Visible = False
            Separator2.Visible = False
            butGuardar.Visible = True
            butSalir.Visible = True
        End If
    End Sub

    'Sub ModificarBindingNavigator(ByVal StrQuery As String, ByVal Pos As Integer)
    '    If Me.Text.Substring(0, 3) = "Ver" Then
    '        GeneralBindingSource. dataSource = dtb.Consultar(StrQuery)
    '        'MAM 14/08/2011 creo que no hace falta esta linea
    '        GeneralBindingSource.Position = Pos

    '        Dim obj1 As Control
    '        Dim obj2 As Control
    '        For Each obj1 In Me.Controls
    '            If TypeName(obj1) = "CuadroDeTexto" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "CuadroDeTexto" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "ComboMAM" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "TextBox" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "ComboBox" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "RichTextBox" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "DateTimePicker" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "CheckBox" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "DomainUpDown" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "Button" Then
    '                obj1.Enabled = False
    '            End If

    '            If TypeName(obj1) = "GroupBox" Then
    '                For Each obj2 In obj1.Controls
    '                    If TypeName(obj2) <> "Label" Then
    '                        obj2.Enabled = False
    '                    End If
    '                Next
    '            End If
    '        Next

    '        BindingNavigatorMoveFirstItem.Visible = True
    '        BindingNavigatorMovePreviousItem.Visible = True
    '        BindingNavigatorSeparator.Visible = True
    '        BindingNavigatorPositionItem.Visible = True
    '        BindingNavigatorCountItem.Visible = True
    '        eparator1.Visible = True
    '        BindingNavigatorMoveNextItem.Visible = True
    '        BindingNavigatorMoveLastItem.Visible = True
    '        Separator2.Visible = True
    '        butGuardar.Visible = False
    '        butSalir.Visible = True
    '    Else
    '        BindingNavigatorMoveFirstItem.Visible = False
    '        BindingNavigatorMovePreviousItem.Visible = False
    '        BindingNavigatorSeparator.Visible = False
    '        BindingNavigatorPositionItem.Visible = False
    '        BindingNavigatorCountItem.Visible = False
    '        eparator1.Visible = False
    '        BindingNavigatorMoveNextItem.Visible = False
    '        BindingNavigatorMoveLastItem.Visible = False
    '        Separator2.Visible = False
    '        butGuardar.Visible = True
    '        butSalir.Visible = True
    '    End If
    'End Sub

    'Sub ModificarBindingNavigator(ByVal StrQuery As String)
    '    If Me.Text.Substring(0, 3) = "Ver" Then
    '        GeneralBindingSource. dataSource = dtb.Consultar(StrQuery)
    '        'MAM 14/08/2011 creo esta sobrecarga para no tener necesidad de pasar el parametro Pos

    '        Dim obj1 As Control
    '        Dim obj2 As Control
    '        For Each obj1 In Me.Controls
    '            If TypeName(obj1) = "CuadroDeTexto" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "CuadroDeTexto" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "ComboMAM" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "TextBox" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "ComboBox" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "RichTextBox" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "DateTimePicker" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "CheckBox" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "DomainUpDown" Then
    '                obj1.Enabled = False
    '            End If
    '            If TypeName(obj1) = "Button" Then
    '                obj1.Enabled = False
    '            End If

    '            If TypeName(obj1) = "GroupBox" Then
    '                For Each obj2 In obj1.Controls
    '                    If TypeName(obj2) <> "Label" Then
    '                        obj2.Enabled = False
    '                    End If
    '                Next
    '            End If
    '        Next

    '        BindingNavigatorMoveFirstItem.Visible = True
    '        BindingNavigatorMovePreviousItem.Visible = True
    '        BindingNavigatorSeparator.Visible = True
    '        BindingNavigatorPositionItem.Visible = True
    '        BindingNavigatorCountItem.Visible = True
    '        eparator1.Visible = True
    '        BindingNavigatorMoveNextItem.Visible = True
    '        BindingNavigatorMoveLastItem.Visible = True
    '        Separator2.Visible = True
    '        butGuardar.Visible = False
    '        butSalir.Visible = True
    '    Else
    '        BindingNavigatorMoveFirstItem.Visible = False
    '        BindingNavigatorMovePreviousItem.Visible = False
    '        BindingNavigatorSeparator.Visible = False
    '        BindingNavigatorPositionItem.Visible = False
    '        BindingNavigatorCountItem.Visible = False
    '        eparator1.Visible = False
    '        BindingNavigatorMoveNextItem.Visible = False
    '        BindingNavigatorMoveLastItem.Visible = False
    '        Separator2.Visible = False
    '        butGuardar.Visible = True
    '        butSalir.Visible = True
    '    End If
    'End Sub

    Sub OcultarBotonGrabar(ByVal Cadena As String)
        If Cadena = "Ver" Then
            OcultarBotonGrabar(True)
        Else
            OcultarBotonGrabar(False)
        End If

    End Sub

    Sub OcultarBotonGrabar(ByVal OcultarGrabar As Boolean)
        BindingNavigatorMoveFirstItem.Visible = False
        BindingNavigatorMovePreviousItem.Visible = False
        BindingNavigatorPositionItem.Visible = False
        BindingNavigatorCountItem.Visible = False
        BindingNavigatorMoveNextItem.Visible = False
        BindingNavigatorMoveLastItem.Visible = False
        Separator2.Visible = False
        If OcultarGrabar Then butGuardar.Visible = False
    End Sub

    Sub OcultarBotonSinRegistros()
        BindingNavigatorMoveFirstItem.Visible = False
        BindingNavigatorMovePreviousItem.Visible = False
        BindingNavigatorPositionItem.Visible = False
        BindingNavigatorCountItem.Visible = False
        BindingNavigatorMoveNextItem.Visible = False
        BindingNavigatorMoveLastItem.Visible = False
        Separator2.Visible = False
        butVer.Visible = False
    End Sub

    Sub cargarGrillaPredefinida(ByVal grillaRef As DataGridView, ByRef grilla As DataGridView)
        With grilla
            .AllowUserToAddRows = grillaRef.AllowUserToAddRows
            .AllowUserToDeleteRows = grillaRef.AllowUserToDeleteRows
            .AllowUserToOrderColumns = grillaRef.AllowUserToOrderColumns
            .AllowUserToResizeColumns = grillaRef.AllowUserToResizeColumns
            .AllowUserToResizeRows = grillaRef.AllowUserToResizeRows
            .AlternatingRowsDefaultCellStyle = grillaRef.AlternatingRowsDefaultCellStyle
            .Anchor = grillaRef.Anchor
            .AutoGenerateColumns = grillaRef.AutoGenerateColumns
            .AutoSize = grillaRef.AutoSize
            .AutoSizeColumnsMode = grillaRef.AutoSizeColumnsMode
            .AutoSizeRowsMode = grillaRef.AutoSizeRowsMode
            .BackgroundColor = grillaRef.BackgroundColor
            .BorderStyle = grillaRef.BorderStyle
            .CellBorderStyle = grillaRef.CellBorderStyle
            .ClipboardCopyMode = grillaRef.ClipboardCopyMode
            .ColumnHeadersBorderStyle = grillaRef.ColumnHeadersBorderStyle
            .ColumnHeadersDefaultCellStyle = grillaRef.ColumnHeadersDefaultCellStyle
            .ColumnHeadersHeight = grillaRef.ColumnHeadersHeight
            .ColumnHeadersHeightSizeMode = grillaRef.ColumnHeadersHeightSizeMode
            .ColumnHeadersVisible = grillaRef.ColumnHeadersVisible
            .DefaultCellStyle = grillaRef.DefaultCellStyle
            .Enabled = grillaRef.Enabled
            .Font = grillaRef.Font
            .ForeColor = grillaRef.ForeColor
            .GridColor = grillaRef.GridColor
            .MaximumSize = grillaRef.MaximumSize
            .MinimumSize = grillaRef.MinimumSize
            .MultiSelect = grillaRef.MultiSelect
            .RowHeadersBorderStyle = grillaRef.RowHeadersBorderStyle
            .RowHeadersDefaultCellStyle = grillaRef.RowHeadersDefaultCellStyle
            .RowHeadersVisible = grillaRef.RowHeadersVisible
            .RowHeadersWidth = grillaRef.RowHeadersWidth
            .RowHeadersWidthSizeMode = grillaRef.RowHeadersWidthSizeMode
            .RowsDefaultCellStyle = grillaRef.RowsDefaultCellStyle
            .ScrollBars = grillaRef.ScrollBars
        End With
    End Sub

    Private Sub butGuardar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butGuardar.Click
        ParaGuardar = True
        Dim obj1 As Control
        Dim obj2 As Control

        For Each obj1 In Me.Controls
            'If obj1.Visible = True And obj1.Enabled = True And TypeName(obj1) <> "RadioButton" Then
            '    obj1.Focus()

            '    If Not obj1.Focused Then
            '        ParaGuardar = False
            '        Return
            '    End If

            'End If

            If TypeName(obj1) = "GroupBox" Then
                For Each obj2 In obj1.Controls
                    If obj2.Visible = True And obj2.Enabled = True And TypeName(obj1) <> "RadioButton" Then
                        If TypeName(obj2) = "CuadroDeTexto" Then
                            obj2.Focus()

                            If Not obj2.Focused Then
                                ParaGuardar = False
                                Return
                            End If
                        End If
                    End If
                Next
            End If
        Next

        txtTodoCorrecto.Focus()
        If txtTodoCorrecto.Focused Then
            Guardar()
        End If
        ParaGuardar = False

    End Sub

    Private Sub txtTodoCorrecto_BackColorChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtTodoCorrecto.BackColorChanged
        If Not ParaGuardar Then
            Dim obj1 As Control
            For Each obj1 In Me.Controls
                If obj1.TabIndex = 1 Then
                    obj1.Focus()
                    Return
                End If
            Next

        End If
    End Sub

    Private Sub butSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butSalir.Click
        Salir()
    End Sub

    Private Sub butImprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butImprimir.Click
        Imprimir()
    End Sub

    Private Sub butExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butExcel.Click
        ExportToExcel()
    End Sub

    Overridable Sub Guardar()
        If Me.ModoDeApertura = INSERCION Then
            Me.Text = Me.Text.Replace("Insertar", "Modificar")
            Me.ModoDeApertura = MODIFICACION
        End If
    End Sub

    Overridable Sub Salir()
        Me.Close()
    End Sub

    Overridable Sub Imprimir()
    End Sub

    Overridable Sub ExportToExcel()
    End Sub

    Overridable Sub Ver()
    End Sub

    Overridable Sub MoveRecord(ByVal Move As String)
        Select Case Move
            Case Is = "First"
                GeneralBindingSource.Position = 0
            Case Is = "Previous"
                GeneralBindingSource.Position -= 1
            Case Is = "Next"
                GeneralBindingSource.Position += 1
            Case Is = "Last"
                GeneralBindingSource.Position = GeneralBindingSource.Count - 1
        End Select

        setValores(GeneralBindingSource(GeneralBindingSource.Position).Item(0), True)
    End Sub

    Private Sub BindingNavigatorMoveFirstItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BindingNavigatorMoveFirstItem.Click
        MoveRecord("First")
    End Sub

    Private Sub BindingNavigatorMovePreviousItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BindingNavigatorMovePreviousItem.Click
        MoveRecord("Previous")
    End Sub

    Private Sub BindingNavigatorMoveNextItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BindingNavigatorMoveNextItem.Click
        MoveRecord("Next")
    End Sub

    Private Sub BindingNavigatorMoveLastItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles BindingNavigatorMoveLastItem.Click
        MoveRecord("Last")
    End Sub

    Overridable Sub setValores(ByVal m_ArticuloID As Integer, ByVal m_SelectRecord As Boolean)
    End Sub

    Overridable Sub setValores()
    End Sub

    Private Sub butVer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butVer.Click
        Ver()
    End Sub

    Private Sub FrmAHeredarEnt_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        Me.gestionarEventosDeTeclado(e)
    End Sub

    Public Sub gestionarEventosDeTeclado(ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = "27" And Me.butSalir.Visible And Me.butSalir.Enabled Then
            Me.Close()
        ElseIf e.KeyCode = "116" And Me.butGuardar.Visible And Me.butGuardar.Enabled Then
            Me.Guardar()
        ElseIf e.KeyCode = "117" And Me.butImprimir.Visible And Me.butImprimir.Enabled Then
            Me.Imprimir()
        ElseIf e.KeyCode = "118" And Me.butVer.Visible And Me.butVer.Enabled Then
            Me.Ver()
        ElseIf e.KeyCode = "119" And Me.butExcel.Visible And Me.butExcel.Enabled Then
            Me.ExportToExcel()
        End If
    End Sub

    Private Sub FrmAHeredarEnt_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Not Me.ModoDeApertura = Nothing Then
            Me.Text = Me.Text.Replace(GridSimpleForm.ACCION_INSERTAR & " ", "")
            Me.Text = Me.Text.Replace(GridSimpleForm.ACCION_MODIFICAR & " ", "")
            Me.Text = Me.Text.Replace(GridSimpleForm.ACCION_VER & " ", "")

            Select Case Me.ModoDeApertura
                Case INSERCION
                    Me.Text = GridSimpleForm.ACCION_INSERTAR & " " & Me.Text
                    Me.Icon = My.Resources.window_new
                Case MODIFICACION
                    Me.Text = GridSimpleForm.ACCION_MODIFICAR & " " & Me.Text
                    Me.Icon = My.Resources.window_new_3
                Case VISION
                    Me.Text = GridSimpleForm.ACCION_VER & " " & Me.Text
                    butGuardar.Visible = False
                    Me.Icon = My.Resources.window_duplicate1
                Case Else
                    Me.Icon = My.Resources.LogoBandejaLa21
            End Select
        Else
            Select Case Me.Text.Substring(0, 3)
                Case GridSimpleForm.ACCION_INSERTAR.Substring(0, 3)
                    'Me.Text = GridSimpleForm.ACCION_INSERTAR & " " & Me.Text
                    Me.Icon = My.Resources.window_new
                Case GridSimpleForm.ACCION_MODIFICAR.Substring(0, 3)
                    ' Me.Text = GridSimpleForm.ACCION_MODIFICAR & " " & Me.Text
                    Me.Icon = My.Resources.window_new_3
                Case GridSimpleForm.ACCION_VER.Substring(0, 3)
                    'Me.Text = GridSimpleForm.ACCION_VER & " " & Me.Text
                    butGuardar.Visible = False
                    Me.Icon = My.Resources.window_duplicate1
                Case Else
                    Me.Icon = My.Resources.LogoBandejaLa21
            End Select
        End If

    End Sub

    Private Sub FrmAHeredarEnt_Shown(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Shown
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Me.Cursor = Cursors.Default
    End Sub
End Class