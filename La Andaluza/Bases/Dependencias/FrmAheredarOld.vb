
'@DEPRECATED
Public Class FrmAheredarOld

    Protected sp As sp
    Protected spSelectDgv As String
    Protected m_MaestroID As Integer
    Public Embebido As Boolean
    Public Const ACCION_INSERTAR As String = "Insertar"
    Public Const ACCION_MODIFICAR As String = "Modificar"
    Public Const ACCION_VER As String = "Ver"
    Protected mse As MicrosoftOfficeExporter
    Protected filterManager As DgvFilterPopup.DgvFilterManager
    Protected dataSource As DataTable
    Protected dtb As DataBase
    Protected LastEntry As Boolean

    Public Sub New()

        ' Llamada necesaria para el diseñador.
        InitializeComponent()
        dtb = New DataBase()
        Me.sp = Nothing
        Me.m_MaestroID = 0
        mse = New MicrosoftOfficeExporter
        filterManager = New DgvFilterPopup.DgvFilterManager(dgvGeneral)
        LastEntry = True

    End Sub
    Public Sub New(ByRef storedProc As sp, Optional ByVal MaestroID As Integer = 0)
        ' This call is required by the designer.
        InitializeComponent()
        dtb = New DataBase()

        If Not storedProc Is Nothing Then Me.sp = storedProc
        Me.m_MaestroID = MaestroID
        mse = New MicrosoftOfficeExporter
        filterManager = New DgvFilterPopup.DgvFilterManager(dgvGeneral)
        LastEntry = True
    End Sub



    Public Sub BindingNavigatorActualizar()
        If dgvGeneral.RowCount <= 0 Then
            BindingNavigatorSinRegistros()
        Else
            BindingNavigatorConRegistros()
        End If
    End Sub

    Sub BindingNavigatorActualizar(ByVal RowCount As Integer)

        If RowCount <= 0 Then
            BindingNavigatorSinRegistros()
        Else
            BindingNavigatorConRegistros()
        End If
    End Sub

    Sub BindingNavigatorCortoActualizar(ByVal RowCount As Integer)
        BindingNavigatorMoveFirstItem.Visible = False
        BindingNavigatorMovePreviousItem.Visible = False
        BindingNavigatorPositionItem.Visible = False
        BindingNavigatorCountItem.Visible = False
        BindingNavigatorMoveNextItem.Visible = False
        BindingNavigatorMoveLastItem.Visible = False
        Separator1.Visible = False
        ToolStripSeparator2.Visible = False
        butVer.Visible = False
        butSalir.Visible = False
        butExcel.Visible = False
        butWord.Visible = False

        If RowCount <= 0 Then
            butModificar.Visible = False
            butEliminar.Visible = False
        Else
            butModificar.Visible = True
            butEliminar.Visible = True
        End If
    End Sub

    Sub BindingNavigatorSinRegistros()
        BindingNavigatorMoveFirstItem.Visible = False
        BindingNavigatorMovePreviousItem.Visible = False
        BindingNavigatorPositionItem.Visible = False
        'BindingNavigatorCountItem.Visible = False
        BindingNavigatorMoveNextItem.Visible = False
        BindingNavigatorMoveLastItem.Visible = False
        Separator1.Visible = False
        ToolStripSeparator2.Visible = False
        butVer.Visible = False
        butModificar.Visible = False
        butEliminar.Visible = False
        butExcel.Visible = False
        butWord.Visible = False
    End Sub

    Sub BindingNavigatorConRegistros()
        BindingNavigatorMoveFirstItem.Visible = True
        BindingNavigatorMovePreviousItem.Visible = True
        BindingNavigatorPositionItem.Visible = True
        BindingNavigatorCountItem.Visible = True
        Separator1.Visible = True
        ToolStripSeparator2.Visible = True
        BindingNavigatorMoveNextItem.Visible = True
        BindingNavigatorMoveLastItem.Visible = True
        butVer.Visible = True
        butModificar.Visible = True
        butEliminar.Visible = True
    End Sub

    Overridable Sub Insertar()
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Me.Cursor = Cursors.WaitCursor
        Try
            If Me.butNuevo.Enabled Then
                Action(ACCION_INSERTAR)
                GeneralBindingSource.MoveLast()
            End If
        Catch ex As Exception
        Finally
            Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Overridable Sub Modificar()
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Me.Cursor = Cursors.WaitCursor
        Try
            If Me.butModificar.Enabled Then Action(ACCION_MODIFICAR)
        Catch ex As Exception
        Finally
            Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Overridable Sub Ver()
        Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Me.Cursor = Cursors.WaitCursor
        Try
            If Me.butVer.Enabled Then Action(ACCION_VER)
        Catch ex As Exception
        Finally
            Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
            Me.Cursor = Cursors.Default
        End Try
    End Sub

    Overridable Sub Eliminar()
    End Sub

    Overridable Sub Imprimir()
    End Sub

    Overridable Sub Guardar()
    End Sub

    Overridable Sub Action(ByVal TipoAction As String)
    End Sub

    Overridable Sub ToExcel()
        Me.ExportarAExcel(MicrosoftOfficeExporter.MANY, Me.Text, Me.dgvGeneral)
    End Sub

    Overridable Sub ToWord()
        Me.ExportarAWord(MicrosoftOfficeExporter.MANY, Me.Text, Me.dgvGeneral)
    End Sub

    Overridable Sub Refrescar()
        dgvFill()
    End Sub

    Private Sub butNuevo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butNuevo.Click
        Insertar()
    End Sub

    Private Sub butModificar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butModificar.Click
        Modificar()
    End Sub

    Private Sub butVer_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butVer.Click
        Ver()
    End Sub

    Private Sub butEliminar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butEliminar.Click
        Eliminar()
    End Sub

    Private Sub butImprimir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butImprimir.Click
        Imprimir()
    End Sub
  
    Private Sub butSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butSalir.Click
        Me.Close()
    End Sub

    Private Sub butGuardar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butGuardar.Click
        Guardar()
    End Sub

    Private Sub butExcel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butExcel.Click
        ToExcel()
    End Sub

    Private Sub butWord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butWord.Click
        ToWord()
    End Sub

    Private Sub FrmAheredar_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyDown
        If Me.Embebido <> True Or Me.Embebido = Nothing Then Me.gestionarEventosDeTeclado(e)
    End Sub

    Public Sub gestionarEventosDeTeclado(ByVal e As System.Windows.Forms.KeyEventArgs)
        Try
            If e.KeyCode = "27" And Me.Embebido = False And Me.butSalir.Visible Then
                Me.Close()
            ElseIf e.KeyCode = "116" And Me.Embebido = False And Me.butGuardar.Visible And Me.butGuardar.Enabled Then
                Me.Guardar()
            ElseIf e.KeyCode = "112" And Me.butNuevo.Enabled Then
                Me.Insertar()
            ElseIf (e.KeyCode = "113" Or e.KeyCode = "13") And Me.butModificar.Enabled Then
                Modificar()
            ElseIf e.KeyCode = "114" And Me.butEliminar.Enabled Then
                Me.Eliminar()
            ElseIf e.KeyCode = "120" And Me.butImprimir.Enabled Then
                Me.Imprimir()
            ElseIf e.KeyCode = "121" And Me.butWord.Enabled Then
                Me.ToWord()
            ElseIf e.KeyCode = "122" And Me.butExcel.Enabled Then
                Me.ToExcel()
            ElseIf e.KeyCode = "123" And Me.butVer.Enabled Then
                Me.Ver()
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FrmAheredar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Me.Embebido Then
            Me.butSalir.Visible = False
            Me.butSalir.Enabled = False
        End If

        dgvFill()
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim cabecera(1) As String
        cabecera(1) = Me.Text
        Dim p As New Printer
        p.imprimirGrilla(Me.dgvGeneral, cabecera, Nothing, False, False)
    End Sub

    Private Sub ImprimirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImprimirToolStripMenuItem.Click
        Me.PrintDocument1.Print()
    End Sub

    'Private Sub butRefrescar_Click(sender As System.Object, e As System.EventArgs) Handles butRefrescar.Click
    '    dgvFill()
    'End Sub

    Private Sub butRefrescar_Click(sender As System.Object, e As System.EventArgs) Handles butRefrescar.Click
        Refrescar()
    End Sub

    Protected Sub ExportarAExcel(ByVal MANY As String, ByVal p2 As String, ByRef dataGridView As DataGridView)
        mse.ExportToExcel(MicrosoftOfficeExporter.MANY, Me.Text, dataGridView)
    End Sub

    Protected Sub ExportarAWord(ByVal MANY As String, ByVal p2 As String, ByRef dataGridView As DataGridView)
        mse.ExportToWord(MicrosoftOfficeExporter.MANY, Me.Text, dataGridView)
    End Sub

    Protected Overridable Sub dgvFill()
        If BackgroundWorker1.IsBusy Then
            Return
            'BackgroundWorker1.CancelAsync()
        End If

        dgvGeneral.Cursor = Cursors.WaitCursor
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Me.cargar_datos()
    End Sub

    Protected Overridable Sub BackgroundWorker1_RunWorkerCompleted(sender As System.Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        Me.dgvGeneral.SuspendLayout()
        Me.BindDataSource()
        BindingNavigatorActualizar()
        dgvGeneral.FormatoGeneral()
        Me.dgvGeneral.ResumeLayout()
        If Me.LastEntry Then
            dgvGeneral.MoveLast()
        Else
            dgvGeneral.MoveFirst()
        End If
        Me.dgvGeneral.Cursor = Cursors.Default
    End Sub

    Protected Overridable Sub cargar_datos()
        ' dataSource = dtb.Consultar(Me.sp.DataGridViewStoredProcedure, True)
    End Sub

    Protected Overridable Sub BindDataSource()
        If Not dataSource Is Nothing Then
            GeneralBindingSource.DataSource = dataSource
            With dgvGeneral
                .DataSource = GeneralBindingSource

            End With
        End If
    End Sub

    Private Sub dgvGeneral_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvGeneral.CellDoubleClick
        Modificar()
    End Sub

    Private Sub FrmAheredarOld_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Me.BackgroundWorker1.IsBusy Then
            Me.BackgroundWorker1.CancelAsync()
        End If
        Me.dtb.Desconectar()
    End Sub
End Class