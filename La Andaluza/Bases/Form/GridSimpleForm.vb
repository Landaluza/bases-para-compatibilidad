Imports DgvFilterPopup

Public Class gridsimpleform
    Protected m_MaestroID As String

    Public Embebido As Boolean
    Public Const ACCION_INSERTAR As String = "Insertar"
    Public Const ACCION_MODIFICAR As String = "Modificar"
    Public Const ACCION_VER As String = "Ver"

    Protected sp As StoredProcedure
    Protected dbo As DataBussines
    Protected newRegForm As DetailedSimpleForm
    Protected filterManager As DgvFilterPopup.DgvFilterManager
    Protected mse As MicrosoftOfficeExporter
    Protected dataSource As DataTable
    Protected dtb As DataBase

    Public Event AfterDelete(ByRef sender As Object, ByRef target As gridsimpleform)
    Public Event BeforeDelete(ByRef sender As Object, ByRef target As gridsimpleform)
    Public Event AfterInsert(ByRef sender As Object, ByRef target As DetailedSimpleForm)
    Public Event BeforeInsert(ByRef sender As Object, ByRef target As DetailedSimpleForm)
    Public Event AfterModify(ByRef sender As Object, ByRef target As DetailedSimpleForm)
    Public Event BeforeModify(ByRef sender As Object, ByRef target As DetailedSimpleForm)
    Public Event AfterView(ByRef sender As Object, ByRef target As DetailedSimpleForm)
    Public Event BeforeView(ByRef sender As Object, ByRef target As DetailedSimpleForm)
    Private handled As Boolean

    Protected Property EventHandeld As Boolean
        Get
            Return handled
        End Get
        Set(ByVal value As Boolean)
            handled = value
        End Set
    End Property

    Public Sub setDbo(ByRef databussinesobject As DataBussines)
        Me.dbo = databussinesobject
    End Sub

    'Public Sub New()

    '    InitializeComponent()

    '    filterManager = New DgvFilterManager(dgvGeneral)
    '    mse = New MicrosoftOfficeExporter
    '    dtb = New DataBase(Config.Server)
    'End Sub

    Public Sub New()

        InitializeComponent()

    End Sub
    Public Sub New(ByRef sp As StoredProcedure, ByVal maestroid As String)
        InitializeComponent()
        Me.sp = sp

        'sp.DataGridViewStoredProcedure = If(m_MaestroID = Nothing, sp.DataGridViewStoredProcedureForSelect, sp.DataGridViewStoredProcedureForFilteredSelect & "'" & Me.m_MaestroID & "'")
        If maestroid = "0" Or maestroid = String.Empty Then
            sp.DataGridViewStoredProcedure = sp.DataGridViewStoredProcedureForSelect
            m_MaestroID = ""
        Else
            m_MaestroID = maestroid
            sp.DataGridViewStoredProcedure = sp.DataGridViewStoredProcedureForFilteredSelect & "'" & maestroid & "'"
        End If

        filterManager = New DgvFilterManager(dgvGeneral)
        mse = New MicrosoftOfficeExporter
        dtb = New DataBase(Config.Server)
        dgvFill()
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
        eparator2.Visible = False
        Separator3.Visible = False
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
        BindingNavigatorCountItem.Visible = False
        BindingNavigatorMoveNextItem.Visible = False
        BindingNavigatorMoveLastItem.Visible = False
        eparator2.Visible = False
        Separator3.Visible = False
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
        Separator3.Visible = True
        BindingNavigatorMoveNextItem.Visible = True
        BindingNavigatorMoveLastItem.Visible = True
        eparator2.Visible = True
        butVer.Visible = True
        butModificar.Visible = True
        butEliminar.Visible = True
    End Sub

    Overridable Sub Insertar()
        'Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        Me.Cursor = Cursors.WaitCursor
        Me.handled = False
        RaiseEvent BeforeInsert(Me, Me.newRegForm)
        If Not Me.handled Then
            Try
                If Me.butNuevo.Enabled Then
                    Action(ACCION_INSERTAR)
                    GeneralBindingSource.MoveLast()
                End If
            Catch ex As Exception
                messagebox.show("Error al insertar el registro. Detalles:" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If
        RaiseEvent AfterInsert(Me, Me.newRegForm)
        'Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Me.Cursor = Cursors.Default
    End Sub

    Overridable Sub Modificar()
        Me.Cursor = Cursors.WaitCursor
        Me.handled = False
        RaiseEvent BeforeModify(Me, Me.newRegForm)
        If Not Me.handled Then
            Try
                If Me.butModificar.Enabled Then Action(ACCION_MODIFICAR)
            Catch ex As Exception
                messagebox.show("Error al modificar el registro. Detalles:" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            RaiseEvent AfterModify(Me, Me.newRegForm)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Overridable Sub Ver()
        Me.Cursor = Cursors.WaitCursor
        Me.handled = False
        RaiseEvent BeforeView(Me, Me.newRegForm)
        If Not Me.handled Then
            Try
                If Me.butVer.Enabled Then Action(ACCION_VER)
            Catch ex As Exception
                messagebox.show("Error al ver el registro. Detalles:" & Environment.NewLine & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            RaiseEvent AfterView(Me, Me.newRegForm)
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Overridable Sub Eliminar(Optional ByVal campoId As String = "Id")
        If MessageBox.Show(" ¿Realmente quieres eliminar este/estos registro ? ", " Eliminar ", MessageBoxButtons.YesNo, MessageBoxIcon.Information) = DialogResult.Yes Then

            Me.handled = False
            RaiseEvent BeforeDelete(CType(dgvGeneral, Control), Me)
            If Not Me.handled Then
                If dgvGeneral.Rows.Count > 0 Then
                    For Each row As DataGridViewRow In dgvGeneral.SelectedRows
                        If Not row.DataGridView Is Nothing Then
                            If sp.Delete(CType(row.Cells(campoId).Value, Integer)) Then
                                RaiseEvent AfterDelete(CType(dgvGeneral, Control), Me)
                            Else
                                MessageBox.Show("No se pudo eliminar el registro", "Atencion", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                Exit For
                            End If
                        End If
                    Next
                End If
            End If
            Me.dgvGeneral.Cursor = Cursors.Default
        End If
    End Sub
    Overridable Sub Imprimir()
    End Sub

    Overridable Sub Guardar()
    End Sub

    Public Overridable Sub Action(ByVal TipoAction As String)
        Me.newRegForm.ModoDeAperturaCadena = TipoAction
        'GUImain.añadirPestaña(Me.newRegForm)
        'Me.newRegForm.ShowDialog()
        Pantalla.mostrarDialogo(Me.newRegForm)
    End Sub

    Overridable Sub ToExcel()
        Me.ExportarAExcel(MicrosoftOfficeExporter.MANY, Me.Text, Me.dgvGeneral)
    End Sub

    Overridable Sub ToWord()
        Me.ExportarAWord(MicrosoftOfficeExporter.MANY, Me.Text, Me.dgvGeneral)
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
            If e.KeyCode = "27" And Not Me.Embebido And Me.butSalir.Visible Then
                Me.Close()
            ElseIf e.KeyCode = "116" And Not Me.Embebido And Me.butGuardar.Visible And Me.butGuardar.Enabled Then
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
            ElseIf e.KeyCode = "115" Then 'F4
                MessageBox.Show(Me.Name)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub FrmAheredar_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Me.Embebido Or Me.m_MaestroID <> "" Then
            Me.butSalir.Visible = False
            Me.butSalir.Enabled = False
        Else
            Me.butSalir.Visible = True
            Me.butSalir.Enabled = True
        End If

        If Not newRegForm Is Nothing Then AddHandler newRegForm.afterSave, AddressOf dgvFill
    End Sub

    Protected Sub dgvGeneral_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvGeneral.CellDoubleClick
        Modificar()
    End Sub

    Protected Sub Deleted(ByRef sender As Object, ByRef target As gridsimpleform) Handles Me.AfterDelete
        dgvFill()
    End Sub

    Private Sub ImprimirToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImprimirToolStripMenuItem.Click
        PrintDocument1.Print()
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim cabecera(1) As String
        cabecera(1) = Me.Text
        Dim p As New Printer
        p.imprimirGrilla(Me.dgvGeneral, cabecera, Nothing, False, False)
    End Sub

    Private Sub butRefrescar_Click(sender As System.Object, e As System.EventArgs) Handles butRefrescar.Click
        dgvFill()
    End Sub

    Protected Sub ExportarAExcel(ByVal MANY As String, ByVal p2 As String, ByRef dataGridView As DataGridView)
        mse.ExportToExcel(MicrosoftOfficeExporter.MANY, Me.Text, dataGridView)
    End Sub

    Protected Sub ExportarAWord(ByVal MANY As String, ByVal p2 As String, ByRef dataGridView As DataGridView)
        mse.ExportToWord(MicrosoftOfficeExporter.MANY, Me.Text, dataGridView)
    End Sub

    Protected Overridable Sub cargar_datos()
        dtb.PrepararConsulta(Me.sp.DataGridViewStoredProcedure)
        dataSource = dtb.Consultar
        'dataSource = dtb.Consultar(Me.sp.DataGridViewStoredProcedure, True)
    End Sub

    Protected Overridable Sub BindDataSource()
        If Not dataSource Is Nothing Then
            GeneralBindingSource.DataSource = dataSource
            With dgvGeneral
                .DataSource = GeneralBindingSource

            End With
        End If
    End Sub

    Protected Sub dgvFill()

        If BackgroundWorker1.IsBusy Then
            Return
        End If

        dgvGeneral.Cursor = Cursors.WaitCursor
        BackgroundWorker1.RunWorkerAsync()
    End Sub

    Private Sub BackgroundWorker1_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles BackgroundWorker1.DoWork
        Me.cargar_datos()
    End Sub

    Private Sub BackgroundWorker1_RunWorkerCompleted(sender As System.Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BackgroundWorker1.RunWorkerCompleted
        bdnGeneral.SuspendLayout()
        'CType(Me.GeneralBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.dgvGeneral.SuspendLayout()

        Try
            Me.BindDataSource()
        Catch ex As Exception

        End Try

        BindingNavigatorActualizar()
        dgvGeneral.FormatoGeneral()

        bdnGeneral.ResumeLayout()
        'CType(Me.GeneralBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.dgvGeneral.ResumeLayout()
        'dgvGeneral.PerformLayout()

        dgvGeneral.MoveLast()
        Me.dgvGeneral.Cursor = Cursors.Default
    End Sub

    Private Sub gridsimpleform_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If Me.BackgroundWorker1.IsBusy Then
            Me.BackgroundWorker1.CancelAsync()
        End If

        Me.dtb.Desconectar()
    End Sub
End Class