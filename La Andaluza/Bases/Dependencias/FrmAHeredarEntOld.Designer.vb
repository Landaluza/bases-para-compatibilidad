<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmAHeredarEntOld
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmAHeredarEntOld))
        Me.bdnGeneral = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.GeneralBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel()
        Me.butGuardar = New System.Windows.Forms.ToolStripButton()
        Me.butSalir = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox()
        Me.Separator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.butImprimir = New System.Windows.Forms.ToolStripButton()
        Me.butVer = New System.Windows.Forms.ToolStripButton()
        Me.butExcel = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton()
        Me.txtTodoCorrecto = New BasesParaCompatibilidad.CuadroDeTexto()
        CType(Me.bdnGeneral, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.bdnGeneral.SuspendLayout()
        CType(Me.GeneralBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'bdnGeneral
        '
        Me.bdnGeneral.AddNewItem = Nothing
        Me.bdnGeneral.BindingSource = Me.GeneralBindingSource
        Me.bdnGeneral.CountItem = Me.BindingNavigatorCountItem
        Me.bdnGeneral.DeleteItem = Nothing
        Me.bdnGeneral.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.bdnGeneral.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.butGuardar, Me.butSalir, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.Separator2, Me.butImprimir, Me.butVer, Me.butExcel, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorMoveFirstItem})
        Me.bdnGeneral.Location = New System.Drawing.Point(0, 0)
        Me.bdnGeneral.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.bdnGeneral.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.bdnGeneral.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.bdnGeneral.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.bdnGeneral.Name = "bdnGeneral"
        Me.bdnGeneral.PositionItem = Me.BindingNavigatorPositionItem
        Me.bdnGeneral.Size = New System.Drawing.Size(466, 25)
        Me.bdnGeneral.TabIndex = 7
        Me.bdnGeneral.Text = "BindingNavigatorGeneral"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(37, 22)
        Me.BindingNavigatorCountItem.Text = "de {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Numero Total de Items"
        '
        'butGuardar
        '
        Me.butGuardar.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.butGuardar.Image = Global.BasesParaCompatibilidad.My.Resources.Resources.document_save_5_16
        Me.butGuardar.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butGuardar.Name = "butGuardar"
        Me.butGuardar.Size = New System.Drawing.Size(23, 22)
        Me.butGuardar.Text = "Guardar"
        Me.butGuardar.ToolTipText = "Guardar (F5)"
        '
        'butSalir
        '
        Me.butSalir.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.butSalir.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.butSalir.Image = Global.BasesParaCompatibilidad.My.Resources.Resources.application_xit_3
        Me.butSalir.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butSalir.Margin = New System.Windows.Forms.Padding(10, 1, 0, 2)
        Me.butSalir.Name = "butSalir"
        Me.butSalir.Size = New System.Drawing.Size(23, 22)
        Me.butSalir.Text = "Salir"
        Me.butSalir.ToolTipText = "Salir (Esc)"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveLastItem.Text = "Ir al Ultimo"
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveNextItem.Text = "Ir al Siguiente"
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 23)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Posicion Actual"
        '
        'Separator2
        '
        Me.Separator2.Name = "Separator2"
        Me.Separator2.Size = New System.Drawing.Size(6, 25)
        '
        'butImprimir
        '
        Me.butImprimir.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.butImprimir.Image = Global.BasesParaCompatibilidad.My.Resources.Resources.document_print_16
        Me.butImprimir.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butImprimir.Name = "butImprimir"
        Me.butImprimir.Size = New System.Drawing.Size(23, 22)
        Me.butImprimir.Text = "&Imprimir"
        Me.butImprimir.ToolTipText = "Imprimir (F6)"
        Me.butImprimir.Visible = False
        '
        'butVer
        '
        Me.butVer.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.butVer.Image = Global.BasesParaCompatibilidad.My.Resources.Resources.view_16
        Me.butVer.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butVer.Name = "butVer"
        Me.butVer.Size = New System.Drawing.Size(23, 22)
        Me.butVer.Text = "Ver"
        Me.butVer.ToolTipText = "Ver (F7)"
        Me.butVer.Visible = False
        '
        'butExcel
        '
        Me.butExcel.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.butExcel.Image = Global.BasesParaCompatibilidad.My.Resources.Resources.page_excel_16
        Me.butExcel.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butExcel.Name = "butExcel"
        Me.butExcel.Size = New System.Drawing.Size(23, 22)
        Me.butExcel.Text = "ToolStripButton1"
        Me.butExcel.ToolTipText = "Exportar a Excel (F9)"
        Me.butExcel.Visible = False
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Ir al Anterior"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Ir al Primero"
        '
        'txtTodoCorrecto
        '
        Me.txtTodoCorrecto.EsUnicoCampo = ""
        Me.txtTodoCorrecto.EsUnicoCampoID = ""
        Me.txtTodoCorrecto.EsUnicoID = 0
        Me.txtTodoCorrecto.EsUnicoTabla = ""
        Me.txtTodoCorrecto.Location = New System.Drawing.Point(-10, 26)
        Me.txtTodoCorrecto.Minimo = 0
        Me.txtTodoCorrecto.Modificado = False
        Me.txtTodoCorrecto.Name = "txtTodoCorrecto"
        Me.txtTodoCorrecto.Numerico_EsNumerico = False
        Me.txtTodoCorrecto.Numerico_NumeroDoublees = 0
        Me.txtTodoCorrecto.Numerico_SeparadorMiles = False
        Me.txtTodoCorrecto.Obligatorio = False
        Me.txtTodoCorrecto.ParametroID = 0
        Me.txtTodoCorrecto.Size = New System.Drawing.Size(10, 20)
        Me.txtTodoCorrecto.TabIndex = 1000000000
        Me.txtTodoCorrecto.TabStop = False
        Me.txtTodoCorrecto.ValorMaximo = 0.0!
        Me.txtTodoCorrecto.ValorMinimo = 0.0!
        '
        'FrmAHeredarEntOld
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(466, 337)
        Me.ControlBox = False
        Me.Controls.Add(Me.bdnGeneral)
        Me.Controls.Add(Me.txtTodoCorrecto)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.KeyPreview = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmAHeredarEntOld"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "FrmAHeredarEnt"
        CType(Me.bdnGeneral, System.ComponentModel.ISupportInitialize).EndInit()
        Me.bdnGeneral.ResumeLayout(False)
        Me.bdnGeneral.PerformLayout()
        CType(Me.GeneralBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Public WithEvents GeneralBindingSource As System.Windows.Forms.BindingSource
    Public WithEvents bdnGeneral As System.Windows.Forms.BindingNavigator
    Public WithEvents BindingNavigatorCountItem As System.Windows.Forms.ToolStripLabel
    Public WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Public WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Public WithEvents BindingNavigatorPositionItem As System.Windows.Forms.ToolStripTextBox
    Public WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Public WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Public WithEvents Separator2 As System.Windows.Forms.ToolStripSeparator
    Public WithEvents butGuardar As System.Windows.Forms.ToolStripButton
    Public WithEvents butSalir As System.Windows.Forms.ToolStripButton
    Public WithEvents txtTodoCorrecto As CuadroDeTexto
    Public WithEvents butImprimir As System.Windows.Forms.ToolStripButton
    Public WithEvents butExcel As System.Windows.Forms.ToolStripButton
    Public WithEvents butVer As System.Windows.Forms.ToolStripButton
End Class
