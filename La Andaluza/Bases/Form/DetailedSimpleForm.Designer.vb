<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DetailedSimpleForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DetailedSimpleForm))
        Me.bdnGeneral = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.GeneralBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.butGuardar = New System.Windows.Forms.ToolStripButton()
        Me.butImprimir = New System.Windows.Forms.ToolStripButton()
        Me.butVer = New System.Windows.Forms.ToolStripButton()
        Me.butExcel = New System.Windows.Forms.ToolStripButton()
        Me.butSalir = New System.Windows.Forms.ToolStripButton()
        Me.MyToolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.Separator1 = New System.Windows.Forms.ToolStripSeparator()
        CType(Me.bdnGeneral, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.bdnGeneral.SuspendLayout()
        CType(Me.GeneralBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'bdnGeneral
        '
        Me.bdnGeneral.AddNewItem = Nothing
        Me.bdnGeneral.BindingSource = Me.GeneralBindingSource
        Me.bdnGeneral.CountItem = Nothing
        Me.bdnGeneral.DeleteItem = Nothing
        Me.bdnGeneral.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden
        Me.bdnGeneral.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.butGuardar, Me.Separator1, Me.butImprimir, Me.butVer, Me.butExcel, Me.butSalir})
        Me.bdnGeneral.Location = New System.Drawing.Point(0, 0)
        Me.bdnGeneral.MoveFirstItem = Nothing
        Me.bdnGeneral.MoveLastItem = Nothing
        Me.bdnGeneral.MoveNextItem = Nothing
        Me.bdnGeneral.MovePreviousItem = Nothing
        Me.bdnGeneral.Name = "bdnGeneral"
        Me.bdnGeneral.PositionItem = Nothing
        Me.bdnGeneral.Size = New System.Drawing.Size(520, 25)
        Me.bdnGeneral.TabIndex = 7
        Me.bdnGeneral.Text = "BindingNavigatorGeneral"
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
        'butSalir
        '
        Me.butSalir.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
        Me.butSalir.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.butSalir.Image = Global.BasesParaCompatibilidad.My.Resources.Resources.application_xit_3
        Me.butSalir.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.butSalir.Name = "butSalir"
        Me.butSalir.Size = New System.Drawing.Size(23, 22)
        Me.butSalir.Text = "Salir"
        Me.butSalir.ToolTipText = "Salir (Esc)"
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        Me.Separator1.Size = New System.Drawing.Size(6, 25)
        '
        'DetailedSimpleForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(520, 337)
        Me.ControlBox = False
        Me.Controls.Add(Me.bdnGeneral)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.Name = "DetailedSimpleForm"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrmAHeredarEnt"
        CType(Me.bdnGeneral, System.ComponentModel.ISupportInitialize).EndInit()
        Me.bdnGeneral.ResumeLayout(False)
        Me.bdnGeneral.PerformLayout()
        CType(Me.GeneralBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Protected WithEvents butGuardar As System.Windows.Forms.ToolStripButton
    Protected WithEvents GeneralBindingSource As System.Windows.Forms.BindingSource
    Protected WithEvents bdnGeneral As System.Windows.Forms.BindingNavigator
    Protected WithEvents butSalir As System.Windows.Forms.ToolStripButton
    Protected WithEvents butImprimir As System.Windows.Forms.ToolStripButton
    Protected WithEvents butExcel As System.Windows.Forms.ToolStripButton
    Protected WithEvents butVer As System.Windows.Forms.ToolStripButton
    Protected WithEvents MyToolTip As System.Windows.Forms.ToolTip
    Protected WithEvents Separator1 As System.Windows.Forms.ToolStripSeparator
End Class
