<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmEntrada
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lentrada = New System.Windows.Forms.Label()
        Me.txtEntrada = New System.Windows.Forms.TextBox()
        Me.btnAceptar = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lentrada
        '
        Me.lentrada.AutoSize = True
        Me.lentrada.Location = New System.Drawing.Point(12, 11)
        Me.lentrada.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lentrada.Name = "lentrada"
        Me.lentrada.Size = New System.Drawing.Size(51, 17)
        Me.lentrada.TabIndex = 0
        Me.lentrada.Text = "Label1"
        '
        'txtEntrada
        '
        Me.txtEntrada.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEntrada.Location = New System.Drawing.Point(12, 31)
        Me.txtEntrada.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.txtEntrada.Name = "txtEntrada"
        Me.txtEntrada.Size = New System.Drawing.Size(127, 22)
        Me.txtEntrada.TabIndex = 1
        '
        'btnAceptar
        '
        Me.btnAceptar.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnAceptar.Image = Global.BasesParaCompatibilidad.My.Resources.Resources.dialog_apply
        Me.btnAceptar.Location = New System.Drawing.Point(148, 31)
        Me.btnAceptar.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.btnAceptar.Name = "btnAceptar"
        Me.btnAceptar.Size = New System.Drawing.Size(39, 27)
        Me.btnAceptar.TabIndex = 2
        Me.btnAceptar.UseVisualStyleBackColor = True
        '
        'frmEntrada
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 16.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(199, 80)
        Me.Controls.Add(Me.btnAceptar)
        Me.Controls.Add(Me.txtEntrada)
        Me.Controls.Add(Me.lentrada)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmEntrada"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "frmEntrada"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents lentrada As System.Windows.Forms.Label
    Private WithEvents txtEntrada As System.Windows.Forms.TextBox
    Private WithEvents btnAceptar As System.Windows.Forms.Button
End Class
