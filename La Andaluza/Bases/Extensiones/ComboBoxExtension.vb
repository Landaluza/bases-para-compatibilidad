Public Module ComboBoxExtension
    Private combotooltip As New ToolTip

    <Extension()> _
    Public Sub mam_DataSource(ByVal Combo As ComboBox, ByVal tb As DataTable, _
                              ByVal PrimeraVacia As Boolean, Optional ByVal Texto As String = "")

        Try
            If Not tb Is Nothing Then

                If tb.Rows.Count = 0 Then
                    Combo.DataSource = Nothing
                Else
                    If PrimeraVacia Then
                        Dim row As DataRow
                        row = tb.NewRow()
                        row.Item(0) = 0

                        row.Item(1) = Texto

                        tb.Rows.InsertAt(row, 0)
                        row = Nothing
                    End If

                    Combo.DataSource = tb
                    Combo.ValueMember = tb.Columns(0).ColumnName
                    Combo.DisplayMember = tb.Columns(1).ColumnName
                    AddHandler Combo.SelectedIndexChanged, AddressOf CambiosDeValor
                End If
            End If
            ''ComprobarCambiosCombo = true

        Catch ex As Exception
        End Try
    End Sub

    <Extension()> _
    Public Sub mam_DataSource(ByVal Combo As ComboBox, ByVal sp As String, ByVal PrimeraVacia As Boolean, Optional ByVal Texto As String = "")

        Try
            Dim tb As New DataTable
            Dim dtb As New DataBase(Config.Server)
            dtb.PrepararConsulta(sp)
            tb = dtb.Consultar
            'tb = dtb.Consultar(sp, True)

            If Not tb Is Nothing Then
                If tb.Rows.Count = 0 Then
                    Combo.DataSource = Nothing
                Else
                    If PrimeraVacia Then
                        Dim row As DataRow
                        row = tb.NewRow()
                        row.Item(0) = 0

                        row.Item(1) = Texto
                        tb.Rows.InsertAt(row, 0)
                        row = Nothing
                    End If


                    Combo.DataSource = tb
                    Combo.ValueMember = tb.Columns(0).ColumnName
                    Combo.DisplayMember = tb.Columns(1).ColumnName
                    AddHandler Combo.SelectedIndexChanged, AddressOf CambiosDeValor
                End If
            End If

        Catch ex As Exception
        End Try
    End Sub

    Private Sub CambiosDeValor(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim combo As System.Windows.Forms.ComboBox = CType(sender, System.Windows.Forms.ComboBox)
        combotooltip.ShowAlways = True
        combotooltip.ToolTipIcon = ToolTipIcon.Info
        combotooltip.UseFading = True
        combotooltip.UseAnimation = True
        combotooltip.SetToolTip(combo, combo.Text)
    End Sub


    <Extension()> _
    Public Sub mam_DataSource(ByVal Combo As ComboBox, ByVal sp As String, ByVal PrimeraVacia As Boolean, ByRef dtb As DataBase, Optional ByVal Texto As String = "")

        Try
            Dim tb As New DataTable
            dtb.PrepararConsulta(sp)
            tb = dtb.Consultar()
            'tb = dtb.Consultar(sp, True)

            If Not tb Is Nothing Then
                If tb.Rows.Count = 0 Then
                    Combo.DataSource = Nothing
                Else
                    If PrimeraVacia Then
                        Dim row As DataRow
                        row = tb.NewRow()
                        row.Item(0) = 0

                        row.Item(1) = Texto
                        tb.Rows.InsertAt(row, 0)
                        row = Nothing
                    End If

                    Combo.DataSource = tb
                    Combo.ValueMember = tb.Columns(0).ColumnName
                    Combo.DisplayMember = tb.Columns(1).ColumnName
                    AddHandler Combo.SelectedIndexChanged, AddressOf CambiosDeValor
                End If
            End If

        Catch ex As Exception
        End Try
    End Sub
End Module
