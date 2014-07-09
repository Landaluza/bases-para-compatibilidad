




Public Enum TiposColumna
    Miles
    Double2
    FechaCorta
    Hora
    Derecha
    Izquierda
    Referencia
    Lote
    Observaciones
    QS
    Juliano
    Centro
    Proveedor
    Double5
    Descripcion
    Persona
End Enum

Public Module DataGridViewExtension

    Private Const Fecha As String = "dd/MM/yyyy"
    Private Const Hora As String = "hh\:mm"
    Private Const Miles As String = "###,###,###"
    Private Const Double2 As String = "0.#0"
    Private Const Double5 As String = "0.#0000"

    <Extension()> _
    Public Sub FormatoGeneral(ByVal dgv As DataGridView, Optional ByVal estrecho As Boolean = False)
        'dgv.SuspendLayout()
        Try
            With dgv
                .AllowUserToAddRows = False
                .AllowUserToDeleteRows = False
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                .AllowUserToResizeColumns = True
                .AutoSize = False
                .ReadOnly = True
                .ScrollBars = ScrollBars.Both

                .ColumnHeadersDefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Bold)
                .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

                .RowHeadersWidth = 14
                .RowsDefaultCellStyle.Font = New Font("Tahoma", 8, FontStyle.Regular)

                If estrecho Then
                    .RowTemplate.Height = 16
                Else
                    .RowTemplate.Height = 21
                End If

            End With
        Catch ex As Exception

        End Try
        'dgv.ResumeLayout()
    End Sub


    <Extension()> _
    Public Sub FormatoColumna(ByRef dgv As DataGridView, ByVal NombreColumna As String, ByVal Titulo As String, ByVal Tipo As TiposColumna, Optional ByVal Ancho As Integer = 9999, Optional ByVal Orden As Integer = 9999)

        Try
            If Not dgv.Columns.Contains(NombreColumna) Then
                Return
            End If
            'dgv.SuspendLayout()
            With dgv.Columns(NombreColumna)
                .HeaderText = Titulo

                Select Case Tipo
                    Case Is = TiposColumna.Miles  '"Miles"
                        If Ancho = 9999 Then
                            .Width = 55
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Format = Miles
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With
                    Case Is = TiposColumna.Double2  '"Double2"
                        With .DefaultCellStyle
                            .Format = Double2
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With
                    Case Is = TiposColumna.FechaCorta  ' "FechaCorta"
                        If Ancho = 9999 Then
                            .Width = 70
                        Else
                            .Width = Ancho
                        End If

                        With .DefaultCellStyle
                            .Format = Fecha
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With
                    Case Is = TiposColumna.Hora  ' "Hora"
                        If Ancho = 9999 Then
                            .Width = 55
                        Else
                            .Width = Ancho
                        End If

                        With .DefaultCellStyle
                            .Format = Hora
                            .Alignment = DataGridViewContentAlignment.MiddleCenter
                        End With
                    Case Is = TiposColumna.Derecha  '"Derecha"
                        .Width = Ancho
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With
                    Case Is = TiposColumna.Izquierda  '"Izquierda" 
                        If Ancho = 9999 Then
                            .Width = 100
                        Else
                            .Width = Ancho
                        End If

                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With
                    Case Is = TiposColumna.Referencia  '"Referencia"
                        .Width = 300
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With
                    Case Is = TiposColumna.Lote  '"Lote" 
                        If Ancho = 9999 Then
                            .Width = 40
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With
                    Case Is = TiposColumna.Observaciones  '"Observaciones"
                        If Ancho = 9999 Then
                            .Width = 200
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With
                    Case Is = TiposColumna.QS  '"QS"
                        If Ancho = 9999 Then
                            .Width = 70
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With
                    Case Is = TiposColumna.Juliano  '"Juliano"
                        If Ancho = 9999 Then
                            .Width = 50
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With
                    Case Is = TiposColumna.Centro  '"Centro"
                        If Ancho = 9999 Then
                            .Width = 50
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleCenter
                        End With
                    Case Is = TiposColumna.Proveedor  '"Proveedor"
                        If Ancho = 9999 Then
                            .Width = 250
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With
                    Case Is = TiposColumna.Double5  '"Double5"
                        With .DefaultCellStyle
                            .Format = Double5
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With
                    Case Else
                        .Width = Ancho
                End Select
                If dgv.Columns.Count > Orden Then
                    .DisplayIndex = Orden
                End If
            End With
        Catch ex As Exception

        End Try
        'dgv.ResumeLayout()
    End Sub

    'Permite no tener que pasar el Titulo de la columna
    <Extension()> _
    Public Sub FormatoColumna(ByRef dgv As DataGridView, ByVal NombreColumna As String, ByVal Tipo As TiposColumna, Optional ByVal Ancho As Integer = 9999, Optional ByVal Orden As Integer = 9999)
        'dgv.SuspendLayout()
        Try
            If Not dgv.Columns.Contains(NombreColumna) Then
                Return
            End If

            With dgv.Columns(NombreColumna)
                Select Case Tipo

                    Case Is = TiposColumna.Miles  '"Miles"
                        If Ancho = 9999 Then
                            .Width = 55
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Format = Miles
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Double2  '"Double2"
                        If Ancho = 9999 Then
                            .Width = 52
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Format = Double2
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.FechaCorta  ' "FechaCorta"
                        If Ancho = 9999 Then
                            .Width = 70
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Format = Fecha
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Hora  ' "Hora"
                        If Ancho = 9999 Then
                            .Width = 55
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Format = Hora
                            .Alignment = DataGridViewContentAlignment.MiddleCenter
                        End With

                    Case Is = TiposColumna.Derecha  '"Derecha"
                        .Width = Ancho
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Izquierda  '"Izquierda"
                        If Ancho = 9999 Then
                            .Width = 100
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.Referencia  '"Referencia"
                        .Width = 300
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.Lote  '"Lote"
                        If Ancho = 9999 Then
                            .Width = 40
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Observaciones  '"Observaciones"
                        If Ancho = 9999 Then
                            .Width = 360
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.QS  '"QS"
                        If Ancho = 9999 Then
                            .Width = 70
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Juliano  '"Juliano"
                        If Ancho = 9999 Then
                            .Width = 50
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Centro  '"Centro"
                        If Ancho = 9999 Then
                            .Width = 50
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleCenter
                        End With

                    Case Is = TiposColumna.Double5  '"Double5"
                        With .DefaultCellStyle
                            .Format = Double5
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Descripcion  '"Descripcion"
                        If Ancho = 9999 Then
                            .Width = 360
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.Persona  '"Persona"
                        If Ancho = 9999 Then
                            .Width = 180
                        Else
                            .Width = Ancho
                        End If
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Else
                        .Width = Ancho
                End Select
                If dgv.Columns.Count > Orden Then
                    .DisplayIndex = Orden
                End If
            End With
        Catch ex As Exception
            '  MessageBox.Show("Error en formato columnas " & dgv.Name & "." & NombreColumna & Environment.NewLine() & Environment.NewLine() & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        'dgv.ResumeLayout()
    End Sub

    'Permite no tener que pasar el Titulo de la columna
    <Extension()> _
    Public Sub FormatoColumna(ByRef dgv As DataGridView, ByVal NombreColumna As String, ByVal NuevoNombreColumna As String, ByVal Tipo As TiposColumna, ByVal auto As Boolean, Optional ByVal Orden As Integer = 9999, Optional ancho_minimo As Integer = 40)
        'dgv.SuspendLayout()
        Try
            If Not dgv.Columns.Contains(NombreColumna) Then
                Return
            End If

            With dgv.Columns(NombreColumna)
                .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                .HeaderText = NuevoNombreColumna
                .MinimumWidth = ancho_minimo

                Select Case Tipo

                    Case Is = TiposColumna.Miles  '"Miles"
                        With .DefaultCellStyle
                            .Format = Miles
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Double2  '"Double2"
                        With .DefaultCellStyle
                            .Format = Double2
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.FechaCorta  ' "FechaCorta"
                        With .DefaultCellStyle
                            .Format = Fecha
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Hora  ' "Hora"
                        With .DefaultCellStyle
                            .Format = Hora
                            .Alignment = DataGridViewContentAlignment.MiddleCenter
                        End With

                    Case Is = TiposColumna.Derecha  '"Derecha"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Izquierda  '"Izquierda"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.Referencia  '"Referencia"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.Lote  '"Lote"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Observaciones  '"Observaciones"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.QS  '"QS"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Juliano  '"Juliano"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Centro  '"Centro"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleCenter
                        End With

                    Case Is = TiposColumna.Double5  '"Double5"
                        With .DefaultCellStyle
                            .Format = Double5
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Descripcion  '"Descripcion"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.Persona  '"Persona"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                End Select
                If dgv.Columns.Count > Orden Then
                    .DisplayIndex = Orden
                End If
            End With
        Catch ex As Exception
            '  MessageBox.Show("Error en formato columnas auto " & dgv.Name & "." & NombreColumna & Environment.NewLine() & Environment.NewLine() & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        'dgv.ResumeLayout()
    End Sub

    'Permite no tener que pasar el Titulo de la columna
    <Extension()> _
    Public Sub FormatoColumna(ByRef dgv As DataGridView, ByVal NombreColumna As String, ByVal Tipo As TiposColumna, ByVal auto As Boolean, Optional ByVal Orden As Integer = 9999, Optional ancho_minimo As Integer = 40)
        'dgv.SuspendLayout()
        Try
            If Not dgv.Columns.Contains(NombreColumna) Then
                Return
            End If

            With dgv.Columns(NombreColumna)
                If auto Then
                    .AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
                    .MinimumWidth = ancho_minimo
                End If

                Select Case Tipo

                    Case Is = TiposColumna.Miles  '"Miles"
                        With .DefaultCellStyle
                            .Format = Miles
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Double2  '"Double2"
                        With .DefaultCellStyle
                            .Format = Double2
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.FechaCorta  ' "FechaCorta"
                        With .DefaultCellStyle
                            .Format = Fecha
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Hora  ' "Hora"
                        With .DefaultCellStyle
                            .Format = Hora
                            .Alignment = DataGridViewContentAlignment.MiddleCenter
                        End With

                    Case Is = TiposColumna.Derecha  '"Derecha"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Izquierda  '"Izquierda"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.Referencia  '"Referencia"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.Lote  '"Lote"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Observaciones  '"Observaciones"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.QS  '"QS"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Juliano  '"Juliano"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Centro  '"Centro"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleCenter
                        End With

                    Case Is = TiposColumna.Double5  '"Double5"
                        With .DefaultCellStyle
                            .Format = Double5
                            .Alignment = DataGridViewContentAlignment.MiddleRight
                        End With

                    Case Is = TiposColumna.Descripcion  '"Descripcion"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With

                    Case Is = TiposColumna.Persona  '"Persona"
                        With .DefaultCellStyle
                            .Alignment = DataGridViewContentAlignment.MiddleLeft
                        End With
                End Select
                If dgv.Columns.Count > Orden Then
                    .DisplayIndex = Orden
                End If
            End With
        Catch ex As Exception
            ' MessageBox.Show("Error en formato columnas auto " & dgv.Name & "." & NombreColumna & Environment.NewLine() & Environment.NewLine() & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        'dgv.ResumeLayout()
    End Sub

    <Extension()> _
    Public Sub SortMode(ByRef dgv As DataGridView, ByVal sortMode As DataGridViewColumnSortMode)
        For Each col As DataGridViewColumn In dgv.Columns
            col.SortMode = sortMode
        Next
    End Sub


    <Extension()> _
    Public Sub MoveLast(ByRef dgv As DataGridView)
        If dgv.Rows.Count > 0 Then
            dgv.BindingContext.Item(dgv.DataSource).Position = dgv.RowCount - 1
        End If
    End Sub

    <Extension()> _
    Public Sub MoveFirst(ByVal dgv As DataGridView)
        If dgv.Rows.Count > 0 Then
            dgv.BindingContext.Item(dgv.DataSource).Position = 0
        End If
    End Sub


End Module
