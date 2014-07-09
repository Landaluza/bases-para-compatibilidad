
Public Class Printer
    Public Sub ImprimirForm(ByRef f As Form)
        Dim pf As New Microsoft.VisualBasic.PowerPacks.Printing.PrintForm
        With pf
            With .PrinterSettings.DefaultPageSettings
                .Landscape = True
                .Margins.Right = 1
                .Margins.Left = 10
                .Margins.Bottom = 1
                .Margins.Top = 1
            End With
            .Form = f
            .PrintAction = Printing.PrintAction.PrintToPrinter
            .Print(f, Microsoft.VisualBasic.PowerPacks.Printing.PrintForm.PrintOption.ClientAreaOnly)
        End With
    End Sub

    Public Sub imprimirGrilla(ByVal Datos As DataGridView, _
                ByVal Encabezado As String(), _
                Optional ByVal PageProperties As System.Drawing.Printing.PageSettings = Nothing, _
                Optional ByVal showPrinters As Boolean = True, _
                Optional ByVal ShowPreview As Boolean = True)
        Try
            Dim pd As System.Drawing.Printing.PrintDocument = New DataGridViewPrintDocument(Datos, Encabezado)
            Try
                If PageProperties Is Nothing Then
                    PageProperties = New System.Drawing.Printing.PageSettings
                    PageProperties = pd.DefaultPageSettings
                    With PageProperties.Margins
                        .Left = 15
                        .Top = 15
                        .Bottom = 20
                        .Right = 5
                    End With
                End If
                pd.DefaultPageSettings = PageProperties
                If showPrinters Then
                    Dim dlg As New PrintDialog()
                    dlg.Document = pd
                    Dim result As DialogResult = dlg.ShowDialog()
                    If result = System.Windows.Forms.DialogResult.OK Then

                        pd = dlg.Document 'documento ya formateado
                        If Not ShowPreview Then
                            pd.Print()
                        Else
                            Dim dlg2 As New PrintPreviewDialog
                            With dlg2
                                .Document = pd
                                .WindowState = FormWindowState.Maximized
                                .PrintPreviewControl.Zoom = 1
                                .Text = "REPORTE GENERAL DE DATOS"
                                .ShowDialog()
                            End With
                        End If
                    End If
                Else
                    If Not ShowPreview Then
                        pd.Print()
                    Else
                        Dim dlg2 As New PrintPreviewDialog
                        With dlg2
                            .Document = pd
                            .WindowState = FormWindowState.Maximized
                            .PrintPreviewControl.Zoom = 1
                            .Text = "REPORTE GENERAL DE DATOS"
                            .ShowDialog()
                        End With
                    End If
                End If
            Finally

            End Try
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
End Class
