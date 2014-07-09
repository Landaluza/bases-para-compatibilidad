Public Module dtpExtension


    <Extension()> _
    Public Sub activarFoco(ByVal dtp As DateTimePicker)

        If dtp.Format <> DateTimePickerFormat.Short And dtp.Format <> DateTimePickerFormat.Long Then
            AddHandler dtp.ValueChanged, AddressOf moverFoco
        End If
    End Sub

    Private Sub moverFoco(sender As Object, args As EventArgs)
        If TypeOf sender Is DateTimePicker Then
            Dim dtp As DateTimePicker = CType(sender, DateTimePicker)
            Try
                If dtp.Focused Then
                    If dtp.Value.ToString("hh").Length = 2 Then
                        System.Windows.Forms.SendKeys.SendWait("{Right}") '"{Right}")
                    End If
                End If
            Catch ex As Exception
            End Try
        End If
    End Sub

End Module
