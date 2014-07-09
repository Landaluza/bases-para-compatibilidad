Public Module TabPageExtension

    <Extension()> _
    Public Function Close(ByVal tp As TabPage) As Boolean
        Dim frm As System.Windows.Forms.Form

        For Each obj As Object In tp.Controls
            If TypeOf obj Is System.Windows.Forms.Form Then
                frm = CType(obj, System.Windows.Forms.Form)
                frm.Close()

                Try
                    If frm.IsDisposed Then
                        Dim tc As TabControl = CType(tp.Parent, TabControl)
                        tc.TabPages.Remove(tp)
                    End If
                Catch ex As Exception
                End Try
                Return True
            End If
        Next


        Return False
    End Function
End Module
