
'Ejemplo uso
'oTabla.mam_FormatoLinea(3, 1, 120, 10, WdParagraphAlignment.wdAlignParagraphLeft, "Lote producto:")
Public Module WordExtension
    <Extension()> _
    Public Sub mam_FormatoLinea(ByVal oTabla As Microsoft.Office.Interop.Word.Table,
                                ByVal m_Fila As Integer,
                                ByVal m_Row As Integer,
                                ByVal m_Pixels As Integer,
                                ByVal m_FontSize As Integer,
                                ByVal m_Alignement As Microsoft.Office.Interop.Word.WdParagraphAlignment,
                                Optional ByVal m_Cadena As String = "")
        Try

            With oTabla
                .Cell(m_Fila, m_Row).Range.Text = m_Cadena
                .Cell(m_Fila, m_Row).Range.Font.Size = m_FontSize
                .Cell(m_Fila, m_Row).Range.ParagraphFormat.Alignment = m_Alignement
                .Cell(m_Fila, m_Row).Width = m_Pixels
            End With
        Catch ex As Exception

        End Try
    End Sub

    <Extension()> _
    Public Sub mam_FormatoLineaTitulo(ByVal oTabla As Microsoft.Office.Interop.Word.Table,
                                ByVal m_Fila As Integer,
                                ByVal m_Row As Integer,
                                ByVal m_Pixels As Integer,
                                ByVal m_FontSize As Integer,
                                ByVal m_Alignement As Microsoft.Office.Interop.Word.WdParagraphAlignment,
                                ByVal Titulo As Boolean,
                                Optional ByVal m_Cadena As String = "")

        With oTabla
            .Cell(m_Fila, m_Row).Width = m_Pixels
            .Cell(m_Fila, m_Row).Range.Font.Size = m_FontSize
            .Cell(m_Fila, m_Row).Range.ParagraphFormat.Alignment = m_Alignement
            .Cell(m_Fila, m_Row).Range.Text = m_Cadena
            If Titulo Then
                .Cell(m_Fila, m_Row).Shading.BackgroundPatternColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdGray25
                .Cell(m_Fila, m_Row).Range.Font.Bold = 1
            End If
        End With
    End Sub
End Module
