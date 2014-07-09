

Public Class MicrosoftOfficeExporter
    Public Const ONE As String = "One"
    Public Const MANY As String = "Many"

    Public Sub ExportToExcel(ByVal Type As String, ByVal Head As String, ByVal Grid As DataGridView)
        Dim frm As New frmEspera("Preparando archivo")
        frm.Show()
        Try
            If Grid.ColumnCount > 0 Then
                Dim oExcel As Microsoft.Office.Interop.Excel.Application
                Dim oBook As Microsoft.Office.Interop.Excel.Workbook
                Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet
                oExcel = CreateObject("Excel.Application")
                oExcel.Visible = False
                oBook = oExcel.Workbooks.Add
                oSheet = oBook.Worksheets(1)

                oSheet.Range("B2").Value = Head
                oSheet.Range("B2").Font.Bold = True

                If Type = MicrosoftOfficeExporter.ONE Then

                    For s As Integer = 0 To Grid.ColumnCount - 1
                        oSheet.Range("B" & 4 + s).Value = Grid.Columns(s).HeaderText
                        oSheet.Range("C" & 4 + s).Value = Grid.SelectedRows.Item(0).Cells.Item(s).Value
                    Next

                    oSheet.Columns("B:B").ColumnWidth = 20
                    oSheet.Columns("C:C").ColumnWidth = 20

                ElseIf Type = MicrosoftOfficeExporter.MANY Then

                    Dim DataArrayHead(0, 0 To Grid.ColumnCount - 1) As Object
                    For s As Integer = 0 To Grid.ColumnCount - 1
                        DataArrayHead(0, s) = Grid.Columns(s).HeaderText
                    Next

                    oSheet.Range("B4").Resize(1, Grid.ColumnCount).Value = DataArrayHead
                    oSheet.Range("B4").Resize(1, Grid.ColumnCount).Font.Bold = True

                    Dim DataArray(0 To Grid.RowCount - 1, 0 To Grid.ColumnCount - 1) As Object
                    For r As Integer = 0 To Grid.RowCount - 1
                        For s As Integer = 0 To Grid.ColumnCount - 1
                            DataArray(r, s) = Grid.Rows.Item(r).Cells.Item(s).Value
                        Next
                    Next

                    oSheet.Range("B5").Resize(Grid.RowCount, Grid.ColumnCount).Value = DataArray
                    oSheet.Range("B5").Resize(Grid.RowCount, Grid.ColumnCount).ColumnWidth = 20

                End If

                oExcel.Visible = True
                oExcel = Nothing
                oSheet = Nothing
            End If
        Catch
            MessageBox.Show("Erroe." & Environment.NewLine, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            frm.Close()
        End Try
    End Sub

    Public Sub ExportToWord(ByVal Type As String, ByVal Head As String, ByVal Grid As DataGridView)
        Dim frm As New frmEspera("Preparando archivo")
        frm.Show()
        Try
            If Grid.ColumnCount > 0 Then
                Dim oWord As Microsoft.Office.Interop.Word.Application
                Dim oDoc As Microsoft.Office.Interop.Word.Document
                Dim oTable As Microsoft.Office.Interop.Word.Table
                Dim oPara1 As Microsoft.Office.Interop.Word.Paragraph

                oWord = CreateObject("Word.Application")
                oWord.Visible = False
                oDoc = oWord.Documents.Add
                With oDoc.PageSetup
                    Try
                        .PaperSize = Microsoft.Office.Interop.Word.WdPaperSize.wdPaperA4
                    Catch ex As Exception
                    End Try

                    Try
                        .Orientation = Microsoft.Office.Interop.Word.WdOrientation.wdOrientLandscape
                    Catch ex As Exception
                    End Try

                    Try
                        .LeftMargin = oWord.CentimetersToPoints(1)
                    Catch ex As Exception
                    End Try

                    Try
                        .BottomMargin = oWord.CentimetersToPoints(1)
                    Catch ex As Exception

                    End Try

                    Try
                        .TopMargin = oWord.CentimetersToPoints(1)
                    Catch ex As Exception
                    End Try

                End With

                oPara1 = oDoc.Content.Paragraphs.Add
                oPara1.Range.Text = Head
                oPara1.Range.Font.Bold = False
                oPara1.Range.Font.Size = 8
                oPara1.Format.SpaceAfter = 8
                oPara1.Range.InsertParagraphAfter()

                If Type = MicrosoftOfficeExporter.ONE Then

                    oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, Grid.ColumnCount, 2)
                    oTable.Range.ParagraphFormat.SpaceAfter = 2

                    Dim cont As Integer = 1

                    For i As Integer = 0 To Grid.ColumnCount - 1
                        If Grid.Columns(i).Visible Then
                            oTable.Cell(1, cont).Range.InsertAfter(Grid.Columns(i).HeaderText)

                            If Grid.Rows(1).Cells(i).Value Is System.Convert.DBNull Then
                                oTable.Cell(2, cont).Range.Text = ""
                            Else
                                oTable.Cell(2, cont).Range.Text = Grid.Rows(0).Cells(i).Value.ToString
                            End If

                            cont = cont + 1
                        End If
                    Next

                    oTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleDouble
                    oTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleDouble

                    oTable.AllowAutoFit = True
                    oTable.Columns.AutoFit()

                ElseIf Type = MicrosoftOfficeExporter.MANY Then

                    oTable = oDoc.Tables.Add(oDoc.Bookmarks.Item("\endofdoc").Range, Grid.RowCount + 1, Grid.ColumnCount)
                    oTable.Range.ParagraphFormat.SpaceAfter = 2

                    Dim cont As Integer = 1

                    For i As Integer = 1 To Grid.ColumnCount - 1
                        If Grid.Columns(i).Visible Then
                            oTable.Cell(1, cont).Range.InsertAfter(Grid.Columns(i).HeaderText)
                            cont = cont + 1
                        End If
                    Next


                    For r As Integer = 0 To Grid.RowCount - 1
                        cont = 1

                        For s As Integer = 0 To Grid.ColumnCount - 1
                            If Grid.Columns(s).Visible Then
                                oTable.Cell(r + 2, cont).Range.Text = If(Grid.Rows.Item(r).Cells.Item(s).Value Is System.Convert.DBNull, "", Grid.Rows.Item(r).Cells.Item(s).Value)
                                cont = cont + 1
                            End If
                        Next
                    Next

                    oTable.Rows.Item(1).Cells.Shading.BackgroundPatternColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdGray25

                    oTable.Borders.InsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleDouble
                    oTable.Borders.OutsideLineStyle = Microsoft.Office.Interop.Word.WdLineStyle.wdLineStyleDouble

                    oTable.AllowAutoFit = True
                    oTable.Columns.AutoFit()

                End If

                oWord.Visible = True
                oWord.NormalTemplate.Saved = True
                oWord = Nothing
                oDoc = Nothing
            End If
        Catch
            MessageBox.Show("Erroe." & Environment.NewLine, "", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            frm.Close()
        End Try
    End Sub

    ''' Crea un archivo de texto delimitado con el contenido de un objeto DataTable.
    ''' 
    ''' <param name="fileName" Ruta y nombre del archivo de texto.
    ''' <param name="dt" Un objeto DataTable válido.
    ''' <param name="separatorChar" El carácter delimitador de los campos.
    ''' <param name="hdr" Indica si la primera fila contiene el nombre de los campos.
    ''' <param name="textDelimiter" Indica si los campos alfanuméricos deben aparecer entre comillas dobles.
    ''' 
    Public Function CreateTextDelimiterFile(ByVal fileName As String, _
                                             ByVal dt As DataTable, _
                                             ByVal separatorChar As Char, _
                                             ByVal hdr As Boolean, _
                                             ByVal textDelimiter As Boolean) As Boolean

        ' Si no se ha especificado un nombre de archivo, o el objeto DataTable no es válido, provocamos
        ' una excepción de argumentos no válidos.
        If fileName = String.Empty OrElse _
           dt Is Nothing Then Throw New System.ArgumentException("Argumentos no válidos.")
        ' Si el archivo existe, solicito confirmación para sobreescribirlo.
        'If IO.File.Exists(fileName) Then
        '    If IO.File.Exists(fileName) Then
        '        If MessageBox.Show("Ya existe un archivo de texto con el mismo nombre." & Environment.NewLine & _
        '                           "¿Desea sobrescribirlo?", _
        '                           "Crear archivo de texto delimitado", _
        '                           MessageBoxButtons.YesNo, _
        '                           MessageBoxIcon.Information) = Windows.Forms.DialogResult.No Then Return False
        '    End If
        'End If

        Dim sw As System.IO.StreamWriter

        Try
            Dim col As Integer = 0
            Dim value As String = String.Empty
            ' Creamos el archivo de texto con la codificación por defecto.
            sw = New System.IO.StreamWriter(fileName, False, System.Text.Encoding.Default)

            If hdr Then
                ' La primera línea del archivo de texto contiene el nombre de los campos.
                For Each dc As DataColumn In dt.Columns
                    If textDelimiter Then
                        ' Incluimos el nombre del campo entre el caracter delimitador de texto especificado.
                        value &= """" & dc.ColumnName & """" & separatorChar
                    Else
                        ' No se incluye caracter delimitador de texto alguno.
                        value &= dc.ColumnName & separatorChar
                    End If
                Next
                sw.WriteLine(value.Remove(value.Length - 1, 1))
                value = String.Empty
            End If

            ' Recorremos todas las filas del objeto DataTable incluido en el conjunto de datos.
            For Each dr As DataRow In dt.Rows
                For Each dc As DataColumn In dt.Columns
                    If dc.DataType Is System.Type.GetType("System.String") And _
                       textDelimiter Then
                        ' Incluimos el dato alfanumérico entre el caracter delimitador de texto especificado.
                        value &= """" & dr.Item(col).ToString & """" & separatorChar
                    Else
                        ' No se incluye caracter delimitador de texto alguno
                        value &= dr.Item(col).ToString & separatorChar
                    End If
                    ' Siguiente columna
                    col += 1
                Next
                ' Al escribir los datos en el archivo, elimino el último carácter delimitador de la fila.
                sw.WriteLine(value.Remove(value.Length - 1, 1))
                value = String.Empty
                col = 0
            Next ' Siguiente fila

            ' Nos aseguramos de cerrar el archivo
            sw.Close()

            ' Se ha creado con éxito el archivo de texto.
            Return True

        Catch ex As Exception
            messageBox.show(ex.ToString)
            Return False
        Finally
            sw = Nothing
        End Try

    End Function

    Public Sub ExportarTablaExcel(ByVal dt As DataTable, ByVal Archivo As String, ByVal NombrePestaña As String, ByVal NombreHoja As String, ByVal NombreTxt As String)
        ' Referenciamos el objeto DataTable enlazado a un control DataGridView
        'Dim dt As DataTable = DirectCast(dgvPalet.DataSource, DataTable)

        ' Indicamos nuestra intención de crear un archivo de texto delimitado por comas,
        ' así como que en la primera línea aparezcan los nombres de los campos, y que
        ' aquellos que sean alfanuméricos, sean encerrados entre comillas dobles.
        '

        'Dim bln As Boolean = CreateTextDelimiterFile(Archivo, dt, ";"c, True, True)  
        'en MAM el separador debe ser ; para que Excel cree todas las columnas , sin embargo en MAM1 el separador es la coma.
        Dim bln As Boolean = CreateTextDelimiterFile(Archivo, dt, ","c, True, True)
        If bln Then
            'MessageBox.Show("Se ha creado satisfactoriamente el archivo de texto.")
        Else
            MessageBox.Show("No se ha creado el archivo de texto delimitado.")
        End If

        ' Creamos una conexión OleDb con el archivo de texto.
        ' Observa que en el parámetro Data Source se especifica
        ' sólo la ruta de la carpeta que almacena el archivo de texto.
        '
        Using cnn As New System.Data.OleDb.OleDbConnection( _
                                         "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                         "Data Source=c:\;" & _
                                         "Extended Properties='TEXT;'")

            Dim sql As String = _
                             "SELECT * INTO " & NombrePestaña & _
                             " IN ''[Excel 8.0;DATABASE=" & NombreHoja & "]" & _
                             "FROM " & NombreTxt
            Try
                ' Creamos un objeto OleDbCommand para ejecutar la consulta
                Dim cmd As New System.Data.OleDb.OleDbCommand(sql, cnn)
                ' Abrimos la conexión
                cnn.Open()
                ' Ejecutamos la consulta.
                Dim n As Int32 = cmd.ExecuteNonQuery
                'MessageBox.Show("Nº Registros afectados: " & CStr(n))
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End Using
    End Sub
End Class
