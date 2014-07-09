Imports System.IO


'ClosedXML makes it easier for developers to create Excel 2007/2010 files. 
'It provides a nice object oriented way to manipulate the files (similar to VBA) without dealing with the hassles of XML Documents.
'It can be used by any .NET language like C# and Visual Basic (VB).
'http://closedxml.codeplex.com/

Public Module ExcelExtension
    ''' <summary>
    ''' Crea un archivo de texto delimitado por comas con
    ''' el contenido de una hoja de cálculo de Excel.
    ''' </summary>
    ''' <author>Enrique Martínez Montejo - 2011</author>
    ''' <param name="fileNameExcel">Ruta del archivo de Excel.</param>
    ''' <param name="sheetName">Nombre de la hoja de cálculo que se desea exportar.</param>
    ''' <param name="fileNameCsv">Ruta del archivo de texto delimitado por comas.</param>
    ''' <remarks></remarks>

    'Cuando desee llamar al procedimiento para crear un archivo CSV, lo haría de la siguiente manera:
    'Try
    '    ExportToCsv("C:\Mis documentos\Libro1.xls", "Hoja1", "C:\Archivos\Archivo.csv")
    'Catch ex As Exception
    '    MessageBox.Show(ex.Message)
    'End Try

    Public Sub ExportToCsv( _
        ByVal fileNameExcel As String, _
        ByVal sheetName As String, _
        ByVal fileNameCsv As String)

        If Not (IO.File.Exists(fileNameExcel)) Then _
                Throw New FileNotFoundException( _
                "No existe el archivo de Excel especificado")

        If String.IsNullOrEmpty(sheetName) Then _
                Throw New ArgumentNullException("sheetName")

        If IO.File.Exists(fileNameCsv) Then _
            Throw New IOException( _
            "Ya existe un archivo csv con el mismo nombre " & _
            "en la carpeta especificada.")

        Dim app As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim wb As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim ws As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Try
            app = New Microsoft.Office.Interop.Excel.Application
            wb = app.Workbooks.Open(fileNameExcel)
            ws = DirectCast(wb.Worksheets(sheetName), Microsoft.Office.Interop.Excel.Worksheet)

            ws.SaveAs(fileNameCsv, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV)

        Catch ex As Exception
            Throw
        Finally
            If Not ws Is Nothing Then _
                ReleaseComObject(ws)

            If Not wb Is Nothing Then
                ' Indicamos que el libro ya ha sido guardado.
                wb.Saved = True
                wb.Close()
                ' Liberamos el objeto.
                ReleaseComObject(wb)
            End If

            If Not app Is Nothing Then
                app.Quit()
                ' Liberamos el objeto.
                ReleaseComObject(app)
            End If

        End Try

    End Sub

    ''' <summary>
    ''' Disminuye el recuento de referencias del contenedor
    ''' RCW asociado al objeto COM especificado.
    ''' </summary>
    ''' <param name="obj">Objeto COM válido.</param>
    ''' <remarks></remarks>
    Private Sub ReleaseComObject(ByVal obj As Object)
        If obj Is Nothing Then Return

        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
        Catch ex As Exception
            ' Deshechamos devolver la posible excepción
            ' si no es un objeto COM válido.
        Finally
            obj = Nothing

        End Try

    End Sub
End Module
