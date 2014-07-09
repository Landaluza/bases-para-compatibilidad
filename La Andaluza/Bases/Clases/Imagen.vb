Option Strict On



Public Class Imagen
    ' Métodos compartidos 
    Public Function Image2Bytes(ByVal img As Image) As Byte()
        Dim sTemp As String = System.IO.Path.GetTempFileName()
        Dim fs As New System.IO.FileStream(sTemp, System.IO.FileMode.OpenOrCreate, System.IO.FileAccess.ReadWrite)
        img.Save(fs, System.Drawing.Imaging.ImageFormat.Png)


        fs.Position = 0
        '
        Dim imgLength As Integer = Convert.ToInt32(fs.Length)
        Dim bytes(0 To imgLength - 1) As Byte
        fs.Read(bytes, 0, imgLength)
        fs.Close()
        fs.Dispose()
        Return bytes
    End Function


    Public Function Bytes2Image(ByVal bytes() As Byte) As Image
        If bytes Is Nothing Then Return Nothing
        '
        Dim ms As New System.IO.MemoryStream(bytes)
        Dim bm As Bitmap = Nothing
        Try
            bm = New Bitmap(ms)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message)
        End Try
        Return bm
    End Function
End Class
