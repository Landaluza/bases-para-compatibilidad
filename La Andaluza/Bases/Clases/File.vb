
Imports System
Imports System.IO
Imports System.Text
Imports System.Collections.Generic
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.Serialization

Public Class File
    Private m_path As String

    Public Property Path As String
        Get
            Return Me.m_path
        End Get
        Set(value As String)
            Me.m_path = value
        End Set
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal path As String)
        Me.Path = path
    End Sub

    Public Sub saveObject(ByRef obj As Object, ByVal filename As String)


        Dim ser As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter

        Using fs As New IO.StreamWriter(filename)
            ser.Serialize(fs.BaseStream, obj)
        End Using
    End Sub

    Public Function loadObject(ByVal filename As String) As Object
        Dim ser As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
        Dim obj As Object = Nothing

        Using fs As New IO.StreamReader(filename)
            obj = ser.Deserialize(fs.BaseStream)
        End Using

        Return obj
    End Function


    'Sub VerDocumento(ByVal Filename As String)
    '    Try
    '        Dim psi As New ProcessStartInfo()
    '        psi.UseShellExecute = True
    '        psi.FileName = Filename
    '        Process.Start(psi)
    '    Catch ex As Exception
    '        messageBox.show(ex.ToString, "Error", MessageBoxButtons.OK,MessageBoxIcon.Error )
    '    End Try
    'End Sub

    Public Sub open(ByVal path As String)
        Dim frm As New frmEspera("Abriendo archivo, espere unos instantes")
        frm.Show()
        Try
            'Dim extension As String = path.Substring(path.Length - 3, 3)

            'If extension = "PDF" Or extension = "pdf" Then
            '    Process.Start(path)

            'ElseIf extension = "doc" Or extension = "ocx" Then
            '    Dim oApp As New Word.Application
            '    Dim oWBa As New Word.Document
            '    oApp.Visible = True
            '    oWBa = oApp.Documents.Open(path.ToString, , False, False, , , True)
            '    oWBa.Activate()

            'Else
            '    Process.Start(path)
            'End If
            If IO.File.Exists(path) Then
                Process.Start(path)
            Else
                MessageBox.Show("No se encontro el archivo especificado. " & Environment.NewLine() & path, "", MessageBoxButtons.OK, MessageBoxIcon.Information) 'Error")
            End If
        Catch ex As Exception
            MessageBox.Show("No se pudo abrir el archivo. Detalles: " & Environment.NewLine() & ex.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Exclamation) 'Error")
        Finally
            frm.Close()
        End Try
    End Sub

    Public Sub open()
        Me.open(Me.Path)
    End Sub

    Public Sub saveImage(ByVal bm As Bitmap, ByVal ruta As String)
        Dim piccaptura As New PictureBox
        piccaptura.Image = bm
        piccaptura.Image.Save(ruta, Imaging.ImageFormat.Jpeg)
    End Sub

    Public Shared Function abrir_carpeta(ByVal ruta As String) As String
        Dim MyFolderBrowser As New System.Windows.Forms.FolderBrowserDialog

        MyFolderBrowser.Description = "Selecciona el archivo"

        MyFolderBrowser.RootFolder = Environment.SpecialFolder.MyComputer
        MyFolderBrowser.SelectedPath = ruta

        MyFolderBrowser.ShowNewFolderButton = False

        Dim dlgResult As DialogResult = MyFolderBrowser.ShowDialog()

        If dlgResult = Windows.Forms.DialogResult.OK Then
            Return MyFolderBrowser.SelectedPath
        Else
            Return ""
        End If
    End Function

    Public Shared Function Elegir_archivo(Optional ByVal directorioInicial As String = "C:\") As String
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String = String.Empty

        fd.Title = "Elija el archivo"
        fd.InitialDirectory = directorioInicial
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName
        End If

        Return strFileName
    End Function

    'Private Function AbrirArchivo(ByVal arch As String) As Integer
    '    Dim hay_error As Integer
    '    hay_error = NativeMethods.ShellExecute(123456, "Open", arch, "", "", 3)
    '    Return hay_error
    'End Function

    Public Function SeleccionarArchivo(ByVal OpcionForm As String, ByVal DoctoUbicacionfisicaCuadroDeTexto As String, ByVal archactual As String, Optional ByVal ruta As String = "") As String
        Dim f As New File
        If OpcionForm = "Ver" Then
            'Dim hay_error As Integer = AbrirArchivo(DoctoUbicacionfisicaCuadroDeTexto)
            'Return DoctoUbicacionfisicaCuadroDeTexto
            open(DoctoUbicacionfisicaCuadroDeTexto)
            Return True
        Else
            Dim ofd As New OpenFileDialog
            Dim arch As String = ""
            ofd.CheckFileExists = True
            ofd.CheckPathExists = True
            ofd.Multiselect = False

            If ruta <> "" Then
                ofd.InitialDirectory = ruta
            End If

            If archactual <> "" Then
                ofd.FileName = archactual
            End If

            Dim ok As DialogResult = ofd.ShowDialog()
            If ok = Windows.Forms.DialogResult.OK Then
                arch = ofd.FileName
                open(arch)
                Return arch
            End If
            Return ""
        End If
    End Function
End Class
