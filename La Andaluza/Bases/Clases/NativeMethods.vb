Public Class NativeMethods
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer

    'http://social.msdn.microsoft.com/Forums/en-US/csharpgeneral/thread/ea8b0fd5-a660-46f9-9dcb-d525cc22dcbd/
    <Runtime.InteropServices.DllImport("user32.dll")> _
    Public Shared Function FindWindow(ByVal lpClassName As String, ByVal lpWindowName As String) As IntPtr
    End Function
End Class
