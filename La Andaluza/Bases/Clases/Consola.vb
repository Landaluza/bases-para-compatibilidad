Public Class Consola

    Public Sub Teclea(ByRef Tecla As String)
        My.Computer.Keyboard.SendKeys(Tecla, True)
        System.Threading.Thread.Sleep(100)
    End Sub

    Public Sub CambioPantalla()
        System.Threading.Thread.Sleep(1000)
    End Sub


    Public Sub Espera(ByRef TiempoEspera As Integer)
        System.Threading.Thread.Sleep(TiempoEspera)
    End Sub


    Private Declare Sub Sleep Lib "kernel32" ( _
           ByVal dwMilliseconds As Long)
    Private Declare Function GetExitCodeProcess Lib "kernel32" ( _
        ByVal hProcess As Long, ByVal lpExitCode As Long) As Long
    Private Declare Function timeGetTime Lib "winmm.dll" () As Long
    Private Declare Function OpenProcess Lib "kernel32" ( _
        ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
    Private Const STILL_ACTIVE As Integer = &H103
    Private Const PROCESS_QUERY_INFORMATION As Integer = &H400
    Private Declare Function CloseHandle Lib "kernel32" ( _
        ByVal hObject As Long) As Long

    '    Public Function ShellAndWait( _
    '     ByVal sShell As String, _
    '            Optional ByVal eWindowStyle As Integer = vbNormalFocus, _
    '            Optional ByRef sError As String = "", _
    '            Optional ByVal lTimeOut As Long = 2000000000 _
    '        ) As Boolean
    '        Dim hProcess As Long
    '        Dim lR As Long
    '        Dim lTimeStart As Long
    '        Dim bSuccess As Boolean

    '        On Error GoTo ShellAndWaitError

    '        ' This is v2 which is somewhat more reliable:
    '        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION, False, Shell(sShell, eWindowStyle))
    '        If hProcess = 0 Then
    '            sError = "This program could not determine whether the process started." & _
    '                 "Please watch the program and check it completes."
    '            ' Only fail if there is an error - this can happen
    '            ' when the program completes too quickly.
    '        Else
    '            bSuccess = True
    '            lTimeStart = timeGetTime()
    '            Do
    '                ' Get the status of the process
    '                GetExitCodeProcess(hProcess, lR)
    '                ' Sleep during wait to ensure the other process gets
    '                ' processor slice:
    'DoEvents:       Sleep(100)
    '                If timeGetTime() - lTimeStart > lTimeOut Then
    '                    ' Too long!
    '                    sError = "The process has timed out."
    '                    lR = 0
    '                    bSuccess = False
    '                End If
    '            Loop While lR = STILL_ACTIVE
    '        End If
    '        ShellAndWait = bSuccess

    '        Exit Function

    'ShellAndWaitError:
    '        sError = Err.Description
    '        Exit Function
    '    End Function
End Class
