Public Class Encriptacion
    'Encripta una cadena de caracteres.
    'S = Cadena a encriptar
    'P = Password
    Public Function EncriptarStr(ByVal S As String) As String
        Dim P As String = "123456789"
        Dim I As Integer, R As String
        Dim C1 As Integer, C2 As Integer
        R = ""
        'If P.Length > 0 Then
        For I = 1 To S.Length
            C1 = Convert.ToInt32(S.Substring(I, 1))
            If I > P.Length Then
                C2 = Convert.ToInt32(P.Substring(I Mod P.Length + 1, 1))
            Else
                C2 = Convert.ToInt32(P.Substring(I, 1))
            End If
            C1 = C1 + C2 + 64
            If C1 > 255 Then C1 = C1 - 256
            R = R + Convert.ToChar(C1)
        Next I
        'Else
        'R = S
        'End If
        EncriptarStr = R
    End Function

    'Desencripta una cadena de caracteres.
    'S = Cadena a desencriptar
    'P = Password
    Public Function DesencriptarStr(ByVal S As String) As String
        Dim P As String = "123456789"
        Dim I As Integer, R As String
        Dim C1 As Integer, C2 As Integer
        R = ""
        'If P.Length > 0 Then
        For I = 1 To S.Length
            C1 = Convert.ToInt32(S.Substring(I, 1))
            If I > P.Length Then
                C2 = Convert.ToInt32(P.Substring(I Mod P.Length + 1, 1))
            Else
                C2 = Convert.ToInt32(P.Substring(I, 1))
            End If
            C1 = C1 - C2 - 64
            If Math.Sign(C1) = -1 Then C1 = 256 + C1
            R = R + Convert.ToChar(C1)
        Next I
        'Else
        '    R = S
        'End If
        DesencriptarStr = R
    End Function
End Class
