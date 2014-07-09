

Public Class SMS
    Public Sub enviarSMS(ByVal numero As String, ByVal mensaje As String, ByVal puerto As String)
        Dim serialport As System.IO.Ports.SerialPort

        Try
            serialport = New System.IO.Ports.SerialPort()

            With serialport
                .PortName = puerto ' Puerto de comunicaciones COM1, COM2...
                .BaudRate = 9600 ' velocidad del puerto
                .Parity = System.IO.Ports.Parity.None ' Paridad
                .StopBits = System.IO.Ports.StopBits.One ' Bit de para
                .DataBits = 8 ' Bits de datos
                Threading.Thread.Sleep(1000) 'Espera 1 segundo
                If Not .IsOpen Then ' Si el puerto no esta abierto
                    .Open() ' Abrir el puerto
                End If
                Threading.Thread.Sleep(1000)
                .Write("ATZ0" & Environment.NewLine)
                .WriteLine("AT+CMGF=1" & Environment.NewLine) ' Configurando envio de SMS modo texto

                .Write("AT+CMGS=" & Convert.ToChar(34) & numero & Convert.ToChar(34) & Environment.NewLine) ' Numero al que se envia el SMS
                .Write(mensaje & Convert.ToChar(26) & Environment.NewLine)

                .Close()
                MessageBox.Show("Mensaje enviado", "", MessageBoxButtons.OK, MessageBoxIcon.Information)

            End With

        Catch Ex As IO.IOException
            Console.WriteLine(Ex.Message)
        End Try

    End Sub
End Class
