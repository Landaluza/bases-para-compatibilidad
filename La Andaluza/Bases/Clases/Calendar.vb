Public Class Calendar
    Public Shared testDateTime As Boolean

    Public Function TestDate() As Boolean


        Dim dtb As New DataBase(Config.Server)
        dtb.TimeOut = 40
        Dim dt As DataTable
        Try
            dt = dtb.Consultar("Update _TestDate set Date='2009/20/1 0:0:0' where Id=1", False)
            'dt = dtb.Consultar("Update _TestDate set Date='2009/20/1' where Id=1", False)
            If Not dt Is Nothing Then
                If CType(dt.Rows(0).Item(0), Integer) = 1 Then
                    testDateTime = True
                Else
                    testDateTime = False
                End If

                Return True
            Else
                testDateTime = False
                Return True
            End If
            'testDateTime = True

        Catch ex As Exception
            Return False
        End Try
        
    End Function

    Public Function ArmarFecha1(ByVal fecha As DateTime) As String

        'El problema es que tenemos diferentes opciones en el lenguaje del SQL server instalado en cada maquina.
        '¿Se solucionaria cambiando todas las maquinas a ingles?

        ' mam1      
        'Microsoft SQL Server Management Studio Express	 	9.00.2047.00
        'Microsoft Data Access Components (MDAC)			2000.085.1117.00 (xpsp_sp2_rtm.040803-2158)
        'Microsoft MSXML						            2.6 3.0 4.0 5.0 6.0 
        'Microsoft Internet Explorer				        7.0.5730.13
        'Microsoft .NET Framework						    2.0.50727.3053
        'Sistema operativo					                5.1.2600
        'Idioma                                             Ingles

        'andaluza1
        'Microsoft SQL Server Management Studio				9.00.1399.00
        'Microsoft Analysis Services Client Tools			2005.090.1399.00
        'Microsoft Data Access Components (MDAC)			2000.085.1132.00 (xpsp.080413-0852)
        'Microsoft MSXML						            2.6 3.0 4.0 5.0 6.0 
        'Microsoft Internet Explorer						7.0.5730.13
        'Microsoft .NET Framework					    	2.0.50727.3053
        'Operating System						            5.1.2600
        'Idioma                                             Español


        'Dim info As String = System.Globalization.CultureInfo.CurrentCulture.ToString
        'Select Case info.ToString
        '    Case Is = "es-ES"
        '        messageBox.show("España")
        '        Return fecha.Year & "/" & fecha.Day & "/" & fecha.Month & " " & fecha.Hour & ":" & fecha.Minute & ":" & fecha.Second ' & fecha.datetime.now.ToString
        '    Case Is = ""
        '        messageBox.show("")
        '        Return fecha.Year & "/" & fecha.Month & "/" & fecha.Day & " " & fecha.Hour & ":" & fecha.Minute & ":" & fecha.Second
        '    Case Else
        '        messageBox.show("Funciones Generales - ArmarFecha() - Por favor especificar que maquina es esta")
        'End Select

        Dim name As String = My.Computer.Name
        If name = "DEALEJANDRO" Or name = "MAM1" Or name = "CALIDAD1" Or name = "ANDALUZA1" Then
            Return fecha.Year & "/" & fecha.Month & "/" & fecha.Day & " " & fecha.Hour & ":" & fecha.Minute & ":" & fecha.Second
        Else
            'If name = "ANDALUZA1" Then
            '    Return fecha.Year & "/" & fecha.Day & "/" & fecha.Month & " " & fecha.Hour & ":" & fecha.Minute & ":" & fecha.Second ' & fecha.datetime.now.ToString
            '    'Return fecha.Year & "/" & fecha.Month & "/" & fecha.Day & " " & fecha.Hour & ":" & fecha.Minute & ":" & fecha.Second
            'Else
            '    messageBox.show("Funciones Generales - ArmarFecha() - Por favor especificar que maquina es esta")
            '    Return "null"
            'End If
            Return Nothing
        End If
    End Function

    Public Shared Function ArmarFecha(ByVal date1 As DateTime) As String
        If testDateTime <> True Then
            Return date1.Year & "/" & date1.Month & "/" & date1.Day & " " & date1.Hour & ":" & date1.Minute & ":" & date1.Second
        Else
            Return date1.Year & "/" & date1.Day & "/" & date1.Month & " " & date1.Hour & ":" & date1.Minute & ":" & date1.Second
        End If
    End Function

    Public Function ArmarFechaHora(ByVal date1 As DateTime, ByVal fecha As String) As String
        If testDateTime <> True Then
            Return date1.Year & "/" & date1.Month & "/" & date1.Day & " " & fecha
        Else
            Return date1.Year & "/" & date1.Day & "/" & date1.Month & " " & fecha
        End If
    End Function


    '**********************************************
    'Función que devuelve la fecha Juliana
    '**********************************************
    Public Function DevuelveFechaJuliana(ByVal Fecha As DateTime) As String
        Dim aux As DateTime
        Dim c As TimeSpan
        aux = New DateTime(Fecha.Year - 1, 12, 31, 0, 0, 0)
        'c = Fecha - aux
        c = Fecha.Subtract(aux)
        'a la fecha Juliana le añadimos al final las dos últimas cifras del año.
        'Return String.Format("000", c.Days) & Fecha.Year.ToString.Substring(Fecha.Year.ToString.Length - 2, 2)
        Return Format(c.Days, "000") & Fecha.Year.ToString.Substring(Fecha.Year.ToString.Length - 2, 2)
    End Function

    '**************************************************************************************
    'Función que devuelve un valor de hora, partiendo del valor de minutos pasado
    '**************************************************************************************
    Public Function DevuelveHora(ByVal Minutos As Double) As String
        Return (Convert.ToInt32(Minutos) \ 60).ToString & ":" & String.Format((Convert.ToInt32(Minutos) Mod 60).ToString, "00")
    End Function

    ' // This function will not convert dates to 19xx, only 20xx so if you need to consider 19xx dates
    ' || you'll need to add some logic to incorporate a switch point and determine whether to add 1000 or
    ' || 2000 in the calculation
    Public Function JulianToSerial(ByVal strJulian As String) As Date
        Try
            JulianToSerial = New DateTime(2000 + CType(CType(strJulian, Integer) / 1000, Integer), 1, CType(strJulian, Integer) Mod 1000)
        Catch ex As Exception
            messageBox.show("Error. " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function
End Class
