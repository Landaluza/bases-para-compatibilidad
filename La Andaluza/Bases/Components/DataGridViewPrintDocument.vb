Public Class DataGridViewPrintDocument    
    Inherits Printing.PrintDocument
    Private printFont As Font
    Private DataFont As Font
    Private Grilla As DataGridView
    Private FilaActual As Integer = 0
    Private anchos As New System.Collections.Hashtable
    Private Registros As Integer
    Private Columnas As Integer
    Private HeaderFont As Font
    Private Encabezado As String()
    Private nPagina As Integer = 1
    Private HoraReporte As String
    Private FechaReporte As String

    Public Sub New(ByVal Grilla As DataGridView, ByVal Header As String())
        MyBase.New()
        'asignar y calcular los anchos de columnas
        Me.Grilla = Grilla
        Me.Registros = Me.Grilla.Rows.Count
        Me.Columnas = Me.Grilla.Columns.Count

        Me.HoraReporte = String.Format("HH:mm:ss", Now.Date)
        Me.FechaReporte = CStr(Today.Day) & "/" & Today.Month & "/" & CStr(Today.Year)
        Me.Encabezado = Header
        For i As Integer = 0 To Me.Columnas - 1
            If Grilla.Columns(i).Visible Then
                anchos.Add(Grilla.Columns(i).Name,
                Math.Round(Grilla.Columns(i).Width / 6, 0)) 'de pixels a espacios
            End If
        Next
    End Sub

    Protected Overrides Sub OnPrintPage(ByVal ev As Printing.PrintPageEventArgs)
        MyBase.OnPrintPage(ev)
        Dim lpp As Double = 0
        Dim yPos As Integer = 0
        Dim count As Integer = 0
        Dim leftMargin As Integer = ev.MarginBounds.Left
        Dim topMargin As Integer = ev.MarginBounds.Top
        Dim line As String = ""
        Dim valor As String = ""

        FilaActual = 0

        'imprimir el numero de pagina.
        yPos = topMargin + (count * Me.DataFont.GetHeight(ev.Graphics))
        ev.Graphics.DrawString("Página:" & nPagina.ToString, Me.DataFont, Brushes.Black, leftMargin, yPos, New StringFormat())
        count += 1
        'imprimir fecha de impresion
        yPos = topMargin + (count * Me.DataFont.GetHeight(ev.Graphics))
        ev.Graphics.DrawString("Fecha:" & Me.FechaReporte, Me.DataFont,Brushes.Black, leftMargin, yPos, New StringFormat())
        count += 1
        'imprimir hora de impresión
        yPos = topMargin + (count * Me.DataFont.GetHeight(ev.Graphics))
        ev.Graphics.DrawString("Hora:" & Me.HoraReporte, Me.DataFont,
        Brushes.Black, leftMargin, yPos, New StringFormat())
        count += 1

        'imprimir el encabezado
        For i As Integer = 0 To Me.Encabezado.Length - 1
            yPos = topMargin + (count *
            Me.HeaderFont.GetHeight(ev.Graphics))
            ev.Graphics.DrawString(LTrim(Me.Encabezado(i)) &
            ControlChars.NewLine, Me.HeaderFont, Brushes.Black, leftMargin, yPos,
            New StringFormat())
            count += 1
        Next

        'linea separadora
        count += 1
        yPos = topMargin + (count *
        Me.HeaderFont.GetHeight(ev.Graphics))
        ev.Graphics.DrawLine(Pens.Black, leftMargin, yPos,
        ev.PageBounds.Width - ev.MarginBounds.Left, yPos)

        'imprimir las cabeceras de las columnas
        For col As Integer = 0 To Me.Columnas - 1
            If Grilla.Columns(col).Visible Then
                valor = if(IsDBNull(Grilla.Columns(col).Name), "", Grilla.Columns(col).Name)
                valor = valor.PadRight(anchos.Item(Grilla.Columns(col).Name), " ")
                line &= valor & " "
            End If
        Next
        yPos = topMargin + (count * HeaderFont.GetHeight(ev.Graphics))
        ev.Graphics.DrawString(line, printFont, Brushes.Black,
        leftMargin, yPos, New StringFormat())

        'linea separadora
        count += 1
        yPos = topMargin + (count *
        Me.HeaderFont.GetHeight(ev.Graphics))
        ev.Graphics.DrawLine(Pens.Black, leftMargin, yPos,
        ev.PageBounds.Width - ev.MarginBounds.Left, yPos)
        count += 1

        'proceso de impresion de lineas para cada pagina
        count += 4
        lpp = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics) - 5 'numero lineas por pagina
        While (count < lpp)
            'producir la fila a imprimir
            line = ""
            For col As Integer = 0 To Me.Columnas - 1
                If Grilla.Columns(col).Visible Then
                    valor = if(IsDBNull(Grilla.Rows(Me.FilaActual).Cells(col).Value), "", Grilla.Rows(Me.FilaActual).Cells(col).Value)
                    valor = valor.PadRight(anchos.Item(Grilla.Columns(col).Name), " ")
                    line &= valor & " "
                End If
            Next
            'calcular posicion de la linea a imprimir
            line &= ControlChars.NewLine
            yPos = topMargin + (count *
            printFont.GetHeight(ev.Graphics))
            ev.Graphics.DrawString(line, printFont, Brushes.Black,
            leftMargin, yPos, New StringFormat())
            count += 1
            Me.FilaActual += 1
            If Me.FilaActual > Me.Registros - 1 Then Exit While
        End While
        'si existen mas lineas, imprimir paginas que faltan
        If Me.FilaActual > Me.Registros - 1 Then
            ev.HasMorePages = False
        Else
            ev.HasMorePages = True
            Me.nPagina += 1
        End If
    End Sub
End Class
