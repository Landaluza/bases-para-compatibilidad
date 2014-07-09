Public Class FechaQs
    Public Function mam_FechaQS(ByRef Fecha As Date) As String
        Dim m_FechaQs As String = String.Empty
        m_FechaQs = Fecha.ToShortDateString.Replace("/", "").Remove(4, 2)
        Return m_FechaQs
    End Function


    Public Function mam_FechaSQL(ByRef Fecha As Date) As String
        Dim m_FechaSQL As String = String.Empty
        m_FechaSQL = Fecha.Year & "-" & Fecha.Month & "-" & Fecha.Day
        Return m_FechaSQL
    End Function
End Class
