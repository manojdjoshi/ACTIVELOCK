Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Runtime.InteropServices

Public Class Daytime
    'Internet Time Server class by Alastair Dallas 01/27/04

    Private Const THRESHOLD_SECONDS As Integer = 15 'Number of seconds
    ' that Windows clock can deviate from NIST and still be okay

    'Server IP addresses from 
    'http://www.boulder.nist.gov/timefreq/service/time-servers.html
    Private Shared Servers() As String = { _
          "129.6.15.28" _
        , "129.6.15.29" _
        , "132.163.4.101" _
        , "132.163.4.102" _
        , "132.163.4.103" _
        , "128.138.140.44" _
        , "192.43.244.18" _
        , "131.107.1.10" _
        , "66.243.43.21" _
        , "216.200.93.8" _
        , "208.184.49.9" _
        , "207.126.98.204" _
        , "205.188.185.33" _
    }

    Public Shared LastHost As String = ""
    Public Shared LastSysTime As DateTime

    Public Shared Function GetTime() As DateTime
        'Returns UTC/GMT using an NIST server if possible, 
        ' degrading to simply returning the system clock

        'If we are successful in getting NIST time, then
        ' LastHost indicates which server was used and
        ' LastSysTime contains the system time of the call
        ' If LastSysTime is not within 15 seconds of NIST time,
        '  the system clock may need to be reset
        ' If LastHost is "", time is equal to system clock

        Dim host As String
        Dim result As DateTime

        LastHost = ""
        For Each host In Servers
            result = GetNISTTime(host)
            If result > DateTime.MinValue Then
                LastHost = host
                Exit For
            End If
        Next

        If LastHost = "" Then
            'No server in list was successful so use system time
            result = DateTime.UtcNow()
        End If

        Return result
    End Function

    Public Shared Function SecondsDifference(ByVal dt1 As DateTime, ByVal dt2 As DateTime) As Integer
        Dim span As TimeSpan = dt1.Subtract(dt2)
        Return span.Seconds + (span.Minutes * 60) + (span.Hours * 3600)
    End Function

    Public Shared Function WindowsClockIncorrect() As Boolean
        Dim nist As DateTime = GetTime()
        If (Math.Abs(SecondsDifference(nist, LastSysTime)) > THRESHOLD_SECONDS) Then
            Return True
        End If
        Return False
    End Function

    Private Shared Function GetNISTTime(ByVal host As String) As DateTime
        'Returns DateTime.MinValue if host unreachable or does not produce time
        Dim timeStr As String

        Try
            Dim reader As New StreamReader(New TcpClient(host, 13).GetStream)
            LastSysTime = DateTime.UtcNow()
            timeStr = reader.ReadToEnd()
            reader.Close()
        Catch ex As SocketException
            'Couldn't connect to server, transmission error
            Debug.WriteLine("Socket Exception [" & host & "]")
            Return DateTime.MinValue
        Catch ex As Exception
            'Some other error, such as Stream under/overflow
            Return DateTime.MinValue
        End Try

        'Parse timeStr
        If (timeStr.Substring(38, 9) <> "UTC(NIST)") Then
            'This signature should be there
            Return DateTime.MinValue
        End If
        If (timeStr.Substring(30, 1) <> "0") Then
            'Server reports non-optimum status, time off by as much as 5 seconds
            Return DateTime.MinValue    'Try a different server
        End If

        Dim jd As Integer = Integer.Parse(timeStr.Substring(1, 5))
        Dim yr As Integer = Integer.Parse(timeStr.Substring(7, 2))
        Dim mo As Integer = Integer.Parse(timeStr.Substring(10, 2))
        Dim dy As Integer = Integer.Parse(timeStr.Substring(13, 2))
        Dim hr As Integer = Integer.Parse(timeStr.Substring(16, 2))
        Dim mm As Integer = Integer.Parse(timeStr.Substring(19, 2))
        Dim sc As Integer = Integer.Parse(timeStr.Substring(22, 2))

        If (jd < 15020) Then
            'Date is before 1900
            Return DateTime.MinValue
        End If
        If (jd > 51544) Then yr += 2000 Else yr += 1900

        Return New DateTime(yr, mo, dy, hr, mm, sc)

    End Function

    <StructLayout(LayoutKind.Sequential)> _
        Public Structure SYSTEMTIME
        Public wYear As Int16
        Public wMonth As Int16
        Public wDayOfWeek As Int16
        Public wDay As Int16
        Public wHour As Int16
        Public wMinute As Int16
        Public wSecond As Int16
        Public wMilliseconds As Int16
    End Structure

    Private Declare Function GetSystemTime Lib "kernel32.dll" (ByRef stru As SYSTEMTIME) As Int32
    Private Declare Function SetSystemTime Lib "kernel32.dll" (ByRef stru As SYSTEMTIME) As Int32

    Public Shared Sub SetWindowsClock(ByVal dt As DateTime)
        'Sets system time. Note: Use UTC time; Windows will apply time zone

        Dim timeStru As SYSTEMTIME
        Dim result As Int32

        timeStru.wYear = CType(dt.Year, Int16)
        timeStru.wMonth = CType(dt.Month, Int16)
        timeStru.wDay = CType(dt.Day, Int16)
        timeStru.wDayOfWeek = CType(dt.DayOfWeek, Int16)
        timeStru.wHour = CType(dt.Hour, Int16)
        timeStru.wMinute = CType(dt.Minute, Int16)
        timeStru.wSecond = CType(dt.Second, Int16)
        timeStru.wMilliseconds = CType(dt.Millisecond, Int16)

        result = SetSystemTime(timeStru)

    End Sub
End Class
