Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Runtime.InteropServices

#Region "Copyright"
' This project is available from SVN on SourceForge.net under the main project, Activelock !
'
' ProjectPage: http://sourceforge.net/projects/activelock
' WebSite: http://www.activeLockSoftware.com
' DeveloperForums: http://forums.activelocksoftware.com
' ProjectManager: Ismail Alkan - http://activelocksoftware.com/simplemachinesforum/index.php?action=profile;u=1
' ProjectLicense: BSD Open License - http://www.opensource.org/licenses/bsd-license.php
' ProjectPurpose: Copy Protection, Software Locking, Anti Piracy
'
' //////////////////////////////////////////////////////////////////////////////////////////
' *   ActiveLock
' *   Copyright 1998-2002 Nelson Ferraz
' *   Copyright 2003-2009 The Activelock - Ismail Alkan (ASG)
' *   All material is the property of the contributing authors.
' *
' *   Redistribution and use in source and binary forms, with or without
' *   modification, are permitted provided that the following conditions are
' *   met:
' *
' *     [o] Redistributions of source code must retain the above copyright
' *         notice, this list of conditions and the following disclaimer.
' *
' *     [o] Redistributions in binary form must reproduce the above
' *         copyright notice, this list of conditions and the following
' *         disclaimer in the documentation and/or other materials provided
' *         with the distribution.
' *
' *   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' *   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' *   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
' *   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
' *   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
' *   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
' *   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
' *   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' *   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' *   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' *   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' *
#End Region

''' <summary>
''' Internet Time Server class by Alastair Dallas 01/27/04
''' </summary>
''' <remarks></remarks>
Public Class Daytime

    ''' <summary>
    ''' Number of seconds that Windows clock can deviate from NIST and still be okay
    ''' </summary>
    ''' <remarks></remarks>
    Private Const THRESHOLD_SECONDS As Integer = 15

    'Server IP addresses from 
    'http://www.boulder.nist.gov/timefreq/service/time-servers.html
    'time-a.nist.gov 129.6.15.28 NIST, Gaithersburg, Maryland 
    'time-b.nist.gov 129.6.15.29  
    'time-a.timefreq.bldrdoc.gov 132.163.4.101 NIST, Boulder, Colorado 
    'time-b.timefreq.bldrdoc.gov 132.163.4.102  
    'time-c.timefreq.bldrdoc.gov 132.163.4.103  
    'utcnist.colorado.edu 128.138.140.44 University of Colorado, Boulder 
    'time.nist.gov 192.43.244.18 NCAR, Boulder, Colorado 
    'time-nw.nist.gov 131.107.13.100 Microsoft, Redmond, Washington 
    'nist1.symmetricom.com 69.25.96.13 Symmetricom, San Jose, California 
    'nist1-dc.WiTime.net 206.246.118.250 WiTime, Virginia 
    'nist1-ny.WiTime.net 208.184.49.9 WiTime, New York City 
    'nist1-sj.WiTime.net 64.125.78.85 WiTime, San Jose, California 
    'nist1.aol-ca.symmetricom.com 207.200.81.113 Symmetricom, AOL facility, Sunnyvale, California 
    'nist1.aol-va.symmetricom.com 64.236.96.53 Symmetricom, AOL facility, Virginia 
    'nist1.columbiacountyga.gov 68.216.79.113 Columbia County, Georgia 
    'nist.expertsmi.com 71.13.91.122 Monroe, Michigan 
    'nist.netservicesgroup.com 64.113.32.5 Southfield, Michigan 

    ' Update this list whenever the server IPs change or new ones are added.

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Private Shared Servers() As String = { _
          "129.6.15.28" _
        , "129.6.15.29" _
        , "132.163.4.101" _
        , "132.163.4.102" _
        , "132.163.4.103" _
        , "128.138.140.44" _
        , "192.43.244.18" _
        , "131.107.13.100" _
        , "69.25.96.13" _
        , "206.246.118.250" _
        , "208.184.49.9" _
        , "64.125.78.85" _
        , "207.200.81.113" _
        , "64.236.96.53" _
        , "68.216.79.113" _
        , "71.13.91.122" _
        , "64.113.32.5" _
    }

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared LastHost As String = ""
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared LastSysTime As DateTime

    ''' <summary>
    ''' Returns UTC/GMT using an NIST server if possible, degrading to simply returning the system clock
    ''' </summary>
    ''' <returns>DateTime - Returns UTC/GMT using an NIST server if possible</returns>
    ''' <remarks></remarks>
    Public Shared Function GetTime() As DateTime

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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="dt1"></param>
    ''' <param name="dt2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function SecondsDifference(ByVal dt1 As DateTime, ByVal dt2 As DateTime) As Integer
        Dim span As TimeSpan = dt1.Subtract(dt2)
        Return span.Seconds + (span.Minutes * 60) + (span.Hours * 3600)
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Shared Function WindowsClockIncorrect() As Boolean
        Dim nist As DateTime = GetTime()
        If (Math.Abs(SecondsDifference(nist, LastSysTime)) > THRESHOLD_SECONDS) Then
            Return True
        End If
        Return False
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="host"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
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

    ''' <summary>
    ''' Retrieves the current system date and time. The system time is expressed in
    ''' Coordinated Universal Time (UTC).
    ''' To retrieve the current system date and time in local time, use the
    ''' GetLocalTime function.
    ''' </summary>
    ''' <param name="stru">A pointer to a SYSTEMTIME structure to receive the current system date and time. The lpSystemTime parameter must not be NULL. Using NULL will result in an access violation.</param>
    ''' <returns>This function does not return a value or provide extended error information.</returns>
    ''' <remarks>To set the current system date and time, use the SetSystemTime function.</remarks>
    Private Declare Function GetSystemTime Lib "kernel32.dll" (ByRef stru As SYSTEMTIME) As Int32
    ''' <summary>
    ''' Sets the current system time and date. The system time is expressed in
    ''' Coordinated Universal Time (UTC).
    ''' </summary>
    ''' <param name="stru">A pointer to a SYSTEMTIME structure that contains the
    ''' new system date and time. The wDayOfWeek member of the SYSTEMTIME structure
    ''' is ignored.</param>
    ''' <returns>If the function succeeds, the return value is nonzero. If the
    ''' function fails, the return value is zero. To get extended error information,
    ''' call GetLastError.</returns>
    ''' <remarks>The calling process must have the SE_SYSTEMTIME_NAME privilege.
    ''' This privilege is disabled by default. The SetSystemTime function enables the
    ''' SE_SYSTEMTIME_NAME privilege before changing the system time and disables the
    ''' privilege before returning. For more information, see Running with Special Privileges.</remarks>
    Private Declare Function SetSystemTime Lib "kernel32.dll" (ByRef stru As SYSTEMTIME) As Int32

    ''' <summary>
    ''' Sets system time.
    ''' </summary>
    ''' <param name="dt"></param>
    ''' <remarks>Note: Use UTC time; Windows will apply time zone</remarks>
    Public Shared Sub SetWindowsClock(ByVal dt As DateTime)

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
