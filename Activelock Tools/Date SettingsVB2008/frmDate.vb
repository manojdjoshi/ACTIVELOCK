Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Imports System.DateTime
Friend Class Form1
	Inherits System.Windows.Forms.Form
	' The following code is used by the locale date format settings
	Private Declare Function GetLocaleInfo Lib "kernel32"  Alias "GetLocaleInfoA"(ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
	Private Declare Function SetLocaleInfo Lib "kernel32"  Alias "SetLocaleInfoA"(ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String) As Boolean
	Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Short
    Const LOCALE_SSHORTDATE As Short = &H1F
	Private regionalSymbol As String
	
    Structure SYSTEMTIME
        Dim wYear As Short
        Dim wMonth As Short
        Dim wDayOfWeek As Short
        Dim wDay As Short
        Dim wHour As Short
        Dim wMinute As Short
        Dim wSecond As Short
        Dim wMilliseconds As Short
    End Structure

    Private Structure TIME_ZONE_INFORMATION
        Dim Bias As Integer
        <VBFixedString(64), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=64)> Public StandardName As String
        Dim StandardDate As SYSTEMTIME
        Dim StandardBias As Integer
        <VBFixedString(64), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=64)> Public DaylightName As String
        Dim DaylightDate As SYSTEMTIME
        Dim DaylightBias As Integer
    End Structure

	Public Enum TimeZoneReturn
		TimeZoneCode = 0
		TimeZoneName = 1
		UTC_BaseOffset = 2
		UTC_Offset = 3
		DST_Active = 4
		DST_Offset = 5
	End Enum
	
	' ----------------- For Time Zone Retrieval ------------------
	Private Const TIME_ZONE_ID_UNKNOWN As Short = 0
	Private Const TIME_ZONE_ID_STANDARD As Short = 1
	Private Const TIME_ZONE_ID_INVALID As Integer = &HFFFFFFFF
	Private Const TIME_ZONE_ID_DAYLIGHT As Short = 2
	
    Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Integer

	Function AmericanDate(ByRef m As Short, ByRef d As Short, ByRef y As Short) As String
		AmericanDate = VB6.Format(y, "0000") & "/" & VB6.Format(m, "00") & "/" & VB6.Format(d, "00")
	End Function
	Function padRight(ByRef s As String, ByRef n As Short) As String
		padRight = VB.Left(s & New String(" ", n), n)
	End Function
	
	Function RegionalDate(ByRef m As Short, ByRef d As Short, ByRef y As Short) As String
		'converts month, day, and year into whatever the current regional format is
        Dim t As String = ""
        Dim s As String = ""
        Dim startPoint As Short
		s = LCase(regionalSymbol) & Space(5)
		
		startPoint = 1
		Do 
			If Mid(s, startPoint, 4) = "yyyy" Then
				t = t & VB6.Format(y, "0000")
				startPoint = startPoint + 4
			ElseIf Mid(s, startPoint, 3) = "yyy" Then 
				t = t & VB6.Format(y, "000")
				startPoint = startPoint + 3
			ElseIf Mid(s, startPoint, 2) = "yy" Then 
				t = t & VB6.Format(y, "00")
				startPoint = startPoint + 2
			ElseIf Mid(s, startPoint, 1) = "y" Then 
				t = t & VB6.Format(y, "0")
				startPoint = startPoint + 1
			ElseIf Mid(s, startPoint, 2) = "mm" Then 
				t = t & VB6.Format(m, "00")
				startPoint = startPoint + 2
			ElseIf Mid(s, startPoint, 1) = "m" Then 
				t = t & VB6.Format(m, "0")
				startPoint = startPoint + 1
			ElseIf Mid(s, startPoint, 2) = "dd" Then 
				t = t & VB6.Format(d, "00")
				startPoint = startPoint + 2
			ElseIf Mid(s, startPoint, 1) = "d" Then 
				t = t & VB6.Format(d, "0")
				startPoint = startPoint + 1
			Else
				t = t & Mid(s, startPoint, 1)
				startPoint = startPoint + 1
			End If
		Loop Until startPoint >= Len(s)
		RegionalDate = Trim(t)
	End Function
	
    Public Function UTC(ByRef dt As Date) As Date
        '  Returns current UTC date-time.
        UTC = DateAdd(Microsoft.VisualBasic.DateInterval.Minute, LocalTimeZone(TimeZoneReturn.UTC_Offset), dt)
    End Function

    Public Function LocalTimeZone(ByVal returnType As TimeZoneReturn) As Object
        Dim x As Integer
        Dim tzi As New TIME_ZONE_INFORMATION
        Dim strName As String
        Dim bDST As Boolean
        Dim rc As Integer

        LocalTimeZone = Nothing

        rc = GetTimeZoneInformation(tzi)
        Select Case rc
            ' if not daylight assume standard
            Case TIME_ZONE_ID_DAYLIGHT
                strName = tzi.DaylightName  'System.Text.UnicodeEncoding.Unicode.GetString(tzi.DaylightName) ' convert to string
                bDST = True
            Case Else
                strName = tzi.StandardName  'System.Text.UnicodeEncoding.Unicode.GetString(tzi.StandardName)
        End Select

        ' name terminates with null
        x = InStr(strName, vbNullChar)
        If x > 0 Then strName = VB.Left(strName, x - 1)

        If returnType = TimeZoneReturn.DST_Active Then
            LocalTimeZone = bDST
        End If

        If returnType = TimeZoneReturn.TimeZoneName Then
            LocalTimeZone = strName
        End If

        If returnType = TimeZoneReturn.TimeZoneCode Then
            LocalTimeZone = VB.Left(strName, 1)
            x = InStr(1, strName, " ")
            Do While x > 0
                LocalTimeZone = LocalTimeZone & Mid(strName, x + 1, 1)
                x = InStr(x + 1, strName, " ")
            Loop
            LocalTimeZone = Trim(LocalTimeZone)
        End If

        If returnType = TimeZoneReturn.UTC_BaseOffset Then
            LocalTimeZone = tzi.bias
        End If

        If returnType = TimeZoneReturn.DST_Offset Then
            LocalTimeZone = tzi.DaylightBias
        End If

        If returnType = TimeZoneReturn.UTC_Offset Then
            If tzi.DaylightBias = -60 Then
                LocalTimeZone = tzi.bias
            Else
                LocalTimeZone = -tzi.bias
            End If
            ' Account for Daylight Savings Time
            If bDST Then LocalTimeZone = LocalTimeZone - 60
        End If
    End Function

	Public Sub Get_locale() ' Retrieve the regional setting
		Dim symbol As String
		Dim iRet1 As Integer
		Dim iRet2 As Integer
        Dim lpLCDataVar As String = ""
		Dim Pos As Short
		Dim Locale As Integer
		Locale = GetUserDefaultLCID()
		iRet1 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, lpLCDataVar, 0)
		symbol = New String(Chr(0), iRet1)
		iRet2 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, symbol, iRet1)
		Pos = InStr(symbol, Chr(0))
		If Pos > 0 Then
			symbol = VB.Left(symbol, Pos - 1)
			If symbol <> "yyyy/MM/dd" Then regionalSymbol = symbol
		End If
	End Sub
	Public Sub Set_locale(Optional ByVal localSymbol As String = "") 'Change the regional setting
		Dim symbol As String
		Dim iRet As Integer
		Dim Locale As Integer
		Locale = GetUserDefaultLCID() 'Get user Locale ID
		If localSymbol = "" Then
			symbol = "yyyy/MM/dd" 'New character for the locale
		Else
			symbol = localSymbol
		End If
		
		iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, symbol)
	End Sub
	
	Private Sub Command1_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command1.Click
		Get_locale()
		Label1.Text = "Under your current settings" & vbCrLf & "dates are formatted as" & vbCrLf & regionalSymbol & vbCrLf & vbCrLf & "Enter a date in this format" & vbCrLf & vbCrLf & "                            or press"
	End Sub
	
	
	Private Sub Command2_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Command2.Click
        Text1.Text = ConvertToActivelockDate(Now)
	End Sub
    Private Sub Form1_Activated(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Activated
        'Command1_Click(Command1, New System.EventArgs())
        'Command2_Click(Command2, New System.EventArgs())
    End Sub
	
	Private Sub Form1_FormClosed(ByVal eventSender As System.Object, ByVal eventArgs As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
		End
	End Sub
	
	
	'UPGRADE_WARNING: Event Text1.TextChanged may fire when form is initialized. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
	Private Sub Text1_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles Text1.TextChanged
        'Dim s, newDate, t As String
        'Dim d, m, y As Short
        's = Text1.Text
        'If IsDate(s) Then
        '	newDate = s
        '          m = VB.Month(CDate(newDate))
        '	d = VB.Day(CDate(newDate))
        '          y = VB.Year(CDate(newDate))
        '	t = "Month: " & m & vbCrLf & "Day: " & d & vbCrLf & "Year: " & y & vbCrLf & "UTC is " & UTC(CDate(RegionalDate(m, d, y))) & vbCrLf & vbCrLf & Space(32) & "AMERICAN" & Space(10) & "REGIONAL" & vbCrLf & padRight("This date is", 30) & AmericanDate(m, d, y) & Space(8) & RegionalDate(m, d, y)

        '	newDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Day, 7, CDate(s)))
        '          m = VB.Month(CDate(newDate))
        '	d = VB.Day(CDate(newDate))
        '          y = VB.Year(CDate(newDate))
        '	t = t & vbCrLf & padRight("A week from this date is", 30) & AmericanDate(m, d, y) & Space(8) & RegionalDate(m, d, y)

        '	newDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Month, 1, CDate(s)))
        '          m = VB.Month(CDate(newDate))
        '	d = VB.Day(CDate(newDate))
        '          y = VB.Year(CDate(newDate))
        '	t = t & vbCrLf & padRight("A month from this date is", 30) & AmericanDate(m, d, y) & Space(8) & RegionalDate(m, d, y)

        '	newDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, 1, CDate(s)))
        '          m = VB.Month(CDate(newDate))
        '	d = VB.Day(CDate(newDate))
        '          y = VB.Year(CDate(newDate))
        '	t = t & vbCrLf & padRight("A year from this date is", 30) & AmericanDate(m, d, y) & Space(8) & RegionalDate(m, d, y)

        '	newDate = CStr(DateAdd(Microsoft.VisualBasic.DateInterval.Year, 10, CDate(s)))
        '          m = VB.Month(CDate(newDate))
        '	d = VB.Day(CDate(newDate))
        '          y = VB.Year(CDate(newDate))
        '	t = t & vbCrLf & padRight("A decade from this date is", 30) & AmericanDate(m, d, y) & Space(8) & RegionalDate(m, d, y)

        '	'UPGRADE_WARNING: DateDiff behavior may be different. Click for more: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6B38EC3F-686D-4B2E-B5A5-9E8E7A762E32"'
        '	t = t & vbCrLf & vbCrLf & "That decade contains " & DateDiff(Microsoft.VisualBasic.DateInterval.Day, CDate(s), CDate(RegionalDate(m, d, y))) & " days"



        '	Label2.Text = t
        'Else
        '	Label2.Text = "This is not a date"
        'End If
    End Sub
    Function ConvertToActivelockDate(ByVal myDate As Object) As String
        Dim newDate As DateTime
        Dim m As Integer
        Dim d As Integer
        Dim y As Integer
        If IsDate(myDate) Then
            newDate = CDate(myDate)
            newDate = UTC(newDate)
            m = newDate.Month
            d = newDate.Day
            y = newDate.Year
            ConvertToActivelockDate = Format$(y, "0000") & "/" & Format$(m, "00") & "/" & Format$(d, "00")
        Else
            ConvertToActivelockDate = "0000/00/00"
        End If
    End Function

End Class