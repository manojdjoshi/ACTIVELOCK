VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Regional Date Setting"
   ClientHeight    =   5625
   ClientLeft      =   5505
   ClientTop       =   1395
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   7350
   Begin VB.CommandButton Command2 
      Caption         =   "Now"
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   2040
      Width           =   1995
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get Date Setting"
      Height          =   555
      Left            =   1740
      TabIndex        =   1
      Top             =   60
      Width           =   3675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   120
      TabIndex        =   4
      Top             =   2460
      Width           =   120
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   3540
      TabIndex        =   0
      Top             =   660
      Width           =   135
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' The following code is used by the locale date format settings
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String) As Boolean
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
Const LOCALE_SSHORTDATE = &H1F
Private regionalSymbol As String

Private Type SYSTEMTIME
     wYear As Integer
     wMonth As Integer
     wDayOfWeek As Integer
     wDay As Integer
     wHour As Integer
     wMinute As Integer
     wSecond As Integer
     wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    bias As Long ' current offset to GMT
    StandardName(1 To 64) As Byte ' unicode string
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(1 To 64) As Byte
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Public Enum TimeZoneReturn
    TimeZoneCode = 0
    TimeZoneName = 1
    UTC_BaseOffset = 2
    UTC_Offset = 3
    DST_Active = 4
    DST_Offset = 5
End Enum

' ----------------- For Time Zone Retrieval ------------------
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Private Declare Sub GetSystemTime Lib "kernel32" _
    (lpSystemTime As SYSTEMTIME)

Private Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long


Function ActivelockDate(m As Integer, d As Integer, y As Integer) As String
ActivelockDate = Format$(y, "0000") & "/" & Format$(m, "00") & "/" & Format$(d, "00")
End Function

Function ConvertToActivelockDate(ByVal myDate As Variant) As String
Dim newDate As Date
Dim m As Long
Dim d As Long
Dim y As Long
If IsDate(myDate) Then
    newDate = CDate(myDate)
    newDate = UTC(newDate)
    m = Month(newDate)
    d = Day(newDate)
    y = Year(newDate)
    ConvertToActivelockDate = Format$(y, "0000") & "/" & Format$(m, "00") & "/" & Format$(d, "00")
Else
    ConvertToActivelockDate = "0000/00/00"
End If
End Function
Function padRight(s As String, n As Integer) As String
padRight = Left$(s & String(n, " "), n)
End Function

Function RegionalDate(m As Integer, d As Integer, y As Integer) As String
'converts month, day, and year into whatever the current regional format is
Dim s As String, t As String, symbol As String
Dim i As Integer, startPoint As Integer, endPoint As Integer
s = LCase(regionalSymbol) & Space(5)

startPoint = 1
Do
    If Mid$(s, startPoint, 4) = "yyyy" Then
        t = t & Format$(y, "0000")
        startPoint = startPoint + 4
    ElseIf Mid$(s, startPoint, 3) = "yyy" Then
        t = t & Format$(y, "000")
        startPoint = startPoint + 3
    ElseIf Mid$(s, startPoint, 2) = "yy" Then
        t = t & Format$(y, "00")
        startPoint = startPoint + 2
    ElseIf Mid$(s, startPoint, 1) = "y" Then
        t = t & Format$(y, "0")
        startPoint = startPoint + 1
    ElseIf Mid$(s, startPoint, 2) = "mm" Then
        t = t & Format$(m, "00")
        startPoint = startPoint + 2
    ElseIf Mid$(s, startPoint, 1) = "m" Then
        t = t & Format$(m, "0")
        startPoint = startPoint + 1
    ElseIf Mid$(s, startPoint, 2) = "dd" Then
        t = t & Format$(d, "00")
        startPoint = startPoint + 2
    ElseIf Mid$(s, startPoint, 1) = "d" Then
        t = t & Format$(d, "0")
        startPoint = startPoint + 1
    Else
        t = t & Mid$(s, startPoint, 1)
        startPoint = startPoint + 1
    End If
Loop Until startPoint >= Len(s)
RegionalDate = Trim(t)
End Function

Public Function UTC(dt As Date) As Date
    '  Returns current UTC date-time.
    UTC = DateAdd("n", LocalTimeZone(UTC_Offset), dt)
End Function

Public Function LocalTimeZone(ByVal returnType As TimeZoneReturn) As Variant
    Dim X As Long
    Dim tzi As TIME_ZONE_INFORMATION
    Dim strName As String
    Dim bDST As Boolean
    Dim rc&
    rc = GetTimeZoneInformation(tzi)
    Select Case rc
        ' if not daylight assume standard
        Case TIME_ZONE_ID_DAYLIGHT
            strName = tzi.DaylightName ' convert to string
            bDST = True
        Case Else
            strName = tzi.StandardName
    End Select
    
    ' name terminates with null
    X = InStr(strName, vbNullChar)
    If X > 0 Then strName = Left$(strName, X - 1)
            
    If returnType = DST_Active Then
        LocalTimeZone = bDST
    End If
    
    If returnType = TimeZoneName Then
        LocalTimeZone = strName
    End If
    
    If returnType = TimeZoneCode Then
        LocalTimeZone = Left(strName, 1)
        X = InStr(1, strName, " ")
        Do While X > 0
            LocalTimeZone = LocalTimeZone & Mid(strName, X + 1, 1)
            X = InStr(X + 1, strName, " ")
        Loop
        LocalTimeZone = Trim(LocalTimeZone)
    End If
            
    If returnType = UTC_BaseOffset Then
        LocalTimeZone = tzi.bias
    End If
        
    If returnType = DST_Offset Then
        LocalTimeZone = tzi.DaylightBias
    End If
    
    If returnType = UTC_Offset Then
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
    Dim iRet1 As Long
    Dim iRet2 As Long
    Dim lpLCDataVar As String
    Dim Pos As Integer
    Dim Locale As Long
    Locale = GetUserDefaultLCID()
    iRet1 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, lpLCDataVar, 0)
    symbol = String$(iRet1, 0)
    iRet2 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, symbol, iRet1)
    Pos = InStr(symbol, Chr$(0))
    If Pos > 0 Then
         symbol = Left$(symbol, Pos - 1)
         If symbol <> "yyyy/MM/dd" Then regionalSymbol = symbol
    End If
End Sub
Public Sub Set_locale(Optional ByVal localSymbol As String = "") 'Change the regional setting
    Dim symbol As String
    Dim iRet As Long
    Dim Locale As Long
    Locale = GetUserDefaultLCID() 'Get user Locale ID
    If localSymbol = "" Then
      symbol = "yyyy/MM/dd" 'New character for the locale
    Else
      symbol = localSymbol
    End If
    
    iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, symbol)
End Sub

Private Sub Command1_Click()
Get_locale
Label1.Caption = "Under your current settings" & vbCrLf & "dates are formatted as" & vbCrLf _
& regionalSymbol & vbCrLf & vbCrLf & "Enter a date in this format" & vbCrLf & vbCrLf & "                            or press"
End Sub


Private Sub Command2_Click()
'Text1.Text = Format$(Now, regionalSymbol)
Text1.Text = ConvertToActivelockDate(Now)
End Sub


Private Sub Form_Activate()
'Command1_Click
'Command2_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub Text1_Change()
'Dim newDate As String, s As String, t As String
'Dim m As Integer, d As Integer, y As Integer
's = Text1.Text
'If IsDate(s) Then
'    newDate = s
'    m = Month(newDate)
'    d = Day(newDate)
'    y = Year(newDate)
'    t = "Month: " & m & vbCrLf & "Day: " & d & vbCrLf & "Year: " & y & vbCrLf & _
'    "UTC is " & UTC(RegionalDate(m, d, y)) & vbCrLf & vbCrLf & _
'    Space(31) & "Activelock" & Space(8) & "REGIONAL" & vbCrLf & _
'    padRight("This date is", 30) & ActivelockDate(m, d, y) & Space(8) & RegionalDate(m, d, y)
'
'    newDate = DateAdd("d", 7, s)
'    m = Month(newDate)
'    d = Day(newDate)
'    y = Year(newDate)
'    t = t & vbCrLf & padRight("A week from this date is", 30) & ActivelockDate(m, d, y) & Space(8) & RegionalDate(m, d, y)
'
'    newDate = DateAdd("m", 1, s)
'    m = Month(newDate)
'    d = Day(newDate)
'    y = Year(newDate)
'    t = t & vbCrLf & padRight("A month from this date is", 30) & ActivelockDate(m, d, y) & Space(8) & RegionalDate(m, d, y)
'
'    newDate = DateAdd("yyyy", 1, s)
'    m = Month(newDate)
'    d = Day(newDate)
'    y = Year(newDate)
'    t = t & vbCrLf & padRight("A year from this date is", 30) & ActivelockDate(m, d, y) & Space(8) & RegionalDate(m, d, y)
'
'    newDate = DateAdd("yyyy", 10, s)
'    m = Month(newDate)
'    d = Day(newDate)
'    y = Year(newDate)
'    t = t & vbCrLf & padRight("A decade from this date is", 30) & ActivelockDate(m, d, y) & Space(8) & RegionalDate(m, d, y)
'
'    t = t & vbCrLf & vbCrLf & "That decade contains " & DateDiff("d", s, RegionalDate(m, d, y)) & " days"
'
'
'
'    Label2.Caption = t
'Else
'    Label2.Caption = "This is not a date"
'End If
End Sub

