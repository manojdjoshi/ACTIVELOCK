VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Locale"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Locale Class
' obtain the Regional Settings for the System.
' ensure that all displays etc are correct for the country of use
'
' Locale Constants
Private Const LOCALE_SDECIMAL As Long = &HE
Private Const LOCALE_SLONGDATE As Long = &H20
Private Const LOCALE_SSHORTDATE As Long = &H1F
Private Const LOCALE_SCURRENCY As Long = &H14
Private Const LOCALE_STHOUSAND As Long = &HF
Private Const LOCALE_SINTLSYMBOL As Long = &H15
Private Const LOCALE_STIMEFORMAT As Long = &H1003
Private Const LOCALE_SNEGATIVESIGN = &H51
Private Const LOCALE_SPOSITIVESIGN = &H50
Private Const LOCALE_SCOUNTRY As Long = &H6
Private Const LOCALE_SDAYNAME1 As Long = &H2A
Private Const LOCALE_SDAYNAME2 As Long = &H2B
Private Const LOCALE_SDAYNAME3 As Long = &H2C
Private Const LOCALE_SDAYNAME4 As Long = &H2D
Private Const LOCALE_SDAYNAME5 As Long = &H2E
Private Const LOCALE_SDAYNAME6 As Long = &H2F
Private Const LOCALE_SDAYNAME7 As Long = &H30
Private Const LOCALE_SENGCOUNTRY As Long = &H1002
Private Const LOCALE_SENGLANGUAGE As Long = &H1001
Private Const LOCALE_SLANGUAGE As Long = &H2
Private Const LOCALE_SMONTHNAME1 As Long = &H38
Private Const LOCALE_SMONTHNAME10 As Long = &H41
Private Const LOCALE_SMONTHNAME11 As Long = &H42
Private Const LOCALE_SMONTHNAME12 As Long = &H43
Private Const LOCALE_SMONTHNAME2 As Long = &H39
Private Const LOCALE_SMONTHNAME3 As Long = &H3A
Private Const LOCALE_SMONTHNAME4 As Long = &H3B
Private Const LOCALE_SMONTHNAME5 As Long = &H3C
Private Const LOCALE_SMONTHNAME6 As Long = &H3D
Private Const LOCALE_SMONTHNAME7 As Long = &H3E
Private Const LOCALE_SMONTHNAME8 As Long = &H3F
Private Const LOCALE_SMONTHNAME9 As Long = &H40
Private Const LOCALE_SABBREVCTRYNAME = &H7
Private Const LOCALE_SABBREVDAYNAME1 = &H31
Private Const LOCALE_SABBREVDAYNAME3 = &H33
Private Const LOCALE_SABBREVDAYNAME2 = &H32
Private Const LOCALE_SABBREVDAYNAME4 = &H34
Private Const LOCALE_SABBREVDAYNAME5 = &H35
Private Const LOCALE_SABBREVDAYNAME6 = &H36
Private Const LOCALE_SABBREVDAYNAME7 = &H37
Private Const LOCALE_SABBREVLANGNAME = &H3
Private Const LOCALE_SABBREVMONTHNAME1 = &H44
Private Const LOCALE_SABBREVMONTHNAME10 = &H4D
Private Const LOCALE_SABBREVMONTHNAME11 = &H4E
Private Const LOCALE_SABBREVMONTHNAME12 = &H4F
Private Const LOCALE_SABBREVMONTHNAME13 = &H100F
Private Const LOCALE_SABBREVMONTHNAME2 = &H45
Private Const LOCALE_SABBREVMONTHNAME3 = &H46
Private Const LOCALE_SABBREVMONTHNAME4 = &H47
Private Const LOCALE_SABBREVMONTHNAME5 = &H48
Private Const LOCALE_SABBREVMONTHNAME6 = &H49
Private Const LOCALE_SABBREVMONTHNAME7 = &H4A
Private Const LOCALE_SABBREVMONTHNAME8 = &H4B
Private Const LOCALE_SABBREVMONTHNAME9 = &H4C
Private Const LOCALE_SNATIVECTRYNAME = &H8
Private Const LOCALE_SNATIVELANGNAME = &H4
Private Const LOCALE_USER_DEFAULT As Long = &H400

' API Declarations for the Locale Methods
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal lLocale As Long, ByVal lLocaleType As Long, ByVal sLCData As String, ByVal lBufferLength As Long) As Long
Private Declare Function GetSystemDefaultLangID Lib "kernel32" () As Integer
Private Declare Function VerLanguageName Lib "kernel32" Alias "VerLanguageNameA" (ByVal wLang As Long, ByVal szLang As String, ByVal nSize As Long) As Long
'------------------
' Public Properties
'------------------
Public Property Get DateFormat() As String
    ' This function will return the Locale date format for the system. Note that the
    ' returned Year is always formatted to 'YYYY' regardless, to ensure Y2k compliance.
    On Error GoTo vbErrorHandler
    DateFormat = GetLocaleString(LOCALE_SSHORTDATE)
    ' Make sure we always have YYYY format for y2k
    If InStr(1, DateFormat, "YYYY", vbTextCompare) = 0 Then DateFormat = Replace(DateFormat, "YY", "YYYY")
Exit Property
vbErrorHandler:
    Err.Raise Err.Number, "CGLocaleInfo GetDateFormat", Err.Description
End Property

Public Property Get TimeFormat() As String
    'This function returns the locale's defined Time Format.
    TimeFormat = Left(GetLocaleString(LOCALE_STIMEFORMAT), 8)
Exit Property
vbErrorHandler:
    Err.Raise Err.Number, "CGLocaleInfo GetTimeFormat", Err.Description
End Property

Public Property Get NumberFormat() As String
' This function returns the Locales defined Decimal Number format
    On Error GoTo vbErrorHandler
    NumberFormat = "##" & ThousandSpecifier & "##0" & DecimalSpecifier & "00"
Exit Property
vbErrorHandler:
    NumberFormat = "##,##,0.00" 'Set a Default
End Property

Public Property Get ThousandSpecifier() As String
    'This function returns the correct Thousand Specifier for the system Locale
    ThousandSpecifier = GetLocaleString(LOCALE_STHOUSAND)
End Property

Public Property Get DecimalSpecifier() As String
    'This function returns the correct Decimal Specifier for the system Locale
    DecimalSpecifier = GetLocaleString(LOCALE_SDECIMAL)
    
End Property

Public Property Get CurrencySpecifier() As String
    'This function returns the correct Currency Specifier for the system Locale
    CurrencySpecifier = GetLocaleString(LOCALE_SCURRENCY)
End Property

Public Property Get SysLanguageID() As Long
    'Returns the System Language ID for the machine
    SysLanguageID = GetSystemDefaultLangID
End Property

Public Property Get SysLanguageName() As String
    'Returns the System Language Name eg : English (United Kingdom)
    Dim lLangID As Long
    Dim sBuffer As String
    Dim lBuffSize As Long
    Dim lRet As Long

    On Error GoTo vbErrorHandler

    lLangID = GetSystemDefaultLangID
    'Setup a buffer to receive the settings
    lBuffSize = 50
    sBuffer = String$(lBuffSize, vbNullChar)
    lRet = VerLanguageName(lLangID, sBuffer, lBuffSize)
    If lRet > 0 Then
        SysLanguageName = Left$(sBuffer, lRet)
    End If
Exit Property
vbErrorHandler:
    Err.Raise Err.Number, "CGLocaleInfo GetSysLanguageName", Err.Description
End Property

Public Property Get Country() As String
    'Returns the Country Name eg. 'United Kingdom'
    Country = GetLocaleString(LOCALE_SENGCOUNTRY)
End Property

Public Property Get LanguageName() As String
    'Returns the Native Language Name eg. 'English'
    LanguageName = GetLocaleString(LOCALE_SNATIVELANGNAME)
End Property

Public Property Get NativeCountryName() As String
    NativeCountryName = GetLocaleString(LOCALE_SNATIVECTRYNAME)
End Property

Public Property Get PositiveSign() As String
    'Returns the symbol used for the positive sign eg. +
    PositiveSign = GetLocaleString(LOCALE_SPOSITIVESIGN)
End Property

Public Property Get NegativeSign() As String
' Returns the symbol used for the negative sign eg. -
    NegativeSign = GetLocaleString(LOCALE_SNEGATIVESIGN)
End Property

'-----------------
' Public Functions
'-----------------
Public Function DayName(ByVal iDayNum As Integer, Optional Short As Boolean = False) As String
    DayName = IIf(Short, GetLocaleString(LOCALE_SABBREVDAYNAME1 + iDayNum - 1), GetLocaleString(LOCALE_SDAYNAME1 + iDayNum - 1))
End Function

Public Function MonthName(ByVal iMonthNum As Integer, Optional Short As Boolean = False) As String
    MonthName = IIf(Short, GetLocaleString(LOCALE_SABBREVMONTHNAME1 - 1 + iMonthNum), GetLocaleString(LOCALE_SMONTHNAME1 + iMonthNum - 1))
End Function

'------------------
' Private Functions
'------------------
Private Function GetLocaleString(ByVal lLocaleNum As Long) As String
    'Generic routine to get the locale string from the Operating system.
    Dim lBuffSize As String
    Dim sBuffer As String
    Dim lRet As Long

    lBuffSize = 256
    sBuffer = String$(lBuffSize, vbNullChar)

    'Get the information from the registry
    lRet = GetLocaleInfo(LOCALE_USER_DEFAULT, lLocaleNum, sBuffer, lBuffSize)
    'If lRet > 0 then success - lret is the size of the string returned
    If lRet > 0 Then
        GetLocaleString = Left$(sBuffer, lRet - 1)
    End If
End Function
