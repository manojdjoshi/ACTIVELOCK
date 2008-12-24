Attribute VB_Name = "mdMMTimeUtils"
'
'Copyright (C) 2003-2005 Funambol
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'

'****************************************************************
'*Module:  mdMMTimeUtils.bas                                    *
'*Author:  Marco Magistrali - Funambol                          *
'****************************************************************
Option Explicit

' System time structure
Public Type SYSTEMTIME
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
        Bias As Long
        StandardName(0 To 31) As Integer
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(0 To 31) As Integer
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)

Private Declare Function SystemTimeToTzSpecificLocalTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpUniversalTime As SYSTEMTIME, lpLocalTime As SYSTEMTIME) As Long
Private Declare Function TzSpecificLocalTimeToSystemTime Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION, lpLocalTime As SYSTEMTIME, lpUniversalTime As SYSTEMTIME) As Long
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long


' Return current time as UTC
Public Function UTCTime() As Date
    Dim t As SYSTEMTIME
    
    GetSystemTime t
    UTCTime = DateSerial(t.wYear, t.wMonth, t.wDay) + TimeSerial(t.wHour, t.wMinute, t.wSecond) + t.wMilliseconds / 86400000#
End Function

' Return current time as local time
Public Function LocalTime() As Date
    Dim t As SYSTEMTIME
    
    GetLocalTime t
    LocalTime = DateSerial(t.wYear, t.wMonth, t.wDay) + TimeSerial(t.wHour, t.wMinute, t.wSecond) + t.wMilliseconds / 86400000
End Function

' Convert local to Utc time from current time zone returning a Date
Public Function LocalToUTC_Date(ByVal tDate As Date) As Date
    Dim tzi As TIME_ZONE_INFORMATION
    Dim stUTC As SYSTEMTIME
    Dim stLocal As SYSTEMTIME
    Dim localFileTime As FILETIME
    Dim utcFileTime As FILETIME
    Dim lRes As Long
    Dim value As String
    
    lRes = GetTimeZoneInformation(tzi)
    stLocal.wYear = Year(tDate)
    stLocal.wMonth = Month(tDate)
    stLocal.wDay = Day(tDate)
    stLocal.wHour = Hour(tDate)
    stLocal.wMinute = Minute(tDate)
    stLocal.wSecond = Second(tDate)
    stLocal.wMilliseconds = 0
          
    If mdMMVariables.OSVersion = WIN_XP Then
        lRes = TzSpecificLocalTimeToSystemTime(tzi, stLocal, stUTC)
    Else
        SystemTimeToFileTime stLocal, localFileTime
        LocalFileTimeToFileTime localFileTime, utcFileTime
        FileTimeToSystemTime utcFileTime, stUTC
    End If
    
    LocalToUTC_Date = DateSerial(stUTC.wYear, stUTC.wMonth, stUTC.wDay) + TimeSerial(stUTC.wHour, stUTC.wMinute, stUTC.wSecond)
    
End Function

' Convert local to Utc time for current time zone
'Public Function LocalToUTC(ByVal tDate As Date) As Date
Public Function LocalToUTC(ByVal tDate As Date) As String
    Dim tzi As TIME_ZONE_INFORMATION
    Dim stUTC As SYSTEMTIME
    Dim stLocal As SYSTEMTIME
    Dim localFileTime As FILETIME
    Dim utcFileTime As FILETIME
    Dim lRes As Long
    Dim value As String
    
    lRes = GetTimeZoneInformation(tzi)
    stLocal.wYear = Year(tDate)
    stLocal.wMonth = Month(tDate)
    stLocal.wDay = Day(tDate)
    stLocal.wHour = Hour(tDate)
    stLocal.wMinute = Minute(tDate)
    stLocal.wSecond = Second(tDate)
    stLocal.wMilliseconds = 0
    
    If mdMMVariables.OSVersion = WIN_XP Then
        lRes = TzSpecificLocalTimeToSystemTime(tzi, stLocal, stUTC)
    Else
        SystemTimeToFileTime stLocal, localFileTime
        LocalFileTimeToFileTime localFileTime, utcFileTime
        FileTimeToSystemTime utcFileTime, stUTC
    End If
                    
    'stLocal.wHour = stLocal.wHour + tzi.Bias
           
    'value = Space(DIM_VALUE)
    'getClientConfiguration "", PROPERTY_SYSTEMTIME_TO_LOCALTIME, value, DIM_VALUE
    
    'If Trim(value) = "1" Then
        'lRes = TzSpecificLocalTimeToSystemTime(tzi, stLocal, stUTC)
        LocalToUTC = Format(DateSerial(stUTC.wYear, stUTC.wMonth, stUTC.wDay) + TimeSerial(stUTC.wHour, stUTC.wMinute, stUTC.wSecond), "yyyyMMddTHHmmssZ")
    '    Exit Function
    'End If
        
    'LocalToUTC = Format(DateSerial(stLocal.wYear, stLocal.wMonth, stLocal.wDay) + TimeSerial(stLocal.wHour, stLocal.wMinute, stLocal.wSecond), "yyyyMMddTHHmmss")
    
End Function
' Convert UTC time to local time for current time zone
Public Function UTCToLocal(ByVal tDate As Date) As Date
    Dim tzi As TIME_ZONE_INFORMATION
    Dim stUTC As SYSTEMTIME
    Dim stLocal As SYSTEMTIME
    Dim lRes As Long
    Dim value As String
    Dim localFileTime As FILETIME
    Dim utcFileTime As FILETIME
    
    lRes = GetTimeZoneInformation(tzi)
    stUTC.wYear = Year(tDate)
    stUTC.wMonth = Month(tDate)
    stUTC.wDay = Day(tDate)
    stUTC.wHour = Hour(tDate)
    stUTC.wMinute = Minute(tDate)
    stUTC.wSecond = Second(tDate)
    stUTC.wMilliseconds = 0
    
    
    If mdMMVariables.OSVersion = WIN_XP Then
        lRes = SystemTimeToTzSpecificLocalTime(tzi, stUTC, stLocal)
    Else
        SystemTimeToFileTime stUTC, utcFileTime
        FileTimeToLocalFileTime utcFileTime, localFileTime
        FileTimeToSystemTime localFileTime, stLocal
    End If
        
    'value = Space(DIM_VALUE)
    'getClientConfiguration "", PROPERTY_SYSTEMTIME_TO_LOCALTIME, value, DIM_VALUE
    '
    'If Trim(value) = "1" Then
        'lRes = SystemTimeToTzSpecificLocalTime(tzi, stUTC, stLocal)
        UTCToLocal = DateSerial(stLocal.wYear, stLocal.wMonth, stLocal.wDay) + TimeSerial(stLocal.wHour, stLocal.wMinute, stLocal.wSecond)
         
        'Exit Function
    'End If
    
    'UTCToLocal = DateSerial(stUTC.wYear, stUTC.wMonth, stUTC.wDay) + TimeSerial(stUTC.wHour, stUTC.wMinute, stUTC.wSecond)

    
End Function

'return True if timezone is negative. it will be used to set Backward to true
Public Function getTimeZoneInfo()
    Dim tzi As TIME_ZONE_INFORMATION
    Dim lRes As Long
    lRes = GetTimeZoneInformation(tzi)
    
    If tzi.Bias < 0 Then
        getTimeZoneInfo = True
    Else
        getTimeZoneInfo = False
    End If
    
End Function

Public Function dateToCanonical(ByRef dateString As String) As String
    
    Dim res As String
    Dim internal As String
    internal = dateString
    res = ""
    
    If InStr(internal, "/") <> 0 Then
        dateToCanonical = internal
        Exit Function
    End If
    'res = Left(internal, 4) + Right(Left(internal, 6), 2) + Right(Left(internal, 8), 2) _
    '    + Right(Left(internal, 11), 2) + Right(Left(internal, 13), 2) + Right(Left(internal, 15), 2)
    
    'res = Right(Left(internal, 8), 2) + "/" + Right(Left(internal, 6), 2) + "/" + Left(internal, 4)
    res = Left(internal, 4) + "/" + Right(Left(internal, 6), 2) + "/" + Right(Left(internal, 8), 2)
    res = res + " " + Right(Left(internal, 11), 2) + ":" + Right(Left(internal, 13), 2) + ":" + Right(Left(internal, 15), 2)
    
    'res = Mid(dateString, 0, 4) & Mid(dateString, 4, 6) & Mid(dateString, 6, 8) & Mid(dateString, 9, 11) & Mid(dateString, 11, 13) & Mid(dateString, 13, 15)
    
    dateToCanonical = res
End Function

Public Function dateToCanonicalBirth(ByRef dateString As String) As String
    
    Dim res As String
    Dim internal As String
    internal = dateString
    res = ""
    
    'res = Right(internal, 2) + "/" + Right(Left(internal, 7), 2) + "/" + Left(internal, 4)
    res = Left(internal, 4) + "/" + Right(Left(internal, 7), 2) + "/" + Right(internal, 2)
    res = res + " " + "00" + ":" + "00" + ":" + "00"
    
    
    dateToCanonicalBirth = res
End Function



' Convert Local time to UTC time if possible.
' Note: The function may return 0, 1 or 2 datetime values.
' During switch from Standard to Daylight time, there is (usually) one hour missing
' and there is an invalid time range of one hour. This function returns an empty collection.
' When time changes from Daylight to Standard, there is an ambguity so one local time
' corresponds to two different UTC times. This function then returns a collection with
' two elements.
'Public Function LocalToUTC(ByVal tDate As Date) As Collection
'    Dim tzi As TIME_ZONE_INFORMATION
'    Dim lRes As Long
'    Dim col As Collection
'    Dim tUTC As Date
'
'    Set col = New Collection
'    lRes = GetTimeZoneInformation(tzi)
'    tUTC = tDate + tzi.Bias / 1440
'    If tzi.StandardDate.wMonth = 0 Then
'        ' No daylight time -- no problem
'        col.Add tUTC
'    Else
'        ' Assume we are fuzzing with +- one daylight bias, which is normally negative.
'        ' So, datetimes will be ordered from earliest to latest in collection.
'        If Round(UTCToLocal(tUTC + tzi.DaylightBias / 1440) * 86400) = Round(tDate * 86400) Then
'            col.Add CDate(tUTC + tzi.DaylightBias / 1440)
'        End If
'        If Round(UTCToLocal(tUTC) * 86400) = Round(tDate * 86400) Then
'            col.Add tUTC
'        End If
'        If Round(UTCToLocal(tUTC - tzi.DaylightBias / 1440) * 86400) = Round(tDate * 86400) Then
'            col.Add CDate(tUTC - tzi.DaylightBias / 1440)
'        End If
'    End If
'    Set LocalToUTC = col
'End Function


