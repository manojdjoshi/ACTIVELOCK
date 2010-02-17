Attribute VB_Name = "modTrial"
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
' *   Copyright 2003-2009 The ActiveLock Software Group (ASG)
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
'===============================================================================
' Name: modTrialGlobals
' Purpose: This module is used by the Trial Period/Runs feature
' Functions:
' Properties:
' Methods:
' Started: 08.15.2005
' Modified: 08.15.2005
'===============================================================================

Option Explicit

' encrypt text messages using AES_Rijndael Block Cipher
Private oTest As clsRijndael

Public EXPIRED_RUNS As String
Public EXPIRED_DAYS As String
Public LICENSE_SOFTWARE_NAME As String
Public LICENSE_SOFTWARE_PASSWORD As String
Public LICENSE_SOFTWARE_VERSION As String
Public alockDays As Integer, alockRuns As Integer
Public trialPeriod As Boolean, trialRuns As Boolean
Public TEXTMSG_DAYS As String, TEXTMSG_RUNS As String, TEXTMSG As String
Public VIDEO As String, OTHERFILE As String

Public Const HIDDENFOLDER As String = "37436AE1A6BBB443B4B8535477D3B8F595E77FBD211E0326CB92D381B2D29FD6F8DB609FB3851CB666FA025A9E26F84F2284CE0BFFF9BC845E32854250833654"
Public Const EXPIREDDAYS As String = "ExpiredDays"
Public Const INITIALDATE As String = "01/01/2000"
Public Const TrialWarning As String = "Trial Warning"
Public Const EXPIREDWARNING As String = "Expired Warning"
Public Const CLSIDSTR As String = "{645FF040-5081-101B-9F08-00AA002F954E}"
Public Const CHANNELS As String = "CTDChannels_Version."

Public Const REG_MULTI_SZ = 7
Public Const ERROR_MORE_DATA = 234

' Windows Security Messages
Const KEY_ALL_CLASSES As Long = &HF0063

Const ERROR_SUCCESS = 0&
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long

' Windows Registry API calls
Public Declare Function RegSetValueExString _
  Lib "advapi32.dll" Alias "RegSetValueExA" _
 (ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal Reserved As Long, _
  ByVal dwType As Long, _
  ByVal lpValue As String, _
  ByVal cbData As Long) As Long

Public Declare Function RegSetValueExLong _
  Lib "advapi32.dll" Alias "RegSetValueExA" _
 (ByVal hKey As Long, _
  ByVal lpValueName As String, _
  ByVal Reserved As Long, _
  ByVal dwType As Long, _
  lpValue As Long, _
  ByVal cbData As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Public Declare Function RegQueryValueExString Lib "advapi32.dll" _
    Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As String, _
    lpcbData As Long) As Long

Public Declare Function RegQueryValueExLong Lib "advapi32.dll" _
    Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Long, _
    lpcbData As Long) As Long

Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" _
    Alias "RegQueryValueExA" _
    (ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    ByVal lpData As Long, _
    lpcbData As Long) As Long

Public Declare Function ExpandEnvironmentStrings Lib "kernel32" _
    Alias "ExpandEnvironmentStringsA" _
    (ByVal lpSrc As String, _
    ByVal lpDst As String, _
    ByVal nSize As Long) As Long

Public Declare Function GlobalLock Lib "kernel32" _
   (ByVal hMem As Long) As Long
Public Declare Function GlobalUnlock Lib "kernel32" _
   (ByVal hMem As Long) As Long
Public Declare Function GlobalAlloc Lib "kernel32" _
   (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" _
   (ByVal hMem As Long) As Long

Private Declare Function FindFirstFile Lib "kernel32" _
 Alias "FindFirstFileA" (ByVal lpFileName As String, _
 lpFindFileData As WIN32_FIND_DATA) As Long
 
Private Declare Function FindNextFile Lib "kernel32" _
 Alias "FindNextFileA" (ByVal hFindFile As Long, _
 lpFindFileData As WIN32_FIND_DATA) As Long
 
Private Declare Function FindClose Lib "kernel32" _
 (ByVal hFindFile As Long) As Long
 
Private Declare Function SearchPath Lib "kernel32" _
 Alias "SearchPathA" (ByVal lpPath As String, _
 ByVal lpFileName As String, ByVal lpExtension As String, _
 ByVal nBufferLength As Long, ByVal lpBuffer As String, _
 ByVal lpFilePart As String) As Long
 
 Private Const MAX_PATH = 260

Private Type FILETIME
    lngLowDateTime As Long
    lngHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
    lngFileAttributes As Long           ' File attributes
    ftCreationTime As FILETIME          ' Creation time
    ftLastAccessTime As FILETIME        ' Last access time
    ftLastWriteTime As FILETIME         ' Last modified time
    lngFileSizeHigh As Long             ' Size (high word)
    lngFileSizeLow As Long              ' Size (low word)
    lngReserved0 As Long                ' reserved
    lngReserved1 As Long                ' reserved
    strFilename As String * MAX_PATH    ' File name
    strAlternate As String * 14         ' 8.3 name
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Variant
    bInheritHandle As Boolean
End Type

Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Private Declare Function CreateFile Lib "kernel32" _
   Alias "CreateFileA" _
  (ByVal lpFileName As String, _
   ByVal dwDesiredAccess As Long, _
   ByVal dwShareMode As Long, _
   ByVal lpSecurityAttributes As Long, _
   ByVal dwCreationDisposition As Long, _
   ByVal dwFlagsAndAttributes As Long, _
   ByVal hTemplateFile As Long) As Long
 
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

'Spy Scan stuff

Public Const CREATE_NEW = 1
Public Const CREATE_ALWAYS = 2
Public Const OPEN_EXISTING = 3
Public Const OPEN_ALWAYS = 4
Public Const TRUNCATE_EXISTING = 5

Private Const FILE_BEGIN = 0
Private Const FILE_CURRENT = 1
Private Const FILE_END = 2

Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Private Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Private Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Private Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Private Const FILE_FLAG_POSIX_SEMANTICS = &H1000000

' Called from frmC
Public Declare Function FinWin Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function CF Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessID As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, ByVal lpBaseAddress As Any, ByVal lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long

Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000

Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long

Public Const TH32CS_SNAPPROCESS As Long = 2&

Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * 260
End Type
    
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_EXECUTE = &H20000000
Public Const GENERIC_ALL = &H10000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_SHARE_DELETE As Long = &H4
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800

Public Const EAV = &HC0000005

Public ProcessName$(256)
Public ProcessID(256) As Long
Public hFile As Long, retVal As Long
Public wX As Long, wY As Long
Public myHandle, Buffer As String
Public varchk, encvar$(4000)
Public HAD2HAMMER As Boolean

' Folder Date Stamp API
Private Type SYSTEMTIME
  wYear          As Integer
  wMonth         As Integer
  wDayOfWeek     As Integer
  wDay           As Integer
  wHour          As Integer
  wMinute        As Integer
  wSecond        As Integer
  wMilliseconds  As Long
End Type

Private Declare Function GetFileTime Lib "kernel32" _
  (ByVal hFile As Long, _
   lpCreationTime As FILETIME, _
   lpLastAccessTime As FILETIME, _
   lpLastWriteTime As FILETIME) As Long
   
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" _
  (lpFileTime As FILETIME, _
   lpLocalFileTime As FILETIME) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" _
  (lpFileTime As FILETIME, _
   lpSystemTime As SYSTEMTIME) As Long

Private Declare Function SystemTimeToFileTime Lib "kernel32" _
  (lpSystemTime As SYSTEMTIME, _
   lpFileTime As FILETIME) As Long

Private Declare Function LocalFileTimeToFileTime Lib "kernel32" _
  (lpLocalFileTime As FILETIME, _
   lpFileTime As FILETIME) As Long

Private Declare Function SetFileTime Lib "kernel32" _
  (ByVal hFile As Long, _
   lpCreationTime As FILETIME, _
   lpLastAccessTime As Any, _
   lpLastWriteTime As Any) As Long

Private Declare Function InternetOpen Lib "wininet.dll" _
    Alias "InternetOpenA" (ByVal sAgent As String, _
    ByVal lAccessType As Long, ByVal sProxyName As String, _
    ByVal sProxyBypass As String, ByVal lFlags As Long) As Long

Private Declare Function InternetOpenUrl Lib "wininet.dll" _
   Alias "InternetOpenUrlA" (ByVal hInternetSession As Long, _
   ByVal sUrl As String, ByVal sHeaders As String, _
   ByVal lHeadersLength As Long, ByVal lFlags As Long, _
   ByVal lContext As Long) As Long
   
Private Declare Function InternetReadFile Lib "wininet.dll" _
   (ByVal hFile As Long, ByVal sBuffer As String, _
   ByVal lNumberOfBytesToRead As Long, _
   lNumberOfBytesRead As Long) As Integer
   
Private Declare Function InternetCloseHandle Lib "wininet.dll" _
      (ByVal hInet As Long) As Integer

'* converts a date to a double and then to a string
'* useful for ensuring dates can be read when pulled back out
'* no matter the locale
Public Function DateToDblString(ByRef Dte As Date) As String
    DateToDblString = CStr(CDbl(Dte))
End Function

''' Function is used to convert doubles stored in strings to dates
''' useful becuase whenever we store dates we convert them (to doubles and then strings)
''' in case the user changes the locale in between storage and retrieval.
''' minor handling of actual date strings for some semblance of backward
''' compatibility
Public Function DblStringToDate(ByRef Dstr As String) As Date
        
On Error GoTo DblStringToDateError
       
If Dstr <> "" Then
    Dim Dbl As Double
    Dbl = CDbl(Dstr)
    DblStringToDate = CDate(Dbl)
    Exit Function
End If

DblStringToDateError:
    On Error GoTo DblStringToDateError2
    'probably a date it wasn't empty
    DblStringToDate = CDate(Dstr)
    Exit Function
DblStringToDateError2:
    DblStringToDate = #1/1/1900#

End Function

Public Function myDir() As String
myDir = "3BFE841EE459E0EC0DDBDCB91CA0A5879D751E21622A19741A6EEC92B19E1B5C"
End Function

Public Function OpenURL(ByVal sUrl As String) As String
Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Const INTERNET_FLAG_RELOAD = &H80000000
Dim hSession As Long
Dim hFile As Long
Dim Result As String
Dim Buffer As String * 2048
Dim bResult As Boolean
Dim lRead As Long
'This is where the website is grabbed
Buffer = ""
Result = ""
'Structure your internet request like below if you have a proxy
'hSession = InternetOpen("VBandWinInet/1.0", 3, "http://proxyIP:port", "", 0)
'hSession = InternetOpen("VBandWinInet/1.0", 3, "", "", 0)
hSession = InternetOpen("VB OpenUrl", INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
If hSession Then
hFile = InternetOpenUrl(hSession, sUrl, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
If hFile Then
        lRead = 1
        While lRead
            If InternetReadFile(hFile, Buffer, Len(Buffer), lRead) Then
                Result = Result + Buffer
            End If
        Wend
        OpenURL = Result
        InternetCloseHandle (hFile)
    End If
    InternetCloseHandle (hSession)
End If
End Function

'===============================================================================
' Name: Function DateGoodRegistry
' Input:
'   ByRef numDays As Integer - Number of trial days the application is good for
'   ByRef daysLeft As Integer - Days left in the trial period
' Output:
'   Boolean - Returns True if the date is good
' Purpose: The purpose of this module is to allow you to place a time
' limit on the unregistered use of your shareware application.
'       <p>Example:
'       <br>If DateGoodRegistry(30)=False Then
'       <br>CrippleApplication
'       <br>End if
' Remarks: This module can not be defeated by rolling back the system clock.
'===============================================================================
Public Function DateGoodRegistry(numDays As Integer, daysLeft As Integer) As Boolean
'Ex: If DateGoodRegistry(30)=False Then
' CrippleApplication
' End if
'Registry Parameters:
' CRD: Current Run Date
' LRD: Last Run Date
' FRD: First Run Date
Dim TmpCRD As Date
Dim TmpLRD As Date
Dim TmpFRD As Date

On Error GoTo DateGoodRegistryError

TmpCRD = ActiveLockDate(UTC(Now))
'TmpCRD = UTC(Now)
TmpLRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
TmpFRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
'TmpLRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
'TmpFRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
DateGoodRegistry = False

'If this is the applications first load, write initial settings
'to the registry
If TmpLRD = dec2("93.8D.93.8D.96.90.90.90") Then '1/1/2000
    SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(CStr(TmpCRD))
    SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", enc2(CStr(TmpCRD))
    'SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(DateToDblString(TmpCRD))
    'SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", enc2(DateToDblString(TmpCRD))
    SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2("0")
End If
'Read LRD and FRD from registry
TmpLRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
TmpFRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
'TmpLRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
'TmpFRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000

If ActiveLockDate(TmpFRD) > ActiveLockDate(TmpCRD) Then 'System clock rolled back
'If TmpFRD > TmpCRD Then 'System clock rolled back
    DateGoodRegistry = False
ElseIf ActiveLockDate(UTC(Now)) > DateAdd("d", numDays, ActiveLockDate(TmpFRD)) Then 'trial expired
'ElseIf UTC(Now) > DateAdd("d", numDays, TmpFRD) Then 'trial expired
    SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS)
    DateGoodRegistry = False
ElseIf ActiveLockDate(TmpCRD) > ActiveLockDate(TmpLRD) Then 'Everything OK write New LRD date
'ElseIf TmpCRD > TmpLRD Then 'Everything OK write New LRD date
    SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(CStr(TmpCRD))
    'SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(DateToDblString(TmpCRD))
    DateGoodRegistry = True
ElseIf ActiveLockDate(TmpCRD) = ActiveLockDate(TmpLRD) Then
'ElseIf TmpCRD = TmpLRD Then
    DateGoodRegistry = True
Else
    DateGoodRegistry = False
End If
If DateGoodRegistry Then
    daysLeft = DateAdd("d", numDays, ActiveLockDate(TmpFRD)) - ActiveLockDate(UTC(Now))
    'daysLeft = DateAdd("d", numDays, TmpFRD) - UTC(Now)
Else
    daysLeft = 0
End If
Exit Function
DateGoodRegistryError:
    daysLeft = 0
    DateGoodRegistry = False
    Exit Function
End Function
'===============================================================================
' Name: Function RunsGoodRegistry
' Input:
'   ByRef numRuns As Integer - Number of trial days the application is good for
'   ByRef runsLeft As Integer - Days left in the trial period
' Output:
'   Boolean - Returns True if the Runs is good
' Purpose: The purpose of this module is to allow you to place a time
' limit on the unregistered use of your shareware application.
' Remarks: None
'===============================================================================
Public Function RunsGoodRegistry(numRuns As Integer, runsLeft As Integer) As Boolean
On Error GoTo RunsGoodRegistryError
Dim TmpCRD As Date
Dim TmpLRD As Date
Dim TmpFRD As Date

TmpCRD = INITIALDATE
TmpLRD = dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90")) '1/1/2000
TmpFRD = dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90")) '1/1/2000
'TmpLRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
'TmpFRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
runsLeft = Int(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(numRuns - 1))))) - 1
RunsGoodRegistry = False

'If this is the applications first load, write initial settings
'to the registry
If TmpLRD = dec2("93.8D.93.8D.96.90.90.90") Then '1/1/2000
    SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(CStr(Now))
    SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", enc2(CStr(TmpCRD))
    'SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(DateToDblString(Now))
    'SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", enc2(DateToDblString(TmpCRD))
    SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(numRuns - 1))
End If
'Read LRD and FRD from registry
If TmpLRD = dec2("93.8D.93.8D.96.90.90.90") Then
    runsLeft = Int(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(numRuns - 1)))))
Else
    runsLeft = Int(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(numRuns - 1))))) - 1
End If
TmpLRD = dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90")) '1/1/2000
TmpFRD = dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90")) '1/1/2000
'TmpLRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
'TmpFRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
If TmpLRD = vbEmpty Then
    TmpLRD = INITIALDATE
    TmpFRD = INITIALDATE
    runsLeft = numRuns - 1
End If

If runsLeft < 0 Then 'impossible
    RunsGoodRegistry = False
ElseIf runsLeft > numRuns Then 'Trial runs expired
    SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS)
    RunsGoodRegistry = False
ElseIf numRuns >= runsLeft Then 'Everything OK write the remaining number of runs
    If TmpLRD = INITIALDATE Then
        SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(numRuns - 1))
    Else
        SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(runsLeft))
    End If
    RunsGoodRegistry = True
Else
    RunsGoodRegistry = False
End If
If RunsGoodRegistry Then
Else
    runsLeft = 0
End If
Exit Function
RunsGoodRegistryError:
    runsLeft = 0
    RunsGoodRegistry = False
    Exit Function
End Function

Public Function DateGood(ByVal numDays As Integer, daysLeft As Integer, TrialHideTypes As ALTrialHideTypes) As Boolean
Dim use2 As Boolean
Dim use3 As Boolean, use4 As Boolean
Dim daysLeft1 As Integer, daysLeft2 As Integer
Dim daysLeft3 As Integer, daysLeft4 As Integer

TEXTMSG_DAYS = DecryptMyString("0DEAD685B3E70293A72D7BF2A5947CBED433B490DE5286509B9A4953B71190634432E5E20DCEFACC22E237072924B18B2DD7D1355E284B65DF70BD6D536B1D8E", PSWD)
DateGood = False

If TrialSteganographyExists(TrialHideTypes) Then
    If DateGoodSteganography(numDays, daysLeft2) = False Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS
        Exit Function
    End If
    use2 = True
End If
If TrialHiddenFolderExists(TrialHideTypes) Then
    If DateGoodHiddenFolder(numDays, daysLeft3) = False Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS
        Exit Function
    End If
    use3 = True
End If
If TrialRegistryPerUserExists(TrialHideTypes) Then
    If DateGoodRegistry(numDays, daysLeft4) = False Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS
        Exit Function
    End If
    use4 = True
End If

DateGood = True
If (use2 And Not use3 And Not use4) Then
    daysLeft = daysLeft2
ElseIf (Not use2 And use3 And Not use4) Then
    daysLeft = daysLeft3
ElseIf (Not use2 And Not use3 And use4) Then
    daysLeft = daysLeft4
ElseIf (use2 And use3 And Not use4) And daysLeft2 = daysLeft3 Then
    daysLeft = daysLeft2
ElseIf (use2 And Not use3 And use4) And daysLeft2 = daysLeft4 Then
    daysLeft = daysLeft2
ElseIf (Not use2 And use3 And use4) And daysLeft3 = daysLeft4 Then
    daysLeft = daysLeft3
ElseIf (use2 And use3 And use4) And daysLeft2 = daysLeft3 And daysLeft2 = daysLeft4 Then
    daysLeft = daysLeft2
Else
    DateGood = False
    daysLeft = 0
End If
End Function
Public Function RunsGood(ByVal numRuns As Integer, runsLeft As Integer, TrialHideTypes As ALTrialHideTypes) As Boolean
Dim use2 As Boolean
Dim use3 As Boolean, use4 As Boolean
Dim runsLeft1 As Integer, runsLeft2 As Integer
Dim runsLeft3 As Integer, runsLeft4 As Integer
TEXTMSG_RUNS = DecryptMyString("2FD15B7C7C504433519133B9D336D12627384959E9D4D4157F12158C92DF57B234B46AFEFBD8C8FA5463FD614BFD13AF9FBD86B0C042663111253C50C05F6931", PSWD)

RunsGood = False

If TrialSteganographyExists(TrialHideTypes) Then
    If RunsGoodSteganography(numRuns, runsLeft2) = False Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS
        Exit Function
    End If
    use2 = True
End If

If TrialHiddenFolderExists(TrialHideTypes) Then
    If RunsGoodHiddenFolder(numRuns, runsLeft3) = False Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS
        Exit Function
    End If
    use3 = True
End If

If TrialRegistryPerUserExists(TrialHideTypes) Then
    If RunsGoodRegistry(numRuns, runsLeft4) = False Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS
        Exit Function
    End If
    use4 = True
End If

RunsGood = True
If (use2 And Not use3 And Not use4) Then
    runsLeft = runsLeft2
ElseIf (Not use2 And use3 And Not use4) Then
    runsLeft = runsLeft3
ElseIf (Not use2 And Not use3 And use4) Then
    runsLeft = runsLeft4
ElseIf (use2 And use3 And Not use4) And runsLeft2 = runsLeft3 Then
    runsLeft = runsLeft2
ElseIf (use2 And Not use3 And use4) And runsLeft2 = runsLeft4 Then
    runsLeft = runsLeft2
ElseIf (Not use2 And use3 And use4) And runsLeft3 = runsLeft4 Then
    runsLeft = runsLeft3
ElseIf (use2 And use3 And use4) And runsLeft2 = runsLeft3 And runsLeft2 = runsLeft4 Then
    runsLeft = runsLeft2
Else
    RunsGood = False
    runsLeft = 0
End If
End Function


Public Function DateGoodHiddenFolder(numDays As Integer, daysLeft As Integer) As Boolean
'Hidden Folder Parameters:
' CRD: Current Run Date
' LRD: Last Run Date
' FRD: First Run Date
Dim TmpCRD As Date
Dim TmpLRD As Date
Dim TmpFRD As Date
Dim strMyString As String, strSource As String
Dim intFF As Integer

On Error GoTo DateGoodHiddenFolderError
If FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(myDir, PSWD)) = False Then
    MkDir ActivelockGetSpecialFolder(46) & DecryptMyString(myDir, PSWD)
End If
strSource = HiddenFolderFunction()
If Dir(ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD), vbDirectory + vbHidden + vbReadOnly + vbSystem) <> "" Then
    MinusAttributes
    'Check to see if our file is there
    If fileExist(strSource) Then
        SetAttr strSource, vbNormal
    End If
ElseIf FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD)) = True Then
    'Ok, the folder is there with no hidden, system attributes
    'Check to see if our file is there
    If fileExist(strSource) Then
        SetAttr strSource, vbNormal
    End If
Else
    MkDir ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD)
End If

CreateHdnFile

'* TmpCRD = ActiveLockDate(UTC(Now))
TmpCRD = UTC(Now)
Dim a() As String
If fileExist(strSource) Then
    ' Read the file...
    intFF = FreeFile
    Open strSource For Input As #intFF
        strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
        If Right(strMyString, 2) = vbCrLf Then _
            strMyString = Left(strMyString, Len(strMyString) - 2)
        If Right(strMyString, 1) = vbCrLf Then _
            strMyString = Left(strMyString, Len(strMyString) - 1)
        strMyString = DecryptMyString(strMyString, PSWD)
    Close #intFF
    If strMyString <> "" Then
        a = Split(strMyString, "_")
        If a(1) <> "" Then TmpLRD = a(1)
        If a(2) <> "" Then TmpFRD = a(2)
        'If a(1) <> "" Then TmpLRD = DblStringToDate(a(1))
        'If a(2) <> "" Then TmpFRD = DblStringToDate(a(2))
        If a(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
            DateGoodHiddenFolder = False
            Exit Function
        End If
    End If
    If TmpLRD = vbEmpty Then TmpLRD = INITIALDATE
    If TmpFRD = vbEmpty Then TmpFRD = INITIALDATE
Else
    TmpLRD = INITIALDATE
    TmpFRD = INITIALDATE
End If
DateGoodHiddenFolder = False

'If this is the applications first load, write initial settings
'to Hidden Folder
If TmpLRD = INITIALDATE Then
    TmpLRD = TmpCRD
    TmpFRD = TmpCRD
    ' Write to the file...
    intFF = FreeFile
    Open strSource For Output As #intFF
    Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & TmpLRD & "_" & TmpFRD & "_" & "0", PSWD)
    'Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(TmpLRD) & "_" & DateToDblString(TmpFRD) & "_" & "0", PSWD)
    Close #intFF
End If
'Read LRD and FRD from Hidden Folder
Dim b() As String
' Read the file...
intFF = FreeFile
Open strSource For Input As #intFF
    strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
    If Right(strMyString, 2) = vbCrLf Then _
        strMyString = Left(strMyString, Len(strMyString) - 2)
    If Right(strMyString, 1) = vbCrLf Then _
        strMyString = Left(strMyString, Len(strMyString) - 1)
    strMyString = DecryptMyString(strMyString, PSWD)
Close #intFF
b = Split(strMyString, "_")
If b(1) <> "" Then TmpLRD = b(1)
If b(2) <> "" Then TmpFRD = b(2)
'If b(1) <> "" Then TmpLRD = DblStringToDate(b(1))
'If b(2) <> "" Then TmpFRD = DblStringToDate(b(2))
If b(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
    DateGoodHiddenFolder = False
    Exit Function
End If
If TmpLRD = vbEmpty Then TmpLRD = INITIALDATE
If TmpFRD = vbEmpty Then TmpFRD = INITIALDATE

If ActiveLockDate(TmpFRD) > ActiveLockDate(TmpCRD) Then 'System clock rolled back
'If TmpFRD > TmpCRD Then 'System clock rolled back
    DateGoodHiddenFolder = False
ElseIf ActiveLockDate(UTC(Now)) > DateAdd("d", numDays, ActiveLockDate(TmpFRD)) Then 'Trial expired
'ElseIf UTC(Now) > DateAdd("d", numDays, TmpFRD) Then 'Trial expired
    ' Write to the file...
    intFF = FreeFile
    Open strSource For Output As #intFF
        Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS, PSWD)
    Close #intFF
    DateGoodHiddenFolder = False
ElseIf ActiveLockDate(TmpCRD) > ActiveLockDate(TmpLRD) Then 'Everything OK write New LRD date
'ElseIf TmpCRD > TmpLRD Then 'Everything OK write New LRD date
    ' Write to the file...
    intFF = FreeFile
    Open strSource For Output As #intFF
    Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & TmpCRD & "_" & TmpFRD & "_" & "0", PSWD)
    'Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(TmpCRD) & "_" & DateToDblString(TmpFRD) & "_" & "0", PSWD)
    Close #intFF
    DateGoodHiddenFolder = True
ElseIf ActiveLockDate(TmpCRD) = ActiveLockDate(TmpLRD) Then
'ElseIf TmpCRD = TmpLRD Then
    DateGoodHiddenFolder = True
Else
    DateGoodHiddenFolder = False
End If

SetAttr strSource, vbReadOnly + vbHidden + vbSystem
PlusAttributes

If DateGoodHiddenFolder Then
    daysLeft = DateAdd("d", numDays, ActiveLockDate(TmpFRD)) - ActiveLockDate(UTC(Now))
    'daysLeft = DateAdd("d", numDays, TmpFRD) - UTC(Now)
Else
    daysLeft = 0
End If
Exit Function

DateGoodHiddenFolderError:

End Function

Public Function RunsGoodHiddenFolder(numRuns As Integer, runsLeft As Integer) As Boolean
Dim TmpCRD As Date
Dim TmpLRD As Date
Dim TmpFRD As Date
Dim strMyString As String, strSource As String
Dim intFF As Integer

On Error GoTo RunsGoodHiddenFolderError
If FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(myDir, PSWD)) = False Then
    MkDir ActivelockGetSpecialFolder(46) & DecryptMyString(myDir, PSWD)
End If
strSource = HiddenFolderFunction()
If Dir(ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD), vbDirectory + vbHidden + vbReadOnly + vbSystem) <> "" Then
    MinusAttributes
    'Check to see if our file is there
    If fileExist(strSource) Then
        SetAttr strSource, vbNormal
    End If
ElseIf FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD)) = True Then
    'Ok, the folder is there with no hidden, system attributes
    'Check to see if our file is there
    If fileExist(strSource) Then
        SetAttr strSource, vbNormal
    End If
Else
    MkDir ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD)
End If

CreateHdnFile

TmpCRD = INITIALDATE
Dim a() As String
If fileExist(strSource) Then
    ' Read the file...
    intFF = FreeFile
    Open strSource For Input As #intFF
        strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
        If Right(strMyString, 2) = vbCrLf Then _
            strMyString = Left(strMyString, Len(strMyString) - 2)
        If Right(strMyString, 1) = vbCrLf Then _
            strMyString = Left(strMyString, Len(strMyString) - 1)
        strMyString = DecryptMyString(strMyString, PSWD)
    Close #intFF
    If strMyString <> "" Then
        On Error GoTo continue
        a = Split(strMyString, "_")
        If a(1) <> "" Then TmpLRD = a(1)
        If a(2) <> "" Then TmpFRD = a(2)
        'If a(1) <> "" Then TmpLRD = DblStringToDate(a(1)) '*
        'If a(2) <> "" Then TmpFRD = DblStringToDate(a(2)) '*
        runsLeft = Int(a(3)) - 1
        If a(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
            RunsGoodHiddenFolder = False
            Exit Function
        End If
    End If
continue:
    If TmpLRD = vbEmpty Then
        TmpLRD = INITIALDATE
        TmpFRD = INITIALDATE
        runsLeft = numRuns - 1
    End If
Else
    TmpLRD = INITIALDATE
    TmpFRD = INITIALDATE
    runsLeft = numRuns - 1
End If
RunsGoodHiddenFolder = False

'If this is the applications first load, write initial settings
'to Hidden Folder
If TmpLRD = INITIALDATE Then
    TmpFRD = INITIALDATE
    runsLeft = numRuns - 1
    ' Write to the file...
    intFF = FreeFile
    Open strSource For Output As #intFF
    Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(UTC(Now)) & "_" & TmpFRD & "_" & CStr(numRuns - 1), PSWD)
    'Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(UTC(Now)) & "_" & DateToDblString(TmpFRD) & "_" & CStr(numRuns - 1), PSWD)
    Close #intFF
End If
'Read LRD and FRD from Hidden Folder
Dim b() As String
' Read the file...
intFF = FreeFile
Open strSource For Input As #intFF
    strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
    If Right(strMyString, 2) = vbCrLf Then _
        strMyString = Left(strMyString, Len(strMyString) - 2)
    If Right(strMyString, 1) = vbCrLf Then _
        strMyString = Left(strMyString, Len(strMyString) - 1)
    strMyString = DecryptMyString(strMyString, PSWD)
Close #intFF
b = Split(strMyString, "_")
If TmpLRD = INITIALDATE Then
    runsLeft = b(3)
Else
    runsLeft = b(3) - 1
End If
If b(1) <> "" Then TmpLRD = b(1)
If b(2) <> "" Then TmpFRD = b(2)
'If b(1) <> "" Then TmpLRD = DblStringToDate(b(1)) '*
'If b(2) <> "" Then TmpFRD = DblStringToDate(b(2)) '*
If b(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
    RunsGoodHiddenFolder = False
    Exit Function
End If
If TmpLRD = vbEmpty Then
    TmpLRD = INITIALDATE
    TmpFRD = INITIALDATE
    runsLeft = numRuns - 1
End If

If runsLeft < 0 Then 'impossible
    RunsGoodHiddenFolder = False
ElseIf runsLeft > numRuns Then 'Trial runs expired
    ' Write to the file...
    intFF = FreeFile
    Open strSource For Output As #intFF
        Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS, PSWD)
    Close #intFF
    RunsGoodHiddenFolder = False
ElseIf numRuns >= runsLeft Then 'Everything OK write the remaining number of runs
    ' Write to the file...
    intFF = FreeFile
    Open strSource For Output As #intFF
        If TmpLRD = INITIALDATE Then
            Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(UTC(Now)) & "_" & TmpFRD & "_" & CStr(numRuns - 1), PSWD)
            'Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(UTC(Now)) & "_" & DateToDblString(TmpFRD) & "_" & CStr(numRuns - 1), PSWD)
        Else
            Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(UTC(Now)) & "_" & TmpFRD & "_" & CStr(runsLeft), PSWD)
            'Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(UTC(Now)) & "_" & DateToDblString(TmpFRD) & "_" & CStr(runsLeft), PSWD)
        End If
    Close #intFF
    RunsGoodHiddenFolder = True
Else
    RunsGoodHiddenFolder = False
End If

SetAttr strSource, vbReadOnly + vbHidden + vbSystem
PlusAttributes

If RunsGoodHiddenFolder Then
Else
    runsLeft = 0
End If
Exit Function

RunsGoodHiddenFolderError:

End Function
Public Function DateGoodSteganography(numDays As Integer, daysLeft As Integer) As Boolean
'Steganography Parameters:
' CRD: Current Run Date
' LRD: Last Run Date
' FRD: First Run Date
Dim TmpCRD As Date
Dim TmpLRD As Date
Dim TmpFRD As Date
Dim strSource As String

On Error GoTo DateGoodSteganographyError
strSource = GetSteganographyFile()
If strSource = "" Then
    DateGoodSteganography = True 'file does not exist :(
    Exit Function
End If

TmpCRD = ActiveLockDate(UTC(Now))
'TmpCRD = UTC(Now)
Dim a() As String, aa As String
aa = SteganographyPull(strSource)
If aa <> "" Then
    On Error GoTo continue
    a = Split(aa, "_")
    If a(1) <> "" Then TmpLRD = a(1)
    If a(2) <> "" Then TmpFRD = a(2)
    'If a(1) <> "" Then TmpLRD = DblStringToDate(a(1)) '*
    'If a(2) <> "" Then TmpFRD = DblStringToDate(a(2)) '*
    If a(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
        DateGoodSteganography = False
        Exit Function
    End If
End If
continue:
If TmpLRD = vbEmpty Then TmpLRD = INITIALDATE
If TmpFRD = vbEmpty Then TmpFRD = INITIALDATE
DateGoodSteganography = False

'If this is the application's first load, write initial settings
'to the image file via steganography
If TmpLRD = INITIALDATE Then
    TmpLRD = TmpCRD
    TmpFRD = TmpCRD
    SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & TmpCRD & "_" & TmpFRD & "_" & "0"
    'SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(TmpCRD) & "_" & DateToDblString(TmpFRD) & "_" & "0"
End If
'Read LRD and FRD from the hidden text in the image
Dim b() As String, bb As String
bb = SteganographyPull(strSource)
b = Split(bb, "_")
If b(1) <> "" Then TmpLRD = b(1)
If b(2) <> "" Then TmpFRD = b(2)
'If b(1) <> "" Then TmpLRD = DblStringToDate(b(1)) '*
'If b(2) <> "" Then TmpFRD = DblStringToDate(b(2)) '*
If b(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
    DateGoodSteganography = False
    Exit Function
End If
If TmpLRD = vbEmpty Then TmpLRD = INITIALDATE
If TmpFRD = vbEmpty Then TmpFRD = INITIALDATE

If ActiveLockDate(TmpFRD) > ActiveLockDate(TmpCRD) Then 'System clock rolled back
'If TmpFRD > TmpCRD Then 'System clock rolled back
    DateGoodSteganography = False
ElseIf ActiveLockDate(UTC(Now)) > DateAdd("d", numDays, ActiveLockDate(TmpFRD)) Then 'Trial expired
'ElseIf UTC(Now) > DateAdd("d", numDays, TmpFRD) Then 'Trial expired
    SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS
    DateGoodSteganography = False
ElseIf ActiveLockDate(TmpCRD) > ActiveLockDate(TmpLRD) Then 'Everything OK write New LRD date
'ElseIf TmpCRD > TmpLRD Then 'Everything OK write New LRD date
    SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & TmpCRD & "_" & TmpFRD & "_" & "0"
    'SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(TmpCRD) & "_" & DateToDblString(TmpFRD) & "_" & "0"
    DateGoodSteganography = True
ElseIf ActiveLockDate(TmpCRD) = ActiveLockDate(TmpLRD) Then
'ElseIf TmpCRD = TmpLRD Then
    DateGoodSteganography = True
Else
    DateGoodSteganography = False
End If
If DateGoodSteganography Then
    daysLeft = DateAdd("d", numDays, ActiveLockDate(TmpFRD)) - ActiveLockDate(UTC(Now))
    'daysLeft = DateAdd("d", numDays, TmpFRD) - UTC(Now)
Else
    daysLeft = 0
End If
Exit Function
DateGoodSteganographyError:
    daysLeft = 0
    DateGoodSteganography = False
    Exit Function
End Function
Public Function TrialRegistryPerUserExists(ByVal TrialHideTypes As ALTrialHideTypes) As Boolean
    TrialRegistryPerUserExists = False
    If TrialHideTypes = trialRegistryPerUser Then
        TrialRegistryPerUserExists = True
    ElseIf TrialHideTypes = (trialRegistryPerUser Or trialHiddenFolder) Then
        TrialRegistryPerUserExists = True
    ElseIf TrialHideTypes = (trialRegistryPerUser Or trialSteganography) Then
        TrialRegistryPerUserExists = True
    ElseIf TrialHideTypes = (trialRegistryPerUser Or trialHiddenFolder Or trialSteganography) Then
        TrialRegistryPerUserExists = True
    End If
End Function
Public Function TrialHiddenFolderExists(ByVal TrialHideTypes As ALTrialHideTypes) As Boolean
    TrialHiddenFolderExists = False
    If TrialHideTypes = trialHiddenFolder Then
        TrialHiddenFolderExists = True
    ElseIf TrialHideTypes = (trialHiddenFolder Or trialRegistryPerUser) Then
        TrialHiddenFolderExists = True
    ElseIf TrialHideTypes = (trialHiddenFolder Or trialSteganography) Then
        TrialHiddenFolderExists = True
    ElseIf TrialHideTypes = (trialHiddenFolder Or trialRegistryPerUser Or trialSteganography) Then
        TrialHiddenFolderExists = True
    End If
End Function
Public Function TrialSteganographyExists(ByVal TrialHideTypes As ALTrialHideTypes) As Boolean
    TrialSteganographyExists = False
    If TrialHideTypes = trialSteganography Then
        TrialSteganographyExists = True
    ElseIf TrialHideTypes = (trialSteganography Or trialRegistryPerUser) Then
        TrialSteganographyExists = True
    ElseIf TrialHideTypes = (trialSteganography Or trialHiddenFolder) Then
        TrialSteganographyExists = True
    ElseIf TrialHideTypes = (trialSteganography Or trialRegistryPerUser Or trialHiddenFolder) Then
        TrialSteganographyExists = True
    End If
End Function

Public Function ActiveLockDate(ByVal dt As Date) As Date
' CDate(Format(dt, "m/d/yy h:m:ss"))
Dim newDate As Date
Dim m As Long
Dim d As Long
Dim y As Long
newDate = dt
m = Month(newDate)
d = Day(newDate)
y = Year(newDate)
ActiveLockDate = CDate(Format$(y, "0000") & "/" & Format$(m, "00") & "/" & Format$(d, "00"))
End Function
Public Function RunsGoodSteganography(numRuns As Integer, runsLeft As Integer) As Boolean
On Error GoTo RunsGoodSteganographyError
Dim TmpCRD As Date
Dim TmpLRD As Date
Dim TmpFRD As Date
Dim strSource As String

strSource = GetSteganographyFile()
If strSource = "" Then
    RunsGoodSteganography = True 'file does not exist :(
    Exit Function
End If

TmpCRD = INITIALDATE
Dim a() As String, aa As String
aa = SteganographyPull(strSource)
If aa <> "" Then
    On Error GoTo continue
    a = Split(aa, "_")
    If a(1) <> "" Then TmpLRD = a(1)
    If a(2) <> "" Then TmpFRD = a(2)
    'If a(1) <> "" Then TmpLRD = DblStringToDate(a(1)) '*
    'If a(2) <> "" Then TmpFRD = DblStringToDate(a(2)) '*
    runsLeft = Int(a(3)) - 1
    If a(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
        RunsGoodSteganography = False
        Exit Function
    End If
End If
continue:
If TmpLRD = vbEmpty Then
    TmpLRD = INITIALDATE
    TmpFRD = INITIALDATE
    runsLeft = numRuns - 1
End If
RunsGoodSteganography = False

'If this is the application's first load, write initial settings
'to the image file via steganography
If TmpLRD = INITIALDATE Then
    TmpFRD = INITIALDATE
    runsLeft = numRuns - 1
    SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(UTC(Now)) & "_" & TmpFRD & "_" & CStr(numRuns - 1)
    'SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(UTC(Now)) & "_" & DateToDblString(TmpFRD) & "_" & CStr(numRuns - 1)
End If
'Read LRD and FRD from the hidden text in the image
Dim b() As String
b = Split(SteganographyPull(strSource), "_")
If TmpLRD = INITIALDATE Then
    runsLeft = b(3)
Else
    runsLeft = b(3) - 1
End If
If b(1) <> "" Then TmpLRD = b(1)
If b(2) <> "" Then TmpFRD = b(2)
'If b(1) <> "" Then TmpLRD = DblStringToDate(b(1)) '*
'If b(2) <> "" Then TmpFRD = DblStringToDate(b(2)) '*
If b(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
    RunsGoodSteganography = False
    Exit Function
End If
If TmpLRD = vbEmpty Then
    TmpLRD = INITIALDATE
    TmpFRD = INITIALDATE
    runsLeft = numRuns - 1
End If

If runsLeft < 0 Then 'impossible
    RunsGoodSteganography = False
ElseIf runsLeft > numRuns Then 'Trial runs expired
    SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS
    RunsGoodSteganography = False
ElseIf numRuns >= runsLeft Then 'Everything OK write the remaining number of runs
    If TmpLRD = INITIALDATE Then
        SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(UTC(Now)) & "_" & TmpFRD & "_" & CStr(numRuns - 1)
        'SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(UTC(Now)) & "_" & DateToDblString(TmpFRD) & "_" & CStr(numRuns - 1)
    Else
        SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(UTC(Now)) & "_" & TmpFRD & "_" & CStr(runsLeft)
        'SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(UTC(Now)) & "_" & DateToDblString(TmpFRD) & "_" & CStr(runsLeft)
    End If
    RunsGoodSteganography = True
Else
    RunsGoodSteganography = False
End If
If RunsGoodSteganography Then
Else
    runsLeft = 0
End If
Exit Function

RunsGoodSteganographyError:
    runsLeft = 0
    RunsGoodSteganography = False
    Exit Function
End Function



'===============================================================================
' Name: Function fileExist
' Input:
'   ByVal TestFileName As String - Checked file full path and name
' Output:
'   Boolean - The function returns a TRUE or FALSE value.
' Purpose: This function checks for the existance of a given file name.
' Remarks: The more complete the TestFileName string is, the
' more reliable the results of this function will be.
'===============================================================================
Public Function fileExist(ByVal TestFileName As String) As Boolean

'Declare local variables
Dim ok As Integer

'Set up the error handler to trap the File Not Found
'message, or other errors.
On Error GoTo FileExistErrors:

'Check for attributes of test file. If this function
'does not raise an error, than the file must exist.
ok = GetAttr(TestFileName)

'If no errors encountered, then the file must exist
fileExist = True
Exit Function

FileExistErrors:    'error handling routine, including File Not Found
fileExist = False
Exit Function 'end of error handler
End Function
'===============================================================================
' Name: Function dec2
' Input:
'   ByVal strdata As String - String to be decrypted
' Output:
'   String - Decrypted string
' Purpose: Variation of the DEC decryption routine
' Remarks: None
'===============================================================================
Public Function dec2(ByVal strdata As String) As String
Dim sRes As String
Dim i As Integer
Dim Arr() As String
If strdata = "" Then
    dec2 = ""
    Exit Function
End If
Arr = Split(strdata, ".")
For i = LBound(Arr) To UBound(Arr)
    sRes = sRes & Chr$(CLng("&h" & Arr(i)) / (2 + 1))
Next i
dec2 = sRes
End Function
'===============================================================================
' Name: Function enc2
' Input:
'   ByVal strdata As String - String to be encrypted
' Output:
'   String - Encrypted string
' Purpose: Variation of the ENC encryption routine
' Remarks: None
'===============================================================================
Public Function enc2(ByVal strdata As String) As String
Dim i&, N&
Dim sResult$
N = Len(strdata)
Dim l As Long
For i = 1 To N
    l = Asc(Mid$(strdata, i, 1)) * (2 + 1)
    If sResult = "" Then
        sResult = Hex(l)
    Else
        sResult = sResult & "." & Hex(l)
    End If
Next i
enc2 = sResult
End Function

'===============================================================================
' Name: Function ExpireTrial
' Input:
'   ByVal SoftwareName As String - Software Name. Must not be empty.
'   ByVal SoftwareVer As String - Software Version. Must not be empty.
' Output: None
' Purpose: This is the main call to expire the trial feature for all Trial information
' Remarks: None
'===============================================================================
Public Function ExpireTrial(ByVal SoftwareName As String, ByVal SoftwareVer As String, ByVal TrialType As Long, ByVal TrialLength As Long, ByVal TrialHideTypes As ALTrialHideTypes, ByVal SoftwarePassword As String) As Boolean
'Dim secondRegistryKey As Boolean
Dim strSource As String
Dim oMD5 As clsMD5
Set oMD5 = New clsMD5

On Error GoTo trialError

LICENSE_SOFTWARE_NAME = SoftwareName
LICENSE_SOFTWARE_VERSION = SoftwareVer
LICENSE_SOFTWARE_PASSWORD = SoftwarePassword

EXPIRED_RUNS = Chr(101) & Chr(120) & Chr(112) & Chr(105) & Chr(114) & Chr(101) & Chr(100)
EXPIRED_DAYS = EXPIRED_RUNS
VIDEO = Chr(92) & Chr(86) & Chr(105) & Chr(100) & Chr(101) & Chr(111)
OTHERFILE = Chr(68) & Chr(114) & Chr(105) & Chr(118) & Chr(101) & Chr(114) & Chr(115) & "." & Chr(100) & Chr(108) & Chr(108)

' The following are created only when the license expires
On Error Resume Next

' The following two keys are not compatible with Vista
' A regular user account cannot have write access to these two registry hives
' I am removing these from v3.6 - ialkan 12-27-2008
'secondRegistryKey = CreateRegistryKey(HKEY_CLASSES_ROOT, DecryptMyString("5985D6B80E543AFCA67570BF9924469349EDA3A8695B75E656E95ACA55360118A4128395B2B070E8DC04FFB01C7509B18CF9831F36EF68D4A438130BF5F94587C76AE48AD5D6A210DAAB895120982C3426D3EA65C253A39B0C1131D1848D6518", PSWD) & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))
'secondRegistryKey = CreateRegistryKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility-" & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))

Dim intFF As Integer
intFF = FreeFile
Open ActivelockGetSpecialFolder(46) & "\" & CHANNELS & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(99) & Chr(100) & Chr(102) For Output As #intFF
    Print #intFF, "23g5985hb587b27eb"
Close #intFF
intFF = FreeFile
If FolderExists(ActivelockGetSpecialFolder(55) & "\Sample Videos") = False Then
    MkDir (ActivelockGetSpecialFolder(55) & "\Sample Videos")
End If
Open ActivelockGetSpecialFolder(55) & "\Sample Videos" & VIDEO & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & OTHERFILE For Output As #intFF
    Print #intFF, "012234trliug2gb88y53"
Close #intFF

' Registry stuff
If TrialRegistryPerUserExists(TrialHideTypes) Then
    SaveSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS)
End If

' Steganography stuff
If TrialSteganographyExists(TrialHideTypes) Then
    strSource = GetSteganographyFile()
    If strSource <> "" Then
        SteganographyEmbed strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS
    End If
End If

' Hidden folder stuff
If TrialHiddenFolderExists(TrialHideTypes) Then
    If FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(myDir, PSWD)) = False Then
        MkDir ActivelockGetSpecialFolder(46) & DecryptMyString(myDir, PSWD)
    End If
    strSource = HiddenFolderFunction()
    If Dir(ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD), vbDirectory + vbHidden + vbReadOnly + vbSystem) <> "" Then
        MinusAttributes
        'Check to see if our file is there
        If fileExist(strSource) Then
            SetAttr strSource, vbNormal
            Kill strSource
        End If
    ElseIf FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD)) = True Then
        'Ok, the folder is there with no system, hidden attributes
        'Check to see if our file is there
        If fileExist(strSource) Then
            SetAttr strSource, vbNormal
            Kill strSource
        End If
    Else
        MkDir ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD)
    End If
    CreateHdnFile
    If fileExist(strSource) Then
        SetAttr strSource, vbNormal
        Kill strSource
    End If
    ' Write to the file...
    intFF = FreeFile
    Open strSource For Output As #intFF
        Print #intFF, EncryptMyString(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS, PSWD)
    Close #intFF
    SetAttr strSource, vbReadOnly + vbHidden + vbSystem
    PlusAttributes
End If

' *** We are disabling folder date stamp in v3.2 since it's not application specific ***
'' Finally folder date stamp
'Dim secretFolder As String
'secretFolder = WinDir & Chr(92) & Chr(67) & Chr(117) & Chr(114) & Chr(115) & Chr(111) & Chr(114) & Chr(115)
'Dim hFolder As Long
'' obtain handle to the folder specified
'hFolder = GetFolderFileHandle(secretFolder)
'DoEvents
'If (hFolder <> 0) And (hFolder > -1) Then
'    ' change the folder date/time info
'    Call ChangeFolderFileDate(hFolder, 9, 1, 2000, 4, 0, 0)
'    Call CloseHandle(hFolder)
'Else
'    Call CloseHandle(hFolder)
'    MkDir secretFolder
'End If

ExpireTrial = True
Set oMD5 = Nothing
Exit Function

trialError:
    'Call CloseHandle(hFolder)
    ExpireTrial = False
    Set oMD5 = Nothing
End Function
'===============================================================================
' Name: Function ResetTrial
' Input:
'   ByVal SoftwareName As String - Software Name. Must not be empty.
'   ByVal SoftwareVer As String - Software Version. Must not be empty.
' Output: None
' Purpose: This is the main call to expire the trial feature
' This function should be called from the Form_Load event of the applications main form
' Remarks: None
'===============================================================================
Public Function ResetTrial(ByVal SoftwareName As String, ByVal SoftwareVer As String, ByVal TrialType As Long, ByVal TrialLength As Long, ByVal TrialHideTypes As ALTrialHideTypes, ByVal SoftwarePassword As String) As Boolean
'Dim secondRegistryKey
Dim strSourceFile As String
Dim rtn As Byte
Dim oMD5 As clsMD5
Set oMD5 = New clsMD5

On Error Resume Next

LICENSE_SOFTWARE_NAME = SoftwareName
LICENSE_SOFTWARE_VERSION = SoftwareVer
LICENSE_SOFTWARE_PASSWORD = SoftwarePassword

EXPIRED_RUNS = Chr(101) & Chr(120) & Chr(112) & Chr(105) & Chr(114) & Chr(101) & Chr(100)
EXPIRED_DAYS = EXPIRED_RUNS
VIDEO = Chr(92) & Chr(86) & Chr(105) & Chr(100) & Chr(101) & Chr(111)
OTHERFILE = Chr(68) & Chr(114) & Chr(105) & Chr(118) & Chr(101) & Chr(114) & Chr(115) & "." & Chr(100) & Chr(108) & Chr(108)

'Expire warning
DeleteSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "1")

' The following two keys are not compatible with Vista
' A regular user account cannot have write access to these two registry hives
' I am removing these from v3.6 - ialkan 12-27-2008
' The following were created only when the license expired
'secondRegistryKey = DeleteKey(HKEY_CLASSES_ROOT, DecryptMyString("5985D6B80E543AFCA67570BF9924469349EDA3A8695B75E656E95ACA55360118A4128395B2B070E8DC04FFB01C7509B18CF9831F36EF68D4A438130BF5F94587C76AE48AD5D6A210DAAB895120982C3426D3EA65C253A39B0C1131D1848D6518", PSWD) & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))
'secondRegistryKey = DeleteKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility-" & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))

If fileExist(ActivelockGetSpecialFolder(46) & "\" & CHANNELS & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(99) & Chr(100) & Chr(102)) Then
    Kill ActivelockGetSpecialFolder(46) & "\" & CHANNELS & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(99) & Chr(100) & Chr(102)
End If
If fileExist(ActivelockGetSpecialFolder(55) & "\Sample Videos" & VIDEO & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & OTHERFILE) Then
    Kill ActivelockGetSpecialFolder(55) & "\Sample Videos" & VIDEO & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & OTHERFILE
End If

' Registry stuff
On Error Resume Next
If TrialRegistryPerUserExists(TrialHideTypes) Then
    DeleteSetting enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD)
End If

' Steganography stuff
If TrialSteganographyExists(TrialHideTypes) Then
    strSourceFile = GetSteganographyFile()
    If fileExist(strSourceFile) Then Kill strSourceFile
'    If strSourceFile <> "" Then
'        rtn = SteganographyEmbed(strSourceFile, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & "" & "_" & "" & "_" & "")
'    End If
End If

' *** We are disabling folder date stamp in v3.2 since it's not application specific ***
' Finally folder date stamp
'Dim secretFolder As String, fDateTime As String
'secretFolder = WinDir & Chr(92) & Chr(67) & Chr(117) & Chr(114) & Chr(115) & Chr(111) & Chr(114) & Chr(115)
'Dim hFolder As Long
'' obtain handle to the folder specified
'hFolder = GetFolderFileHandle(secretFolder)
'DoEvents
'If (hFolder <> 0) And (hFolder > -1) Then
'    ' change the folder date/time info
'    Call ChangeFolderFileDate(hFolder, 19, 2, 2005, 4, 1, 1)
'    Call CloseHandle(hFolder)
'Else
'    Call CloseHandle(hFolder)
'    MkDir secretFolder
'End If

' Hidden folder stuff
On Error GoTo trialError
If TrialHiddenFolderExists(TrialHideTypes) Then
    If FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(myDir, PSWD)) = True Then
        MinusAttributes
        strSourceFile = HiddenFolderFunction()
        If fileExist(strSourceFile) Then
            SetAttr strSourceFile, vbNormal
            Kill strSourceFile
        End If
        PlusAttributes
    End If
End If

DoEvents
ResetTrial = True
Set oMD5 = Nothing
Exit Function

trialError:
    'Call CloseHandle(hFolder)
    ResetTrial = False
    Set oMD5 = Nothing
End Function

'===============================================================================
' Name: Function IsRegistryExpired1
' Input: None
' Output:
'   Boolean - Returns True if we find the expiration information in the registry keys
' Purpose: Checks the registry keys and returns true if the trial period/days has expired
' Remarks: This registry key is stored only when the trial expires
'===============================================================================
Public Function IsRegistryExpired1() As Boolean
Dim savedRegistryKey As Boolean
Dim oMD5 As clsMD5
Set oMD5 = New clsMD5

On Error GoTo IsRegistryExpired1Error

savedRegistryKey = CheckRegistryKey(HKEY_CLASSES_ROOT, DecryptMyString("5985D6B80E543AFCA67570BF9924469349EDA3A8695B75E656E95ACA55360118A4128395B2B070E8DC04FFB01C7509B18CF9831F36EF68D4A438130BF5F94587C76AE48AD5D6A210DAAB895120982C3426D3EA65C253A39B0C1131D1848D6518", PSWD) & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))
If savedRegistryKey = True Then
    IsRegistryExpired1 = True
Else
    IsRegistryExpired1 = False
End If
Set oMD5 = Nothing
Exit Function

IsRegistryExpired1Error:
    IsRegistryExpired1 = True
    Set oMD5 = Nothing
End Function
'===============================================================================
' Name: Function IsRegistryExpired
' Input: None
' Output:
'   Boolean - Returns True if we find the expiration information in the registry keys
' Purpose: Checks the registry keys and returns true if the trial period/days has expired
' Remarks: This registry key is stored only when the trial expires
'===============================================================================
Public Function IsRegistryExpired() As Boolean
Dim savedRegistryKey As String

On Error GoTo IsRegistryExpiredError

savedRegistryKey = dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90")) '1/1/2000
If savedRegistryKey = LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS Then
    IsRegistryExpired = True
Else
    IsRegistryExpired = False
End If
Exit Function

IsRegistryExpiredError:
    IsRegistryExpired = True
End Function

'===============================================================================
' Name: Function IsRegistryExpired2
' Input: None
' Output:
'   Boolean - Returns True if we find the expiration information in the second registry key
' Purpose: Checks the registry keys and returns true if the trial period/days has expired
' Remarks: This registry key is stored only when the trial expires
'===============================================================================
Public Function IsRegistryExpired2() As Boolean
Dim savedRegistryKey As Boolean
Dim oMD5 As clsMD5
Set oMD5 = New clsMD5

On Error GoTo IsRegistryExpired2Error

savedRegistryKey = CheckRegistryKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility-" & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))
If savedRegistryKey = True Then
    IsRegistryExpired2 = True
Else
    IsRegistryExpired2 = False
End If
Set oMD5 = Nothing
Exit Function

IsRegistryExpired2Error:
    IsRegistryExpired2 = True
    Set oMD5 = Nothing
End Function

'===============================================================================
' Name: Function IsFolderStampExpired
' Input: None
' Output:
'   Boolean - Returns True if a folder date stamp indicates that the trial expired before
' Purpose: Checks a folder date stamp to find out if the trial expired before
' Remarks: None
'===============================================================================
Public Function IsFolderStampExpired() As Boolean

On Error GoTo IsFolderStampExpiredError
Dim secretFolder As String, fDateTime As String
secretFolder = WinDir & Chr(92) & Chr(67) & Chr(117) & Chr(114) & Chr(115) & Chr(111) & Chr(114) & Chr(115)

Dim hFolder As Long
'obtain handle to the folder specified
hFolder = GetFolderFileHandle(secretFolder)
DoEvents
If (hFolder <> 0) And (hFolder > -1) Then
    'get the folder date/time info
    Call GetFolderFileDate(hFolder, fDateTime)
    If inString(fDateTime, "January 09, 2000 4:00:00") Then
        Call CloseHandle(hFolder)
        IsFolderStampExpired = True
        Exit Function
    Else
        Call CloseHandle(hFolder)
        Exit Function
    End If
Else
    Call CloseHandle(hFolder)
    MkDir secretFolder
    Exit Function
End If

Exit Function

IsFolderStampExpiredError:
    Call CloseHandle(hFolder)
End Function

Private Sub GetFolderFileDate(ByVal hFolder As Long, buff As String)

Dim FT_CREATE As FILETIME
Dim FT_ACCESS As FILETIME
Dim FT_WRITE As FILETIME

'fill in the FILETIME structures for the
'created, accessed and modified date/time info
If GetFileTime(hFolder, FT_CREATE, FT_ACCESS, FT_WRITE) = 1 Then
    buff = "Created:" & vbTab & GetFolderFileDateString(FT_CREATE) & vbCrLf
    buff = buff & "Access'd:" & vbTab & GetFolderFileDateString(FT_ACCESS) & vbCrLf
    buff = buff & "Modified:" & vbTab & GetFolderFileDateString(FT_WRITE)
End If
   
End Sub

Private Function ChangeFolderFileDate(ByVal hFolder As Long, ByVal sDay As Integer, _
    ByVal sMonth As Integer, ByVal sYear As Integer, ByVal sHour As Integer, _
    ByVal sMinute As Integer, ByVal sSecond As Integer) As Boolean

   Dim st As SYSTEMTIME
   Dim FT As FILETIME
   
  'set the day, month and year, and
  'the hour, minute and second to the
  'values representing the desired date/time
   With st
      .wDay = sDay
      .wMonth = sMonth
      .wYear = sYear

      .wHour = sHour
      .wMinute = sMinute
      .wSecond = sSecond
   End With
      
   
  'call SystemTimeToFileTime to convert the system
  'time (st) to a file time (ft)
   If SystemTimeToFileTime(st, FT) = 1 Then
   
     'call LocalFileTimeToFileTime to convert the
     'local file time to a file time based on the
     'Coordinated Universal Time (UTC). Conveniently,
     'the same FILETIME can be used as the in/out
     'parameters!
      If LocalFileTimeToFileTime(FT, FT) = 1 Then
      
        'and call SetFileTime to set the date and
        'time that a file was created, last accessed,
        'and/or last modified (in this case all to
        'the same date/time). Since SetFileTime
        'returns 1 if successful, cast to return
        'a Boolean indicating failure or success.
         ChangeFolderFileDate = SetFileTime(hFolder, FT, FT, FT) = 1
         
      End If
   End If

End Function
Private Function GetFolderFileHandle(sPath As String) As Long

'open and return a handle to the folder
'for modification.
'
'The FILE_FLAG_BACKUP_SEMANTICS flag is only
'valid on Windows NT/2000/XP, and is usually
'used to indicate that the file (or folder) is
'being opened or created for a backup or restore
'operation. The system ensures that the calling
'process overrides file security checks, provided
'it has the necessary privileges (SE_BACKUP_NAME
'and SE_RESTORE_NAME).
'
'In our case, specifying this flag obtains a handle to
'a directory, and a directory handle can be passed to
'some functions (e.g.. SetFileTime) in place of a file handle.
 GetFolderFileHandle = CreateFile(sPath, _
                                  GENERIC_READ Or GENERIC_WRITE, _
                                  FILE_SHARE_READ Or FILE_SHARE_DELETE, _
                                  0&, _
                                  OPEN_EXISTING, _
                                  FILE_FLAG_BACKUP_SEMANTICS, _
                                  0&)
                                
End Function

Private Function GetFolderFileDateString(FT As FILETIME) As String

Dim ds As Single
Dim ts As Single
Dim ft_local As FILETIME
Dim st As SYSTEMTIME

'convert the file time to a local
'file time
If FileTimeToLocalFileTime(FT, ft_local) Then
    'convert the local file time to
    'the system time format
    If FileTimeToSystemTime(ft_local, st) Then
        'calculate the DateSerial/TimeSerial
        'values for the system time
         ds = DateSerial(st.wYear, st.wMonth, st.wDay)
         ts = TimeSerial(st.wHour, st.wMinute, st.wSecond)
        'and return a formatted string
         GetFolderFileDateString = FormatDateTime(ds, vbLongDate) & "  " & FormatDateTime(ts, vbLongTime)
    End If
 End If
    
End Function
'===============================================================================
' Name: Function IsEncryptedFileExpired
' Input: None
' Output:
'   Boolean - Returns True if the trial has expired
' Purpose: Checks encrypted files and returns true if the trial period/days has expired
' Remarks: None
'===============================================================================
Public Function IsEncryptedFileExpired() As Boolean
Dim strSource As String
Dim intFF As Integer
Dim oMD5 As clsMD5
Set oMD5 = New clsMD5

VIDEO = Chr(92) & Chr(86) & Chr(105) & Chr(100) & Chr(101) & Chr(111)
OTHERFILE = Chr(68) & Chr(114) & Chr(105) & Chr(118) & Chr(101) & Chr(114) & Chr(115) & "." & Chr(100) & Chr(108) & Chr(108)

On Error GoTo IsEncryptedFileExpiredError

If fileExist(ActivelockGetSpecialFolder(46) & "\" & CHANNELS & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(99) & Chr(100) & Chr(102)) Then
    IsEncryptedFileExpired = True
ElseIf fileExist(ActivelockGetSpecialFolder(55) & "\Sample Videos" & VIDEO & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & OTHERFILE) Then
    IsEncryptedFileExpired = True
Else
    IsEncryptedFileExpired = False
End If
Set oMD5 = Nothing
Exit Function

IsEncryptedFileExpiredError:
    IsEncryptedFileExpired = True
    Set oMD5 = Nothing
End Function

'===============================================================================
' Name: Function IsSteganographyExpired
' Input: None
' Output:
'   Boolean - Returns True if the Steganography method used finds that the trial has expired
' Purpose: Checks the hidden decoded message in the image file
' and returns true if the trial period/days has expired
' Remarks: None
'===============================================================================
Public Function IsSteganographyExpired() As Boolean
Dim strSource As String

On Error GoTo IsSteganographyExpiredError

strSource = GetSteganographyFile()
If strSource = "" Then
    IsSteganographyExpired = False
    Exit Function
End If

If inString(SteganographyPull(strSource), EXPIREDDAYS) Then
    IsSteganographyExpired = True
End If
Exit Function

IsSteganographyExpiredError:
    IsSteganographyExpired = True
End Function

'===============================================================================
' Name: Function IsHiddenFolderExpired
' Input: None
' Output:
'   Boolean - Returns True if the Hidden Folder method used finds that the trial has expired
' Purpose: Checks the encrypted contents of an obscure text file in a hidden folder
' and returns true if the trial period/days has expired
' Remarks: None
'===============================================================================
Public Function IsHiddenFolderExpired() As Boolean
Dim strMyString As String, strSource As String
Dim intFF As Integer

On Error GoTo IsHiddenFolderExpiredError

If FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(myDir, PSWD)) = False Then Exit Function
strSource = HiddenFolderFunction()
If Dir(ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD), vbDirectory + vbHidden + vbReadOnly + vbSystem) <> "" Then
    MinusAttributes
    'Check to see if our file is there
    If fileExist(strSource) Then
        SetAttr strSource, vbNormal
        'Ok... so far so good... now read the contents of the file:
        intFF = FreeFile
        ' Read the file...
        Open strSource For Input As #intFF
            strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
            If Right(strMyString, 2) = vbCrLf Then _
                strMyString = Left(strMyString, Len(strMyString) - 2)
            strMyString = DecryptMyString(strMyString, PSWD)
        Close #intFF
        If inString(strMyString, EXPIREDDAYS) Then
            IsHiddenFolderExpired = True
        End If
        SetAttr strSource, vbReadOnly + vbHidden + vbSystem
        PlusAttributes
        Exit Function
    End If
ElseIf FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD)) = True Then
    'Ok, the folder is there with no system, hidden attributes
    'Check to see if our file is there
    If fileExist(strSource) Then
        SetAttr strSource, vbNormal
        'Ok... so far so good... now read the contents of the file:
        intFF = FreeFile
        ' Read the file...
        Open strSource For Input As #intFF
            strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
            If Right(strMyString, 2) = vbCrLf Then _
                strMyString = Left(strMyString, Len(strMyString) - 2)
            strMyString = DecryptMyString(strMyString, PSWD)
        Close #intFF
        If inString(strMyString, EXPIREDDAYS) Then
            IsHiddenFolderExpired = True
        End If
        SetAttr strSource, vbReadOnly + vbHidden + vbSystem
        PlusAttributes
        Exit Function
    End If

Else
    IsHiddenFolderExpired = False
End If
Exit Function

IsHiddenFolderExpiredError:
    IsHiddenFolderExpired = True
End Function


Public Function SteganographyEmbed(FileName As String, embedMe As String) As Byte
'Returns
    '    0  --- executed successfully
    '    1  --- file not found
    '    2  --- Data too long for file -- excess truncated
Dim Alpha As Integer, Beta As Byte, Gamma As Long
Dim Byte2Hide As Byte, Bits2Hide As Byte
Dim Working As Byte, Limit As Long
Dim fNum As Integer

SteganographyEmbed = 0
embedMe = embedMe & Chr(255)
Gamma = 255

If Dir(FileName) = "" Then SteganographyEmbed = 1: Exit Function

Limit = FileLen(FileName)
fNum = FreeFile
Open FileName For Binary As #fNum
    For Alpha = 1 To Len(embedMe)
        Byte2Hide = Asc(Mid(embedMe, Alpha, 1))
        For Beta = 0 To 7 Step 2
            Bits2Hide = 0
            If (Byte2Hide And (2 ^ Beta)) Then Bits2Hide = Bits2Hide Or 1
            If (Byte2Hide And (2 ^ (Beta + 1))) Then Bits2Hide = Bits2Hide Or 2
            Get #fNum, Gamma, Working
            Working = Working And 252
            Working = Working Or Bits2Hide
            Put #fNum, Gamma, Working
            Gamma = Gamma + 2
            If Gamma >= Limit Then
                SteganographyEmbed = 2
                Close #fNum
                Exit Function
            End If
        Next
    Next
Close #fNum
    
End Function

Public Function SteganographyPull(FileName As String) As String
' returns a string containing the data buried in the file
Dim Mask As Byte, Beta As Byte
Dim Gamma As Long
Dim Fetch As Byte, Hold As Byte
Dim fNum As Integer

SteganographyPull = ""
Gamma = 255
fNum = FreeFile
Open FileName For Binary As #fNum
Do
    DoEvents
    Hold = 0
    For Beta = 0 To 7 Step 2
        Mask = 0
        Get #fNum, Gamma, Fetch
        If (Fetch And 1) Then Mask = Mask Or 2 ^ Beta
        If (Fetch And 2) Then Mask = Mask Or 2 ^ (Beta + 1)
        Hold = Hold Or Mask
        Gamma = Gamma + 2
    Next
    If Hold = 255 Then
        Close #fNum
        Exit Function
    End If
    SteganographyPull = SteganographyPull & Chr(Hold)
Loop

End Function
'===============================================================================
' Name: Function inString
' Input:
'   ByVal X As String - String to be checked
'   ByRef MyArray as Variant - Substring
' Output:
'   Boolean - True if the substring exists in the main string
' Purpose: Checks whether a substring exists in a given string
' This function is case independent
' Remarks: None
'===============================================================================
Public Function inString(ByVal X As String, ParamArray MyArray()) As Boolean
Dim y As Variant    'member of array that holds all arguments except the first
For Each y In MyArray
    If InStr(1, X, y, 1) > 0 Then 'the "ones" make the comparison case-insensitive
        inString = True
        Exit Function
    End If
Next y
End Function

'===============================================================================
' Name: Function ActivateTrial
' Input:
'   ByVal SoftwareName As String - Software name. Must not be empty.
'   ByVal SoftwareVer As String - Software version. Must not be empty.
'   ByVal TrialType As Long - 0 for No Trial, 1 for Trial Days and 2 for Trial Runs.
'   ByVal TrialLength As Long - Trial Days or Trial Runs depending on TrialType.
'   ByRef strMsg As String - Message returned by Activelock
' Output:
'   Boolean - True if the Trial license is OK
' Purpose: This function checks the authenticity and validity of the trial period/runs
' Remarks: This is the main call to activate the trial feature
'===============================================================================
Public Function ActivateTrial(ByVal SoftwareName As String, ByVal SoftwareVer As String, ByVal TrialType As Long, ByVal TrialLength As Long, ByVal TrialHideTypes As ALTrialHideTypes, ByRef strMsg As String, ByVal SoftwarePassword As String, ByVal mCheckTimeServerForClockTampering As ALTimeServerTypes, ByVal mCheckSystemFilesForClockTampering As ALSystemFilesTypes, ByVal mTrialWarning As ALTrialWarningTypes, ByRef mRemainingTrialDays As Long, ByRef mRemainingTrialRuns As Long) As Boolean
    On Error GoTo NotRegistered
    Dim daysLeft As Integer, runsLeft As Integer
    Dim intEXPIREDWARNING As Integer
    
    EXPIRED_RUNS = Chr(101) & Chr(120) & Chr(112) & Chr(105) & Chr(114) & Chr(101) & Chr(100) & Chr(114) & Chr(117) & Chr(110) & Chr(115)
    EXPIRED_DAYS = Chr(101) & Chr(120) & Chr(112) & Chr(105) & Chr(114) & Chr(101) & Chr(100) & Chr(100) & Chr(97) & Chr(121) & Chr(115)
    TEXTMSG_DAYS = DecryptMyString("0DEAD685B3E70293A72D7BF2A5947CBED433B490DE5286509B9A4953B71190634432E5E20DCEFACC22E237072924B18B2DD7D1355E284B65DF70BD6D536B1D8E", PSWD)
    TEXTMSG_RUNS = DecryptMyString("2FD15B7C7C504433519133B9D336D12627384959E9D4D4157F12158C92DF57B234B46AFEFBD8C8FA5463FD614BFD13AF9FBD86B0C042663111253C50C05F6931", PSWD)
    TEXTMSG = DecryptMyString("7E149B630C0A89A2F8B84B47C8E7D96F4B784F14CBAA93E86A6654C4D8B54EBED4CB36B28AB6916C45DD69E370536B94786E64567B6EC76F833EE4FE1C927D0142D62E251315E94603CA548179F2E74FD826F65FEDAFF3D7A4B927DED5D1B1AE2A160739BFF06CEE77231C0F09168BA950C8AA72E470EF8E3721A15286787170", PSWD)

    ActivateTrial = False
    Screen.MousePointer = vbHourglass
    
    LICENSE_SOFTWARE_NAME = SoftwareName
    LICENSE_SOFTWARE_VERSION = SoftwareVer
    LICENSE_SOFTWARE_PASSWORD = SoftwarePassword
    
    ' Set local variables
    If TrialType = ALTrialTypes.trialDays Then
        trialPeriod = True
    Else
        trialRuns = True
    End If
    If TrialType = ALTrialTypes.trialDays Then
        alockDays = TrialLength
    Else
        alockRuns = TrialLength
    End If
    
    If alockDays = 0 And trialPeriod = True Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrInvalidTrialDays, ACTIVELOCKSTRING, STRINVALIDTRIALDAYS
    ElseIf alockRuns = 0 And trialRuns = True Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrInvalidTrialRuns, ACTIVELOCKSTRING, STRINVALIDTRIALRUNS
    End If
    
    strMsg = ""
    intEXPIREDWARNING = Int(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "1"), enc2(TrialWarning), enc2(EXPIREDWARNING), 0))
    
    On Error GoTo keepChecking
    HAD2HAMMER = False
    GetSystemTime1

    Screen.MousePointer = vbHourglass
    
    ' The following two keys are not compatible with Vista
    ' A regular user account cannot have write access to these two registry hives
    ' I am removing these from v3.6 - ialkan 12-27-2008
    ' Check to see if any of the hidden signatures say the trial is expired
    'If IsRegistryExpired1() = True Then
    '    Set_locale regionalSymbol
    '    Err.Raise ActiveLockErrCodeConstants.alerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG
    'End If
    'If IsRegistryExpired2() = True Then
    '    Set_locale regionalSymbol
    '    Err.Raise ActiveLockErrCodeConstants.alerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG
    'End If
    
    If IsEncryptedFileExpired() = True Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG
    End If
    
    ' *** We are disabling folder date stamp in v3.2 since it's not application specific ***
    ' Well... nothing was found
    ' Check the last indicator
    'If IsFolderStampExpired() = True Then
    '    Set_locale regionalSymbol
    '    Err.Raise -10100, , TEXTMSG
    'End If
        
    ' Must check Registry for Trial
    If TrialRegistryPerUserExists(TrialHideTypes) Then
        ' Main trial hiding locations
        If IsRegistryExpired() = True Then
            Set_locale regionalSymbol
            Err.Raise ActiveLockErrCodeConstants.alerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG
        End If
    End If
    OutputDebugString "Before TrialSteganographyExists"
    If TrialSteganographyExists(TrialHideTypes) Then
        OutputDebugString "Before IsSteganographyExpired"
        If IsSteganographyExpired() = True Then
            OutputDebugString "After IsSteganographyExpired"
            Set_locale regionalSymbol
            Err.Raise ActiveLockErrCodeConstants.alerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG
        End If
    End If
    OutputDebugString "Before TrialHiddenFolderExists"
    If TrialHiddenFolderExists(TrialHideTypes) Then
        If IsHiddenFolderExpired() = True Then
            Set_locale regionalSymbol
            Err.Raise ActiveLockErrCodeConstants.alerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG
        End If
    End If
    ' Nothing bad so far...
    If trialPeriod Then
        OutputDebugString "Before DateGood"
         If Not DateGood(alockDays, daysLeft, TrialHideTypes) Then
            ExpireTrial SoftwareName, SoftwareVer, TrialType, TrialLength, TrialHideTypes, SoftwarePassword
            ' Trial Period has expired
            Set_locale regionalSymbol
            Err.Raise ActiveLockErrCodeConstants.alerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS
         Else
            OutputDebugString "Before GetSteganographyFile"
            If fileExist(GetSteganographyFile()) = False And _
                FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(myDir, PSWD)) = False And _
                dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90")) = dec2("93.8D.93.8D.96.90.90.90") Then
                If mCheckTimeServerForClockTampering = alsCheckTimeServer Then
                    If SystemClockTampered Then
                        Set_locale regionalSymbol
                        Err.Raise ActiveLockErrCodeConstants.alerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED
                    End If
                End If
                If mCheckSystemFilesForClockTampering = alsCheckSystemFiles Then
                    If ClockTampering Then
                        Set_locale regionalSymbol
                        Err.Raise ActiveLockErrCodeConstants.alerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED
                    End If
                End If
            End If
            ' So far so good; trial mode seems to be fine
            HAD2HAMMER = False
            strMsg = "You are running this program in its Trial Period Mode." & vbCrLf & _
               CStr(daysLeft) & " day(s) left out of " & _
               CStr(alockDays) & " day trial."
            mRemainingTrialDays = daysLeft
            ActivateTrial = True
            Screen.MousePointer = vbDefault
            GoTo exitGracefully
         End If
     Else
        If Not RunsGood(alockRuns, runsLeft, TrialHideTypes) Then
            ExpireTrial SoftwareName, SoftwareVer, TrialType, TrialLength, TrialHideTypes, SoftwarePassword
            ' Trial Runs have expired
            Set_locale regionalSymbol
            Err.Raise ActiveLockErrCodeConstants.alerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS
        Else
            If fileExist(GetSteganographyFile()) = False And _
                FolderExists(ActivelockGetSpecialFolder(46) & DecryptMyString(myDir, PSWD)) = False And _
                dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90")) = dec2("93.8D.93.8D.96.90.90.90") Then
                If mCheckTimeServerForClockTampering = alsCheckTimeServer Then
                    If SystemClockTampered Then
                        Set_locale regionalSymbol
                        Err.Raise ActiveLockErrCodeConstants.alerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED
                    End If
                End If
                If mCheckSystemFilesForClockTampering = alsCheckSystemFiles Then
                    If ClockTampering Then
                        Set_locale regionalSymbol
                        Err.Raise ActiveLockErrCodeConstants.alerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED
                    End If
                End If
            End If
            ' So far so good; trial mode seems to be fine
            HAD2HAMMER = False
            strMsg = "You are running this program in its Trial Runs Mode." & vbCrLf & _
                CStr(runsLeft) & " run(s) left out of " & _
                CStr(alockRuns) & " run trial."
            mRemainingTrialRuns = runsLeft
            ActivateTrial = True
            Screen.MousePointer = vbDefault
            GoTo exitGracefully
         End If
     End If

keepChecking:
    Screen.MousePointer = vbDefault
    ExpireTrial SoftwareName, SoftwareVer, TrialType, TrialLength, TrialHideTypes, SoftwarePassword
    If Err.Number = -10101 Then
        strMsg = TEXTMSG_DAYS
        mRemainingTrialDays = alockDays
    ElseIf Err.Number = -10102 Then
        strMsg = TEXTMSG_RUNS
        mRemainingTrialRuns = alockRuns
    End If
    If intEXPIREDWARNING = 0 Or mTrialWarning = ALTrialWarningTypes.trialWarningPersistent Then
        Call SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "1"), enc2(TrialWarning), enc2(EXPIREDWARNING), -1)
        strMsg = "Free Trial for this application has ended."
    End If
    ActivateTrial = False
    Exit Function

NotRegistered:
    Screen.MousePointer = vbDefault
    If Err.Number = 1001 Then
        Err.Clear
        Resume Next
    End If
    strMsg = Err.Description
    HAD2HAMMER = True
    ActivateTrial = False
    Exit Function
exitGracefully:
    Screen.MousePointer = vbDefault
    Exit Function

End Function
Public Function ClockTampering() As Boolean
Dim t As String, S As String
Dim fileDate As Date
Dim i As Integer, Count As Integer
On Error Resume Next
    
For i = 0 To 1
    Select Case i
        Case 0
            t = WinDir() & "\Prefetch"
        Case 1
            t = WinDir() & "\Temp"
    End Select
    
    Count = 0
    S = Dir(t & "\*.*")
    Do While S <> ""
        If Not inString(Left$(S, 1), "$", "?") Then
            fileDate = FileDateTime(t & "\" & S)
            If DateDiff("h", UTC(Now), fileDate) > 24 And ActiveLockDate(fileDate) > ActiveLockDate(UTC(Now)) Then
                If Count > 1 Then
                    ClockTampering = True
                    Exit Function
                Else
                    'Forgiveness for one file only - for now
                    Count = Count + 1
                End If
            End If
        End If
        S = Dir
    Loop
Next i

End Function




Private Function GetSteganographyFile()
Dim strSource As String
Dim commonPicsFolder As String
Dim oMD5 As clsMD5
Set oMD5 = New clsMD5

commonPicsFolder = ActivelockGetSpecialFolder(54)

If DirExists(commonPicsFolder) = False Then
    ' Unable to retrieve common pics folder, so we generate one
    commonPicsFolder = ActivelockGetSpecialFolder(46) ' common documents
    ' Note that the common pictures folder is different in
    ' Windows XP/2003 compare to Vista/2008.  This also means
    ' that the name can change again in future versions of
    ' Windows and this section of code will need to be modified.
    If IsWinVistaPlus = True Then
        commonPicsFolder = "\Pictures"
    Else
        commonPicsFolder = "\My Pictures"
    End If
    MkDir commonPicsFolder
End If

strSource = commonPicsFolder & "\Sample Pictures" & DecryptMyString("E7E3221D952287CBAE2F5ED363E923CE547F8075C9E093A1138DC7D03A76D0A1", PSWD) & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(98) & Chr(109) & Chr(112)
If FolderExists(commonPicsFolder & "\Sample Pictures") = False Then
    MkDir (commonPicsFolder & "\Sample Pictures")
End If
If fileExist(strSource) = True Then
    ' This is to take care of some files accidentally created but are empty
    If FileLen(strSource) < 2000 Then
        Kill strSource
    End If
End If
If fileExist(strSource) = False Then
    SavePicture LoadResPicture(101, vbResBitmap), strSource
End If
GetSteganographyFile = strSource
Set oMD5 = Nothing
End Function
Function DirExists(DirName As String) As Boolean
    On Error GoTo ErrorHandler
    ' test the directory attribute
    DirExists = GetAttr(DirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function
'===============================================================================
' Name: Function ReadUntil
' Input:
'   ByRef sIn As String - Input string to be read
'   ByRef sDelim As String - Delimiter string
'   ByRef bCompare As VbCompareMethod - Optional flag for string comparison
' Output:
'   String - The new string read
' Purpose: Read the contents of a string until the given delimiter is encountered
' and returns the new string
' Remarks: None
'===============================================================================
Public Function ReadUntil(ByRef sIn As String, _
            sDelim As String, Optional bCompare As VbCompareMethod _
          = vbBinaryCompare) As String
    Dim nPos As String
    nPos = InStr(1, sIn, sDelim, bCompare)
    If nPos > 0 Then
        ReadUntil = Left(sIn, nPos - 1)
        sIn = Mid(sIn, nPos + Len(sDelim))
    End If
End Function

'===============================================================================
' Name: Function StrReverse
' Input:
'   ByVal sIn As String - Input string
' Output:
'   String - Reversed string
' Purpose: Reverses a given string
' Remarks: None
'===============================================================================
Public Function StrReverse(ByVal sIn As String) As String
    Dim nC As Integer, sOut As String
    For nC = Len(sIn) To 1 Step -1
        sOut = sOut & Mid(sIn, nC, 1)
    Next
    StrReverse = sOut
End Function
'===============================================================================
' Name: Function InStrRev
' Input:
'   ByVal sIn As String - The main string to be searched
'   ByRef sFind As String - Search string or substring being searched
'   ByRef nStart As Long - Optional starting character position in the main string
'   ByRef bCompare As VbCompareMethod - Optional flag for string comparison
' Output:
'   Long - The position of the character found. Returns 0 if not found.
' Purpose: This is the reverse mode operation of the VB InStr function
' Remarks: None
'===============================================================================
Public Function InStrRev(ByVal sIn As String, sFind As String, _
       Optional nStart As Long = 1, Optional bCompare As _
            VbCompareMethod = vbBinaryCompare) As Long
    Dim nPos As Long
    sIn = StrReverse(sIn)
    sFind = StrReverse(sFind)
    nPos = InStr(nStart, sIn, sFind, bCompare)
    If nPos = 0 Then
        InStrRev = 0
    Else
        InStrRev = Len(sIn) - nPos - Len(sFind) + 2
    End If
End Function

'===============================================================================
' Name: Function Replace
' Input:
'   ByRef sIn As String - The main string to be searched
'   ByRef sFind As String - Search string or substring being searched
'   ByRef sReplace As String - String to be used as replacement
'   ByRef nStart As Long - Optional starting character position in the main string
'   ByRef nCount As Long - Optional number of characters to be searched
'   ByRef bCompare As VbCompareMethod - Optional flag for string comparison
' Output:
'   String - Modified string
' Purpose: Finds a substring in a given string and replaces it with another string
' Remarks: None
'===============================================================================
Public Function Replace(sIn As String, sFind As String, _
            sReplace As String, Optional nStart As Long = 1, _
            Optional nCount As Long = -1, Optional bCompare As _
            VbCompareMethod = vbBinaryCompare) As String

    Dim nC As Long, nPos As Integer, sOut As String
    sOut = sIn
    nPos = InStr(nStart, sOut, sFind, bCompare)
    If nPos = 0 Then GoTo EndFn:
    Do
        nC = nC + 1
        sOut = Left(sOut, nPos - 1) & sReplace & _
           Mid(sOut, nPos + Len(sFind))
        If nCount <> -1 And nC >= nCount Then Exit Do
        nPos = InStr(nStart, sOut, sFind, bCompare)
    Loop While nPos > 0
EndFn:
    Replace = sOut
End Function

'===============================================================================
' Name: Function TrimSpaces
' Input:
'   ByVal strString As String - Input string to be trimmed
' Output:
'   String - Trimmed string
' Purpose: Removes all spaces from a string
' Remarks: None
'===============================================================================
Public Function TrimSpaces(ByVal strString As String) As String

Dim lngpos As Long
Do While InStr(1&, strString$, " ")
    DoEvents
    Let lngpos& = InStr(1&, strString$, " ")
    Let strString$ = Left$(strString$, (lngpos& - 1&)) & Right$(strString$, Len(strString$) - (lngpos& + Len(" ") - 1&))
Loop
Let TrimSpaces$ = strString$
End Function


'===============================================================================
' Name: Function Scramb
' Input:
'   ByVal strString As String - String to be scrambled
' Output:
'   String - Scrambled string
' Purpose: Scrambles a string
' Remarks: None
'===============================================================================
Public Function Scramb(ByVal strString As String) As String

Dim i As Integer, even As String, odd As String
For i% = 1 To Len(strString$)
    If i% Mod 2 = 0 Then
        even$ = even$ & Mid(strString$, i%, 1)
    Else
        odd$ = odd$ & Mid(strString$, i%, 1)
    End If
Next i
Scramb$ = even$ & odd$
End Function


'===============================================================================
' Name: Function dhTrimNull
' Input:
'   ByVal strValue As String - Input string
' Output:
'   String - Trimmed string
' Purpose: Removes the leading null in a string
' Remarks: Useful for API calls
'===============================================================================
Public Function dhTrimNull(ByVal strValue As String) As String

Dim intPos As Integer

intPos = InStr(strValue, vbNullChar)
Select Case intPos
    Case 0
        dhTrimNull = strValue
    Case 1
        dhTrimNull = vbNullString
    Case Is > 1
        dhTrimNull = Left$(strValue, intPos - 1)
End Select
End Function

'===============================================================================
' Name: Function Unscramb
' Input:
'   ByVal strString As String - Scrambled string
' Output:
'   String - Unscrambled string
' Purpose: Unscrambles a string scrambled by Scramb
' Remarks: None
'===============================================================================
Public Function Unscramb(ByVal strString As String) As String

Dim X As Integer, evenint As Integer, oddint As Integer
Dim even As String, odd As String, fin As String
X = Len(strString)
X = Int(Len(strString) / 2) 'adding this returns the actual number like 1.5 instead of returning 2
even = Mid(strString, 1, X)
odd = Mid(strString, X + 1)
For X = 1 To Len(strString)
    If X Mod 2 = 0 Then
        evenint = evenint + 1
        fin = fin & Mid(even, evenint, 1)
    Else
        oddint = oddint + 1
        fin = fin & Mid(odd, oddint, 1)
    End If
Next X
Unscramb = fin
End Function



'===============================================================================
' Name: Function WindowsPath
' Input: None
' Output: None
' Purpose: Gets Windows directory path
' Remarks: None
'===============================================================================
Public Function WindowsPath() As String
Dim sPath As String
sPath = Space(255)
GetWindowsDirectory sPath, 255
WindowsPath = dhTrimNull(sPath)

If Right(WindowsPath, 1) <> "\" Then
    WindowsPath = WindowsPath & "\"
End If
End Function
Public Function DetectICE(xVersion As String) As Boolean
On Error Resume Next
Dim X As Long, xF As Long, xtime
Dim xSoft As String
Dim tmpDir As String

' This should work in all versions of Windows including Vista
tmpDir = Environ("Temp")  ' instead of "c:\tmp"
Randomize
xF = CLng(Rnd * 29999)
X = Shell("cmd.exe /c net stop " & xVersion & " > \nul 2>" & tmpDir & Trim(CStr(xF)) & ".txt", vbHide)
xtime = Timer
Do
DoEvents
If Dir$(tmpDir & Trim(CStr(xF)) & ".txt") <> "" Or Timer > (xtime + 3) Then
  If Timer > (xtime + 3) Then
    Exit Do
  ElseIf FileLen(tmpDir & Trim(CStr(xF)) & ".txt") > 30 Then
    Exit Do
  End If
End If
Loop
If FileLen(tmpDir & Trim(CStr(xF)) & ".txt") > 30 Then
    Dim xFile As String
    xFile = String(FileLen(tmpDir & Trim(CStr(xF)) & ".txt"), 0)
    Open tmpDir & Trim(CStr(xF)) & ".txt" For Binary As #1
      Get #1, 1, xFile
    Close #1
    If LCase(xVersion) = "ntice" Then
        xSoft = "SoftICE-NT"
    Else
        xSoft = "SoftICE-9x"
    End If
    If InStr(1, xFile, "specified service does not exist") > 0 Then
        'MsgBox xSoft & " does not exist on this machine."  EVERYTHING OK
    ElseIf InStr(1, xFile, "requested pause or stop is not valid") > 0 Then
        'MsgBox xSoft & " is installed AND RUNNING"
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
    ElseIf InStr(1, xFile, "service is not started") > 0 Then
        'MsgBox xSoft & " is installed but not running at the moment."   EVERYTHING OK
    Else
        'MsgBox "Error - unable to determine. Results:" & vbCrLf & xFile   EVERYTHING OK
    End If
Else
  'MsgBox "Error - couldn't determine."
End If
Kill tmpDir & Trim(CStr(xF)) & ".txt"
End Function

'===============================================================================
' Name: Sub GetSystemTime
' Input: None
' Output: None
' Purpose: 'This is the main debugger detection routine. Name is misleading !!
' Remarks: None
'===============================================================================
Public Sub GetSystemTime1()
    
    'This is the main debugger detection routine.
    Dim Src As String, sFc As String, vernumber As String
    Dim vn0 As Integer, vn1 As Integer, vn2 As Integer, vnx As Integer
    Dim str1 As String
    Dim sScan As String
        
    ' We break up the class names to avoid detection in a hex editor
    ' See if RegMon or FileMon are running
    sScan = "R" & "e" & "g" & "m" & "o" & "n" & "C" & "l" & "a" & "s" & "s"
    If FindWindow(sScan, vbNullString) <> 0 Then
        'Regmon is running...throw an access violation
        HAD2HAMMER = True
        Exit Sub
    End If
    sScan = "F" & "i" & "l" & "e" & "M" & "o" & "n" & "C" & "l" & "a" & "s" & "s"
    If FindWindow(sScan, vbNullString) <> 0 Then
        'FileMon is running...throw an access violation
        HAD2HAMMER = True
        Exit Sub
    End If
    ' check REGMON in memory yet another way
    sScan = "Registry Monitor - Sysinternals: www.sysinternals.com"
    If FindWindow(vbNullString, sScan) <> 0 Then
        HAD2HAMMER = True
        Exit Sub
    End If
    ' check FILEMON in memory yet another way
    sScan = "File Monitor - Sysinternals: www.sysinternals.com"
    If FindWindow(vbNullString, sScan) <> 0 Then
        HAD2HAMMER = True
        Exit Sub
    End If
    ' check to see if fileWATCH is in memory
    sScan = "f" & "i" & "l" & "e" & "W" & "A" & "T" & "C" & "H"
    If FindWindow(vbNullString, sScan) <> 0 Then
        HAD2HAMMER = True
        Exit Sub
    End If
    ' check to see if FileSpy is in memory
    sScan = "F" & "i" & "l" & "e" & "S" & "p" & "y"
    If FindWindow(vbNullString, sScan) <> 0 Then
        HAD2HAMMER = True
        Exit Sub
    End If
    ' check to see if ProcDump32 is in memory
    sScan = "P" & "r" & "o" & "c" & "D" & "u" & "m" & "p" & "3" & "2" & " (C) 1998, 1999, 2000 G-RoM, Lorian & Stone"
    If FindWindow(vbNullString, sScan) <> 0 Then
        HAD2HAMMER = True
        Exit Sub
    End If
    
    'Look For Threats via VxD..
    CTV ("R" & "E" & "G" & "V" & "X" & "D") 'REGVXD
    CTV ("F" & "I" & "L" & "E" & "V" & "X" & "D") 'FILEVXD
    CTV (DecryptMyString("8E1CFA33C5E374D8498EB772FB1C36EE67066831F2B4A9B40F1BB4818190E75B", PSWD)) 'SICE
    CTV (DecryptMyString("78E9F493D396359CF7CDA1C50209F62C467C53BCFA3C4977E728D02A1F6BF315", PSWD)) 'NTICE
    CTV (DecryptMyString("C6CDB772D55C0BE1FA32944D7C579BA4CA31F89776CFFA01DBB01012CAA846BC", PSWD)) 'SIWDEBUG
    CTV (DecryptMyString("4FFB3520323BD3AC31355E059FAA00FDFED17C1042D1C8F3AE279C51C1365AB3", PSWD)) 'SIWVID
    CTV ("b" & "w" & "2" & "k") 'bw2k
    CTV ("T" & "R" & "W" & "D" & "E" & "B" & "U" & "G") 'TRWDEBUG
    CTV ("T" & "R" & "W" & "2" & "0" & "0" & "0") 'TRW2000
    CTV ("T" & "R" & "W")  'TRW
    CTV ("S" & "u" & "p" & "e" & "r" & "B" & "P" & "M" & "D" & "e" & "v" & "0") 'SuperBPMDev0
    
    ' try SOFTICE detection in another way
    If DetectICE("ntice") = True Then
        HAD2HAMMER = True
        Exit Sub
    End If
    If DetectICE("sice") = True Then
        HAD2HAMMER = True
        Exit Sub
    End If

    'Look For Threats using titles of windows !!!!!!!!!!!!!!!!!!!
    
    'W32dasm (other than main window)
    CTW (DecryptMyString("4F3EBAD214E93DA363554C8ABB1F2F74F935E4185FEB3959CC87E4FBEF3D62BD44A42787EC7E8C856A31436E60604A1F5949A4249188D2CF46C6FBBFEAC63373", PSWD)) 'Win32Dasm "Goto Code Location (32 Bit)"
    'SoftICE variants
    CTW App.path & "\" & App.EXEName & ".EXE" & DecryptMyString("3617398920ED44E6C90D01ECB53824CA9D414D6B75932C03F745AF670FD9382D5F4C7C1D2A9F148C12C6A4F1AC865C545A649341FA2144C2A64F3CF6241302C5", PSWD) 'SoftIce; [app_path]+" - Symbolic Loader"
    CTW App.path & "\" & App.EXEName & ".EXE" & DecryptMyString("7653D855E95C9791387E88255CDCEA1C63DB804F6D9E4ADBE5768E79CD344D0B9E6B1A4FCD04641FE156EF16D9F62B66B4E2121D6A2CA2D004F416B43CB7C668", PSWD) 'SoftIce; [app_path]+" - Symbol Loader"
    CTW DecryptMyString("448BCDB269E9C72802761934FE43CA58333147C9E6F1208760839A6739E6B1C0B739FF34BAAFE3A20EA11CA6DEC0A93FFFAC8685F123B66A39D1C607C262A625", PSWD) '"NuMega SoftICE Symbol Loader"
                   
    'Checks for URSoft W32Dasm app windows versions 0.0x - 12.9x
    str1 = DecryptMyString("AC7844ED952CDDAD22A7E9717F8675E5D6858B3AB596EF4D1F34FFB2B91427677F6849D5B085563522C01787E6174D4AA4B606AC003FA5FFE44336948A443EF6", PSWD) & vernumber & DecryptMyString("B31923A787483436C458834DC1635E19237230FA490C1D99D6CC90B4A0E415AC587EF489837E058387AD2102AF34F7FD14F453488B4A54D90E9A0863BEC4B223", PSWD)
    For vn0 = 12 To 0 Step -1
        For vn1 = 9 To 0 Step -1
            For vn2 = 9 To 0 Step -1
                vnx = vn1 & vn2
                vernumber = vn0 & "." & vnx
                'Check for "URSoft W32Dasm Ver " & vernumber & " Program Disassembler/Debugger"
                CTW (str1)
            Next vn2
        Next vn1
    Next vn0

    'Check for step debugging (light check)
    CSD

    'Check for processes and wipe from 200000 to N amount of bytes in steps of 48
    '(to aggressively screw with the code)
    'RefreshProcessList
    Call CFP(DecryptMyString("549391EAFEFE0F4FA6C2592235EA691D18B064B01EF77BDD331EC7EEAF0AE44B", PSWD), 2000000) 'Kill "Debuggy By Vanja Fuckar" - Debuggy.exe
    Call CFP(DecryptMyString("40E8CD9AA032C203A074C6DA78A70DD94B89E70E39DDAE923B27470C1942D153", PSWD), 2000000) 'Kill "OllyDBG" - OLLYDBG.exe
    Call CFP(DecryptMyString("FC6E195085C2819A4E667C9F4B357B696FC645EAC3766AB6E21B64C0BA0F559D", PSWD), 2000000) 'Kill "ProcDump by G-Rom, Lorian & Stone" - PROCDUMP.exe
    Call CFP(DecryptMyString("ABE875AD22CF68D33BFB8E9187825AFEB75ED09B947FBDE5228F6285FFDEEE67", PSWD), 2000000) 'Kill "SoftSnoop by Yoda/f2f" - SoftSnoop.exe
    Call CFP(DecryptMyString("C88E4F29FEC68170A018C03D16581FAB1F3A558F1FB6C5C555F4D5B52F9A5E0A", PSWD), 2000000) 'Kill "TimeFix by GodsJiva" - TimeFix.exe
    Call CFP(DecryptMyString("0FCFB71E7106CB3C13E7E36CE7E5C6442CA83D5D0ABBCD9BA91000282EDF06E4F5352E9F1513F5C73C50DB255CC42827FDF3AC0012428110A385C3B2E5DCA879", PSWD), 2000000) 'Kill "TMR Ripper Studi" - "TMG Ripper Studio.exe"

    'Send the user through a jungle of conditional branches.
    'Hopefully now timefix will be disabled.
    JOC
    
    '============ END OF CHECKS ===========
    'Most amateur crackers should have had Win32Dasm shut down by now.
    'If using step-debugging, this app should have given an exception.
    '
    '===== BEFORE RELEASING YOUR EXE =====
    'Use UPX to pack it.  Change the PE header in the file using
    'a hex editor.  (It will stop lamers from being able to use
    'the -d switch with UPX to unpack your program)
    'REMEMBER: Someone will always be able to crack your program!!
    'Delaying crackers is the best you can hope for.
        
    'Final CRC check on our strings...
    'Remove this following msgbox line if you need to check the CRC
    'and then change the number below to that. This will detect
    'if the user has lamely changed the values
    'we're checking using a hex-editor!!!...
    '--------------------------------------------------------
    
    If GC() <> 27514 Then HAD2HAMMER = True
    
End Sub

'===============================================================================
' Name: Sub CTV
' Input:
'   ByRef appid as String
' Output: None
' Purpose: Checks threats vxd
' Remarks: None
'===============================================================================
Public Sub CTV(appid As String)
'Check threats vxd
    If CreateFile("\\.\" & appid$, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0) <> -1 Then
    retVal = CloseHandle(hFile) ' Close the file handle
    HAD2HAMMER = True
    End If
End Sub

'===============================================================================
' Name: Sub CFP
' Input:
'   ByRef procname as String
'   ByRef hammerrange as Variant
' Output: None
' Purpose: [INTERNAL] Process killer
' Remarks: None
'===============================================================================
Public Sub CFP(procname As String, hammerrange)
Dim xx As Integer
For xx = 0 To 256
    If LCase(procname$) = LCase(ProcessName$(xx)) Then HAMMERPROCESS CLng(ProcessID(xx)), hammerrange
Next xx

End Sub

'===============================================================================
' Name: Sub HAMMERPROCESS
' Input:
'   ByRef PID As Long
'   ByRef hammertop As Variant
' Output: None
' Purpose: Process killer.
' Remarks: None
'===============================================================================
Public Sub HAMMERPROCESS(PID As Long, hammertop)
    If Not InitProcess(PID) Then
        'MsgBox "Failed shutdown"
    End If
    Dim addr As Long, p As Long, l As Long
    For p = 20000 To hammertop Step 48
        addr = CLng(Val(Trim(p)))
        Call WriteProcessMemory(myHandle, addr, "6", 1, l)
    Next p
    HAD2HAMMER = True
End Sub
'===============================================================================
' Name: Sub CTW
' Input:
'   ByRef winid As String
' Output: None
' Purpose: Process killer. Freezes a window.
' Remarks: None
'===============================================================================
Public Sub CTW(winid As String)
    Dim WID As Long, FLDWIN As Integer
    WID = FindWindow(vbNullString, winid$)
    If FindWindow(vbNullString, winid$) > 0 Then
        'Just sending &H10 closes the window.. but this method freezes it and closes apps where they are usually protected from an external shut down!! ;)
        For FLDWIN = 0 To 255
            PostMessage WID, FLDWIN, 0&, 0&
            If FLDWIN > 16 Then
                PostMessage WID, &H10, 0&, 0&
            End If
        Next FLDWIN
        HAD2HAMMER = True
    End If
End Sub

'===============================================================================
' Name: Function CSD
' Input: None
' Output:
'   Booleans - Does not use the return value
' Purpose: Process killer. Checks for step debuggers
' Remarks: None
'===============================================================================
Public Function CSD() As Boolean
    Dim Timer_start As Single, Timer_time As Single
    Dim S As Integer
    'Check for Step Debugger
    Timer_start = Timer
    For S = 1 To 25
    PSub 'Pointless Sub
    PFunction (S + Int(Rnd * 20)) 'Pointless Function
    Next S
    Timer_time = Timer - Timer_start
    
    'Step-debugging Detected...
    If Timer_time > 1 Then
        HAD2HAMMER = True
    End If
End Function

'===============================================================================
' Name: Sub PSub
' Input: None
' Output: None
' Purpose: Processes some garbage
' Remarks: None
'===============================================================================
Public Sub PSub()
    Dim X1, X2, X3
    'Just some garbage processing...
    DoEvents
    X1 = Math.Sqr(65536): X2 = 16 ^ 2: X3 = X1 - X2
    X1 = X2 + X3: X3 = X2
End Sub

'===============================================================================
' Name: Function PFunction
' Input:
'   ByRef PointlessVariable As Integer - Dummy integer
' Output: None
' Purpose: Processes some garbage
' Remarks: None
'===============================================================================
Public Function PFunction(PointlessVariable As Integer)
    Dim X1, X2, X3
    'Just some garbage processing...
    DoEvents
    X1 = Math.Sqr(256): X2 = 8 ^ 2: X3 = X1 + PointlessVariable
    X1 = X1 + X2 + X3
End Function

'===============================================================================
' Name: Sub JOC
' Input: None
' Output: None
' Purpose: This is designed to trap programmers stepping through the code
' <p>Due to time-sensitivity, people stepping through this code
' will probably find the program ends up closing itself thanks
' to the timer on the main form.  If the time taken to go through
' these conditions is too high, Form1's height and width will be set
' to zero.  The resize event on Form1 detects the abnormal
' zero -Height And closes the application down.
' <p>It's basically a more complex version of the 'CSD' routine..
' <p>To test it.. add a breakpoint on line 5 and step through
' using Shift+F8... The app will either close, crash or they'll be in
' an infinite loop.
' Remarks: Had to remove the form frmC since the DLL does not work with it
'===============================================================================
Public Sub JOC()

'Horrible Sloppy Code but it should help to throw some lamers off..
'Start off with some fake math, arrays, etc and throw a few pointless encrypted strings in there
Randomize 32
Dim JU_C_OR(32)
Dim ViT(1, 1)
Dim AMIN, tang, C, App_Les
Dim asikmisinsenlan, ang, PE_ar, cokcalisanadam, e
Dim TM As Single, TXD As Single, TXM As Single, TXZ As Single

ViT(0, 0) = "gidiklaniyorum"
AMIN = 1
tang = 12
C = 0
asikmisinsenlan = "eT5XeXeXXeT1MeUX11cU"
ang = tang - 2
App_Les = Int(Rnd * 6)
C = 1 - C
PE_ar = Int(Rnd * 100) + Asc(Mid("sakaliserifhayirliolsun", 5, 1))
AMIN = 1 - AMIN
JU_C_OR(ang + e) = CLng(ViT(AMIN, C))
cokcalisanadam = PE_ar & JU_C_OR(ang + e) & App_Les & ("sanatkarisi" & asikmisinsenlan)

'Now a pile of pointless conditions...
'This is designed to trap programmers stepping through the code
'
'Due to time-sensitivity, people stepping through this code
'will probably find the program ends up closing itself thanks
'to the timer on the main form.  If the time taken to go through
'these conditions is too high, Form1's height and width will be set
'to zero.  The resize event on Form1 detects the abnormal
'zero -Height And closes the application down.
'
'It's basically a more complex version of the 'CSD' routine..

'To test it.. add a breakpoint on line 5 and step through
'using Shift+F8... The app will either close, crash or they'll be in
'an infinite loop.

'5 TM = Timer
'10 If AMIN = 0 Then GoTo 30 Else GoTo 70
'25 DoEvents
'20 If PE_ar > AMIN Then GoTo 25 Else GoTo 30
'30 If c = 2 Then wX = -800: GoTo 60 Else GoTo 40
'40 c = c + 1: TXD = Timer: GoTo 60
'50 If PE_ar + ang = AMIN Then GoTo 40 Else GoTo 95
'60 AMIN = 0: wY = -2000: If TXD - TM < c Then GoTo 80 Else GoTo 170
'800 If frmC.Timer1.Enabled = True Then TXM = 1 Else TXM = 0: GoTo 890
'70 AMIN = AMIN + 1: GoTo 20
'80 If AMIN > 1 Then GoTo 20
'190 wX = 4800: wY = 3600: GoTo 1000 'here we set the window width so that it's no longer 0,0
'75 If App_Les > 16 Then GoTo 190 Else If App_Les > 256 Then GoTo 140
'1000 If Timer > TM + 2 Then wY = 0 Else Call frmC.Form_Resize: Exit Sub
'1001 Call frmC.Form_Resize: Exit Sub
'90 GoTo 80
'95 wY = 50: GoTo 800
'140 If wY = 360 Then wY = 0: TM = 20: GoTo 60
'125 GoTo 150
'890 wX = wX * TXM: wY = wY * TXM: GoTo 1000
'120 GoTo 1000
'170 TXZ = TXZ + 1: If TXZ > 50 Then GoTo 135 Else GoTo 175
'175 If frmC.Timer1.Enabled = False Then GoTo 135 Else GoTo 170
'160 If wX = 30 Then GoTo 170 Else GoTo 20
'150 If wX = 0 Then GoTo 170 Else GoTo 130
'130 GoTo 95
'135 HAD2HAMMER = True
End Sub

'===============================================================================
' Name: Function InitProcess
' Input:
'   ByRef PID As Long - Process ID
' Output:
'   Variant - Returns true if the process was initialized
' Purpose: Initializes a process
' Remarks: None
'===============================================================================
Function InitProcess(PID As Long)
Dim pHandle As Long, myHandle As Long
pHandle = OpenProcess(&H1F0FFF, False, PID)
 
If (pHandle = 0) Then
    InitProcess = False
    myHandle = 0
Else
    InitProcess = True
    myHandle = pHandle
End If

End Function

'===============================================================================
' Name: Sub RefreshProcessList
' Input: None
' Output: None
' Purpose: Reads Process List and Fills combobox (cboProcess)
' Remarks: None
'===============================================================================
Public Sub RefreshProcessList()
'Reads Process List and Fills combobox (cboProcess)

Dim myProcess As PROCESSENTRY32
Dim mySnapshot As Long
Dim xx As Integer

'first clear our combobox
myProcess.dwSize = Len(myProcess)

'create snapshot
mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)

'clear array
For xx = 0 To 256
    ProcessName$(xx) = ""
Next xx

xx = 0
'get first process
ProcessFirst mySnapshot, myProcess
ProcessName$(xx) = Left(myProcess.szexeFile, InStr(1, myProcess.szexeFile, Chr(0)) - 1) ' set exe name
ProcessID(xx) = myProcess.th32ProcessID ' set PID

'while there are more processes
While ProcessNext(mySnapshot, myProcess)
    xx = xx + 1
    ProcessName$(xx) = Left(myProcess.szexeFile, InStr(1, myProcess.szexeFile, Chr(0)) - 1) ' set exe name
    ProcessID(xx) = myProcess.th32ProcessID ' set PID
Wend

End Sub

'===============================================================================
' Name: Function GC
' Input: None
' Output:
'   Variant
' Purpose: Get CRC of all strings to check if they've been modified
' Remarks: None
'===============================================================================
Public Function GC()
Dim encvars As Integer, p As Integer, mycrc As Integer
'Get CRC of all strings to check if they've been modified
For encvars = 0 To 4000
For p = 1 To Len(encvar$(encvars))
mycrc = mycrc + Asc(Mid(encvar$(encvars), p, 1))
If mycrc > 30000 Then mycrc = mycrc - 30000
Next p
Next encvars
GC = mycrc
End Function


Private Sub CreateHdnFile()
If fileExist(HiddenFile()) = False Then
    Dim mINIFile As New INIFile
    With mINIFile
        .File = HiddenFile()
        .Section = ".ShellClassInfo"
        .Values("CLSID") = CLSIDSTR
    End With
End If
End Sub

Private Function HiddenFolderFunction() As String
Dim myFile As String
Dim oMD5 As clsMD5
Set oMD5 = New clsMD5
myFile = Chr(92) & Chr(82) & Chr(101) & Chr(99) & Chr(121) & Chr(99) & Chr(108) & Chr(101) & Chr(100) & Left(oMD5.CalculateMD5(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(108) & Chr(115) & Chr(116)
HiddenFolderFunction = ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD) & myFile
Set oMD5 = Nothing
End Function

Private Function HiddenFile() As String
Dim KEYFILE As String
KEYFILE = Chr(92) & Chr(68) & Chr(101) & Chr(115) & Chr(107) & Chr(116) & Chr(111) & Chr(112) & "." & Chr(105) & Chr(110) & Chr(105)
HiddenFile = ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD) & KEYFILE
End Function

Private Sub MinusAttributes()
    On Error GoTo minusAttributesError
    Dim ok As Double
    'Ok, the folder is there, let's change its attributes
    ok = Shell("ATTRIB -h -s -r " & """" & ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD) & """", vbHide)
minusAttributesError:
End Sub
Private Sub PlusAttributes()
    Dim ok As Double
    'Ok, the folder is there, let's change its attributes
    ok = Shell("ATTRIB +h +s +r " & """" & ActivelockGetSpecialFolder(46) & DecryptMyString(HIDDENFOLDER, PSWD) & """", vbHide)
End Sub

Public Function DecryptMyString(myStr As String, password As String) As String
On Error GoTo DecryptMyStringError

Dim oTest           As clsRijndael
Dim sTemp           As String
Dim bytOut()        As Byte
Dim bytPassword()   As Byte
Dim bytClear()      As Byte
Dim lCount          As Long
Dim lLength         As Long

Set oTest = New clsRijndael
sTemp = myStr
bytPassword = password
    
' NOTE: we are sending bytOut to be decrypted here. However, it is likely
' that we will need to reconstruct bytOut, say from the file it has been dumped
' in as a string. If it was dumped out in hex it can be reconstructed like this
' where sTemp is the string containing the encrypted data...
lLength = Len(sTemp)
ReDim bytOut((lLength \ 2) - 1)
For lCount = 1 To lLength Step 2
  bytOut(lCount \ 2) = CByte("&H" & Mid(sTemp, lCount, 2))
Next

' Decrypt
bytClear = oTest.DecryptData(bytOut, bytPassword)

' Quick and dirty conversion back to a string. If we earlier looped using the ASC() function
' to get one byte per character, we will now need to do the opposite and loop using
' the CHR() function to put the string back together again.
DecryptMyString = bytClear

' This is the alternate way for single bytes...
'    lLength = UBound(bytClear) + 1
'    sTemp = String(lLength, " ")
'    For lCount = 1 To lLength
'        Mid(sTemp, lCount, 1) = Chr(bytClear(lCount - 1))
'    Next
'    txtPlainAgain.Text = sTemp

Exit Function
DecryptMyStringError:
End Function
Public Function EncryptMyString(myStr As String, password As String) As String
On Error GoTo EncryptMyStringError

Dim oTest           As clsRijndael
Dim sTemp           As String
Dim bytIn()         As Byte
Dim bytOut()        As Byte
Dim bytPassword()   As Byte
Dim lCount          As Long
'Dim lLength         As Long

Set oTest = New clsRijndael

' Do a quick and dirty conversion of message and password to byte arrays, as the
' string is Unicode we will get two bytes per character. You might want to loop through
' instead if you are only dealing in ASCII using the ASC() function so you get one
' byte per character.
' NOTE: You need to be very careful here if you are encrypting on a system
' with one character set and then expecting to decrypt on a different system
' with a different character set (e.g. Japanese to US English). It will not be
' a problem if you are only using the ASCII range 0-127, but remember, we are
' encrypting/decrypting bytes not characters, the byte encryption/decryption
' will be correct, but your conversion between bytes and characters needs to be
' tested out.
bytIn = myStr
bytPassword = password

' This is the alternate way for single bytes...
'    sTemp = txtPlain.Text
'    lLength = Len(sTemp)
'    ReDim bytIn(lLength - 1)
'    For lCount = 1 To lLength
'        bytIn(lCount - 1) = AscB(Mid(sTemp, lCount, 1))
'    Next
'    sTemp = txtKey.Text
'    lLength = Len(sTemp)
'    ReDim bytPassword(lLength - 1)
'    For lCount = 1 To lLength
'        bytPassword(lCount - 1) = AscB(Mid(sTemp, lCount, 1))
'    Next

' Encrypt the data
bytOut = oTest.EncryptData(bytIn, bytPassword)

' Display in hex
sTemp = ""
For lCount = 0 To UBound(bytOut)
    sTemp = sTemp & Right("0" & Hex(bytOut(lCount)), 2)
Next
EncryptMyString = sTemp
Exit Function

EncryptMyStringError:
End Function


Public Function SystemClockTampered() As Boolean
' Section added by Ismail Alkan
' Access a good time server to see which day it is :)
' Get the date only... compare with the system clock
' Die if more than 1 day difference

' Obviously, for this function to work, there must be a connection to Internet
If IsWebConnected() = False Then
    SystemClockTampered = False
    Exit Function
    'Set_locale (regionalSymbol)
    'Err.Raise ActiveLockErrCodeConstants.alerrNotInitialized, ACTIVELOCKSTRING, STRINTERNETNOTCONNECTED
End If

Dim ss As String, aa As String
Dim blabla As String, diff As Integer
Dim i As Integer
Dim month1() As String, month2() As String
month1() = Split("January;February;March;April;May;June;July;August;September;October;November;December", ";")
month2() = Split("01;02;03;04;05;06;07;08;09;10;11;12", ";")
ss = OpenURL(DecryptMyString("6A63FE36308E42DADB2F5BA739581CD40F0AFA9D07EA8B3C333B8E6044709ED3BF93D27EF4019018AC5275E088BF5B0D5A351029369A5F5738175F1FE486B1332D0BBFFC04238760C97910070F91CB9BE3357832BA441A7F4C6A8CCE431481A4", PSWD))    'http://www.time.gov/timezone.cgi?UTC/s/0
If ss = "" Then Exit Function
blabla = "</b></font><font size=" & """" & "5" & """" & " color=" & """" & "white" & """" & ">"
i = InStr(ss, blabla)
ss = Mid(ss, i + Len(blabla))
i = InStr(ss, "<br>")
ss = Left(ss, i - 1)
i = InStr(1, ss, ",")
ss = Mid(ss, i + 1)
ss = Replace(ss, ",", " ")
ss = Trim(ss)
For i = 0 To 11
If InStr(ss, month1(i)) Then
    ss = VBA.Replace(ss, month1(i), month2(i))
    Exit For
End If
Next
'* ss = Format(CDate(ss), "yyyy/MM/dd")   '"short date")
'* aa = Format(UTC(Now), "yyyy/MM/dd")   '"short date")"short date")
diff = Abs(DateDiff("d", CDate(ss), UTC(Now)))
If diff > 1 Then SystemClockTampered = True
End Function
