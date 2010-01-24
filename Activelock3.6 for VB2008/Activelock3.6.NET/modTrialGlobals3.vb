Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text
Imports System.String
Imports System.DateTime
Imports System.Runtime.InteropServices
Imports Microsoft.VisualBasic.ControlChars

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
#End Region

Module modTrial

    '===============================================================================
    ' Name: modTrialGlobals
    ' Purpose: This module is used by the Trial Period/Runs feature
    ' Functions:
    ' Properties:
    ' Methods:
    ' Started: 08.15.2005
    ' Modified: 03.25.2006
    '===============================================================================

    Public EXPIRED_RUNS As String
    Public EXPIRED_DAYS As String
    Public LICENSE_SOFTWARE_NAME As String
    Public LICENSE_SOFTWARE_PASSWORD As String
    Public LICENSE_SOFTWARE_CODE As String
    Public LICENSE_SOFTWARE_VERSION As String
    Public LICENSE_SOFTWARE_LOCKTYPE As String
    Public alockDays, alockRuns As Short
    Public trialPeriod, trialRuns As Boolean
    Public TEXTMSG_RUNS, TEXTMSG_DAYS, TEXTMSG As String
    Public VIDEO, OTHERFILE As String

    Public StegInfo As String

    Public Const HIDDENFOLDER As String = "SM8YnnHzkjsvBayVJjIexcUpH5+7aO1WosnkqOTm8ZU="
    Public Const EXPIREDDAYS As String = "ExpiredDays"
    Public Const INITIALDATE As String = "01/01/2000"
    Public Const TRIALWARNING As String = "Trial Warning"
    Public Const EXPIREDWARNING As String = "Expired Warning"
    Public Const CLSIDSTR As String = "{645FF040-5081-101B-9F08-00AA002F954E}"
    Public Const CHANNELS As String = "CTDChannels_Version."

    Public Const REG_MULTI_SZ As Short = 7
    Public Const ERROR_MORE_DATA As Short = 234

    ' Windows Security Messages
    Const KEY_ALL_CLASSES As Integer = &HF0063

    Const ERROR_SUCCESS As Short = 0
    Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Integer

    ' Windows Registry API calls
    Public Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByVal lpValue As String, ByVal cbData As Integer) As Integer

    Public Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByRef lpValue As Integer, ByVal cbData As Integer) As Integer
    Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Integer) As Integer
    Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
    Public Declare Function RegOpenKey Lib "advapi32.dll" (ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
    Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
    Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Integer, ByRef lpcbData As Integer) As Integer
    Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
    Public Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Integer, ByRef lpcbData As Integer) As Integer
    Public Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As Integer, ByRef lpcbData As Integer) As Integer
    Public Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Integer) As Integer

    Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Integer) As Integer
    Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Integer) As Integer
    Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Integer, ByVal dwBytes As Integer) As Integer
    Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Integer) As Integer

    Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
    Private Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Integer, ByRef lpFindFileData As WIN32_FIND_DATA) As Integer
    Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Integer) As Integer
    Private Declare Function SearchPath Lib "kernel32" Alias "SearchPathA" (ByVal lpPath As String, ByVal lpFileName As String, ByVal lpExtension As String, ByVal nBufferLength As Integer, ByVal lpBuffer As String, ByVal lpFilePart As String) As Integer

    Private Const MAX_PATH As Short = 260



    Private Structure FILETIME
        Dim lngLowDateTime As Integer
        Dim lngHighDateTime As Integer
    End Structure

    Private Structure WIN32_FIND_DATA
        Dim lngFileAttributes As Integer ' File attributes
        Dim ftCreationTime As FILETIME ' Creation time
        Dim ftLastAccessTime As FILETIME ' Last access time
        Dim ftLastWriteTime As FILETIME ' Last modified time
        Dim lngFileSizeHigh As Integer ' Size (high word)
        Dim lngFileSizeLow As Integer ' Size (low word)
        Dim lngReserved0 As Integer ' reserved
        Dim lngReserved1 As Integer ' reserved
        <VBFixedString(MAX_PATH), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> Public strFilename As String ' File name
        <VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=14)> Public strAlternate As String ' 8.3 name
    End Structure

    Private Structure SECURITY_ATTRIBUTES
        Dim nLength As Integer
        Dim lpSecurityDescriptor As Object
        Dim bInheritHandle As Boolean
    End Structure

    Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, ByRef lpVolumeSerialNumber As Integer, ByRef lpMaximumComponentLength As Integer, ByRef lpFileSystemFlags As Integer, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Integer) As Integer

    'Some common variables
    Public sWinPath As String
    Public sLocalKey As String
    Public sMainFile As String

    'Error variables
    Public bTampered As Boolean
    Public bAccessDenied As Boolean
    Public bFatal As Boolean
    Public intProgress As Short

    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByVal lpSecurityAttributes As Integer, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As Integer) As Integer

    Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Integer) As Integer
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Integer, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer

    'Spy Scan stuff

    Public Const CREATE_NEW As Short = 1
    Public Const CREATE_ALWAYS As Short = 2
    Public Const OPEN_EXISTING As Short = 3
    Public Const OPEN_ALWAYS As Short = 4
    Public Const TRUNCATE_EXISTING As Short = 5

    Private Const FILE_BEGIN As Short = 0
    Private Const FILE_CURRENT As Short = 1
    Private Const FILE_END As Short = 2

    Private Const FILE_FLAG_WRITE_THROUGH As Integer = &H80000000
    Private Const FILE_FLAG_OVERLAPPED As Integer = &H40000000
    Private Const FILE_FLAG_NO_BUFFERING As Integer = &H20000000
    Private Const FILE_FLAG_RANDOM_ACCESS As Integer = &H10000000
    Private Const FILE_FLAG_SEQUENTIAL_SCAN As Integer = &H8000000
    Private Const FILE_FLAG_DELETE_ON_CLOSE As Integer = &H4000000
    Private Const FILE_FLAG_BACKUP_SEMANTICS As Integer = &H2000000
    Private Const FILE_FLAG_POSIX_SEMANTICS As Integer = &H1000000

    ' Called from frmC
    Public Declare Function FinWin Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function CF Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByRef lpSecurityAttributes As Integer, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As Integer) As Integer
    Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
    Public Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Integer
    Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, ByRef lParam As Integer) As Integer
    Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Integer, ByRef lpdwProcessId As Integer) As Integer
    Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Integer, ByVal bInheritHandle As Integer, ByVal dwProcessID As Integer) As Integer
    Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddress As Long, ByVal lpBuffer As Integer, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
    Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Integer, ByVal lpBaseAddress As Long, ByVal lpBuffer As Long, ByVal nSize As Integer, ByRef lpNumberOfBytesWritten As Integer) As Integer
    Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, ByRef lpSource As Long, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As String, ByVal nSize As Integer, ByRef Arguments As Integer) As Integer
    Public Declare Function GetLastError Lib "kernel32" () As Integer

    Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Short = &H100S
    Public Const FORMAT_MESSAGE_FROM_SYSTEM As Short = &H1000S

    'Structure PROCESSENTRY32 may require marshalling attributes to be passed as an argument in this Declare statement
    Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Integer, ByRef uProcess As PROCESSENTRY32) As Integer
    'Structure PROCESSENTRY32 may require marshalling attributes to be passed as an argument in this Declare statement
    Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Integer, ByRef uProcess As PROCESSENTRY32) As Integer
    Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Integer, ByRef lProcessID As Integer) As Integer

    Public Const TH32CS_SNAPPROCESS As Integer = 2

    Structure PROCESSENTRY32
        Dim dwSize As Integer
        Dim cntUsage As Integer
        Dim th32ProcessID As Integer
        Dim th32DefaultHeapID As Integer
        Dim th32ModuleID As Integer
        Dim cntThreads As Integer
        Dim th32ParentProcessID As Integer
        Dim pcPriClassBase As Integer
        Dim dwFlags As Integer
        <VBFixedString(260), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=260)> Public szexeFile As String
    End Structure

    Public Const GENERIC_WRITE As Integer = &H40000000
    Public Const GENERIC_READ As Integer = &H80000000
    Public Const GENERIC_EXECUTE As Integer = &H20000000
    Public Const GENERIC_ALL As Integer = &H10000000
    Public Const FILE_SHARE_READ As Short = &H1S
    Public Const FILE_SHARE_WRITE As Short = &H2S
    Public Const FILE_SHARE_DELETE As Integer = &H4S
    Public Const FILE_ATTRIBUTE_NORMAL As Short = &H80S
    Public Const FILE_ATTRIBUTE_READONLY As Short = &H1S
    Public Const FILE_ATTRIBUTE_HIDDEN As Short = &H2S
    Public Const FILE_ATTRIBUTE_SYSTEM As Short = &H4S
    Public Const FILE_ATTRIBUTE_DIRECTORY As Short = &H10S
    Public Const FILE_ATTRIBUTE_ARCHIVE As Short = &H20S
    Public Const FILE_ATTRIBUTE_TEMPORARY As Short = &H100S
    Public Const FILE_ATTRIBUTE_COMPRESSED As Short = &H800S

    Public Const EAV As Integer = &HC0000005

    Public ProcessName(256) As String
    Public ProcessID(256) As Integer
    Public retVal, hFile, TimerStart As Integer
    Public wX, wY As Integer
    Public myHandle As Object
    'Public Buffer As String
    Public varchk As Object
    Public encvar(4000) As String
    Public HAD2HAMMER As Boolean

    ' Folder Date Stamp API
    Private Structure SYSTEMTIME
        Dim wYear As Short
        Dim wMonth As Short
        Dim wDayOfWeek As Short
        Dim wDay As Short
        Dim wHour As Short
        Dim wMinute As Short
        Dim wSecond As Short
        Dim wMilliseconds As Integer
    End Structure

    'Structure FILETIME may require marshalling attributes to be passed as an argument in this Declare statement
    'Structure SYSTEMTIME may require marshalling attributes to be passed as an argument in this Declare statement
    Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Integer, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Integer
    Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Integer
    Private Declare Function FileTimeToSystemTime Lib "kernel32" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Integer
    Private Declare Function SystemTimeToFileTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Integer
    Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (ByRef lpLocalFileTime As FILETIME, ByRef lpFileTime As FILETIME) As Integer
    Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Integer, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Integer

    'Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Integer, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Integer) As Integer
    'Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Integer, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Integer, ByVal lFlags As Integer, ByVal lContext As Integer) As Integer
    'Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Integer, ByVal sBuffer As String, ByVal lNumberOfBytesToRead As Integer, ByRef lNumberOfBytesRead As Integer) As Integer
    'Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Integer) As Integer

    <DllImport("WinInet.dll", _
EntryPoint:="InternetOpenA", _
CharSet:=CharSet.Ansi, ExactSpelling:=True, SetLastError:=True)> _
Public Function InternetOpen( _
ByVal agent As String, _
ByVal accessType As Int32, _
ByVal proxyName As String, _
ByVal proxyBypass As String, _
ByVal flags As Int32) As IntPtr
    End Function


    <DllImport("WinInet.dll", _
    EntryPoint:="InternetOpenUrlA", _
    CharSet:=CharSet.Ansi, ExactSpelling:=True, SetLastError:=True)> _
    Public Function InternetOpenUrl( _
    ByVal session As IntPtr, _
    ByVal url As String, _
    ByVal header As String, _
    ByVal headerLength As Int32, _
    ByVal flags As Int32, _
    ByVal context As Int32) As Int32
    End Function


    'InternetReadFile
    <DllImport("WinInet.dll", _
    EntryPoint:="InternetReadFile", _
    CharSet:=CharSet.Auto, SetLastError:=True)> _
    Public Function InternetReadFile( _
    ByVal handle As Int32, _
    <MarshalAs(UnmanagedType.LPArray)> _
    ByVal newBuffer() As Byte, _
    ByVal bufferLength As Int32, _
    ByRef bytesRead As Int32) As Int32
    End Function

    <DllImport("WinInet.dll", _
    EntryPoint:="InternetCloseHandle", _
    CharSet:=CharSet.Ansi, ExactSpelling:=True, SetLastError:=True)> _
    Public Function InternetCloseHandle( _
    ByVal hInternet As Int32) As Int32
    End Function


    Public Function OpenURL(ByVal sUrl As String) As String
        Const INTERNET_ACCESS_TYPE_DIRECT As Integer = 1
        Const INTERNET_FLAG_RELOAD As Integer = &H80000000
        Const USER_AGENT As String = "IE"

        Dim length As Int32 = 2048
        Dim handle As IntPtr
        Dim session As Int32
        Dim header As String = "Accept: */*" & Cr & Cr
        Dim newBuffer() As Byte
        Dim bytesRead As Int32
        Dim response As Int32
        Dim context As Integer = 0
        Dim flags As Integer = 0
        Dim result As String
        handle = InternetOpen(USER_AGENT, INTERNET_ACCESS_TYPE_DIRECT, vbNullString, vbNullString, flags)
        session = InternetOpenUrl(handle, sUrl, header, header.Length, INTERNET_FLAG_RELOAD, context)
        If session = 0 Then
            result = "Error: " & Marshal.GetLastWin32Error()
        Else
            ReDim newBuffer(length - 1)
            response = InternetReadFile(session, newBuffer, length, bytesRead)
            If response = 0 Then
                result = ""   '"Error Reading File: " & Marshal.GetLastWin32Error()
            Else
                ' Use appropriate Encoding here to get string from byte array
                result = System.Text.UTF8Encoding.UTF8.GetString(newBuffer)
            End If
        End If
        OpenURL = result
        InternetCloseHandle(session)
        InternetCloseHandle(handle.ToInt32)
        Return result

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
    Public Function DateGoodRegistry(ByRef numDays As Short, ByRef daysLeft As Short) As Boolean
        'Ex: If DateGoodRegistry(30)=False Then
        ' CrippleApplication
        ' End if
        'Registry Parameters:
        ' CRD: Current Run Date
        ' LRD: Last Run Date
        ' FRD: First Run Date
        Dim TmpCRD As System.DateTime
        Dim TmpLRD As System.DateTime
        Dim TmpFRD As System.DateTime

        On Error GoTo DateGoodRegistryError

        TmpCRD = Date.UtcNow '* TmpCRD = ActiveLockDate(Date.UtcNow)
        '*TmpLRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        '*TmpFRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        TmpLRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        TmpFRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        DateGoodRegistry = False

        'If this is the applications first load, write initial settings
        'to the registry
        If TmpLRD = CDate(dec2("93.8D.93.8D.96.90.90.90")) Then '1/1/2000
            '*  SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(CStr(TmpCRD)))
            '* SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", enc2(CStr(TmpCRD)))
            SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(DateToDblString(TmpCRD))) '*
            SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", enc2(DateToDblString(TmpCRD))) '*
            SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2("0"))
        End If
        'Read LRD and FRD from registry
        '*   TmpLRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        '*  TmpFRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        TmpLRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        TmpFRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000

        '* If ActiveLockDate(TmpFRD) > ActiveLockDate(TmpCRD) Then 'System clock rolled back
        If TmpFRD > TmpCRD Then 'System clock rolled back
            DateGoodRegistry = False
            '*        ElseIf ActiveLockDate(Date.UtcNow) > ActiveLockDate(TmpFRD).AddDays(numDays) Then  'trial expired
        ElseIf Date.UtcNow > (TmpFRD).AddDays(numDays) Then  'trial expired
            SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS))
            DateGoodRegistry = False
            '* ElseIf ActiveLockDate(TmpCRD) > ActiveLockDate(TmpLRD) Then  'Everything OK write New LRD date
            '*   SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(CStr(TmpCRD)))
        ElseIf TmpCRD > TmpLRD Then  'Everything OK write New LRD date
            SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(DateToDblString(TmpCRD)))
            DateGoodRegistry = True
            '* ElseIf ActiveLockDate(TmpCRD) = ActiveLockDate(TmpLRD) Then
        ElseIf TmpCRD = TmpLRD Then
            DateGoodRegistry = True
        Else
            DateGoodRegistry = False
        End If
        If DateGoodRegistry Then
            '* daysLeft = numDays - ActiveLockDate(Date.UtcNow).Subtract(ActiveLockDate(TmpFRD)).Days
            daysLeft = numDays - Date.UtcNow.Subtract(TmpFRD).Days '*
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
    Public Function RunsGoodRegistry(ByRef numRuns As Short, ByRef runsLeft As Short) As Boolean
        On Error GoTo RunsGoodRegistryError
        Dim TmpCRD As Date
        Dim TmpLRD As Date
        Dim TmpFRD As Date

        TmpCRD = CDate(INITIALDATE)
        '* TmpLRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        '* TmpFRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        TmpLRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        TmpFRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        runsLeft = Int(CDbl(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(numRuns - 1)))))) - 1
        RunsGoodRegistry = False

        'If this is the applications first load, write initial settings
        'to the registry
        If TmpLRD = CDate(dec2("93.8D.93.8D.96.90.90.90")) Then '1/1/2000
            '*SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(CStr(Now)))
            '*SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", enc2(CStr(TmpCRD)))
            SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(DateToDblString(Now)))
            SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", enc2(DateToDblString(TmpCRD)))
            SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(numRuns - 1)))
        End If
        'Read LRD and FRD from registry
        '*TmpLRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        '*TmpFRD = CDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        If TmpLRD = CDate(dec2("93.8D.93.8D.96.90.90.90")) Then
            runsLeft = Int(CDbl(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(numRuns - 1))))))
        Else
            runsLeft = Int(CDbl(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(numRuns - 1)))))) - 1
        End If
        TmpLRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        TmpFRD = DblStringToDate(dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"))) '1/1/2000
        If TmpLRD = "#12:00:00 AM#" Then
            TmpLRD = CDate(INITIALDATE)
            TmpFRD = CDate(INITIALDATE)
            runsLeft = numRuns - 1
        End If

        If runsLeft < 0 Then 'impossible
            RunsGoodRegistry = False
        ElseIf runsLeft > numRuns Then  'Trial runs expired
            SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS))
            RunsGoodRegistry = False
        ElseIf numRuns >= runsLeft Then  'Everything OK write the remaining number of runs
            If TmpLRD = CDate(INITIALDATE) Then
                SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(numRuns - 1)))
            Else
                SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2(CStr(runsLeft)))
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

    Public Function DateGood(ByVal numDays As Short, ByRef daysLeft As Short, ByRef TrialHideTypes As IActiveLock.ALTrialHideTypes) As Boolean
        Dim use2 As Boolean
        Dim use3, use4 As Boolean
        Dim daysLeft2 As Short
        Dim daysLeft3, daysLeft4 As Short

        TEXTMSG_DAYS = DecryptString128Bit("sQvYYRLPon5IyH6BQRAUBuCLTq/5VkH3kl7HUwJLZ2M=", PSWD)
        DateGood = False

        If TrialSteganographyExists(TrialHideTypes) Then
            If DateGoodSteganography(numDays, daysLeft2) = False Then
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS)
                'MsgBox "DateGoodSteganography " & daysLeft2
                Exit Function
            End If
            use2 = True
        End If
        If TrialHiddenFolderExists(TrialHideTypes) Then
            If DateGoodHiddenFolder(numDays, daysLeft3) = False Then
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS)
                'MsgBox "DateGoodHiddenFolder " & daysLeft3
                Exit Function
            End If
            use3 = True
        End If
        If TrialRegistryPerUserExists(TrialHideTypes) Then
            If DateGoodRegistry(numDays, daysLeft4) = False Then
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS)
                'MsgBox "DateGoodRegistry " & daysLeft4
                Exit Function
            End If
            use4 = True
        End If
        'MsgBox "DateGoodSteganography " & daysLeft2
        'MsgBox "DateGoodHiddenFolder " & daysLeft3
        'MsgBox "DateGoodRegistry " & daysLeft4

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
    Public Function RunsGood(ByVal numRuns As Short, ByRef runsLeft As Short, ByRef TrialHideTypes As IActiveLock.ALTrialHideTypes) As Boolean
        Dim use2 As Boolean
        Dim use3, use4 As Boolean
        Dim runsLeft2 As Short
        Dim runsLeft3, runsLeft4 As Short
        TEXTMSG_RUNS = DecryptString128Bit("6urN2+xbgqbLLsOoC4hbGpLT3bnvY3YPGW299cOnqfo=", PSWD)

        RunsGood = False

        If TrialSteganographyExists(TrialHideTypes) Then
            If RunsGoodSteganography(numRuns, runsLeft2) = False Then
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS)
                'MsgBox "RunsGoodSteganography " & runsLeft2
                Exit Function
            End If
            use2 = True
        End If

        If TrialHiddenFolderExists(TrialHideTypes) Then
            If RunsGoodHiddenFolder(numRuns, runsLeft3) = False Then
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS)
                'MsgBox "RunsGoodHiddenFolder " & runsLeft3
                Exit Function
            End If
            use3 = True
        End If

        If TrialRegistryPerUserExists(TrialHideTypes) Then
            If RunsGoodRegistry(numRuns, runsLeft4) = False Then
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS)
                'MsgBox "RunsGoodRegistry " & runsLeft4
                Exit Function
            End If
            use4 = True
        End If

        'MsgBox "RunsGoodSteganography " & runsLeft2
        'MsgBox "RunsGoodHiddenFolder " & runsLeft3
        'MsgBox "RunsGoodRegistry " & runsLeft4

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

    Public Function DateGoodHiddenFolder(ByRef numDays As Short, ByRef daysLeft As Short) As Boolean
        'Hidden Folder Parameters:
        ' CRD: Current Run Date
        ' LRD: Last Run Date
        ' FRD: First Run Date
        Dim TmpCRD As Date
        Dim TmpLRD As Date
        Dim TmpFRD As Date
        Dim strMyString, strSource As String
        Dim intFF As Short
        Dim ok As Double

        On Error GoTo DateGoodHiddenFolderError
        If Directory.Exists(ActivelockGetSpecialFolder(46) & DecryptString128Bit(myDir(), PSWD)) = False Then MkDir(ActivelockGetSpecialFolder(46) & DecryptString128Bit(myDir, PSWD))
        strSource = HiddenFolderFunction()

        Dim checkFile As System.IO.DirectoryInfo
        Dim dirPath As String
        dirPath = ActivelockGetSpecialFolder(46) & DecryptString128Bit(HIDDENFOLDER, PSWD)
        checkFile = New System.IO.DirectoryInfo(dirPath)
        Dim attributeReader As System.IO.FileAttributes
        attributeReader = checkFile.Attributes

        If Directory.Exists(dirPath) = True And (attributeReader And System.IO.FileAttributes.Directory And System.IO.FileAttributes.Hidden And System.IO.FileAttributes.ReadOnly And System.IO.FileAttributes.System) > 0 Then
            MinusAttributes()
            'Check to see if our file is there
            If fileExist(strSource) Then
                SetAttr(strSource, FileAttribute.Normal)
                '    Else
                '        ' User found the file and deleted it; expire
                '        PlusAttributes
                '        DateGoodHiddenFolder = False
                '        Exit Function
            End If
        ElseIf Directory.Exists(dirPath) = True Then
            'Ok, the folder is there with no hidden, system attributes
            'Check to see if our file is there
            If fileExist(strSource) Then
                SetAttr(strSource, FileAttribute.Normal)
                '    Else
                '        ' User found the file and deleted it; expire
                '        PlusAttributes
                '        DateGoodHiddenFolder = False
                '        Exit Function
            End If
        Else
            MkDir(dirPath)
        End If

        CreateHdnFile()

        '*TmpCRD = ActiveLockDate(Date.UtcNow)
        TmpCRD = Date.UtcNow
        Dim a() As String
        Dim aa As String
        If fileExist(strSource) Then

            Dim strContents1 As String
            Dim objReader1 As StreamReader
            objReader1 = New StreamReader(strSource)
            strMyString = objReader1.ReadToEnd()
            objReader1.Close()
            If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
            If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
            strMyString = DecryptString128Bit(strMyString, PSWD)

            '' Read the file...
            'intFF = FreeFile()
            'FileOpen(intFF, strSource, OpenMode.Input)
            'strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
            'If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
            'If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
            'strMyString = DecryptString128Bit(strMyString, PSWD)
            'FileClose(intFF)

            If strMyString <> "" Then
                a = strMyString.Split("_")
                '*    If a(1) <> "" Then TmpLRD = CDate(a(1))
                '*  If a(2) <> "" Then TmpFRD = CDate(a(2))
                If a(1) <> "" Then TmpLRD = DblStringToDate(a(1))
                If a(2) <> "" Then TmpFRD = DblStringToDate(a(2))
                If a(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
                    DateGoodHiddenFolder = False
                    Exit Function
                End If
            End If
            If TmpLRD = "#12:00:00 AM#" Then TmpLRD = CDate(INITIALDATE)
            If TmpFRD = "#12:00:00 AM#" Then TmpFRD = CDate(INITIALDATE)
        Else
            TmpLRD = CDate(INITIALDATE)
            TmpFRD = CDate(INITIALDATE)
        End If
        DateGoodHiddenFolder = False

        'If this is the applications first load, write initial settings
        'to Hidden Folder
        If TmpLRD = CDate(INITIALDATE) Then
            TmpLRD = TmpCRD
            TmpFRD = TmpCRD
            ' Write to the file...
            intFF = FreeFile()
            FileOpen(intFF, strSource, OpenMode.Output)
            '* PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & TmpLRD & "_" & TmpFRD & "_" & "0", PSWD))
            PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(TmpLRD) & "_" & DateToDblString(TmpFRD) & "_" & "0", PSWD)) '*
            FileClose(intFF)
        End If
        'Read LRD and FRD from Hidden Folder
        Dim b() As String

        Dim strContents As String
        Dim objReader As StreamReader
        objReader = New StreamReader(strSource)
        strMyString = objReader.ReadToEnd()
        objReader.Close()
        If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
        If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
        strMyString = DecryptString128Bit(strMyString, PSWD)

        '' Read the file...
        'intFF = FreeFile()
        'FileOpen(intFF, strSource, OpenMode.Input)
        'strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
        'If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
        'If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
        'strMyString = DecryptString128Bit(strMyString, PSWD)
        'FileClose(intFF)

        b = strMyString.Split("_")
        '* If b(1) <> "" Then TmpLRD = CDate(b(1))
        '* If b(2) <> "" Then TmpFRD = CDate(b(2))
        If b(1) <> "" Then TmpLRD = DblStringToDate(b(1))
        If b(2) <> "" Then TmpFRD = DblStringToDate(b(2))
        If b(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
            DateGoodHiddenFolder = False
            Exit Function
        End If
        If TmpLRD = "#12:00:00 AM#" Then TmpLRD = CDate(INITIALDATE)
        If TmpFRD = "#12:00:00 AM#" Then TmpFRD = CDate(INITIALDATE)

        '*If ActiveLockDate(TmpFRD) > ActiveLockDate(TmpCRD) Then 'System clock rolled back
        If TmpFRD > TmpCRD Then 'System clock rolled back
            DateGoodHiddenFolder = False
            '* ElseIf ActiveLockDate(Date.UtcNow) > ActiveLockDate(TmpFRD).AddDays(numDays) Then  'trial expired
        ElseIf Date.UtcNow > TmpFRD.AddDays(numDays) Then  'trial expired
            ' Write to the file...
            intFF = FreeFile()
            FileOpen(intFF, strSource, OpenMode.Output)
            PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS, PSWD))
            FileClose(intFF)
            DateGoodHiddenFolder = False
            '* ElseIf ActiveLockDate(TmpCRD) > ActiveLockDate(TmpLRD) Then  'Everything OK write New LRD date
        ElseIf TmpCRD > TmpLRD Then  'Everything OK write New LRD date
            ' Write to the file...
            intFF = FreeFile()
            FileOpen(intFF, strSource, OpenMode.Output)
            '* PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & TmpCRD & "_" & TmpFRD & "_" & "0", PSWD))
            PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(TmpCRD) & "_" & DateToDblString(TmpFRD) & "_" & "0", PSWD)) '*
            FileClose(intFF)
            DateGoodHiddenFolder = True
            '*ElseIf ActiveLockDate(TmpCRD) = ActiveLockDate(TmpLRD) Then
        ElseIf TmpCRD = TmpLRD Then '*
            DateGoodHiddenFolder = True
        Else
            DateGoodHiddenFolder = False
        End If

        SetAttr(strSource, FileAttribute.ReadOnly + FileAttribute.Hidden + FileAttribute.System)
        PlusAttributes()

        If DateGoodHiddenFolder Then
            '* daysLeft = numDays - ActiveLockDate(Date.UtcNow).Subtract(ActiveLockDate(TmpFRD)).Days
            daysLeft = numDays - Date.UtcNow.Subtract(TmpFRD).Days '*
        Else
            daysLeft = 0
        End If
        Exit Function

DateGoodHiddenFolderError:

    End Function
    Public Function RunsGoodHiddenFolder(ByRef numRuns As Short, ByRef runsLeft As Short) As Boolean
        Dim TmpCRD As Date
        Dim TmpLRD As Date
        Dim TmpFRD As Date
        Dim strMyString, strSource As String
        Dim intFF As Short
        Dim ok As Double

        On Error GoTo RunsGoodHiddenFolderError
        If Directory.Exists(ActivelockGetSpecialFolder(46) & DecryptString128Bit(myDir, PSWD)) = False Then MkDir(ActivelockGetSpecialFolder(46) & DecryptString128Bit(myDir, PSWD))
        strSource = HiddenFolderFunction()

        Dim checkFile As System.IO.DirectoryInfo
        Dim dirPath As String
        dirPath = ActivelockGetSpecialFolder(46) & DecryptString128Bit(HIDDENFOLDER, PSWD)
        checkFile = New System.IO.DirectoryInfo(dirPath)
        Dim attributeReader As System.IO.FileAttributes
        attributeReader = checkFile.Attributes

        If Directory.Exists(dirPath) = True And (attributeReader And System.IO.FileAttributes.Directory And System.IO.FileAttributes.Hidden And System.IO.FileAttributes.ReadOnly And System.IO.FileAttributes.System) > 0 Then
            MinusAttributes()
            'Check to see if our file is there
            If fileExist(strSource) Then
                SetAttr(strSource, FileAttribute.Normal)
                'Else
                '    ' User found the file and deleted it; expire
                '    PlusAttributes()
                '    RunsGoodHiddenFolder = False
                '    Exit Function
            End If
        ElseIf Directory.Exists(dirPath) = True Then
            'Ok, the folder is there with no hidden, system attributes
            'Check to see if our file is there
            If fileExist(strSource) Then
                SetAttr(strSource, FileAttribute.Normal)
                'Else
                '    ' User found the file and deleted it; expire
                '    PlusAttributes()
                '    RunsGoodHiddenFolder = False
                '    Exit Function
            End If
        Else
            MkDir(dirPath)
        End If

        CreateHdnFile()

        TmpCRD = CDate(INITIALDATE)
        Dim a() As String
        If fileExist(strSource) Then

            Dim strContents2 As String
            Dim objReader2 As StreamReader
            objReader2 = New StreamReader(strSource)
            strMyString = objReader2.ReadToEnd()
            objReader2.Close()
            If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
            If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
            strMyString = DecryptString128Bit(strMyString, PSWD)

            '' Read the file...
            'intFF = FreeFile()
            'FileOpen(intFF, strSource, OpenMode.Input)
            'strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
            'If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
            'If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
            'strMyString = DecryptString128Bit(strMyString, PSWD)
            'FileClose(intFF)

            If strMyString <> "" Then
                On Error GoTo [continue]
                a = strMyString.Split("_")
                '* If a(1) <> "" Then TmpLRD = CDate(a(1))
                '*If a(2) <> "" Then TmpFRD = CDate(a(2))
                If a(1) <> "" Then TmpLRD = DblStringToDate(a(1)) '*
                If a(2) <> "" Then TmpFRD = DblStringToDate(a(2)) '*
                runsLeft = Int(CDbl(a(3))) - 1
                If a(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
                    RunsGoodHiddenFolder = False
                    Exit Function
                End If
            End If
[continue]:
            If TmpLRD = "#12:00:00 AM#" Then
                TmpLRD = CDate(INITIALDATE)
                TmpFRD = CDate(INITIALDATE)
                runsLeft = numRuns - 1
            End If
        Else
            TmpLRD = CDate(INITIALDATE)
            TmpFRD = CDate(INITIALDATE)
            runsLeft = numRuns - 1
        End If
        RunsGoodHiddenFolder = False

        'If this is the applications first load, write initial settings
        'to Hidden Folder
        If TmpLRD = CDate(INITIALDATE) Then
            TmpFRD = CDate(INITIALDATE)
            runsLeft = numRuns - 1
            ' Write to the file...
            intFF = FreeFile()
            FileOpen(intFF, strSource, OpenMode.Output)
            '*PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(Date.UtcNow) & "_" & TmpFRD & "_" & CStr(numRuns - 1), PSWD))
            PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(Date.UtcNow) & "_" & DateToDblString(TmpFRD) & "_" & CStr(numRuns - 1), PSWD))
            FileClose(intFF)
        End If
        'Read LRD and FRD from Hidden Folder
        Dim b() As String

        Dim strContents As String
        Dim objReader As StreamReader
        objReader = New StreamReader(strSource)
        strMyString = objReader.ReadToEnd()
        objReader.Close()
        If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
        If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
        strMyString = DecryptString128Bit(strMyString, PSWD)

        '' Read the file...
        'intFF = FreeFile()
        'FileOpen(intFF, strSource, OpenMode.Input)
        'strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
        'If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
        'If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
        'strMyString = DecryptString128Bit(strMyString, PSWD)
        'FileClose(intFF)

        b = strMyString.Split("_")
        '* If b(1) <> "" Then TmpLRD = CDate(b(1))
        '* If b(2) <> "" Then TmpFRD = CDate(b(2))
        If TmpLRD = CDate(INITIALDATE) Then
            runsLeft = Int(CDbl(b(3)))
        Else
            runsLeft = Int(CDbl(b(3))) - 1
        End If
        If b(1) <> "" Then TmpLRD = DblStringToDate(b(1)) '*
        If b(2) <> "" Then TmpFRD = DblStringToDate(b(2)) '*
        If b(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
            RunsGoodHiddenFolder = False
            Exit Function
        End If
        If TmpLRD = "#12:00:00 AM#" Then
            TmpLRD = CDate(INITIALDATE)
            TmpFRD = CDate(INITIALDATE)
            runsLeft = numRuns - 1
        End If

        If runsLeft < 0 Then 'impossible
            RunsGoodHiddenFolder = False
        ElseIf runsLeft > numRuns Then  'Trial runs expired
            ' Write to the file...
            intFF = FreeFile()
            FileOpen(intFF, strSource, OpenMode.Output)
            PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS, PSWD))
            FileClose(intFF)
            RunsGoodHiddenFolder = False
        ElseIf numRuns >= runsLeft Then  'Everything OK write the remaining number of runs
            ' Write to the file...
            intFF = FreeFile()
            FileOpen(intFF, strSource, OpenMode.Output)
            If TmpLRD = CDate(INITIALDATE) Then
                '*  PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(Date.UtcNow) & "_" & TmpFRD & "_" & CStr(numRuns - 1), PSWD))
                PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(Date.UtcNow) & "_" & DateToDblString(TmpFRD) & "_" & CStr(numRuns - 1), PSWD)) '*
            Else
                '* PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(Date.UtcNow) & "_" & TmpFRD & "_" & CStr(runsLeft), PSWD))
                PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(Date.UtcNow) & "_" & DateToDblString(TmpFRD) & "_" & CStr(runsLeft), PSWD)) '*
            End If
            FileClose(intFF)
            RunsGoodHiddenFolder = True
        Else
            RunsGoodHiddenFolder = False
        End If

        SetAttr(strSource, FileAttribute.ReadOnly + FileAttribute.Hidden + FileAttribute.System)
        PlusAttributes()

        If RunsGoodHiddenFolder Then
        Else
            runsLeft = 0
        End If
        Exit Function

RunsGoodHiddenFolderError:

    End Function

    Public Function DateGoodSteganography(ByRef numDays As Short, ByRef daysLeft As Short) As Boolean
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

        If StegInfo = String.Empty Then
            GetSteganographyInfo()
        End If

        '* TmpCRD = ActiveLockDate(Date.UtcNow)
        TmpCRD = Date.UtcNow
        Dim a() As String
        Dim aa As String
        aa = StegInfo
        If aa <> "" Then
            On Error GoTo [continue]
            a = aa.Split("_")
            '* If a(1) <> "" Then TmpLRD = CDate(a(1))
            '* If a(2) <> "" Then TmpFRD = CDate(a(2))
            If a(1) <> "" Then TmpLRD = DblStringToDate(a(1)) '*
            If a(2) <> "" Then TmpFRD = DblStringToDate(a(2)) '*
            If a(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
                DateGoodSteganography = False
                Exit Function
            End If
        End If
[continue]:
        If TmpLRD = "#12:00:00 AM#" Then TmpLRD = CDate(INITIALDATE)
        If TmpFRD = "#12:00:00 AM#" Then TmpFRD = CDate(INITIALDATE)
        DateGoodSteganography = False

        'If this is the application's first load, write initial settings
        'to the image file via steganography
        If TmpLRD = CDate(INITIALDATE) Then
            TmpLRD = TmpCRD
            TmpFRD = TmpCRD
            '*SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & TmpCRD & "_" & TmpFRD & "_" & "0")
            SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(TmpCRD) & "_" & DateToDblString(TmpFRD) & "_" & "0")
            GetSteganographyInfo()
        End If
        'Read LRD and FRD from the hidden text in the image
        Dim b() As String
        Dim bb As String
        bb = StegInfo
        b = bb.Split("_")
        '* If b(1) <> "" Then TmpLRD = CDate(b(1))
        '*If b(2) <> "" Then TmpFRD = CDate(b(2))
        If b(1) <> "" Then TmpLRD = DblStringToDate(b(1)) '*
        If b(2) <> "" Then TmpFRD = DblStringToDate(b(2)) '*
        If b(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
            DateGoodSteganography = False
            Exit Function
        End If
        If TmpLRD = "#12:00:00 AM#" Then TmpLRD = CDate(INITIALDATE)
        If TmpFRD = "#12:00:00 AM#" Then TmpFRD = CDate(INITIALDATE)

        '*If ActiveLockDate(TmpFRD) > ActiveLockDate(TmpCRD) Then 'System clock rolled back
        If TmpFRD > TmpCRD Then 'System clock rolled back
            DateGoodSteganography = False
            '*ElseIf ActiveLockDate(Date.UtcNow) > ActiveLockDate(TmpFRD).AddDays(numDays) Then  'trial expired
        ElseIf Date.UtcNow > TmpFRD.AddDays(numDays) Then  'trial expired
            SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS)
            DateGoodSteganography = False
            '* ElseIf ActiveLockDate(TmpCRD) > ActiveLockDate(TmpLRD) Then  'Everything OK write New LRD date
            '* SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & TmpCRD & "_" & TmpFRD & "_" & "0")
        ElseIf TmpCRD > TmpLRD Then  'Everything OK write New LRD date
            SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(TmpCRD) & "_" & DateToDblString(TmpFRD) & "_" & "0") '*
            DateGoodSteganography = True
            '* ElseIf ActiveLockDate(TmpCRD) = ActiveLockDate(TmpLRD) Then
        ElseIf TmpCRD = TmpLRD Then '*
            DateGoodSteganography = True
        Else
            DateGoodSteganography = False
        End If
        If DateGoodSteganography Then
            '* daysLeft = numDays - ActiveLockDate(Date.UtcNow).Subtract(ActiveLockDate(TmpFRD)).Days
            daysLeft = numDays - Date.UtcNow.Subtract(TmpFRD).Days '*
        Else
            daysLeft = 0
        End If
        Exit Function
DateGoodSteganographyError:
        daysLeft = 0
        DateGoodSteganography = False
        Exit Function
    End Function
    '''* <summary>
    '''* converts a date to a double and then to a string
    '''* useful for ensureing dates can be read when pulled back out
    '''* no matter the locale
    '''* </summary>
    '''* <param name="Dte">the date to convert</param>
    '''* <returns>a double number in a string</returns>
    '''* <remarks>can blow up if passed an invalid date?</remarks>
    Public Function DateToDblString(ByRef Dte As Date) As String
#If VBC_VER > 6.0 Then
        Dim nfi As System.Globalization.NumberFormatInfo = New System.Globalization.CultureInfo("en-US", False).NumberFormat()
        Return Dte.ToOADate().ToString(nfi)
        'Return Dte.ToOADate().ToString
#Else
        Return CStr(CDbl(Dte))
#End If

    End Function

    '''* <summary>
    '''* Function is used to convert doubles stored in strings to dates
    '''* useful because whenever we store dates we convert them (to doubles and then strings)
    '''* in case the user changes the locale in between storage and retrieval.
    '''* minor handling of actual date strings for some semblance of backward 
    '''* compatibility
    '''* </summary>
    '''* <param name="Dstr">The string to pass</param>
    '''* <returns>a date representation on the passed Dstr or
    '''*  1/1/1900 if a conversion error occurred</returns>
    '''* <remarks></remarks>
    Public Function DblStringToDate(ByRef Dstr As String) As Date
#If VBC_VER > 6.0 Then

        ' double varaible to hold our return if tryparse succeeds
        ' culture to ensure we do this the same way everytime
        Dim Dr As Double
        Dim Drd As Date
        Dim culture As Globalization.CultureInfo = Globalization.CultureInfo.GetCultureInfo("en-US") 'CurrentUICulture

        If Double.TryParse(Dstr, Globalization.NumberStyles.Float, culture, Dr) Then
            Return Date.FromOADate(Dr)
        Else
            'this is an older version we need to start the upgrade process
            If Date.TryParse(Dstr, culture, Globalization.DateTimeStyles.AdjustToUniversal, Drd) Then
                Return Drd
            Else
                Try
                    Drd = Date.FromOADate(CDbl(Val(Dstr)))
                Catch ex As Exception
                    Return #1/1/1900#
                End Try
                Return Drd
            End If
        End If
#Else

        Try
            ' do something similar
            If Dstr <> "" Then
                Dim Dbl As Double = CDbl(Dstr)
                Return CDate(Dbl)
            End If

        Catch ex As Exception
            'probably a date it wasn't empty
            Try
                Return CDate(Dstr)
            Catch ex As Exception
                Return #1/1/1900#
            End Try

        End Try

#End If
    End Function

    Public Function ActiveLockDate(ByVal dt As Date) As Date
        ' CDate(Format(dt, "m/d/yy h:m:ss"))
        Dim newDate As Date
        Dim m As Integer
        Dim d As Integer
        Dim y As Integer
        newDate = dt
        m = newDate.Month
        d = newDate.Day
        y = newDate.Year
        ActiveLockDate = CDate(y.ToString("0000") & "/" & m.ToString("00") & "/" & d.ToString("00"))
        'Err.Raise(Globals.ActiveLockErrCodeConstants.alerrDateError, ACTIVELOCKSTRING, STRLICENSEINVALID)
    End Function
    Public Function RunsGoodSteganography(ByRef numRuns As Short, ByRef runsLeft As Short) As Boolean
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

        TmpCRD = CDate(INITIALDATE)
        Dim a() As String
        Dim aa As String
        aa = SteganographyPull(strSource)
        If aa <> "" Then
            On Error GoTo [continue]
            a = aa.Split("_")
            '* If a(1) <> "" Then TmpLRD = CDate(a(1))
            '* If a(2) <> "" Then TmpFRD = CDate(a(2))
            If a(1) <> "" Then TmpLRD = DblStringToDate(a(1)) '*
            If a(2) <> "" Then TmpFRD = DblStringToDate(a(2)) '*
            runsLeft = Int(CDbl(a(3))) - 1
            If a(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
                RunsGoodSteganography = False
                Exit Function
            End If
        End If
[continue]:
        If TmpLRD = "#12:00:00 AM#" Then
            TmpLRD = CDate(INITIALDATE)
            TmpFRD = CDate(INITIALDATE)
            runsLeft = numRuns - 1
        End If
        RunsGoodSteganography = False

        'If this is the application's first load, write initial settings
        'to the image file via steganography
        If TmpLRD = CDate(INITIALDATE) Then
            TmpFRD = CDate(INITIALDATE)
            runsLeft = numRuns - 1
            '*SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(Date.UtcNow) & "_" & TmpFRD & "_" & CStr(numRuns - 1))
            SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(Date.UtcNow) & "_" & DateToDblString(TmpFRD) & "_" & CStr(numRuns - 1)) '*
        End If
        'Read LRD and FRD from the hidden text in the image
        Dim b() As String
        b = SteganographyPull(strSource).Split("_")
        '* If b(1) <> "" Then TmpLRD = CDate(b(1))
        '* If b(2) <> "" Then TmpFRD = CDate(b(2))
        If TmpLRD = CDate(INITIALDATE) Then
            runsLeft = Int(CDbl(b(3)))
        Else
            runsLeft = Int(CDbl(b(3))) - 1
        End If
        If b(1) <> "" Then TmpLRD = DblStringToDate(b(1)) '*
        If b(2) <> "" Then TmpFRD = DblStringToDate(b(2)) '*
        If b(0) <> LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD Then
            RunsGoodSteganography = False
            Exit Function
        End If
        If TmpLRD = "#12:00:00 AM#" Then
            TmpLRD = CDate(INITIALDATE)
            TmpFRD = CDate(INITIALDATE)
            runsLeft = numRuns - 1
        End If

        If runsLeft < 0 Then 'impossible
            RunsGoodSteganography = False
        ElseIf runsLeft > numRuns Then  'Trial runs expired
            SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS)
            RunsGoodSteganography = False
        ElseIf numRuns >= runsLeft Then  'Everything OK write the remaining number of runs
            If TmpLRD = CDate(INITIALDATE) Then
                '*SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(Date.UtcNow) & "_" & TmpFRD & "_" & CStr(numRuns - 1))
                SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(Date.UtcNow) & "_" & DateToDblString(TmpFRD) & "_" & CStr(numRuns - 1)) '*
            Else
                '*SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & ActiveLockDate(Date.UtcNow) & "_" & TmpFRD & "_" & CStr(runsLeft))
                SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & DateToDblString(Date.UtcNow) & "_" & DateToDblString(TmpFRD) & "_" & CStr(runsLeft)) '*
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
        Dim ok As Short
        'Set up the error handler to trap the File Not Found
        'message, or other errors.
        On Error GoTo FileExistErrors
        'Check for attributes of test file. If this function
        'does not raise an error, than the file must exist.
        ok = GetAttr(TestFileName)
        'If no errors encountered, then the file must exist
        fileExist = True
        Exit Function
FileExistErrors:  'error handling routine, including File Not Found
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
        Dim sRes As String = Nothing
        Dim i As Short
        Dim arr() As String
        If strdata = "" Then
            dec2 = ""
            Exit Function
        End If
        arr = strdata.Split(".")
        For i = LBound(arr) To UBound(arr)
            sRes = sRes & Chr(CInt("&h" & arr(i)) / (2 + 1))
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
        Dim i, N As Integer
        Dim sResult As String = String.Empty

        N = Len(strdata)
        Dim l As Integer
        For i = 1 To N
            l = Asc(Mid(strdata, i, 1)) * (2 + 1)
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
    Public Function ExpireTrial(ByVal SoftwareName As String, ByVal SoftwareVer As String, ByVal TrialType As Integer, ByVal TrialLength As Integer, ByVal TrialHideTypes As IActiveLock.ALTrialHideTypes, ByVal SoftwarePassword As String) As Boolean
        'Dim secondRegistryKey As Boolean
        Dim ok As Double
        Dim strSource As String

        On Error GoTo triAlerror

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
        'secondRegistryKey = CreateRegistryKey(HKEY_CLASSES_ROOT, DecryptString128Bit("5985D6B80E543AFCA67570BF9924469349EDA3A8695B75E656E95ACA55360118A4128395B2B070E8DC04FFB01C7509B18CF9831F36EF68D4A438130BF5F94587C76AE48AD5D6A210DAAB895120982C3426D3EA65C253A39B0C1131D1848D6518", PSWD) & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))
        'secondRegistryKey = CreateRegistryKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility-" & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))

        Dim intFF As Short
        intFF = FreeFile()
        FileOpen(intFF, ActivelockGetSpecialFolder(46) & "\" & CHANNELS & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(99) & Chr(100) & Chr(102), OpenMode.Output)
        PrintLine(intFF, "23g5985hb587b27eb")
        FileClose(intFF)

        If Directory.Exists(ActivelockGetSpecialFolder(55) & "\Sample Videos") = False Then
            My.Computer.FileSystem.CreateDirectory(ActivelockGetSpecialFolder(55) & "\Sample Videos")
        End If
        intFF = FreeFile()
        FileOpen(intFF, ActivelockGetSpecialFolder(55) & "\Sample Videos" & VIDEO & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & OTHERFILE, OpenMode.Output)
        PrintLine(intFF, "012234trliug2gb88y53")
        FileClose(intFF)

        ' Registry stuff
        If TrialRegistryPerUserExists(TrialHideTypes) Then
            SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS))
        End If

        ' Steganography stuff
        If TrialSteganographyExists(TrialHideTypes) Then
            strSource = GetSteganographyFile()
            If strSource <> "" Then
                SteganographyEmbed(strSource, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS)
            End If
        End If

        ' Hidden folder stuff
        If TrialHiddenFolderExists(TrialHideTypes) Then
            If Directory.Exists(ActivelockGetSpecialFolder(46) & DecryptString128Bit(myDir, PSWD)) = False Then
                MkDir(ActivelockGetSpecialFolder(46) & DecryptString128Bit(myDir, PSWD))
            End If
            strSource = HiddenFolderFunction()

            Dim checkFile As System.IO.DirectoryInfo
            Dim dirPath As String
            dirPath = ActivelockGetSpecialFolder(46) & DecryptString128Bit(HIDDENFOLDER, PSWD)
            checkFile = New System.IO.DirectoryInfo(dirPath)
            Dim attributeReader As System.IO.FileAttributes
            attributeReader = checkFile.Attributes

            If Directory.Exists(dirPath) = True And (attributeReader And System.IO.FileAttributes.Directory And System.IO.FileAttributes.Hidden And System.IO.FileAttributes.ReadOnly And System.IO.FileAttributes.System) > 0 Then
                MinusAttributes()
                'Check to see if our file is there
                If fileExist(strSource) Then
                    SetAttr(strSource, FileAttribute.Normal)
                    Kill(strSource)
                End If
            ElseIf Directory.Exists(dirPath) = True Then
                'Ok, the folder is there with no system, hidden attributes
                'Check to see if our file is there
                If fileExist(strSource) Then
                    SetAttr(strSource, FileAttribute.Normal)
                    Kill(strSource)
                End If
            Else
                MkDir(dirPath)
            End If
            CreateHdnFile()
            If fileExist(strSource) Then
                SetAttr(strSource, FileAttribute.Normal)
                Kill(strSource)
            End If
            ' Write to the file...
            intFF = FreeFile()
            FileOpen(intFF, strSource, OpenMode.Output)
            PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & EXPIRED_DAYS & "_" & EXPIRED_DAYS & "_" & EXPIRED_RUNS, PSWD))
            FileClose(intFF)
            SetAttr(strSource, FileAttribute.ReadOnly + FileAttribute.Hidden + FileAttribute.System)
            PlusAttributes()
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
        Exit Function

triAlerror:
        'Call CloseHandle(hFolder)
        ExpireTrial = False
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
    Public Function ResetTrial(ByVal SoftwareName As String, ByVal SoftwareVer As String, ByVal TrialType As Integer, ByVal TrialLength As Integer, ByVal TrialHideTypes As IActiveLock.ALTrialHideTypes, ByVal SoftwarePassword As String) As Boolean
        Dim ok As Integer
        Dim strSourceFile As String
        Dim rtn As Byte

        On Error Resume Next

        LICENSE_SOFTWARE_NAME = SoftwareName
        LICENSE_SOFTWARE_VERSION = SoftwareVer
        LICENSE_SOFTWARE_PASSWORD = SoftwarePassword

        EXPIRED_RUNS = Chr(101) & Chr(120) & Chr(112) & Chr(105) & Chr(114) & Chr(101) & Chr(100)
        EXPIRED_DAYS = EXPIRED_RUNS
        VIDEO = Chr(92) & Chr(86) & Chr(105) & Chr(100) & Chr(101) & Chr(111)
        OTHERFILE = Chr(68) & Chr(114) & Chr(105) & Chr(118) & Chr(101) & Chr(114) & Chr(115) & "." & Chr(100) & Chr(108) & Chr(108)

        'Expire warning
        DeleteSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "1"))

        ' The following two keys are not compatible with Vista
        ' A regular user account cannot have write access to these two registry hives
        ' I am removing these from v3.6 - ialkan 12-27-2008
        ' The following were created only when the license expired
        'secondRegistryKey = DeleteKey(HKEY_CLASSES_ROOT, DecryptString128Bit("5985D6B80E543AFCA67570BF9924469349EDA3A8695B75E656E95ACA55360118A4128395B2B070E8DC04FFB01C7509B18CF9831F36EF68D4A438130BF5F94587C76AE48AD5D6A210DAAB895120982C3426D3EA65C253A39B0C1131D1848D6518", PSWD) & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))
        'secondRegistryKey = DeleteKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility-" & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))

        If fileExist(ActivelockGetSpecialFolder(46) & "\" & CHANNELS & strLeft(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(99) & Chr(100) & Chr(102)) Then
            Kill(ActivelockGetSpecialFolder(46) & "\" & CHANNELS & strLeft(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(99) & Chr(100) & Chr(102))
        End If
        If fileExist(ActivelockGetSpecialFolder(55) & "\Sample Videos" & VIDEO & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & OTHERFILE) Then
            Kill(ActivelockGetSpecialFolder(55) & "\Sample Videos" & VIDEO & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & OTHERFILE)
        End If

        ' Registry stuff
        On Error Resume Next
        If TrialRegistryPerUserExists(TrialHideTypes) Then
            DeleteSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD))
        End If

        ' Steganography stuff
        If TrialSteganographyExists(TrialHideTypes) Then
            strSourceFile = GetSteganographyFile()
            If File.Exists(strSourceFile) Then Kill(strSourceFile)
            'If strSourceFile <> "" Then
            '    rtn = SteganographyEmbed(strSourceFile, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & "" & "_" & "" & "_" & "")
            'End If
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
        On Error GoTo triAlerror
        If TrialHiddenFolderExists(TrialHideTypes) Then
            If Directory.Exists(ActivelockGetSpecialFolder(46) & DecryptString128Bit(myDir, PSWD)) = True Then
                MinusAttributes()
                strSourceFile = HiddenFolderFunction()
                If fileExist(strSourceFile) Then
                    SetAttr(strSourceFile, FileAttribute.Normal)
                    Kill(strSourceFile)
                End If
                PlusAttributes()
            End If
        End If

        System.Windows.Forms.Application.DoEvents()
        ResetTrial = True
        Exit Function

triAlerror:
        'Call CloseHandle(hFolder)
        ResetTrial = False
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

        On Error GoTo IsRegistryExpired1Error

        savedRegistryKey = CheckRegistryKey(HKEY_CLASSES_ROOT, DecryptString128Bit("5985D6B80E543AFCA67570BF9924469349EDA3A8695B75E656E95ACA55360118A4128395B2B070E8DC04FFB01C7509B18CF9831F36EF68D4A438130BF5F94587C76AE48AD5D6A210DAAB895120982C3426D3EA65C253A39B0C1131D1848D6518", PSWD) & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))
        If savedRegistryKey = True Then
            IsRegistryExpired1 = True
        Else
            IsRegistryExpired1 = False
        End If
        Exit Function

IsRegistryExpired1Error:
        IsRegistryExpired1 = True
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

        On Error GoTo IsRegistryExpired2Error

        savedRegistryKey = CheckRegistryKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility-" & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))
        If savedRegistryKey = True Then
            IsRegistryExpired2 = True
        Else
            IsRegistryExpired2 = False
        End If
        Exit Function

IsRegistryExpired2Error:
        IsRegistryExpired2 = True
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
        Dim secretFolder As String
        Dim fDateTime As String = String.Empty
        secretFolder = WinDir() & Chr(92) & Chr(67) & Chr(117) & Chr(114) & Chr(115) & Chr(111) & Chr(114) & Chr(115)

        Dim hFolder As Integer
        'obtain handle to the folder specified
        hFolder = GetFolderFileHandle(secretFolder)
        System.Windows.Forms.Application.DoEvents()
        If (hFolder <> 0) And (hFolder > -1) Then
            'get the folder date/time info
            Call GetFolderFileDate(hFolder, fDateTime)
            If fDateTime.IndexOf("January 09, 2000 4:00:00") > 0 Then
                Call CloseHandle(hFolder)
                IsFolderStampExpired = True
                Exit Function
            Else
                Call CloseHandle(hFolder)
                Exit Function
            End If
        Else
            Call CloseHandle(hFolder)
            MkDir(secretFolder)
            Exit Function
        End If

        Exit Function

IsFolderStampExpiredError:
        Call CloseHandle(hFolder)
    End Function
    Private Sub GetFolderFileDate(ByVal hFolder As Integer, ByRef buff As String)

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
    Private Function ChangeFolderFileDate(ByVal hFolder As Integer, ByVal sDay As Short, ByVal sMonth As Short, ByVal sYear As Short, ByVal sHour As Short, ByVal sMinute As Short, ByVal sSecond As Short) As Boolean

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
    Private Function GetFolderFileHandle(ByRef sPath As String) As Integer

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
        GetFolderFileHandle = CreateFile(sPath, GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_DELETE, 0, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0)

    End Function
    Private Function GetFolderFileDateString(ByRef FT As FILETIME) As String

        Dim ds As Single
        Dim ts As Single
        Dim ft_local As FILETIME
        Dim st As SYSTEMTIME

        GetFolderFileDateString = String.Empty

        'convert the file time to a local
        'file time
        If FileTimeToLocalFileTime(FT, ft_local) Then
            'convert the local file time to
            'the system time format
            If FileTimeToSystemTime(ft_local, st) Then
                'calculate the DateSerial/TimeSerial
                'values for the system time
                ds = DateSerial(st.wYear, st.wMonth, st.wDay).ToOADate
                ts = TimeSerial(st.wHour, st.wMinute, st.wSecond).ToOADate
                'and return a formatted string
                GetFolderFileDateString = FormatDateTime(System.DateTime.FromOADate(ds), DateFormat.LongDate) & "  " & FormatDateTime(System.DateTime.FromOADate(ts), DateFormat.LongTime)
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
        Dim strMyString As String
        Dim strSource, strDestination As String
        Dim intFF As Short
        Dim ok As Double
        VIDEO = Chr(92) & Chr(86) & Chr(105) & Chr(100) & Chr(101) & Chr(111)
        OTHERFILE = Chr(68) & Chr(114) & Chr(105) & Chr(118) & Chr(101) & Chr(114) & Chr(115) & "." & Chr(100) & Chr(108) & Chr(108)

        On Error GoTo IsEncryptedFileExpiredError

        If fileExist(ActivelockGetSpecialFolder(46) & "\" & CHANNELS & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(99) & Chr(100) & Chr(102)) Then
            IsEncryptedFileExpired = True
        ElseIf fileExist(ActivelockGetSpecialFolder(55) & "\Sample Videos" & VIDEO & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & OTHERFILE) Then
            IsEncryptedFileExpired = True
        Else
            IsEncryptedFileExpired = False
        End If

        Exit Function

IsEncryptedFileExpiredError:
        IsEncryptedFileExpired = True
    End Function
    '===============================================================================
    ' Name: Function GetSteganographyInfo
    ' Input: None
    ' Output:
    '   None
    ' Purpose: Sets the global StegInfo value
    ' Remarks: None
    '===============================================================================
    Public Sub GetSteganographyInfo()
        Dim strSource As String

        On Error GoTo GetSteganographyInfoError

        StegInfo = String.Empty

        strSource = GetSteganographyFile()
        If strSource = "" Then
            Exit Sub
        End If

        StegInfo = SteganographyPull(strSource)

        Exit Sub

GetSteganographyInfoError:
        Exit Sub

    End Sub

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

        On Error GoTo IsSteganographyExpiredError

        If StegInfo = String.Empty Then
            GetSteganographyInfo()
        End If

        If StegInfo.IndexOf(EXPIREDDAYS) > 0 Then
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
        Dim strMyString, strSource As String
        Dim intFF As Short
        Dim ok As Double
        On Error GoTo IsHiddenFolderExpiredError

        Dim checkFile As System.IO.DirectoryInfo
        Dim dirPath As String
        dirPath = ActivelockGetSpecialFolder(46) & DecryptString128Bit(HIDDENFOLDER, PSWD)
        checkFile = New System.IO.DirectoryInfo(dirPath)
        Dim attributeReader As System.IO.FileAttributes
        attributeReader = checkFile.Attributes

        If Directory.Exists(ActivelockGetSpecialFolder(46) & DecryptString128Bit(myDir, PSWD)) = False Then Exit Function
        strSource = HiddenFolderFunction()
        If Directory.Exists(dirPath) = True And (attributeReader And System.IO.FileAttributes.Directory And System.IO.FileAttributes.Hidden And System.IO.FileAttributes.ReadOnly And System.IO.FileAttributes.System) > 0 Then
            MinusAttributes()
            'Check to see if our file is there
            If fileExist(strSource) Then
                SetAttr(strSource, FileAttribute.Normal)

                Dim strContents As String
                Dim objReader As StreamReader
                objReader = New StreamReader(strSource)
                strMyString = objReader.ReadToEnd()
                objReader.Close()
                If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
                If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
                strMyString = DecryptString128Bit(strMyString, PSWD)

                ''Ok... so far so good... now read the contents of the file:
                'intFF = FreeFile()
                '' Read the file...
                'FileOpen(intFF, strSource, OpenMode.Input)
                'strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
                'If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
                'strMyString = DecryptString128Bit(strMyString, PSWD)
                'FileClose(intFF)

                If strMyString.IndexOf(EXPIREDDAYS) > 0 Then
                    IsHiddenFolderExpired = True
                End If
                SetAttr(strSource, FileAttribute.ReadOnly + FileAttribute.Hidden + FileAttribute.System)
                PlusAttributes()
                Exit Function
            End If
        ElseIf Directory.Exists(dirPath) = True Then
            'Ok, the folder is there with no system, hidden attributes
            'Check to see if our file is there
            If fileExist(strSource) Then
                SetAttr(strSource, FileAttribute.Normal)

                Dim strContents As String
                Dim objReader As StreamReader
                objReader = New StreamReader(strSource)
                strMyString = objReader.ReadToEnd()
                objReader.Close()
                If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
                If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
                strMyString = DecryptString128Bit(strMyString, PSWD)

                ''Ok... so far so good... now read the contents of the file:
                'intFF = FreeFile()
                '' Read the file...
                'FileOpen(intFF, strSource, OpenMode.Input)
                'strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
                'If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
                'strMyString = DecryptString128Bit(strMyString, PSWD)
                'FileClose(intFF)

                If strMyString.IndexOf(EXPIREDDAYS) > 0 Then
                    IsHiddenFolderExpired = True
                End If
                SetAttr(strSource, FileAttribute.ReadOnly + FileAttribute.Hidden + FileAttribute.System)
                PlusAttributes()
                Exit Function
            End If

        Else
            IsHiddenFolderExpired = False
        End If
        Exit Function

IsHiddenFolderExpiredError:
        IsHiddenFolderExpired = True
    End Function
    Public Function SteganographyEmbed(ByRef FileName As String, ByRef embedMe As String) As Byte
        ''Returns
        ''    0  --- executed successfully
        ''    1  --- file not found
        ''    2  --- Data too long for file -- excess truncated
        If File.Exists(FileName) = False Then SteganographyEmbed = 1 : Exit Function
        Dim objCoder As CCoder = Nothing
        Dim keyFileName As String
        Try
            objCoder = New CCoder
            keyFileName = ActivelockGetSpecialFolder(46) & "\rock." & Chr(98) & Chr(109) & Chr(112)

            Dim b As System.Drawing.Bitmap
            Dim r As New Resources.ResourceManager("ActiveLock3_6Net.ProjectResources", System.Reflection.Assembly.GetExecutingAssembly)
            b = r.GetObject("bmp101")
            If fileExist(keyFileName) = False Then b.Save(keyFileName)
            'If fileExist(keyFileName) = False Then VB6.LoadResPicture(101, VB6.LoadResConstants.ResBitmap).Save(keyFileName)

            objCoder.Code(keyFileName, embedMe, FileName)
            Kill(keyFileName)

        Catch exc As Exception
            'MsgBox(exc.ToString)
        Finally
            Try
                objCoder.Dispose()
            Catch
            End Try
        End Try

    End Function
    Public Function SteganographyPull(ByVal FileName As String) As String
        '' returns a string containing the data buried in the file
        'Dim Mask, Beta As Byte
        'Dim Gamma As Integer
        'Dim Fetch, Hold As Byte
        'Dim fNum As Short

        'SteganographyPull = ""
        'Gamma = 255
        'fNum = FreeFile()
        'FileOpen(fNum, FileName, OpenMode.Binary)
        'Do
        '    System.Windows.Forms.Application.DoEvents()
        '    Hold = 0
        '    For Beta = 0 To 7 Step 2
        '        Mask = 0
        '        FileGet(fNum, Fetch, Gamma)
        '        If (Fetch And 1) Then Mask = Mask Or 2 ^ Beta
        '        If (Fetch And 2) Then Mask = Mask Or 2 ^ (Beta + 1)
        '        Hold = Hold Or Mask
        '        Gamma = Gamma + 2
        '    Next
        '    If Hold = 255 Then
        '        FileClose(fNum)
        '        Exit Function
        '    End If
        '    SteganographyPull = SteganographyPull & Chr(Hold)
        'Loop

        Dim objCoder As CCoder = Nothing
        Dim keyFileName As String
        Dim a As String = String.Empty

        SteganographyPull = String.Empty
        Try
            objCoder = New CCoder
            keyFileName = ActivelockGetSpecialFolder(46) & "\rock." & Chr(98) & Chr(109) & Chr(112)

            Dim b As System.Drawing.Bitmap
            Dim r As New Resources.ResourceManager("ActiveLock3_6Net.ProjectResources", System.Reflection.Assembly.GetExecutingAssembly)
            b = r.GetObject("bmp101")
            If fileExist(keyFileName) = False Then b.Save(keyFileName)
            'If fileExist(keyFileName) = False Then VB6.LoadResPicture(101, VB6.LoadResConstants.ResBitmap).Save(keyFileName)

            objCoder.Decode(keyFileName, FileName, a)
            SteganographyPull = a
            Kill(keyFileName)

        Catch exc As Exception
            'MsgBox(exc.ToString)
        Finally
            Try
                objCoder.Dispose()
            Catch
            End Try
        End Try

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
    Public Function ActivateTrial(ByVal SoftwareName As String, ByVal SoftwareVer As String, ByVal TrialType As Integer, ByVal TrialLength As Integer, ByVal TrialHideTypes As IActiveLock.ALTrialHideTypes, ByRef strMsg As String, ByVal SoftwarePassword As String, ByVal mCheckTimeServerForClockTampering As IActiveLock.ALTimeServerTypes, ByVal mChecksystemfilesForClockTampering As IActiveLock.ALSystemFilesTypes, ByVal mTrialWarning As IActiveLock.ALTrialWarningTypes, ByRef mRemainingTrialDays As Integer, ByRef mRemainingTrialRuns As Integer) As Boolean
        On Error GoTo NotRegistered
        Dim strVal As String
        Dim daysLeft, runsLeft As Short
        Dim intEXPIREDWARNING As Short

        EXPIRED_RUNS = Chr(101) & Chr(120) & Chr(112) & Chr(105) & Chr(114) & Chr(101) & Chr(100) & Chr(114) & Chr(117) & Chr(110) & Chr(115)
        EXPIRED_DAYS = Chr(101) & Chr(120) & Chr(112) & Chr(105) & Chr(114) & Chr(101) & Chr(100) & Chr(100) & Chr(97) & Chr(121) & Chr(115)
        TEXTMSG_DAYS = DecryptString128Bit("sQvYYRLPon5IyH6BQRAUBuCLTq/5VkH3kl7HUwJLZ2M=", PSWD)
        TEXTMSG_RUNS = DecryptString128Bit("6urN2+xbgqbLLsOoC4hbGpLT3bnvY3YPGW299cOnqfo=", PSWD)
        TEXTMSG = DecryptString128Bit("7MKOm40JXXXc6f3svOrlNeFuWWdWD56E3k7FwvTH2i/oyG2MDdVKjLMcQKsNbfXforIQwFlJEEgCMOfWiSI0sw==", PSWD)

        ActivateTrial = False
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        LICENSE_SOFTWARE_NAME = SoftwareName
        LICENSE_SOFTWARE_VERSION = SoftwareVer
        LICENSE_SOFTWARE_PASSWORD = SoftwarePassword

        ' Set local variables
        If TrialType = IActiveLock.ALTrialTypes.trialDays Then
            trialPeriod = True
        Else
            trialRuns = True
        End If
        If TrialType = IActiveLock.ALTrialTypes.trialDays Then
            alockDays = TrialLength
        Else
            alockRuns = TrialLength
        End If

        If alockDays = 0 And trialPeriod = True Then
            Change_Culture("")
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrInvalidTrialDays, ACTIVELOCKSTRING, STRINVALIDTRIALDAYS)
        ElseIf alockRuns = 0 And trialRuns = True Then
            Change_Culture("")
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrInvalidTrialRuns, ACTIVELOCKSTRING, STRINVALIDTRIALRUNS)
        End If

        strMsg = ""
        intEXPIREDWARNING = Int(CDbl(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "1"), enc2(TRIALWARNING), enc2(EXPIREDWARNING), CStr(0))))

        On Error GoTo keepChecking
        HAD2HAMMER = False
        GetSystemTime1()

        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor

        ' Check to see if any of the hidden signatures say the trial is expired
        ' The following two keys are not compatible with Vista
        ' A regular user account cannot have write access to these two registry hives
        ' I am removing these from v3.6 - ialkan 12-27-2008
        'If IsRegistryExpired1() = True Then
        '    Change_Culture("")
        '    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG)
        'End If
        'If IsRegistryExpired2() = True Then
        '    Change_Culture("")
        '    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG)
        'End If
        If IsEncryptedFileExpired() = True Then
            Change_Culture("")
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG)
        End If
        ' *** We are disabling folder date stamp in v3.2 since it's not application specific ***
        ' Well... nothing was found
        ' Check the last indicator
        'If IsFolderStampExpired() = True Then
        '    Change_Culture("")
        '    Err.Raise -10100, , TEXTMSG
        'End If

        ' Must check Registry for Trial
        If TrialRegistryPerUserExists(TrialHideTypes) Then
            If IsRegistryExpired() = True Then
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG)
            End If
        End If

        ' Must check picture for Trial
        If TrialSteganographyExists(TrialHideTypes) Then
            If IsSteganographyExpired() = True Then
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG)
            End If
        End If

        ' Must check folder for Trial
        If TrialHiddenFolderExists(TrialHideTypes) Then
            If IsHiddenFolderExpired() = True Then
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG)
            End If
        End If

        ' Nothing bad so far...
        If trialPeriod Then
            If Not DateGood(alockDays, daysLeft, TrialHideTypes) Then
                ExpireTrial(SoftwareName, SoftwareVer, TrialType, TrialLength, TrialHideTypes, SoftwarePassword)
                ' Trial Period has expired
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS)
            Else
                If fileExist(GetSteganographyFile()) = False And Directory.Exists(ActivelockGetSpecialFolder(46) & DecryptString128Bit(myDir, PSWD)) = False And dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90")) = dec2("93.8D.93.8D.96.90.90.90") Then
                    If mCheckTimeServerForClockTampering = IActiveLock.ALTimeServerTypes.alsCheckTimeServer Then
                        If SystemClockTampered() Then
                            Change_Culture("")
                            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                        End If
                    End If
                    If mChecksystemfilesForClockTampering = IActiveLock.ALSystemFilesTypes.alsCheckSystemFiles Then
                        If ClockTampering() Then
                            Change_Culture("")
                            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                        End If
                    End If
                End If
                ' So far so good; trial mode seems to be fine
                HAD2HAMMER = False
                strMsg = "You are running this program in its Trial Period Mode." & vbCrLf & CStr(daysLeft) & " day(s) left out of " & CStr(alockDays) & " day trial."
                mRemainingTrialDays = daysLeft
                ActivateTrial = True
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                GoTo exitGracefully
            End If
        Else
            If Not RunsGood(alockRuns, runsLeft, TrialHideTypes) Then
                ExpireTrial(SoftwareName, SoftwareVer, TrialType, TrialLength, TrialHideTypes, SoftwarePassword)
                ' Trial Runs have expired
                Change_Culture("")
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS)
            Else
                If fileExist(GetSteganographyFile()) = False And Directory.Exists(ActivelockGetSpecialFolder(46) & DecryptString128Bit(myDir, PSWD)) = False And dec2(GetSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90")) = dec2("93.8D.93.8D.96.90.90.90") Then
                    If mCheckTimeServerForClockTampering = IActiveLock.ALTimeServerTypes.alsCheckTimeServer Then
                        If SystemClockTampered() Then
                            Change_Culture("")
                            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                        End If
                    End If
                    If ClockTampering() Then
                        Change_Culture("")
                        Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                    End If
                End If
                ' So far so good; trial mode seems to be fine
                HAD2HAMMER = False
                strMsg = "You are running this program in its Trial Runs Mode." & vbCrLf & CStr(runsLeft) & " run(s) left out of " & CStr(alockRuns) & " run trial."
                mRemainingTrialRuns = runsLeft
                ActivateTrial = True
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                GoTo exitGracefully
            End If
        End If

keepChecking:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        ExpireTrial(SoftwareName, SoftwareVer, TrialType, TrialLength, TrialHideTypes, SoftwarePassword)
        If Err.Number = -10101 Then
            strMsg = TEXTMSG_DAYS
            mRemainingTrialDays = alockDays
        ElseIf Err.Number = -10102 Then
            strMsg = TEXTMSG_RUNS
            mRemainingTrialRuns = alockRuns
        End If
        If intEXPIREDWARNING = 0 Or mTrialWarning = IActiveLock.ALTrialWarningTypes.trialWarningPersistent Then
            Call SaveSetting(enc2(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "1"), enc2(TRIALWARNING), enc2(EXPIREDWARNING), CStr(-1))
            strMsg = "Free Trial for this application has ended."
        End If
        ActivateTrial = False
        Exit Function

NotRegistered:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        If Err.Number = 1001 Then
            Err.Clear()
            Resume Next
        End If
        strMsg = Err.Description
        HAD2HAMMER = True
        ActivateTrial = False
        Exit Function
exitGracefully:
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Function

    End Function
    Public Function ClockTampering() As Boolean
        Dim t As String = Nothing
        Dim s As String = Nothing
        Dim fileDate As Date
        Dim i, Count As Short
        On Error Resume Next

        For i = 0 To 1
            Select Case i
                Case 0
                    t = WinDir() & "\Prefetch"
                Case 1
                    t = WinDir() & "\Temp"
            End Select

            Count = 0
            s = Dir(t & "\*.*")
            Do While s <> ""
                If Left(s, 1) <> "$" And Left(s, 1) <> "?" Then
                    fileDate = FileDateTime(t & "\" & s)
                    Dim difHours As Long
                    '* difHours = CDate(fileDate.Date.ToString("yyyy/MM/dd")).Subtract(CDate(Date.UtcNow.ToString("yyyy/MM/dd"))).Hours
                    difHours = fileDate.Subtract(Date.UtcNow).Hours '*
                    If difHours > 24 Then
                        If Count > 1 Then
                            ClockTampering = True
                            Exit Function
                        Else
                            'Forgiveness for one file only - for now
                            Count = Count + 1
                        End If
                    End If
                End If
                s = Dir()
            Loop
        Next i

    End Function
    Private Function GetSteganographyFile() As Object

        Dim strSource As String
        Dim commonPicsFolder As String

        GetSteganographyFile = Nothing
        Try
            commonPicsFolder = ActivelockGetSpecialFolder(54)

            If Directory.Exists(commonPicsFolder) = False Then
                ' Unable to retrieve common pics folder, so we generate one
                commonPicsFolder = ActivelockGetSpecialFolder(46) ' common documents

                ' Note that the common pictures folder is different in
                ' Windows XP/2003 compare to Vista/2008.  This also means
                ' that the name can change again in future versions of
                ' Windows and this section of code will need to be modified.
                Dim version_info As New CWindows.OperatingSystemVersion()
                If (version_info.IsWinVistaPlus()) Then
                    commonPicsFolder &= "\Pictures"
                Else
                    commonPicsFolder &= "\My Pictures"
                End If

                My.Computer.FileSystem.CreateDirectory(commonPicsFolder)
            End If

            strSource = commonPicsFolder & "\Sample Pictures" & DecryptString128Bit("Qspq9Tu3sG/IE+ugm+o1RQ==", PSWD) & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(98) & Chr(109) & Chr(112)
            Dim b As System.Drawing.Bitmap
            Dim r As New System.Resources.ResourceManager("ActiveLock3_6Net.ProjectResources", System.Reflection.Assembly.GetExecutingAssembly)
            b = r.GetObject("bmp101")
            If Directory.Exists(commonPicsFolder & "\Sample Pictures") = False Then
                My.Computer.FileSystem.CreateDirectory(commonPicsFolder & "\Sample Pictures")
            End If
            If fileExist(strSource) = False Then
                b.Save(strSource)
            End If
            'If fileExist(strSource) = False Then VB6.LoadResPicture(101, VB6.LoadResConstants.ResBitmap).Save(strSource)
            GetSteganographyFile = strSource
        Catch ex As Exception
            Change_Culture("")
            Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
        End Try

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
    Public Function ReadUntil(ByRef sIn As String, ByRef sDelim As String, Optional ByRef bCompare As CompareMethod = CompareMethod.Binary) As String
        Dim nPos As String
        ReadUntil = String.Empty

        nPos = CStr(InStr(1, sIn, sDelim, bCompare))
        If CDbl(nPos) > 0 Then
            ReadUntil = Left(sIn, CDbl(nPos) - 1)
            sIn = Mid(sIn, CDbl(nPos) + Len(sDelim))
        End If
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
    'InStrRev was upgraded to InStrRev_Renamed
    Public Function InStrRev_Renamed(ByVal sIn As String, ByRef sFind As String, Optional ByRef nStart As Integer = 1, Optional ByRef bCompare As CompareMethod = CompareMethod.Binary) As Integer
        Dim nPos As Integer
        sIn = StrReverse_Renamed(sIn)
        sFind = StrReverse_Renamed(sFind)
        nPos = InStr(nStart, sIn, sFind, bCompare)
        If nPos = 0 Then
            InStrRev_Renamed = 0
        Else
            InStrRev_Renamed = Len(sIn) - nPos - Len(sFind) + 2
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
    'StrReverse was upgraded to StrReverse_Renamed
    Public Function StrReverse_Renamed(ByVal sIn As String) As String
        Dim nC As Short
        Dim sOut As String = Nothing
        For nC = Len(sIn) To 1 Step -1
            sOut = sOut & Mid(sIn, nC, 1)
        Next
        StrReverse_Renamed = sOut
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
    'Replace was upgraded to Replace_Renamed
    Public Function Replace_Renamed(ByRef sIn As String, ByRef sFind As String, ByRef sReplace As String, Optional ByRef nStart As Integer = 1, Optional ByRef nCount As Integer = -1, Optional ByRef bCompare As CompareMethod = CompareMethod.Binary) As String
        Dim nC As Integer
        Dim nPos As Short
        Dim sOut As String
        sOut = sIn
        nPos = InStr(nStart, sOut, sFind, bCompare)
        If nPos = 0 Then GoTo EndFn
        Do
            nC = nC + 1
            sOut = Left(sOut, nPos - 1) & sReplace & Mid(sOut, nPos + Len(sFind))
            If nCount <> -1 And nC >= nCount Then Exit Do
            nPos = InStr(nStart, sOut, sFind, bCompare)
        Loop While nPos > 0
EndFn:
        Replace_Renamed = sOut
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
        Dim lngpos As Integer
        Do While InStr(1, strString, " ")
            System.Windows.Forms.Application.DoEvents()
            lngpos = InStr(1, strString, " ")
            strString = Left(strString, lngpos - 1) & Right(strString, Len(strString) - (lngpos + Len(" ") - 1))
        Loop
        TrimSpaces = strString
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
        Dim i As Short
        Dim even As String = String.Empty
        Dim odd As String = String.Empty
        For i = 1 To Len(strString)
            If i Mod 2 = 0 Then
                even = even & Mid(strString, i, 1)
            Else
                odd = odd & Mid(strString, i, 1)
            End If
        Next i
        Scramb = even & odd
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
        Dim intPos As Short
        dhTrimNull = String.Empty

        intPos = InStr(strValue, vbNullChar)
        Select Case intPos
            Case 0
                dhTrimNull = strValue
            Case 1
                dhTrimNull = vbNullString
            Case Is > 1
                dhTrimNull = Left(strValue, intPos - 1)
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
        Dim evenint, x, oddint As Short
        Dim odd, even As String
        Dim fin As String = String.Empty

        x = Len(strString)
        x = Int(Len(strString) / 2) 'adding this returns the actual number like 1.5 instead of returning 2
        even = Mid(strString, 1, x)
        odd = Mid(strString, x + 1)
        For x = 1 To Len(strString)
            If x Mod 2 = 0 Then
                evenint = evenint + 1
                fin = fin & Mid(even, evenint, 1)
            Else
                oddint = oddint + 1
                fin = fin & Mid(odd, oddint, 1)
            End If
        Next x
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
        GetWindowsDirectory(sPath, 255)
        WindowsPath = dhTrimNull(sPath)
        If Right(WindowsPath, 1) <> "\" Then
            WindowsPath = WindowsPath & "\"
        End If
    End Function
    '===============================================================================
    ' Name: Function DoScan
    ' Input: None
    ' Output:
    '   Boolean - True if a debugger or file monitor is found
    ' Purpose: Will scan the memory for common debuggers or File Monitoring software
    ' Remarks: None
    '===============================================================================
    Public Function DoScan() As Boolean

        Dim hFile, retVal As Integer
        Dim sScan As String
        ' Dim sBuffer As String

        Dim sRegMonClass, sFileMonClass As String
        '\\We break up the class names to avoid
        '     detection in a hex editor
        sRegMonClass = "R" & "e" & "g" & "m" & "o" & "n" & "C" & "l" & "a" & "s" & "s"
        sFileMonClass = "F" & "i" & "l" & "e" & "M" & "o" & "n" & "C" & "l" & "a" & "s" & "s"
        '\\See if RegMon or FileMon are running

        Select Case True
            Case FindWindow(sRegMonClass, vbNullString) <> 0
                'Regmon is running...throw an access vio
                '     lation
                DoScan = True
                Exit Function
            Case FindWindow(sFileMonClass, vbNullString) <> 0
                'FileMon is running...throw an access vi
                '     olation
                DoScan = True
                Exit Function
        End Select
        '\\So far so good...check for SoftICE in memory
        hFile = CreateFile("\\.\SICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

        If hFile <> -1 Then
            ' SoftICE is detected.
            retVal = CloseHandle(hFile) ' Close the file handle
            DoScan = True
            Exit Function
        Else
            ' SoftICE is not found for windows 9x, check for NT.
            hFile = CreateFile("\\.\NTICE", GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

            If hFile <> -1 Then
                ' SoftICE is detected.
                retVal = CloseHandle(hFile) ' Close the file handle
                DoScan = True
                Exit Function
            End If
        End If

        sScan = "f" & "i" & "l" & "e" & "W" & "A" & "T" & "C" & "H"

        If FindWindow(vbNullString, sScan) <> 0 Then
            DoScan = True
            Exit Function
        End If

        sScan = "F" & "i" & "l" & "e" & "S" & "p" & "y"
        If FindWindow(vbNullString, sScan) <> 0 Then
            DoScan = True
            Exit Function
        End If
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
        Dim sFc, src, str1 As String
        Dim vernumber As String = String.Empty
        Dim vn2, vn0, vn1, vnx As Short

        src = DecryptString128Bit("kSPnKBkg9LARQZgo61A3ow==", PSWD) 'RegmonClass
        sFc = DecryptString128Bit("SWRcv8sZJnnQDZHkS+Ewnw==", PSWD) 'FileMonClass

        'Check For RegMon
        If FinWin(src, vbNullString) <> 0 Then
            HAD2HAMMER = True
            Exit Sub
        End If
        If FinWin(sFc, vbNullString) <> 0 Then
            HAD2HAMMER = True
            Exit Sub
        End If

        'Look For Threats via VxD..
        CTV(DecryptString128Bit("bVW9MtZbzBzH3f8HeSCO9g==", PSWD)) 'SICE
        CTV(DecryptString128Bit("Jf4nLowTbc8fqW0lTNQ/6Q==", PSWD)) 'NTICE
        CTV(DecryptString128Bit("OO6LYadZGIJ57F7HphxBPw==", PSWD)) 'SIWDEBUG
        CTV(DecryptString128Bit("Bv5XEuseX6OXpMbpZbV+cg==", PSWD)) 'SIWVID

        'Look For Threats using titles of windows !!!!!!!!!!!!!!!!!!!

        'W32dasm (other than main window)
        CTW(DecryptString128Bit("aDHIGwWBTB0qErkuL9yzyd3jTfPbCStMz2Kpd8WgMq8=", PSWD)) 'Win32Dasm "Goto Code Location (32 Bit)"
        'SoftICE variants
        CTW(My.Application.Info.DirectoryPath & "\" & System.IO.Path.GetFileName(Application.ExecutablePath) & ".EXE" & DecryptString128Bit("28u3ww4pEzuawNHeN8kJWscUrblvk5v4Af2gaVLRD0M=", PSWD)) 'SoftIce; [app_path]+" - Symbolic Loader"
        CTW(My.Application.Info.DirectoryPath & "\" & System.IO.Path.GetFileName(Application.ExecutablePath) & ".EXE" & DecryptString128Bit("jk980I+TMZFGIcGEFPSxCJydPCzofApnewH3ORNKRQI=", PSWD)) 'SoftIce; [app_path]+" - Symbol Loader"
        'CTW(VB6.GetPath & "\" & VB6.GetEXEName() & ".EXE" & DecryptString128Bit("28u3ww4pEzuawNHeN8kJWscUrblvk5v4Af2gaVLRD0M=", PSWD)) 'SoftIce; [app_path]+" - Symbolic Loader"
        'CTW(VB6.GetPath & "\" & VB6.GetEXEName() & ".EXE" & DecryptString128Bit("jk980I+TMZFGIcGEFPSxCJydPCzofApnewH3ORNKRQI=", PSWD)) 'SoftIce; [app_path]+" - Symbol Loader"
        CTW(DecryptString128Bit("4WuQVAq9HLlmPg1lsTQN7m9vintF1mbSLUERgq1nONo=", PSWD)) '"NuMega SoftICE Symbol Loader"

        'Checks for URSoft W32Dasm app windows versions 0.0x - 12.9x
        str1 = DecryptString128Bit("PhxMXRKD5iLJ6eNYZBKeg3BRdQyRx+fa0S1RijBZRJ4=", PSWD) & vernumber & DecryptString128Bit("SOsZIVbDRg8v0mcr7pD9w1xrKHCI3xFpof9LET3zFv4=", PSWD)
        For vn0 = 12 To 0 Step -1
            For vn1 = 9 To 0 Step -1
                For vn2 = 9 To 0 Step -1
                    vnx = CShort(vn1 & vn2)
                    vernumber = vn0 & "." & vnx
                    'Check for "URSoft W32Dasm Ver " & vernumber & " Program Disassembler/Debugger"
                    CTW(str1)
                Next vn2
            Next vn1
        Next vn0

        'Check for step debugging (light check)
        CSD()

        'Check for processes and wipe from 200000 to N amount of bytes in steps of 48
        '(to aggressively screw with the code)
        'RefreshProcessList()
        CFP(DecryptString128Bit("OiOALAEiNTL0BDcDFRa5+A==", PSWD), 2000000) 'Kill "Debuggy By Vanja Fuckar" - Debuggy.exe
        CFP(DecryptString128Bit("3idyYtU1s0wY5HvE9AztZQ==", PSWD), 2000000) 'Kill "OllyDBG" - OLLYDBG.exe
        CFP(DecryptString128Bit("2UqaM1jVTgltRHWy59DcyA==", PSWD), 2000000) 'Kill "ProcDump by G-Rom, Lorian & Stone" - PROCDUMP.exe
        CFP(DecryptString128Bit("cXf/bBlwI4BzR2+996P6rw==", PSWD), 2000000) 'Kill "SoftSnoop by Yoda/f2f" - SoftSnoop.exe
        CFP(DecryptString128Bit("kjLerFCcQQ3qcRsJNjb5qA==", PSWD), 2000000) 'Kill "TimeFix by GodsJiva" - TimeFix.exe
        CFP(DecryptString128Bit("lCu+rzLgHvPml0ceQaYnc6Ga/wmZDry4uyL8UmPQVdk=", PSWD), 2000000) 'Kill "TMR Ripper Studi" - "TMG Ripper Studio.exe"

        'Send the user through a jungle of conditional branches.
        'Hopefully now timefix will be disabled.
        JOC()

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

        If GC() = 27514 Then HAD2HAMMER = False
        ' One final check
        If DoScan() = True Then HAD2HAMMER = True

    End Sub
    '===============================================================================
    ' Name: Sub CTV
    ' Input:
    '   ByRef appid as String
    ' Output: None
    ' Purpose: Checks threats vxd
    ' Remarks: None
    '===============================================================================
    Public Sub CTV(ByRef appid As String)
        'Check threats vxd
        If CreateFile("\\.\" & appid, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0) <> -1 Then
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
    Public Sub CFP(ByRef procname As String, ByRef hammerrange As Object)
        Dim xx As Short
        For xx = 0 To 256
            If LCase(procname) = LCase(ProcessName(xx)) Then HAMMERPROCESS(CInt(ProcessID(xx)), hammerrange)
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
    Public Sub HAMMERPROCESS(ByRef PID As Integer, ByRef hammertop As Object)
        If Not InitProcess(PID) Then
            'MsgBox "Failed shutdown"
        End If
        Dim p, addr, l As Integer
        For p = 20000 To hammertop Step 48
            addr = CInt(Val((CStr(p)).Trim))
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
    Public Sub CTW(ByRef winid As String)
        Dim WID As Integer
        Dim FLDWIN As Short
        WID = FindWindow(vbNullString, winid)
        If FindWindow(vbNullString, winid) > 0 Then
            'Just sending &H10 closes the window.. but this method freezes it and closes apps where they are usually protected from an external shut down!! ;)
            For FLDWIN = 0 To 255
                PostMessage(WID, FLDWIN, 0, 0)
                If FLDWIN > 16 Then
                    PostMessage(WID, &H10S, 0, 0)
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
        Dim Timer_start, Timer_time As Single
        Dim s As Short
        'Check for Step Debugger
        Timer_start = Microsoft.VisualBasic.Timer()
        For s = 1 To 25
            PSub() 'Pointless Sub
            PFunction(s + Int(Rnd() * 20)) 'Pointless Function
        Next s
        Timer_time = Microsoft.VisualBasic.Timer() - Timer_start

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
        Dim X2, X1, X3 As Object
        'Just some garbage processing...
        System.Windows.Forms.Application.DoEvents()
        X1 = System.Math.Sqrt(65536)
        X2 = 16 ^ 2
        X3 = X1 - X2
        X1 = X2 + X3
        X3 = X2
    End Sub
    '===============================================================================
    ' Name: Function PFunction
    ' Input:
    '   ByRef PointlessVariable As Integer - Dummy integer
    ' Output: None
    ' Purpose: Processes some garbage
    ' Remarks: None
    '===============================================================================
    Public Function PFunction(ByRef PointlessVariable As Short) As Object
        Dim X2, X1, X3 As Object
        'Just some garbage processing...
        System.Windows.Forms.Application.DoEvents()
        X1 = System.Math.Sqrt(256)
        X2 = 8 ^ 2
        X3 = X1 + PointlessVariable
        X1 = X1 + X2 + X3

        PFunction = Nothing
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
    ' Remarks: None
    '===============================================================================
    Public Sub JOC()

        'Horrible Sloppy Code but it should help to throw some lamers off..
        'Start off with some fake math, arrays, etc and throw a few pointless encrypted strings in there
        Randomize(32)
        Dim JU_C_OR(32) As Object
        Dim ViT(1, 1) As Object
        Dim c, AMIN, tang, App_Les As Object
        Dim yergisiz, ang, gerinmeoyle, PE_ar As Object
        Dim e As Object = Nothing
        ' Dim TXM, TM, TXD, TXZ As Single

        ViT(0, 0) = "yalnizlikzorseydostum"
        AMIN = 1
        tang = 12
        c = 0
        gerinmeoyle = "eT5XeXeXXeT1MeUX11cU"
        ang = tang - 2
        App_Les = Int(Rnd() * 6)
        c = 1 - c
        PE_ar = Int(Rnd() * 100) + Asc(Mid("saygisizbiradam", 5, 1))
        AMIN = 1 - AMIN
        JU_C_OR(ang + e) = CInt(ViT(AMIN, c))
        yergisiz = PE_ar & JU_C_OR(ang + e) & App_Les & ("yalnizlikzor" & gerinmeoyle)

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

        '5:
        '        TM = VB.Timer()
        '10:
        '        If AMIN = 0 Then GoTo 30 Else GoTo 70
        '25:     System.Windows.Forms.Application.DoEvents()
        '20:
        '        If PE_ar > AMIN Then GoTo 25 Else GoTo 30
        '30:
        '        If c = 2 Then wX = -800 : GoTo 60 Else GoTo 40
        '40:
        '        c = c + 1
        '        TXD = VB.Timer() : GoTo 60
        '50:
        '        If PE_ar + ang = AMIN Then GoTo 40 Else GoTo 95
        '60:
        '        AMIN = 0 : wY = -2000
        '        If TXD - TM < c Then GoTo 80 Else GoTo 170
        '800:
        '        If frmC.DefInstance.Timer1.Enabled = True Then
        '            TXM = 1
        '        Else
        '            TXM = 0 : GoTo 890
        '        End If
        '70:
        '        AMIN = AMIN + 1 : GoTo 20
        '80:
        '        If AMIN > 1 Then GoTo 20
        '190:    wX = 4800 : wY = 3600 : GoTo 1000 'here we set the window width so that it's no longer 0,0
        '75:
        '        If App_Les > 16 Then
        '            GoTo 190
        '        Else
        '            If App_Les > 256 Then GoTo 140
        '        End If
        '1000:
        '        If VB.Timer() > TM + 2 Then wY = 0 Else Call frmC.DefInstance.frmC_Resize(frmC.DefInstance, New System.EventArgs) : Exit Sub
        '1001:   Call frmC.DefInstance.frmC_Resize(frmC.DefInstance, New System.EventArgs) : Exit Sub
        '90:     GoTo 80
        '95:     wY = 50 : GoTo 800
        '140:    If wY = 360 Then
        '            wY = 0
        '            TM = 20 : GoTo 60
        '        End If
        '125:    GoTo 150
        '890:
        '        wX = wX * TXM
        '        wY = wY * TXM : GoTo 1000
        '120:    GoTo 1000
        '170:
        '        TXZ = TXZ + 1
        '        If TXZ > 50 Then GoTo 135 Else GoTo 175
        '175:    If frmC.DefInstance.Timer1.Enabled = False Then GoTo 135 Else GoTo 170
        '160:    If wX = 30 Then GoTo 170 Else GoTo 20
        '150:    If wX = 0 Then GoTo 170 Else GoTo 130
        '130:    GoTo 95
        '135:    HAD2HAMMER = True
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
    Function InitProcess(ByRef PID As Integer) As Object
        Dim pHandle, myHandle As Integer
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

        Dim myProcess As PROCESSENTRY32 = Nothing
        Dim mySnapshot As Integer
        Dim xx As Short

        'first clear our combobox
        myProcess.dwSize = Len(myProcess)

        'create snapshot
        mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)

        'clear array
        For xx = 0 To 256
            ProcessName(xx) = ""
        Next xx

        xx = 0
        'get first process
        ProcessFirst(mySnapshot, myProcess)
        ProcessName(xx) = Left(myProcess.szexeFile, InStr(1, myProcess.szexeFile, Chr(0)) - 1) ' set exe name
        ProcessID(xx) = myProcess.th32ProcessID ' set PID

        'while there are more processes
        While ProcessNext(mySnapshot, myProcess)
            xx = xx + 1
            ProcessName(xx) = Left(myProcess.szexeFile, InStr(1, myProcess.szexeFile, Chr(0)) - 1) ' set exe name
            ProcessID(xx) = myProcess.th32ProcessID ' set PID
        End While

    End Sub
    '===============================================================================
    ' Name: Function GC
    ' Input: None
    ' Output:
    '   Variant
    ' Purpose: Get CRC of all strings to check if they've been modified
    ' Remarks: None
    '===============================================================================
    Public Function GC() As Object
        Dim p, encvars, mycrc As Short
        'Get CRC of all strings to check if they've been modified
        For encvars = 0 To 4000
            For p = 1 To Len(encvar(encvars))
                mycrc = mycrc + Asc(Mid(encvar(encvars), p, 1))
                If mycrc > 30000 Then mycrc = mycrc - 30000
            Next p
        Next encvars
        GC = mycrc
    End Function
    Private Sub CreateHdnFile()
        Dim mINIFile As New INIFile
        If fileExist(HiddenFile()) = False Then
            With mINIFile
                .File = HiddenFile()
                .Section = ".ShellClassInfo"
                .Values("CLSID") = CLSIDSTR
            End With
        End If
    End Sub
    Private Function HiddenFolderFunction() As String
        Dim myFile As String
        myFile = Chr(92) & Chr(82) & Chr(101) & Chr(99) & Chr(121) & Chr(99) & Chr(108) & Chr(101) & Chr(100) & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8) & "." & Chr(108) & Chr(115) & Chr(116)
        HiddenFolderFunction = ActivelockGetSpecialFolder(46) & DecryptString128Bit(HIDDENFOLDER, PSWD) & myFile
    End Function
    Private Function HiddenFile() As String
        Dim KEYFILE As String
        KEYFILE = Chr(92) & Chr(68) & Chr(101) & Chr(115) & Chr(107) & Chr(116) & Chr(111) & Chr(112) & "." & Chr(105) & Chr(110) & Chr(105)
        HiddenFile = ActivelockGetSpecialFolder(46) & DecryptString128Bit(HIDDENFOLDER, PSWD) & KEYFILE
    End Function
    Private Sub MinusAttributes()
        On Error GoTo minusAttributesError
        Dim ok As Double
        'Ok, the folder is there, let's change its attributes
        ok = Shell("ATTRIB -h -s -r " & ActivelockGetSpecialFolder(46) & DecryptString128Bit(HIDDENFOLDER, PSWD), AppWinStyle.Hide)
minusAttributesError:
    End Sub
    Private Sub PlusAttributes()
        Dim ok As Double
        'Ok, the folder is there, let's change its attributes
        ok = Shell("ATTRIB +h +s +r " & ActivelockGetSpecialFolder(46) & DecryptString128Bit(HIDDENFOLDER, PSWD), AppWinStyle.Hide)
    End Sub
    Public Function SystemClockTampered() As Boolean

        ' Section added by Ismail Alkan
        ' Access a good time server to see which day it is :)
        ' Get the date only... compare with the system clock
        ' Die if more than 1 day difference'
        ' UTC Time and Date of course...

        ' Obviously, for this function to work, there must be a connection to Internet
        If IsWebConnected() = False Then
            SystemClockTampered = False
            Exit Function
            'Change_Culture("")
            'Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrInternetConnectionError, ACTIVELOCKSTRING, STRINTERNETNOTCONNECTED)
        End If

        Dim ss, aa As String
        Dim blabla As String
        Dim diff As Short
        Dim i As Short
        Dim month1() As String
        Dim month2() As String
        SystemClockTampered = False
        month1 = "January;February;March;April;May;June;July;August;September;October;November;December".Split(";")
        month2 = "01;02;03;04;05;06;07;08;09;10;11;12".Split(";")
        ss = OpenURL(DecryptString128Bit("xcZfCWJLnPOl2V5kTJpBOi4ysgfzBj1H3nUzyhxODITU7s9QX0xe23TQ9ue3ypmT", PSWD)) 'http://www.time.gov/timezone.cgi?UTC/s/0
        If ss = "" Then
            ' The fast method above did not do its job
            ' Call an existing Daytime class from "The Code Project"
            ' to check if the system clock was adjusted
            Dim systemClock As New Daytime
            If Daytime.WindowsClockIncorrect = True Then
                Return True
            End If
        End If
        blabla = "</b></font><font size=" & """" & "5" & """" & " color=" & """" & "white" & """" & ">"
        i = InStr(ss, blabla)
        ss = Mid(ss, i + Len(blabla))
        i = InStr(ss, "<br>")
        ss = Left(ss, i - 1)
        i = InStr(1, ss, ",")
        ss = Mid(ss, i + 1)
        ss = Replace_Renamed(ss, ",", " ")
        ss = ss.Trim
        For i = 0 To 11
            If InStr(ss, month1(i)) Then
                ss = ss.Replace(month1(i), month2(i))
                Exit For
            End If
        Next
        ' I am leaving this alone because it works - yes it's crazy
        '* ss = Format(CDate(ss), "yyyy/MM/dd")   '"short date")
        '* aa = Date.UtcNow.Year.ToString("0000") & "/" & Date.UtcNow.Month.ToString("00") & "/" & Date.UtcNow.Day.ToString("00")
        'Ok, I'll give it a try
        ss = CDate(ss).Date.ToString '*
        aa = Date.UtcNow.ToString '*
        '*diff = CDate(ss).Subtract(CDate(aa)).Days
        diff = CDate(ss).Date.Subtract(Date.UtcNow).Days '*

        If diff > 1 Then SystemClockTampered = True

    End Function
    Public Function VarPtr(ByVal o As Object) As Integer
        ' Undocumented VarPtr in VB.NET !!!
        Dim GC As System.Runtime.InteropServices.GCHandle = System.Runtime.InteropServices.GCHandle.Alloc(o, System.Runtime.InteropServices.GCHandleType.Pinned)
        Dim ret As Integer = GC.AddrOfPinnedObject.ToInt32
        GC.Free()
        Return ret
    End Function
    Public Function EncryptString128Bit(ByVal vstrTextToBeEncrypted As String, _
                                    ByVal vstrEncryptionKey As String) As String

        Dim bytValue() As Byte
        Dim bytKey() As Byte
        Dim bytEncoded() As Byte = Nothing
        Dim bytIV() As Byte = {121, 241, 10, 1, 132, 74, 11, 39, 255, 91, 45, 78, 14, 211, 22, 62}
        Dim intLength As Integer
        Dim intRemaining As Integer
        Dim objMemoryStream As New MemoryStream
        Dim objCryptoStream As CryptoStream
        Dim objRijndaelManaged As RijndaelManaged


        '   **********************************************************************
        '   ******  Strip any null character from string to be encrypted    ******
        '   **********************************************************************

        vstrTextToBeEncrypted = StripNullCharacters(vstrTextToBeEncrypted)

        '   **********************************************************************
        '   ******  Value must be within ASCII range (i.e., no DBCS chars)  ******
        '   **********************************************************************

        bytValue = Encoding.ASCII.GetBytes(vstrTextToBeEncrypted.ToCharArray)

        intLength = Len(vstrEncryptionKey)

        '   ********************************************************************
        '   ******   Encryption Key must be 256 bits long (32 bytes)      ******
        '   ******   If it is longer than 32 bytes it will be truncated.  ******
        '   ******   If it is shorter than 32 bytes it will be padded     ******
        '   ******   with upper-case Xs.                                  ****** 
        '   ********************************************************************

        If intLength >= 32 Then
            vstrEncryptionKey = Strings.Left(vstrEncryptionKey, 32)
        Else
            intLength = Len(vstrEncryptionKey)
            intRemaining = 32 - intLength
            vstrEncryptionKey = vstrEncryptionKey & Strings.StrDup(intRemaining, "X")
        End If

        bytKey = Encoding.ASCII.GetBytes(vstrEncryptionKey.ToCharArray)

        objRijndaelManaged = New RijndaelManaged

        '   ***********************************************************************
        '   ******  Create the encryptor and write value to it after it is   ******
        '   ******  converted into a byte array                              ******
        '   ***********************************************************************

        Try
            objCryptoStream = New CryptoStream(objMemoryStream, _
              objRijndaelManaged.CreateEncryptor(bytKey, bytIV), _
              CryptoStreamMode.Write)
            objCryptoStream.Write(bytValue, 0, bytValue.Length)

            objCryptoStream.FlushFinalBlock()

            bytEncoded = objMemoryStream.ToArray
            objMemoryStream.Close()
            objCryptoStream.Close()
        Catch

        End Try

        '   ***********************************************************************
        '   ******   Return encryptes value (converted from  byte Array to   ******
        '   ******   a base64 string).  Base64 is MIME encoding)             ******
        '   ***********************************************************************

        Return Convert.ToBase64String(bytEncoded)

    End Function
    Public Function DecryptString128Bit(ByVal vstrStringToBeDecrypted As String, _
                                        ByVal vstrDecryptionKey As String) As String

        Dim bytDataToBeDecrypted() As Byte
        Dim bytTemp() As Byte
        Dim bytIV() As Byte = {121, 241, 10, 1, 132, 74, 11, 39, 255, 91, 45, 78, 14, 211, 22, 62}
        Dim objRijndaelManaged As New RijndaelManaged
        Dim objMemoryStream As MemoryStream
        Dim objCryptoStream As CryptoStream
        Dim bytDecryptionKey() As Byte

        Dim intLength As Integer
        Dim intRemaining As Integer
        Dim strReturnString As String = String.Empty

        '   *****************************************************************
        '   ******   Convert base64 encrypted value to byte array      ******
        '   *****************************************************************

        bytDataToBeDecrypted = Convert.FromBase64String(vstrStringToBeDecrypted)

        '   ********************************************************************
        '   ******   Encryption Key must be 256 bits long (32 bytes)      ******
        '   ******   If it is longer than 32 bytes it will be truncated.  ******
        '   ******   If it is shorter than 32 bytes it will be padded     ******
        '   ******   with upper-case Xs.                                  ****** 
        '   ********************************************************************

        intLength = Len(vstrDecryptionKey)

        If intLength >= 32 Then
            vstrDecryptionKey = Strings.Left(vstrDecryptionKey, 32)
        Else
            intLength = Len(vstrDecryptionKey)
            intRemaining = 32 - intLength
            vstrDecryptionKey = vstrDecryptionKey & Strings.StrDup(intRemaining, "X")
        End If

        bytDecryptionKey = Encoding.ASCII.GetBytes(vstrDecryptionKey.ToCharArray)

        ReDim bytTemp(bytDataToBeDecrypted.Length)

        objMemoryStream = New MemoryStream(bytDataToBeDecrypted)

        '   ***********************************************************************
        '   ******  Create the decryptor and write value to it after it is   ******
        '   ******  converted into a byte array                              ******
        '   ***********************************************************************

        Try

            objCryptoStream = New CryptoStream(objMemoryStream, _
               objRijndaelManaged.CreateDecryptor(bytDecryptionKey, bytIV), _
               CryptoStreamMode.Read)

            objCryptoStream.Read(bytTemp, 0, bytTemp.Length)

            objCryptoStream.FlushFinalBlock()
            objMemoryStream.Close()
            objCryptoStream.Close()

        Catch

        End Try

        '   *****************************************
        '   ******   Return decypted value     ******
        '   *****************************************

        Return StripNullCharacters(Encoding.ASCII.GetString(bytTemp))

    End Function

    Public Function StripNullCharacters(ByVal vstrStringWithNulls As String) As String

        Dim intPosition As Integer
        Dim strStringWithOutNulls As String = ""

        If Not (vstrStringWithNulls Is Nothing) Then
            intPosition = 1
            strStringWithOutNulls = vstrStringWithNulls

            Do While intPosition > 0
                intPosition = InStr(intPosition, vstrStringWithNulls, vbNullChar)

                If intPosition > 0 Then
                    strStringWithOutNulls = Strings.Left(strStringWithOutNulls, intPosition - 1) & Strings.Right(strStringWithOutNulls, strStringWithOutNulls.Length - intPosition)
                End If

                If intPosition > strStringWithOutNulls.Length Then
                    Exit Do
                End If
            Loop
        End If
        Return strStringWithOutNulls

    End Function
    Public Function TrialRegistryPerUserExists(ByRef TrialHideTypes As IActiveLock.ALTrialHideTypes) As Boolean
        TrialRegistryPerUserExists = False
        If TrialHideTypes = IActiveLock.ALTrialHideTypes.trialRegistryPerUser Then
            TrialRegistryPerUserExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialHiddenFolder) Then
            TrialRegistryPerUserExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialSteganography) Then
            TrialRegistryPerUserExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage) Then
            TrialRegistryPerUserExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage) Then
            TrialRegistryPerUserExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialSteganography) Then
            TrialRegistryPerUserExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage) Then
            TrialRegistryPerUserExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage Or IActiveLock.ALTrialHideTypes.trialHiddenFolder) Then
            TrialRegistryPerUserExists = True
        End If
    End Function
    Public Function TrialHiddenFolderExists(ByRef TrialHideTypes As IActiveLock.ALTrialHideTypes) As Boolean
        TrialHiddenFolderExists = False
        If TrialHideTypes = IActiveLock.ALTrialHideTypes.trialHiddenFolder Then
            TrialHiddenFolderExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser) Then
            TrialHiddenFolderExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialSteganography) Then
            TrialHiddenFolderExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage) Then
            TrialHiddenFolderExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage) Then
            TrialHiddenFolderExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialSteganography) Then
            TrialHiddenFolderExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage) Then
            TrialHiddenFolderExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser) Then
            TrialHiddenFolderExists = True
        End If
    End Function
    Public Function TrialSteganographyExists(ByRef TrialHideTypes As IActiveLock.ALTrialHideTypes) As Boolean
        TrialSteganographyExists = False
        If TrialHideTypes = IActiveLock.ALTrialHideTypes.trialSteganography Then
            TrialSteganographyExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser) Then
            TrialSteganographyExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialHiddenFolder) Then
            TrialSteganographyExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage) Then
            TrialSteganographyExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage) Then
            TrialSteganographyExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialHiddenFolder) Then
            TrialSteganographyExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage) Then
            TrialSteganographyExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialIsolatedStorage Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser) Then
            TrialSteganographyExists = True
        End If
    End Function
    Public Function TrialIsolatedStorageExists(ByRef TrialHideTypes As IActiveLock.ALTrialHideTypes) As Boolean
        TrialIsolatedStorageExists = False
        If TrialHideTypes = IActiveLock.ALTrialHideTypes.trialIsolatedStorage Then
            TrialIsolatedStorageExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialIsolatedStorage Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser) Then
            TrialIsolatedStorageExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialIsolatedStorage Or IActiveLock.ALTrialHideTypes.trialHiddenFolder) Then
            TrialIsolatedStorageExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialIsolatedStorage Or IActiveLock.ALTrialHideTypes.trialSteganography) Then
            TrialIsolatedStorageExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialIsolatedStorage Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialSteganography) Then
            TrialIsolatedStorageExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialIsolatedStorage Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialHiddenFolder) Then
            TrialIsolatedStorageExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialIsolatedStorage Or IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialSteganography) Then
            TrialIsolatedStorageExists = True
        ElseIf TrialHideTypes = (IActiveLock.ALTrialHideTypes.trialIsolatedStorage Or IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialSteganography Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser) Then
            TrialIsolatedStorageExists = True
        End If
    End Function
    Public Function myDir() As String
        myDir = "mPY+Que6efQvkZsstJlvvw=="
    End Function
End Module