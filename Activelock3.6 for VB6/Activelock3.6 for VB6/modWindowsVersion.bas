Attribute VB_Name = "modWindowsVersion"
Option Explicit
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2006 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'dwPlatformId
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

'os product type values
Private Const VER_NT_WORKSTATION = &H1
Private Const VER_NT_DOMAIN_CONTROLLER = &H2
Private Const VER_NT_SERVER = &H3

'product types
Private Const VER_SERVER_NT = &H80000000
Private Const VER_WORKSTATION_NT = &H40000000

Private Const VER_SUITE_SMALLBUSINESS = &H1
Private Const VER_SUITE_ENTERPRISE = &H2
Private Const VER_SUITE_BACKOFFICE = &H4
Private Const VER_SUITE_COMMUNICATIONS = &H8
Private Const VER_SUITE_TERMINAL = &H10
Private Const VER_SUITE_SMALLBUSINESS_RESTRICTED = &H20
Private Const VER_SUITE_EMBEDDEDNT = &H40
Private Const VER_SUITE_DATACENTER = &H80
Private Const VER_SUITE_SINGLEUSERTS = &H100
Private Const VER_SUITE_PERSONAL = &H200
Private Const VER_SUITE_BLADE = &H400

Private Const OSV_LENGTH As Long = 148
Private Const OSVEX_LENGTH As Long = 156

Private Type OSVERSIONINFO
  OSVSize         As Long         'size, in bytes, of this data structure
  dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
  dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
  dwBuildNumber   As Long         'NT: build number of the OS
                                  'Win9x: build number of the OS in low-order word.
                                  '       High-order word contains major & minor ver nos.
  PlatformID      As Long         'Identifies the operating system platform.
  szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3"
                                  'Win9x: string providing arbitrary additional information
End Type

Private Type OSVERSIONINFOEX
  OSVSize            As Long
  dwVerMajor        As Long
  dwVerMinor         As Long
  dwBuildNumber      As Long
  PlatformID         As Long
  szCSDVersion       As String * 128
  wServicePackMajor  As Integer
  wServicePackMinor  As Integer
  wSuiteMask         As Integer
  wProductType       As Byte
  wReserved          As Byte
End Type

'defined As Any to support OSVERSIONINFO and OSVERSIONINFOEX
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" _
  (lpVersionInformation As Any) As Long
Public Function GetWinVer() As String

'returns a string representing the version,
'ie "95", "98", "NT4", "WinXP"
Dim osv As OSVERSIONINFO
Dim r As Long
Dim Pos As Integer
Dim sVer As String
Dim sBuild As String

osv.OSVSize = Len(osv)
If GetVersionEx(osv) = 1 Then
    'PlatformId contains a value representing the OS
    Select Case osv.PlatformID
    Case VER_PLATFORM_WIN32s: GetWinVer = "32s"
    Case VER_PLATFORM_WIN32_NT:
        'dwVerMajor = NT version.
        'dwVerMinor = minor version
        Select Case osv.dwVerMajor
        Case 3:
            Select Case osv.dwVerMinor
            Case 0:  GetWinVer = "WinNT3"
            Case 1:  GetWinVer = "WinNT3.1"
            Case 5:  GetWinVer = "WinNT3.5"
            Case 51: GetWinVer = "WinNT3.51"
            End Select
        Case 4: GetWinVer = "WinNT4"
        Case 5:
            Select Case osv.dwVerMinor
            Case 0:  GetWinVer = "Win2000"
            Case 1:  GetWinVer = "WinXP"
            Case 2:  GetWinVer = "Win2003"
            End Select
        Case 6: GetWinVer = "WinVista"
        End Select
    Case VER_PLATFORM_WIN32_WINDOWS:
        'dwVerMinor bit tells if its 95 or 98.
        Select Case osv.dwVerMinor
        Case 0:    GetWinVer = "Win95"
        Case 90:   GetWinVer = "WinME"
        Case Else: GetWinVer = "Win98"
        End Select
    End Select
End If
   
''Get the build
'win.BuildNo = (osv.dwBuildNumber And &HFFFF&)
'
''Any additional info. In Win9x, this can be
''"any arbitrary string" provided by the
''manufacturer. In NT, this is the service pack.
'pos = InStr(osv.szCSDVersion, Chr$(0))
'If pos Then
'    win.ServicePack = Left$(osv.szCSDVersion, pos - 1)
'End If


End Function

Public Function IsBackOfficeServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Microsoft BackOffice components are installed
  
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWinNT4Plus() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsBackOfficeServer = (osv.wSuiteMask And VER_SUITE_BACKOFFICE)
      End If
   End If

End Function

Public Function IsBladeServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Windows Server 2003 Web Edition is installed
  
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWin2003Server() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsBladeServer = (osv.wSuiteMask And VER_SUITE_BLADE)
      End If
   End If

End Function

Public Function IsDomainController() As Boolean
  
   Dim osv As OSVERSIONINFOEX
  'Returns True if the server is a domain
  'controller (Win 2000 or later), including
  'under active directory
   
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWin2000Server() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsDomainController = (osv.wProductType = VER_NT_SERVER) And _
                              (osv.wProductType = VER_NT_DOMAIN_CONTROLLER)
      End If
   End If

End Function

Public Function IsEnterpriseServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Windows NT 4.0 Enterprise Edition,
  'Windows 2000 Advanced Server, or Windows Server 2003
  'Enterprise Edition is installed.
   
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWinNT4Plus() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsEnterpriseServer = (osv.wProductType = VER_NT_SERVER) And _
                              (osv.wSuiteMask And VER_SUITE_ENTERPRISE)
      End If
   End If

End Function

Public Function IsWin2000AdvancedServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Windows 2000 Advanced Server
   
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWin2000Plus() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsWin2000AdvancedServer = ((osv.wProductType = VER_NT_SERVER) Or _
                                    (osv.wProductType = VER_NT_DOMAIN_CONTROLLER)) And _
                                    (osv.wSuiteMask And VER_SUITE_ENTERPRISE)
      End If
   End If

End Function

Public Function IsWin2000Server() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Windows 2000 Server
   
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWin2000() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsWin2000Server = (osv.wProductType = VER_NT_SERVER)
      End If
   End If

End Function

Public Function IsSmallBusinessServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Microsoft Small Business Server is installed
  
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWinNT4Plus() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsSmallBusinessServer = (osv.wSuiteMask And VER_SUITE_SMALLBUSINESS)
      End If
   End If

End Function

Public Function IsSmallBusinessRestrictedServer() As Boolean

   Dim osv As OSVERSIONINFOEX
  'Returns True if Microsoft Small Business Server
  'is installed with the restrictive client license
  'in force
  
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWinNT4Plus() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsSmallBusinessRestrictedServer = (osv.wSuiteMask And VER_SUITE_SMALLBUSINESS_RESTRICTED)
      End If
   End If

End Function

Public Function IsTerminalServer() As Boolean
  
   Dim osv As OSVERSIONINFOEX
  'Returns True if Terminal Services is installed
   
  'OSVERSIONINFOEX supported on NT4 or
  'later only, so a test is required
  'before using
   If IsWinNT4Plus() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsTerminalServer = (osv.wSuiteMask And VER_SUITE_TERMINAL)
      End If
   End If

End Function

Public Function IsWin95() As Boolean

  'returns True if running Win95
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWin95 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
                (osv.dwBuildNumber = 950)
   End If

End Function

Public Function IsWin95OSR2() As Boolean

  'returns True if running Win95
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWin95OSR2 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                    (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
                    (osv.dwBuildNumber = 1111)
   End If

End Function

Public Function IsWin98() As Boolean

  'returns True if running Win98
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWin98 = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                (osv.dwVerMajor = 4 And osv.dwVerMinor = 10) And _
                (osv.dwBuildNumber >= 2222)
   End If

End Function

Public Function IsWinME() As Boolean

  'returns True if running Windows ME
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWinME = (osv.PlatformID = VER_PLATFORM_WIN32_WINDOWS) And _
                (osv.dwVerMajor = 4 And osv.dwVerMinor = 90) And _
                (osv.dwBuildNumber >= 3000)
   End If

End Function

Public Function IsWinNT4() As Boolean

  'returns True if running WinNT4
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWinNT4 = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                 (osv.dwVerMajor = 4 And osv.dwVerMinor = 0) And _
                 (osv.dwBuildNumber >= 1381)
   End If

End Function

'===============================================================================
' Name: Function IsWinNT4Plus
' Input: None
' Output:
'   Boolean - True if WinNT4 or later
' Purpose: Returns True if running Windows NT4 or later
' Remarks: None
'===============================================================================
Public Function IsWinNT4Plus() As Boolean

  'returns True if running Windows NT4 or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWinNT4Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                     (osv.dwVerMajor >= 4)
   End If

End Function

Public Function IsWinNT4Server() As Boolean

  'returns True if running Windows NT4 Server
   Dim osv As OSVERSIONINFOEX
      
   If IsWinNT4() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsWinNT4Server = (osv.wProductType And VER_NT_SERVER)
      End If
   End If

End Function

Public Function IsWinNT4Workstation() As Boolean

  'returns True if running Windows NT4 Workstation
   Dim osv As OSVERSIONINFOEX
      
   If IsWinNT4() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsWinNT4Workstation = (osv.wProductType And VER_NT_WORKSTATION)
      End If
   End If

End Function

Public Function IsWin2000() As Boolean

  'returns True if running Win2000 (NT5)
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWin2000 = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                  (osv.dwVerMajor = 5 And osv.dwVerMinor = 0) And _
                  (osv.dwBuildNumber >= 2195)
   End If

End Function

Public Function IsWin2000Plus() As Boolean

  'returns True if running Windows 2000 or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWin2000Plus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                      (osv.dwVerMajor = 5 And osv.dwVerMinor >= 0)
   End If

End Function

Public Function IsWin2003Server() As Boolean

  'returns True if running Windows 2003 (.NET) Server
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWin2003Server = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                        (osv.dwVerMajor = 5 And osv.dwVerMinor = 2) And _
                        (osv.dwBuildNumber = 3790)
   End If

End Function

Public Function IsWin2000Workstation() As Boolean

  'returns True if running Windows NT4 Workstation
   Dim osv As OSVERSIONINFOEX
      
   If IsWin2000() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsWin2000Workstation = (osv.wProductType And VER_NT_WORKSTATION)
      End If
   End If

End Function

Public Function IsWinXP() As Boolean

  'returns True if running Windows XP (NT5.1)
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWinXP = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                (osv.dwVerMajor = 5 And osv.dwVerMinor = 1) And _
                (osv.dwBuildNumber >= 2600)
   End If

End Function
Public Function IsWinVista() As Boolean

  'returns True if running Windows Vista (NT6.0)
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWinVista = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                (osv.dwVerMajor = 6 And osv.dwVerMinor = 0)
   End If

End Function

Public Function IsWinXPSP2() As Boolean

  'returns True if running Windows XP SP2 (Service Pack 2)
   Dim osv As OSVERSIONINFOEX
      
   If IsWinXP() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsWinXPSP2 = InStr(osv.szCSDVersion, "Service Pack 2") > 0
      End If
   End If

End Function

Public Function IsWinXPPlus() As Boolean

  'returns True if running Windows XP or later
   Dim osv As OSVERSIONINFO

   osv.OSVSize = Len(osv)
   If GetVersionEx(osv) = 1 Then
      IsWinXPPlus = (osv.PlatformID = VER_PLATFORM_WIN32_NT) And _
                    (osv.dwVerMajor >= 5 And osv.dwVerMinor >= 1)
   End If

End Function

Public Function IsWinXPHomeEdition() As Boolean

  'returns True if running Windows XP Home Edition
   Dim osv As OSVERSIONINFOEX
      
   If IsWinXP() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsWinXPHomeEdition = ((osv.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL)
      End If
   End If

End Function

Public Function IsWinXPProEdition() As Boolean

  'returns True if running WinXP Pro
   Dim osv As OSVERSIONINFOEX
      
   If IsWinXP() Then
      osv.OSVSize = Len(osv)
      If GetVersionEx(osv) = 1 Then
         IsWinXPProEdition = Not ((osv.wSuiteMask And VER_SUITE_PERSONAL) = VER_SUITE_PERSONAL)
      End If
   End If

End Function



