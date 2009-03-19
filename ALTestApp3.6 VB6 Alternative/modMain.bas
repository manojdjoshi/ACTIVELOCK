Attribute VB_Name = "modMain"
'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003 The ActiveLock Software Group (ASG)
'*   All material is the property of the contributing authors.
'*
'*   Redistribution and use in source and binary forms, with or without
'*   modification, are permitted provided that the following conditions are
'*   met:
'*
'*     [o] Redistributions of source code must retain the above copyright
'*         notice, this list of conditions and the following disclaimer.
'*
'*     [o] Redistributions in binary form must reproduce the above
'*         copyright notice, this list of conditions and the following
'*         disclaimer in the documentation and/or other materials provided
'*         with the distribution.
'*
'*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
'*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
'*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
'*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'*
'*

''
' This module handles contains common utility routines that can be shared
' between ActiveLock and the client application.
'
' @author th2tran
' @version 2.0.2
' @date 20030715
'
'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
'
'

'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' @history
' <pre>
'   07.15.03 - th2tran - Created
'   08.15.03 - th2tran - Value() - Following vbclassicforever's suggestion:
'                        Compute the expected CRC instead leaving it as a plain
'                        value to make it more difficult to spot in a hex editor.
'   09.21.03 - th2tran - Dumped PRIV_KEY. PRIV_KEY should only be accessible to ALUGEN.
'   10.13.03 - th2tran - Copied a small number of functions from modActiveLock.bas into here
'                        so that our test app doesn't need to depend on modActiveLock.
'   11.02.03 - th2tran - Added simple encrypt/decrypt routines to be used by frmMain
'   11.08.03 - th2tran - Previously, GetTypeLibPathFromObject() used to retrieve the ActiveLock2
'                        TypeLib path using the TLI library (tlbinfo.dll).  This was proven to be unsecure because
'                        tlbinfo32.dll is a non-system DLL and therefore can be easily replace with
'                        a dummy DLL, thereby thwarting our checksum scheme.
'                        Thanks to Peter Young (pyoung@vbadvance.com) for pointing this out.
'                        I have now replaced the TLI implementation with a simpler registry lookup
'                        implementation.
' </pre>

'* ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit
Option Private Module

'**************************************************************************************************
' Win32 Structs & Enums
'**************************************************************************************************
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type SIZE
    X As Long
    Y As Long
End Type

'Public Type POINTAPI
'    x As Long
'    y As Long
'End Type

Public Type NOTIFYICONDATAA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

'Public Type NOTIFYICONDATAW
'    cbSize As Long
'    hwnd As Long
'    uId As Long
'    uFlags As Long
'    uCallbackMessage As Long
'    hIcon As Long
'    szTip(0 To 127) As Byte
'End Type

Public Type APPBARDATA
    cbSize As Long
    hWnd As Long
    uCallbackMessage As Long
    uEdge As SHAPPBAR_EDGES
    rc As RECT
    lParam As Long
End Type

Public Type APPBAR_DETAILS
    at_clkHwnd As Long
    at_clkRECT As RECT
    at_CurEdge As SHAPPBAR_EDGES
    at_IconHeight As Long
    at_IsAutoHide As Boolean
    at_IsHidden As Boolean
    at_LastEdge As SHAPPBAR_EDGES
    at_NumRows As Long
    at_saHwnd As Long
    at_saHeight As Long
    at_saLastHeight As Long
    at_saWidth As Long
    at_saLastWidth As Long
    at_saRECT As RECT
    at_saRECTPRE As RECT
    at_saRECTPOST As RECT
    at_tbHwnd As Long
End Type

Public Enum SHAPPBAR_MESSAGES
    ABM_NEW = &H0
    ABM_REMOVE = &H1
    ABM_QUERYPOS = &H2
    ABM_SETPOS = &H3
    ABM_GETSTATE = &H4
    ABM_GETTASKBARPOS = &H5
    ABM_ACTIVATE = &H6
    ABM_GETAUTOHIDEBAR = &H7
    ABM_SETAUTOHIDEBAR = &H8
    ABM_WINDOWPOSCHANGED = &H9
End Enum

Public Enum SHAPPBAR_NOTIFICATIONS
    ABN_STATECHANGE = &H0
    ABN_POSCHANGED = &H1
    ABN_FULLSCREENAPP = &H2
    ABN_WINDOWARRANGE = &H3
End Enum

Public Enum SHAPPBAR_STATES
    ABS_AUTOHIDE = &H1
    ABS_ALWAYSONTOP = &H2
End Enum

Public Enum SHAPPBAR_EDGES
    ABE_LEFT = 0
    ABE_TOP = 1
    ABE_RIGHT = 2
    ABE_BOTTOM = 3
End Enum


'**************************************************************************************************
' atViewPort Control Property Types/Enums:
'**************************************************************************************************
' Control placement
Public Type AT_CTLPOSITION
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type
' Control placement in status area
Public Type AT_CTLSAPOSITION
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type
' Is the world round or flat
Public Enum AT_CTLTICKERAPPEARANCE
    [Flat]
    [3D]
End Enum
' Border Or No Border
Public Enum AT_CTLBORDER
    [None]
    [FixedSingle]
End Enum
' Voice gender if speech TickerEnabled
Public Enum AT_CTLGENDER
    [Male]
    [Female]
End Enum
' Where is control sited
Public Enum AT_CTLHOST
    [HostContainer]
    [StatusArea]
End Enum
' Scroll speed
Public Enum AT_CTLSPEED
    [Slowest]
    [Slow]
    [Normal]
    [Fast]
    [Fastest]
End Enum
' ShowTicker Constants
Public Enum AT_CTLSTATE
    AT_ADDICONS = 0
    AT_REMOVEICONS = 1
    AT_SHOW = 2
    AT_HIDE = 3
    AT_RESIZE = 4
End Enum

'**************************************************************************************************
' Balloontip Structures
'**************************************************************************************************
Public Type TOOLINFO
    tiSize As Long
    tiFlags As Long
    tiHwnd As Long
    tiID As Long
    tiRect As RECT
    tiInst As Long
    tiSzText As String
    #If WIN32_IE >= &H300 Then
        tiParam As Long
    #End If
End Type

Public Enum INFOTITLE
    NoIcon
    InfoIcon
    WarningIcon
    ErrorIcon
End Enum

Public Type INITCOMMONCONTROLEXSTRUCT
    iccSize As Long
    iccICC As Long
End Type

Public Type OLECOLOR
    RedOrSys As Byte
    Green As Byte
    Blue As Byte
    Type As Byte
End Type

Public Enum DELAYTIME
    Automatic = &H0
    Reshow = &H1
    AutoPop = &H2
    Initial = &H3
End Enum

' Windows API Declares
Private Declare Function MapFileAndCheckSum Lib "imagehlp" Alias "MapFileAndCheckSumA" (ByVal FileName As String, HeaderSum As Long, CheckSum As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' Application Encryption keys:
' !!!WARNING!!!
' It is alright to use these same keys for testing your application.  But it is highly recommended
' that you generate your own set of keys to use before deploying your app.
'Enc("AAAAB3NzaC1yc2EAAAABJQAAAIB8/B2KWoai2WSGTRPcgmMoczeXpd8nv0Y4r1sJ1wV3vH21q4rTpEYuBiD4HFOpkbNBSRdpBHJGWec7jUi8ISV0pM6i2KznjhCms5CEtYHRybbiYvRXleGzFsAAP817PLN3JYo3WkErT2ofR5RCkfhmx060BT8waPoqnn3AB7sZ0Q==")
'TestApp, Version 2.0
'Public PUB_KEY As String
'You need to encrypt you software code
'SATYA
Public Const PUB_KEY = "386.391.2CB.21B.210.226.23C.2D6.46D.323.2CB.2CB.2CB.2CB.499.2CB.2CB.2D6.391.3A7.210.2F7.528.2CB.2CB.37B.2CB.2CB.2CB.2F7.2CB.2CB.37B.2EC.39C.512.2EC.2F7.507.48E.4BA.318.48E.35A.21B.302.478.30D.37B.23C.205.386.3DE.2EC.3DE.3D3.512.2F7.2E1.32E.3C8.252.370.53E.247.3BD.507.499.4AF.370.4FC.205.2D6.210.25D.391.365.35A.344.4A4.32E.32E.48E.4D0.35A.507.3DE.25D.3DE.4AF.441.4D0.35A.35A.37B.4AF.512.3A7.51D.339.391.3BD.4DB.2CB.365.4D0.528.32E.247.3DE.23C.3DE.2EC.39C.268.462.3C8.3BD.30D.4BA.436.3A7.4F1.2F7.4D0.323.51D.247.365.2D6.4D0.53E.3B2.34F.4FC.247.23C.268.441.4FC.386.4C5.4E6.32E.391.2F7.478.365.344.268.499.44C.44C.3A7.4F1.3B2.273.302.323.4A4.533.30D.4F1.46D.528.391.39C.2CB.4AF.205.3A7.2F7.273.231.507.478.273.35A.4A4.4FC.25D.53E.318.1D9.44C.483.34F.441.2E1.1D9.3DE.30D.2EC.2EC.4D0.46D.44C.273.2EC.483.3A7.2F7.3B2.21B.2CB.29F.29F"

' Verifies the checksum of the typelib containing the specified object.
' Returns the checksum.
'
Public Function VerifyActiveLockdll() As String
    Dim crc As Long
    crc = CRCCheckSumTypeLib()
    'Debug.Print "Hash: " & crc
    If crc <> Value Then
        ' Encrypted version of "Activelock DLL has been corrupted." If you were running a real application, it should terminate at this point.
        MsgBox Dec("2CB.441.4FC.483.512.457.4A4.4C5.441.499.160.2EC.344.344.160.478.42B.4F1.160.436.457.457.4BA.160.441.4C5.4E6.4E6.507.4D0.4FC.457.44C.1FA"), vbExclamation
        End
    End If
    VerifyActiveLockdll = CStr(crc)
End Function

''
' Simple encrypt of a string
Public Function Enc(strdata As String) As String
    Dim i&, n&
    Dim sResult$
    n = Len(strdata)
    Dim l As Long
    For i = 1 To n
        l = Asc(Mid$(strdata, i, 1)) * 11
        If sResult = "" Then
            sResult = Hex(l)
        Else
            sResult = sResult & "." & Hex(l)
        End If
    Next i
    Enc = sResult
End Function

Public Function Dec(strdata As String) As String
    Dim arr() As String
    arr = Split(strdata, ".")
    Dim sRes As String
    Dim i&
    For i = LBound(arr) To UBound(arr)
        sRes = sRes & Chr$(CLng("&h" & arr(i)) / 11)
    Next
    Dec = sRes
End Function


''
' Returns the expected CRC value of ActiveLock3.dll
'
Private Property Get Value() As Long
    Value = 710000 + 118      ' compute it so that it can't be easily spotted via a Hex Editor
End Property

' Callback function for rsa_generate()
'
Public Sub ProgressUpdate(ByVal param As Long, ByVal action As Long, ByVal phase As Long, ByVal iprogress As Long)
    'frmMain.UpdateStatus "Progress Update received " & param & ", action: " & action & ", iprogress: " & iprogress
End Sub


Public Function GetTypeLibPathFromObject() As String
    Dim strDllPath As String
    ' Read DLL Path using a Registry Lookup:
    '  Second parm = HKEY_CLASSES_ROOT\CLSID\{F749C3AE-19CC-4209-AE71-1A24D3F710F6}\InprocServer32
    '                   {F749C3AE-19CC-4209-AE71-1A24D3F710F6}= ClsID for ActiveLock3.Globals
'    strDllPath = modRegistryAPIs.ReadRegVal(HKEY_CLASSES_ROOT, _
'                                            Dec("2E1.344.391.323.2EC.3F4.549.302.25D.23C.273.2E1.231.2CB.2F7.1EF.21B.273.2E1.2E1.1EF.23C.226.210.273.1EF.2CB.2F7.25D.21B.1EF.21B.2CB.226.23C.2EC.231.302.25D.21B.210.302.252.55F.3F4.323.4BA.4D0.4E6.4C5.441.391.457.4E6.512.457.4E6.231.226"), _
'                                            "", _
'                                            Dec("42B.441.4FC.483.512.457.4A4.4C5.441.499.231.1FA.226.1FA.44C.4A4.4A4"))
    'Debug.Print "DLL Path: " + strDllPath
'    GetTypeLibPathFromObject = WinSysDir() & "\activelock" & CStr(App.Major) & "." & CStr(App.Minor) & ".dll" 'strDllPath
'SATYA: Temporarily lets use 3.5 as hard coded.
GetTypeLibPathFromObject = WinSysDir() & "\activelock" & CStr(3) & "." & CStr(6) & ".dll" 'strDllPath
End Function

'===============================================================================
' Name: Function WinSysDir
' Input: None
' Output:
'   String - Windows system directory path
' Purpose: Gets the Windows system directory
' Remarks: None
'===============================================================================
Public Function WinSysDir() As String
    Const FIX_LENGTH% = 4096
    Dim Length As Integer
    Dim Buffer As String * FIX_LENGTH

    Length = GetSystemDirectory(Buffer, FIX_LENGTH - 1)
    WinSysDir = Left$(Buffer, Length)
End Function

''
' Performs CRC checksum on the type library containing the object.
' @param obj    COM object used to determine the file path to the type library
'
Public Function CRCCheckSumTypeLib() As Long
    Dim strDllPath As String
    strDllPath = GetTypeLibPathFromObject()
    Dim HeaderSum As Long, RealSum As Long
    MapFileAndCheckSum strDllPath, HeaderSum, RealSum
    CRCCheckSumTypeLib = RealSum
End Function
