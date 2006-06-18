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
    x As Long
    y As Long
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
'TestApp
Public Const PUB_KEY$ = "2CB.2CB.2CB.2CB.2D6.231.35A.53E.42B.2E1.21B.533.441.226.2F7.2CB.2CB.2CB.2CB.2D6.32E.37B.2CB.2CB.2CB.323.2D6.268.205.2D6.226.339.3BD.4C5.42B.483.226.3BD.391.30D.39C.386.370.441.46D.4AF.34F.4C5.441.53E.457.3C8.4D0.44C.268.4BA.512.210.3D3.23C.4E6.21B.4F1.32E.21B.51D.3B2.231.512.318.226.21B.4DB.23C.4E6.39C.4D0.2F7.3D3.507.2D6.483.2EC.23C.318.302.365.4D0.499.436.35A.2D6.391.386.44C.4D0.2D6.318.32E.30D.3BD.457.441.25D.48E.3A7.483.268.323.391.3B2.210.4D0.34F.252.483.226.339.53E.4BA.48E.478.2E1.4AF.4F1.247.2E1.2F7.4FC.3D3.318.386.533.436.436.483.3D3.512.386.3C8.4A4.457.30D.53E.302.4F1.2CB.2CB.370.268.21B.25D.370.344.35A.231.32E.3D3.4C5.231.3BD.499.2F7.4E6.39C.226.4C5.462.386.247.386.2E1.499.462.478.4AF.528.210.252.210.2D6.39C.268.51D.42B.370.4C5.4DB.4BA.4BA.231.2CB.2D6.25D.4F1.3DE.210.37B.29F.29F"
'Test_App
'Public Const PUB_KEY$ = "2CB.2CB.2CB.2CB.2D6.231.35A.53E.42B.2E1.21B.533.441.226.2F7.2CB.2CB.2CB.2CB.2D6.32E.37B.2CB.2CB.2CB.323.2E1.48E.4A4.3DE.252.210.39C.210.37B.51D.344.441.386.252.2EC.1D9.51D.323.23C.25D.3A7.2E1.210.35A.1D9.339.39C.2E1.25D.3A7.339.44C.507.32E.231.441.2F7.499.3DE.2E1.318.46D.231.4A4.2EC.365.533.2CB.3D3.37B.3C8.4AF.53E.4A4.51D.25D.23C.4A4.4F1.48E.32E.42B.457.3A7.318.365.4F1.21B.2F7.3B2.339.436.318.34F.4AF.512.4A4.302.23C.507.365.42B.302.323.1D9.4FC.252.457.35A.2D6.4DB.3BD.3D3.4C5.268.4AF.210.441.3A7.528.386.318.3BD.386.231.210.4C5.23C.457.4C5.226.3DE.2D6.273.3DE.302.533.48E.48E.370.4A4.478.2F7.4E6.25D.231.3A7.4D0.4AF.51D.226.226.478.205.25D.48E.32E.533.35A.323.21B.462.4AF.4FC.46D.2D6.34F.37B.3B2.4AF.323.3B2.457.365.528.1D9.51D.23C.46D.25D.441.339.318.4A4.4AF.386.457.35A.4F1.507.21B.37B.29F.29F"

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
    Value = 690000 + 127       ' compute it so that it can't be easily spotted via a Hex Editor
End Property

' Callback function for rsa_generate()
'
Public Sub ProgressUpdate(ByVal param As Long, ByVal action As Long, ByVal phase As Long, ByVal iprogress As Long)
    frmMain.UpdateStatus "Progress Update received " & param & ", action: " & action & ", iprogress: " & iprogress
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
    GetTypeLibPathFromObject = WinSysDir() & "\activelock" & CStr(App.Major) & "." & CStr(App.Minor) & ".dll" 'strDllPath
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
