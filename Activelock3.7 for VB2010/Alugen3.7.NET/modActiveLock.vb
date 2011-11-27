Option Strict Off
Option Explicit On 

Imports System.IO
Imports System.Security.Cryptography
Imports System.text
Imports ActiveLock37Net

Module modActiveLock
    ' *   ActiveLock
    ' *   Copyright 1998-2002 Nelson Ferraz
    ' *   Copyright 2003-2008 The ActiveLock Software Group (ASG)
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
    ' *
    '===============================================================================
	' Name: modActiveLock
	' Purpose: This module contains common utility routines that can be shared
	' between ActiveLock and the client application.
	' Functions:
	' Properties:
	' Methods:
	' Started: 04.21.2005
    ' Modified: 03.25.2006
	'===============================================================================
	' @author activelock-admins
    ' @version 3.3.0
    ' @date 03.25.2006
	'
    Public Const STRRSAERROR As String = "Internal RSA Error."
    Public Const RETVAL_ON_ERROR As Integer = -999
    Public Const STRWRONGIPADDRESS As String = "Wrong IP Address."

    ' The following constants and declares are used to Get/Set Locale Date format
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
    Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String) As Boolean
    Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Short
    Const LOCALE_SSHORTDATE As Short = &H1FS
    Public regionalSymbol As String
    Public current_culture As System.Globalization.CultureInfo
    Public current_culture_name As String


    '===============================================================================
    ' Name: Function WinDir
    ' Input: None
    ' Output:
    '   String - Windows directory path
    ' Purpose: Gets the windows directory
    ' Remarks: None
    '===============================================================================
    Public Function WinDir() As String
        Dim WinSysPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.System)
        WinDir = WinSysPath.Substring(0, WinSysPath.LastIndexOf("\"))
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
        WinSysDir = System.Environment.GetFolderPath(Environment.SpecialFolder.System)
        ' or could use WinSysDir = System.Environment.SystemDirectory
    End Function
    '===============================================================================
    ' Name: Function FolderExists
    ' Input:
    '   ByVal sFolder As String -  Name of the folder in question
    ' Output:
    '   Boolean - Returns true if the Folder Exists
    ' Purpose: Checks if a Folder Exists
    ' Remarks: None
    '===============================================================================
    Public Function FolderExists(ByVal sFolder As String) As Boolean
        FolderExists = Directory.Exists(sFolder)
    End Function
    Public Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Short
        '===============================================================================
        '   MakeWord - Packs 2 8-bit integers into a 16-bit integer.
        '===============================================================================

        If (HiByte And &H80S) <> 0 Then
            MakeWord = ((HiByte * 256) + LoByte) Or &HFFFF0000
        Else
            MakeWord = (HiByte * 256) + LoByte
        End If

    End Function
    Public Function HiByte(ByVal w As Short) As Byte
        HiByte = (w And &HFF00) \ 256
    End Function
    Public Function LoByte(ByVal w As Short) As Byte
        LoByte = w And &HFFS
    End Function
    'Public Sub Get_locale() ' Retrieve the regional setting
    '    Dim Symbol As String
    '    Dim iRet1 As Integer
    '    Dim iRet2 As Integer
    '    Dim lpLCDataVar As String = String.Empty
    '    Dim Pos As Short
    '    Dim Locale As Integer
    '    Locale = GetUserDefaultLCID()
    '    iRet1 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, lpLCDataVar, 0)
    '    Symbol = New String(Chr(0), iRet1)
    '    iRet2 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol, iRet1)
    '    Pos = InStr(Symbol, Chr(0))
    '    If Pos > 0 Then
    '        Symbol = Left(Symbol, Pos - 1)
    '        If Symbol <> "yyyy/MM/dd" Then regionalSymbol = Symbol
    '    End If
    'End Sub
    'Public Sub Set_locale(Optional ByVal localSymbol As String = "") 'Change the regional setting
    '    Dim Symbol As String
    '    Dim iRet As Integer
    '    Dim Locale As Integer
    '    Locale = GetUserDefaultLCID() 'Get user Locale ID
    '    If localSymbol = "" Then
    '        Symbol = "yyyy/MM/dd" 'New character for the locale
    '    Else
    '        Symbol = localSymbol
    '    End If

    '    iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol)
    'End Sub
    Public Sub Change_Culture(ByVal c As String) 'Change the regional setting
        If c = "en-US" Then
            'current_culture = System.Threading.Thread.CurrentThread.CurrentCulture
            'Dim us_culture As New System.Globalization.CultureInfo("en-US")
            'System.Threading.Thread.CurrentThread.CurrentCulture = us_culture
            If System.Threading.Thread.CurrentThread.CurrentCulture.Name <> "en-US" Then
                current_culture_name = System.Threading.Thread.CurrentThread.CurrentCulture.Name
                My.Application.ChangeCulture("en-US")
            End If
        Else
            'System.Threading.Thread.CurrentThread.CurrentCulture = current_culture
            If current_culture_name <> "en-US" And current_culture_name <> Nothing Then
                My.Application.ChangeCulture(current_culture_name)
            End If
        End If
    End Sub

End Module