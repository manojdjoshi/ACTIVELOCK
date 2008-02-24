'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2008 The ActiveLock Software Group (ASG)
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
Option Strict On
Option Explicit On 
Imports System.IO

Module modINIFile

  'Declare API calls needed to read/write to INI files
  Private Declare Function KRN32_GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Integer) As Integer
  Private Declare Function KRN32_GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer) As Integer
  Private Declare Function KRN32_WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Integer
  Private Declare Function KRN32_GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Integer, ByVal lpFileName As String) As Integer
  Private Declare Function KRN32_GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Integer, ByVal lpFileName As String) As Integer
  Private Declare Function KRN32_WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Integer


  Public Function ProfileString32(ByRef sININame As String, ByRef sSection As String, ByRef sKeyword As String, ByRef vsDefault As String) As String
    '
    '   This routine will get a string from an INI file. The user must pass the
    '   following information:
    '
    '   Input Values
    '       sININame   - The name of the INI file (with path if desired)
    '       sSection   - The section head to search for
    '       sKeyword   - The keyword within the section to return
    '       vsDefault   - The default value to return if nothing is found
    '
    '   Output Value
    '       A string with the value found (max length = 512)
    '
    Dim sReturnValue As String
    Dim nReturnSize As Integer
    Dim nValidSize As Integer

    Try
      If IsNothing(sININame) Or IsNothing(sSection) Or IsNothing(sKeyword) Then
        ProfileString32 = vsDefault
        Exit Function
      End If

      sReturnValue = Space(512)
      nReturnSize = sReturnValue.Length

      If sININame.ToUpper.Substring(sININame.Length - 7) = "WIN.INI" Then
        nValidSize = KRN32_GetProfileString(sSection, sKeyword, vsDefault, sReturnValue, nReturnSize)
      Else
        nValidSize = KRN32_GetPrivateProfileString(sSection, sKeyword, vsDefault, sReturnValue, nReturnSize, sININame)
      End If

      ProfileString32 = sReturnValue.Substring(0, nValidSize)
    Catch ex As Exception
      ProfileString32 = vsDefault
    End Try
  End Function

  Public Function SetProfileString32(ByRef sININame As String, ByRef sSection As String, ByRef sKeyword As String, ByRef vsEntry As String) As Integer
    '
    '   This routine will write a string to an INI file. The user must pass the
    '   following information:
    '
    '   Input Values
    '       sININame   - The name of the INI file (with path if desired)
    '       sSection   - The section head to search for
    '       sKeyword   - The keyword within the section to return
    '       vsEntry     - The value to write for the keyword
    '
    '   Output Value
    '       non-zero if successful, zero if failure.
    '
    Try
      If IsNothing(sININame) Or IsNothing(sSection) Or IsNothing(sKeyword) Then
        SetProfileString32 = 0
        Exit Function
      End If

      If sININame.ToUpper.Substring(sININame.Length - 7) = "WIN.INI" Then
        SetProfileString32 = KRN32_WriteProfileString(sSection, sKeyword, vsEntry)
      Else
        SetProfileString32 = KRN32_WritePrivateProfileString(sSection, sKeyword, vsEntry, sININame)
      End If
    Catch ex As Exception
      SetProfileString32 = 0
    End Try
  End Function
End Module
