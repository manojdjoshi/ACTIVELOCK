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
Option Strict Off
Option Explicit On 

Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Namespace ASPNETAlugen3

    Public Module modActiveLock

        Public Const STRRSAERROR As String = "Internal RSA Error."
        Public Const RETVAL_ON_ERROR As Integer = -999
        Public Const STRWRONGIPADDRESS As String = "Wrong IP Address."

        ' The following constants and declares are used to Get/Set Locale Date format
        Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
        Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String) As Boolean
        Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Short
        Const LOCALE_SSHORTDATE As Short = &H1FS
        Public regionalSymbol As String

        Public Function RSASign(ByVal strPub As String, ByVal strPriv As String, ByVal strdata As String) As String
            ' Performs RSA signing of <code>strData</code> using the specified key.
            ' @param strPub     RSA Public key blob
            ' @param strPriv    RSA Private key blob
            ' @param strData    Data to be signed
            ' @return           Signature string.
            ' 05.13.05    - alkan  - Removed the modActiveLock references
            '
            Dim Key As RSAKey
            ReDim Key.data(32)
            ' create the key from the key blobs
            rsa_createkey(strPub, strPub.Length, strPriv, strPriv.Length, Key)
            ' sign the data using the created key
            Dim sLen As Integer
            rsa_sign(Key, strdata, strdata.Length, vbNullString, sLen)
            Dim strSig As New String(Chr(0), sLen)
            rsa_sign(Key, strdata, strdata.Length, strSig, sLen)
            ' throw away the key
            rsa_freekey(Key)
            RSASign = strSig
        End Function

        Public Function RSAVerify(ByVal strPub As String, ByVal strdata As String, ByVal strSig As String) As Integer
            ' Verifies an RSA signature.
            ' @param strPub     Public key blob
            ' @param strData    Data to be signed
            ' @param strSig     Private key blob
            ' @return           Zero if verification is successful; Non-zero otherwise.
            '
            Dim Key As RSAKey = Nothing
            Dim rc As Integer
            ' create the key from the public key blob
            rsa_createkey(strPub, strPub.Length, vbNullString, 0, Key)
            ' validate the key
            rc = rsa_verifysig(Key, strSig, strSig.Length, strdata, strdata.Length)
            ' de-allocate memory used by the key
            rsa_freekey(Key)
            RSAVerify = rc
        End Function
        Public Sub Get_locale() ' Retrieve the regional setting
            Dim Symbol As String
            Dim iRet1 As Integer
            Dim iRet2 As Integer
            Dim lpLCDataVar As String = String.Empty
            Dim Pos As Short
            Dim Locale As Integer
            Locale = GetUserDefaultLCID()
            iRet1 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, lpLCDataVar, 0)
            Symbol = New String(Chr(0), iRet1)
            iRet2 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol, iRet1)
            Pos = InStr(Symbol, Chr(0))
            If Pos > 0 Then
                Symbol = Left(Symbol, Pos - 1)
                If Symbol <> "yyyy/MM/dd" Then regionalSymbol = Symbol
            End If
        End Sub
        Public Sub Set_locale(Optional ByVal localSymbol As String = "") 'Change the regional setting
            Dim Symbol As String
            Dim iRet As Integer
            Dim Locale As Integer
            Locale = GetUserDefaultLCID() 'Get user Locale ID
            If localSymbol = "" Then
                Symbol = "yyyy/MM/dd" 'New character for the locale
            Else
                Symbol = localSymbol
            End If

            iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol)
        End Sub


    End Module
End Namespace
