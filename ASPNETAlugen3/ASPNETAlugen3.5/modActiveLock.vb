'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
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
Option Explicit On 
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography


Module modActiveLock

  Public Const STRKEYSTOREINVALID As String = "A license property contains an invalid value."
  Public Const STRLICENSEEXPIRED As String = "License expired."
  Public Const STRLICENSEINVALID As String = "License invalid."
  Public Const STRNOLICENSE As String = "No valid license."
  Public Const STRLICENSETAMPERED As String = "License may have been tampered."
  Public Const STRNOTINITIALIZED As String = "ActiveLock has not been initialized."
  Public Const STRNOTIMPLEMENTED As String = "Not implemented."
  Public Const STRCLOCKCHANGED As String = STRLICENSEINVALID & " System clock has been tampered."
  Public Const STRINVALIDTRIALDAYS As String = "Zero Free Trial days allowed."
  Public Const STRINVALIDTRIALRUNS As String = "Zero Free Trial runs allowed."
  Public Const STRFILETAMPERED As String = "Alcrypto3.dll has been tampered."
  Public Const STRKEYSTOREUNINITIALIZED As String = "Key Store Provider hasn't been initialized yet."
  Public Const STRNOSOFTWARECODE As String = "Software code has not been set."
  Public Const STRNOSOFTWARENAME As String = "Software Name has not been set."
  Public Const STRNOSOFTWAREVERSION As String = "Software Version has not been set."
  Public Const STRUSERNAMETOOLONG As String = "User Name > 2000 characters."
  Public Const STRRSAERROR As String = "Internal RSA Error."
  Public Const RETVAL_ON_ERROR As Integer = -999


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
    Dim Key As RSAKey
    Dim rc As Integer
    ' create the key from the public key blob
    rsa_createkey(strPub, strPub.Length, vbNullString, 0, Key)
    ' validate the key
    rc = rsa_verifysig(Key, strSig, strSig.Length, strdata, strdata.Length)
    ' de-allocate memory used by the key
    rsa_freekey(Key)
    RSAVerify = rc
  End Function

  Public Function ReadFile(ByVal sPath As String, ByRef sData As String) As Integer

    Dim c As New CRC32
    Dim crc As Integer = 0

    ' CRC32 Hash:
    Dim f As FileStream = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
    crc = c.GetCrc32(f)
    f.Close()

    ' File size:
    'f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
    'txtSize.Text = String.Format("{0}", f.Length)
    'f.Close()
    'txtCrc32.Text = String.Format("{0:X8}", crc)
    'txtTime.Text = String.Format("{0}", h.ElapsedTime)

    ' Run MD5 Hash
    f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
    Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider
    md5.ComputeHash(f)
    f.Close()

    Dim hash As Byte() = md5.Hash
    Dim buff As StringBuilder = New StringBuilder
    Dim hashByte As Byte
    For Each hashByte In hash
      buff.Append(String.Format("{0:X1}", hashByte))
    Next
    sData = buff.ToString() 'MD5 String

    ' Run SHA-1 Hash
    'f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
    'Dim sha1 As SHA1CryptoServiceProvider = New SHA1CryptoServiceProvider
    'sha1.ComputeHash(f)
    'f.Close()
    'hash = SHA1.Hash
    'buff = New StringBuilder
    'For Each hashByte In hash
    '    buff.Append(String.Format("{0:X1}", hashByte))
    'Next
    'txtSHA1.Text = buff.ToString()

    ReadFile = sData.Length
    Exit Function
Hell:
    Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
  End Function
End Module