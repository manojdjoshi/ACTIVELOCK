Option Strict Off
Option Explicit On 
Imports System.IO
Imports System.Security.Cryptography
Imports System.text
Module modActiveLock
	'*   ActiveLock
	'*   Copyright 1998-2002 Nelson Ferraz
	'*   Copyright 2005 The ActiveLock Software Group (ASG)
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
    ' This module handles contains common utility routines that can be shared
    ' between ActiveLock and the client application.
    ' This code has been overhauled to VB.NET
	'
	' @author activelock-users-admin
	' @version 3.0.0
	' @date 20050421
	'
	'* ///////////////////////////////////////////////////////////////////////
	'  /                        MODULE TO DO LIST                            /
	'  ///////////////////////////////////////////////////////////////////////
	'
	' @bug rsa_createkey() sometimes causes crash.  This is due to a bug in
    '      ALCrypto3NET.dll in which a bad keyset is sometimes generated
	'      (either caused by <code>rsa_generate()</code> or one of <code>rsa_private_key_blob()</code>
	'      and <code>rsa_public_key_blob()</code>--we're not sure which is the culprit yet.
	'      This causes the <code>rsa_createkey()</code> call encryption routines to crash.
	'      The work-around for the time being is to keep regenerating the keyset
	'      until eventually you'll get a valid keyset that no longer causes a crash.
	'      You only need to go through this keyset generation step once.
	'      Once you have a valid keyset, you should store it inside your app for later use.
	'
	
	'  ///////////////////////////////////////////////////////////////////////
	'  /                        MODULE CHANGE LOG                            /
	'  ///////////////////////////////////////////////////////////////////////
	' @history
	' <pre>
	'   07.07.03 - mcrute  - Updated the header comments for this file.
	'   07.30.03 - th2tran - New routines to do MD5 hashes of TypeLibs.
	'   08.02.03 - th2tran - wizzardme2000 found a gaping security hole with using md5_hash().
	'                        So now I&#39;m using CRC checksums and MapFileAndCheckSum() API call
	'                        instead.  This approach was suggested by Peter Young (vbclassicforever)
	'                        in the forum and mailing list a while back.
	'   08.02.03 - th2tran - VBdox&#39;d this module.
	'   04.17.04 - th2tran - Added FileExists() routine.
	'   07.11.04 - th2tran - New routines for handling GMT date-time.
	'   05.13.05 - ialkan  - Modified this module to comment out duplicate entries in modAlugen
    '   08.04.05 - ialkan  - Renamed ALCrypto3NET.dll as ALCrypto3NET.dll for V3 of Activelock
    '   08.09.05 - ialkan  - Codebase overhauled to VB.NET using the managed Activelock3NET.DLL
	' </pre>
	
	'  ///////////////////////////////////////////////////////////////////////
	'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
	'  ///////////////////////////////////////////////////////////////////////

	' RSA encrypts the data.
	' @param CryptType CryptType = 0 for public&#59; 1 for private
	' @param Data   Data to be encrypted
	' @param dLen   [in/out] Length of data, in bytes. This parameter will contain length of encrypted data when returned.
	' @param ptrKey Key to be used for encryption
    Public Declare Function rsa_encrypt Lib "ALCrypto3NET" (ByVal CryptType As Integer, ByVal data As String, ByRef dLen As Integer, ByRef ptrKey As RSAKey) As Integer

	' RSA decrypts the data.
	' @param CryptType CryptType = 0 for public&#59; 1 for private
	' @param Data   Data to be encrypted
	' @param dLen   [in/out] Length of data, in bytes. This parameter will contain length of encrypted data when returned.
	' @param ptrKey Key to be used for encryption
    Public Declare Function rsa_decrypt Lib "ALCrypto3NET" (ByVal CryptType As Integer, ByVal data As String, ByRef dLen As Integer, ByRef ptrKey As RSAKey) As Integer

	' Computes an MD5 hash from the data.
	' @param inData Data to be hashed
	' @param nDataLen   Length of inData
	' @param outData    [out] 32-byte Computed hash code
    Public Declare Function md5_hash Lib "ALCrypto3NET" (ByVal inData As String, ByVal nDataLen As Integer, ByVal outData As String) As Integer

    ' ActiveLock Encryption Key
	' !!!WARNING!!! It is highly recommended that you change this key for your version of ActiveLock before deploying your app.
	Public Const ENCRYPT_KEY As String = "AAAAgEPRFzhQEF7S91vt2K6kOcEdDDe5BfwNiEL30/+ozTFHc7cZctB8NIlS++ZR//D3AjSMqScjh7xUF/gwvUgGCjiExjj1DF/XWFWnPOCfF8UxYAizCLZ9fdqxb1FRpI5NoW0xxUmvxGjmxKwazIW4P4XVi/+i1Bvh2qQ6ri3whcsNAAAAQQCyWGsbJKO28H2QLYH+enb7ehzwBThqfAeke/Gv1Te95yIAWme71I9aCTTlLsmtIYSk9rNrp3sh9ItD2Re67SE7AAAAQQCAookH1nws1gS2XP9cZTPaZEmFLwuxlSVsLQ5RWmd9cuxpgw5y2gIskbL4c+4oBuj0IDwKtnMrZq7UfV9I5VfVAAAAQQCEnyAuO0ahXH3KhAboop9+tCmRzZInTrDYdMy23xf3PLCLd777dL/Y2Y+zmaH1VO03m6iOog7WLiN4dCL7m+Im" ' RSA Private Key
	
	Public Const MAGICNUMBER_YES As Integer = &HEFCDAB89
	Public Const MAGICNUMBER_NO As Integer = &H98BADCFE

	Private Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
	Private Const KEY_CONTAINER As String = "ActiveLock"
	Private Const PROV_RSA_FULL As Integer = 1
	
	Private fInit As Boolean ' flag to indicate that module initialization has been done
	
    Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Integer, ByRef source As Integer, ByVal length As Integer)
	Private Declare Function GetModuleFileName Lib "kernel32"  Alias "GetModuleFileNameA"(ByVal hModule As Integer, ByVal lpFileName As String, ByVal nSize As Integer) As Integer
	Private Declare Function MapFileAndCheckSum Lib "imagehlp"  Alias "MapFileAndCheckSumA"(ByVal FileName As String, ByRef HeaderSum As Integer, ByRef CheckSum As Integer) As Integer

	' Trims Null characters from the string.
	' @param startstr   String to be trimmed
	' @return Trimmed string
	Public Function TrimNulls(ByRef startstr As String) As String
		Dim pos As Short
		pos = InStr(startstr, Chr(0))
		If pos Then
			TrimNulls = Trim(Left(startstr, pos - 1))
		Else
			TrimNulls = Trim(startstr)
		End If
	End Function

	' Reads a binary file into the sData buffer.
	' Returns the number of bytes read.
	' @param sPath  Path to the file to be read
	' @param sData  Output parameter contains the data that has been read
	' @return number of bytes read, 0 if no file was read
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

        ReadFile = Len(sData)
        Exit Function
Hell:
        Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End Function

    ' [INTERNAL] Call-back routine used by ALCrypto3NET.dll during key generation process.
    ' @param param TBD
    ' @param action Action being performed
    ' @param phase  Current phase
    ' @param iprogress  Percent complete
    Public Sub CryptoProgressUpdate(ByVal param As Integer, ByVal action As Integer, ByVal phase As Integer, ByVal iprogress As Integer)
        System.Diagnostics.Debug.WriteLine("Progress Update received " & param & ", action: " & action & ", iprogress: " & iprogress)
    End Sub

    ' Computes an MD5 hash of the specified file.
    ' @param strPath    File path
    ' @return    MD5 Hash Value
    Public Function MD5HashFile(ByVal strPath As String) As String
        System.Diagnostics.Debug.WriteLine("Hashing file " & strPath)
        System.Diagnostics.Debug.WriteLine("File Date: " & FileDateTime(strPath))
        ' read and hash the content
        Dim sData As String
        Dim nFileLen As Integer
        nFileLen = ReadFile(strPath, sData)
        ' use the .NET's native MD5 functions instead of our own MD5 hashing routine
        ' and instead of ALCrypto's md5_hash() function.
        MD5HashFile = LCase(sData)    '<--- ReadFile procedure already computes the MD5.Hash
    End Function

    ' Converts a local date-time into UTC/GMT date-time
    ' @param dt   Date-Time input to be converted into UTC Date-Time
    ' @return     UTC Date-Time
    Public Function UTC(ByRef dt As Date) As Date
        '  Returns current UTC date-time.
        UTC = Date.UtcNow
    End Function
End Module