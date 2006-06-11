Option Strict Off
Option Explicit On 
Imports System.IO
Imports System.Security.Cryptography
Imports System.text
Module modActiveLock
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
	'*
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
	'* ///////////////////////////////////////////////////////////////////////
	'  /                        MODULE TO DO LIST                            /
	'  ///////////////////////////////////////////////////////////////////////
	'
	' @bug rsa_createkey() sometimes causes crash.  This is due to a bug in
	'      ALCrypto3.dll in which a bad keyset is sometimes generated
	'      (either caused by <code>rsa_generate()</code> or one of <code>rsa_private_key_blob()</code>
	'      and <code>rsa_public_key_blob()</code>--we're not sure which is the culprit yet.
	'      This causes the <code>rsa_createkey()</code> call encryption routines to crash.
	'      The work-around for the time being is to keep regenerating the keyset
	'      until eventually you'll get a valid keyset that no longer causes a crash.
	'      You only need to go through this keyset generation step once.
	'      Once you have a valid keyset, you should store it inside your app for later use.
	'

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
    Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Integer, ByVal lpFileName As String, ByVal nSize As Integer) As Integer
    Private Declare Function MapFileAndCheckSum Lib "imagehlp" Alias "MapFileAndCheckSumA" (ByVal FileName As String, ByRef HeaderSum As Integer, ByRef CheckSum As Integer) As Integer

    Structure SYSTEMTIME
        Dim wYear As Short
        Dim wMonth As Short
        Dim wDayOfWeek As Short
        Dim wDay As Short
        Dim wHour As Short
        Dim wMinute As Short
        Dim wSecond As Short
        Dim wMilliseconds As Short
    End Structure

    Private Structure TIME_ZONE_INFORMATION
        Dim bias As Integer ' current offset to GMT
        <VBFixedArray(64)> Dim StandardName() As Byte ' unicode string
        Dim StandardDate As SYSTEMTIME
        Dim StandardBias As Integer
        <VBFixedArray(64)> Dim DaylightName() As Byte
        Dim DaylightDate As SYSTEMTIME
        Dim DaylightBias As Integer
        Public Sub Initialize()
            ReDim StandardName(64)
            ReDim DaylightName(64)
        End Sub
    End Structure

    Public Enum TimeZoneReturn
        TimeZoneCode = 0
        TimeZoneName = 1
        UTC_BaseOffset = 2
        UTC_Offset = 3
        DST_Active = 4
        DST_Offset = 5
    End Enum

    ' ----------------- For Time Zone Retrieval ------------------
    Private Const TIME_ZONE_ID_UNKNOWN As Short = 0
    Private Const TIME_ZONE_ID_STANDARD As Short = 1
    Private Const TIME_ZONE_ID_INVALID As Integer = &HFFFFFFFF
    Private Const TIME_ZONE_ID_DAYLIGHT As Short = 2

    Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Integer

    ' To Report API errors:
    Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Short = &H100S
    Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Short = &H2000S
    Private Const FORMAT_MESSAGE_FROM_HMODULE As Short = &H800S
    Private Const FORMAT_MESSAGE_FROM_STRING As Short = &H400S
    Private Const FORMAT_MESSAGE_FROM_SYSTEM As Short = &H1000S
    Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Short = &H200S
    Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Short = &HFFS

    Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, ByRef lpSource As Long, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As String, ByVal nSize As Integer, ByRef Arguments As Integer) As Integer
    Public Declare Function GeneralWinDirApi Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

    Public Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    '===============================================================================
    ' Name: Function TrimNulls
    ' Input:
    '   ByRef startstr As String - String to be trimmed
    ' Output:
    '   String - Trimmed string
    ' Purpose: Trims Null characters from the string.
    ' Remarks: None
    '===============================================================================
    Public Function TrimNulls(ByRef startstr As String) As String
        Dim pos As Short
        pos = InStr(startstr, Chr(0))
        If pos Then
            TrimNulls = Trim(Left(startstr, pos - 1))
        Else
            TrimNulls = Trim(startstr)
        End If
    End Function
    '===============================================================================
    ' Name: Function ReadFile
    ' Input:
    '   ByVal sPath As String - Path to the file to be read
    '   ByRef sData As String - Output parameter contains the data that has been read
    ' Output:
    '   Long - Number of bytes read, 0 if no file was read
    ' Purpose: Reads a binary file into the sData buffer. Returns the number of bytes read.
    ' Remarks: None
    '===============================================================================
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
    '===============================================================================
    ' Name: Sub CryptoProgressUpdate
    ' Input:
    '   ByVal param As Long - TBD
    '   ByVal action As Long - Action being performed
    '   ByVal phase As Long - Current phase
    '   ByVal iprogress As Long - Percent complete
    ' Output: None
    ' Purpose: [INTERNAL] Call-back routine used by ALCrypto3.dll during key generation process.
    ' Remarks: None
    '===============================================================================
    Public Sub CryptoProgressUpdate(ByVal param As Integer, ByVal action As Integer, ByVal phase As Integer, ByVal iprogress As Integer)
        System.Diagnostics.Debug.WriteLine("Progress Update received " & param & ", action: " & action & ", iprogress: " & iprogress)
    End Sub
    '===============================================================================
    ' Name: Sub EndSub
    ' Input: None
    ' Output: None
    ' Purpose: This is a dummy sub. Used to circumvent the End statement restriction in COM DLLs.
    ' Remarks: None
    '===============================================================================
    Public Sub EndSub()
        'Dummy sub
    End Sub
    '===============================================================================
    ' Name: Function MD5HashFile
    ' Input:
    '   ByVal strPath As String - File path
    ' Output:
    '   String - MD5 Hash Value
    ' Purpose: Computes an MD5 hash of the specified file.
    ' Remarks: None
    '===============================================================================
    Public Function MD5HashFile(ByVal strPath As String) As String
        System.Diagnostics.Debug.WriteLine("Hashing file " & strPath)
        System.Diagnostics.Debug.WriteLine("File Date: " & FileDateTime(strPath))
        ' read and hash the content
        Dim sData As String = String.Empty
        Dim nFileLen As Integer
        nFileLen = ReadFile(strPath, sData)
        ' use the .NET's native MD5 functions instead of our own MD5 hashing routine
        ' and instead of ALCrypto's md5_hash() function.
        MD5HashFile = LCase(sData)    '<--- ReadFile procedure already computes the MD5.Hash
    End Function
    '===============================================================================
    ' Name: Function FileExists
    ' Input:
    '   ByVal strFile As String - File path and name
    ' Output:
    '   Boolean - True if file exists, False if it doesn't
    ' Purpose: Checks if a file exists in the system.
    ' Remarks: None
    '===============================================================================
    Public Function FileExists(ByVal strFile As String) As Boolean
        FileExists = False
        If File.Exists(strFile) = True Then
            FileExists = True
        End If
    End Function
    '===============================================================================
    ' Name: Function LocalTimeZone
    ' Input:
    '   ByVal returnType As TimeZoneReturn - Type of time zone information being requested
    '       UTC_BaseOffset = UTC offset, not including DST <br>
    '       UTC_Offset = UTC offset, including DST if active <br>
    '       DST_Active = True if DST is currently active, otherwise false <br>
    '       DST_Offset = Offset value for DST (generally -60, if in US)
    ' Output:
    '    Variant - Return type varies depending on returnValue parameter.
    ' Purpose: Retrieves the local time zone.
    ' Remarks: None
    '===============================================================================
    Public Function LocalTimeZone(ByVal returnType As TimeZoneReturn) As Object
        Dim x As Integer
        Dim tzi As TIME_ZONE_INFORMATION
        Dim strName As String
        Dim bDST As Boolean
        Dim rc As Integer

        LocalTimeZone = Nothing

        rc = GetTimeZoneInformation(tzi)
        Select Case rc
            ' if not daylight assume standard
        Case TIME_ZONE_ID_DAYLIGHT
                strName = System.Text.UnicodeEncoding.Unicode.GetString(tzi.DaylightName) ' convert to string
                bDST = True
            Case Else
                strName = System.Text.UnicodeEncoding.Unicode.GetString(tzi.StandardName)
        End Select

        ' name terminates with null
        x = InStr(strName, vbNullChar)
        If x > 0 Then strName = Left(strName, x - 1)

        If returnType = TimeZoneReturn.DST_Active Then
            LocalTimeZone = bDST
        End If

        If returnType = TimeZoneReturn.TimeZoneName Then
            LocalTimeZone = strName
        End If

        If returnType = TimeZoneReturn.TimeZoneCode Then
            LocalTimeZone = Left(strName, 1)
            x = InStr(1, strName, " ")
            Do While x > 0
                LocalTimeZone = LocalTimeZone & Mid(strName, x + 1, 1)
                x = InStr(x + 1, strName, " ")
            Loop
            LocalTimeZone = Trim(LocalTimeZone)
        End If

        If returnType = TimeZoneReturn.UTC_BaseOffset Then
            LocalTimeZone = tzi.bias
        End If

        If returnType = TimeZoneReturn.DST_Offset Then
            LocalTimeZone = tzi.DaylightBias
        End If

        If returnType = TimeZoneReturn.UTC_Offset Then
            If tzi.DaylightBias = -60 Then
                LocalTimeZone = tzi.bias
            Else
                LocalTimeZone = -tzi.bias
            End If
            ' Account for Daylight Savings Time
            If bDST Then LocalTimeZone = LocalTimeZone - 60
        End If
    End Function
    '===============================================================================
    ' Name: Function RSASign
    ' Input:
    '   ByVal strPub As String - RSA Public key blob
    '   ByVal strPriv As String - RSA Private key blob
    '   ByVal strdata As String - Data to be signed
    ' Output:
    '   String - Signature string
    ' Purpose: Performs RSA signing of <code>strData</code> using the specified key.
    ' Remarks: 05.13.05    - alkan  - Removed the modActiveLock references
    '===============================================================================
    Public Function RSASign(ByVal strPub As String, ByVal strPriv As String, ByVal strdata As String) As String
        Dim KEY As RSAKey
        ' create the key from the key blobs
        If rsa_createkey(strPub, Len(strPub), strPriv, Len(strPriv), KEY) = RETVAL_ON_ERROR Then
            Err.Raise(Globals_Renamed.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If

        ' sign the data using the created key
        Dim sLen As Integer
        If rsa_sign(KEY, strdata, Len(strdata), vbNullString, sLen) = RETVAL_ON_ERROR Then
            Err.Raise(Globals_Renamed.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        Dim strSig As String : strSig = New String(Chr(0), sLen)
        If rsa_sign(KEY, strdata, Len(strdata), strSig, sLen) = RETVAL_ON_ERROR Then
            Err.Raise(Globals_Renamed.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        ' throw away the key
        If rsa_freekey(KEY) = RETVAL_ON_ERROR Then
            Err.Raise(Globals_Renamed.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        RSASign = strSig
    End Function
    '===============================================================================
    ' Name: Function RSAVerify
    ' Input:
    '   ByVal strPub As String - Public key blob
    '   ByVal strdata As String - Data to be signed
    '   ByVal strSig As String - Private key blob
    ' Output:
    '   Long - Zero if verification is successful, non-zero otherwise.
    ' Purpose: Verifies an RSA signature.
    ' Remarks: None
    '===============================================================================
    Public Function RSAVerify(ByVal strPub As String, ByVal strdata As String, ByVal strSig As String) As Integer
        Dim KEY As RSAKey = Nothing
        Dim rc As Integer
        ' create the key from the public key blob
        If rsa_createkey(strPub, Len(strPub), vbNullString, 0, KEY) = RETVAL_ON_ERROR Then
            Err.Raise(Globals_Renamed.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        ' validate the key
        rc = rsa_verifysig(KEY, strSig, Len(strSig), strdata, Len(strdata))
        If rc = RETVAL_ON_ERROR Then
            Err.Raise(Globals_Renamed.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        ' de-allocate memory used by the key
        If rsa_freekey(KEY) = RETVAL_ON_ERROR Then
            Err.Raise(Globals_Renamed.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        RSAVerify = rc
    End Function
    '===============================================================================
    ' Name: Function WinError
    ' Input:
    '   ByVal lLastDLLError As Long - Last DLL error as an input
    ' Output:
    '   String - Error message string
    ' Purpose: Retrieves the error text for the specified Windows error code
    ' Remarks: None
    '===============================================================================
    Public Function WinError(ByVal lLastDLLError As Integer) As String
        Dim sBuff As String
        Dim lCount As Integer

        WinError = String.Empty

        ' Return the error message associated with LastDLLError:
        sBuff = New String(Chr(0), 256)
        lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0, sBuff, Len(sBuff), 0)
        If lCount Then
            WinError = Left(sBuff, lCount)
        End If
    End Function
    '===============================================================================
    ' Name: Function WinDir
    ' Input: None
    ' Output:
    '   String - Windows directory path
    ' Purpose: Gets the windows directory
    ' Remarks: None
    '===============================================================================
    Public Function WinDir() As String
        Const FIX_LENGTH As Short = 4096
        Dim Length As Short
        Dim Buffer As New VB6.FixedLengthString(FIX_LENGTH)

        Length = GeneralWinDirApi(Buffer.Value, FIX_LENGTH - 1)
        WinDir = Left(Buffer.Value, Length)
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
    '===============================================================================
    ' Name: Function WinSysDir
    ' Input: None
    ' Output:
    '   String - Windows system directory path
    ' Purpose: Gets the Windows system directory
    ' Remarks: None
    '===============================================================================
    Public Function WinSysDir() As String
        Const FIX_LENGTH As Short = 4096
        Dim Length As Short
        Dim Buffer As New VB6.FixedLengthString(FIX_LENGTH)

        Length = GetSystemDirectory(Buffer.Value, FIX_LENGTH - 1)
        WinSysDir = Left(Buffer.Value, Length)
    End Function
End Module