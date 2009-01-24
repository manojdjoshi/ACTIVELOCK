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
    Public Const STRKEYSTOREPATHISEMPTY As String = "Key Store Path (LIC file path) not specified."
    Public Const STRNOSOFTWARECODE As String = "Software code has not been set."
	Public Const STRNOSOFTWARENAME As String = "Software Name has not been set."
	Public Const STRNOSOFTWAREVERSION As String = "Software Version has not been set."
    Public Const STRNOSOFTWAREPASSWORD As String = "Software Password has not been set."
    Public Const STRUSERNAMETOOLONG As String = "User Name > 2000 characters."
    Public Const STRUSERNAMEINVALID As String = "User Name invalid."
    Public Const STRRSAERROR As String = "Internal RSA Error."
    Public Const RETVAL_ON_ERROR As Integer = -999
    Public Const STRWRONGIPADDRESS As String = "Wrong IP Address."
    Public Const STRUNDEFINEDSPECIALFOLDER As String = "Undefined Special Folder."
    Public Const STRDATEERROR As String = "Date Error."

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

    Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As IntPtr, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean


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

    ' The following constants and declares are used to Get/Set Locale Date format
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
    Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String) As Boolean
    Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Short
    Const LOCALE_SSHORTDATE As Short = &H1FS
    Public regionalSymbol As String

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
        Set_Locale(regionalSymbol)
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
        Dim tzi As TIME_ZONE_INFORMATION = Nothing
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
        Dim KEY As RSAKey = Nothing
        ' create the key from the key blobs
        If rsa_createkey(strPub, Len(strPub), strPriv, Len(strPriv), KEY) = RETVAL_ON_ERROR Then
            Set_Locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If

        ' sign the data using the created key
        Dim sLen As Integer
        If rsa_sign(KEY, strdata, Len(strdata), vbNullString, sLen) = RETVAL_ON_ERROR Then
            Set_Locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        Dim strSig As String : strSig = New String(Chr(0), sLen)
        If rsa_sign(KEY, strdata, Len(strdata), strSig, sLen) = RETVAL_ON_ERROR Then
            Set_Locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        ' throw away the key
        If rsa_freekey(KEY) = RETVAL_ON_ERROR Then
            Set_Locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
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
            Set_Locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        ' validate the key
        rc = rsa_verifysig(KEY, strSig, Len(strSig), strdata, Len(strdata))
        If rc = RETVAL_ON_ERROR Then
            Set_Locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        ' de-allocate memory used by the key
        If rsa_freekey(KEY) = RETVAL_ON_ERROR Then
            Set_Locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
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
    '===============================================================================
    ' Name: Function UTC
    ' Input:
    '   ByRef dt As Date - Date-Time input to be converted into UTC Date-Time
    ' Output:
    '   Date - UTC Date-Time
    ' Purpose: Converts a local date-time into UTC/GMT date-time
    ' Remarks: None
    '===============================================================================
    Public Function UTC(ByVal dt As Date) As Date
        '  Returns current UTC date-time.
        UTC = dt.AddMinutes(LocalTimeZone(TimeZoneReturn.UTC_Offset))
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
    Public Function ActivelockGetSpecialFolder(ByVal CSIDL As Long) As String
        'value  ID                      Result
        '2      CSIDL_PROGRAMS          C:\Documents and Settings\<USERNAME>\Start Menu\Programs
        '5      CSIDL_PERSONAL          C:\Documents and Settings\<USERNAME>\My Documents
        '6      CSIDL_FAVORITES         C:\Documents and Settings\<USERNAME>\Favorites
        '7      CSIDL_STARTUP           C:\Documents and Settings\<USERNAME>\Start Menu\Programs\Startup
        '8      CSIDL_RECENT            C:\Documents and Settings\<USERNAME>\Recent
        '9      CSIDL_SENDTO            C:\Documents and Settings\<USERNAME>\SendTo
        '11     CSIDL_STARTMENU         C:\Documents and Settings\<USERNAME>\Start Menu
        '13     CSIDL_MYMUSIC           C:\Documents and Settings\<USERNAME>\My Documents\My Music
        '14     CSIDL_MYVIDEO           C:\Documents and Settings\<USERNAME>\My Documents\My Video
        '16     CSIDL_DESKTOPDIRECTORY  C:\Documents and Settings\<USERNAME>\Desktop
        '19     CSIDL_NETHOOD           C:\Documents and Settings\<USERNAME>\NetHood
        '20     Virtual - don't use     CSIDL_FONTS             C:\Windows\Fonts
        '21     CSIDL_TEMPLATES         C:\Documents and Settings\<USERNAME>\Templates
        '22     CSIDL_COMMON_STARTMENU  C:\Documents and Settings\All Users\Start Menu
        '23     CSIDL_COMMON_PROGRAMS   C:\Documents and Settings\All Users\Start Menu\Programs
        '24     CSIDL_COMMON_STARTUP    C:\Documents and Settings\All Users\Start Menu\Programs\Startup
        '25     CSIDL_COMMON_DESKTOPDIRECTORY     C:\Documents and Settings\All Users\Desktop
        '26     CSIDL_APPDATA           C:\Documents and Settings\<USERNAME>\Application Data
        '27     CSIDL_PRINTHOOD         C:\Documents and Settings\<USERNAME>\PrintHood
        '28     CSIDL_LOCAL_APPDATA     C:\Documents and Settings\<USERNAME>\Local Settings\Application Data
        '31     CSIDL_COMMON_FAVORITES  C:\Documents and Settings\All Users\Favorites
        '32     CSIDL_INTERNET_CACHE    C:\Documents and Settings\<USERNAME>\Local Settings\Temporary Internet Files
        '33     CSIDL_COOKIES           C:\Documents and Settings\<USERNAME>\Cookies
        '34     CSIDL_HISTORY           C:\Documents and Settings\<USERNAME>\Local Settings\History
        '35     CSIDL_COMMON_APPDATA    C:\Documents and Settings\All Users\Application Data
        '36     CSIDL_WINDOWS           C:\Windows
        '37     CSIDL_SYSTEM            C:\Windows\system32
        '38     CSIDL_PROGRAM_FILES     C:\Program Files
        '39     CSIDL_MYPICTURES        C:\Documents and Settings\<USERNAME>\My Documents\My Pictures
        '40     CSIDL_PROFILE           C:\Documents and Settings\<USERNAME>
        '41     CSIDL_SYSTEMX86         C:\Windows\system32  'x86 system directory on RISC
        '43     CSIDL_PROGRAM_FILES_COMMON     C:\Program Files\Common Files
        '45     CSIDL_COMMON_TEMPLATES  C:\Documents and Settings\All Users\Templates
        '46     CSIDL_COMMON_DOCUMENTS  C:\Documents and Settings\All Users\Documents
        '47     CSIDL_COMMON_ADMINTOOLS C:\Documents and Settings\All Users\Start Menu\Programs\Administrative Tools
        '48     CSIDL_ADMINTOOLS        C:\Documents and Settings\<USERNAME>\Start Menu\Programs\Administrative Tools
        '53     CSIDL_COMMON_MUSIC      C:\Documents and Settings\All Users\My Documents\My Music
        '54     CSIDL_COMMON_PICTURES   C:\Documents and Settings\All Users\My Documents\My Pictures
        '55     CSIDL_COMMON_VIDEO      C:\Documents and Settings\All Users\My Documents\My Video

        ' CLSIDs for XP and older
        'CSIDL_DESKTOP = &H0
        'CSIDL_INTERNET = &H1
        'CSIDL_PROGRAMS = &H2
        'CSIDL_CONTROLS = &H3
        'CSIDL_PRINTERS = &H4
        'CSIDL_PERSONAL = &H5
        'CSIDL_FAVORITES = &H6
        'CSIDL_STARTUP = &H7
        'CSIDL_RECENT = &H8
        'CSIDL_SENDTO = &H9
        'CSIDL_BITBUCKET = &HA
        'CSIDL_STARTMENU = &HB
        'CSIDL_MYDOCUMENTS = &HC
        'CSIDL_MYMUSIC = &HD
        'CSIDL_MYVIDEO = &HE
        'CSIDL_DESKTOPDIRECTORY = &H10
        'CSIDL_DRIVES = &H11
        'CSIDL_NETWORK = &H12
        'CSIDL_NETHOOD = &H13
        'CSIDL_FONTS = &H14
        'CSIDL_TEMPLATES = &H15
        'CSIDL_COMMON_STARTMENU = &H16
        'CSIDL_COMMON_PROGRAMS = &H17
        'CSIDL_COMMON_STARTUP = &H18
        'CSIDL_COMMON_DESKTOPDIRECTORY = &H19
        'CSIDL_APPDATA = &H1A
        'CSIDL_PRINTHOOD = &H1B
        'CSIDL_LOCAL_APPDATA = &H1C
        'CSIDL_ALTSTARTUP = &H1D
        'CSIDL_COMMON_ALTSTARTUP = &H1E
        'CSIDL_COMMON_FAVORITES = &H1F
        'CSIDL_INTERNET_CACHE = &H20
        'CSIDL_COOKIES = &H21
        'CSIDL_HISTORY = &H22
        'CSIDL_COMMON_APPDATA = &H23
        'CSIDL_WINDOWS = &H24
        'CSIDL_SYSTEM = &H25
        'CSIDL_PROGRAM_FILES = &H26
        'CSIDL_MYPICTURES = &H27
        'CSIDL_PROFILE = &H28
        'CSIDL_SYSTEMX86 = &H29
        'CSIDL_PROGRAM_FILESX86 = &H2A
        'CSIDL_PROGRAM_FILES_COMMON = &H2B
        'CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
        'CSIDL_COMMON_TEMPLATES = &H2D
        'CSIDL_COMMON_DOCUMENTS = &H2E
        'CSIDL_COMMON_ADMINTOOLS = &H2F
        'CSIDL_ADMINTOOLS = &H30
        'CSIDL_CONNECTIONS = &H31
        'CSIDL_COMMON_MUSIC = &H35
        'CSIDL_COMMON_PICTURES = &H36
        'CSIDL_COMMON_VIDEO = &H37
        'CSIDL_RESOURCES = &H38
        'CSIDL_RESOURCES_LOCALIZED = &H39
        'CSIDL_COMMON_OEM_LINKS = &H3A
        'CSIDL_CDBURN_AREA = &H3B
        'CSIDL_COMPUTERSNEARME = &H3D
        'CSIDL_FLAG_PER_USER_INIT = &H800
        'CSIDL_FLAG_NO_ALIAS = &H1000
        'CSIDL_FLAG_DONT_VERIFY = &H4000
        'CSIDL_FLAG_CREATE = &H8000
        'CSIDL_FLAG_MASK = &HFF00

        'KNOWNFOLDER IDs for Vista
        'Select Case CSIDL

        '    Case CSIDL_NETWORK, CSIDL_COMPUTERSNEARME  ' VIRTUAL
        '        FOLDERID_NetworkFolder = "{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}"
        '        sRfid = FOLDERID_NetworkFolder
        '    Case CSIDL_DRIVES  ' VIRTUAL
        '        FOLDERID_ComputerFolder = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"
        '        sRfid = FOLDERID_ComputerFolder
        '    Case CSIDL_INTERNET  ' VIRTUAL
        '        FOLDERID_InternetFolder = "{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}"
        '        sRfid = FOLDERID_InternetFolder
        '    Case CSIDL_CONTROLS  ' VIRTUAL
        '        FOLDERID_ControlPanelFolder = "{82A74AEB-AEB4-465C-A014-D097EE346D63}"
        '        sRfid = FOLDERID_ControlPanelFolder
        '    Case CSIDL_PRINTERS  ' VIRTUAL
        '        FOLDERID_PrintersFolder = "{76FC4E2D-D6AD-4519-A663-37BD56068185}"
        '        sRfid = FOLDERID_PrintersFolder
        '    Case 101  ' VIRTUAL
        '        FOLDERID_SyncManagerFolder = "{43668BF8-C14E-49B2-97C9-747784D784B7}"
        '        sRfid = FOLDERID_SyncManagerFolder
        '    Case 102  ' VIRTUAL
        '        FOLDERID_SyncSetupFolder = "{0F214138-B1D3-4a90-BBA9-27CBC0C5389A}"
        '        sRfid = FOLDERID_SyncSetupFolder
        '    Case 103  ' VIRTUAL
        '        FOLDERID_ConflictFolder = "{4bfefb45-347d-4006-a5be-ac0cb0567192}"
        '        sRfid = FOLDERID_ConflictFolder
        '    Case 104  ' VIRTUAL
        '        FOLDERID_SyncResultsFolder = "{289a9a43-be44-4057-a41b-587a76d7e7f9}"
        '        sRfid = FOLDERID_SyncResultsFolder
        '    Case CSIDL_BITBUCKET  ' VIRTUAL
        '        FOLDERID_RecycleBinFolder = "{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}"
        '        sRfid = FOLDERID_RecycleBinFolder
        '    Case CSIDL_CONNECTIONS  ' VIRTUAL
        '        FOLDERID_ConnectionsFolder = "{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}"
        '        sRfid = FOLDERID_ConnectionsFolder
        '        ' VISTA - %windir%\Fonts
        '        ' XP - %windir%\Fonts
        '    Case CSIDL_FONTS  ' FIXED
        '        FOLDERID_Fonts = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
        '        sRfid = FOLDERID_Fonts
        '        ' VISTA - %USERPROFILE%\Desktop
        '        ' XP - %USERPROFILE%\Desktop
        '    Case CSIDL_DESKTOP, CSIDL_DESKTOPDIRECTORY  ' PERUSER
        '        FOLDERID_Desktop = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
        '        sRfid = FOLDERID_Desktop
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu\Programs\StartUp
        '        ' XP - %USERPROFILE%\Start Menu\Programs\StartUp
        '    Case CSIDL_STARTUP, CSIDL_ALTSTARTUP  ' PERUSER
        '        FOLDERID_Startup = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"
        '        sRfid = FOLDERID_Startup
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu\Programs
        '        ' XP - %USERPROFILE%\Start Menu\Programs
        '    Case CSIDL_PROGRAMS  ' PERUSER
        '        FOLDERID_Programs = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"
        '        sRfid = FOLDERID_Programs
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu
        '        ' XP - %USERPROFILE%\Start Menu
        '    Case CSIDL_STARTMENU  'PERUSER
        '        FOLDERID_StartMenu = "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"
        '        sRfid = FOLDERID_StartMenu
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Recent
        '        ' XP - %USERPROFILE%\Recent
        '    Case CSIDL_RECENT  ' PERUSER
        '        FOLDERID_Recent = "{AE50C081-EBD2-438A-8655-8A092E34987A}"
        '        sRfid = FOLDERID_Recent
        '        ' VISTA - %APPDATA%\Microsoft\Windows\SendTo
        '        ' XP - %USERPROFILE%\SendTo
        '    Case CSIDL_SENDTO  ' PERUSER
        '        FOLDERID_SendTo = "{8983036C-27C0-404B-8F08-102D10DCFD74}"
        '        sRfid = FOLDERID_SendTo
        '        ' VISTA - %USERPROFILE%\Documents
        '        ' XP - %USERPROFILE%\My Documents
        '    Case CSIDL_MYDOCUMENTS, CSIDL_PERSONAL  ' PERUSER
        '        FOLDERID_Documents = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
        '        sRfid = FOLDERID_Documents
        '        ' VISTA - %USERPROFILE%\Documents
        '        ' XP - %USERPROFILE%\My Documents
        '    Case CSIDL_FAVORITES, CSIDL_COMMON_FAVORITES  ' PERUSER
        '        FOLDERID_Favorites = "{1777F761-68AD-4D8A-87BD-30B759FA33DD}"
        '        sRfid = FOLDERID_Favorites
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Network Shortcuts
        '        ' XP - %USERPROFILE%\NetHood
        '    Case CSIDL_NETHOOD  ' PERUSER
        '        FOLDERID_NetHood = "{C5ABBF53-E17F-4121-8900-86626FC2C973}"
        '        sRfid = FOLDERID_NetHood
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Printer Shortcuts
        '        ' XP - %USERPROFILE%\PrintHood
        '    Case CSIDL_PRINTHOOD  ' PERUSER
        '        FOLDERID_PrintHood = "{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"
        '        sRfid = FOLDERID_PrintHood
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Templates
        '        ' XP - %USERPROFILE%\Templates
        '    Case CSIDL_TEMPLATES  ' PERUSER
        '        FOLDERID_Templates = "{A63293E8-664E-48DB-A079-DF759E0509F7}"
        '        sRfid = FOLDERID_Templates
        '        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\StartUp
        '        ' XP - %ALLUSERSPROFILE%\Start Menu\Programs\StartUp
        '    Case CSIDL_COMMON_STARTUP, CSIDL_COMMON_ALTSTARTUP  ' COMMON
        '        FOLDERID_CommonStartup = "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"
        '        sRfid = FOLDERID_CommonStartup
        '        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs
        '        ' XP - %ALLUSERSPROFILE%\Start Menu\Programs
        '    Case CSIDL_COMMON_PROGRAMS  ' COMMON
        '        FOLDERID_CommonPrograms = "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"
        '        sRfid = FOLDERID_CommonPrograms
        '        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu
        '        ' XP - %ALLUSERSPROFILE%\Start Menu
        '    Case CSIDL_COMMON_STARTMENU  ' COMMON
        '        FOLDERID_CommonStartMenu = "{A4115719-D62E-491D-AA7C-E74B8BE3B067}"
        '        sRfid = FOLDERID_CommonStartMenu
        '        ' VISTA - %PUBLIC%\Desktop
        '        ' XP - %ALLUSERSPROFILE%\Desktop
        '    Case 201  ' COMMON
        '        FOLDERID_PublicDesktop = "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"
        '        sRfid = FOLDERID_PublicDesktop
        '        ' VISTA - %ALLUSERSPROFILE% (%ProgramData%, %SystemDrive%\ProgramData)
        '        ' XP - %ALLUSERSPROFILE%\Application Data
        '    Case CSIDL_COMMON_APPDATA  ' FIXED
        '        FOLDERID_ProgramData = "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"
        '        sRfid = FOLDERID_ProgramData
        '        ' VISTA - %ALLUSERSPROFILE%\Templates
        '        ' XP - %ALLUSERSPROFILE%\Templates
        '    Case CSIDL_COMMON_TEMPLATES  ' COMMON
        '        FOLDERID_CommonTemplates = "{B94237E7-57AC-4347-9151-B08C6C32D1F7}"
        '        sRfid = FOLDERID_CommonTemplates
        '        ' VISTA - %PUBLIC%\Documents
        '        ' XP - %ALLUSERSPROFILE%\Documents
        '    Case CSIDL_COMMON_DOCUMENTS  ' COMMON
        '        FOLDERID_PublicDocuments = "{ED4824AF-DCE4-45A8-81E2-FC7965083634}"
        '        sRfid = FOLDERID_PublicDocuments
        '        ' VISTA - %APPDATA% (%USERPROFILE%\AppData\Roaming)
        '        ' XP - %APPDATA% (%USERPROFILE%\Application Data)
        '    Case CSIDL_APPDATA  ' PERUSER
        '        FOLDERID_RoamingAppData = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
        '        sRfid = FOLDERID_RoamingAppData
        '        ' VISTA - %LOCALAPPDATA% (%USERPROFILE%\AppData\Local)
        '        ' XP - %USERPROFILE%\Local Settings\Application Data
        '    Case CSIDL_LOCAL_APPDATA  ' PERUSER
        '        FOLDERID_LocalAppData = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
        '        sRfid = FOLDERID_LocalAppData
        '        ' VISTA - %USERPROFILE%\AppData\LocalLow
        '        ' XP - NONE
        '    Case 301  ' PERUSER
        '        FOLDERID_LocalAppDataLow = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
        '        sRfid = FOLDERID_LocalAppDataLow
        '        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\Temporary Internet Files
        '        ' XP - %USERPROFILE%\Local Settings\Temporary Internet Files
        '    Case CSIDL_INTERNET_CACHE  ' PERUSER
        '        FOLDERID_InternetCache = "{352481E8-33BE-4251-BA85-6007CAEDCF9D}"
        '        sRfid = FOLDERID_InternetCache
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Cookies
        '        ' XP - %USERPROFILE%\Cookies
        '    Case CSIDL_COOKIES  ' PERUSER
        '        FOLDERID_Cookies = "{2B0F765D-C0E9-4171-908E-08A611B84FF6}"
        '        sRfid = FOLDERID_Cookies
        '        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\History
        '        ' XP - %USERPROFILE%\Local Settings\History
        '    Case CSIDL_HISTORY  ' PERUSER
        '        FOLDERID_History = "{D9DC8A3B-B784-432E-A781-5A1130A75963}"
        '        sRfid = FOLDERID_History
        '        ' VISTA - %windir%\system32
        '        ' XP - %windir%\system32
        '    Case CSIDL_SYSTEM  ' FIXED
        '        FOLDERID_System = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
        '        sRfid = FOLDERID_System
        '        ' VISTA - %windir%\system32
        '        ' XP - %windir%\system32
        '    Case CSIDL_SYSTEMX86  ' FIXED
        '        FOLDERID_SystemX86 = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"
        '        sRfid = FOLDERID_SystemX86
        '        ' VISTA - %windir%
        '        ' XP - %windir%
        '    Case CSIDL_WINDOWS  ' FIXED
        '        FOLDERID_Windows = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
        '        sRfid = FOLDERID_Windows
        '        ' VISTA - %USERPROFILE% (%SystemDrive%\Users\%USERNAME%)
        '        ' XP - %USERPROFILE% (%SystemDrive%\Documents and Settings\%USERNAME%)
        '    Case CSIDL_PROFILE  ' FIXED
        '        FOLDERID_Profile = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
        '        sRfid = FOLDERID_Profile
        '        ' VISTA - %USERPROFILE%\Pictures
        '        ' XP - %USERPROFILE%\My Documents\My Pictures
        '    Case CSIDL_MYPICTURES  ' PERUSER
        '        FOLDERID_Pictures = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
        '        sRfid = FOLDERID_Pictures
        '        ' VISTA - %ProgramFiles% (%SystemDrive%\Program Files)
        '        ' XP - %ProgramFiles% (%SystemDrive%\Program Files)
        '    Case CSIDL_PROGRAM_FILESX86  ' FIXED
        '        FOLDERID_ProgramFilesX86 = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
        '        sRfid = FOLDERID_ProgramFilesX86
        '        ' VISTA - %ProgramFiles%\Common Files
        '        ' XP - %ProgramFiles%\Common Files
        '    Case CSIDL_PROGRAM_FILES_COMMONX86  ' FIXED
        '        FOLDERID_ProgramFilesCommonX86 = "{DE974D24-D9C6-4D3E-BF91-F4455120B917}"
        '        sRfid = FOLDERID_ProgramFilesCommonX86
        '        ' VISTA - %ProgramFiles% (%SystemDrive%\Program Files)
        '        ' XP - %ProgramFiles% (%SystemDrive%\Program Files)
        '    Case 401  ' FIXED
        '        FOLDERID_ProgramFilesX64 = "{6D809377-6AF0-444b-8957-A3773F02200E}"
        '        sRfid = FOLDERID_ProgramFilesX64
        '        ' VISTA - %ProgramFiles%\Common Files
        '        ' XP - %ProgramFiles%\Common Files
        '    Case 402
        '        FOLDERID_ProgramFilesCommonX64 = "{6365D5A7-0F0D-45e5-87F6-0DA56B6A4F7D}"
        '        sRfid = FOLDERID_ProgramFilesCommonX64
        '        ' VISTA - %ProgramFiles% (%SystemDrive%\Program Files)
        '        ' XP - %ProgramFiles% (%SystemDrive%\Program Files)
        '    Case CSIDL_PROGRAM_FILES  ' FIXED
        '        FOLDERID_ProgramFiles = "{905e63b6-c1bf-494e-b29c-65b732d3d21a}"
        '        sRfid = FOLDERID_ProgramFiles
        '        ' VISTA - %ProgramFiles%\Common Files
        '        ' XP - %ProgramFiles%\Common Files
        '    Case CSIDL_PROGRAM_FILES_COMMON  ' FIXED
        '        FOLDERID_ProgramFilesCommon = "{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"
        '        sRfid = FOLDERID_ProgramFilesCommon
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu\Programs\Administrative Tools
        '        ' XP - %USERPROFILE%\Start Menu\Programs\Administrative Tools
        '    Case CSIDL_ADMINTOOLS  ' PERUSER
        '        FOLDERID_AdminTools = "{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"
        '        sRfid = FOLDERID_AdminTools
        '        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Administrative Tools
        '        ' XP - %ALLUSERSPROFILE%\Start Menu\Programs\Administrative Tools
        '    Case CSIDL_COMMON_ADMINTOOLS  ' COMMON
        '        FOLDERID_CommonAdminTools = "{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"
        '        sRfid = FOLDERID_CommonAdminTools
        '        ' VISTA - %USERPROFILE%\Music
        '        ' XP - %USERPROFILE%\My Documents\My Music
        '    Case CSIDL_MYMUSIC  ' PERUSER
        '        FOLDERID_Music = "{4BD8D571-6D19-48D3-BE97-422220080E43}"
        '        sRfid = FOLDERID_Music
        '        ' VISTA - %USERPROFILE%\Videos
        '        ' XP - %USERPROFILE%\My Documents\My Videos
        '    Case CSIDL_MYVIDEO  ' PERUSER
        '        FOLDERID_Videos = "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"
        '        sRfid = FOLDERID_Videos
        '        ' VISTA - %PUBLIC%\Pictures
        '        ' XP - %ALLUSERSPROFILE%\Documents\My Pictures
        '    Case CSIDL_COMMON_PICTURES  ' COMMON
        '        FOLDERID_PublicPictures = "{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"
        '        sRfid = FOLDERID_PublicPictures
        '        ' VISTA - %PUBLIC%\Music
        '        ' XP - %ALLUSERSPROFILE%\Documents\My Music
        '    Case CSIDL_COMMON_MUSIC  ' COMMON
        '        FOLDERID_PublicMusic = "{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"
        '        sRfid = FOLDERID_PublicMusic
        '        'VISTA - %PUBLIC%\Videos
        '        ' XP - %ALLUSERSPROFILE%\Documents\My Videos
        '    Case CSIDL_COMMON_VIDEO  ' COMMON
        '        FOLDERID_PublicVideos = "{2400183A-6185-49FB-A2D8-4A392A602BA3}"
        '        sRfid = FOLDERID_PublicVideos
        '        ' VISTA - %windir%\Resources
        '        ' XP - %windir%\Resources
        '    Case CSIDL_RESOURCES  ' FIXED
        '        FOLDERID_ResourceDir = "{8AD10C31-2ADB-4296-A8F7-E4701232C972}"
        '        sRfid = FOLDERID_ResourceDir
        '        ' VISTA - %windir%\resources\0409 (code page)
        '        ' XP - %windir%\resources\0409 (code page)
        '    Case CSIDL_RESOURCES_LOCALIZED  ' FIXED
        '        FOLDERID_LocalizedResourcesDir = "{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"
        '        sRfid = FOLDERID_LocalizedResourcesDir
        '        ' VISTA - %ALLUSERSPROFILE%\OEM Links
        '        ' XP - %ALLUSERSPROFILE%\OEM Links
        '    Case CSIDL_COMMON_OEM_LINKS  ' COMMON
        '        FOLDERID_CommonOEMLinks = "{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"
        '        sRfid = FOLDERID_CommonOEMLinks
        '        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\Burn\Burn
        '        ' XP - %USERPROFILE%\Local Settings\Application Data\Microsoft\CD Burning
        '    Case CSIDL_CDBURN_AREA  ' PERUSER
        '        FOLDERID_CDBurning = "{9E52AB10-F80D-49DF-ACB8-4330F5687855}"
        '        sRfid = FOLDERID_CDBurning
        '        ' VISTA - %SystemDrive%\Users
        '        ' XP - NONE
        '    Case 501  ' FIXED
        '        FOLDERID_UserProfiles = "{0762D272-C50A-4BB0-A382-697DCD729B80}"
        '        sRfid = FOLDERID_UserProfiles
        '        ' VISTA - %USERPROFILE%\Music\Playlists
        '        ' XP - NONE
        '    Case 502  ' PERUSER
        '        FOLDERID_Playlists = "{DE92C1C7-837F-4F69-A3BB-86E631204A23}"
        '        sRfid = FOLDERID_Playlists
        '        ' VISTA - %PUBLIC%\Music\Sample Playlists
        '        ' XP - NONE
        '    Case 503  ' COMMON
        '        FOLDERID_SamplePlaylists = "{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"
        '        sRfid = FOLDERID_SamplePlaylists
        '        ' VISTA - %PUBLIC%\Music\Sample Music
        '        ' XP - %ALLUSERSPROFILE%\Documents\My Music\Sample Music
        '    Case 504  ' COMMON
        '        FOLDERID_SampleMusic = "{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"
        '        sRfid = FOLDERID_SampleMusic
        '        ' VISTA - %PUBLIC%\Pictures\Sample Pictures
        '        ' XP - %ALLUSERSPROFILE%\Documents\My Pictures\Sample Pictures
        '    Case 505  ' COMMON
        '        FOLDERID_SamplePictures = "{C4900540-2379-4C75-844B-64E6FAF8716B}"
        '        sRfid = FOLDERID_SamplePictures
        '        ' VISTA - %PUBLIC%\Videos\Sample Videos
        '        ' XP - NONE
        '    Case 506  ' COMMON
        '        FOLDERID_SampleVideos = "{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"
        '        sRfid = FOLDERID_SampleVideos
        '        ' VISTA - %USERPROFILE%\Pictures\Slide Shows
        '        ' XP - NONE
        '    Case 507  ' PERUSER
        '        FOLDERID_PhotoAlbums = "{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"
        '        sRfid = FOLDERID_PhotoAlbums
        '        ' VISTA - %PUBLIC% (%SystemDrive%\Users\Public)
        '        ' XP - NONE
        '    Case 508  'FIXED
        '        FOLDERID_Public = "{DFDF76A2-C82A-4D63-906A-5644AC457385}"
        '        sRfid = FOLDERID_Public
        '    Case 509  ' VIRTUAL
        '        FOLDERID_ChangeRemovePrograms = "{df7266ac-9274-4867-8d55-3bd661de872d}"
        '        sRfid = FOLDERID_ChangeRemovePrograms
        '    Case 510  ' VIRTUAL
        '        FOLDERID_AppUpdates = "{a305ce99-f527-492b-8b1a-7e76fa98d6e4}"
        '        sRfid = FOLDERID_AppUpdates
        '    Case 511  ' VIRTUAL
        '        FOLDERID_AddNewPrograms = "{de61d971-5ebc-4f02-a3a9-6c82895e5c04}"
        '        sRfid = FOLDERID_AddNewPrograms
        '        ' VISTA - %USERPROFILE%\Downloads
        '        ' XP - NONE
        '    Case 512  ' PERUSER
        '        FOLDERID_Downloads = "{374DE290-123F-4565-9164-39C4925E467B}"
        '        sRfid = FOLDERID_Downloads
        '        ' VISTA - %PUBLIC%\Downloads
        '        ' XP - NONE
        '    Case 513  ' COMMON
        '        FOLDERID_PublicDownloads = "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}"
        '        sRfid = FOLDERID_PublicDownloads
        '        ' VISTA - %USERPROFILE%\Searches
        '        ' XP - NONE
        '    Case 514  ' PERUSER
        '        FOLDERID_SavedSearches = "{7d1d3a04-debb-4115-95cf-2f29da2920da}"
        '        sRfid = FOLDERID_SavedSearches
        '        'VISTA - %APPDATA%\Microsoft\Internet Explorer\Quick Launch
        '        'XP - %APPDATA%\Microsoft\Internet Explorer\Quick Launch
        '    Case 515  ' PERUSER
        '        FOLDERID_QuickLaunch = "{52a4f021-7b75-48a9-9f6b-4b87a210bc8f}"
        '        sRfid = FOLDERID_QuickLaunch
        '        ' VISTA - %USERPROFILE%\Contacts
        '        ' XP - NONE
        '    Case 516  ' PERUSER
        '        FOLDERID_Contacts = "{56784854-C6CB-462b-8169-88E350ACB882}"
        '        sRfid = FOLDERID_Contacts
        '        ' VISTA -%LOCALAPPDATA%\Microsoft\Windows Sidebar\Gadgets
        '        ' XP - NONE
        '    Case 517  ' PERUSER
        '        FOLDERID_SidebarParts = "{A75D362E-50FC-4fb7-AC2C-A8BEAA314493}"
        '        sRfid = FOLDERID_SidebarParts
        '        ' VISTA - %ProgramFiles%\Windows Sidebar\Gadgets
        '        ' XP - NONE
        '    Case 518  ' COMMON
        '        FOLDERID_SidebarDefaultParts = "{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"
        '        sRfid = FOLDERID_SidebarDefaultParts
        '    Case 519  ' NOT USED
        '        FOLDERID_TreeProperties = "{5b3749ad-b49f-49c1-83eb-15370fbd4882}"
        '        sRfid = FOLDERID_TreeProperties
        '        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\GameExplorer
        '        ' XP - NONE
        '    Case 520  ' COMMON
        '        FOLDERID_PublicGameTasks = "{DEBF2536-E1A8-4c59-B6A2-414586476AEA}"
        '        sRfid = FOLDERID_PublicGameTasks
        '        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\GameExplorer
        '        ' XP - NONE
        '    Case 521  ' PERUSER
        '        FOLDERID_GameTasks = "{054FAE61-4DD8-4787-80B6-090220C4B700}"
        '        sRfid = FOLDERID_GameTasks
        '        ' VISTA - %USERPROFILE%\Saved Games
        '        ' XP - NONE
        '    Case 522  ' PERUSER
        '        FOLDERID_SavedGames = "{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}"
        '        sRfid = FOLDERID_SavedGames
        '    Case 523  ' VIRTUAL
        '        FOLDERID_Games = "{CAC52C1A-B53D-4edc-92D7-6B2E8AC19434}"
        '        sRfid = FOLDERID_Games
        '    Case 524  ' NOT USED
        '        FOLDERID_RecordedTV = "{bd85e001-112e-431e-983b-7b15ac09fff1}"
        '        sRfid = FOLDERID_RecordedTV
        '    Case 525  ' VIRTUAL
        '        FOLDERID_SEARCH_MAPI = "{98ec0e18-2098-4d44-8644-66979315a281}"
        '        sRfid = FOLDERID_SEARCH_MAPI
        '    Case 526  ' VIRTUAL
        '        FOLDERID_SEARCH_CSC = "{ee32e446-31ca-4aba-814f-a5ebd2fd6d5e}"
        '        sRfid = FOLDERID_SEARCH_CSC
        '        ' VISTA - %USERPROFILE%\Links
        '        ' XP - NONE
        '    Case 527  ' PERUSER
        '        FOLDERID_Links = "{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}"
        '        sRfid = FOLDERID_Links
        '    Case 528  ' VIRTUAL
        '        FOLDERID_UsersFiles = "{f3ce0f7c-4901-4acc-8648-d5d44b04ef8f}"
        '        sRfid = FOLDERID_UsersFiles
        '    Case 529  ' VIRTUAL
        '        FOLDERID_SearchHome = "{190337d1-b8ca-4121-a639-6d472d16972a}"
        '        sRfid = FOLDERID_SearchHome
        '        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows Photo Gallery\Original Images
        '        ' XP - NONE
        '    Case 530  ' PERUSER
        '        FOLDERID_OriginalImages = "{2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}"
        '        sRfid = FOLDERID_OriginalImages
        'End Select
        Dim Ret As Long
        Dim Trash As String
        Trash = Space$(260)
        ActivelockGetSpecialFolder = ""
        Try
            Ret = SHGetSpecialFolderPath(0, Trash, CSIDL, False)
            If Trim$(Trash) <> Chr(0) Then
                Trash = Left$(Trash, InStr(Trash, Chr(0)) - 1)
            End If
            ActivelockGetSpecialFolder = Trash
            Return ActivelockGetSpecialFolder
        Catch ex As Exception
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrUndefinedSpecialFolder, ACTIVELOCKSTRING, STRUNDEFINEDSPECIALFOLDER)
        End Try

    End Function
    Public Function PSWD() As String
        ' Do not modify this unless you change all encrypted strings in the entire project
        PSWD = Chr(109) & Chr(121) & Chr(108) & Chr(111) & Chr(118) & Chr(101) & Chr(97) & Chr(99) & Chr(116) & Chr(105) & Chr(118) & Chr(101) & "lock"
    End Function
    Function IsNumberIncluded(ByVal n1 As Long, ByVal n2 As Long) As Boolean
        ' n1 = the larger number which may include n2
        ' n2 = the number we're checking as "is this a component?"
        Dim binary1 As String = String.Empty
        Dim binary2 As String = String.Empty
        IsNumberIncluded = False

        If n1 < n2 Then
            Exit Function
        ElseIf n1 <= 0 Or n2 <= 0 Then
            Exit Function
        ElseIf n1 = n2 Then
            IsNumberIncluded = True
        Else

            Do Until n1 = 0
                If (n1 Mod 2) Then
                    binary1 = binary1 & "1"   'write binary number BACKWARDS
                Else
                    binary1 = binary1 & "0"
                End If
                n1 = n1 \ 2
            Loop

            Do Until n2 = 0
                If (n2 Mod 2) Then
                    binary2 = binary2 & "1"   'write binary number BACKWARDS
                Else
                    binary2 = binary2 & "0"
                End If
                n2 = n2 \ 2
            Loop
            IsNumberIncluded = CBool(Mid$(binary1, Len(binary2), 1) = "1")
            If binary1.Substring(binary2.Length - 1, 1) = "1" Then IsNumberIncluded = True

        End If


    End Function

End Module