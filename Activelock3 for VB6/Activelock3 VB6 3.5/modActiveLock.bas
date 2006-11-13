Attribute VB_Name = "modActiveLock"
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
' Modified: 08.15.2005
'===============================================================================
' @author activelock-admins
' @version 3.0.0
' @date 20050421
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
'   08.04.05 - ialkan  - Renamed ALCrypto.dll as alcrypto3.dll for V3 of Activelock
' </pre>

'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Private Module
Option Explicit
Option Base 0

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
Public Const STRUSERNAMETOOLONG As String = "User Name > 2000 characters."
Public Const STRUSERNAMEINVALID As String = "User Name invalid."
Public Const STRRSAERROR As String = "Internal RSA Error."
Public Const RETVAL_ON_ERROR As Long = -999

' RSA encrypts the data.
' @param CryptType CryptType = 0 for public&#59; 1 for private
' @param Data   Data to be encrypted
' @param dLen   [in/out] Length of data, in bytes. This parameter will contain length of encrypted data when returned.
' @param ptrKey Key to be used for encryption
Public Declare Function rsa_encrypt Lib "ALCrypto3" (ByVal CryptType As Long, ByVal data As String, dLen As Long, ptrKey As RSAKey) As Long


' RSA decrypts the data.
' @param CryptType CryptType = 0 for public&#59; 1 for private
' @param Data   Data to be encrypted
' @param dLen   [in/out] Length of data, in bytes. This parameter will contain length of encrypted data when returned.
' @param ptrKey Key to be used for encryption
Public Declare Function rsa_decrypt Lib "ALCrypto3" (ByVal CryptType As Long, ByVal data As String, dLen As Long, ptrKey As RSAKey) As Long


' Computes an MD5 hash from the data.
' @param inData Data to be hashed
' @param nDataLen   Length of inData
' @param outData    [out] 32-byte Computed hash code
Public Declare Function md5_hash Lib "ALCrypto3" (ByVal inData As String, ByVal nDataLen As Long, ByVal outData As String) As Long

Type SYSTEMTIME
     wYear As Integer
     wMonth As Integer
     wDayOfWeek As Integer
     wDay As Integer
     wHour As Integer
     wMinute As Integer
     wSecond As Integer
     wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
    bias As Long ' current offset to GMT
    StandardName(1 To 64) As Byte ' unicode string
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(1 To 64) As Byte
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type

Public Enum TimeZoneReturn
    TimeZoneCode = 0
    TimeZoneName = 1
    UTC_BaseOffset = 2
    UTC_Offset = 3
    DST_Active = 4
    DST_Offset = 5
End Enum

' ----------------- For Time Zone Retrieval ------------------
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Private Declare Sub GetSystemTime Lib "kernel32" _
    (lpSystemTime As SYSTEMTIME)

Private Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

Public Const MAGICNUMBER_YES& = &HEFCDAB89
Public Const MAGICNUMBER_NO& = &H98BADCFE

Private Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
Private Const KEY_CONTAINER As String = "ActiveLock"
Private Const PROV_RSA_FULL As Long = 1

Private fInit As Boolean ' flag to indicate that module initialization has been done

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Declare Function MapFileAndCheckSum Lib "imagehlp" Alias "MapFileAndCheckSumA" (ByVal FileName As String, HeaderSum As Long, CheckSum As Long) As Long

' To Report API errors:
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF

Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Declare Function GeneralWinDirApi Lib "kernel32" _
        Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, _
        ByVal nSize As Long) As Long

Public Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Integer
'===============================================================================
'   MakeWord - Packs 2 8-bit integers into a 16-bit integer.
'===============================================================================

    If (HiByte And &H80) <> 0 Then
        MakeWord = ((HiByte * 256&) + LoByte) Or &HFFFF0000
    Else
        MakeWord = (HiByte * 256) + LoByte
    End If
    
End Function

Public Function HiByte(ByVal w As Integer) As Byte
    HiByte = (w And &HFF00&) \ 256
End Function
Public Function LoByte(ByVal w As Integer) As Byte
    LoByte = w And &HFF
End Function

'===============================================================================
' Name: Function TrimNulls
' Input:
'   ByRef startstr As String - String to be trimmed
' Output:
'   String - Trimmed string
' Purpose: Trims Null characters from the string.
' Remarks: None
'===============================================================================
Public Function TrimNulls(startstr As String) As String
    Dim Pos As Integer
    Pos = InStr(startstr, Chr$(0))
    If Pos Then
        TrimNulls = Trim(Left$(startstr, Pos - 1))
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
Public Function ReadFile(ByVal sPath As String, ByRef sData As String) As Long
    Dim hFile As Integer
    ' obtain next free file handle
    hFile = FreeFile
    ' read file content
    Open sPath For Binary Access Read As #hFile
    On Error GoTo Hell
    Debug.Print "File len: " & LOF(hFile)
    ' allocate enough memory to hold the data
    sData = String(LOF(hFile), 0)
    ' read from file
    Get hFile, , sData
    Debug.Print "Bytes read: " & Len(sData)
    Close #hFile
    ReadFile = Len(sData)
    Exit Function
Hell:
    Close #hFile
    Err.Raise Err.Number, "Activelock3", Err.Description
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
Public Sub CryptoProgressUpdate(ByVal param As Long, ByVal action As Long, ByVal phase As Long, ByVal iprogress As Long)
    Debug.Print "Progress Update received " & param & ", action: " & action & ", iprogress: " & iprogress
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
' Name: Function MD5HashTypeLib
' Input:
'   ByRef obj As IUnknown - COM object used to determine the file path to the type library
' Output:
'   String - MD5 Hash Type Library File Path
' Purpose: Computes an MD5 hash of the type library containing the object.
' Remarks: None
'===============================================================================
Public Function MD5HashTypeLib(obj As IUnknown) As String
    Dim strDllPath As String
    strDllPath = GetTypeLibPathFromObject(obj)
    MD5HashTypeLib = MD5HashFile(strDllPath)
End Function


'===============================================================================
' Name: Function GetTypeLibPathFromObject
' Input:
'   ByRef obj As IUnknown - Name of the object for which the class info is needed
' Output:
'   String - Type library path for the given object
' Purpose: Retrieves TypeLib info using TLI library (tlbinfo.dll)
' Remarks: Uses late-binding so that the user doesn't have to add it to their project reference
'===============================================================================
Private Function GetTypeLibPathFromObject(obj As IUnknown) As String
    ' Retrieve TypeLib info using TLI library (tlbinfo.dll)
    ' Use late-binding so that the user doesn't have to add it to their project reference
    Dim tliApp As Object
    Set tliApp = CreateObject("TLI.TLIApplication")
    Dim ti As Object ' actually TLI.TypeInfo
    Set ti = tliApp.ClassInfoFromObject(obj)
    Dim strDllPath As String
    GetTypeLibPathFromObject = ti.Parent.ContainingFile
End Function


'===============================================================================
' Name: Function CRCCheckSumTypeLib
' Input:
'   ByRef obj As IUnknown - COM object used to determine the file path to the type library
' Output:
'   Long - CRC value of the DLL
' Purpose: Performs CRC checksum on the type library containing the object.
' Remarks: None
'===============================================================================
Public Function CRCCheckSumTypeLib(obj As IUnknown) As Long
    Dim strDllPath As String
    strDllPath = GetTypeLibPathFromObject(obj)
    Dim HeaderSum As Long, RealSum As Long
    MapFileAndCheckSum strDllPath, HeaderSum, RealSum
    CRCCheckSumTypeLib = RealSum
End Function



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
Debug.Print "Hashing file " & strPath
Debug.Print "File Date: " & FileDateTime(strPath)
    ' read and hash the content
    Dim sData As String, nFileLen
    nFileLen = ReadFile(strPath, sData)
    Dim sHash As String: sHash = String(32, 0)
    ' hash it
    md5_hash sData, nFileLen, sHash
    MD5HashFile = sHash
End Function


'===============================================================================
' Name: Function IsRunningInIDE
' Input: None
' Output:
'   Boolean - True if running in the IDE
' Purpose: Check if we&#39;re running inside the VB6 IDE
' Remarks: None
'===============================================================================
Public Function IsRunningInIDE() As Boolean
    Dim strFilename As String
    Dim lngCount As Long
    
    strFilename = String(255, 0)
    lngCount = GetModuleFileName(App.hInstance, strFilename, 255)
    strFilename = Left(strFilename, lngCount)
    IsRunningInIDE = UCase(Right(strFilename, 7)) = "VB6.EXE"
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
    If Not Dir(strFile) = "" Then
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
Public Function LocalTimeZone(ByVal returnType As TimeZoneReturn) As Variant
    Dim x As Long
    Dim tzi As TIME_ZONE_INFORMATION
    Dim strName As String
    Dim bDST As Boolean
    Dim rc&
    rc = GetTimeZoneInformation(tzi)
    Select Case rc
        ' if not daylight assume standard
        Case TIME_ZONE_ID_DAYLIGHT
            strName = tzi.DaylightName ' convert to string
            bDST = True
        Case Else
            strName = tzi.StandardName
    End Select
    
    ' name terminates with null
    x = InStr(strName, vbNullChar)
    If x > 0 Then strName = Left$(strName, x - 1)
            
    If returnType = DST_Active Then
        LocalTimeZone = bDST
    End If
    
    If returnType = TimeZoneName Then
        LocalTimeZone = strName
    End If
    
    If returnType = TimeZoneCode Then
        LocalTimeZone = Left(strName, 1)
        x = InStr(1, strName, " ")
        Do While x > 0
            LocalTimeZone = LocalTimeZone & Mid(strName, x + 1, 1)
            x = InStr(x + 1, strName, " ")
        Loop
        LocalTimeZone = Trim(LocalTimeZone)
    End If
            
    If returnType = UTC_BaseOffset Then
        LocalTimeZone = tzi.bias
    End If
        
    If returnType = DST_Offset Then
        LocalTimeZone = tzi.DaylightBias
    End If
    
    If returnType = UTC_Offset Then
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
' Name: Function UTC
' Input:
'   ByRef dt As Date - Date-Time input to be converted into UTC Date-Time
' Output:
'   Date - UTC Date-Time
' Purpose: Converts a local date-time into UTC/GMT date-time
' Remarks: None
'===============================================================================
Public Function UTC(dt As Date) As Date
    '  Returns current UTC date-time.
    UTC = DateAdd("n", LocalTimeZone(UTC_Offset), dt)
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
    Dim Key As RSAKey
    ' create the key from the key blobs
    If rsa_createkey(strPub, Len(strPub), strPriv, Len(strPriv), Key) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If

    ' sign the data using the created key
    Dim sLen&
    If rsa_sign(Key, strdata, Len(strdata), vbNullString, sLen) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
    Dim strSig As String: strSig = String(sLen, 0)
    If rsa_sign(Key, strdata, Len(strdata), strSig, sLen) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
    ' throw away the key
    If rsa_freekey(Key) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
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
Public Function RSAVerify(ByVal strPub As String, ByVal strdata As String, ByVal strSig As String) As Long
    Dim Key As RSAKey
    Dim rc As Long
    ' create the key from the public key blob
    If rsa_createkey(strPub, Len(strPub), vbNullString, 0, Key) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
    ' validate the key
    rc = rsa_verifysig(Key, strSig, Len(strSig), strdata, Len(strdata))
    If rc = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
    ' de-allocate memory used by the key
    If rsa_freekey(Key) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
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
Public Function WinError(ByVal lLastDLLError As Long) As String
    Dim sBuff As String
    Dim lCount As Long

    ' Return the error message associated with LastDLLError:
    sBuff = String$(256, 0)
    lCount = FormatMessage( _
        FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
        0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
      WinError = Left$(sBuff, lCount)
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
    Const FIX_LENGTH% = 4096
    Dim Length As Integer
    Dim Buffer As String * FIX_LENGTH

    Length = GeneralWinDirApi(Buffer, FIX_LENGTH - 1)
    WinDir = Left$(Buffer, Length)
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
Dim sResult As String

On Error Resume Next
sResult = Dir(sFolder, vbDirectory)

On Error GoTo 0
FolderExists = sResult <> ""
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



