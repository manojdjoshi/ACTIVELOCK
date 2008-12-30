Option Strict Off
Option Explicit On 
Imports System.IO
Friend Class FileKeyStoreProvider
	Implements _IKeyStoreProvider
	'*   ActiveLock
	'*   Copyright 1998-2002 Nelson Ferraz
    '*   Copyright 2006 The ActiveLock Software Group (ASG)
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
	' Name: FileKeyStoreProvider
	' Purpose: This IKeyStoreProvider implementation is used to  maintain the license keys on a file system.
	' Functions:
	' Properties:
	' Methods:
	' Started: 21.04.2005
    ' Modified: 03.24.2006
	'===============================================================================
	' @author activelock-admins
    ' @version 3.3.0
    ' @date 03.24.2006
	'

	' Implements IKeyStoreProvider interface.
	
	Private mstrPath As String
	Private mINIFile As New INIFile
	
	' License File Key names
	Private Const KEY_PRODKEY As String = "ProductKey"
	Private Const KEY_PRODNAME As String = "ProductName"
	Private Const KEY_PRODVER As String = "ProductVersion"
	Private Const KEY_LICENSEE As String = "Licensee"
	Private Const KEY_REGISTERED_LEVEL As String = "RegisteredLevel"
	Private Const KEY_MAXCOUNT As String = "MaxCount" ' Maximum number of users
	Private Const KEY_LICTYPE As String = "LicenseType"
	Private Const KEY_LICCLASS As String = "LicenseClass"
	Private Const KEY_LICKEY As String = "LicenseKey"
	Private Const KEY_LICCODE As String = "LicenseCode" ' New in v3.1
	Private Const KEY_EXP As String = "Expiration"
	Private Const KEY_REGISTERED_DATE As String = "RegisteredDate"
	Private Const KEY_LASTRUN_DATE As String = "LastUsed" ' date and time stamp
	Private Const KEY_LASTRUN_DATE_HASH As String = "Hash1" ' Hash of LastRunDate
    '===============================================================================
	' Name: Property Let IKeyStoreProvider_KeyStorePath
	' Input:
	'   ByRef RHS As String - File path and name
	' Output: None
	' Purpose: Creates an empty file if it doesn't exist
	' Remarks: None
	'===============================================================================
	Private WriteOnly Property IKeyStoreProvider_KeyStorePath() As String Implements _IKeyStoreProvider.KeyStorePath
		Set(ByVal Value As String)
			If Not FileExists(Value) Then
				' Create an empty file if it doesn't exists
				CreateEmptyFile(Value)
			Else 'the file exists, but check to see if it has read-only attribute
				If (GetAttr(Value) And FileAttribute.ReadOnly) Or (GetAttr(Value) And FileAttribute.ReadOnly And FileAttribute.Archive) Then
                    'Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                End If
            End If
            mstrPath = Value
            mINIFile.File = mstrPath
        End Set
    End Property
    '===============================================================================
    ' Name: Sub CreateEmptyFile
    ' Input:
    '   ByVal sFilePath As String - File path and name
    ' Output: None
    ' Purpose: Creates an empty file
    ' Remarks: None
    '===============================================================================
    Private Sub CreateEmptyFile(ByVal sFilePath As String)
        Dim hFile As Integer
        hFile = FreeFile
        FileOpen(hFile, sFilePath, OpenMode.Output)
        FileClose(hFile)
    End Sub
    '===============================================================================
    ' Name: Sub IKeyStoreProvider_Store
    ' Input:
    '   ByRef Lic As ProductLicense - Product license object
    ' Output: None
    ' Purpose: Write license properties to INI file section
    ' Remarks: TODO: Perhaps we need to lock the file first.?
    '===============================================================================
    Private Sub IKeyStoreProvider_Store(ByRef Lic As ProductLicense, ByVal mLicenseFileType As IActiveLock.ALLicenseFileTypes) Implements _IKeyStoreProvider.Store
        ' Write license properties to INI file section
        mINIFile.Section = Lic.ProductName

        If mLicenseFileType = IActiveLock.ALLicenseFileTypes.alsLicenseFileEncrypted Then
            With Lic
                mINIFile.Values(KEY_PRODVER) = EncryptString128Bit(.ProductVer, PSWD)
                mINIFile.Values(KEY_LICTYPE) = EncryptString128Bit(.LicenseType, PSWD)
                mINIFile.Values(KEY_LICCLASS) = EncryptString128Bit(.LicenseClass, PSWD)
                mINIFile.Values(KEY_LICENSEE) = EncryptString128Bit(.Licensee, PSWD)
                mINIFile.Values(KEY_REGISTERED_LEVEL) = EncryptString128Bit(.RegisteredLevel, PSWD)
                mINIFile.Values(KEY_MAXCOUNT) = EncryptString128Bit(CStr(.MaxCount), PSWD)
                mINIFile.Values(KEY_LICKEY) = EncryptString128Bit(.LicenseKey, PSWD)
                mINIFile.Values(KEY_REGISTERED_DATE) = EncryptString128Bit(.RegisteredDate, PSWD)
                mINIFile.Values(KEY_LASTRUN_DATE) = EncryptString128Bit(.LastUsed, PSWD)
                mINIFile.Values(KEY_LASTRUN_DATE_HASH) = EncryptString128Bit(.Hash1, PSWD)
                mINIFile.Values(KEY_EXP) = EncryptString128Bit(.Expiration, PSWD)
                mINIFile.Values(KEY_LICCODE) = EncryptString128Bit(.LicenseCode, PSWD)
            End With
        Else
            With Lic
                mINIFile.Values(KEY_PRODVER) = .ProductVer
                mINIFile.Values(KEY_LICTYPE) = .LicenseType
                mINIFile.Values(KEY_LICCLASS) = .LicenseClass
                mINIFile.Values(KEY_LICENSEE) = .Licensee
                mINIFile.Values(KEY_REGISTERED_LEVEL) = .RegisteredLevel
                mINIFile.Values(KEY_MAXCOUNT) = CStr(.MaxCount)
                mINIFile.Values(KEY_LICKEY) = .LicenseKey
                mINIFile.Values(KEY_REGISTERED_DATE) = .RegisteredDate
                mINIFile.Values(KEY_LASTRUN_DATE) = .LastUsed
                mINIFile.Values(KEY_LASTRUN_DATE_HASH) = .Hash1
                mINIFile.Values(KEY_EXP) = .Expiration
                mINIFile.Values(KEY_LICCODE) = .LicenseCode
            End With

        End If

        ' This was the original idea... to read the file, and encrypt and write again
        ' But since this is a stream based operation, it conflicts with the
        ' ADS therefore should NOT be used
        'If mLicenseFileType = IActiveLock.ALLicenseFileTypes.alsLicenseFileEncrypted Then
        '    ' Read the LIC file again
        '    Dim stream_reader2 As New IO.StreamReader(mINIFile.File)
        '    Dim fileContents2 As String = stream_reader2.ReadToEnd()
        '    stream_reader2.Close()
        '    ' Encrypt the file and save it
        '    Dim encryptedFile2 As String
        '    encryptedFile2 = EncryptString128Bit(fileContents2, PSWD)
        '    Dim stream_writer2 As New IO.StreamWriter(mINIFile.File, False)
        '    stream_writer2.Write(encryptedFile2)
        '    stream_writer2.Close()
        'End If

    End Sub
    '===============================================================================
    ' Name: Function IKeyStoreProvider_Retrieve
    ' Input:
    '   ByRef ProductName As String - Product or application name
    ' Output:
    '   ProductLicense - Returns the product license object
    ' Purpose: Retrieves the registered license for the specified product.
    ' Remarks: None
    '===============================================================================
    Private Function IKeyStoreProvider_Retrieve(ByRef ProductName As String, ByVal mLicenseFileType As IActiveLock.ALLicenseFileTypes) As ProductLicense Implements _IKeyStoreProvider.Retrieve

        IKeyStoreProvider_Retrieve = Nothing

        mINIFile.Section = ProductName

        On Error GoTo InvalidValue
        ' No license found
        If mINIFile.GetValue(KEY_LICKEY) = "" Then Exit Function

        Dim Lic As New ProductLicense
        If mLicenseFileType = IActiveLock.ALLicenseFileTypes.alsLicenseFileEncrypted And mINIFile.GetValue(KEY_LICCLASS) <> "Single" And mINIFile.GetValue(KEY_LICCLASS) <> "MultiUser" Then

            ' This was the original idea... to read the file, and encrypt and write again
            ' But since this is a stream based operation, it conflicts with the
            ' ADS therefore should NOT be used
            '' Read the LIC file
            'Dim stream_reader1 As New IO.StreamReader(mINIFile.File)
            'Dim fileContents1 As String = stream_reader1.ReadToEnd()
            'stream_reader1.Close()
            '' Decrypt the LIC file
            'Dim decryptedFile1 As String
            'decryptedFile1 = DecryptString128Bit(fileContents1, PSWD)
            '' Save the LIC file again
            'Dim stream_writer1 As New IO.StreamWriter(mINIFile.File, False)
            'stream_writer1.Write(decryptedFile1)
            'stream_writer1.Close()

            ' Read license properties from INI file section
            With Lic
                .ProductName = ProductName
                .ProductVer = DecryptString128Bit(mINIFile.GetValue(KEY_PRODVER), PSWD)
                .Licensee = DecryptString128Bit(mINIFile.GetValue(KEY_LICENSEE), PSWD)
                .RegisteredLevel = DecryptString128Bit(mINIFile.GetValue(KEY_REGISTERED_LEVEL), PSWD)
                .MaxCount = CType(DecryptString128Bit(mINIFile.Values(KEY_MAXCOUNT), PSWD), Integer)
                .LicenseType = DecryptString128Bit(mINIFile.GetValue(KEY_LICTYPE), PSWD)
                .LicenseClass = DecryptString128Bit(mINIFile.GetValue(KEY_LICCLASS), PSWD)
                .LicenseKey = DecryptString128Bit(mINIFile.GetValue(KEY_LICKEY), PSWD)
                .Expiration = DecryptString128Bit(mINIFile.GetValue(KEY_EXP), PSWD)
                .RegisteredDate = DecryptString128Bit(mINIFile.Values(KEY_REGISTERED_DATE), PSWD)
                .LastUsed = DecryptString128Bit(mINIFile.Values(KEY_LASTRUN_DATE), PSWD)
                .Hash1 = DecryptString128Bit(mINIFile.Values(KEY_LASTRUN_DATE_HASH), PSWD)
                .LicenseCode = DecryptString128Bit(mINIFile.Values(KEY_LICCODE), PSWD)
            End With
        Else ' INI file (LIC) is in PLAIN format

            ' Read license properties from INI file section
            With Lic
                .ProductName = ProductName
                .ProductVer = mINIFile.GetValue(KEY_PRODVER)
                .Licensee = mINIFile.GetValue(KEY_LICENSEE)
                .RegisteredLevel = mINIFile.GetValue(KEY_REGISTERED_LEVEL)
                .MaxCount = CInt(mINIFile.Values(KEY_MAXCOUNT))
                .LicenseType = mINIFile.GetValue(KEY_LICTYPE)
                .LicenseClass = mINIFile.GetValue(KEY_LICCLASS)
                .LicenseKey = mINIFile.GetValue(KEY_LICKEY)
                .Expiration = mINIFile.GetValue(KEY_EXP)
                .RegisteredDate = mINIFile.Values(KEY_REGISTERED_DATE)
                .LastUsed = mINIFile.Values(KEY_LASTRUN_DATE)
                .Hash1 = mINIFile.Values(KEY_LASTRUN_DATE_HASH)
                .LicenseCode = mINIFile.Values(KEY_LICCODE) ' New in v3.1
            End With
        End If
        IKeyStoreProvider_Retrieve = Lic

        Exit Function
InvalidValue:
        'Set_locale(regionalSymbol)
        Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrKeyStoreInvalid, ACTIVELOCKSTRING, STRKEYSTOREINVALID)
    End Function
End Class