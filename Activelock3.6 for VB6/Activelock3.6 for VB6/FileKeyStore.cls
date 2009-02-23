VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileKeyStoreProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
'===============================================================================
' Name: FileKeyStoreProvider
' Purpose: This IKeyStoreProvider implementation is used to  maintain the license keys on a file system.
' Functions:
' Properties:
' Methods:
' Started: 21.04.2005
' Modified: 08.07.2005
'===============================================================================
' @author activelock-admins
' @version 3.0.0
' @date 20050807
'
'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
'
'   [ ] IKeyStoreProvider_Store() - need to handle file locking for concurrent licensing.
'
'
'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
'
'   07.07.03 - mcrute    - Updated the header comments for this file.
'   11.02.03 - th2tran   - Handle loading and saving of new MaxCount property
'   04.17.04 - th2tran   - Don't save the product key to the license file.
'                        - Now creates an empty license file if it doesn't exist,
'                          instead of raising alerrKeyStoreInvalid error.

'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////

Option Explicit


' Implements IKeyStoreProvider interface.
Implements IKeyStoreProvider

Private mstrPath As String
Private mINIFile As New INIFile

' License File Key names
Private Const KEY_PRODKEY$ = "ProductKey"
Private Const KEY_PRODNAME$ = "ProductName"
Private Const KEY_PRODVER$ = "ProductVersion"
Private Const KEY_LICENSEE$ = "Licensee"
Private Const KEY_REGISTERED_LEVEL$ = "RegisteredLevel"
Private Const KEY_MAXCOUNT$ = "MaxCount" ' Maximum number of users
Private Const KEY_LICTYPE$ = "LicenseType"
Private Const KEY_LICCLASS$ = "LicenseClass"
Private Const KEY_LICKEY$ = "LicenseKey"
Private Const KEY_LICCODE$ = "LicenseCode" ' New in v3.1
Private Const KEY_EXP$ = "Expiration"
Private Const KEY_REGISTERED_DATE$ = "RegisteredDate"
Private Const KEY_LASTRUN_DATE$ = "LastUsed" ' date and time stamp
Private Const KEY_LASTRUN_DATE_HASH$ = "Hash1" ' Hash of LastRunDate

'===============================================================================
' Name: Property Let IKeyStoreProvider_KeyStorePath
' Input:
'   ByRef RHS As String - File path and name
' Output: None
' Purpose: Creates an empty file if it doesn't exist
' Remarks: None
'===============================================================================
Private Property Let IKeyStoreProvider_KeyStorePath(RHS As String)
    If Not FileExists(RHS) Then
        ' Create an empty file if it doesn't exists
        CreateEmptyFile RHS
    Else 'the file exists, but check to see if it has read-only attribute
        If (GetAttr(RHS) And vbReadOnly) Or (GetAttr(RHS) And vbReadOnly And vbArchive) Then
            Set_locale regionalSymbol
            Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
        End If
    End If
    mstrPath = RHS
    mINIFile.File = mstrPath
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
    Dim hFile As Long
    hFile = FreeFile
    Open sFilePath For Output As #hFile
    Close #hFile
End Sub

'===============================================================================
' Name: Sub IKeyStoreProvider_Store
' Input:
'   ByRef Lic As ProductLicense - Product license object
' Output: None
' Purpose: Write license properties to INI file section
' Remarks: TODO: Perhaps we need to lock the file first.?
'===============================================================================
Private Sub IKeyStoreProvider_Store(Lic As ProductLicense, mLicenseFileType As ALLicenseFileTypes)
    Dim arrProdVer() As String
    Dim actualLicensee As String
    arrProdVer = Split(Lic.Licensee, "&&&")
    actualLicensee = arrProdVer(0)
    ' Write license properties to INI file section
    mINIFile.Section = Lic.ProductName
    
    If mLicenseFileType = ALLicenseFileTypes.alsLicenseFileEncrypted Then
        With Lic
            mINIFile.Values(KEY_PRODVER) = EncryptMyString(.ProductVer, PSWD)
            mINIFile.Values(KEY_LICTYPE) = EncryptMyString(.LicenseType, PSWD)
            mINIFile.Values(KEY_LICCLASS) = EncryptMyString(.LicenseClass, PSWD)
            mINIFile.Values(KEY_LICENSEE) = EncryptMyString(actualLicensee, PSWD)
            mINIFile.Values(KEY_REGISTERED_LEVEL) = EncryptMyString(.RegisteredLevel, PSWD)
            mINIFile.Values(KEY_MAXCOUNT) = EncryptMyString(CStr(.MaxCount), PSWD)
            mINIFile.Values(KEY_LICKEY) = EncryptMyString(.LicenseKey, PSWD)
            mINIFile.Values(KEY_REGISTERED_DATE) = EncryptMyString(.RegisteredDate, PSWD)
            mINIFile.Values(KEY_LASTRUN_DATE) = EncryptMyString(.LastUsed, PSWD)
            mINIFile.Values(KEY_LASTRUN_DATE_HASH) = EncryptMyString(.Hash1, PSWD)
            mINIFile.Values(KEY_EXP) = EncryptMyString(.Expiration, PSWD)
            mINIFile.Values(KEY_LICCODE) = EncryptMyString(.LicenseCode, PSWD)
        End With
    Else
        With Lic
            mINIFile.Values(KEY_PRODVER) = .ProductVer
            mINIFile.Values(KEY_LICTYPE) = .LicenseType
            mINIFile.Values(KEY_LICCLASS) = .LicenseClass
            mINIFile.Values(KEY_LICENSEE) = actualLicensee
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
Private Function IKeyStoreProvider_Retrieve(ProductName As String, mLicenseFileType As ALLicenseFileTypes) As ProductLicense

    Set IKeyStoreProvider_Retrieve = Nothing
    
    mINIFile.Section = ProductName

    On Error GoTo InvalidValue
    ' No license found
    If mINIFile.GetValue(KEY_LICKEY) = "" Then Exit Function

    ' Read the LIC file quickly and decide whether this is a valid LIC file
    Dim fileText As String
    fileText = FileLoad(mINIFile.File)
    Dim astring As String
    If fileText = "" Then Exit Function
    If mLicenseFileType = ALLicenseFileTypes.alsLicenseFileEncrypted Then
        astring = DecryptMyString(fileText, PSWD)
    Else
        astring = fileText
    End If
    If inString(LCase(astring), "productversion") = False Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrKeyStoreInvalid, ACTIVELOCKSTRING, STRKEYSTOREINVALID
    End If

    Dim Lic As New ProductLicense
    If mLicenseFileType = ALLicenseFileTypes.alsLicenseFileEncrypted And _
        mINIFile.GetValue(KEY_LICCLASS) <> "Single" And mINIFile.GetValue(KEY_LICCLASS) <> "MultiUser" Then
        ' Read license properties from INI file section
        With Lic
            .ProductName = ProductName
            .ProductVer = DecryptMyString(mINIFile.GetValue(KEY_PRODVER), PSWD)
            .Licensee = DecryptMyString(mINIFile.GetValue(KEY_LICENSEE), PSWD) & "&&&" & .ProductName & " - " & .ProductVer
            .RegisteredLevel = DecryptMyString(mINIFile.GetValue(KEY_REGISTERED_LEVEL), PSWD)
            .MaxCount = CInt(DecryptMyString(mINIFile.Values(KEY_MAXCOUNT), PSWD))
            .LicenseType = DecryptMyString(mINIFile.GetValue(KEY_LICTYPE), PSWD)
            .LicenseClass = DecryptMyString(mINIFile.GetValue(KEY_LICCLASS), PSWD)
            .LicenseKey = DecryptMyString(mINIFile.GetValue(KEY_LICKEY), PSWD)
            .Expiration = DecryptMyString(mINIFile.GetValue(KEY_EXP), PSWD)
            .RegisteredDate = DecryptMyString(mINIFile.Values(KEY_REGISTERED_DATE), PSWD)
            .LastUsed = DecryptMyString(mINIFile.Values(KEY_LASTRUN_DATE), PSWD)
            .Hash1 = DecryptMyString(mINIFile.Values(KEY_LASTRUN_DATE_HASH), PSWD)
            .LicenseCode = DecryptMyString(mINIFile.Values(KEY_LICCODE), PSWD)
        End With
    Else ' INI file (LIC) is in PLAIN format

        ' Read license properties from INI file section
        With Lic
            .ProductName = ProductName
            .ProductVer = mINIFile.GetValue(KEY_PRODVER)
            .Licensee = mINIFile.GetValue(KEY_LICENSEE) & "&&&" & .ProductName & " - " & .ProductVer
            .RegisteredLevel = mINIFile.GetValue(KEY_REGISTERED_LEVEL)
            .MaxCount = CInt(mINIFile.Values(KEY_MAXCOUNT))
            .LicenseType = mINIFile.GetValue(KEY_LICTYPE)
            .LicenseClass = mINIFile.GetValue(KEY_LICCLASS)
            .LicenseKey = mINIFile.GetValue(KEY_LICKEY)
            .Expiration = mINIFile.GetValue(KEY_EXP)
            .RegisteredDate = mINIFile.Values(KEY_REGISTERED_DATE)
            .LastUsed = mINIFile.Values(KEY_LASTRUN_DATE)
            .Hash1 = mINIFile.Values(KEY_LASTRUN_DATE_HASH)
            .LicenseCode = mINIFile.Values(KEY_LICCODE)
        End With
    End If
    Set IKeyStoreProvider_Retrieve = Lic
    
    Exit Function

InvalidValue:
    Set_locale regionalSymbol
    Err.Raise ActiveLockErrCodeConstants.alerrKeyStoreInvalid, ACTIVELOCKSTRING, STRKEYSTOREINVALID
End Function
