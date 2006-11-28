VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Globals"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Name: Globals
' Purpose: This class contains global object factory and utility methods and constants.
' <p>It is a global class so its routines in here can be accessed directly
' from the ActiveLock3 namespace.
' For example, the <code>NewInstance()</code> function can be accessed via
' <code>ActiveLock3.NewInstance()</code>.
' Functions:
' Properties:
' Methods:
' Started: 04.21.2005
' Modified: 08.08.2005
'===============================================================================
' @author activelock-admins
' @version 3.0.0
' @date 20050808
'
'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
' @todo GetLicTypeString().  Currently only supports &#34;Single&#34;.  Need to implement the rest.
'

'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' @history
' <pre>
'   07.07.03 - mcrute        - Updated the header comments for this file.
'   07.21.03 - th2tran       - Added ActiveLockErrCodeConstants to contain
'                              a list of error code constants raised by ActiveLock.
'   07.28.03 - wizzardme2000 - Implemented lockComp, lockHD, and lockWindows
'   07.31.03 - th2tran       - Added alerrFileTampered
'                            - CreateProductLicense() to ignore Expiration date
'                              for Permanent license type
'   08.03.03 - th2tran       - VBDox&#39;ed this class.
'   10.13.03 - th2tran       - Corrections to ActiveLockErrCodeConstants vbdox errors.
'                            - Added VB descriptions for VB Object Browser's sake.
'   11.02.03 - th2tran       - Added ActiveLockErrCodeConstants.alerrNotImplemented
'   04.17.04 - th2tran       - NewInstance() no longer calls Init().  Callers will have to
'                              init the instance themselves.
' </pre>
'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit


' ActiveLock Error Codes.
' These error codes are used for <code>Err.Number</code> whenever ActiveLock raises an error.
'
' @param alerrOK                    No error. Operation was successful.
' @param alerrNoLicense             No license available.
' @param alerrLicenseInvalid        License is invalid.
' @param alerrLicenseExpired        License has expired.
' @param alerrLicenseTampered       License has been tampered.
' @param alerrClockChanged          System clock has been changed.
' @param alerrWrongIPaddress        Wrong IP Address.
' @param alerrKeyStoreInvalid       Key Store Provider has not been initialized yet.
' @param alerrKeyStorePathInvalid   Key Store Path (LIC file path) hasn't been specified.
' @param alerrFileTampered          ActiveLock DLL file has been tampered.
' @param alerrNotInitialized        ActiveLock has not been initialized yet.
' @param alerrNotImplemented        An ActiveLock operation has not been implemented.
' @param alerrUserNameTooLong       Maximum User Name length of 2000 characters has been exceeded.
' @param alerrUserNameInvalid       Used User name does not match with the license key.
' @param alerrInvalidTrialDays      Specified number of Free Trial Days is invalid (possibly <=0).
' @param alerrInvalidTrialRuns      Specified number of Free Trial Runs is invalid (possibly <=0).
' @param alerrTrialInvalid          Trial is invalid.
' @param alerrTrialDaysExpired      Trial Days have expired.
' @param alerrTrialRunsExpired      Trial Runs have expired.
' @param alerrNoSoftwareName        Software Name has not been specified.
' @param alerrNoSoftwareVersion     Software Version has not been specified.
' @param alerrRSAError              Something went wrong in the RSA routines.

' @param alerrCryptoAPIError        Crypto API error in CryptoAPI class.

Public Enum ActiveLockErrCodeConstants
    alerrOK = 0   ' successful
    alerrNoLicense = &H80040001         ' vbObjectError (&H80040000) + 1
    alerrLicenseInvalid = &H80040002
    alerrLicenseExpired = &H80040003
    alerrLicenseTampered = &H80040004
    alerrClockChanged = &H80040005
    alerrWrongIPaddress = &H80040006
    
    alerrKeyStoreInvalid = &H80040010
    alerrFileTampered = &H80040011
    alerrNotInitialized = &H80040012
    alerrNotImplemented = &H80040013
    alerrUserNameTooLong = &H80040014
    alerrUserNameInvalid = &H80040015
    
    alerrInvalidTrialDays = &H80040020
    alerrInvalidTrialRuns = &H80040021
    alerrTrialInvalid = &H80040022
    alerrTrialDaysExpired = &H80040023
    alerrTrialRunsExpired = &H80040024
    alerrNoSoftwareName = &H80040025
    alerrNoSoftwareVersion = &H80040026
    alerrRSAError = &H80040027
    alerrKeyStorePathInvalid = &H80040028
    alerrCryptoAPIError = &H80040029
    alerrNoSoftwarePassword = &H80040030
    
End Enum

' USED BY CRYPTOAPI
Private cryptSession As clsCryptoAPI
Private strCypherText As String
Private bCypherOn As Boolean

Public Function ContainerChange(ByVal containerName As String)
Dim sKeyDongle As String, sKeyDonglePrivate As String

' At one point check the characters in the container name and do the following:
' Not acceptable ---->    Case 8 To 10, 32, 44, 160, 163, 169, 174
' Use zero  --------->    Case 0 To 47, 58 To 63, 91 To 96, 123 To 255

strCypherText = vbNullString

If containerName <> vbNullString Then
    Set cryptSession = New clsCryptoAPI
    If cryptSession.VerifyContainer(containerName, False) Then
        
        Dim publicKey As String, privateKey As String
        'privateKey = "UzMwODoHAgAAACQAAPh4QW25OhuOEKDIksffuUuJYQTsMeu/lKoemvJz8Oj56ShbVvou69rbAybVHXbhKnhrDnS9wjuIqOt9MWNmzld4G2tr33ll/8aQ9MaOZT/0G7hDmOVFPtl3L/3OtbGR9nj5oSoGlNWTrsgYNa9fV2M5kD4tVnheVqa55pf23K2iN40nufeAd2Oqnk7p8JHx58nITbCkQ1Q1W1zGm4z6jIktn0sPdgGuxmI5Rnx0onBtVp9XMIkwJgyzWvWUqrUKG4v2bu98Vmms3+3o9WOJDNYbZB1MeaGfBA+B6lUonRATh7jfefSh/8WCKxnu7JJPtp23rM64CpclIhjq2s+YAYhUPe9J/u+lx0hE61psS9rlAhBFu8nqOPVLiPuJSu+p3K5kjF59O5u1Sp9nxA"
        'publicKey = "Uzg3fDg0OgYCAAAAJAAAUlNBMQACAAABAAEAp6+IziBn6pKjSMh+PBAvUUtiSbUqJPUpcBJWdSkgeYwfJo2FBE0PuuIqtS3vFUK9PLnGsQ1k7/0F/RnaGzhC6A"
        '             Uzg3fDg0OgYCAAAAJAAAUlNBMQACAAABAAEAq9jBcc/NfSSpBcNdTcXUnp/aEm/wLr5jIJ1u/GqzFCQi6uFNCr90OvC3KlR2hGK3S/4s9UsjbEZ9lQTpX4vaww
        'privateKey = Base64_Decode(privateKey)
        'publicKey = Base64_Decode(publicKey)
        'privateBlob = cryptSession.StringToByteArray(privateKey, True, False)
        'publicBlob = cryptSession.StringToByteArray(publicKey, True, False)
        'cryptSession.ValuePrivateKey = privateKey
        'cryptSession.ValuePublicKey = publicKey
        'privateKey = cryptSession.ByteArrayToString(privateBlob)
        'publicKey = cryptSession.ByteArrayToString(publicBlob)
        
        'cryptSession.ImportPublicKey False
        'cryptSession.ImportPrivateKey True, False
        
        cryptSession.Generate False
        cryptSession.ExportPublicKey False
        cryptSession.ExportPrivateKey True, False
        'sKeyDongle = cryptSession.ValuePublicKey
        'sKeyDonglePrivate = cryptSession.ValuePrivateKey
        
        'publicKey = Base64_Encode(sKeyDongle)
        'privateKey = Base64_Encode(sKeyDonglePrivate)
        
        cryptSession.SessionEnd
    End If
    Set cryptSession = Nothing
End If

End Function

Public Function CryptoAPIAction(Index As Integer, containerName As String, plainText As String, signedText As String, publicKey As String, privateKey As String, modulus As Long)
' This sub generates Public/Private Exchange or Signature keys, Signs/Validates Exchange or Signature keys,
' Creates containers if they don't exist and destroys containers

Set cryptSession = New clsCryptoAPI
Dim sKeyDongle As String, lX As Long

' Process the encrypted or signed string such that:
' Not acceptable ----->    Case 3, 8 To 10, 13, 22, 24, 32, 44, 160, 163, 169, 174
' Use zero  ---------->    Case 0 To 31, 127 To 255

Select Case Index
Case 0   'create container
    Set cryptSession = New clsCryptoAPI
    If cryptSession.SessionStart(containerName, True) Then
        cryptSession.Generate False 'Signature Key
        publicPrivateKeyBlob = privateKey
        cryptSession.ImportPrivateKey True, False
        cryptSession.ExportPrivateKey True, False
        publicPublicKeyBlob = publicKey
        cryptSession.ImportPublicKey False
        cryptSession.ExportPublicKey False
        cryptSession.SessionEnd
    Else
        Set cryptSession = Nothing
        Err.Raise ActiveLockErrCodeConstants.alerrCryptoAPIError, ACTIVELOCKSTRING, "Crypto API Error: This System Unable to Perform Cryptographic Functions"
    End If
    Set cryptSession = Nothing

Case 1   'generate exchange or signature key
    Set cryptSession = New clsCryptoAPI
    If cryptSession.SessionStart(containerName) Then
'        If optStrength(0).Value = True Then
'            modulus = 4096
'        ElseIf optStrength(1).Value = True Then
'            modulus = 2048
'        ElseIf optStrength(2).Value = True Then
'            modulus = 1536
'        ElseIf optStrength(3).Value = True Then
'            modulus = 1024
'        ElseIf optStrength(4).Value = True Then
'            modulus = 512
'        End If
        Screen.MousePointer = vbHourglass
        cryptSession.Generate False, modulus 'Signature Key
        Screen.MousePointer = vbDefault
        cryptSession.ExportPrivateKey True, False
        cryptSession.ExportPublicKey False
        privateKey = "RSA" & modulus & publicPrivateKeyBlob
        publicKey = "RSA" & modulus & publicPublicKeyBlob
        cryptSession.SessionEnd
    Else
        Set cryptSession = Nothing
        Err.Raise ActiveLockErrCodeConstants.alerrCryptoAPIError, ACTIVELOCKSTRING, "Crypto API Error: This System Unable to Perform Cryptographic Functions"
    End If
    Set cryptSession = Nothing

Case 2    'destroy container
    Set cryptSession = New clsCryptoAPI
    If cryptSession.SessionStart(containerName) Then
        cryptSession.DestroyContainer containerName
        cryptSession.SessionEnd
        bCypherOn = False
        plainText = vbNullString
        signedText = vbNullString
        strCypherText = vbNullString
    Else
        Set cryptSession = Nothing
        Err.Raise ActiveLockErrCodeConstants.alerrCryptoAPIError, ACTIVELOCKSTRING, "Crypto API Error: This System Unable to Perform Cryptographic Functions"
    End If
    Set cryptSession = Nothing

Case 3   'encrypt/decrypt
    Dim iChr As Integer
    Set cryptSession = New clsCryptoAPI
    If cryptSession.SessionStart(containerName) Then
        If Not bCypherOn Then
            bCypherOn = True
            strCypherText = cryptSession.EncryptString(plainText)
            plainText = vbNullString
            sKeyDongle = strCypherText
            For lX = 0 To 31
                sKeyDongle = Replace(sKeyDongle, Chr(lX), Chr(lX + 32))
            Next lX
            plainText = sKeyDongle
        Else
            plainText = cryptSession.DecryptString(strCypherText)
            strCypherText = vbNullString
            bCypherOn = False
        End If
        cryptSession.SessionEnd
    Else
        Err.Raise ActiveLockErrCodeConstants.alerrCryptoAPIError, ACTIVELOCKSTRING, "Crypto API Error"
    End If
    Set cryptSession = Nothing

Case 4    'sign the key
    If plainText = "" Then
        Err.Raise ActiveLockErrCodeConstants.alerrCryptoAPIError, ACTIVELOCKSTRING, "Crypto API Error: There's no string to sign"
        Exit Function
    End If
    Set cryptSession = New clsCryptoAPI
    If cryptSession.VerifyContainer(containerName, False) = True Then
        cryptSession.DestroyContainer containerName
        cryptSession.SessionEnd
    End If
    If cryptSession.SessionStart(containerName, True) Then
        cryptSession.Generate False 'Signature Key
        publicPrivateKeyBlob = privateKey
        cryptSession.ImportPrivateKey True, False
        
        Dim inputText As String
        inputText = plainText
        Call cryptSession.SignString(inputText, True)
        signedText = inputText
    End If

Case 5   'validate the key
    Set cryptSession = New clsCryptoAPI
    If cryptSession.VerifyContainer(containerName, False) = True Then
        cryptSession.DestroyContainer containerName
        cryptSession.SessionEnd
    End If
    If cryptSession.SessionStart(containerName, True) Then
        cryptSession.Generate False 'Signature Key
        publicPublicKeyBlob = publicKey
        cryptSession.ImportPublicKey False
        
        inputText = signedText
        If inputText = "" Then
            Err.Raise ActiveLockErrCodeConstants.alerrCryptoAPIError, ACTIVELOCKSTRING, "Crypto API Error: There's no string to validate"
        End If
        If cryptSession.ValidateString(inputText) = True Then
            ' Stay quiet here... everything is allright
            'MsgBox "Valid signature !!", vbInformation
        Else
            Err.Raise ActiveLockErrCodeConstants.alerrCryptoAPIError, ACTIVELOCKSTRING, "Crypto API Error: Invalid signature"
        End If
    End If

End Select
Set cryptSession = Nothing
End Function


'===============================================================================
' Name: Function NewInstance
' Input: None
' Output: ActiveLock interface.
' Purpose: Obtains a new instance of an object that implements IActiveLock interface.
' <p>As of 2.0.5, this method will no longer initialize the instance automatically.
' Callers will have to call Init() by themselves subsequent to obtaining the instance.
' Remarks: None
'===============================================================================
Public Function NewInstance() As IActiveLock
Attribute NewInstance.VB_Description = "Returns a new instance of an object that implements IActiveLock interface."
    Dim NewInst As IActiveLock
    Set NewInst = New ActiveLock
    Set NewInstance = NewInst
End Function

'===============================================================================
' Name: Function CreateProductLicense
' Input:
'   ByVal name As String - Product/Software Name
'   ByVal Ver As String - Product version
'   ByVal Code As String - Product/Software Code
'   ByVal Flags As ActiveLock3.LicFlags - License Flag
'   ByVal LicType As ActiveLock3.ALLicType - License type
'   ByVal Licensee As String - Registered party for which the license has been issued
'   ByVal RegisteredLevel As String - Registered level
'   ByVal Expiration As String - Expiration date
'   ByVal LicKey As String - License key
'   ByVal RegisteredDate As String - Date on which the product is registered
'   ByVal Hash1 As String - Hash-1 code
'   ByVal MaxUsers As Integer - Maximum number of users allowed to use this license
' Output:
'   ProductLicense - License object
' Purpose: Instantiates a new ProductLicense object from the specified parameters.
' <p>If <code>LicType</code> is <i>Permanent</i>, then <code>Expiration</code> date parameter will be ignored.
' Remarks: None
'===============================================================================
Public Function CreateProductLicense(ByVal name As String, _
                                     ByVal Ver As String, _
                                     ByVal Code As String, _
                                     ByVal Flags As ActiveLock3.LicFlags, _
                                     ByVal LicType As ActiveLock3.ALLicType, _
                                     ByVal Licensee As String, _
                                     ByVal RegisteredLevel As String, _
                                     ByVal Expiration As String, _
                                     Optional ByVal LicKey As String, _
                                     Optional ByVal RegisteredDate As String, _
                                     Optional ByVal Hash1 As String, _
                                     Optional ByVal MaxUsers As Integer = 1, _
                                     Optional ByVal LicCode As String _
                                     ) As ProductLicense
    Dim NewLic As New ProductLicense
    With NewLic
        .ProductName = name
        .ProductKey = Code
        .ProductVer = Ver
        .LicenseClass = GetClassString(Flags)
        .LicenseType = LicType
        .Licensee = Licensee
        .RegisteredLevel = RegisteredLevel
        .MaxCount = MaxUsers
        ' ignore expiration date if license type is "permanent"
        If LicType <> allicPermanent Then
            .Expiration = Expiration
        End If
        If Not IsMissing(LicKey) Then
            .LicenseKey = LicKey
        End If
        If Not IsMissing(RegisteredDate) Then
            .RegisteredDate = RegisteredDate
        End If
        If Not IsMissing(Hash1) Then
            .Hash1 = Hash1
        End If
        ' New in v3.1
        ' LicenseCode is appended to the end so that we can know
        ' Alugen specified the hardware keys, and LockType
        ' was not specified by the protected app
        If Not IsMissing(LicCode) Then
            If LicCode <> "" Then .LicenseCode = LicCode
        End If
    End With
    Set CreateProductLicense = NewLic
End Function

'===============================================================================
' Name: Function GetClassString
' Input:
'   ByRef Flags As ActiveLock3.LicFlags - License flag string
' Output:
'   String - License flag string
' Purpose: Gets the license flag string such as MultiUser or Single
' Remarks: None
'===============================================================================
Private Function GetClassString(Flags As ActiveLock3.LicFlags) As String
    ' TODO: Decide the class numbers.
    ' lockMAC should probably be last,
    ' like it is in the enum. (IActivelock.cls)
    If Flags = alfMulti Then
        GetClassString = "MultiUser"
    Else ' default
        GetClassString = "Single"
    End If
End Function


'===============================================================================
' Name: Function GetLicTypeString
' Input:
'   LicType As ALLicType - License type object
' Output:
'   String - License type, such as Period, Permanent, Timed Expiry or None
' Purpose: Returns a string version of LicType
' Remarks: None
'===============================================================================
Private Function GetLicTypeString(LicType As ALLicType) As String
    'TODO: Implement this properly.
    If LicType = allicPeriodic Then
        GetLicTypeString = "Periodic"
    ElseIf LicType = allicPermanent Then
        GetLicTypeString = "Permanent"
    ElseIf LicType = allicTimeLocked Then
        GetLicTypeString = "Timed Expiry"
    Else ' default
        GetLicTypeString = "None"
    End If
End Function


'===============================================================================
' Name: Function TrimNulls
' Input:
'   ByVal str As String - String to be trimmed.
' Output:
'   String - Trimmed string.
' Purpose: Removes Null characters from the string.
' Remarks: None
'===============================================================================
Public Function TrimNulls(ByVal str As String) As String
Attribute TrimNulls.VB_Description = "Trim Null characters from the string."
    TrimNulls = modActiveLock.TrimNulls(str)
End Function

'===============================================================================
' Name: Function MD5Hash
' Input:
'   ByVal str As String - String to be hashed.
' Output:
'   String - Computed hash code.
' Purpose: Computes an MD5 hash of the specified string.
' Remarks: None
'===============================================================================
Public Function MD5Hash(ByVal str As String) As String
Attribute MD5Hash.VB_Description = "Computes an MD5 hash of the specified string."
    MD5Hash = modMD5.Hash(str)
End Function
'===============================================================================
' Name: Function Base64Encode
' Input:
'   ByVal str As String - String to be encoded.
' Output:
'   String - Encoded string.
' Purpose: Encodes a base64-decoded string.
' Remarks: None
'===============================================================================
Public Function Base64Encode(ByVal str As String) As String
Attribute Base64Encode.VB_Description = "Performs Base-64 encoding of the specified string."
    Base64Encode = modBase64.Base64_Encode(str)
End Function


'===============================================================================
' Name: Function Base64Decode
' Input:
'   ByVal strEncoded As String - String to be decoded.
' Output:
'   String - Decoded string.
' Purpose: Decodes a base64-encoded string.
' Remarks: None
'===============================================================================
Public Function Base64Decode(ByVal strEncoded As String) As String
Attribute Base64Decode.VB_Description = "Performs Base-64 decoding of the specified string."
    Base64Decode = modBase64.Base64_Decode(strEncoded)
End Function


