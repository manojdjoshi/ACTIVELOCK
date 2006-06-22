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
' @param alerrKeyStoreInvalid       Key Store Provider has not been initialized yet.
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

Public Enum ActiveLockErrCodeConstants
    alerrOK = 0   ' successful
    alerrNoLicense = &H80040001         ' vbObjectError (&H80040000) + 1
    alerrLicenseInvalid = &H80040002
    alerrLicenseExpired = &H80040003
    alerrLicenseTampered = &H80040004
    alerrClockChanged = &H80040005
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
End Enum

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
        'If LicType = allicNetwork Then
        '    .LicenseClass = alfMulti
        'Else
            .LicenseClass = GetClassString(Flags)
        'End If
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


