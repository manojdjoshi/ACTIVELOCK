Option Strict Off
Option Explicit On
'Class instancing was changed to public
<System.Runtime.InteropServices.ProgId("Globals_Renamed_NET.Globals_Renamed")> Public Class Globals_Renamed
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
    ' Modified: 03.24.2006
	'===============================================================================
	' @author activelock-admins
    ' @version 3.3.0
    ' @date 03.24.2006
	'

	' ActiveLock Error Codes.
	' These error codes are used for <code>Err.Number</code> whenever ActiveLock raises an error.
	'
	' @param AlerrOK                    No error. Operation was successful.
	' @param AlerrNoLicense             No license available.
	' @param AlerrLicenseInvalid        License is invalid.
	' @param AlerrLicenseExpired        License has expired.
	' @param AlerrLicenseTampered       License has been tampered.
	' @param AlerrClockChanged          System clock has been changed.
	' @param AlerrKeyStoreInvalid       Key Store Provider has not been initialized yet.
	' @param AlerrFileTampered          ActiveLock DLL file has been tampered.
	' @param AlerrNotInitialized        ActiveLock has not been initialized yet.
	' @param AlerrNotImplemented        An ActiveLock operation has not been implemented.
	' @param AlerrUserNameTooLong       Maximum User Name length of 2000 characters has been exceeded.
	' @param AlerrInvalidTrialDays      Specified number of Free Trial Days is invalid (possibly <=0).
	' @param AlerrInvalidTrialRuns      Specified number of Free Trial Runs is invalid (possibly <=0).
	' @param AlerrTrialInvalid          Trial is invalid.
	' @param AlerrTrialDaysExpired      Trial Days have expired.
	' @param AlerrTrialRunsExpired      Trial Runs have expired.
	' @param AlerrNoSoftwareName        Software Name has not been specified.
	' @param AlerrNoSoftwareVersion     Software Version has not been specified.
    ' @param AlerrRSAError              Something went wrong in the RSA routines.

	Public Enum ActiveLockErrCodeConstants
		AlerrOK = 0 ' successful
		AlerrNoLicense = &H80040001 ' vbObjectError (&H80040000) + 1
		AlerrLicenseInvalid = &H80040002
		AlerrLicenseExpired = &H80040003
		AlerrLicenseTampered = &H80040004
		AlerrClockChanged = &H80040005
		AlerrKeyStoreInvalid = &H80040010
		AlerrFileTampered = &H80040011
		AlerrNotInitialized = &H80040012
		AlerrNotImplemented = &H80040013
		AlerrUserNameTooLong = &H80040014
		AlerrInvalidTrialDays = &H80040020
		AlerrInvalidTrialRuns = &H80040021
		AlerrTrialInvalid = &H80040022
		AlerrTrialDaysExpired = &H80040023
		AlerrTrialRunsExpired = &H80040024
		AlerrNoSoftwareName = &H80040025
		AlerrNoSoftwareVersion = &H80040026
        AlerrRSAError = &H80040027
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
    Public Function NewInstance() As _IActiveLock
        Dim NewInst As _IActiveLock
        NewInst = New ActiveLock
        NewInstance = NewInst
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
    Public Function CreateProductLicense(ByVal Name As String, ByVal Ver As String, ByVal Code As String, ByVal Flags As ProductLicense.LicFlags, ByVal LicType As ProductLicense.ALLicType, ByVal Licensee As String, ByVal RegisteredLevel As String, ByVal Expiration As String, Optional ByVal LicKey As String = "", Optional ByVal RegisteredDate As String = "", Optional ByVal Hash1 As String = "", Optional ByVal MaxUsers As Short = 1, Optional ByVal LicCode As String = "") As ProductLicense
        Dim NewLic As New ProductLicense
        With NewLic
            .ProductName = Name
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
            If LicType <> ProductLicense.ALLicType.allicPermanent Then
                .Expiration = Expiration
            End If
            'IsMissing() was changed to IsNothing()
            If Not IsNothing(LicKey) Then
                .LicenseKey = LicKey
            End If
            'IsMissing() was changed to IsNothing()
            If Not IsNothing(RegisteredDate) Then
                .RegisteredDate = RegisteredDate
            End If
            'IsMissing() was changed to IsNothing()
            If Not IsNothing(Hash1) Then
                .Hash1 = Hash1
            End If
            ' New in v3.1
            ' LicenseCode is appended to the end so that we can know
            ' Alugen specified the hardware keys, and LockType
            ' was not specified by the protected app
            'IsMissing() was changed to IsNothing()
            If Not IsNothing(LicCode) Then
                If LicCode <> "" Then .LicenseCode = LicCode
            End If
        End With
        CreateProductLicense = NewLic
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
    Private Function GetClassString(ByRef Flags As ProductLicense.LicFlags) As String
        ' TODO: Decide the class numbers.
        ' lockMAC should probably be last,
        ' like it is in the enum. (IActivelock.cls)
        If Flags = ProductLicense.LicFlags.alfMulti Then
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
    Private Function GetLicTypeString(ByRef LicType As ProductLicense.ALLicType) As String
        'TODO: Implement this properly.
        If LicType = ProductLicense.ALLicType.allicPeriodic Then
            GetLicTypeString = "Periodic"
        ElseIf LicType = ProductLicense.ALLicType.allicPermanent Then
            GetLicTypeString = "Permanent"
        ElseIf LicType = ProductLicense.ALLicType.allicTimeLocked Then
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
    'str was upgraded to str_Renamed
    Public Function TrimNulls(ByVal str_Renamed As String) As String
        TrimNulls = modActiveLock.TrimNulls(str_Renamed)
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
    'str was upgraded to str_Renamed
    Public Function MD5Hash(ByVal str_Renamed As String) As String
        MD5Hash = modMD5.Hash(str_Renamed)
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
    'str was upgraded to str_Renamed
    Public Function Base64Encode(ByVal str_Renamed As String) As String
        Base64Encode = modBase64.Base64_Encode(str_Renamed)
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
        Base64Decode = modBase64.Base64_Decode(strEncoded)
    End Function
End Class