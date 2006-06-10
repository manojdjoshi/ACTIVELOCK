Option Strict Off
Option Explicit On 
Imports System.IO
Imports Microsoft.visualbasic.compatibility
Friend Class ActiveLock
    Implements _IActiveLock
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
    ' Name: Activelock
    ' Purpose: This is an implementation of IActiveLock.<p>It is not public-creatable, and so must only
    '   be accessed via ActiveLock.NewInstance() method.<p>Includes Key generation and validation routines.
    ' Remarks: If you want to turn off dll-checksumming, add this compilation flag to the Project Properties (Make tab) AL_DEBUG = 1
    ' Functions:
    ' Properties:
    ' Methods:
    ' Started: 21.04.2005
    ' Modified: 03.235.2006
    '===============================================================================

    ' @author: activelock-admins
    ' @version: 3.3.0
    ' @date: 03.23.2006

    ' Implements the IActiveLock interface.

    Private mSoftwareName As String
    Private mSoftwareVer As String
    Private mSoftwarePassword As String
    Private mSoftwareCode As String
    Private mRegisteredLevel As String
    Private mLockTypes As IActiveLock.ALLockTypes
    Private mUsedLockTypes As IActiveLock.ALLockTypes
    Private mTrialType As Integer
    Private mTrialLength As Integer
    Private mTrialHideTypes As IActiveLock.ALTrialHideTypes
    Private mKeyStore As _IKeyStoreProvider
    Private mKeyStorePath As String
    Private MyNotifier As New ActiveLockEventNotifier
    Private MyGlobals As New Globals_Renamed
    Private mLibKeyPath As String

    ' Registry hive used to store Activelock settings.
    Private Const AL_REGISTRY_HIVE As String = "Software\ActiveLock\ActiveLock3"

    ' Transients
    Private mfInit As Boolean ' flag to indicate that ActiveLock has been initialized
    '===============================================================================
    ' Name: Property Get IActiveLock_RegisteredLevel
    ' Input: None
    ' Output:
    '   String - License RegisteredLevel
    ' Purpose: Gets the Registered Level for the license after validating it.
    ' Remarks: None
    '===============================================================================
    Private ReadOnly Property IActiveLock_RegisteredLevel() As String Implements _IActiveLock.RegisteredLevel
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName)
            If Lic Is Nothing Then
                Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If
            ' Validate the License.
            ValidateLic(Lic)
            Return Lic.RegisteredLevel
        End Get
    End Property
    '===============================================================================
    ' Name: Property Let IActiveLock_AutoRegisterKeyPath
    ' Input:
    '   ByVal RHS As String - Liberation key file auto path name
    ' Output: None
    ' Purpose: IActiveLock Interface implementation
    '   <p>Specifies the liberation key auto file path name
    ' Remarks: None
    '===============================================================================
    ' IActiveLock Interface implementations
    ' @param RHS
    Private WriteOnly Property IActiveLock_AutoRegisterKeyPath() As String Implements _IActiveLock.AutoRegisterKeyPath
        Set(ByVal Value As String)
            mLibKeyPath = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Get AutoRegisterKeyPath
    ' Input: None
    ' Output: None
    ' Purpose: Sets the auto register file full path
    ' Remarks: None
    '===============================================================================
    Private ReadOnly Property AutoRegisterKeyPath() As String
        Get
            Return mLibKeyPath
        End Get
    End Property

    '===============================================================================
    ' Name: Property Get IActiveLock_EventNotifier
    ' Input: None
    ' Output: None
    ' Purpose: Gets a notification from Activelock
    ' Remarks: None
    '===============================================================================
    Private ReadOnly Property IActiveLock_EventNotifier() As ActiveLockEventNotifier Implements _IActiveLock.EventNotifier
        Get
            IActiveLock_EventNotifier = MyNotifier
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get IActiveLock_RegisteredDate
    ' Input: None
    ' Output:
    '   String - License registration date.
    ' Purpose: Gets the license registration date after validating it.
    ' Remarks: This is the date the license was generated by Alugen. NOT the date the license was activated.
    '===============================================================================
    Private ReadOnly Property IActiveLock_RegisteredDate() As String Implements _IActiveLock.RegisteredDate
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName)
            If Lic Is Nothing Then
                Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If
            ' Validate the License.
            ValidateLic(Lic)
            Return Lic.RegisteredDate
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get IActiveLock_RegisteredUser
    ' Input: None
    ' Output:
    '   String - Registered user name
    ' Purpose: Gets the registered user name after validating the license
    ' Remarks: None
    '===============================================================================
    Private ReadOnly Property IActiveLock_RegisteredUser() As String Implements _IActiveLock.RegisteredUser
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName)
            If Lic Is Nothing Then
                Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If
            ' Validate the License.
            ValidateLic(Lic)
            Return Lic.Licensee
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get IActiveLock_ExpirationDate
    ' Input: None
    ' Output:
    '   String - Expiration date of the license
    ' Purpose: Returns the expiration date of the license after validating it
    ' Remarks: None
    '===============================================================================
    Private ReadOnly Property IActiveLock_ExpirationDate() As String Implements _IActiveLock.ExpirationDate
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName)
            If Lic Is Nothing Then
                Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If
            ' Validate the License.
            ValidateLic(Lic)
            Return Lic.Expiration
        End Get
    End Property
    '===============================================================================
    ' Name: Property Let IActiveLock_KeyStorePath
    ' Input:
    '   ByVal RHS As String - License file path name
    ' Output: None
    ' Purpose: Specifies the license file path name
    ' Remarks: None
    '===============================================================================
    Private WriteOnly Property IActiveLock_KeyStorePath() As String Implements _IActiveLock.KeyStorePath
        Set(ByVal Value As String)
            If Not mKeyStore Is Nothing Then
                mKeyStore.KeyStorePath = Value
            End If
            mKeyStorePath = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Let IActiveLock_KeyStoreType
    ' Input:
    '   ByVal RHS As LicStoreType - License store type
    ' Output: None
    ' Purpose: Specifies the key store type
    '   <p>This version of Activelock does not work with the registry
    ' Remarks: Portions of this (RegistryKeyStoreProvider) not implemented yet
    '===============================================================================
    Private WriteOnly Property IActiveLock_KeyStoreType() As IActiveLock.LicStoreType Implements _IActiveLock.KeyStoreType
        Set(ByVal Value As IActiveLock.LicStoreType)
            ' Instantiate Key Store Provider
            If Value = IActiveLock.LicStoreType.alsFile Then
                mKeyStore = New FileKeyStoreProvider
            Else
                ' Set mKeyStore = New RegistryKeyStoreProvider
                ' TODO: Implement me!
                Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNotImplemented, ACTIVELOCKSTRING, STRNOTIMPLEMENTED)
            End If
            ' Set Key Store Path in KeyStoreProvider
            If mKeyStorePath <> "" Then
                mKeyStore.KeyStorePath = mKeyStorePath
            End If
        End Set
    End Property
    '===============================================================================
    ' Name: Property Let IActiveLock_LockType
    ' Input:
    '   ByVal RHS As ALLockTypes - ALLockTypes type
    ' Output: None
    ' Purpose: Specifies the ALLockTypes type
    ' Remarks: None
    '===============================================================================
    '===============================================================================
    ' Name: Property Get IActiveLock_LockType
    ' Input: None
    ' Output:
    '   ALLockTypes - Lock types type
    ' Purpose: Gets the ALLockTypes type
    ' Remarks: None
    '===============================================================================
    Private Property IActiveLock_LockType() As IActiveLock.ALLockTypes Implements _IActiveLock.LockType
        Get
            Return mLockTypes
        End Get
        Set(ByVal Value As IActiveLock.ALLockTypes)
            mLockTypes = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Let IActiveLock_UsedLockType
    ' Input:
    '   ByVal RHS As ALLockTypes - ALLockTypes type
    ' Output: None
    ' Purpose: Specifies the ALLockTypes type
    ' Remarks: None
    '===============================================================================
    '===============================================================================
    ' Name: Property Get IActiveLock_UsedLockType
    ' Input: None
    ' Output:
    '   ALLockTypes - Used Lock types type
    ' Purpose: Gets the ALLockTypes type
    ' Remarks: None
    '===============================================================================
    Private Property IActiveLock_UsedLockType() As IActiveLock.ALLockTypes Implements _IActiveLock.UsedLockType
        Get
            Return mUsedLockTypes
        End Get
        Set(ByVal Value As IActiveLock.ALLockTypes)
            mUsedLockTypes = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Get IActiveLock_TrialHideType
    ' Input: None
    ' Output:
    '   ALTrialHideTypes - Trial Hide types type
    ' Purpose: Gets the ALTrialHideTypes type
    ' Remarks: None
    '===============================================================================
    Private Property IActiveLock_TrialHideType() As IActiveLock.ALTrialHideTypes Implements _IActiveLock.TrialHideType
        Get
            Return mTrialHideTypes
        End Get
        Set(ByVal Value As IActiveLock.ALTrialHideTypes)
            mTrialHideTypes = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Get IActiveLock_SoftwareName
    ' Input: None
    ' Output:
    '   String - Software name  for the license
    ' Purpose: Gets the SoftwareName for the license
    ' Remarks: None
    '===============================================================================
    Private Property IActiveLock_SoftwareName() As String Implements _IActiveLock.SoftwareName
        Get
            Return mSoftwareName
        End Get
        Set(ByVal Value As String)
            mSoftwareName = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Let IActiveLock_SoftwarePassword
    ' Input:
    '   ByVal RHS As String - Software Password for the license
    ' Output: None
    ' Purpose: Specifies the SoftwarePassword for the license
    ' Remarks: None
    '===============================================================================

    '===============================================================================
    ' Name: Property Get IActiveLock_SoftwarePassword
    ' Input: None
    ' Output:
    '   String - Software Password for the license
    ' Purpose: Gets the SoftwarePassword for the license
    ' Remarks: None
    '===============================================================================
    Private Property IActiveLock_SoftwarePassword() As String Implements _IActiveLock.SoftwarePassword
        Get
            Return mSoftwarePassword
        End Get
        Set(ByVal Value As String)
            mSoftwarePassword = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Get IActiveLock_TrialType
    ' Input: None
    ' Output:
    '   ALTrialTypes - Trial Type  for the license
    ' Purpose: Gets the TrialType for the license
    ' Remarks: None
    '===============================================================================
    Private Property IActiveLock_TrialType() As IActiveLock.ALTrialTypes Implements _IActiveLock.TrialType
        Get
            Return mTrialType
        End Get
        Set(ByVal Value As IActiveLock.ALTrialTypes)
            mTrialType = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Get IActiveLock_TrialLength
    ' Input: None
    ' Output:
    '   Integer - Trial Length  for the license
    ' Purpose: Gets the TrialLength for the license
    ' Remarks: None
    '===============================================================================
    Private Property IActiveLock_TrialLength() As Integer Implements _IActiveLock.TrialLength
        Get
            Return mTrialLength
        End Get
        Set(ByVal Value As Integer)
            mTrialLength = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Get IActiveLock_InstallationCode
    ' Input:
    '   ByVal User As String - User name
    ' Output:
    '   String - Installation code
    ' Purpose: Combines the user name with the lock code and returns it as the installation code
    ' Remarks: None
    '===============================================================================
    Private ReadOnly Property IActiveLock_InstallationCode(Optional ByVal User As String = vbNullString, Optional ByVal Lic As ProductLicense = Nothing) As String Implements _IActiveLock.InstallationCode
        Get
            ' Generate Request code to Lock
            Dim strReq, strLock As String

            'Restrict user name to 2000 characters; need more? why?
            If Len(User) > 2000 Then
                Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrUserNameTooLong, ACTIVELOCKSTRING, STRUSERNAMETOOLONG)
            End If

            ' New in v3.1
            ' Version 3.1 and above of Activelock will append the "+" sign
            ' in front of the installation code whenever lockNone is used or
            ' lockType is not specified in the protected app.
            ' When "+" is not found at the beginning of the installation code,
            ' Alugen will not allow users pick the hardware lock method since this
            ' corresponds to an installation code which
            ' utilizes a hardware lock option specified inside the protected app.
            If mLockTypes = IActiveLock.ALLockTypes.lockNone Then
                strLock = "+" & IActiveLock_LockCode()
            Else
                strLock = IActiveLock_LockCode()
            End If

            ' combine with user name
            strReq = strLock & vbLf & User

            ' base-64 encode the request
            Dim strReq2 As String
            strReq2 = modBase64.Base64_Encode(strReq)
            Return strReq2

            ' New in v3.1
            ' If there's a license and the LicenseCode exists, then use it
            ' LicenseCode is actually the Installation Code modified by Alugen
            ' LicenseCode is appended to the end of the lic file so that we can know
            ' Alugen specified the hardware keys, and LockType
            ' was not specified inside the protected app
            If Not Lic Is Nothing Then
                If Lic.LicenseCode <> "" Then
                    Return Lic.LicenseCode
                    If Left(IActiveLock_InstallationCode, 1) = "+" Then Return Mid(IActiveLock_InstallationCode, 2)
                    ' We won't do the following in order to maintain backwards compatibility with existing licenses
                    ' ElseIf Lic.LicenseCode = "" And mLockTypes = lockNone Then
                    ' Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
                End If
            End If
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get IActiveLock_SoftwareVersion
    ' Input: None
    ' Output:
    '   String - Software version  for the license
    ' Purpose: Gets the SoftwareVersion for the license
    ' Remarks: None
    '===============================================================================
    Private Property IActiveLock_SoftwareVersion() As String Implements _IActiveLock.SoftwareVersion
        Get
            Return mSoftwareVer
        End Get
        Set(ByVal Value As String)
            mSoftwareVer = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Let IActiveLock_SoftwareCode
    ' Input:
    '   ByVal RHS As String - Software code for the license
    ' Output: None
    ' Purpose: Specifies the SoftwareCode for the license
    ' Remarks: SoftwareCode is an RSA public key.  This code will be used to verify license keys later on
    '===============================================================================
    Private WriteOnly Property IActiveLock_SoftwareCode() As String Implements _IActiveLock.SoftwareCode
        Set(ByVal Value As String)
            mSoftwareCode = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Get IActiveLock_UsedDays
    ' Input: None
    ' Output: None
    ' Purpose: Gets the number of days the license was used after validating it.
    ' Remarks: None
    '===============================================================================
    Private ReadOnly Property IActiveLock_UsedDays() As Integer Implements _IActiveLock.UsedDays
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName)

            If Lic Is Nothing Then Exit Property

            ' validate the license
            ValidateLic(Lic)
            'IActiveLock_UsedDays = CInt(DateDiff("d", Lic.RegisteredDate, Now.UtcNow))
            IActiveLock_UsedDays = CInt(DateDiff("d", CDate(Replace(Lic.RegisteredDate, ".", "-")), Now.UtcNow))
            If IActiveLock_UsedDays < 0 Then
                Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
            End If
        End Get
    End Property
    'Class_Initialize was upgraded to Class_Initialize_Renamed
    Private Sub Class_Initialize_Renamed()
        ' Default to alsFile
        IActiveLock_KeyStoreType = IActiveLock.LicStoreType.alsFile
    End Sub
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

    '===============================================================================
    ' Name: Sub IActiveLock_Init
    ' Input:
    '   ByRef autoLicString As String - Returned License Key of AutoRegister is successful
    ' Output: None
    ' Purpose: Initalizes Activelock
    ' Remarks: Performs CRC check on Alcrypto
    '   <p>Performs auto license registration if the license file is found
    '===============================================================================
    Private Sub IActiveLock_Init(Optional ByVal strPath As String = "", Optional ByRef autoLicString As String = "") Implements _IActiveLock.Init
        ' If running in Debug mode, don't bother with dll authentication
#If CBool(AL_DEBUG) <> False Then
		GoTo Done
#End If
        ' Checksum ALCrypto3NET.dll
        Const ALCRYPTO_MD5 As String = "54BED793A0E24D3E71706EEC4FA1B0FC"
        'Const ALCRYPTO_MD5$ = "be299ad0f52858fdd9ea3626468dc05c"

        Dim strdata, strMD5, usedFile As String
        ' .NET version of Activelock Init() now supports an optional path string
        ' for the Alcrypto3NET.dll
        ' This is needed for the cases where the user does not have the luxury of
        ' placing this file in the system32 directory
        If strPath <> "" Then
            usedFile = strPath & "\alcrypto3NET.dll"
        Else
            usedFile = WinSysDir() & "\alcrypto3NET.dll"
        End If
        If File.Exists(usedFile) = False Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrFileTampered, "ActiveLock3", "Alcrypto3Net.dll could not be found.")
        End If
        Call modActiveLock.ReadFile(usedFile, strdata)
        ' use the .NET's native MD5 functions instead of our own MD5 hashing routine
        ' and instead of ALCrypto's md5_hash() function.
        strMD5 = UCase(strdata)    '<--- ReadFile procedure already computes the MD5.Hash
        System.Diagnostics.Debug.WriteLine("ALCrypto Hash: " & strMD5)
        System.Diagnostics.Debug.WriteLine("strdata: " & strdata)

        If strMD5 <> ALCRYPTO_MD5 Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrFileTampered, ACTIVELOCKSTRING, STRFILETAMPERED)
        End If
        ' Perform automatic license registration
        If AutoRegisterKeyPath <> "" Then
            DoAutoRegistration(autoLicString)
            If Err.Number <> 0 Then autoLicString = ""
        End If
Done:
        mfInit = True
    End Sub
    '===============================================================================
    ' Name: Sub DoAutoRegistration
    ' Input:
    '   strLibKey As String - Returned liberation key if auto register is successful
    ' Output: None
    ' Purpose: Checks the specified path to see if the auto registration liberation file is there
    ' Remarks: None
    '===============================================================================
    Private Sub DoAutoRegistration(ByRef strLibKey As String)

        ' Don't bother to proceed unless the file is there.
        If Not File.Exists(AutoRegisterKeyPath) Then Exit Sub

        ReadLibKey(AutoRegisterKeyPath, strLibKey)
        IActiveLock_Register(strLibKey)

        ' If registration is successful, delete the liberation file so we won't register the same file on next startup
        Kill(AutoRegisterKeyPath)
    End Sub
    '===============================================================================
    ' Name: Sub ReadLibKey
    ' Input:
    '   ByVal sFileName As String - File name to read the liberation key from
    '   ByRef strLibKey As String -  Liberation key returned
    ' Output: None
    ' Purpose: Reads the liberation key from a file
    ' Remarks: None
    '===============================================================================
    Private Sub ReadLibKey(ByVal sFileName As String, ByRef strLibKey As String)
        Dim hFile As Integer
        hFile = FreeFile()
        FileOpen(hFile, sFileName, OpenMode.Input)
        On Error GoTo finally_Renamed
        strLibKey = InputString(hFile, LOF(hFile))
finally_Renamed:
        FileClose(hFile)
    End Sub
    '===============================================================================
    ' Name: Sub IActiveLock_Acquire
    ' Input:
    '   ByRef SoftwareName As String - Software name.
    '   ByRef SoftwareVer As String - Software version.
    ' Output: None
    ' Purpose: Acquires an Activelock License.
    '<p>This is the main method that retrieves an Activelock license, validates it, and ends the trial license if it exists.
    ' Remarks: None
    '===============================================================================
    Private Sub IActiveLock_Acquire(Optional ByRef strMsg As String = "") Implements _IActiveLock.Acquire
        Dim trialActivated As Boolean
        'Check the Key Store Provider
        If mKeyStore Is Nothing Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrKeyStoreInvalid, ACTIVELOCKSTRING, STRKEYSTOREUNINITIALIZED)
        End If

        Dim Lic As ProductLicense
        Lic = mKeyStore.Retrieve(mSoftwareName)

        Dim trialStatus As Boolean

        If Lic Is Nothing Then
            ' There's no valid license, so let's see if we can grant this user a "Trial License"
            If mTrialType = IActiveLock.ALTrialTypes.trialNone Then 'No Trial
                GoTo noRegistration
            End If

            On Error GoTo noRegistration
            strMsg = ""
            If mTrialHideTypes = 0 Then
                mTrialHideTypes = IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialRegistry Or IActiveLock.ALTrialHideTypes.trialSteganography
            End If

            ' Using this trick to temporarily set the date format to mm/dd/yyyy
            Get_locale() ' Get the current date format and save it to regionalSymbol variable
            Set_locale((""))
            trialStatus = ActivateTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, strMsg, mSoftwarePassword)
            ' Set the locale date format to what we had before; can't leave changed
            Set_locale((regionalSymbol))
            If trialStatus = True Then
                Exit Sub
            End If
            GoTo continueRegistration

noRegistration:
            Set_locale((regionalSymbol))
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
        End If

continueRegistration:
        ' Validate license
        ValidateLic(Lic)
    End Sub
    '===============================================================================
    ' Name: Sub ValidateKey
    ' Input:
    '   Lic As ProductLicense - Product license
    ' Output: None
    ' Purpose: Validates the License Key using RSA signature verification.
    '   <p>License key contains the RSA signature of IActiveLock_LockCode.
    ' Remarks: None
    '===============================================================================
    Private Sub ValidateKey(ByRef Lic As ProductLicense)
        ' make sure software code is set
        If mSoftwareCode = "" Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNotInitialized, ACTIVELOCKSTRING, STRNOSOFTWARECODE)
        End If

        Dim Key As RSAKey
        Dim strPubKey As String
        strPubKey = mSoftwareCode

        Dim strSig As String
        Dim strLic As String
        Dim strLicKey As String

        strLic = IActiveLock_LockCode(Lic)
        strLicKey = Lic.LicenseKey

        ' decode the license key
        strSig = Base64_Decode(strLicKey)

        ' Print out some info for debugging purposes
        'System.Diagnostics.Debug.WriteLine("Code1: " & strPubKey)
        'System.Diagnostics.Debug.WriteLine("Lic: " & strLic)
        'System.Diagnostics.Debug.WriteLine("Lic hash: " & modMD5.Hash(strLic))
        'System.Diagnostics.Debug.WriteLine("LicKey: " & strLicKey)
        'System.Diagnostics.Debug.WriteLine("Sig: " & strSig)
        'System.Diagnostics.Debug.WriteLine("Verify: " & RSAVerify(strPubKey, strLic, strSig))
        'System.Diagnostics.Debug.WriteLine("====================================================")

        ' validate the key
        Dim rc As Integer
        rc = RSAVerify(strPubKey, strLic, strSig)
        If rc <> 0 Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
        End If

        ' Check if license has not expired
        ' but don't do it if there's no expiration date
        If Lic.Expiration = "" Then Exit Sub
        Dim dtExp As Date
        dtExp = CDate(Lic.Expiration)
        If Now.UtcNow > dtExp And Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
            ' ialkan - 9-23-2005 added the following to update and store the license
            ' with the new LastUsed property; otherwise setting the clock back next time
            ' might bypass the protection mechanism
            ' Update last used date
            UpdateLastUsed(Lic)
            mKeyStore.Store(Lic)
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseExpired, ACTIVELOCKSTRING, STRLICENSEEXPIRED)
        End If
    End Sub
    '===============================================================================
    ' Name: Sub ValidateLic
    ' Input:
    '  Lic As ProductLicense - Product License
    ' Output: None
    ' Purpose: Validates the entire license (including lastused, etc.)
    ' Remarks: None
    '===============================================================================
    Private Sub ValidateLic(ByRef Lic As ProductLicense)
        ' make sure we're initialized.
        If Not mfInit Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNotInitialized, ACTIVELOCKSTRING, STRNOTINITIALIZED)
        End If

        ' validate license key first
        ValidateKey(Lic)

        Dim strEncrypted, strHash As String
        ' Validate last run date
        strEncrypted = Lic.LastUsed
        MyNotifier.Notify("ValidateValue", strEncrypted)
        strHash = modMD5.Hash(strEncrypted)
        If strHash <> Lic.Hash1 Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseTampered, ACTIVELOCKSTRING, STRLICENSETAMPERED)
        End If

        ' try to detect the user setting their system clock back
        ' Need to account for Daylight Savings Time
        Dim strNow As String
        ' Normalize to the format of the saved date-time, before we compare
        strNow = Microsoft.VisualBasic.Compatibility.VB6.Format(Now.UtcNow, "YYYY/MM/DD")
        'If strNow < Lic.LastUsed And Lic.LicenseType <> allicPermanent Then
        If DateValue(strNow) < DateValue(Microsoft.VisualBasic.Compatibility.VB6.Format(Lic.LastUsed, "YYYY/MM/DD")) And Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
            'System.Diagnostics.Debug.WriteLine("UTC Now: " & strNow)
            'System.Diagnostics.Debug.WriteLine("LastUsed: " & CDate(Lic.LastUsed))
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
        End If
        UpdateLastUsed(Lic)
        mKeyStore.Store(Lic)
    End Sub
    '===============================================================================
    ' Name: Sub UpdateLastUsed
    ' Input:
    '   Lic As ProductLicense - Product License
    ' Output: None
    ' Purpose: Updates LastUsed property with current date stamp.
    ' Remarks: None
    '===============================================================================
    Private Sub UpdateLastUsed(ByRef Lic As ProductLicense)
        ' Update license store with LastRunDate
        Dim strEncrypted As String
        Dim strLastUsed As String
        strLastUsed = Microsoft.VisualBasic.Compatibility.VB6.Format(Now.UtcNow, "YYYY/MM/DD")
        Lic.LastUsed = strLastUsed
        MyNotifier.Notify("ValidateValue", strLastUsed)
        Lic.Hash1 = modMD5.Hash(strLastUsed)
    End Sub
    '===============================================================================
    ' Name: Sub IActiveLock_Register
    ' Input:
    '   ByVal LibKey As String - Liberation Key
    ' Output: None
    ' Purpose: Registers Activelock license with a given liberation key
    ' Remarks: None
    '===============================================================================
    Private Sub IActiveLock_Register(ByVal LibKey As String) Implements _IActiveLock.Register

        Dim Lic As ActiveLock3_5NET.ProductLicense = New ActiveLock3_5NET.ProductLicense

        Lic.Load(LibKey)

        ' Validate that the license key.
        '   - registered user
        '   - expiry date
        Dim varResult As Object
        ValidateKey(Lic)

        ' License was validated successfuly. Check clock tampering for non-permanent licenses.
        If Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
            If ClockTampering() Then
                If SystemClockTampered() Then
                    Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                End If
            End If
        End If

        ' License was validated successfuly.  Store it.
        If mKeyStore Is Nothing Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrKeyStoreInvalid, ACTIVELOCKSTRING, STRKEYSTOREUNINITIALIZED)
        End If

        ' Update last used date
        UpdateLastUsed(Lic)
        mKeyStore.Store(Lic)

        ' Expire all trial licenses
        On Error Resume Next
        ' Expire the Trial
        Dim trialStatus As Boolean
        If mTrialType <> IActiveLock.ALTrialTypes.trialNone Then
            trialStatus = ExpireTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, mSoftwarePassword)
        End If

    End Sub
    '===============================================================================
    ' Name: Sub IActiveLock_KillTrial
    ' Input: None
    ' Output: None
    ' Purpose: Kills a Trial License
    ' Remarks: None
    '===============================================================================
    Private Sub IActiveLock_KillTrial() Implements _IActiveLock.KillTrial
        On Error Resume Next
        'Expire the Trial
        Dim trialStatus As Boolean
        If mSoftwareName = "" Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNoSoftwareName, ACTIVELOCKSTRING, STRNOSOFTWARENAME)
        ElseIf mSoftwareVer = "" Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNoSoftwareVersion, ACTIVELOCKSTRING, STRNOSOFTWAREVERSION)
        Else
            trialStatus = ExpireTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, mSoftwarePassword)
        End If
    End Sub
    '===============================================================================
    ' Name: Sub IActiveLock_ResetTrial
    ' Input: None
    ' Output: None
    ' Purpose: Resets a Trial License
    ' Remarks: None
    '===============================================================================
    Private Sub IActiveLock_ResetTrial() Implements _IActiveLock.ResetTrial
        On Error Resume Next
        'Reset the Trial
        Dim trialStatus As Boolean
        If mSoftwareName = "" Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNoSoftwareName, ACTIVELOCKSTRING, STRNOSOFTWARENAME)
        ElseIf mSoftwareVer = "" Then
            Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNoSoftwareVersion, ACTIVELOCKSTRING, STRNOSOFTWAREVERSION)
        Else
            trialStatus = ResetTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, mSoftwarePassword)
        End If
    End Sub
    '===============================================================================
    ' Name: Function IActiveLock_LockCode
    ' Input:
    '   ByRef Lic As ProductLicense - Product License
    ' Output:
    '   String - Lock code
    ' Purpose: Returns the lock code from a given Activelock license
    ' Remarks: v3 includes the new lockHDFirmware option
    '===============================================================================
    Private Function IActiveLock_LockCode(Optional ByRef Lic As ProductLicense = Nothing) As String Implements _IActiveLock.LockCode
        Dim strLock As String
        Dim noKey As String
        Dim userFromInstallCode, usedcode As String
        Dim tmpLockType As IActiveLock.ALLockTypes
        Dim bLockNone As Boolean
        Dim j As Short
        Dim Index, i As Short
        Dim a() As String
        Dim aString As String ' we have lockNone ' use temp in case of failure.

        noKey = Chr(110) & Chr(111) & Chr(107) & Chr(101) & Chr(121)
        If Lic Is Nothing Then
            ' New in v3.1
            ' Modified this function on 1-13-2006 to append ALL hardware keys
            ' to the Installation Code. This way, it will be decided in Alugen which
            ' hardware keys will be used to lock the license to
            ' If there's already a lock selected in the protected app,
            ' such as lockHDfirmware or lockComputer, then Alugen will show
            ' these two options and will gray them out (fix these two selections)
            If mLockTypes = IActiveLock.ALLockTypes.lockNone Then
                strLock = ""
                AppendLockString(strLock, modHardware.GetMACAddress())
                AppendLockString(strLock, modHardware.GetComputerName())
                AppendLockString(strLock, modHardware.GetHDSerial())
                AppendLockString(strLock, modHardware.GetHDSerialFirmware())
                AppendLockString(strLock, modHardware.GetWindowsSerial())
                AppendLockString(strLock, modHardware.GetBIOSserial())
                AppendLockString(strLock, modHardware.GetMotherboardSerial())
                AppendLockString(strLock, modHardware.GetIPaddress())
            Else
                If mLockTypes And IActiveLock.ALLockTypes.lockMAC Then
                    AppendLockString(strLock, modHardware.GetMACAddress())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If mLockTypes And IActiveLock.ALLockTypes.lockComp Then
                    AppendLockString(strLock, modHardware.GetComputerName())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If mLockTypes And IActiveLock.ALLockTypes.lockHD Then
                    AppendLockString(strLock, modHardware.GetHDSerial())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If mLockTypes And IActiveLock.ALLockTypes.lockHDFirmware Then
                    AppendLockString(strLock, modHardware.GetHDSerialFirmware())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If mLockTypes And IActiveLock.ALLockTypes.lockWindows Then
                    AppendLockString(strLock, modHardware.GetWindowsSerial())
                Else
                    AppendLockString(strLock, noKey)
                End If
                ' New in v3.1
                If mLockTypes And IActiveLock.ALLockTypes.lockBIOS Then
                    AppendLockString(strLock, modHardware.GetBIOSserial())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If mLockTypes And IActiveLock.ALLockTypes.lockMotherboard Then
                    AppendLockString(strLock, modHardware.GetMotherboardSerial())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If mLockTypes And IActiveLock.ALLockTypes.lockIP Then
                    AppendLockString(strLock, modHardware.GetIPaddress())
                Else
                    AppendLockString(strLock, noKey)
                End If
            End If

            If Left(strLock, 1) = vbLf Then strLock = Mid(strLock, 2)

            ' Append lockcode.
            ' Note: The logic here must match the corresponding logic
            '       in ALUGENLib.Generator_GenKey()
            IActiveLock_LockCode = strLock
            Else
            ' We have a License
            ' New in v3.1
                ' In such cases when Alugen modifies the Installation Code and sends it
                ' back, we need to retrieve in here and process it
                ' Modified Installation Code is appended to the end of the Liberation Key
                ' The modified Installation Code is also stored in the license file
                ' otherwise we'd never know which hardware leys were used to lock the license
            'IActiveLock_LockCode = Lic.ToString_Renamed() & vbLf & strLock

                ' Per David Weatherall ' New in v3.3
                tmpLockType = IActiveLock.ALLockTypes.lockNone ' lockNone = 0 so starting value

                If Lic.LicenseCode <> "" Then
                    If Left(Lic.LicenseCode, 1) = "+" Then
                        usedcode = modBase64.Base64_Decode(Mid(Lic.LicenseCode, 2))
                        bLockNone = True ' per David Weatherall
                    Else
                        usedcode = modBase64.Base64_Decode((Lic.LicenseCode))
                        bLockNone = False ' per David Weatherall
                    End If

                    a = Split(usedcode, vbLf)
                    For j = LBound(a) To UBound(a) - 1
                        aString = a(j)
                        If Left(aString, 1) = "+" Then aString = Mid(aString, 2)
                        If j = LBound(a) Then
                            If aString <> noKey Then
                                If Not bLockNone Then tmpLockType = tmpLockType Or IActiveLock.ALLockTypes.lockMAC ' build up lockType per David Weatherall
                                If aString <> modHardware.GetMACAddress() Then
                                    Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                                End If
                            End If
                        ElseIf j = LBound(a) + 1 Then
                            If aString <> noKey Then
                                If Not bLockNone Then tmpLockType = tmpLockType Or IActiveLock.ALLockTypes.lockComp ' build up lockType per David Weatherall
                                If aString <> modHardware.GetComputerName() Then
                                    Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                                End If
                            End If
                        ElseIf j = LBound(a) + 2 Then
                            If aString <> noKey Then
                                If Not bLockNone Then tmpLockType = tmpLockType Or IActiveLock.ALLockTypes.lockHD ' build up lockType per David Weatherall
                                If aString <> modHardware.GetHDSerial() Then
                                    Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                                End If
                            End If
                        ElseIf j = LBound(a) + 3 Then
                            If aString <> noKey Then
                                If Not bLockNone Then tmpLockType = tmpLockType Or IActiveLock.ALLockTypes.lockHDFirmware ' build up lockType per David Weatherall
                                If aString <> modHardware.GetHDSerialFirmware() Then
                                    Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                                End If
                            End If
                        ElseIf j = LBound(a) + 4 Then
                            If aString <> noKey Then
                                If Not bLockNone Then tmpLockType = tmpLockType Or IActiveLock.ALLockTypes.lockWindows ' build up lockType per David Weatherall
                                If aString <> modHardware.GetWindowsSerial() Then
                                    Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                                End If
                            End If
                        ElseIf j = LBound(a) + 5 Then
                            If aString <> noKey Then
                                If Not bLockNone Then tmpLockType = tmpLockType Or IActiveLock.ALLockTypes.lockBIOS ' build up lockType per David Weatherall
                                If aString <> modHardware.GetBIOSserial() Then
                                    Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                                End If
                            End If
                        ElseIf j = LBound(a) + 6 Then
                            If aString <> noKey Then
                                If Not bLockNone Then tmpLockType = tmpLockType Or IActiveLock.ALLockTypes.lockMotherboard ' build up lockType per David Weatherall
                                If aString <> modHardware.GetMotherboardSerial() Then
                                    Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                                End If
                            End If
                        ElseIf j = LBound(a) + 7 Then
                            If aString <> noKey Then
                                If Not bLockNone Then tmpLockType = tmpLockType Or IActiveLock.ALLockTypes.lockIP ' build up lockType per David Weatherall
                                If aString <> modHardware.GetIPaddress() Then
                                    Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                                End If
                            End If
                        End If
                    Next j

                    Index = 0
                    i = 1
                    ' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
                    Do While i > 0
                        i = InStr(Index + 1, usedcode, vbLf)
                        If i > 0 Then Index = i
                    Loop
                    ' user name starts from Index+1 to the end
                    userFromInstallCode = Mid(usedcode, Index + 1)
                    ' Check to see if this user name matches the one in the liberation key
                    If userFromInstallCode <> Lic.Licensee Then
                        Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                    End If
                    ' above is last possible failure point
                    mUsedLockTypes = tmpLockType ' per David Weatherall

                    usedcode = Mid(usedcode, 1, Len(usedcode) - Len(userFromInstallCode) - 1)

                    IActiveLock_LockCode = Lic.ToString_Renamed() & vbLf & usedcode
                Else
                    Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                End If
        End If
    End Function
    '===============================================================================
    ' Name: Sub AppendLockString
    ' Input:
    '   ByRef strLock As String - The lock string to be appended to, returns as an output
    '   ByVal newSubString As String - The string to be appended to the lock string if strLock is empty string
    ' Output:
    '   Appended lock string and installation code
    ' Purpose: Appends the lock string to the given installation code
    ' Remarks: None
    '===============================================================================
    Private Sub AppendLockString(ByRef strLock As String, ByVal newSubString As String)
        If strLock = "" Then
            strLock = newSubString
        Else
            strLock = strLock & vbLf & newSubString
        End If
    End Sub
    '===============================================================================
    ' Name: Function IActiveLock_Transfer
    ' Input:
    '   ByVal OtherSoftwareCode As String - Installation code from another machine/software
    ' Output: None
    ' Purpose: Not implemented yet
    ' Remarks: Transfers an Activelock license from one machine/software to another
    '===============================================================================
    Private Function IActiveLock_Transfer(ByVal OtherSoftwareCode As String) As String Implements _IActiveLock.Transfer
        ' TODO: Implement me!
        Err.Raise(Globals_definst.ActiveLockErrCodeConstants.alerrNotImplemented, ACTIVELOCKSTRING, STRNOTIMPLEMENTED)
    End Function
End Class