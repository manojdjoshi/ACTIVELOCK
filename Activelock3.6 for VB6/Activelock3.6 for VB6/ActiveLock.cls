VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActiveLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Modified: 08.05.2005
'===============================================================================

' @author: activelock-admins
' @version: 3.0.0
' @date: 08.05.2005
'
'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
'
' @todo Implement IActiveLock_Transfer. This is the functionality to transfer
'       a license from one machine to another.

'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' @history
' <pre>
' 04.21.05 - sentax        - Added Hard Drive Firmware Serial module.
' 07.07.03 - mcrute        - Updated the header comments for this file.
' 07.28.03 - th2tran       - Using RSA signature for license key.
' 07.28.03 - wizzardme2000 - Implemented lockComp, lockHD, and lockWindows
' 07.31.03 - th2tran       - Perform checksum on ALCrypto3.dll on Init()
'                          - ValidateKey() to bypass checking expiry if there's no expiration date.
' 08.21.03 - th2tran       - Added #If directive in IActiveLock_Init() to bypass dll authentication when
'                            running from inside VB.  This makes debugging a bit easier.
'                          - Refactored IActiveLock_LockCode()
' 10.06.03 - th2tran       - For ALCrypto checksumming, we must put ALCrypto3.dll in the Windows directory
'                            instead of the working directory.  Otherwise, users will have to put ALCrypto3.dll in
'                            each of their apps' working directory, which is OK if they only have 1 app,
'                            but not so OK when multiple apps are involved.
' 10.13.03 - th2tran       - Fixed bug: If ValidateValue event was not handled by the client, then we ended up
'                            hashing an empty string.
' 11.02.03 - th2tran       - ALCrypto3.dll now resides in the Windows System directory
'                          - IActiveLock_Transfer raise alerrNotImplemented error until we get this implemented
' 02.24.04 - th2tran       - Fixed bug: mismatched logic between IActiveLock.LockCode() and ALUGENLib.Generator_GenKey()
'                            implementations caused the key validation process to fail for keys generated
'                            from blank LockCode.
' 02.25.04 - th2tran       - Changed IActiveLock_Register() not to validate the license (leave this task to IActiveLock_Acquire()).
'                            This allows ALUGEN to use this method to save the generated license key to the license file.
'                            Potentially, with this approach we can do away with handing liberation keys to the users.
'                            Instead, we would be handing them a license file to plop into their application directory...a much simplier
'                            process than manually registering the liberation key.
' 04.17.04 - th2tran       - Rolled back changes from 02.25.04. The above implementation smalll breached in security:
'                            User can blank out the LastUsed property, set their clock back and continue to use an otherwise-expired
'                            license.
' 05.13.04 - th2tran       - Fixed "Object variable not set" error in IActiveLock_UsedDays() caused when retrieved Lic is Nothing.
' 07.11.04 - th2tran       - For regional setting in which Daylight Savings Time (DST) applies,
'                            There was a chance that ActiveLock could misbehave when DST goes into effect, due to
'                            its usage of local time in the LastUsed property tripping over when Daylights Savings Time (DST) goes into effect.
'                            To avoid this problem, all date-time sensitive values are now to be handled in UTC format.
'                           - Raise alerrNotImplemented for KeyStoreType other than alsFile
' 07.19.04 - th2tran       - Fixed bug: Bad InstallationCode was  generated when UserName="" and LockType is other than lockNone.
'                            This problem was because, previously, the vbLf delimiter wasn't included in the Installation Code when the
'                            user name is blank, which causes ALUGEN to (correctly) take the LockCode part of InstallationCode
'                            (the part after the last vbLf delimiter) for the user name, resulting in an invalid liberation key
'                            being generated.
' </pre>

'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit


' Implements the IActiveLock interface.
'
Implements IActiveLock

Private mSoftwareName As String
Private mSoftwareVer As String
Private mSoftwarePassword As String
Private mSoftwareCode As String
Private mLockTypes As ALLockTypes
Private mLicenseKeyTypes As ALLicenseKeyTypes
Private mUsedLockTypes() As ALLockTypes
Private mTrialType As Long
Private mTrialLength As Long
Private mRemainingTrialDays As Long
Private mRemainingTrialRuns As Long
Private mTrialHideTypes As ALTrialHideTypes
Private mKeyStore As IKeyStoreProvider
Private mKeyStorePath As String
Private MyNotifier As New ActiveLockEventNotifier
Private mLibKeyPath As String
Private mCheckTimeServerForClockTampering As ALTimeServerTypes
Private mCheckSystemFilesForClockTampering As ALSystemFilesTypes
Private mLicenseFileType As ALLicenseFileTypes
Private mAutoRegister As ALAutoRegisterTypes
Private mTrialWarning As ALTrialWarningTypes

' Registry hive used to store Activelock settings.
Private Const AL_REGISTRY_HIVE$ = "Software\ActiveLock Software Group\ActiveLock3"

' Transients
Private mfInit As Boolean   ' flag to indicate that ActiveLock has been initialized
'===============================================================================
' Name: Function AppendLockString
' Input:
'   ByRef strdata As String - String to be disected
' Output:
'   String - 64-byte lines
' Purpose: Breaks a long string into chunks of 64-byte lines.
' Remarks: None
'===============================================================================
Private Function Make64ByteChunks(strdata As String) As String
    Dim i As Long
    Dim Count As Long
    Count = Len(strdata)
    Dim sResult As String
    sResult = Left$(strdata, 64)
    i = 65
    While i <= Count
        sResult = sResult & vbCrLf & Mid$(strdata, i, 64)
        i = i + 64
    Wend
    Make64ByteChunks = sResult
End Function

'===============================================================================
' Name: Property Let CheckTimeServerForClockTampering
' Input:
'   ByVal RHS As ALTimeServerTypes - Flag being passed to check the time server
' Output: None
' Purpose: Specifies whether a Time Server should be used to check Clock Tampering
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_CheckTimeServerForClockTampering(ByVal RHS As ALTimeServerTypes)
    mCheckTimeServerForClockTampering = RHS
End Property
'===============================================================================
' Name: Property Let TrialWarning
' Input:
'   ByVal RHS As ALTrialWarningTypes - Trial Warning is either Persistent or Temporary
' Output: None
' Purpose: Specifies whether the Trial Warning is either Persistent or Temporary
' Remarks: None
'===============================================================================
Private Property Let TrialWarning(ByVal RHS As ALTrialWarningTypes)
    mTrialWarning = RHS
End Property

'===============================================================================
' Name: Property Let AutoRegister
' Input:
'   ByVal RHS As ALAutoRegisterTypes - Flag to indicate whether auto registration via an ALL file should be performed
' Output: None
' Purpose: Specifies whether auto registration via an ALL file should be performed
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_AutoRegister(ByVal RHS As ALAutoRegisterTypes)
    mAutoRegister = RHS
End Property

'===============================================================================
' Name: Property Let CheckSystemFilesForClockTampering
' Input:
'   ByVal RHS As ALSystemFilesTypes - Flag being passed to check the system files for clock tampering
' Output: None
' Purpose: Specifies whether the system files should be scanned to check Clock Tampering
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_CheckSystemFilesForClockTampering(ByVal RHS As ALSystemFilesTypes)
    mCheckSystemFilesForClockTampering = RHS
End Property

'===============================================================================
' Name: Property Let IActiveLock_LicenseFileType
' Input:
'   ByVal RHS As ALLicenseFileTypes - ALLicenseFileTypes type
' Output: None
' Purpose: Specifies the ALLicenseFileTypes type
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_LicenseFileType(ByVal RHS As ALLicenseFileTypes)
    mLicenseFileType = RHS
End Property

'===============================================================================
' Name: Property Let IActiveLock_LockType
' Input:
'   ByVal RHS As ALLockTypes - ALLockTypes type
' Output: None
' Purpose: Specifies the ALLockTypes type
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_LockType(ByVal RHS As ALLockTypes)
    mLockTypes = RHS
End Property

'===============================================================================
' Name: Property Let IActiveLock_LicenseKeyType
' Input:
'   ByVal RHS As ALLicenseKeyTypes - ALLicenseKeyTypes type
' Output: None
' Purpose: Specifies the ALLicenseKeyTypes type
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_LicenseKeyType(ByVal RHS As ALLicenseKeyTypes)
    mLicenseKeyTypes = RHS
End Property

'===============================================================================
' Name: Property Get IActiveLock_RegisteredLevel
' Input: None
' Output:
'   String - License RegisteredLevel
' Purpose: Gets the Registered Level for the license after validating it.
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_RegisteredLevel() As String
    Dim Lic As ProductLicense
    Set Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
    If Lic Is Nothing Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE
    End If
    ' Validate the License.
    ValidateLic Lic
    IActiveLock_RegisteredLevel = Lic.RegisteredLevel
End Property

Private Sub Class_Initialize()
    ' Default to alsFile
    IActiveLock_KeyStoreType = alsFile
End Sub

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
Private Property Let IActiveLock_AutoRegisterKeyPath(ByVal RHS As String)
    mLibKeyPath = RHS
End Property


'===============================================================================
' Name: Property Get AutoRegisterKeyPath
' Input: None
' Output: None
' Purpose: Sets the auto register file full path
' Remarks: None
'===============================================================================
Private Property Get AutoRegisterKeyPath() As String
    AutoRegisterKeyPath = mLibKeyPath
End Property

'===============================================================================
' Name: Property Get IActiveLock_EventNotifier
' Input: None
' Output: None
' Purpose: Gets a notification from Activelock
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_EventNotifier() As ActiveLockEventNotifier
    Set IActiveLock_EventNotifier = MyNotifier
End Property

'===============================================================================
' Name: Sub IActiveLock_Init
' Input:
'   ByRef autoLicString As String - Returned License Key of AutoRegister is successful
' Output: None
' Purpose: Initalizes Activelock
' Remarks: Performs CRC check on Alcrypto
'   <p>Performs auto license registration if the license file is found
'===============================================================================
Private Sub IActiveLock_Init(Optional autoLicString As String)
' If running in Debug mode, don't bother with dll authentication
#If AL_DEBUG = 1 Then
    GoTo Done
#End If
    ' Checksum ALCrypto3.dll
    'Const ALCRYPTO_MD5$ = "a86c30" & "57d" & "c51ca0cb8204b0ce493c5b4"
    Const ALCRYPTO_MD5$ = "D052CE549A594CC7F2AEC1AE50CDD947"
    Dim strdata As String, strMD5 As String
    Call modActiveLock.ReadFile(WinSysDir() & "\alcrypto3.dll", strdata)
    strMD5 = UCase(modMD5.Hash(strdata)) ' use our own MD5 hashing routine instead of ALCrypto's md5_hash() function.
'Debug.Print "ALCrypto Hash: " & strMD5
'Debug.Print "strdata: " & strdata
'MsgBox strMD5
    If strMD5 <> ALCRYPTO_MD5 Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrFileTampered, ACTIVELOCKSTRING, STRFILETAMPERED
    End If
    ' Perform automatic license registration
    If AutoRegisterKeyPath <> "" And mAutoRegister = alsEnableAutoRegistration Then
        DoAutoRegistration autoLicString
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
Private Sub DoAutoRegistration(strLibKey As String)
    
    ' Don't bother to proceed unless the file is there.
    If Not FileExists(AutoRegisterKeyPath) Then Exit Sub
    
    ReadLibKey AutoRegisterKeyPath, strLibKey
    IActiveLock_Register strLibKey
    
    ' If registration is successful, delete the liberation file so we won't register the same file on next startup
    Kill AutoRegisterKeyPath
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
    Dim hFile As Long
    hFile = FreeFile
    Open sFileName For Input As #hFile
        On Error GoTo finally
        strLibKey = Input$(LOF(hFile), hFile)
finally:
    Close #hFile
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
Private Sub IActiveLock_Acquire(Optional strMsg As String)
    Dim adsText As String
    Dim strStream As String
    Dim Lic As ProductLicense
    Dim trialStatus As Boolean
    
    Dim dt1 As Date
    Dim dt2 As Date
    Dim fileno As Integer
    Dim run As String

    ' This sample code is for debugging purposes
    ' For debugging via DBGVIEW.EXE, etc.
    ' Use calls like this in every location you want to debug
    OutputDebugString "Entered the IActiveLock_Acquire function"
    
    strStream = mSoftwareName & mSoftwareVer & mSoftwarePassword
    
    ' Get the current date format and save it to regionalSymbol variable
    Get_locale
    ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
    Set_locale ""

    'Check the Key Store Provider
    If mKeyStore Is Nothing Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrKeyStoreInvalid, ACTIVELOCKSTRING, STRKEYSTOREUNINITIALIZED
    'Check the Key Store Path (LIC file path)
    ElseIf mKeyStorePath = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrKeyStorePathInvalid, ACTIVELOCKSTRING, STRKEYSTOREPATHISEMPTY
    ElseIf mSoftwareName = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoSoftwareName, ACTIVELOCKSTRING, STRNOSOFTWARENAME
    ElseIf mSoftwareVer = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoSoftwareVersion, ACTIVELOCKSTRING, STRNOSOFTWAREVERSION
    ElseIf mSoftwarePassword = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoSoftwarePassword, ACTIVELOCKSTRING, STRNOSOFTWAREPASSWORD
    End If
    
    Set Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
    If Lic Is Nothing Then
        
        ' There's no valid license, so let's see if we can grant this user a "Trial License"
        If mTrialType = ALTrialTypes.trialNone Then 'No Trial
            GoTo noRegistration
        End If
            
        On Error GoTo noRegistration
        strMsg = ""
        If mTrialHideTypes = 0 Then
            mTrialHideTypes = trialHiddenFolder Or trialRegistryPerUser Or trialSteganography
        End If
        
        trialStatus = ActivateTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, strMsg, mSoftwarePassword, mCheckTimeServerForClockTampering, mCheckSystemFilesForClockTampering, mTrialWarning, mRemainingTrialDays, mRemainingTrialRuns)
        
        ' Set the locale date format to what we had before; can't leave changed
        Set_locale regionalSymbol
        If trialStatus = True Then
            Exit Sub
        End If
        GoTo continueRegistration

noRegistration:
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE
    
    Else  'Lic exists therefore we'll check the LIC file ADS
        
        ' Check clock tampering for non-permanent licenses.
        If Lic.LicenseType <> allicPermanent Then
            If mCheckTimeServerForClockTampering = alsCheckTimeServer Then
                If SystemClockTampered Then
                    Set_locale regionalSymbol
                    Err.Raise ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED
                End If
            End If
            If mCheckSystemFilesForClockTampering = alsCheckSystemFiles Then
                If ClockTampering Then
                    Set_locale regionalSymbol
                    Err.Raise ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED
                End If
            End If
        End If
        
        If CheckStreamCapability() And Lic.LicenseType <> allicPermanent Then

            If FileLen(mKeyStorePath) = 0 Then GoTo continueRegistration
            adsText = ViewStream(mKeyStorePath & ":" & strStream) ' View the stream
            If adsText = "" Then
                Set_locale regionalSymbol
                Err.Raise ActiveLockErrCodeConstants.alerrLicenseTampered, ACTIVELOCKSTRING, STRLICENSETAMPERED
            End If
            dt1 = ActiveLockDate(CDate(adsText))
            dt2 = ActiveLockDate(UTC(Now))
            If dt2 < dt1 Then
                Set_locale regionalSymbol
                Err.Raise ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED
            End If
            ' Write the stream to the file
            fileno = FreeFile
            run = App.path & "\run.bat"
            If fileExist(run) Then Kill run
            Open run For Append As fileno
            Print #fileno, "@ECHO OFF"
            Print #fileno, "echo " & ActiveLockDate(UTC(Now)) & " > " & """" & mKeyStorePath & ":" & strStream & """"
            Close fileno
            Dim rc As Double
            rc = Shell(run, vbMinimizedFocus)
            DoEvents
            
            GoTo continueRegistration
        End If
    
    End If

continueRegistration:
    ' Validate license
    Set_locale regionalSymbol
    ValidateLic Lic
    DoEvents
    
End Sub
'===============================================================================
' Name: Sub IActiveLock_AddLockCode
' Input:
'   ByVal LockType As ALLockTypes _ to be added to array
'   ByRef Byref LockTypes() As ALLockTypes - array of used LockTypes being built up
'   ByRef SizeLT as Integer - size of array of used LockTypes being built up
' Output:
' Purpose: Helper function to build up array of used LockType s
' Remarks:
'===============================================================================
Private Sub IActiveLock_AddLockCode(ByVal LockType As ALLockTypes, ByRef SizeLT As Integer)
  ReDim Preserve mUsedLockTypes(SizeLT)
  mUsedLockTypes(SizeLT) = LockType
  SizeLT = SizeLT + 1
End Sub

'===============================================================================
' Name: Property Get IActiveLock_RegisteredDate
' Input: None
' Output:
'   String - License registration date.
' Purpose: Gets the license registration date after validating it.
' Remarks: This is the date the license was generated by Alugen. NOT the date the license was activated.
'===============================================================================
Private Property Get IActiveLock_RegisteredDate() As String
    Dim Lic As ProductLicense
    Set Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
    If Lic Is Nothing Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE
    End If
    ' Validate the License.
    ValidateLic Lic
    IActiveLock_RegisteredDate = Lic.RegisteredDate
End Property

'===============================================================================
' Name: Property Get IActiveLock_RegisteredUser
' Input: None
' Output:
'   String - Registered user name
' Purpose: Gets the registered user name after validating the license
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_RegisteredUser() As String
    Dim Lic As ProductLicense
    Set Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
    If Lic Is Nothing Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE
    End If
    ' Validate the License.
    ValidateLic Lic
    IActiveLock_RegisteredUser = Lic.Licensee
End Property



'===============================================================================
' Name: Sub ValidateKey
' Input:
'   Lic As ProductLicense - Product license
' Output: None
' Purpose: Validates the License Key using RSA signature verification.
'   <p>License key contains the RSA signature of IActiveLock_LockCode.
' Remarks: None
'===============================================================================
Private Sub ValidateKey(Lic As ProductLicense)
    ' make sure software code is set
    If mSoftwareCode = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNotInitialized, ACTIVELOCKSTRING, STRNOSOFTWARECODE
    End If

    Dim strPubKey As String
    Dim strSig As String
    Dim strLic As String
    Dim strLicKey As String
    
    strPubKey = mSoftwareCode
    strLic = IActiveLock_LockCode(Lic)
    strLicKey = Lic.LicenseKey

    ' decode the license key
    strSig = Base64_Decode(strLicKey)
    
    If Left(strPubKey, 3) <> "RSA" Then 'ALCrypto
        ' Print out some info for debugging purposes
        Debug.Print "Code1: " & strPubKey
        Debug.Print "Lic: " & strLic
        Debug.Print "Lic hash: " & modMD5.Hash(strLic)
        Debug.Print "LicKey: " & strLicKey
        Debug.Print "Sig: " & strSig
        Debug.Print "Verify: " & RSAVerify(strPubKey, strLic, strSig)
        Debug.Print "===================================================="
        
        ' validate the key
        Dim rc As Long
        rc = RSAVerify(strPubKey, strLic, strSig)
        If rc <> 0 Then
            Set_locale regionalSymbol
            Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
        End If
    Else    'CryptoAPI
        
        Dim strPublicBlob As String, modulus As Long
        Dim MyGlobals As Globals
        Dim ok As String
        Set MyGlobals = New Globals
        ' Validate the key with CryptoAPI
        If Left(strPubKey, 6) = "RSA512" Then
            strPublicBlob = Right(strPubKey, Len(strPubKey) - 6)
            modulus = CInt(Mid(strPubKey, 4, 3))
        Else
            strPublicBlob = Right(strPubKey, Len(strPubKey) - 7)
            modulus = CInt(Mid(strPubKey, 4, 4))
        End If
        ok = MyGlobals.CryptoAPIAction(5, mSoftwareName & mSoftwareVer, strLic, strSig, strPublicBlob, "", modulus)
    End If

    ' Check if license has not expired
    ' but don't do it if there's no expiration date
    If Lic.Expiration = "" Then Exit Sub
    Dim dtExp As Date
    dtExp = ActiveLockDate(CDate(Lic.Expiration))
    If ActiveLockDate(UTC(Now)) > dtExp And Lic.LicenseType <> allicPermanent Then
        ' ialkan - 9-23-2005 added the following to update and store the license
        ' with the new LastUsed property; otherwise setting the clock back next time
        ' might bypass the protection mechanism
        ' Update last used date
        UpdateLastUsed Lic
        mKeyStore.Store Lic, mLicenseFileType
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrLicenseExpired, ACTIVELOCKSTRING, STRLICENSEEXPIRED
    End If
End Sub

'===============================================================================
' Name: Sub ValidateShortKey
' Input:
'   Lic As ProductLicense - Product license
' Output: None
' Purpose: Validates the License Key using the Short Key MD5 verification.
' Remarks: None
'===============================================================================
Private Sub ValidateShortKey(Lic As ProductLicense, ByVal user As String)
    
    Dim oReg As clsShortSerial
    Dim m_Key As clsShortLicenseKey
    Dim sKey As String
    Dim m_ProdCode As Long
    Dim SerialNumber As String
    Dim ExpireDate As Date
    Dim UserData As Integer
    Dim RegisteredLevel As Long
    
    ' make sure software code is set
    If mSoftwareCode = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNotInitialized, ACTIVELOCKSTRING, STRNOSOFTWARECODE
    End If

    'This is a short key
    Set m_Key = New clsShortLicenseKey
    
    m_Key.AddSwapBits 0, 0, 1, 0
    m_Key.AddSwapBits 0, 2, 1, 1
    m_Key.AddSwapBits 0, 4, 2, 0
    m_Key.AddSwapBits 0, 5, 2, 1
    m_Key.AddSwapBits 2, 0, 3, 0
    m_Key.AddSwapBits 2, 6, 3, 1
    m_Key.AddSwapBits 2, 7, 1, 3
    
    Set oReg = New clsShortSerial
    sKey = oReg.GenerateKey("", Left(mSoftwareCode, Len(mSoftwareCode) - 2)) 'Do not include the last 2 possible == paddings
    m_ProdCode = CLng(Left(sKey, 4))
    
    SerialNumber = oReg.GenerateKey(mSoftwareName & mSoftwareVer & mSoftwarePassword, modHardware.GetHDSerialFirmware())

    ' verify the key is valid
    If m_Key.ValidateShortKey(Lic.LicenseKey, SerialNumber, _
        user, m_ProdCode, ExpireDate, UserData, RegisteredLevel) = True Then
            
            ' After the key is disassembled it fills the output
            ' variables with expire date and license counter.
            Lic.LicenseType = CLng(CStr(modActiveLock.HiByte(UserData)))
            Lic.ProductName = mSoftwareName
            Lic.ProductVer = mSoftwareVer
            Lic.LicenseClass = LicFlags.alfSingle 'Multi User License will be available with network version
            Lic.Licensee = user
            If Lic.RegisteredLevel = "" Then
                Lic.RegisteredLevel = CStr(RegisteredLevel)
            ElseIf Lic.RegisteredLevel <> CStr(RegisteredLevel) Then
                Set_locale regionalSymbol
                Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
            End If
           If Lic.RegisteredDate = "" Then
                Lic.RegisteredDate = CStr(ActiveLockDate(UTC(Now)))
            End If
         ' ignore expiration date if license type is "permanent"
            If Lic.LicenseType <> allicPermanent Then
                Lic.Expiration = CStr(ActiveLockDate(ExpireDate))
            End If
            Lic.MaxCount = CLng(CStr(modActiveLock.LoByte(UserData)))
            ' Finally check if the serial number is Ok
            If Not oReg.IsKeyOK(SerialNumber, mSoftwareName & mSoftwareVer & mSoftwarePassword, modHardware.GetHDSerialFirmware()) Then
                ' Something wrong with the serial number used
                Set_locale regionalSymbol
                Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
            End If
            Lic.LicenseCode = "Short Key"
            '"Key is valid."
    Else
         'MsgBox "Invalid license key."
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
    End If

    ' Check if license has not expired
    ' but don't do it if there's no expiration date
    If Lic.Expiration = "" Then Exit Sub
    Dim dtExp As Date
    dtExp = ActiveLockDate(CDate(Lic.Expiration))
    If ActiveLockDate(UTC(Now)) > dtExp And Lic.LicenseType <> allicPermanent Then
        ' ialkan - 9-23-2005 added the following to update and store the license
        ' with the new LastUsed property; otherwise setting the clock back next time
        ' might bypass the protection mechanism
        ' Update last used date
        UpdateLastUsed Lic
        mKeyStore.Store Lic, mLicenseFileType
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrLicenseExpired, ACTIVELOCKSTRING, STRLICENSEEXPIRED
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
Private Sub ValidateLic(Lic As ProductLicense)
    
    ' Get the current date format and save it to regionalSymbol variable
    Get_locale
    ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
    Set_locale ""
    
    ' make sure we're initialized.
    If Not mfInit Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNotInitialized, ACTIVELOCKSTRING, STRNOTINITIALIZED
    End If

    ' validate license key first
    If Mid(Lic.LicenseKey, 5, 1) = "-" And Mid(Lic.LicenseKey, 10, 1) = "-" And Mid(Lic.LicenseKey, 15, 1) = "-" And Mid(Lic.LicenseKey, 20, 1) = "-" Then
        ValidateShortKey Lic, Lic.Licensee
    Else 'ALCrypto RSA key
        ValidateKey Lic
    End If
    
    Dim strEncrypted As String, strHash As String
    ' Validate last run date
    strEncrypted = Lic.LastUsed
    MyNotifier.Notify "ValidateValue", strEncrypted
    strHash = modMD5.Hash(strEncrypted)
    If strHash <> Lic.Hash1 Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrLicenseTampered, ACTIVELOCKSTRING, STRLICENSETAMPERED
    End If
    ' try to detect the user setting their system clock back
    ' Need to account for Daylight Savings Time
    Dim strNow As String
    ' Normalize to the format of the saved date-time, before we compare
    strNow = ActiveLockDate(UTC(Now))
    'If strNow < Lic.LastUsed And Lic.LicenseType <> allicPermanent Then
    If strNow < ActiveLockDate(CDate(Lic.LastUsed)) And Lic.LicenseType <> allicPermanent Then
        'MsgBox "UTC Now: " & strNow
        'MsgBox "LastUsed: " & CDate(Lic.LastUsed)
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED
    End If
    UpdateLastUsed Lic
    mKeyStore.Store Lic, mLicenseFileType
    Set_locale regionalSymbol
End Sub


'===============================================================================
' Name: Sub UpdateLastUsed
' Input:
'   Lic As ProductLicense - Product License
' Output: None
' Purpose: Updates LastUsed property with current date stamp.
' Remarks: None
'===============================================================================
Private Sub UpdateLastUsed(Lic As ProductLicense)
    ' Update license store with LastRunDate
    Dim strLastUsed As String
    Set_locale ""
    strLastUsed = ActiveLockDate(UTC(Now))
    strLastUsed = ActiveLockDate(CDate(strLastUsed))
    Lic.LastUsed = strLastUsed
    MyNotifier.Notify "ValidateValue", strLastUsed
    Lic.Hash1 = modMD5.Hash(strLastUsed)
End Sub

'===============================================================================
' Name: Property Get IActiveLock_ExpirationDate
' Input: None
' Output:
'   String - Expiration date of the license
' Purpose: Returns the expiration date of the license after validating it
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_ExpirationDate() As String
    Dim Lic As ProductLicense
    Set Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
    If Lic Is Nothing Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE
    End If
    ' Validate the License.
    ValidateLic Lic
    IActiveLock_ExpirationDate = Lic.Expiration
End Property

'===============================================================================
' Name: Property Let IActiveLock_KeyStorePath
' Input:
'   ByVal RHS As String - License file path name
' Output: None
' Purpose: Specifies the license file path name
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_KeyStorePath(ByVal RHS As String)
    If Not mKeyStore Is Nothing Then
        mKeyStore.KeyStorePath = RHS
    End If
    mKeyStorePath = RHS
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
Private Property Let IActiveLock_KeyStoreType(ByVal RHS As LicStoreType)
    ' Instantiate Key Store Provider
    If RHS = alsFile Then
        Set mKeyStore = New FileKeyStoreProvider
    Else
'        Set mKeyStore = New RegistryKeyStoreProvider
        ' TODO: Implement me!
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNotImplemented, ACTIVELOCKSTRING, STRNOTIMPLEMENTED
    End If
    ' Set Key Store Path in KeyStoreProvider
    If mKeyStorePath <> "" Then
        mKeyStore.KeyStorePath = mKeyStorePath
    End If
End Property

'===============================================================================
' Name: Property Let IActiveLock_TrialHideType
' Input:
'   ByVal RHS As ALTrialHideTypes - ALTrialHideTypes type
' Output: None
' Purpose: Specifies the ALTrialHideTypes type
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_TrialHideType(ByVal RHS As ALTrialHideTypes)
    mTrialHideTypes = RHS
End Property

'===============================================================================
' Name: Property Get IActiveLock_LockType
' Input: None
' Output:
'   ALLockTypes - Lock types type
' Purpose: Gets the ALLockTypes type
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_LockType() As ALLockTypes
    IActiveLock_LockType = mLockTypes
End Property

'===============================================================================
' Name: Property Let TrialWarning
' Input:
'   ByVal Value As ALTrialWarningTypes - Trial Warning is either Persistent or Temporary
' Output: None
' Purpose: Specifies whether the Trial Warning is either Persistent or Temporary
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_TrialWarning(ByVal RHS As ALTrialWarningTypes)
    mTrialWarning = RHS
End Property

'===============================================================================
' Name: Property Get IActiveLock_UsedLockType
' Input: None
' Output:
'   ALLockTypes - Used Lock types type
' Purpose: Gets the ALLockTypes type
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_UsedLockType() As ALLockTypes()
    IActiveLock_UsedLockType = mUsedLockTypes
End Property

'===============================================================================
' Name: Property Get IActiveLock_TrialHideType
' Input: None
' Output:
'   ALTrialHideTypes - Trial Hide types type
' Purpose: Gets the ALTrialHideTypes type
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_TrialHideType() As ALTrialHideTypes
    IActiveLock_TrialHideType = mTrialHideTypes
End Property

'===============================================================================
' Name: Sub IActiveLock_Register
' Input:
'   ByVal LibKey As String - Liberation Key
' Output: None
' Purpose: Registers Activelock license with a given liberation key
' Remarks: None
'===============================================================================
Private Sub IActiveLock_Register(ByVal LibKey As String, Optional user As String = "")
    
    Dim Lic As New ActiveLock3.ProductLicense
    Dim trialStatus As Boolean
    Dim ok As Integer
    Dim strStream As String
    Dim fileno As Integer
    Dim run As String
    Dim rc As Double
    
    ' Get the current date format and save it to regionalSymbol variable
    Get_locale
    ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
    Set_locale ""
    
    ' Check to see if this is a Short License Key
    If Mid(LibKey, 5, 1) = "-" And Mid(LibKey, 10, 1) = "-" And Mid(LibKey, 15, 1) = "-" And Mid(LibKey, 20, 1) = "-" Then
        Lic.LicenseKey = UCase(LibKey)
        ValidateShortKey Lic, user
    Else 'ALCrypto RSA key
        Lic.Load LibKey
        ' Validate that the license key.
        '   - registered user
        '   - expiry date
        ValidateKey Lic
    End If
    
    ' License was validated successfuly. Check clock tampering for non-permanent licenses.
    If Lic.LicenseType <> allicPermanent Then
        If mCheckTimeServerForClockTampering = alsCheckTimeServer Then
            If SystemClockTampered Then
                Set_locale regionalSymbol
                Err.Raise ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED
            End If
        End If
        If mCheckSystemFilesForClockTampering = alsCheckSystemFiles Then
            If ClockTampering Then
                Set_locale regionalSymbol
                Err.Raise ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED
            End If
        End If
    End If
    
    ' License was validated successfuly.  Store it.
    If mKeyStore Is Nothing Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrKeyStoreInvalid, ACTIVELOCKSTRING, STRKEYSTOREUNINITIALIZED
    End If

    ' Update last used date
    UpdateLastUsed Lic
    mKeyStore.Store Lic, mLicenseFileType
        
    ' This works under NTFS and is needed to prevent clock tampering
    If CheckStreamCapability() And Lic.LicenseType <> allicPermanent Then
        ' Write the current date and time into the ads
        strStream = mSoftwareName & mSoftwareVer & mSoftwarePassword
        ' Write the stream to the file
        fileno = FreeFile
        run = App.path & "\run.bat"
        If fileExist(run) Then Kill run
        Open run For Append As fileno
        Print #fileno, "@ECHO OFF"
        Print #fileno, "echo " & CStr(ActiveLockDate(UTC(Now))) & " > " & """" & mKeyStorePath & ":" & strStream & """"
        Close fileno
        rc = Shell(run, vbMinimizedFocus)
        DoEvents
    
    End If
   
    ' Expire all trial licenses
    On Error Resume Next
    ' Expire the Trial
    If mTrialType <> ALTrialTypes.trialNone Then
        trialStatus = ExpireTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, mSoftwarePassword)
    End If
    Set_locale regionalSymbol

End Sub

'===============================================================================
' Name: Sub IActiveLock_KillTrial
' Input: None
' Output: None
' Purpose: Kills a Trial License
' Remarks: None
'===============================================================================
Private Sub IActiveLock_KillTrial()
    On Error Resume Next
    'Expire the Trial
    Dim trialStatus As Boolean
    If mSoftwareName = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoSoftwareName, ACTIVELOCKSTRING, STRNOSOFTWARENAME
    ElseIf mSoftwareVer = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoSoftwareVersion, ACTIVELOCKSTRING, STRNOSOFTWAREVERSION
    ElseIf mSoftwarePassword = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoSoftwarePassword, ACTIVELOCKSTRING, STRNOSOFTWAREPASSWORD
    Else
        trialStatus = ExpireTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, mSoftwarePassword)
    End If
End Sub
Private Function IActiveLock_GenerateShortKey(ByVal SoftwareCode As String, ByVal SerialNumber As String, ByVal LicenseeAndRegisteredLevel As String, ByVal Expiration As String, ByVal LicType As ActiveLock3.ALLicType, ByVal RegisteredLevel As Long, Optional ByVal MaxUsers As Integer = 1) As String

    On Error GoTo ErrHandler
    
    Dim m_Key As clsShortLicenseKey
    Set m_Key = New clsShortLicenseKey
    
    m_Key.AddSwapBits 0, 0, 1, 0
    m_Key.AddSwapBits 0, 2, 1, 1
    m_Key.AddSwapBits 0, 4, 2, 0
    m_Key.AddSwapBits 0, 5, 2, 1
    m_Key.AddSwapBits 2, 0, 3, 0
    m_Key.AddSwapBits 2, 6, 3, 1
    m_Key.AddSwapBits 2, 7, 1, 3

    Dim oReg As clsShortSerial
    Set oReg = New clsShortSerial
    Dim sKey As String
    Dim m_ProdCode As Long
    
    sKey = oReg.GenerateKey("", Left(SoftwareCode, Len(SoftwareCode) - 2)) 'Do not include the last 2 possible == paddings
    m_ProdCode = CLng(Left(sKey, 4))
    
    ' create a new key
    IActiveLock_GenerateShortKey = m_Key.CreateShortKey(SerialNumber, _
        LicenseeAndRegisteredLevel, m_ProdCode, Expiration, _
        MakeWord(CStr(MaxUsers), CStr(LicType)), RegisteredLevel)

    Exit Function
ErrHandler:
    Set oReg = Nothing
    Set m_Key = Nothing
    
End Function

'*******************************************************************************
' Sub GenerateShortSerial
'
' Input:
' appNameVersionPassword
' HDDfirmwareSerial
'
' DESCRIPTION:
' Generates a Short Key (Serial Number)
'*******************************************************************************
Private Function IActivelock_GenerateShortSerial(ByVal HDDfirmwareSerial As String) As String
    Dim oReg As clsShortSerial
    Dim sKey As String
    
    Set oReg = New clsShortSerial
    sKey = oReg.GenerateKey(mSoftwareName & mSoftwareVer & mSoftwarePassword, HDDfirmwareSerial)
    IActivelock_GenerateShortSerial = sKey
    ' If longer serial is used, possible to break up into sections
    'Left(sKey, 4) & "-" & Mid(sKey, 5, 4) & "-" & Mid(sKey, 9, 4) & "-" & Mid(sKey, 13, 4)
    
    Set oReg = Nothing
End Function


'===============================================================================
' Name: Sub IActiveLock_ResetTrial
' Input: None
' Output: None
' Purpose: Resets a Trial License
' Remarks: None
'===============================================================================
Private Sub IActiveLock_ResetTrial()
    On Error Resume Next
    'Reset the Trial
    Dim trialStatus As Boolean
    If mSoftwareName = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoSoftwareName, ACTIVELOCKSTRING, STRNOSOFTWARENAME
    ElseIf mSoftwareVer = "" Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoSoftwareVersion, ACTIVELOCKSTRING, STRNOSOFTWAREVERSION
    Else
        trialStatus = ResetTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, mSoftwarePassword)
    End If
End Sub

'===============================================================================
' Name: Function RandomNumber
' Input: None
' Output:
'   String - Random number generated
' Purpose: This function generates a random number
' Remarks: None
'===============================================================================
Private Function RandomNumber() As String
    On Error Resume Next
    Randomize
    RandomNumber = CStr(CLng(Rnd(1) * 2147483647))
End Function
'===============================================================================
' Name: Property Let IActiveLock_SoftwareName
' Input:
'   ByVal RHS As String - Software name for the license
' Output: None
' Purpose: Specifies the SoftwareName for the license
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_SoftwareName(ByVal RHS As String)
    mSoftwareName = RHS
End Property

'===============================================================================
' Name: Property Let IActiveLock_SoftwarePassword
' Input:
'   ByVal RHS As String - Software Password for the license
' Output: None
' Purpose: Specifies the SoftwarePassword for the license
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_SoftwarePassword(ByVal RHS As String)
    mSoftwarePassword = RHS
End Property

'===============================================================================
' Name: Property Let IActiveLock_TrialType
' Input:
'   ByVal RHS As ALTrialTypes - Trial Type for the license
' Output: None
' Purpose: Specifies the TrialType for the license
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_TrialType(ByVal RHS As ALTrialTypes)
    mTrialType = RHS
End Property
'===============================================================================
' Name: Property Let IActiveLock_TrialLength
' Input:
'   ByVal RHS As Long - Trial Length for the license
' Output: None
' Purpose: Specifies the TrialLength for the license
' Remarks: The length specified corresponds to the TrialType selected
'===============================================================================
Private Property Let IActiveLock_TrialLength(ByVal RHS As Long)
    mTrialLength = RHS
End Property
'===============================================================================
' Name: Property Get IActiveLock_SoftwareName
' Input: None
' Output:
'   String - Software name  for the license
' Purpose: Gets the SoftwareName for the license
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_SoftwareName() As String
    IActiveLock_SoftwareName = mSoftwareName
End Property
'===============================================================================
' Name: Property Get IActiveLock_MaxCount
' Input: None
' Output:
'   Long - Number of concurrent users for the networked license
' Purpose: Gets the Number of concurrent users for the networked license
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_MaxCount() As Long
    Dim Lic As ProductLicense
    Set Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
    If Lic Is Nothing Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE
    End If
    ' Validate the License.
    ValidateLic Lic
    IActiveLock_MaxCount = Lic.MaxCount
End Property

'===============================================================================
' Name: Property Get IActiveLock_LicenseClass
' Input: None
' Output:
'   String - LicenseClass
' Purpose: Gets the LicenseClass for the license
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_LicenseClass() As String
    Dim Lic As ProductLicense
    Set Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
    If Lic Is Nothing Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE
    End If
    ' Validate the License.
    ValidateLic Lic
    IActiveLock_LicenseClass = Lic.LicenseClass
End Property

'===============================================================================
' Name: Property Get IActiveLock_SoftwarePassword
' Input: None
' Output:
'   String - Software Password for the license
' Purpose: Gets the SoftwarePassword for the license
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_SoftwarePassword() As String
    IActiveLock_SoftwarePassword = mSoftwarePassword
End Property

'===============================================================================
' Name: Property Get IActiveLock_TrialType
' Input: None
' Output:
'   ALTrialTypes - Trial Type  for the license
' Purpose: Gets the TrialType for the license
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_TrialType() As ALTrialTypes
    IActiveLock_TrialType = mTrialType
End Property
'===============================================================================
' Name: Property Get IActiveLock_TrialLength
' Input: None
' Output:
'   Integer - Trial Length  for the license
' Purpose: Gets the TrialLength for the license
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_TrialLength() As Long
    IActiveLock_TrialLength = mTrialLength
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
Private Property Get IActiveLock_InstallationCode(Optional ByVal user As String = "", Optional Lic As ProductLicense = Nothing) As String
    
    'Before we generate the installation code, let's check if this app is using a short key
    If mLicenseKeyTypes = alsShortKeyMD5 Then
        IActiveLock_InstallationCode = IActivelock_GenerateShortSerial(modHardware.GetHDSerialFirmware())
    
    ElseIf mLicenseKeyTypes = alsRSA Then

        ' Generate Request code to Lock
        Dim strReq As String, strLock  As String
        
        'Restrict user name to 2000 characters; need more? why?
        If Len(user) > 2000 Then
            Set_locale regionalSymbol
            Err.Raise ActiveLockErrCodeConstants.alerrUserNameTooLong, ACTIVELOCKSTRING, STRUSERNAMETOOLONG
        End If
        
        ' New in v3.1
        ' Version 3.1 and above of Activelock will append the "+" sign
        ' in front of the installation code whenever lockNone is used or
        ' lockType is not specified in the protected app.
        ' When "+" is not found at the beginning of the installation code,
        ' Alugen will not allow users pick the hardware lock method since this
        ' corresponds to an installation code which
        ' utilizes a hardware lock option specified inside the protected app.
        If mLockTypes = lockNone Then
            strLock = "+" & IActiveLock_LockCode()
        Else
            strLock = IActiveLock_LockCode()
        End If
        
        ' combine with user name
        strReq = strLock & vbLf & user
        
        ' base-64 encode the request
        Dim strReq2 As String
        strReq2 = modBase64.Base64_Encode(strReq)
        IActiveLock_InstallationCode = strReq2
    
        ' New in v3.1
        ' If there's a license and the LicenseCode exists, then use it
        ' LicenseCode is actually the Installation Code modified by Alugen
        ' LicenseCode is appended to the end of the lic file so that we can know
        ' Alugen specified the hardware keys, and LockType
        ' was not specified inside the protected app
        If Not Lic Is Nothing Then
            If Lic.LicenseCode <> "" Then
                IActiveLock_InstallationCode = Lic.LicenseCode
                If Left(IActiveLock_InstallationCode, 1) = "+" Then IActiveLock_InstallationCode = Mid(IActiveLock_InstallationCode, 2)
            ' We won't do the following in order to maintain backwards compatibility with existing licenses
            'ElseIf Lic.LicenseCode = "" And mLockTypes = lockNone Then
            '    Set_locale regionalSymbol
            '    Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
            End If
        End If
    
    End If
    
End Property

'===============================================================================
' Name: Function IActiveLock_LockCode
' Input:
'   ByRef Lic As ProductLicense - Product License
' Output:
'   String - Lock code
' Purpose: Returns the lock code from a given Activelock license
' Remarks: v3 includes the new lockHDFirmware option
'===============================================================================
Private Function IActiveLock_LockCode(Optional Lic As ProductLicense = Nothing) As String
    Dim strLock As String
    Dim noKey As String
    Dim userFromInstallCode As String, usedcode As String
    Dim a() As String, aString As String
    Dim j As Integer
    
    noKey = Chr(110) & Chr(111) & Chr(107) & Chr(101) & Chr(121)
    If Lic Is Nothing Then
    ' New in v3.1
    ' Modified this function on 1-13-2006 to append ALL hardware keys
    ' to the Installation Code. This way, it will be decided in Alugen which
    ' hardware keys will be used to lock the license to
    ' If there's already a lock selected in the protected app,
    ' such as lockHDfirmware or lockComputer, then Alugen will show
    ' these two options and will gray them out (fix these two selections)
    If mLockTypes = lockNone Then
        strLock = ""
        AppendLockString strLock, modHardware.GetMACAddress()
        AppendLockString strLock, modHardware.GetComputerName()
        AppendLockString strLock, modHardware.GetHDSerial()
        AppendLockString strLock, modHardware.GetHDSerialFirmware()
        AppendLockString strLock, modHardware.GetWindowsSerial()
        AppendLockString strLock, modHardware.GeetBiosVersion()
        AppendLockString strLock, modHardware.GetMotherboardSerial()
        AppendLockString strLock, modHardware.GetIPaddress()
    Else
        If mLockTypes And lockMAC Then
            AppendLockString strLock, modHardware.GetMACAddress()
        Else
            AppendLockString strLock, noKey
        End If
        If mLockTypes And lockComp Then
            AppendLockString strLock, modHardware.GetComputerName()
        Else
            AppendLockString strLock, noKey
        End If
        If mLockTypes And lockHD Then
            AppendLockString strLock, modHardware.GetHDSerial()
        Else
            AppendLockString strLock, noKey
        End If
        If mLockTypes And lockHDFirmware Then
            AppendLockString strLock, modHardware.GetHDSerialFirmware()
        Else
            AppendLockString strLock, noKey
        End If
        If mLockTypes And lockWindows Then
            AppendLockString strLock, modHardware.GetWindowsSerial()
        Else
            AppendLockString strLock, noKey
        End If
        If mLockTypes And lockBIOS Then
            AppendLockString strLock, modHardware.GeetBiosVersion()
        Else
            AppendLockString strLock, noKey
        End If
        If mLockTypes And lockMotherboard Then
            AppendLockString strLock, modHardware.GetMotherboardSerial()
        Else
            AppendLockString strLock, noKey
        End If
        If mLockTypes And lockIP Then
            AppendLockString strLock, modHardware.GetIPaddress()
        Else
            AppendLockString strLock, noKey
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
        
        ReDim mUsedLockTypes(0)  ' remove all previous
        Dim SizeLockType As Integer  ' use to build up LockCode.
        SizeLockType = 0
    
        If Lic.LicenseCode <> "" Then
            If Left(Lic.LicenseCode, 1) = "+" Then
                usedcode = modBase64.Base64_Decode(Mid(Lic.LicenseCode, 2))
                'bLockNone = True ' per David Weatherall
                IActiveLock_AddLockCode lockNone, SizeLockType   'dw1 build up lockTypes - start with lockNone
            Else
                usedcode = modBase64.Base64_Decode(Lic.LicenseCode)
                'bLockNone = False ' per David Weatherall
            End If
            
            a = Split(usedcode, vbLf)
            For j = LBound(a) To UBound(a) - 1
                aString = a(j)
                If Left(aString, 1) = "+" Then aString = Mid(aString, 2)
                If j = LBound(a) Then
                    If aString <> noKey Then
                        IActiveLock_AddLockCode lockMAC, SizeLockType
                        If aString <> modHardware.GetMACAddress() Then
                            Set_locale regionalSymbol
                            Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
                        End If
                    End If
                ElseIf j = LBound(a) + 1 Then
                    If aString <> noKey Then
                        IActiveLock_AddLockCode lockComp, SizeLockType 'dw1 build up lockTypes
                        If aString <> modHardware.GetComputerName() Then
                            Set_locale regionalSymbol
                            Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
                        End If
                    End If
                ElseIf j = LBound(a) + 2 Then
                    If aString <> noKey Then
                        IActiveLock_AddLockCode lockHD, SizeLockType  'dw1 build up lockTypes
                        If aString <> modHardware.GetHDSerial() Then
                            Set_locale regionalSymbol
                            Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
                        End If
                    End If
                ElseIf j = LBound(a) + 3 Then
                    If aString <> noKey Then
                        IActiveLock_AddLockCode lockHDFirmware, SizeLockType  'dw1 build up lockTypes
                        If aString <> modHardware.GetHDSerialFirmware() Then
                            Set_locale regionalSymbol
                            Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
                        End If
                    End If
                ElseIf j = LBound(a) + 4 Then
                    If aString <> noKey Then
                        IActiveLock_AddLockCode lockWindows, SizeLockType  'dw1 build up lockTypes
                        If aString <> modHardware.GetWindowsSerial() Then
                            Set_locale regionalSymbol
                            Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
                        End If
                    End If
                ElseIf j = LBound(a) + 5 Then
                    If aString <> noKey Then
                        IActiveLock_AddLockCode lockBIOS, SizeLockType  'dw1 build up lockTypes
                        If aString <> modHardware.GeetBiosVersion() Then
                            Set_locale regionalSymbol
                            Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
                        End If
                    End If
                ElseIf j = LBound(a) + 6 Then
                    If aString <> noKey Then
                        IActiveLock_AddLockCode lockMotherboard, SizeLockType  'dw1 build up lockTypes
                        If aString <> modHardware.GetMotherboardSerial() Then
                            Set_locale regionalSymbol
                            Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
                        End If
                    End If
                ElseIf j = LBound(a) + 7 Then
                    If aString <> noKey Then
                        IActiveLock_AddLockCode lockIP, SizeLockType  'dw1 build up lockTypes
                        If aString <> modHardware.GetIPaddress() Then
                            Set_locale regionalSymbol
                            Err.Raise ActiveLockErrCodeConstants.alerrWrongIPaddress, ACTIVELOCKSTRING, STRWRONGIPADDRESS
                        End If
                    End If
                End If
            Next j
            
            Dim Index As Integer, i As Integer
            Index = 0
            i = 1
            ' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
            Do While i > 0
                i = InStr(Index + 1, usedcode, vbLf)
                If i > 0 Then Index = i
            Loop
            ' user name starts from Index+1 to the end
            userFromInstallCode = Mid$(usedcode, Index + 1)
            ' Check to see if this user name matches the one in the liberation key
            If userFromInstallCode <> Lic.Licensee Then
                Set_locale regionalSymbol
                Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
            End If
            ' above is last possible failure point
            'mUsedLockTypes = tmpLockType ' per David Weatherall
                        
            usedcode = Mid(usedcode, 1, Len(usedcode) - Len(userFromInstallCode) - 1)
            
            IActiveLock_LockCode = Lic.ToString() & vbLf & usedcode
        Else
            Set_locale regionalSymbol
            Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
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
' Name: Property Let IActiveLock_SoftwareVersion
' Input:
'   ByVal RHS As String - Software version for the license
' Output: None
' Purpose: Specifies the SoftwareVersion for the license
' Remarks: None
'===============================================================================
Private Property Let IActiveLock_SoftwareVersion(ByVal RHS As String)
    mSoftwareVer = RHS
End Property
'===============================================================================
' Name: Property Get IActiveLock_SoftwareVersion
' Input: None
' Output:
'   String - Software version  for the license
' Purpose: Gets the SoftwareVersion for the license
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_SoftwareVersion() As String
    IActiveLock_SoftwareVersion = mSoftwareVer
End Property
'===============================================================================
' Name: Property Get IActiveLock_RemainingTrialDays
' Input: None
' Output:
'   Long - Number of Used Trial Days
' Purpose: Gets the Number of Used Trial Days
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_RemainingTrialDays() As Long
    IActiveLock_RemainingTrialDays = mRemainingTrialDays
End Property

'===============================================================================
' Name: Property Let IActiveLock_SoftwareCode
' Input:
'   ByVal RHS As String - Software code for the license
' Output: None
' Purpose: Specifies the SoftwareCode for the license
' Remarks: SoftwareCode is an RSA public key.  This code will be used to verify license keys later on
'===============================================================================
Private Property Let IActiveLock_SoftwareCode(ByVal RHS As String)
    mSoftwareCode = RHS
End Property


'===============================================================================
' Name: Function IActiveLock_Transfer
' Input:
'   ByVal OtherSoftwareCode As String - Installation code from another machine/software
' Output: None
' Purpose: Not implemented yet
' Remarks: Transfers an Activelock license from one machine/software to another
'===============================================================================
Private Function IActiveLock_Transfer(ByVal OtherSoftwareCode As String) As String
    ' TODO: Implement me!
    Set_locale regionalSymbol
    Err.Raise ActiveLockErrCodeConstants.alerrNotImplemented, ACTIVELOCKSTRING, STRNOTIMPLEMENTED
End Function




'===============================================================================
' Name: Property Get IActiveLock_UsedDays
' Input: None
' Output: None
' Purpose: Gets the number of days the license was used after validating it.
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_UsedDays() As Long
    Dim Lic As ProductLicense
    Set Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
    
    If Lic Is Nothing Then Exit Property
    
    ' validate the license
    ValidateLic Lic
    'IActiveLock_UsedDays = CLng(DateDiff("d", Lic.RegisteredDate, modActiveLock.UTC(Now)))
    'IActiveLock_UsedDays = CLng(DateDiff("d", CDate(Replace(Lic.RegisteredDate, ".", "-")), modActiveLock.UTC(Now())))
    IActiveLock_UsedDays = CLng(DateDiff("d", ActiveLockDate(CDate(Lic.RegisteredDate)), ActiveLockDate(UTC(Now))))
    If IActiveLock_UsedDays < 0 Then
        Set_locale regionalSymbol
        Err.Raise ActiveLockErrCodeConstants.alerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
    End If
End Property

'===============================================================================
' Name: Property Get IActiveLock_RemainingTrialRuns
' Input: None
' Output:
'   Long - Number of Used Trial Runs
' Purpose: Gets the Number of Used Trial Runs
' Remarks: None
'===============================================================================
Private Property Get IActiveLock_RemainingTrialRuns() As Long
    IActiveLock_RemainingTrialRuns = mRemainingTrialRuns
End Property


