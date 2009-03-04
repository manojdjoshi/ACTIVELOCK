Option Strict Off
Option Explicit On 
Imports System.IO
Imports ActiveLock3_6NET
Imports System.Security.Cryptography
Imports System.Management
Imports System.TimeSpan
Imports System.Text ' For StringBuilder
Imports System.Runtime.InteropServices ' For DLL Call

#Region "Copyright"
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
#End Region

''' <summary>
''' <para>This is an implementation of IActiveLock.</para>
''' <para>It is not public-creatable, and so must only be accessed via ActiveLock.NewInstance() method.</para>
''' <para>Includes Key generation and validation routines.</para>
''' </summary>
''' <remarks>If you want to turn off dll-checksumming, add this compilation flag to the Project Properties (Make tab) AL_DEBUG = 1</remarks>
Friend Class ActiveLock
    Implements _IActiveLock
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
    Private mLicenseKeyTypes As IActiveLock.ALLicenseKeyTypes
    Private mUsedLockTypes() As IActiveLock.ALLockTypes
    Private mTrialType As Integer
    Private mTrialLength As Integer
    Private mRemainingTrialDays As Integer
    Private mRemainingTrialRuns As Integer
    Private mTrialHideTypes As IActiveLock.ALTrialHideTypes
    Private mKeyStore As _IKeyStoreProvider
    Private mKeyStorePath As String
    Private MyNotifier As New ActiveLockEventNotifier
    Private MyGlobals As New Globals
    Private mLibKeyPath As String
    Private mCheckTimeServerForClockTampering As IActiveLock.ALTimeServerTypes
    Private mChecksystemfilesForClockTampering As IActiveLock.ALSystemFilesTypes
    Private mLicenseFileType As IActiveLock.ALLicenseFileTypes
    Private mAutoRegister As IActiveLock.ALAutoRegisterTypes
    Private mTrialWarning As IActiveLock.ALTrialWarningTypes
    Private mUsedLockType As Integer
    Private dontValidateLicense As Boolean

	''' <summary>
	''' Registry hive used to store Activelock settings.
	''' </summary>
	''' <remarks></remarks>
    Private Const AL_REGISTRY_HIVE As String = "Software\ActiveLock\ActiveLock3"

	' Transients

	''' <summary>
	''' flag to indicate that ActiveLock has been initialized
	''' </summary>
	''' <remarks></remarks>
    Private mfInit As Boolean ' flag to indicate that ActiveLock has been initialized

	''' <summary>
	''' <para>GetVolumeInformation</para>
	''' </summary>
	''' <param name="lpRootPathName">String - A pointer to a string that contains the root directory of the volume to be described.</param>
	''' <param name="lpVolumeNameBuffer">A pointer to a buffer that receives the name of a specified volume. The maximum buffer size is MAX_PATH+1.</param>
	''' <param name="nVolumeNameSize">The length of a volume name buffer, in TCHARs. The maximum buffer size is MAX_PATH+1.</param>
	''' <param name="lpVolumeSerialNumber">A pointer to a variable that receives the volume serial number.</param>
	''' <param name="lpMaximumComponentLength">A pointer to a variable that receives the maximum length, in TCHARs, of a file name component that a specified file system supports.</param>
	''' <param name="lpFileSystemFlags">A pointer to a variable that receives flags associated with the specified file system.</param>
	''' <param name="lpFileSystemNameBuffer">A pointer to a buffer that receives the name of the file system, for example, the FAT file system or the NTFS file system. The maximum buffer size is MAX_PATH+1.</param>
	''' <param name="nFileSystemNameSize">The length of the file system name buffer, in TCHARs. The maximum buffer size is MAX_PATH+1.</param>
	''' <returns>
	''' <para>If all the requested information is retrieved, the return value is nonzero.</para>
	''' <para>If not all the requested information is retrieved, the return value is zero (0). To get extended error information, call GetLastError.</para>
	''' </returns>
	''' <remarks>
	''' <para>See <a href="http://msdn.microsoft.com/en-us/library/aa364993(VS.85).aspx">http://msdn.microsoft.com/en-us/library/aa364993(VS.85).aspx</a></para>
	''' </remarks>
    Public Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
       (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As StringBuilder, _
       ByVal nVolumeNameSize As Integer, ByVal lpVolumeSerialNumber As Integer, _
       ByVal lpMaximumComponentLength As Integer, ByVal lpFileSystemFlags As Integer, _
       ByVal lpFileSystemNameBuffer As StringBuilder, ByVal nFileSystemNameSize As Integer) As Integer

	''' <summary>
	''' IActiveLock_LicenseKeyType - Specifies the ALLicenseKeyTypes type
	''' </summary>
	''' <value>ByVal RHS As ALLicenseKeyTypes - ALLicenseKeyTypes type</value>
	''' <remarks>None</remarks>
    Private WriteOnly Property IActiveLock_LicenseKeyType() As IActiveLock.ALLicenseKeyTypes Implements _IActiveLock.LicenseKeyType
        Set(ByVal Value As IActiveLock.ALLicenseKeyTypes)
            mLicenseKeyTypes = Value
        End Set
	End Property

	''' <summary>
	''' Gets the Registered Level for the license after validating it.
	''' </summary>
	''' <value></value>
	''' <returns>String - License RegisteredLevel</returns>
	''' <remarks>None</remarks>
    Private ReadOnly Property IActiveLock_RegisteredLevel() As String Implements _IActiveLock.RegisteredLevel
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
            If Lic Is Nothing Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If
            ' Validate the License.
            ValidateLic(Lic)
            Return Lic.RegisteredLevel
        End Get
	End Property

	''' <summary>
	''' Gets the LicenseClass
	''' </summary>
	''' <value></value>
	''' <returns>String - LicenseClass</returns>
	''' <remarks></remarks>
    Private ReadOnly Property IActiveLock_LicenseClass() As String Implements _IActiveLock.LicenseClass
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
            If Lic Is Nothing Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If
            ' Validate the License.
            ValidateLic(Lic)
            Return Lic.LicenseClass
        End Get
	End Property

	''' <summary>
	''' Gets the Number of Used Trial Days
	''' </summary>
	''' <value></value>
	''' <returns>Integer - License Used Trial Days</returns>
	''' <remarks></remarks>
    Private ReadOnly Property IActiveLock_RemainingTrialDays() As Integer Implements _IActiveLock.RemainingTrialDays
        Get
            Return mRemainingTrialDays
        End Get
	End Property

	''' <summary>
	''' Gets the Number of Used Trial Runs
	''' </summary>
	''' <value></value>
	''' <returns>Integer - License Used Trial Runs</returns>
	''' <remarks></remarks>
    Private ReadOnly Property IActiveLock_RemainingTrialRuns() As Integer Implements _IActiveLock.RemainingTrialRuns
        Get
            Return mRemainingTrialRuns
        End Get
	End Property

	''' <summary>
	''' Gets the Number of concurrent users for the networked license
	''' </summary>
	''' <value></value>
	''' <returns>Integer - Number of concurrent users for the networked license</returns>
	''' <remarks></remarks>
    Private ReadOnly Property IActiveLock_MaxCount() As Integer Implements _IActiveLock.MaxCount
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
            If Lic Is Nothing Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If
            ' Validate the License.
            ValidateLic(Lic)
            Return Lic.MaxCount
        End Get
	End Property

	''' <summary>
	''' <para>IActiveLock Interface implementation</para>
	''' <para>Specifies the liberation key auto file path name</para>
	''' </summary>
	''' <value>ByVal RHS As String - Liberation key file auto path name</value>
	''' <remarks></remarks>
    Private WriteOnly Property IActiveLock_AutoRegisterKeyPath() As String Implements _IActiveLock.AutoRegisterKeyPath
        Set(ByVal Value As String)
            mLibKeyPath = Value
        End Set
	End Property

	' TODO: ActiveLock.vb - Property AutoRegisterKeyPath - I think this should read Gets, not Sets: japreja
	''' <summary>
	''' Sets the auto register file full path
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
    Private ReadOnly Property AutoRegisterKeyPath() As String
        Get
            Return mLibKeyPath
        End Get
    End Property

	' TODO: ActiveLock.vb - Property IActiveLock_EventNotifier - Update the Returns comment!
	''' <summary>
	''' Gets a notification from Activelock
	''' </summary>
	''' <value></value>
	''' <returns>ActiveLockEventNotifier - ???</returns>
	''' <remarks></remarks>
    Private ReadOnly Property IActiveLock_EventNotifier() As ActiveLockEventNotifier Implements _IActiveLock.EventNotifier
        Get
            IActiveLock_EventNotifier = MyNotifier
        End Get
	End Property

	''' <summary>
	''' Gets the license registration date after validating it.
	''' </summary>
	''' <value></value>
	''' <returns>String - License registration date.</returns>
	''' <remarks>This is the date the license was generated by Alugen. NOT the date the license was activated.</remarks>
    Private ReadOnly Property IActiveLock_RegisteredDate() As String Implements _IActiveLock.RegisteredDate
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
            If Lic Is Nothing Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If
            ' Validate the License.
            ValidateLic(Lic)
            Return Lic.RegisteredDate
        End Get
	End Property

	''' <summary>
	''' Gets the registered user name after validating the license
	''' </summary>
	''' <value></value>
	''' <returns>String - Registered user name</returns>
	''' <remarks></remarks>
    Private ReadOnly Property IActiveLock_RegisteredUser() As String Implements _IActiveLock.RegisteredUser
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
            If Lic Is Nothing Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If
            ' Validate the License.
            ValidateLic(Lic)
            Return Lic.Licensee
        End Get
	End Property

	''' <summary>
	''' Returns the expiration date of the license after validating it
	''' </summary>
	''' <value></value>
	''' <returns>String - Expiration date of the license</returns>
	''' <remarks></remarks>
    Private ReadOnly Property IActiveLock_ExpirationDate() As String Implements _IActiveLock.ExpirationDate
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
            If Lic Is Nothing Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If
            ' Validate the License.
            ValidateLic(Lic)
            Return Lic.Expiration
        End Get
	End Property

	''' <summary>
	''' Specifies the license file path name
	''' </summary>
	''' <value>ByVal RHS As String - License file path name</value>
	''' <remarks></remarks>
    Private WriteOnly Property IActiveLock_KeyStorePath() As String Implements _IActiveLock.KeyStorePath
        Set(ByVal Value As String)
            If Not mKeyStore Is Nothing Then
                mKeyStore.KeyStorePath = Value
            End If
            mKeyStorePath = Value
        End Set
	End Property

	''' <summary>
	''' <para>Specifies the key store type</para>
	''' <para>This version of Activelock does not work with the registry</para>
	''' </summary>
	''' <value>ByVal RHS As LicStoreType - License store type</value>
	''' <remarks>Portions of this (RegistryKeyStoreProvider) not implemented yet</remarks>
    Private WriteOnly Property IActiveLock_KeyStoreType() As IActiveLock.LicStoreType Implements _IActiveLock.KeyStoreType
        Set(ByVal Value As IActiveLock.LicStoreType)
            ' Instantiate Key Store Provider
            If Value = IActiveLock.LicStoreType.alsFile Then
                mKeyStore = New FileKeyStoreProvider
            Else
                ' Set mKeyStore = New RegistryKeyStoreProvider
				' TODO: ActiveLock.vb - Property IActiveLock_KeyStoreType - Implement me!
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNotImplemented, ACTIVELOCKSTRING, STRNOTIMPLEMENTED)
            End If
            ' Set Key Store Path in KeyStoreProvider
            If mKeyStorePath <> "" Then
                mKeyStore.KeyStorePath = mKeyStorePath
            End If
        End Set
	End Property

	''' <summary>
	''' Gets or Sets the ALLockTypes type
	''' </summary>
	''' <value>ByVal RHS As ALLockTypes - ALLockTypes type</value>
	''' <returns>ALLockTypes - Lock types type</returns>
	''' <remarks></remarks>
    Private Property IActiveLock_LockType() As IActiveLock.ALLockTypes Implements _IActiveLock.LockType
        Get
            Return mLockTypes
        End Get
        Set(ByVal Value As IActiveLock.ALLockTypes)
            mLockTypes = Value
        End Set
	End Property

	''' <summary>
	''' Helper function to build up array of used LockType s
	''' </summary>
	''' <param name="LockType"><para>ByVal LockType As ALLockTypes _ to be added to array.</para><para>ByRef Byref LockTypes() As ALLockTypes - array of used LockTypes being built up.</para></param>
	''' <param name="SizeLT">ByRef SizeLT as Integer - size of array of used LockTypes being built up</param>
	''' <remarks></remarks>
    Private Sub IActiveLock_AddLockCode(ByVal LockType As IActiveLock.ALLockTypes, ByRef SizeLT As Integer)
        ReDim Preserve mUsedLockTypes(SizeLT)
        mUsedLockTypes(SizeLT) = LockType
        SizeLT = SizeLT + 1
	End Sub

	''' <summary>
	''' Gets the ALTrialHideTypes type
	''' </summary>
	''' <value></value>
	''' <returns>ALTrialHideTypes - Trial Hide types type</returns>
	''' <remarks></remarks>
    Private Property IActiveLock_TrialHideType() As IActiveLock.ALTrialHideTypes Implements _IActiveLock.TrialHideType
        Get
            Return mTrialHideTypes
        End Get
        Set(ByVal Value As IActiveLock.ALTrialHideTypes)
            mTrialHideTypes = Value
        End Set
	End Property

	''' <summary>
	''' Gets the SoftwareName for the license
	''' </summary>
	''' <value></value>
	''' <returns>String - Software name  for the license</returns>
	''' <remarks></remarks>
    Private Property IActiveLock_SoftwareName() As String Implements _IActiveLock.SoftwareName
        Get
            Return mSoftwareName
        End Get
        Set(ByVal Value As String)
            mSoftwareName = Value
        End Set
	End Property

	''' <summary>
	''' Gets/Sets the SoftwarePassword for the license
	''' </summary>
	''' <value>ByVal RHS As String - Software Password for the license</value>
	''' <returns>String - Software Password for the license</returns>
	''' <remarks></remarks>
    Private Property IActiveLock_SoftwarePassword() As String Implements _IActiveLock.SoftwarePassword
        Get
            Return mSoftwarePassword
        End Get
        Set(ByVal Value As String)
            mSoftwarePassword = Value
        End Set
	End Property

	''' <summary>
	''' Specifies whether a Time Server should be used to check Clock Tampering
	''' </summary>
	''' <value>ByVal iServer As Integer - Flag being passed to check the time server</value>
	''' <remarks></remarks>
    Private WriteOnly Property IActiveLock_CheckTimeServerForClockTampering() As IActiveLock.ALTimeServerTypes Implements _IActiveLock.CheckTimeServerForClockTampering
        Set(ByVal Value As IActiveLock.ALTimeServerTypes)
            mCheckTimeServerForClockTampering = Value
        End Set
	End Property

	''' <summary>
	''' Specifies whether a Time Server should be used to check Clock Tampering
	''' </summary>
	''' <value>ByVal iServer As Integer - Flag being passed to check the time server</value>
	''' <remarks></remarks>
    Private WriteOnly Property IActiveLock_CheckSystemFilesForClockTampering() As IActiveLock.ALSystemFilesTypes Implements _IActiveLock.CheckSystemFilesForClockTampering
        Set(ByVal Value As IActiveLock.ALSystemFilesTypes)
            mChecksystemfilesForClockTampering = Value
        End Set
	End Property

	' TODO: ActiveLock.vb - Property IActiveLock_LicenseFileType - Update return value comment!
	''' <summary>
	''' Specifies whether the License File should be encrypted or not
	''' </summary>
	''' <value>ByVal Value As IActiveLock.ALLicenseFileTypes - Flag to indicate the license file will be encrypted or not</value>
	''' <returns></returns>
	''' <remarks></remarks>
    Private Property IActiveLock_LicenseFileType() As IActiveLock.ALLicenseFileTypes Implements _IActiveLock.LicenseFileType
        Set(ByVal Value As IActiveLock.ALLicenseFileTypes)
            mLicenseFileType = Value
        End Set
        Get
            Return mLicenseFileType
        End Get
	End Property

	' TODO: ActiveLock.vb - Property IActiveLock_AutoRegister - Update Comment - Not Documented!
	''' <summary>
	''' Not Documented!
	''' </summary>
	''' <value>ALAutoRegisterTypes - ALAutoRegisterType</value>
	''' <remarks></remarks>
    Private WriteOnly Property IActiveLock_AutoRegister() As IActiveLock.ALAutoRegisterTypes Implements _IActiveLock.AutoRegister
        Set(ByVal Value As IActiveLock.ALAutoRegisterTypes)
            mAutoRegister = Value
        End Set
	End Property

	''' <summary>
	''' Specifies whether the License File should be encrypted or not
	''' </summary>
	''' <value>ByVal Value As IActiveLock.ALTrialWarningTypes - Flag to indicate the license file will be encrypted or not.</value>
	''' <remarks></remarks>
    Private WriteOnly Property IActiveLock_TrialWarning() As IActiveLock.ALTrialWarningTypes Implements _IActiveLock.TrialWarning
        Set(ByVal Value As IActiveLock.ALTrialWarningTypes)
            mTrialWarning = Value
        End Set
	End Property

	''' <summary>
	''' Gets/Sets the TrialType for the license
	''' </summary>
	''' <value>ByVal Value As IActiveLock.ALTrialTypes</value>
	''' <returns>ALTrialTypes - Trial Type  for the license</returns>
	''' <remarks></remarks>
    Private Property IActiveLock_TrialType() As IActiveLock.ALTrialTypes Implements _IActiveLock.TrialType
        Get
            Return mTrialType
        End Get
        Set(ByVal Value As IActiveLock.ALTrialTypes)
            mTrialType = Value
        End Set
	End Property

	''' <summary>
	''' Gets/Sets the TrialLength for the license
	''' </summary>
	''' <value></value>
	''' <returns>Integer - Trial Length  for the license</returns>
	''' <remarks></remarks>
    Private Property IActiveLock_TrialLength() As Integer Implements _IActiveLock.TrialLength
        Get
            Return mTrialLength
        End Get
        Set(ByVal Value As Integer)
            mTrialLength = Value
        End Set
	End Property

	''' <summary>
	''' Combines the user name with the lock code and returns it as the installation code
	''' </summary>
	''' <param name="User">Optional - String - User name</param>
	''' <param name="Lic">Optional - ProductLicense - Product License</param>
	''' <value></value>
	''' <returns>String - Installation code</returns>
	''' <remarks></remarks>
    Private ReadOnly Property IActiveLock_InstallationCode(Optional ByVal User As String = vbNullString, Optional ByVal Lic As ProductLicense = Nothing) As String Implements _IActiveLock.InstallationCode
        Get
            'Before we generate the installation code, let's check if this app is using a short key
            Dim strReq, strLock As String
            Dim strReq2 As String
            If mLicenseKeyTypes = IActiveLock.ALLicenseKeyTypes.alsShortKeyMD5 Then
                ' Shortkeys are no longer using the HDD firmware serial number
                ' they are using the Computer Fingerprint after v3.6
                Return IActivelock_GenerateShortSerial(modHardware.GetFingerprint())

            ElseIf mLicenseKeyTypes = IActiveLock.ALLicenseKeyTypes.alsRSA Then

                ' Generate Request code to Lock

                'Restrict user name to 2000 characters; need more? why?
                If Len(User) > 2000 Then
                    Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrUserNameTooLong, ACTIVELOCKSTRING, STRUSERNAMETOOLONG)
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

                ' combine with app name and version
                strReq = strReq & "&&&" & IActiveLock_SoftwareName & " (" & IActiveLock_SoftwareVersion & ")"

                ' base-64 encode the request
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
                        ' Set_locale(regionalSymbol)
                        ' Err.Raise ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
                    End If
                End If
            End If
            Return Nothing
        End Get

	End Property

	''' <summary>
	''' Gets the SoftwareVersion for the license
	''' </summary>
	''' <value></value>
	''' <returns>String - Software version  for the license</returns>
	''' <remarks></remarks>
    Private Property IActiveLock_SoftwareVersion() As String Implements _IActiveLock.SoftwareVersion
        Get
            Return mSoftwareVer
        End Get
        Set(ByVal Value As String)
            mSoftwareVer = Value
        End Set
	End Property

	''' <summary>
	''' Specifies the SoftwareCode for the license
	''' </summary>
	''' <value>ByVal RHS As String - Software code for the license</value>
	''' <remarks>SoftwareCode is an RSA public key.  This code will be used to verify license keys later on.</remarks>
    Private WriteOnly Property IActiveLock_SoftwareCode() As String Implements _IActiveLock.SoftwareCode
        Set(ByVal Value As String)
            mSoftwareCode = Value
        End Set
	End Property

	' TODO: ActiveLock.vb - Property IActiveLock_UsedDays - returns comment not documented!
	''' <summary>
	''' Gets the number of days the license was used after validating it.
	''' </summary>
	''' <value></value>
	''' <returns>Integer - ?</returns>
	''' <remarks></remarks>
    Public ReadOnly Property IActiveLock_UsedDays() As Integer Implements _IActiveLock.UsedDays
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
            If Lic Is Nothing Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If

            ' validate the license
            If dontValidateLicense = False Then ValidateLic(Lic)

            'IActiveLock_UsedDays = CInt(DateDiff("d", Lic.RegisteredDate, Now.UtcNow))
            Dim mydate As DateTime = ActiveLockDate(Date.UtcNow)
            IActiveLock_UsedDays = mydate.Subtract(ActiveLockDate(CDate(Lic.RegisteredDate))).Days      'CInt(DateDiff("d", CDate(Replace(Lic.RegisteredDate, ".", "-")), Date.UtcNow))
            If IActiveLock_UsedDays < 0 Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
            End If
        End Get
	End Property

	''' <summary>
	''' Gets the Lock Type selected in Alugen.
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
    Public ReadOnly Property IActiveLock_UsedLockType() As Integer Implements _IActiveLock.UsedLockType
        Get
            Dim Lic As ProductLicense
            Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)

            If Lic Is Nothing Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)
            End If

            ' validate the license
            If dontValidateLicense = False Then ValidateLic(Lic)

            If Lic.LicenseCode = "Short Key" Then
                IActiveLock_UsedLockType = 0
            Else
                Dim usedcode As String
                If Left(Lic.LicenseCode, 1) = "+" Then
                    usedcode = modBase64.Base64_Decode(Mid(Lic.LicenseCode, 2))
                Else
                    usedcode = modBase64.Base64_Decode((Lic.LicenseCode))
                End If
                Dim Index, i As Integer
                Index = 0 : i = 1
                ' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
                Do While i > 0
                    i = InStr(Index + 1, usedcode, vbLf)
                    If i > 0 Then Index = i
                Loop

                If Index <= 0 Then Exit Property
                ' lockcode is from beginning to Index-1
                usedcode = Left(usedcode, Index - 1)
                If Left(usedcode, 1) = "+" Then
                    usedcode = Right(usedcode, usedcode.Length - 1)
                End If
                Dim myarray() As String
                Dim counter As Integer = 0
                myarray = usedcode.Split(vbLf)
                For i = 0 To UBound(myarray)
                    If myarray(i) <> "nokey" Then
                        counter = counter + 2 ^ i
                    End If
                Next
                IActiveLock_UsedLockType = counter
            End If
        End Get
	End Property

	' TODO: ActiveLock.vb - Sub Class_Initialize_Renamed - Add documentation!
	''' <summary>
	''' Not documented!
	''' </summary>
	''' <remarks>Class_Initialize was upgraded to Class_Initialize_Renamed</remarks>
    Private Sub Class_Initialize_Renamed()
        ' Default to alsFile
        IActiveLock_KeyStoreType = IActiveLock.LicStoreType.alsFile
	End Sub

	' TODO: ActiveLock.vb - Sub New - Add documentation!
	''' <summary>
	''' Not Documented!
	''' </summary>
	''' <remarks></remarks>
    Public Sub New()
        MyBase.New()
        Class_Initialize_Renamed()
    End Sub

	' TODO: ActiveLock.vb - Sub IActiveLock_Init - Update comment for strPath
	''' <summary>
	''' Initalizes Activelock
	''' </summary>
	''' <param name="strPath"></param>
	''' <param name="autoLicString">ByRef autoLicString As String - Returned License Key of AutoRegister is successful.</param>
	''' <remarks>
	''' <para>Performs CRC check on Alcrypto.</para>
	''' <para>Performs auto license registration if the license file is found.</para>
	''' </remarks>
    Private Sub IActiveLock_Init(Optional ByVal strPath As String = "", Optional ByRef autoLicString As String = "") Implements _IActiveLock.Init
        ' If running in Debug mode, don't bother with dll authentication
#If CBool(AL_DEBUG) <> False Then
		GoTo Done
#End If
        ' ALL file generatiand software usage on PCs with different cultures 
        ' does not work due to the usage of Chr() function in Base64_Decode in
        ' modBase64.vb
        ' The following is necessary to fix the problem
        My.Application.ChangeCulture("en-US")

        ' Checksum ALCrypto3NET.dll
        'Const ALCRYPTO_MD5 As String = "54BED793A0E24D3E71706EEC4FA1B0FC"
        'Const ALCRYPTO_MD5$ = "be299ad0f52858fdd9ea3626468dc05c"
        Const ALCRYPTO_MD5 As String = "6E5C849489281E47A9B4BB8375506D" 'mod for VB2005'
        Dim strdata As String = String.Empty
        Dim strMD5, usedFile As String
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
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrFileTampered, ACTIVELOCKSTRING, "Alcrypto3Net.dll could not be found in system32 directory.")
        End If
        Call modActiveLock.ReadFile(usedFile, strdata)
        ' use the .NET's native MD5 functions instead of our own MD5 hashing routine
        ' and instead of ALCrypto's md5_hash() function.
        strMD5 = UCase(strdata)    '<--- ReadFile procedure already computes the MD5.Hash

        If strMD5 <> ALCRYPTO_MD5 Then

            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrFileTampered, ACTIVELOCKSTRING, STRFILETAMPERED)
        End If
        ' Perform automatic license registration
        If AutoRegisterKeyPath <> "" And mAutoRegister = IActiveLock.ALAutoRegisterTypes.alsEnableAutoRegistration Then
            DoAutoRegistration(autoLicString)
            If Err.Number <> 0 Then autoLicString = ""
        End If
Done:
        mfInit = True
	End Sub

	''' <summary>
	''' Checks the specified path to see if the auto registration liberation file is there
	''' </summary>
	''' <param name="strLibKey">strLibKey As String - Returned liberation key if auto register is successful.</param>
	''' <remarks></remarks>
    Private Sub DoAutoRegistration(ByRef strLibKey As String)

        ' Don't bother to proceed unless the file is there.
        If Not File.Exists(AutoRegisterKeyPath) Then Exit Sub

        ReadLibKey(AutoRegisterKeyPath, strLibKey)
        IActiveLock_Register(strLibKey)

        ' If registration is successful, delete the liberation file so we won't register the same file on next startup
        Kill(AutoRegisterKeyPath)
	End Sub

	''' <summary>
	''' Reads the liberation key from a file
	''' </summary>
	''' <param name="sFileName">ByVal sFileName As String - File name to read the liberation key from.</param>
	''' <param name="strLibKey">ByRef strLibKey As String -  Liberation key returned</param>
	''' <remarks></remarks>
    Private Sub ReadLibKey(ByVal sFileName As String, ByRef strLibKey As String)
        Dim hFile As Integer
        hFile = FreeFile()
        FileOpen(hFile, sFileName, OpenMode.Input)
        On Error GoTo finally_Renamed
        strLibKey = InputString(hFile, LOF(hFile))
finally_Renamed:
        FileClose(hFile)
	End Sub

	' TODO: ActiveLock.vb - Sub IActiveLock_Acquire - Update the input paramaters comments!
	''' <summary>
	''' <para>Acquires an Activelock License.</para>
	''' <para>This is the main method that retrieves an Activelock license, validates it, and ends the trial license if it exists.</para>
	''' </summary>
	''' <param name="strMsg"></param>
	''' <param name="strRemainingTrialDays"></param>
	''' <param name="strRemainingTrialRuns"></param>
	''' <param name="strTrialLength"></param>
	''' <param name="strUsedDays"></param>
	''' <param name="strExpirationDate"></param>
	''' <param name="strRegisteredUser"></param>
	''' <param name="strRegisteredLevel"></param>
	''' <param name="strLicenseClass"></param>
	''' <param name="strMaxCount"></param>
	''' <param name="strLicenseFileType"></param>
	''' <param name="strLicenseType"></param>
	''' <param name="strUsedLockType"></param>
	''' <remarks></remarks>
    Private Sub IActiveLock_Acquire(Optional ByRef strMsg As String = "", Optional ByRef strRemainingTrialDays As String = "", Optional ByRef strRemainingTrialRuns As String = "", Optional ByRef strTrialLength As String = "", Optional ByRef strUsedDays As String = "", Optional ByRef strExpirationDate As String = "", Optional ByRef strRegisteredUser As String = "", Optional ByRef strRegisteredLevel As String = "", Optional ByRef strLicenseClass As String = "", Optional ByRef strMaxCount As String = "", Optional ByRef strLicenseFileType As String = "", Optional ByRef strLicenseType As String = "", Optional ByRef strUsedLockType As String = "") Implements _IActiveLock.Acquire
        Dim trialActivated As Boolean
        Dim adsText As String = String.Empty
        Dim strStream As String = String.Empty
        Dim Lic As ProductLicense
        Dim trialStatus As Boolean

        strStream = mSoftwareName & mSoftwareVer & mSoftwarePassword

        ' Get the current date format and save it to regionalSymbol variable
        Get_locale()
        ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
        Set_locale("")

        'Check the Key Store Provider
        If mKeyStore Is Nothing Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrKeyStoreInvalid, ACTIVELOCKSTRING, STRKEYSTOREUNINITIALIZED)
            'Check the Key Store Path (LIC file path)
        ElseIf mKeyStorePath = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrKeyStorePathInvalid, ACTIVELOCKSTRING, STRKEYSTOREPATHISEMPTY)
        ElseIf mSoftwareName = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareName, ACTIVELOCKSTRING, STRNOSOFTWARENAME)
        ElseIf mSoftwareVer = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareVersion, ACTIVELOCKSTRING, STRNOSOFTWAREVERSION)
        ElseIf mSoftwarePassword = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwarePassword, ACTIVELOCKSTRING, STRNOSOFTWAREPASSWORD)
        ElseIf specialChar(mSoftwarePassword) Or mSoftwarePassword.Length > 255 Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrSoftwarePasswordInvalid, ACTIVELOCKSTRING, STRSOFTWAREPASSWORDINVALID)
        End If

        Lic = mKeyStore.Retrieve(mSoftwareName, mLicenseFileType)
        If Lic Is Nothing Then
            ' There's no valid license, so let's see if we can grant this user a "Trial License"
            If mTrialType = IActiveLock.ALTrialTypes.trialNone Then 'No Trial
                GoTo noRegistration
            End If

            On Error GoTo noRegistration
            strMsg = ""
            If mTrialHideTypes = 0 Then
                mTrialHideTypes = IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialSteganography
            End If

            If mCheckTimeServerForClockTampering = IActiveLock.ALTimeServerTypes.alsCheckTimeServer Then
                If SystemClockTampered() Then
                    Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                End If
            End If
            If mChecksystemfilesForClockTampering = IActiveLock.ALSystemFilesTypes.alsCheckSystemFiles Then
                If ClockTampering() Then
                    Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                End If
            End If

            trialStatus = ActivateTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, strMsg, mSoftwarePassword, mCheckTimeServerForClockTampering, mChecksystemfilesForClockTampering, mTrialWarning, mRemainingTrialDays, mRemainingTrialRuns)
            strRemainingTrialDays = mRemainingTrialDays.ToString
            strRemainingTrialRuns = mRemainingTrialRuns.ToString
            strTrialLength = mTrialLength.ToString
            ' Set the locale date format to what we had before; can't leave changed
            Set_locale((regionalSymbol))
            If trialStatus = True Then
                Exit Sub
            End If
            GoTo continueRegistration

noRegistration:
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense, ACTIVELOCKSTRING, STRNOLICENSE)

        Else  'Lic exists therefore we'll check the LIC file ADS

            If Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
                If mCheckTimeServerForClockTampering = IActiveLock.ALTimeServerTypes.alsCheckTimeServer Then
                    If SystemClockTampered() Then
                        Set_locale(regionalSymbol)
                        Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                    End If
                End If
                If mChecksystemfilesForClockTampering = IActiveLock.ALSystemFilesTypes.alsCheckSystemFiles Then
                    If ClockTampering() Then
                        Set_locale(regionalSymbol)
                        Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                    End If
                End If
            End If

            If CheckStreamCapability() And Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
                Dim fi As New FileInfo(mKeyStorePath)
                If fi.Length = 0 Then GoTo continueRegistration
                adsText = ADSFile.Read(mKeyStorePath, strStream)
                If adsText = "" Then
                    Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseTampered, ACTIVELOCKSTRING, STRLICENSETAMPERED)
                End If
                Dim dt1 As DateTime = Convert.ToDateTime(adsText)
                Dim dt2 As DateTime = ActiveLockDate(Date.UtcNow)
                Dim span As TimeSpan = dt2.Subtract(dt1)
                If span.TotalHours < 0 Then
                    Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                End If
                Dim ok As Integer
                ok = ADSFile.Write(ActiveLockDate(Date.UtcNow), mKeyStorePath, strStream)
                GoTo continueRegistration
            End If
        End If

continueRegistration:
        Set_locale(regionalSymbol)
        ' Validate license
        ValidateLic(Lic)
        ' Return all needed properties for faster form loading
        dontValidateLicense = True
        Dim mydate As DateTime = ActiveLockDate(Date.UtcNow)
        strUsedDays = CStr(mydate.Subtract(ActiveLockDate(CDate(Lic.RegisteredDate))).Days)     'CInt(DateDiff("d", CDate(Replace(Lic.RegisteredDate, ".", "-")), Date.UtcNow))
        strExpirationDate = Lic.Expiration
        strRegisteredUser = Lic.Licensee
        strRegisteredLevel = Lic.RegisteredLevel
        strLicenseClass = Lic.LicenseClass
        strMaxCount = Lic.MaxCount.ToString
        strLicenseFileType = Val(IActiveLock_LicenseFileType).ToString
        strLicenseType = Lic.LicenseType.ToString
        strUsedLockType = IActiveLock_UsedLockType.ToString
        dontValidateLicense = False

	End Sub

	' TODO: ActiveLock.vb - Function CheckStreamCapability - Add Comments - Not Documented!
	''' <summary>
	''' Not Documented!
	''' </summary>
	''' <returns></returns>
	''' <remarks></remarks>
    Public Function CheckStreamCapability() As Boolean
        ' The following WMI call also works but it seems to be a bit slower than the GetVolumeInformation
        ' especially when it checks the A: drive
        ' METHOD 1 - WMI
        'Dim mc As New ManagementClass("Win32_LogicalDisk")
        'Dim moc As ManagementObjectCollection = mc.GetInstances()
        'Dim strFileSystem As String = String.Empty
        'Dim mo As ManagementObject
        'For Each mo In moc
        '    If strFileSystem = String.Empty Then ' only return the file system
        '        If mo("Name").ToString = "C:" Then
        '            strFileSystem = mo("FileSystem").ToString
        '            Exit For
        '        End If
        '    End If
        '    mo.Dispose()
        'Next mo
        'If strFileSystem = "NTFS" Then
        '    CheckStreamCapability = True
        'End If

        ' METHOD 2 - GetVolumeInformation API
        Dim lsRootPathName As String = IO.Directory.GetDirectoryRoot(Application.StartupPath)
        Const MAX_PATH As Integer = 260
        Dim iSerial As Integer
        Dim iLength As Integer
        Dim iFlags As Integer
        Dim sbVol As New StringBuilder(MAX_PATH)
        Dim sbFil As New StringBuilder(MAX_PATH)
        GetVolumeInformation("c:\", sbVol, MAX_PATH, iSerial, iLength, iFlags, sbFil, MAX_PATH)
        If sbFil.ToString = "NTFS" Then
            CheckStreamCapability = True
        End If

	End Function

	''' <summary>
	''' <para>Validates the License Key using RSA signature verification.</para>
	''' <para>License key contains the RSA signature of IActiveLock_LockCode.</para>
	''' </summary>
	''' <param name="Lic">Lic As ProductLicense - Product license</param>
	''' <remarks></remarks>
    Private Sub ValidateKey(ByRef Lic As ProductLicense)
        Dim strPubKey As String
        Dim strSig As String
        Dim strLic As String
        Dim strLicKey As String

        strPubKey = mSoftwareCode

        ' make sure software code is set
        If mSoftwareCode = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNotInitialized, ACTIVELOCKSTRING, STRNOSOFTWARECODE)
        End If

        strLic = IActiveLock_LockCode(Lic)
        strLicKey = Lic.LicenseKey

        If Left(strPubKey, 3) <> "RSA" Then 'ALCrypto
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
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
            End If
        Else    ' .NET RSA

            Try


                ' Verify Signature
                Dim rsaCSP As New System.Security.Cryptography.RSACryptoServiceProvider
                Dim rsaPubParams As RSAParameters 'stores public key
                Dim strPublicBlob As String
                If strLeft(strPubKey, 6) = "RSA512" Then
                    strPublicBlob = strRight(strPubKey, Len(strPubKey) - 6)
                Else
                    strPublicBlob = strRight(strPubKey, Len(strPubKey) - 7)
                End If
                rsaCSP.FromXmlString(strPublicBlob)
                rsaPubParams = rsaCSP.ExportParameters(False)
                ' import public key params into instance of RSACryptoServiceProvider
                rsaCSP.ImportParameters(rsaPubParams)

                Dim userData As Byte() = Encoding.UTF8.GetBytes(strLic)

                Dim newsignature() As Byte
                newsignature = Convert.FromBase64String(strLicKey)
                Dim asd As AsymmetricSignatureDeformatter = New RSAPKCS1SignatureDeformatter(rsaCSP)
                Dim algorithm As HashAlgorithm = New SHA1Managed
                asd.SetHashAlgorithm(algorithm.ToString)
                Dim newhashedData() As Byte ' a byte array to store hash value
                Dim newhashedDataString As String
                newhashedData = algorithm.ComputeHash(userData)
                newhashedDataString = BitConverter.ToString(newhashedData).Replace("-", String.Empty)
                Dim verified As Boolean
                verified = asd.VerifySignature(algorithm, newsignature)
                If verified Then
                    'MsgBox("Signature Valid", MsgBoxStyle.Information)
                Else
                    Set_locale(regionalSymbol)
                    Err.Raise(AlugenGlobals.alugenErrCodeConstants.alugenProdInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                    'MsgBox("Invalid Signature", MsgBoxStyle.Exclamation)
                End If
            Catch ex As Exception
                Set_locale(regionalSymbol)
                Err.Raise(AlugenGlobals.alugenErrCodeConstants.alugenProdInvalid, ACTIVELOCKSTRING, ex.Message)
            End Try

        End If

        ' Check if license has not expired
        ' but don't do it if there's no expiration date
        If Lic.Expiration = "" Then Exit Sub
        If ActiveLockDate(Date.UtcNow) > ActiveLockDate(CDate(Lic.Expiration)) And Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
            ' ialkan - 9-23-2005 added the following to update and store the license
            ' with the new LastUsed property; otherwise setting the clock back next time
            ' might bypass the protection mechanism
            ' Update last used date
            UpdateLastUsed(Lic)
            mKeyStore.Store(Lic, mLicenseFileType)
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseExpired, ACTIVELOCKSTRING, STRLICENSEEXPIRED)
        End If
	End Sub

	''' <summary>
	''' Validates the License Key using the Short Key MD5 verification.
	''' </summary>
	''' <param name="Lic">Lic As ProductLicense - Product license</param>
	''' <param name="user">String - User</param>
	''' <remarks></remarks>
    Private Sub ValidateShortKey(ByRef Lic As ProductLicense, ByVal user As String)

        Dim oReg As clsShortSerial
        Dim m_Key As clsShortLicenseKey
        Dim sKey As String
        Dim m_ProdCode As Integer
        Dim SerialNumber As String
        Dim ExpireDate As Date
        Dim UserData As Short
        Dim RegisteredLevel As Integer

        ' make sure software code is set
        If mSoftwareCode = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNotInitialized, ACTIVELOCKSTRING, STRNOSOFTWARECODE)
        End If

        'This is a short key
        m_Key = New clsShortLicenseKey

        m_Key.AddSwapBits(0, 0, 1, 0)
        m_Key.AddSwapBits(0, 2, 1, 1)
        m_Key.AddSwapBits(0, 4, 2, 0)
        m_Key.AddSwapBits(0, 5, 2, 1)
        m_Key.AddSwapBits(2, 0, 3, 0)
        m_Key.AddSwapBits(2, 6, 3, 1)
        m_Key.AddSwapBits(2, 7, 1, 3)

        oReg = New clsShortSerial
        sKey = oReg.GenerateKey("", Left(mSoftwareCode, Len(mSoftwareCode) - 2)) 'Do not include the last 2 possible == paddings
        m_ProdCode = CInt(Left(sKey, 4))

        ' Shortkeys are no longer using the HDD firmware serial number
        ' they are using the Computer Fingerprint after v3.6
        SerialNumber = oReg.GenerateKey(mSoftwareName & mSoftwareVer & mSoftwarePassword, modHardware.GetFingerprint())

        ' verify the key is valid
        If m_Key.ValidateShortKey(Lic.LicenseKey, SerialNumber, user, m_ProdCode, ExpireDate, UserData, RegisteredLevel) = True Then
            ' After the key is disassembled it fills the output
            ' variables with expire date and license counter.
            Lic.LicenseType = CInt(CStr(modActiveLock.HiByte(UserData)))
            Lic.ProductName = mSoftwareName
            Lic.ProductVer = mSoftwareVer
            Lic.LicenseClass = ProductLicense.LicFlags.alfSingle 'Multi User License will be available with network version
            Lic.Licensee = user
            If Lic.RegisteredLevel = 0 Then
                Lic.RegisteredLevel = CStr(RegisteredLevel)
            ElseIf Lic.RegisteredLevel <> CStr(RegisteredLevel) Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
            End If
            If Lic.RegisteredDate = "" Then
                Lic.RegisteredDate = ActiveLockDate(Date.UtcNow).ToString("yyyy/MM/dd")
            End If
            ' ignore expiration date if license type is "permanent"
            If Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
                Lic.Expiration = ActiveLockDate(ExpireDate).ToString("yyyy/MM/dd")
            End If
            Lic.MaxCount = CInt(CStr(modActiveLock.LoByte(UserData)))

            ' Finally check if the serial number is Ok
            ' Shortkeys are no longer using the HDD firmware serial number
            ' they are using the Computer Fingerprint after v3.6
            If Not oReg.IsKeyOK(SerialNumber, mSoftwareName & mSoftwareVer & mSoftwarePassword, modHardware.GetFingerprint()) Then
                ' Something wrong with the serial number used
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
            End If
            Lic.LicenseCode = "Short Key"
            '"Key is valid."
        Else
            'MsgBox "Invalid license key."
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
        End If

        ' Check if license has not expired
        ' but don't do it if there's no expiration date
        If Lic.Expiration = "" Then Exit Sub
        Dim dtExp As Date
        dtExp = ActiveLockDate(CDate(Lic.Expiration))
        If ActiveLockDate(Date.UtcNow) > dtExp And Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
            ' ialkan - 9-23-2005 added the following to update and store the license
            ' with the new LastUsed property; otherwise setting the clock back next time
            ' might bypass the protection mechanism
            ' Update last used date
            UpdateLastUsed(Lic)
            mKeyStore.Store(Lic, mLicenseFileType)
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseExpired, ACTIVELOCKSTRING, STRLICENSEEXPIRED)
        End If
	End Sub

	''' <summary>
	''' Validates the entire license (including lastused, etc.)
	''' </summary>
	''' <param name="Lic">ProductLicense - Product License</param>
	''' <remarks></remarks>
    Private Sub ValidateLic(ByRef Lic As ProductLicense)

        ' Get the current date format and save it to regionalSymbol variable
        Get_locale()
        ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
        Set_locale("")

        ' make sure we're initialized.
        If Not mfInit Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNotInitialized, ACTIVELOCKSTRING, STRNOTINITIALIZED)
        End If

        ' validate license key first
        If Mid(Lic.LicenseKey, 5, 1) = "-" And Mid(Lic.LicenseKey, 10, 1) = "-" And Mid(Lic.LicenseKey, 15, 1) = "-" And Mid(Lic.LicenseKey, 20, 1) = "-" Then
            Dim arrProdVer() As String
            Dim actualLicensee As String
            arrProdVer = Lic.Licensee.Split("&&&")
            actualLicensee = arrProdVer(0)
            ValidateShortKey(Lic, actualLicensee)
            'ValidateShortKey(Lic, Lic.Licensee)
        Else 'ALCrypto RSA key
            ValidateKey(Lic)
        End If

        Dim strEncrypted, strHash As String
        ' Validate last run date
        strEncrypted = Lic.LastUsed
        MyNotifier.Notify("ValidateValue", strEncrypted)
        strHash = modMD5.Hash(strEncrypted)
        If strHash <> Lic.Hash1 Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseTampered, ACTIVELOCKSTRING, STRLICENSETAMPERED)
        End If

        ' We will compare several important dates with each other 
        ' to see if anything is wrong in the license
        If Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
            ' Must have NOW>LASTUSED
            If DateValue(ActiveLockDate(Date.UtcNow)) < DateValue(ActiveLockDate(CDate(Lic.LastUsed))) Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
            End If
            ' Must have NOW<EXPIRATION
            If DateValue(ActiveLockDate(Date.UtcNow)) > DateValue(ActiveLockDate(CDate(Lic.Expiration))) Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseExpired, ACTIVELOCKSTRING, STRLICENSEEXPIRED)
            End If
            ' Must have LASTUSED>=REGISTERED
            If DateValue(ActiveLockDate(Lic.LastUsed)) < DateValue(ActiveLockDate(CDate(Lic.RegisteredDate))) Then
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
            End If
        End If
        UpdateLastUsed(Lic)
        mKeyStore.Store(Lic, mLicenseFileType)
        Set_locale(regionalSymbol)
	End Sub

	''' <summary>
	''' Updates LastUsed property with current date stamp.
	''' </summary>
	''' <param name="Lic">ProductLicense - Product License</param>
	''' <remarks></remarks>
    Private Sub UpdateLastUsed(ByRef Lic As ProductLicense)
        ' Update license store with LastRunDate
        Dim strLastUsed As String
        Set_locale("")
        strLastUsed = ActiveLockDate(Date.UtcNow)
        strLastUsed = ActiveLockDate(CDate(strLastUsed))
        Lic.LastUsed = strLastUsed
        MyNotifier.Notify("ValidateValue", strLastUsed)
        Lic.Hash1 = modMD5.Hash(strLastUsed)
	End Sub

	''' <summary>
	''' Registers Activelock license with a given liberation key
	''' </summary>
	''' <param name="LibKey">String - Liberation Key</param>
	''' <param name="user">Optional - String - User</param>
	''' <remarks></remarks>
    Private Sub IActiveLock_Register(ByVal LibKey As String, Optional ByRef user As String = "") Implements _IActiveLock.Register

        Dim Lic As ActiveLock3_6NET.ProductLicense = New ActiveLock3_6NET.ProductLicense
        Dim varResult As Object
        Dim trialStatus As Boolean

        ' Get the current date format and save it to regionalSymbol variable
        Get_locale()
        ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
        Set_locale("")

        ' Check to see if this is a Short License Key
        If Mid(LibKey, 5, 1) = "-" And Mid(LibKey, 10, 1) = "-" And Mid(LibKey, 15, 1) = "-" And Mid(LibKey, 20, 1) = "-" Then
            Lic.LicenseKey = UCase(LibKey)
            ValidateShortKey(Lic, user)
        Else ' RSA key
            Lic.Load(LibKey)
            ' Validate that the license key.
            '   - registered user
            '   - expiry date
            ValidateKey(Lic)
        End If

        ' License was validated successfuly. Check clock tampering for non-permanent licenses.
        If Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
            If mCheckTimeServerForClockTampering = IActiveLock.ALTimeServerTypes.alsCheckTimeServer Then
                If SystemClockTampered() Then
                    Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                End If
            End If
            If mChecksystemfilesForClockTampering = IActiveLock.ALSystemFilesTypes.alsCheckSystemFiles Then
                If ClockTampering() Then
                    Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, STRCLOCKCHANGED)
                End If
            End If
        End If

        ' License was validated successfuly.  Store it.
        If mKeyStore Is Nothing Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrKeyStoreInvalid, ACTIVELOCKSTRING, STRKEYSTOREUNINITIALIZED)
        End If

        ' Update last used date
        UpdateLastUsed(Lic)
        mKeyStore.Store(Lic, mLicenseFileType)

        ' This works under NTFS and is needed to prevent clock tampering
        If CheckStreamCapability() And Lic.LicenseType <> ProductLicense.ALLicType.allicPermanent Then
            ' Write the current date and time into the ads
            Dim ok As Integer
            Dim strStream As String = String.Empty
            strStream = mSoftwareName & mSoftwareVer & mSoftwarePassword
            ok = ADSFile.Write(ActiveLockDate(Date.UtcNow).ToString, mKeyStorePath, strStream)
        End If

        ' Expire all trial licenses
        On Error Resume Next
        ' Expire the Trial
        If mTrialType <> IActiveLock.ALTrialTypes.trialNone Then
            trialStatus = ExpireTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, mSoftwarePassword)
        End If
        Set_locale(regionalSymbol)

	End Sub

	''' <summary>
	''' Kills a Trial License
	''' </summary>
	''' <remarks></remarks>
    Private Sub IActiveLock_KillTrial() Implements _IActiveLock.KillTrial
        On Error Resume Next
        'Expire the Trial
        Dim trialStatus As Boolean
        If mSoftwareName = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareName, ACTIVELOCKSTRING, STRNOSOFTWARENAME)
        ElseIf mSoftwareVer = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareVersion, ACTIVELOCKSTRING, STRNOSOFTWAREVERSION)
        ElseIf mSoftwarePassword = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwarePassword, ACTIVELOCKSTRING, STRNOSOFTWAREPASSWORD)
        Else
            trialStatus = ExpireTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, mSoftwarePassword)
        End If
	End Sub

	' TODO: ActiveLock.vb - Function IActiveLock_GenerateShortKey - Needs Documentation!
	''' <summary>
	''' Not Documented!
	''' </summary>
	''' <param name="SoftwareCode"></param>
	''' <param name="SerialNumber"></param>
	''' <param name="LicenseeAndRegisteredLevel"></param>
	''' <param name="Expiration"></param>
	''' <param name="LicType"></param>
	''' <param name="RegisteredLevel"></param>
	''' <param name="MaxUsers"></param>
	''' <returns></returns>
	''' <remarks></remarks>
    Private Function IActiveLock_GenerateShortKey(ByVal SoftwareCode As String, ByVal SerialNumber As String, ByVal LicenseeAndRegisteredLevel As String, ByVal Expiration As String, ByVal LicType As ProductLicense.ALLicType, ByVal RegisteredLevel As Integer, Optional ByVal MaxUsers As Short = 1) As String Implements _IActiveLock.GenerateShortKey

        On Error GoTo ErrHandler

        Dim m_Key As clsShortLicenseKey
        m_Key = New clsShortLicenseKey

        m_Key.AddSwapBits(0, 0, 1, 0)
        m_Key.AddSwapBits(0, 2, 1, 1)
        m_Key.AddSwapBits(0, 4, 2, 0)
        m_Key.AddSwapBits(0, 5, 2, 1)
        m_Key.AddSwapBits(2, 0, 3, 0)
        m_Key.AddSwapBits(2, 6, 3, 1)
        m_Key.AddSwapBits(2, 7, 1, 3)

        Dim oReg As clsShortSerial
        oReg = New clsShortSerial
        Dim sKey As String
        Dim m_ProdCode As Integer

        sKey = oReg.GenerateKey("", Left(SoftwareCode, Len(SoftwareCode) - 2)) 'Do not include the last 2 possible == paddings
        m_ProdCode = CInt(Left(sKey, 4))

        ' create a new key
        IActiveLock_GenerateShortKey = m_Key.CreateShortKey(SerialNumber, LicenseeAndRegisteredLevel, m_ProdCode, CDate(Expiration), MakeWord(CStr(MaxUsers), CStr(LicType)), RegisteredLevel)

        Exit Function
ErrHandler:
        oReg = Nothing
        m_Key = Nothing

	End Function

	''' <summary>
	''' Resets a Trial License
	''' </summary>
	''' <remarks></remarks>
    Private Sub IActiveLock_ResetTrial() Implements _IActiveLock.ResetTrial
        On Error Resume Next
        'Reset the Trial
        Dim trialStatus As Boolean
        If mSoftwareName = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareName, ACTIVELOCKSTRING, STRNOSOFTWARENAME)
        ElseIf mSoftwareVer = "" Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareVersion, ACTIVELOCKSTRING, STRNOSOFTWAREVERSION)
        Else
            trialStatus = ResetTrial(mSoftwareName, mSoftwareVer, mTrialType, mTrialLength, mTrialHideTypes, mSoftwarePassword)
        End If
	End Sub

	''' <summary>
	''' Returns the lock code from a given Activelock license
	''' </summary>
	''' <param name="Lic">ProductLicense - Product License</param>
	''' <returns>String - Lock code</returns>
	''' <remarks>v3 includes the new lockHDFirmware option</remarks>
    Private Function IActiveLock_LockCode(Optional ByRef Lic As ProductLicense = Nothing) As String Implements _IActiveLock.LockCode
        Dim strLock As String = String.Empty
        Dim noKey As String
        Dim userFromInstallCode, usedcode As String
        'Dim tmpLockType As IActiveLock.ALLockTypes
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
                AppendLockString(strLock, modHardware.GetBiosVersion())
                AppendLockString(strLock, modHardware.GetMotherboardSerial())
                AppendLockString(strLock, modHardware.GetIPaddress())
                AppendLockString(strLock, modHardware.GetExternalIP())
                AppendLockString(strLock, modHardware.GetFingerprint())
                AppendLockString(strLock, modHardware.GetMemoryID())
                AppendLockString(strLock, modHardware.GetCPUID())
                AppendLockString(strLock, modHardware.GetBaseBoardID())
                AppendLockString(strLock, modHardware.GetVideoID())
            Else
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockMAC) Then    'mLockTypes And IActiveLock.ALLockTypes.lockMAC Then
                    AppendLockString(strLock, modHardware.GetMACAddress())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockComp) Then    'mLockTypes And IActiveLock.ALLockTypes.lockComp Then
                    AppendLockString(strLock, modHardware.GetComputerName())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockHD) Then    'mLockTypes And IActiveLock.ALLockTypes.lockHD Then
                    AppendLockString(strLock, modHardware.GetHDSerial())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockHDFirmware) Then    'mLockTypes And IActiveLock.ALLockTypes.lockHDFirmware Then
                    AppendLockString(strLock, modHardware.GetHDSerialFirmware())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockWindows) Then    'mLockTypes And IActiveLock.ALLockTypes.lockWindows Then
                    AppendLockString(strLock, modHardware.GetWindowsSerial())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockBIOS) Then    'mLockTypes And IActiveLock.ALLockTypes.lockBIOS Then
                    AppendLockString(strLock, modHardware.GetBiosVersion())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockMotherboard) Then    'mLockTypes And IActiveLock.ALLockTypes.lockMotherboard Then
                    AppendLockString(strLock, modHardware.GetMotherboardSerial())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockIP) Then    'mLockTypes And IActiveLock.ALLockTypes.lockIP Then
                    AppendLockString(strLock, modHardware.GetIPaddress())
                Else
                    AppendLockString(strLock, noKey)
                End If
                ' new in v3.6
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockExternalIP) Then
                    AppendLockString(strLock, modHardware.GetExternalIP())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockFingerprint) Then
                    AppendLockString(strLock, modHardware.GetFingerprint())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockMemory) Then
                    AppendLockString(strLock, modHardware.GetMemoryID())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockCPUID) Then
                    AppendLockString(strLock, modHardware.GetCPUID())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockBaseboardID) Then
                    AppendLockString(strLock, modHardware.GetBaseBoardID())
                Else
                    AppendLockString(strLock, noKey)
                End If
                If IsNumberIncluded(mLockTypes, IActiveLock.ALLockTypes.lockVideoID) Then
                    AppendLockString(strLock, modHardware.GetVideoID())
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
            'tmpLockType = IActiveLock.ALLockTypes.lockNone ' lockNone = 0 so starting value

            ReDim mUsedLockTypes(0)  ' remove all previous
            Dim SizeLockType As Integer  ' use to build up LockCode.
            SizeLockType = 0

            If Lic.LicenseCode <> "" Then
                If Left(Lic.LicenseCode, 1) = "+" Then
                    usedcode = modBase64.Base64_Decode(Mid(Lic.LicenseCode, 2))
                    'bLockNone = True ' per David Weatherall
                    IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockNone, SizeLockType)  'dw1 build up lockTypes - start with lockNone
                Else
                    usedcode = modBase64.Base64_Decode((Lic.LicenseCode))
                    'bLockNone = False ' per David Weatherall
                End If
                a = Split(usedcode, vbLf)
                For j = LBound(a) To UBound(a) - 1
                    aString = a(j)
                    If Left(aString, 1) = "+" Then aString = Mid(aString, 2)
                    If j = LBound(a) Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockMAC, SizeLockType)
                            If aString <> modHardware.GetMACAddress() Then
                                ' Ok MAC address did not match
                                ' Maybe the laptop owner turned on the wireless connection
                                ' and it wasn't on when the license was registered
                                If CheckMACaddress(aString) = False Then
                                    ' we truly failed
                                    Set_locale(regionalSymbol)
                                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                                End If
                            End If
                        End If
                    ElseIf j = LBound(a) + 1 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockComp, SizeLockType)
                            If aString <> modHardware.GetComputerName() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 2 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockHD, SizeLockType)
                            If aString <> modHardware.GetHDSerial() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 3 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockHDFirmware, SizeLockType)
                            If aString <> modHardware.GetHDSerialFirmware() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 4 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockWindows, SizeLockType)
                            If aString <> modHardware.GetWindowsSerial() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 5 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockBIOS, SizeLockType)
                            If aString <> modHardware.GetBiosVersion() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 6 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockMotherboard, SizeLockType)
                            If aString <> modHardware.GetMotherboardSerial() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 7 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockIP, SizeLockType)
                            Dim returnedIP As String = modHardware.GetIPaddress()
                            If returnedIP = "-1" Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrInternetConnectionError, ACTIVELOCKSTRING, STRINTERNETNOTCONNECTED)
                            End If
                            If aString <> returnedIP Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrWrongIPaddress, ACTIVELOCKSTRING, STRWRONGIPADDRESS)
                            End If
                        End If
                        'added v3.6
                    ElseIf j = LBound(a) + 8 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockExternalIP, SizeLockType)
                            If aString <> modHardware.GetExternalIP() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 9 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockFingerprint, SizeLockType)
                            If aString <> modHardware.GetFingerprint() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 10 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockMemory, SizeLockType)
                            If aString <> modHardware.GetMemoryID() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 11 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockCPUID, SizeLockType)
                            If aString <> modHardware.GetCPUID() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 12 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockBaseboardID, SizeLockType)
                            If aString <> modHardware.GetBaseBoardID() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                            End If
                        End If
                    ElseIf j = LBound(a) + 13 Then
                        If aString <> noKey Then
                            IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockvideoID, SizeLockType)
                            If aString <> modHardware.GetVideoID() Then
                                Set_locale(regionalSymbol)
                                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
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
                    Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                End If

                usedcode = Mid(usedcode, 1, Len(usedcode) - Len(userFromInstallCode) - 1)
                IActiveLock_LockCode = Lic.ToString_Renamed() & vbLf & usedcode
            Else
                Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID)
                Return Nothing
            End If
        End If
	End Function

	''' <summary>
	''' Appends the lock string to the given installation code
	''' </summary>
	''' <param name="strLock">String - The lock string to be appended to, returns as an output</param>
	''' <param name="newSubString">String - The string to be appended to the lock string if strLock is empty string</param>
	''' <remarks></remarks>
    Private Sub AppendLockString(ByRef strLock As String, ByVal newSubString As String)
        If strLock = "" Then
            strLock = newSubString
        Else
            strLock = strLock & vbLf & newSubString
        End If
	End Sub

	''' <summary>
	''' Not implemented yet
	''' </summary>
	''' <param name="OtherSoftwareCode">String - Installation code from another machine/software</param>
	''' <returns></returns>
	''' <remarks>Transfers an Activelock license from one machine/software to another</remarks>
    Private Function IActiveLock_Transfer(ByVal OtherSoftwareCode As String) As String Implements _IActiveLock.Transfer
		' TODO: ActiveLock.vb - Function IActiveLock_Transfer - Implement me!
        Set_locale(regionalSymbol)
        Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNotImplemented, ACTIVELOCKSTRING, STRNOTIMPLEMENTED)
        Return Nothing
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
	''' <summary>
	''' Generates a Short Key (Serial Number)
	''' </summary>
	''' <param name="HDDfirmwareSerial"></param>
	''' <returns></returns>
	''' <remarks></remarks>
    Private Function IActivelock_GenerateShortSerial(ByVal HDDfirmwareSerial As String) As String Implements _IActiveLock.GenerateShortSerial
        Dim oReg As clsShortSerial
        Dim sKey As String

        oReg = New clsShortSerial
        sKey = oReg.GenerateKey(mSoftwareName & mSoftwareVer & mSoftwarePassword, HDDfirmwareSerial)
        IActivelock_GenerateShortSerial = sKey
        ' If longer serial is used, possible to break up into sections
        'Left(sKey, 4) & "-" & Mid(sKey, 5, 4) & "-" & Mid(sKey, 9, 4) & "-" & Mid(sKey, 13, 4)

        oReg = Nothing
	End Function

	' TODO: ActiveLock.vb - Function specialChar - Add documentation - Not Documented!
	''' <summary>
	''' Not Documented!
	''' </summary>
	''' <param name="s"></param>
	''' <returns></returns>
	''' <remarks></remarks>
    Private Function specialChar(ByVal s As String) As Boolean
        Dim k As Integer
        s = s & Space(1) 'check against null-strings
        For k = 1 To Len(s)
            Select Case Asc(Mid$(s, k))
                Case 32 To 126
                    'continue
                Case Else
                    specialChar = True
                    Exit Function
            End Select
        Next k
    End Function


End Class