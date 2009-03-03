Option Strict Off
Option Explicit On

#Region "Copyright"
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
#End Region

#Region "Changes"
'* Reserverd for documenting changes
'*
'* japreja - March 03, 2009 - Updated comments for Intellisense & Object browser
'*                            Comments should now be available when referancing the DLL
'*                            Updated SVN Between 12:30PM and 1:30PM
'*
#End Region

''' <summary>
''' _IActiveLock - Interface - Implimented by IActiveLock
''' </summary>
''' <remarks>
''' <para> - MaintainedBy:</para>
''' <para> - LastRevisionDate:</para>
''' <para> - Comments:</para></remarks>
Public Interface _IActiveLock
    ReadOnly Property RemainingTrialDays() As Integer
    ReadOnly Property RemainingTrialRuns() As Integer
    ReadOnly Property MaxCount() As Integer
    ReadOnly Property RegisteredLevel() As String
    ReadOnly Property LicenseClass() As String
    Property LockType() As IActiveLock.ALLockTypes
    WriteOnly Property LicenseKeyType() As IActiveLock.ALLicenseKeyTypes
    ReadOnly Property UsedLockType() As Integer
    Property TrialHideType() As IActiveLock.ALTrialHideTypes
    Property TrialType() As IActiveLock.ALTrialTypes
    Property TrialLength() As Integer
    Property SoftwareName() As String
    Property SoftwarePassword() As String
    WriteOnly Property CheckTimeServerForClockTampering() As IActiveLock.ALTimeServerTypes
    WriteOnly Property CheckSystemFilesForClockTampering() As IActiveLock.ALSystemFilesTypes
    Property LicenseFileType() As IActiveLock.ALLicenseFileTypes
    WriteOnly Property AutoRegister() As IActiveLock.ALAutoRegisterTypes
    WriteOnly Property TrialWarning() As IActiveLock.ALTrialWarningTypes
    WriteOnly Property SoftwareCode() As String
	Property SoftwareVersion() As String

    WriteOnly Property KeyStoreType() As IActiveLock.LicStoreType
    WriteOnly Property KeyStorePath() As String
    ReadOnly Property InstallationCode(Optional ByVal User As String = vbNullString, Optional ByVal Lic As ProductLicense = Nothing) As String
    WriteOnly Property AutoRegisterKeyPath() As String
    Function LockCode(Optional ByRef Lic As ProductLicense = Nothing) As String
    Sub Register(ByVal LibKey As String, Optional ByRef user As String = "")
    Function Transfer(ByVal InstallCode As String) As String
    Sub Init(Optional ByVal strPath As String = "", Optional ByRef autoLicString As String = "")
    Sub Acquire(Optional ByRef strMsg As String = "", Optional ByRef strRemainingTrialDays As String = "", Optional ByRef strRemainingTrialRuns As String = "", Optional ByRef strTrialLength As String = "", Optional ByRef strUsedDays As String = "", Optional ByRef strExpirationDate As String = "", Optional ByRef strRegisteredUser As String = "", Optional ByRef strRegisteredLevel As String = "", Optional ByRef strLicenseClass As String = "", Optional ByRef strMaxCount As String = "", Optional ByRef strLicenseFileType As String = "", Optional ByRef strLicenseType As String = "", Optional ByRef strUsedLockType As String = "")
    Sub ResetTrial()
    Sub KillTrial()
    Function GenerateShortSerial(ByVal HDDfirmwareSerial As String) As String
    Function GenerateShortKey(ByVal SoftwareCode As String, ByVal SerialNumber As String, ByVal LicenseeAndRegisteredLevel As String, ByVal Expiration As String, ByVal LicType As ProductLicense.ALLicType, ByVal RegisteredLevel As Integer, Optional ByVal MaxUsers As Short = 1) As String
    ReadOnly Property EventNotifier() As ActiveLockEventNotifier
    ReadOnly Property UsedDays() As Integer
    ReadOnly Property RegisteredDate() As String
    ReadOnly Property RegisteredUser() As String
    ReadOnly Property ExpirationDate() As String
End Interface

''' <summary>
''' IActiveLock - Impliments _IActiveLock
''' </summary>
''' <remarks>Class instancing was changed to public.</remarks>
<System.Runtime.InteropServices.ProgId("IActiveLock_NET.IActiveLock")> Public Class IActiveLock
    Implements _IActiveLock

#Region "Notes"
    '===============================================================================
    ' Name: IActivelock
    ' Purpose: This is the main interface into ActiveLock&#39;s functionalities.
    ' The user application interacts with ActiveLock primarily through this IActiveLock interface.
    ' Typically, the application would obtain an instance of this interface via the
    ' <a href="Globals.NewInstance.html">ActiveLock3.NewInstance()</a> accessor method. From there, initialization calls are done,
    ' and then various method such as <a href="IActiveLock.Register.html">Register()</a>, <a href="IActiveLock.Acquire.html">Acquire()</a>, etc..., can be used.
    ' <p>
    ' ActiveLock also sends COM event notifications to the user application whenever it needs help to perform
    ' some action, such as license property validation/encryption.  The user application can intercept
    ' these events via the ActiveLockEventNotifier object, which can be obtained from
    ' <a href="IActiveLock.Get.EventNotifier.html">IActiveLock.EventNotifier</a> property.
    ' <p>
    ' <b>Important Note</b><br>
    ' The user application is strongly advised to perform a checksum on the
    ' ActiveLock DLL prior to accessing and interacting with ActiveLock. Using the checksum, you can tell if
    ' the DLL has been tampered. Please refer to sample code below on how the checksumming can be done.
    ' <p>
    ' The sample code fragments below illustrate the typical usage flow between your application and ActiveLock.
    ' Please note that the code shown is only for illustration purposes and is not meant to be a complete
    ' compilable program. You may have to add variable declarations and function definitions around the code
    ' fragments before you can compile it.
    ' <p>
    ' <pre>
    '   Form1.frm:
    '   ...
    '   Private MyActiveLock As ActiveLock3.IActiveLock
    '   Private WithEvents ActiveLockEventSink As ActiveLockEventNotifier
    '   Private Const AL_CRC& = 123+ &#39; ActiveLock3.dll&#39;s CRC checksum to be used for validation
    '
    '   &#39; This key will be used to set <a href="IActiveLock.Let.SoftwareCode.html">IActiveLock.SoftwareCode</a> property.
    '   &#39; NOTE: This is NOT a complete key (complete key is to long to put in documentation).
    '   &#39; You will generate your own product code using ALUGEN.  This is the <code>VCode</code> generated
    '   by ALUGEN.
    '   Private Const PROD_CODE$ = "AAAAB3NzaC1yc2EAAAABJQAAAIBZnXD4IKfrBH25ekwLWQMs5mJ..."
    '
    '   ....
    '
    '   Private Sub Form_Load()
    '       On Error GoTo ErrHandler
    '       &#39; Obtain an instance of AL. NewInstance() also calls IActiveLock.Init()
    '       Set MyActiveLock = ActiveLock3.NewInstance()
    '       &#39; At this point, either AL has been initialized or an error would have already been raised
    '       &#39; if there were problems (such as ActiveLock3.dll having been tampered).
    '
    '       &#39; Verify AL&#39;s authenticity
    '       &#39; modActiveLock.CRCCheckSumTypeLib() requires a public-creatable object to be passed in
    '       &#39; so that it can determine the Type Library DLL on which to perform the checksum.
    '       &#39; So can&#39;t use MyActiveLock object to authenticate since it is not a public creatable object.
    '       &#39; So we&#39;ll use ActiveLock3.Globals, which is just as good because they are in the same DLL.
    '       Dim crc As Long
    '       crc = CRCCheckSumTypeLib(New ActiveLock3.Globals)
    '       Debug.Print "Hash: " & crc
    '       If crc &lt;&gt; AL_CRC Then
    '           MsgBox "ActiveLock3.dll has been corrupted."
    '           End ' terminate
    '       End If
    '
    '      &#39; Initialize the keystore. We use a File keystore in this case.
    '      MyActiveLock.KeyStoreType = alsFile
    '      MyActiveLock.KeyStorePath = App.path & "\myapp.lic"
    '
    '      &#39; Obtain the EventNotifier so that we can receive notifications from AL.
    '      Set ActiveLockEventSink = MyActiveLock.EventNotifier
    '
    '      &#39; Specify the name of the product that will be locked through AL.
    '      MyActiveLock.SoftwareName = "MyApp"
    '
    '      &#39; Specify our product code.  This code will be used later by ActiveLock to validate license keys.
    '      MyActiveLock.SoftwareCode = PROD_CODE
    '
    '      &#39; Specify product version
    '       MyActiveLock.SoftwareVersion = txtVersion
    '
    '      &#39; Specify Lock Type
    '      MyActiveLock.LockType = lockHD
    '
    '      &#39; Sets path to liberation key file for automatic registration
    '      MyActiveLock.AutoRegisterKeyPath = App.path & "\myapp.alb"
    '
    '      &#39; Initialize the instance
    '      MyActiveLock.Init
    '
    '      &#39; Check registration status by calling Acquire()
    '      &#39; Note: Calling Acquire() may trigger ActiveLockEventNotifier_ValidateValue() event.
    '      &#39; So we should be prepared to handle that.
    '      MyActiveLock.Acquire
    '
    '      &#39; By now, if the product is not registered, then an error would have been raised,
    '      &#39; which means if we get to here, then we're registered.
    '
    '      &#39; Just for fun, print out some registration status info
    '      Debug.Print "Registered User: " & MyActiveLock.RegisteredUser
    '      Debug.Print "Used Days: " & MyActiveLock.UsedDays
    '      Debug.Print "Expiration Date: " & MyActiveLock.ExpirationDate
    '      Exit Sub
    '   ErrHandler:
    '       MsgBox Err.Number & ": " & Err.Description
    '       &#39; End program
    '       End
    '   End Sub
    '   ...
    '   <p>
    '   (Optional) ActiveLock raises this event typically when it needs a value to be encrypted.
    '   We can use any kind of encryption we&#39;d like here, as long as it&#39;s deterministic.
    '   i.e. there&#39;s a one-to-one correspondence between unencrypted value and encrypted value.
    '   NOTE: BlowFish is NOT an example of deterministic encryption so you can&#39;t use it here.
    '   You are allowed to use asymmetric algorithm since you will never be asked to decrypt a value,
    '   only to encrypt.
    '   You don&#39;t have to handle this event if you don&#39;t want to; it just means that the value WILL NOT
    '   be encrypted when it is saved to the keystore.
    '
    '   Private Sub ActiveLockEventSink_ValidateValue(ByVal Value As String, Result As String)
    '       Result = Encrypt(Value)
    '   End Sub
    '<br>
    '   &#39; Roll our own simple-yet-weird encryption routine.
    '   &#39; Must keep in mind that our encryption algorithm must be deterministic.
    '   &#39; In other words, given the same uncrypted string, it must always yield the same encrypted string.
    '   Private Function Encrypt(strData As String) As String
    '       Dim i&, n&
    '       dim sResult$
    '       n = Len(strData)
    '       For i = 1 to n
    '           sResult = sResult & Asc(Mid$(strData, i, 1)) * 7
    '       Next i
    '       Encrypt = sResult
    '   End Function
    '   ...
    '   &#39; Returns the CRC checksum of the ActiveLock3.dll.
    '   Private Property Get ALCRC() As Long
    '       &#39; Don&#39;t just return a single value, but rather compute it using some simple arithmetic
    '       &#39; so that hackers can&#39;t easily find it with a hex editor.
    '       &#39; Of course, the values below will not make up the real checksum. For the most up-to-date
    '       &#39; checksum, please refer to the ActiveLock Release Notes.
    '       ALCRC = 123 + 456
    '   End Property
    ' </pre>
    '
    ' <p>Generating registration code from the user application to be sent to the vendor in exchange for
    ' a liberation key.
    ' <pre>
    '    &#39; Generate Installation code
    '    Dim strInstCode As String
    '    strInstCode = MyActiveLock.InstallCode(txtUser)
    '    &#39; strInstCode now contains the request code to be sent to the vendor for activation.
    ' </pre>
    '
    ' <p>Key Registration functionality - register using a liberation key.
    ' <pre>
    '    On Error GoTo ErrHandler
    '    &#39; Register this key
    '    &#39; txtLibKey contains the liberation key entered by the user.
    '    &#39; This key could have be sent via an email to the user or a program that automatically
    '    &#39; requests the key from a registration website.
    '    MyActiveLock.Register txtLibKey
    '    MsgBox "Registration successful!"
    '    Exit Sub
    'ErrHandler:
    '    MsgBox Err.Number & ": " & Err.Description
    ' </pre>
    ' Remarks:
    ' Functions:
    ' Properties:
    ' Methods:
    ' Started: 21.04.2005
    ' Modified: 08.05.2005
    '===============================================================================
    ' @author activelock-admins
    ' @version 3.3.0
    ' @date 3-23-2006
#End Region

#Region "Public Enums"

    ''' <summary>
    ''' License Lock Types.
    ''' </summary>
    ''' <remarks>Values can be combined (OR ed) together.</remarks>
    Public Enum ALLockTypes
        ''' <summary>No locking - not recommended</summary>
        ''' <remarks></remarks>
        lockNone = 0
        ''' <summary>
        ''' Lock to Network Interface Card Address
        ''' </summary>
        ''' <remarks></remarks>
        lockMAC = 1             '8
        ''' <summary>
        ''' Lock to Computer Name
        ''' </summary>
        ''' <remarks></remarks>
        lockComp = 2
        ''' <summary>
        ''' Lock to Hard Drive Serial Number (Volume Serial Number)
        ''' </summary>
        ''' <remarks></remarks>
        lockHD = 4
        ''' <summary>
        ''' Lock to Hard Disk Firmware Serial (HDD Manufacturer's Serial Number)
        ''' </summary>
        ''' <remarks></remarks>
        lockHDFirmware = 8      '256
        ''' <summary>
        ''' Lock to Windows Serial Number
        ''' </summary>
        ''' <remarks></remarks>
        lockWindows = 16        '1
        ''' <summary>
        ''' Lock to BIOS Serial Number
        ''' </summary>
        ''' <remarks></remarks>
        lockBIOS = 32           '16
        ''' <summary>
        ''' Lock to Motherboard Serial Number
        ''' </summary>
        ''' <remarks></remarks>
        lockMotherboard = 64
        ''' <summary>
        ''' Lock to Computer Local IP Address
        ''' </summary>
        ''' <remarks></remarks>
        lockIP = 128            '32
        ''' <summary>
        ''' Lock to External IP Address
        ''' </summary>
        ''' <remarks></remarks>
        lockExternalIP = 256    '128
        ''' <summary>
        ''' Lock to Fingerprint (Activelock Combination)
        ''' </summary>
        ''' <remarks></remarks>
        lockFingerprint = 512
        ''' <summary>
        ''' Lock to Memory ID
        ''' </summary>
        ''' <remarks></remarks>
        lockMemory = 1024
        ''' <summary>
        ''' Lock to CPU ID
        ''' </summary>
        ''' <remarks></remarks>
        lockCPUID = 2048
        ''' <summary>
        ''' Lock to Baseboard Name and Serial Number
        ''' </summary>
        ''' <remarks></remarks>
        lockBaseboardID = 4096
        ''' <summary>
        ''' Lock to Video Controller Name and Drive Version Number
        ''' </summary>
        ''' <remarks></remarks>
        lockvideoID = 8192
    End Enum

    ''' <summary>
    ''' License Key Type specifies the length/type
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ALLicenseKeyTypes
        ''' <summary>
        ''' 1024-bit Long keys by RSA via ALCrypto DLL
        ''' </summary>
        ''' <remarks></remarks>
        alsRSA = 0
        ''' <summary>
        ''' Short license keys by MD5
        ''' </summary>
        ''' <remarks></remarks>
        alsShortKeyMD5 = 1
    End Enum

    ''' <summary>
    ''' License Store Type specifies where to store license keys
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum LicStoreType
        ''' <summary>
        ''' Store in Windows Registry
        ''' </summary>
        ''' <remarks></remarks>
        alsRegistry = 0
        ''' <summary>
        ''' Store in a license file
        ''' </summary>
        ''' <remarks></remarks>
        alsFile = 1
    End Enum

    ''' <summary>
    ''' Products Store Type specifies where to store products infos
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ProductsStoreType
        ''' <summary>
        ''' Store in INI file (licenses.ini)
        ''' </summary>
        ''' <remarks></remarks>
        alsINIFile = 0
        ''' <summary>
        ''' Store in XML file (licenses.xml)
        ''' </summary>
        ''' <remarks></remarks>
        alsXMLFile = 1
        ''' <summary>
        ''' Store in MDB file (licenses.mdb)
        ''' </summary>
        ''' <remarks>mdb file should contain a table named products with structure: ID(autonumber), name(text,150), version (text,50), vccode(memo), gcode(memo)</remarks>
        alsMDBFile = 2
        ' TODO: IActiveLock.vb - Enum ProductStoreType - Store in MSSQL database 'not implemented
        '''' <summary>
        '''' TODO: Store in MSSQL database 'not implemented
        '''' </summary>
        '''' <remarks></remarks>
        'alsMSSQL = 3 'TODO
    End Enum

    ''' <summary>
    ''' Trial Type specifies what kind of Trial Feature is used
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ALTrialTypes
        ''' <summary>
        ''' No trial used
        ''' </summary>
        ''' <remarks></remarks>
        trialNone = 0
        ''' <summary>
        ''' Trial by Days
        ''' </summary>
        ''' <remarks></remarks>
        trialDays = 1
        ''' <summary>
        ''' Trial by Runs
        ''' </summary>
        ''' <remarks></remarks>
        trialRuns = 2
    End Enum

    ''' <summary>
    ''' Trial Hide Mode Type specifies what kind of Trial Hiding Mode is used
    ''' </summary>
    ''' <remarks>Values can be combined (OR'ed) together.</remarks>
    Public Enum ALTrialHideTypes
        ''' <summary>
        ''' Trial information is hidden in BMP files
        ''' </summary>
        ''' <remarks></remarks>
        trialSteganography = 1
        ''' <summary>
        ''' Trial information is hidden in a folder which uses a default namespace
        ''' </summary>
        ''' <remarks></remarks>
        trialHiddenFolder = 2
        ''' <summary>
        ''' Trial information is encrypted and hidden in registry (per user)
        ''' </summary>
        ''' <remarks></remarks>
        trialRegistryPerUser = 4
        ' TODO: IActiveLock.vb - Enum ALTrialHideYypes - Please update this comment
        ''' <summary>
        ''' Not documented! Please Update!
        ''' </summary>
        ''' <remarks></remarks>
        trialIsolatedStorage = 8
    End Enum

    ''' <summary>
    ''' Enum for accessing the Time Server to check Clock Tampering
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ALTimeServerTypes
        ''' <summary>
        ''' Skips checking a Time Server
        ''' </summary>
        ''' <remarks></remarks>
        alsDontCheckTimeServer = 0
        ''' <summary>
        ''' Checks a Time Server
        ''' </summary>
        ''' <remarks></remarks>
        alsCheckTimeServer = 1
    End Enum

    ''' <summary>
    ''' Enum for scanning the system folders/files to detect clock tampering
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ALSystemFilesTypes
        ''' <summary>
        ''' Skips checking system files
        ''' </summary>
        ''' <remarks></remarks>
        alsDontCheckSystemFiles = 0
        ''' <summary>
        ''' Checks system files
        ''' </summary>
        ''' <remarks></remarks>
        alsCheckSystemFiles = 1
    End Enum

    ''' <summary>
    ''' Enum for license file encryption
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ALLicenseFileTypes
        ''' <summary>
        ''' Encrypts the license file
        ''' </summary>
        ''' <remarks></remarks>
        alsLicenseFilePlain = 0
        ''' <summary>
        ''' Leaves the license file readable
        ''' </summary>
        ''' <remarks></remarks>
        alsLicenseFileEncrypted = 1
    End Enum

    ''' <summary>
    ''' Enum for Auto Registeration via ALL files
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ALAutoRegisterTypes
        ''' <summary>
        ''' Enables auto license registration
        ''' </summary>
        ''' <remarks></remarks>
        alsEnableAutoRegistration = 0
        ''' <summary>
        ''' Disables auto license registration
        ''' </summary>
        ''' <remarks></remarks>
        alsDisableAutoRegistration = 1
    End Enum

    ''' <summary>
    ''' Trial Warning can be persistent or temporary
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum ALTrialWarningTypes
        ''' <summary>
        ''' Trial Warning is Temporary (1-time only)
        ''' </summary>
        ''' <remarks></remarks>
        trialWarningTemporary = 0
        ''' <summary>
        ''' Trial Warning is Persistent
        ''' </summary>
        ''' <remarks></remarks>
        trialWarningPersistent = 1
    End Enum

#End Region

#Region "Public Properties"

#Region "Read Only"

    ''' <summary>
    ''' RemainingTrialDays - Read Only - Returns the Number of Used Trial Days.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Integer - Number of Used Trial Days</returns>
    ''' <remarks>None</remarks>
    Public ReadOnly Property RemainingTrialDays() As Integer Implements _IActiveLock.RemainingTrialDays
        Get
            RemainingTrialDays = 0
        End Get
    End Property

    ''' <summary>
    ''' RemainingTrialRuns - Read Only - Returns the Number of Used Trial Runs.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Integer - Number of Used Trial Runs</returns>
    ''' <remarks>None</remarks>
    Public ReadOnly Property RemainingTrialRuns() As Integer Implements _IActiveLock.RemainingTrialRuns
        Get
            RemainingTrialRuns = 0
        End Get
    End Property

    ''' <summary>
    ''' RegisteredLevel - Read Only - Returns the registered level.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String - Registered level</returns>
    ''' <remarks>None</remarks>
    Public ReadOnly Property RegisteredLevel() As String Implements _IActiveLock.RegisteredLevel
        Get
            RegisteredLevel = String.Empty
        End Get
    End Property

    ''' <summary>
    ''' MaxCount - Read Only - Returns the Number of concurrent users for the networked license
    ''' </summary>
    ''' <value></value>
    ''' <returns>Integer - Number of concurrent users for the networked license</returns>
    ''' <remarks>None</remarks>
    Public ReadOnly Property MaxCount() As Integer Implements _IActiveLock.MaxCount
        Get
            MaxCount = 0
        End Get
    End Property

    ''' <summary>
    ''' LicenseClass - Read Only - Returns the LicenseClass
    ''' </summary>
    ''' <value></value>
    ''' <returns>String - LicenseClass</returns>
    ''' <remarks>None</remarks>
    Public ReadOnly Property LicenseClass() As String Implements _IActiveLock.LicenseClass
        Get
            LicenseClass = "Single"
        End Get
    End Property

    ''' <summary>
    ''' UsedLockType - Read Only - Returns the Current Lock Type being used in this instance.
    ''' </summary>
    ''' <value></value>
    ''' <returns>ALLockTypes - lock type object corresponding to the current lock type(s) being used</returns>
    ''' <remarks>None</remarks>
    Public ReadOnly Property UsedLockType() As Integer Implements _IActiveLock.UsedLockType
        Get

        End Get
    End Property

    ''' <summary>
    ''' InstallationCode - Read Only - Returns the installation-specific code needed to obtain the liberation key.
    ''' </summary>
    ''' <param name="User">Optional - String - User</param>
    ''' <param name="Lic">Optional - ProductLicense - License</param>
    ''' <value>ByVal User As String - Optionally tailors the installation code specific to this user.</value>
    ''' <returns>String - Installation Code</returns>
    ''' <remarks></remarks>
    Public ReadOnly Property InstallationCode(Optional ByVal User As String = vbNullString, Optional ByVal Lic As ProductLicense = Nothing) As String Implements _IActiveLock.InstallationCode
        Get
            InstallationCode = String.Empty
        End Get
    End Property

    ''' <summary>
    ''' EventNotifier - Read Only - Retrieves the event notifier.
    ''' </summary>
    ''' <value></value>
    ''' <returns>ActiveLockEventNotifier - An object that can be used as a COM event source. i.e. can be used in <code>WithEvents</code> statements in VB.</returns>
    ''' <remarks>
    ''' Client applications uses this Notifier to handle event notifications sent by ActiveLock,
    ''' including license property validation and encryption events.
    ''' </remarks>
    Public ReadOnly Property EventNotifier() As ActiveLockEventNotifier Implements _IActiveLock.EventNotifier
        Get
            EventNotifier = Nothing
        End Get
    End Property

    ''' <summary>
    ''' UsedDays - Read Only - Returns the number of days this product has been used since its registration.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Long - Used days for the license</returns>
    ''' <remarks>None</remarks>
    Public ReadOnly Property UsedDays() As Integer Implements _IActiveLock.UsedDays
        Get
            UsedDays = 0
        End Get
    End Property

    ''' <summary>
    ''' RegisteredDate - Read Only - Retrieves the registration date.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String - Date on which the product is registered.</returns>
    ''' <remarks>None</remarks>
    Public ReadOnly Property RegisteredDate() As String Implements _IActiveLock.RegisteredDate
        Get
            RegisteredDate = String.Empty
        End Get
    End Property

    ''' <summary>
    ''' RegisteredUser - Read Only - Returns the registered user.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String - Registered user name</returns>
    ''' <remarks>None</remarks>
    Public ReadOnly Property RegisteredUser() As String Implements _IActiveLock.RegisteredUser
        Get
            RegisteredUser = String.Empty
        End Get
    End Property

    ''' <summary>
    ''' ExpirationDate - Read Only - Retrieves the expiration date.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String - Date on which the license will expire.</returns>
    ''' <remarks>None</remarks>
    Public ReadOnly Property ExpirationDate() As String Implements _IActiveLock.ExpirationDate
        Get
            ExpirationDate = String.Empty
        End Get
    End Property

#End Region

#Region "Write Only"

    ''' <summary>
    ''' LicenseKeyType - Write Only - Interface Property. Specifies the license key type for this instance of ActiveLock.
    ''' </summary>
    ''' <value>ByVal LicenseKeyTypes As ALLicenseKeyType - License Key Types object</value>
    ''' <remarks>None</remarks>
    Public WriteOnly Property LicenseKeyType() As ALLicenseKeyTypes Implements _IActiveLock.LicenseKeyType
        Set(ByVal Value As ALLicenseKeyTypes)

        End Set
    End Property

    ''' <summary>
    ''' CheckTimeServerForClockTampering - Write Only - Specifies whether a Time Server should be used to check Clock Tampering
    ''' </summary>
    ''' <value>ByVal Value As ALTimeServerTypes - Flag to use a Time Server or not</value>
    ''' <remarks></remarks>
    Public WriteOnly Property CheckTimeServerForClockTampering() As ALTimeServerTypes Implements _IActiveLock.CheckTimeServerForClockTampering
        Set(ByVal Value As ALTimeServerTypes)

        End Set
    End Property

    ''' <summary>
    ''' CheckSystemFilesForClockTampering - Write Only - Specifies whether the system files should be checked for Clock Tampering
    ''' </summary>
    ''' <value>ByVal Value As ALSystemFilesTypes - Flag to check system files or not</value>
    ''' <remarks></remarks>
    Public WriteOnly Property CheckSystemFilesForClockTampering() As ALSystemFilesTypes Implements _IActiveLock.CheckSystemFilesForClockTampering
        Set(ByVal Value As ALSystemFilesTypes)

        End Set
    End Property

    ''' <summary>
    ''' AutoRegister - Write Only - Specifies whether the auto register mechanism via an ALL file should be enabled or disabled
    ''' </summary>
    ''' <value>ByVal Value As ALAutoRegisterTypes - Flag to auto register a license or not</value>
    ''' <remarks></remarks>
    Public WriteOnly Property AutoRegister() As ALAutoRegisterTypes Implements _IActiveLock.AutoRegister
        Set(ByVal Value As ALAutoRegisterTypes)

        End Set
    End Property

    ''' <summary>
    ''' TrialWarning - Write Only - Specifies whether the Trial Warning is either Persistent or Temporary
    ''' </summary>
    ''' <value>ByVal Value As ALTrialWarningTypes - Trial Warning is either Persistent or Temporary</value>
    ''' <remarks></remarks>
    Public WriteOnly Property TrialWarning() As ALTrialWarningTypes Implements _IActiveLock.TrialWarning
        Set(ByVal Value As ALTrialWarningTypes)

        End Set
    End Property

    ''' <summary>
    ''' SoftwareCode - Write Only - Specifies the software code (product code)
    ''' </summary>
    ''' <value>ByVal sCode As String - Software Code</value>
    ''' <remarks></remarks>
    Public WriteOnly Property SoftwareCode() As String Implements _IActiveLock.SoftwareCode
        Set(ByVal Value As String)

        End Set
    End Property

    ''' <summary>
    ''' KeyStoreType - Write Only - Specifies the key store type.
    ''' </summary>
    ''' <value>ByVal KeyStore As LicStoreType - Key store type</value>
    ''' <remarks></remarks>
    Public WriteOnly Property KeyStoreType() As LicStoreType Implements _IActiveLock.KeyStoreType
        Set(ByVal Value As LicStoreType)

        End Set
    End Property

    ''' <summary>
    ''' KeyStorePath - Write Only - Specifies the key store path.
    ''' </summary>
    ''' <value>ByVal sPath As String - The path to be used for the specified KeyStoreType.</value>
    ''' <remarks>
    ''' <para>@param sPath - The path to be used for the specified KeyStoreType.</para>
    ''' <para>e.g. If <a href="IActiveLock.LicStoreType.html">alsFile</a> is used for <a href="IActiveLock.Let.KeyStoreType.html">KeyStoreType</a>,</para>
    ''' <para>then <code>Path</code> specifies the path to the license file.</para>
    ''' <para>If <a href="IActiveLock.LicStoreType.html">alsRegistry</a> is used,</para>
    ''' <para>the Path specifies the Registry hive where license information is stored.</para>
    ''' </remarks>
    Public WriteOnly Property KeyStorePath() As String Implements _IActiveLock.KeyStorePath
        Set(ByVal Value As String)

        End Set
    End Property

    ''' <summary>
    ''' AutoRegisterKeyPath - Write Only - Specifies the file path that contains the liberation key.
    ''' </summary>
    ''' <value>ByVal sPath As String - Full path to where the liberation file may reside.</value>
    ''' <remarks>
    ''' <para>If this file exists, ActiveLock will attempt to register the key automatically during its initialization.</para>
    ''' <para>Upon successful registration, the liberation file WILL be deleted.</para>
    ''' <para><b>Note</b>: This property is only effective if it is set prior to calling <code>Init</code>.</para>
    ''' </remarks>
    Public WriteOnly Property AutoRegisterKeyPath() As String Implements _IActiveLock.AutoRegisterKeyPath
        Set(ByVal Value As String)

        End Set
    End Property

#End Region

#Region "Read Write"

    ''' <summary>
    ''' LockType - Read/Write - Returns the Lock Type being used in this instance.
    ''' </summary>
    ''' <value>See TODO:</value>
    ''' <returns>ALLockTypes - lock type object corresponding to the lock type(s) being used</returns>
    ''' <remarks></remarks>
    Public Property LockType() As ALLockTypes Implements _IActiveLock.LockType
        Get

        End Get
        Set(ByVal Value As ALLockTypes)

        End Set
    End Property

    ''' <summary>
    ''' TrialHideType - Read/Write - Returns the Trial Hide Type being used in this instance.
    ''' </summary>
    ''' <value></value>
    ''' <returns>ALTrialHideTypes - trial hide type object corresponding to the trial hide type(s) being used</returns>
    ''' <remarks></remarks>
    Public Property TrialHideType() As ALTrialHideTypes Implements _IActiveLock.TrialHideType
        Get

        End Get
        Set(ByVal Value As ALTrialHideTypes)

        End Set
    End Property

    ''' <summary>
    ''' TrialType - Read/Write - Returns the Trial Type being used in this instance.
    ''' </summary>
    ''' <value></value>
    ''' <returns>ALTrialTypes - Trial Type (TrialNone, TrialByDays, TrialByRuns)</returns>
    ''' <remarks></remarks>
    Public Property TrialType() As ALTrialTypes Implements _IActiveLock.TrialType
        Get

        End Get
        Set(ByVal Value As ALTrialTypes)

        End Set
    End Property

    ''' <summary>
    ''' TrialLength - Read/Write - Returns the Trial Length being used in this instance.
    ''' </summary>
    ''' <value></value>
    ''' <returns>Integer - Trial Length (Number of Days or Runs)</returns>
    ''' <remarks></remarks>
    Public Property TrialLength() As Integer Implements _IActiveLock.TrialLength
        Get

        End Get
        Set(ByVal Value As Integer)

        End Set
    End Property

    ''' <summary>
    ''' SoftwareName - Read/Write - Returns the Software Name being used in this instance.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String - Software Name</returns>
    ''' <remarks></remarks>
    Public Property SoftwareName() As String Implements _IActiveLock.SoftwareName
        Get
            SoftwareName = String.Empty
        End Get
        Set(ByVal Value As String)

        End Set
    End Property

    ''' <summary>
    ''' SoftwarePassword - Read/Write - Returns the Software Password being used in this instance.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String - Software Password</returns>
    ''' <remarks></remarks>
    Public Property SoftwarePassword() As String Implements _IActiveLock.SoftwarePassword
        Get
            SoftwarePassword = String.Empty
        End Get
        Set(ByVal Value As String)

        End Set
    End Property

    ''' <summary>
    ''' LicenseFileType - Read/Write - Specifies whether the system files should be checked for Clock Tampering
    ''' </summary>
    ''' <value>ByVal Value As ALLicenseFileTypes - Encrypt License File or Leave it Plain</value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property LicenseFileType() As ALLicenseFileTypes Implements _IActiveLock.LicenseFileType
        Set(ByVal Value As ALLicenseFileTypes)

        End Set
        Get
            'LicenseFileType = String.Empty
        End Get
    End Property

    ''' <summary>
    ''' SoftwareVersion - Read/Write - Returns the Software Version being used in this instance.
    ''' </summary>
    ''' <value></value>
    ''' <returns>String - Software Version</returns>
    ''' <remarks></remarks>
    Public Property SoftwareVersion() As String Implements _IActiveLock.SoftwareVersion
        Get
            SoftwareVersion = String.Empty
        End Get
        Set(ByVal Value As String)

        End Set
    End Property

#End Region

#End Region

#Region "Functions"

    ''' <summary>
    ''' <para>LockCode - Interface Method. Computes a lock code corresponding to the specified Lock Types, License Class, etc.</para>
    ''' <para>Optionally, if a product license is specified, then a lock string specific to that license is returned.</para>
    ''' </summary>
    ''' <param name="Lic">Optional - ByRef Lic As ProductLicense - Product License for which to compute the lock code.</param>
    ''' <returns>String - Lock code</returns>
    ''' <remarks></remarks>
    Public Function LockCode(Optional ByRef Lic As ProductLicense = Nothing) As String Implements _IActiveLock.LockCode
        LockCode = String.Empty
    End Function

    ' TODO: IActiveLock.vb - Function Transfer - Not Implimented?
    ''' <summary>
    ''' Transfer - Not Implimented? - Transfers the current license to another computer.
    ''' </summary>
    ''' <param name="InstallCode">ByVal InstallCode As String - Installation Code generated from the other computer.</param>
    ''' <returns>String - The liberation key tailored for the request code generated from the other machine.</returns>
    ''' <remarks></remarks>
    Public Function Transfer(ByVal InstallCode As String) As String Implements _IActiveLock.Transfer
        Transfer = String.Empty
    End Function

    ' TODO: IActiveLock.vb - Function GenerateShortSerial - Update Comment
    ''' <summary>
    ''' GenerateShortSerial ? Undocumented...
    ''' </summary>
    ''' <param name="HDDfirmwareSerial"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GenerateShortSerial(ByVal HDDfirmwareSerial As String) As String Implements _IActiveLock.GenerateShortSerial
        Return Nothing
    End Function

    ' TODO: IActiveLock.vb - Function GenerateShortKey - Update Comment
    ''' <summary>
    ''' GenerateShortKey ? Undocumented...
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
    Public Function GenerateShortKey(ByVal SoftwareCode As String, ByVal SerialNumber As String, ByVal LicenseeAndRegisteredLevel As String, ByVal Expiration As String, ByVal LicType As ProductLicense.ALLicType, ByVal RegisteredLevel As Integer, Optional ByVal MaxUsers As Short = 1) As String Implements _IActiveLock.GenerateShortKey

        Return Nothing
    End Function

#End Region

#Region "Methods"

    ''' <summary>
    ''' Register - Registers the product using the specified liberation key.
    ''' </summary>
    ''' <param name="LibKey">ByVal LibKey As String - Liberation key</param>
    ''' <param name="user">Optional - String - User</param>
    ''' <remarks></remarks>
    Public Sub Register(ByVal LibKey As String, Optional ByRef user As String = "") Implements _IActiveLock.Register

    End Sub

    ''' <summary>
    ''' Init - Purpose: Initializes ActiveLock before use. Some of the routines, including <a href="IActiveLock.Acquire.html">Acquire()</a> and <a href="IActiveLock.Register.html">Register()</a> requires <code>Init()</code> to be called first.
    ''' </summary>
    ''' <param name="strPath">Optional - ?Undocumented!</param>
    ''' <param name="autoLicString">Optional - autoLicString As String - license key if autoregister is successful</param>
    ''' <remarks></remarks>
    Public Sub Init(Optional ByVal strPath As String = "", Optional ByRef autoLicString As String = "") Implements _IActiveLock.Init

    End Sub

    ''' <summary>
    ''' <para>Acquires a valid license token.</para>
    ''' <para>If no valid license can be found, an appropriate error will be raised, specifying the cause.</para>
    ''' </summary>
    ''' <param name="strMsg">Optional - ByRef strMsg As String - String returned by Activelock</param>
    ''' <param name="strRemainingTrialDays">Optional - ?Undocumented!</param>
    ''' <param name="strRemainingTrialRuns">Optional - ?Undocumented!</param>
    ''' <param name="strTrialLength">Optional - ?Undocumented!</param>
    ''' <param name="strUsedDays">Optional - ?Undocumented!</param>
    ''' <param name="strExpirationDate">Optional - ?Undocumented!</param>
    ''' <param name="strRegisteredUser">Optional - ?Undocumented!</param>
    ''' <param name="strRegisteredLevel">Optional - ?Undocumented!</param>
    ''' <param name="strLicenseClass">Optional - ?Undocumented!</param>
    ''' <param name="strMaxCount">Optional - ?Undocumented!</param>
    ''' <param name="strLicenseFileType">Optional - ?Undocumented!</param>
    ''' <param name="strLicenseType">Optional - ?Undocumented!</param>
    ''' <param name="strUsedLockType">Optional - ?Undocumented!</param>
    ''' <remarks></remarks>
    Public Sub Acquire(Optional ByRef strMsg As String = "", Optional ByRef strRemainingTrialDays As String = "", Optional ByRef strRemainingTrialRuns As String = "", Optional ByRef strTrialLength As String = "", Optional ByRef strUsedDays As String = "", Optional ByRef strExpirationDate As String = "", Optional ByRef strRegisteredUser As String = "", Optional ByRef strRegisteredLevel As String = "", Optional ByRef strLicenseClass As String = "", Optional ByRef strMaxCount As String = "", Optional ByRef strLicenseFileType As String = "", Optional ByRef strLicenseType As String = "", Optional ByRef strUsedLockType As String = "") Implements _IActiveLock.Acquire

    End Sub

    ''' <summary>
    ''' ResetTrial - Resets the Trial Mode
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ResetTrial() Implements _IActiveLock.ResetTrial

    End Sub

    ''' <summary>
    ''' KillTrial - Kills the Trial Mode
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub KillTrial() Implements _IActiveLock.KillTrial

    End Sub

#End Region



End Class