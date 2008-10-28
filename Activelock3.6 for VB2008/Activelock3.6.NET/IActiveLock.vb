Option Strict Off
Option Explicit On
Public Interface _IActiveLock
    ReadOnly Property RemainingTrialDays() As Integer
    ReadOnly Property RemainingTrialRuns() As Integer
    ReadOnly Property MaxCount() As Integer
    ReadOnly Property RegisteredLevel() As String
    ReadOnly Property LicenseClass() As String
    Property LockType() As IActiveLock.ALLockTypes
    WriteOnly Property LicenseKeyType() As IActiveLock.ALLicenseKeyTypes
    Property UsedLockType() As IActiveLock.ALLockTypes()
    Property TrialHideType() As IActiveLock.ALTrialHideTypes
    Property TrialType() As IActiveLock.ALTrialTypes
    Property TrialLength() As Integer
    Property SoftwareName() As String
    Property SoftwarePassword() As String
    WriteOnly Property CheckTimeServerForClockTampering() As IActiveLock.ALTimeServerTypes
    WriteOnly Property CheckSystemFilesForClockTampering() As IActiveLock.ALSystemFilesTypes
    WriteOnly Property LicenseFileType() As IActiveLock.ALLicenseFileTypes
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
    Sub Acquire(Optional ByRef strMsg As String = "")
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
'Class instancing was changed to public.
<System.Runtime.InteropServices.ProgId("IActiveLock_NET.IActiveLock")> Public Class IActiveLock
	Implements _IActiveLock
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
	'
    ' License Lock Types.  Values can be combined (OR&#39;ed) together.
	'
	' @param lockNone       No locking - not recommended
	' @param lockWindows    Lock to Windows Serial Number
	' @param lockComp       Lock to Computer Name
	' @param lockHD         Lock to Hard Drive Serial Number (Volume Serial Number)
	' @param lockMAC        Lock to Network Interface Card Address
	' @param lockBIOS       Lock to BIOS Serial Number
	' @param lockIP         Lock to Computer IP Address
	' @param lockMotherboard   Lock to Motherboard Serial Number
	' @param lockCustom     Lock to Custom String (Not Implemented)
	' @param lockHDFirmware Lock to Hard Disk Firmware Serial (HDD Manufacturer&#39;s Serial Number)
	' @param lockSID        Lock to Computer Security ID (SID)(Not Implemented)
	
	'###############################################################
    'ADDED lockHDFirmware = 256
	'Date: 4/21/05
	'By: Scott Nelson (Sentax)
	Public Enum ALLockTypes
		lockNone = 0
		lockWindows = 1
		lockComp = 2
		lockHD = 4
		lockMAC = 8
		lockBIOS = 16
		lockIP = 32
		lockMotherboard = 64
		lockHDFirmware = 256
	End Enum
	'    lockCustom = 128
	'    lockSID = 512
	'
    '###############################################################

    ' License Key Type specifies the length/type
    '
    ' @param alsRSA    ' 1024-bit Long keys by RSA via ALCrypto DLL
    ' @param alsShortKeyMD5    ' Short license keys by MD5
    Public Enum ALLicenseKeyTypes
        alsRSA = 0
        alsShortKeyMD5 = 1
    End Enum
    '###############################################################
    ' License Store Type specifies where to store license keys
    '
    ' @param alsRegistry    ' Store in Windows Registry
    ' @param alsFile        ' Store in a license file
    Public Enum LicStoreType
        alsRegistry = 0
        alsFile = 1
    End Enum

    '###############################################################

    ' Products Store Type specifies where to store products infos
    '
    ' @param alsINIFile    ' Store in INI file (licenses.ini)
    ' @param alsXMLFile    ' Store in XML file (licenses.xml)
    ' @param alsMDBFile    ' Store in MDB file (licenses.mdb)
    '                       mdb file should contain a table named products with structure:
    '                       ID(autonumber), name(text,150), version (text,50), vccode(memo), gcode(memo)
    ' @param alsMSSQL      ' Store in MSSQL database 'not implemented 'TODO
    Public Enum ProductsStoreType
        alsINIFile = 0
        alsXMLFile = 1
        alsMDBFile = 2
        'alsMSSQL = 3 'TODO
    End Enum

    '###############################################################

    ' Trial Type specifies what kind of Trial Feature is used
    '
    ' @param trialNone      ' No trial used
    ' @param trialDays      ' Trial by Days
    ' @param trialRuns      ' Trial by Runs
    Public Enum ALTrialTypes
        trialNone = 0
        trialDays = 1
        trialRuns = 2
    End Enum
    '###############################################################

    ' Trial Hide Mode Type specifies what kind of Trial Hiding Mode is used
    ' Values can be combined (OR&#39;ed) together.
    '
    ' @param trialSteganography     ' Trial information is hidden in BMP files
    ' @param trialHiddenFolder      ' Trial information is hidden in a folder which uses a default namespace
    ' @param trialRegistry          ' Trial information is encrypted and hidden in several registry locations
    Public Enum ALTrialHideTypes
        trialSteganography = 1
        trialHiddenFolder = 2
        trialRegistry = 4
    End Enum
    '###############################################################

    ' Enum for accessing the Time Server to check Clock Tampering
    '
    ' @param alsDontCheckTimeServer  ' Skips checking a Time Server
    ' @param alsCheckTimeServer      ' Checks a Time Server
    Public Enum ALTimeServerTypes
        alsDontCheckTimeServer = 0
        alsCheckTimeServer = 1
    End Enum
    '###############################################################

    ' Enum for scanning the system folders/files to detect clock tampering
    '
    ' @param alsDontCheckSystemFiles  ' Skips checking system files
    ' @param alsCheckSystemFiles      ' Checks system files
    Public Enum ALSystemFilesTypes
        alsDontCheckSystemFiles = 0
        alsCheckSystemFiles = 1
    End Enum

    '###############################################################

    ' Enum for license file encryption
    '
    ' @param alsLicenseFileEncrypt    ' Encrypts the license file
    ' @param alsLicenseFilePlain      ' Leaves the license file readable
    Public Enum ALLicenseFileTypes
        alsLicenseFilePlain = 0
        alsLicenseFileEncrypted = 1
    End Enum
    '###############################################################

    ' Enum for Auto Registeration via ALL files
    '
    ' @param alsEnableAutoRegistration      ' Enables auto license registration
    ' @param alsDisableAutoRegistration     ' Disables auto license registration
    Public Enum ALAutoRegisterTypes
        alsEnableAutoRegistration = 0
        alsDisableAutoRegistration = 1
    End Enum
    '###############################################################

    ' Trial Warning can be persistent or temporary
    '
    ' @param trialWarningTemporary       ' Trial Warning is Temporary (1-time only)
    ' @param trialWarningPersistent      ' Trial Warning is Persistent
    Public Enum ALTrialWarningTypes
        trialWarningTemporary = 0
        trialWarningPersistent = 1
    End Enum
    '===============================================================================
    ' Name: Property Get RemainingTrialDays
    ' Input: None
    ' Output:
    '   Integer - Number of Used Trial Days
    ' Purpose: Returns the Number of Used Trial Days.
    ' Remarks: None
    '===============================================================================
    Public ReadOnly Property RemainingTrialDays() As Integer Implements _IActiveLock.RemainingTrialDays
        Get
            RemainingTrialDays = 0
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get RemainingTrialRuns
    ' Input: None
    ' Output:
    '   Integer - Number of Used Trial Runs
    ' Purpose: Returns the Number of Used Trial Runs.
    ' Remarks: None
    '===============================================================================
    Public ReadOnly Property RemainingTrialRuns() As Integer Implements _IActiveLock.RemainingTrialRuns
        Get
            RemainingTrialRuns = 0
        End Get
    End Property

    '===============================================================================
    ' Name: Property Get RegisteredLevel
    ' Input: None
    ' Output:
    '   String - Registered level
    ' Purpose: Returns the registered level.
    ' Remarks: None
    '===============================================================================
    Public ReadOnly Property RegisteredLevel() As String Implements _IActiveLock.RegisteredLevel
        Get
            RegisteredLevel = String.Empty
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get MaxCount
    ' Input: None
    ' Output:
    '   Integer - Number of concurrent users for the networked license
    ' Purpose: Returns the Number of concurrent users for the networked license
    ' Remarks: None
    '===============================================================================
    Public ReadOnly Property MaxCount() As Integer Implements _IActiveLock.MaxCount
        Get
            MaxCount = 0
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get LicenseClass
    ' Input: None
    ' Output:
    '   String - LicenseClass
    ' Purpose: Returns the LicenseClass
    ' Remarks: None
    '===============================================================================
    Public ReadOnly Property LicenseClass() As String Implements _IActiveLock.LicenseClass
        Get
            LicenseClass = "Single"
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get LockType
    ' Input: None
    ' Output:
    '   ALLockTypes - lock type object corresponding to the lock type(s) being used
    ' Purpose: Returns the Lock Type being used in this instance.
    ' Remarks: None
    '===============================================================================
    Public Property LockType() As ALLockTypes Implements _IActiveLock.LockType
        Get

        End Get
        Set(ByVal Value As ALLockTypes)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Let LicenseKeyType
    ' Input:
    '    ByVal LicenseKeyTypes As ALLicenseKeyType - License Key Types object
    ' Output: None
    ' Purpose: Interface Property. Specifies the license key type for this instance of ActiveLock.
    ' Remarks: None
    '===============================================================================
    Public WriteOnly Property LicenseKeyType() As ALLicenseKeyTypes Implements _IActiveLock.LicenseKeyType
        Set(ByVal Value As ALLicenseKeyTypes)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Get UsedLockType
    ' Input: None
    ' Output:
    '   ALLockTypes - lock type object corresponding to the current lock type(s) being used
    ' Purpose: Returns the Current Lock Type being used in this instance.
    ' Remarks: None
    '===============================================================================
    Public Property UsedLockType() As ALLockTypes() Implements _IActiveLock.UsedLockType
        Get
            Return Nothing
        End Get
        Set(ByVal Value As ALLockTypes())

        End Set
    End Property

    '===============================================================================
    ' Name: Property Get TrialHideType
    ' Input: None
    ' Output:
    '   ALTrialHideTypes - trial hide type object corresponding to the trial hide type(s) being used
    ' Purpose: Returns the Trial Hide Type being used in this instance.
    ' Remarks: None
    '===============================================================================
    Public Property TrialHideType() As ALTrialHideTypes Implements _IActiveLock.TrialHideType
        Get

        End Get
        Set(ByVal Value As ALTrialHideTypes)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Get TrialType
    ' Input: None
    ' Output:
    '   ALTrialTypes - Trial Type (TrialNone, TrialByDays, TrialByRuns)
    ' Purpose: Returns the Trial Type being used in this instance.
    ' Remarks: None
    '===============================================================================
    Public Property TrialType() As ALTrialTypes Implements _IActiveLock.TrialType
        Get

        End Get
        Set(ByVal Value As ALTrialTypes)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Get TrialLength
    ' Input: None
    ' Output:
    '   Integer - Trial Length (Number of Days or Runs)
    ' Purpose: Returns the Trial Length being used in this instance.
    ' Remarks: None
    '===============================================================================
    Public Property TrialLength() As Integer Implements _IActiveLock.TrialLength
        Get

        End Get
        Set(ByVal Value As Integer)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Get SoftwareName
    ' Input: None
    ' Output:
    '   String - Software Name
    ' Purpose: Returns the Software Name being used in this instance.
    ' Remarks: None
    '===============================================================================
    Public Property SoftwareName() As String Implements _IActiveLock.SoftwareName
        Get
            SoftwareName = String.Empty
        End Get
        Set(ByVal Value As String)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Get SoftwarePassword
    ' Input: None
    ' Output:
    '   String - Software Password
    ' Purpose: Returns the Software Password being used in this instance.
    ' Remarks: None
    '===============================================================================
    Public Property SoftwarePassword() As String Implements _IActiveLock.SoftwarePassword
        Get
            SoftwarePassword = String.Empty
        End Get
        Set(ByVal Value As String)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Let CheckTimeServerForClockTampering
    ' Input:
    '   ByVal Value As ALTimeServerTypes - Flag to use a Time Server or not
    ' Output:
    '   None
    ' Purpose: Specifies whether a Time Server should be used to check Clock Tampering
    ' Remarks: None
    '===============================================================================
    Public WriteOnly Property CheckTimeServerForClockTampering() As ALTimeServerTypes Implements _IActiveLock.CheckTimeServerForClockTampering
        Set(ByVal Value As ALTimeServerTypes)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Let CheckSystemFilesForClockTampering
    ' Input:
    '   ByVal Value As ALSystemFilesTypes - Flag to check system files or not
    ' Output:
    '   None
    ' Purpose: Specifies whether the system files should be checked for Clock Tampering
    ' Remarks: None
    '===============================================================================
    Public WriteOnly Property CheckSystemFilesForClockTampering() As ALSystemFilesTypes Implements _IActiveLock.CheckSystemFilesForClockTampering
        Set(ByVal Value As ALSystemFilesTypes)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Let LicenseFileType
    ' Input:
    '   ByVal Value As ALLicenseFileTypes - Encrypt License File or Leave it Plain
    ' Output:
    '   None
    ' Purpose: Specifies whether the system files should be checked for Clock Tampering
    ' Remarks: None
    '===============================================================================
    Public WriteOnly Property LicenseFileType() As ALLicenseFileTypes Implements _IActiveLock.LicenseFileType
        Set(ByVal Value As ALLicenseFileTypes)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Let AutoRegister
    ' Input:
    '   ByVal Value As ALAutoRegisterTypes - Flag to auto register a license or not
    ' Output:
    '   None
    ' Purpose: Specifies whether the auto register mechanism via an ALL file should be enabled or disabled
    ' Remarks: None
    '===============================================================================
    Public WriteOnly Property AutoRegister() As ALAutoRegisterTypes Implements _IActiveLock.AutoRegister
        Set(ByVal Value As ALAutoRegisterTypes)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Let TrialWarning
    ' Input:
    '   ByVal Value As ALTrialWarningTypes - Trial Warning is either Persistent or Temporary
    ' Output: None
    ' Purpose: Specifies whether the Trial Warning is either Persistent or Temporary
    ' Remarks: None
    '===============================================================================
    Public WriteOnly Property TrialWarning() As ALTrialWarningTypes Implements _IActiveLock.TrialWarning
        Set(ByVal Value As ALTrialWarningTypes)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Let SoftwareCode
    ' Input:
    '   ByVal sCode As String - Software Code
    ' Output: None
    ' Purpose: Specifies the software code (product code)
    ' Remarks: None
    '===============================================================================
    Public WriteOnly Property SoftwareCode() As String Implements _IActiveLock.SoftwareCode
        Set(ByVal Value As String)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Get SoftwareVersion
    ' Input: None
    ' Output:
    '   String - Software Version
    ' Purpose: Returns the Software Version being used in this instance.
    ' Remarks: None
    '===============================================================================
    Public Property SoftwareVersion() As String Implements _IActiveLock.SoftwareVersion
        Get
            SoftwareVersion = String.Empty
        End Get
        Set(ByVal Value As String)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Let KeyStoreType
    ' Input:
    '   ByVal KeyStore As LicStoreType - Key store type
    ' Output: None
    ' Purpose: Specifies the key store type.
    ' Remarks: None
    '===============================================================================
    Public WriteOnly Property KeyStoreType() As LicStoreType Implements _IActiveLock.KeyStoreType
        Set(ByVal Value As LicStoreType)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Let KeyStorePath
    ' Input:
    '   ByVal sPath As String - File path and name
    ' Output: None
    ' Purpose: Specifies the key store path.
    ' Remarks: None
    '===============================================================================
    ' @param sPath  The path to be used for the specified KeyStoreType.
    '               e.g. If <a href="IActiveLock.LicStoreType.html">alsFile</a> is used for <a href="IActiveLock.Let.KeyStoreType.html">KeyStoreType</a>,
    '               then <code>Path</code> specifies the path to the license file.
    '               If <a href="IActiveLock.LicStoreType.html">alsRegistry</a> is used,
    '               the Path specifies the Registry hive where license information is stored.
    Public WriteOnly Property KeyStorePath() As String Implements _IActiveLock.KeyStorePath
        Set(ByVal Value As String)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Get InstallationCode
    ' Input:
    '   ByVal User As String - Optionally tailors the installation code specific to this user.
    ' Output:
    '   String - Installation Code
    ' Purpose: Returns the installation-specific code needed to obtain the liberation key.
    ' Remarks: None
    '===============================================================================
    Public ReadOnly Property InstallationCode(Optional ByVal User As String = vbNullString, Optional ByVal Lic As ProductLicense = Nothing) As String Implements _IActiveLock.InstallationCode
        Get
            InstallationCode = String.Empty
        End Get
    End Property
    '===============================================================================
    ' Name: Property Let AutoRegisterKeyPath
    ' Input:
    '   ByVal sPath As String - Full path to where the liberation file may reside.
    ' Output: None
    ' Purpose: Specifies the file path that contains the liberation key.
    ' <p>If this file exists, ActiveLock will attempt to register the key automatically during its initialization.
    ' <p>Upon successful registration, the liberation file WILL be deleted.
    ' Remarks: Note: This property is only effective if it is set prior to calling <code>Init</code>.
    '===============================================================================
    Public WriteOnly Property AutoRegisterKeyPath() As String Implements _IActiveLock.AutoRegisterKeyPath
        Set(ByVal Value As String)

        End Set
    End Property
    '===============================================================================
    ' Name: Property Get EventNotifier
    ' Input: None
    ' Output:
    '   ActiveLockEventNotifier - An object that can be used as a COM event source. i.e. can be used in <code>WithEvents</code> statements in VB.
    ' Purpose: Retrieves the event notifier.
    ' <p>Client applications uses this Notifier to handle event notifications sent by ActiveLock,
    ' including license property validation and encryption events.
    ' Remarks: See ActiveLockEventNotifier for more information.
    '===============================================================================
    Public ReadOnly Property EventNotifier() As ActiveLockEventNotifier Implements _IActiveLock.EventNotifier
        Get
            EventNotifier = Nothing
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get UsedDays
    ' Input: None
    ' Output:
    '   Long - Used days for the license
    ' Purpose: Returns the number of days this product has been used since its registration.
    ' Remarks: None
    '===============================================================================
    Public ReadOnly Property UsedDays() As Integer Implements _IActiveLock.UsedDays
        Get

        End Get
    End Property
    '===============================================================================
    ' Name: Property Get RegisteredDate
    ' Input: None
    ' Output:
    '   String - Date on which the product is registered.
    ' Purpose: Retrieves the registration date.
    ' Remarks: None
    '===============================================================================
    Public ReadOnly Property RegisteredDate() As String Implements _IActiveLock.RegisteredDate
        Get
            RegisteredDate = String.Empty
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get RegisteredUser
    ' Input: None
    ' Output:
    '   String - Registered user name
    ' Purpose: Returns the registered user.
    ' Remarks: None
    '===============================================================================
    Public ReadOnly Property RegisteredUser() As String Implements _IActiveLock.RegisteredUser
        Get
            RegisteredUser = String.Empty
        End Get
    End Property
    '===============================================================================
    ' Name: Property Get ExpirationDate
    ' Input: None
    ' Output:
    '   String - Date on which the license will expire.
    ' Purpose: Retrieves the expiration date.
    ' Remarks: None
    '===============================================================================
    Public ReadOnly Property ExpirationDate() As String Implements _IActiveLock.ExpirationDate
        Get
            ExpirationDate = String.Empty
        End Get
    End Property
    '===============================================================================
    ' Name: Function LockCode
    ' Input:
    '   ByRef Lic As ProductLicense - Product License for which to compute the lock code.
    ' Output:
    '   String - Lock code
    ' Purpose: Interface Method. Computes a lock code corresponding to the specified Lock Types, License Class, etc.
    ' Optionally, if a product license is specified, then a lock string specific to that license is returned.
    ' Remarks: None
    '===============================================================================
    Public Function LockCode(Optional ByRef Lic As ProductLicense = Nothing) As String Implements _IActiveLock.LockCode
        LockCode = String.Empty
    End Function
    '===============================================================================
    ' Name: Sub Register
    ' Input:
    '   ByVal LibKey As String - Liberation key
    ' Output: None
    ' Purpose: Registers the product using the specified liberation key.
    ' Remarks: None
    '===============================================================================
    Public Sub Register(ByVal LibKey As String, Optional ByRef user As String = "") Implements _IActiveLock.Register

    End Sub
    '===============================================================================
    ' Name: Function Transfer
    ' Input:
    '   ByVal InstallCode As String - Installation Code generated from the other computer.
    ' Output:
    '   String - The liberation key tailored for the request code generated from the other machine.
    ' Purpose: Transfers the current license to another computer.
    ' Remarks: Not implemented yet.
    '===============================================================================
    Public Function Transfer(ByVal InstallCode As String) As String Implements _IActiveLock.Transfer
        Transfer = String.Empty
    End Function
    '===============================================================================
    ' Name: Sub Init
    ' Input:
    '    autoLicString As String - license key if autoregister is successful
    ' Output: None
    ' Purpose: Initializes ActiveLock before use. Some of the routines, including <a href="IActiveLock.Acquire.html">Acquire()</a>
    ' and <a href="IActiveLock.Register.html">Register()</a> requires <code>Init()</code> to be called first.
    ' Remarks: None
    '===============================================================================
    Public Sub Init(Optional ByVal strPath As String = "", Optional ByRef autoLicString As String = "") Implements _IActiveLock.Init

    End Sub
    '===============================================================================
    ' Name: Sub Acquire
    ' Input:
    '   ByRef strMsg As String - String returned by Activelock
    ' Output: None
    ' Purpose: Acquires a valid license token.
    ' <p>If no valid license can be found, an appropriate error will be raised, specifying the cause.
    ' Remarks: None
    '===============================================================================
    Public Sub Acquire(Optional ByRef strMsg As String = "") Implements _IActiveLock.Acquire

    End Sub
    '===============================================================================
    ' Name: Sub ResetTrial
    ' Input: None
    ' Output: None
    ' Purpose: Resets the Trial Mode
    ' Remarks: None
    '===============================================================================
    Public Sub ResetTrial() Implements _IActiveLock.ResetTrial

    End Sub
    '===============================================================================
    ' Name: Sub KillTrial
    ' Input: None
    ' Output: None
    ' Purpose: Kills the Trial Mode
    ' Remarks: None
    '===============================================================================
    Public Sub KillTrial() Implements _IActiveLock.KillTrial

    End Sub
    Public Function GenerateShortSerial(ByVal HDDfirmwareSerial As String) As String Implements _IActiveLock.GenerateShortSerial
        Return Nothing
    End Function

    Public Function GenerateShortKey(ByVal SoftwareCode As String, ByVal SerialNumber As String, ByVal LicenseeAndRegisteredLevel As String, ByVal Expiration As String, ByVal LicType As ProductLicense.ALLicType, ByVal RegisteredLevel As Integer, Optional ByVal MaxUsers As Short = 1) As String Implements _IActiveLock.GenerateShortKey

        Return Nothing
    End Function
End Class