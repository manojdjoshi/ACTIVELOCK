VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IActiveLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Modified: 06.19.2006
'===============================================================================

Option Explicit
'###############################################################

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
'
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

'###############################################################
    
' License Key Type specifies the length/type
'
' @param alsALCryptoRSA    ' 1024-bit Long keys by RSA via ALCrypto DLL
' @param alsShortKeyMD5    ' Short license keys by MD5
Public Enum ALLicenseKeyTypes
    alsALCryptoRSA = 0
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


'===============================================================================
' Name: Property Get RegisteredLevel
' Input: None
' Output:
'   String - Registered level
' Purpose: Returns the registered level.
' Remarks: None
'===============================================================================
Public Property Get RegisteredLevel() As String

End Property
'===============================================================================
' Name: Property Get MaxCount
' Input: None
' Output:
'   Long - MaxCount
' Purpose: Returns the Number of concurrent users for the license
' Remarks: None
'===============================================================================
Public Property Get MaxCount() As Long

End Property

'===============================================================================
' Name: Property Get LicenseClass
' Input: None
' Output:
'   String - LicenseClass
' Purpose: Returns the LicenseClass for the license
' Remarks: None
'===============================================================================
Public Property Get LicenseClass() As String

End Property

'===============================================================================
' Name: Property Let LockType
' Input:
'    ByVal LockTypes As ALLockTypes - Lock Types object
' Output: None
' Purpose: Interface Property. Specifies the lock type for this instance of ActiveLock.
' Remarks: None
'===============================================================================
Public Property Let LockType(ByVal LockTypes As ALLockTypes)
Attribute LockType.VB_Description = "Lock type for this instance of ActiveLock."

End Property
'===============================================================================
' Name: Property Let LicenseKeyType
' Input:
'    ByVal LicenseKeyTypes As ALLicenseKeyType - License Key Types object
' Output: None
' Purpose: Interface Property. Specifies the license key type for this instance of ActiveLock.
' Remarks: None
'===============================================================================
Public Property Let LicenseKeyType(ByVal LicenseKeyTypes As ALLicenseKeyTypes)

End Property

'===============================================================================
' Name: Property Let TrialHideType
' Input:
'    ByVal TrialHideTypes As ALTrialHideTypes - Trial Hide Types object
' Output: None
' Purpose: Interface Property. Specifies the trial hide type for this instance of ActiveLock.
' Remarks: None
'===============================================================================
Public Property Let TrialHideType(ByVal TrialHideTypes As ALTrialHideTypes)

End Property

'===============================================================================
' Name: Property Let TrialType
' Input:
'    ByVal TrialType As ALTrialTypes - Trial Type (TrialNone, TrialByDays, TrialByRuns)
' Output: None
' Purpose: Interface Property. Specifies the trial type for this instance of ActiveLock.
' Remarks: None
'===============================================================================
Public Property Let TrialType(ByVal TrialType As ALTrialTypes)

End Property

'===============================================================================
' Name: Property Let TrialLength
' Input:
'    ByVal TrialLength As Long - Trial Length (Number of Days or Runs)
' Output: None
' Purpose: Interface Property. Specifies the trial length for this instance of ActiveLock.
' Remarks: None
'===============================================================================
Public Property Let TrialLength(ByVal TrialLength As Long)

End Property

'===============================================================================
' Name: Property Get LockType
' Input: None
' Output:
'   ALLockTypes - lock type object corresponding to the lock type(s) being used
' Purpose: Returns the Lock Type being used in this instance.
' Remarks: None
'===============================================================================
Public Property Get LockType() As ALLockTypes

End Property

'===============================================================================
' Name: Property Get UsedLockType
' Input: None
' Output:
'   ALLockTypes - Used lock type object corresponding to the lock type(s) currently used
' Purpose: Returns the Used Lock Type in this instance.
' Remarks: None
'===============================================================================
Public Property Get UsedLockType() As ALLockTypes()

End Property

'===============================================================================
' Name: Property Get TrialHideType
' Input: None
' Output:
'   ALTrialHideTypes - trial hide type object corresponding to the trial hide type(s) being used
' Purpose: Returns the Trial Hide Type being used in this instance.
' Remarks: None
'===============================================================================
Public Property Get TrialHideType() As ALTrialHideTypes

End Property

'===============================================================================
' Name: Property Get TrialType
' Input: None
' Output:
'   ALTrialTypes - Trial Type (TrialNone, TrialByDays, TrialByRuns)
' Purpose: Returns the Trial Type being used in this instance.
' Remarks: None
'===============================================================================
Public Property Get TrialType() As ALTrialTypes

End Property

'===============================================================================
' Name: Property Get TrialLength
' Input: None
' Output:
'   Integer - Trial Length (Number of Days or Runs)
' Purpose: Returns the Trial Length being used in this instance.
' Remarks: None
'===============================================================================
Public Property Get TrialLength() As Long

End Property
'===============================================================================
' Name: Property Let SoftwareName
' Input:
'   ByVal sName As String - Software Name
' Output: None
' Purpose: Specifies the name of the product being locked.
' Remarks: None
'===============================================================================
Public Property Let SoftwareName(ByVal sName As String)
Attribute SoftwareName.VB_Description = "Name of the product being locked."

End Property

'===============================================================================
' Name: Property Let SoftwarePassword
' Input:
'   ByVal sName As String - Software Password
' Output: None
' Purpose: Specifies the Password of the product being locked.
' Remarks: None
'===============================================================================
Public Property Let SoftwarePassword(ByVal sPassword As String)

End Property

'===============================================================================
' Name: Property Get SoftwareName
' Input: None
' Output:
'   String - Software Name
' Purpose: Returns the Software Name being used in this instance.
' Remarks: None
'===============================================================================
Public Property Get SoftwareName() As String

End Property
'===============================================================================
' Name: Property Get SoftwarePassword
' Input: None
' Output:
'   String - Software Password
' Purpose: Returns the Software Password being used in this instance.
' Remarks: None
'===============================================================================
Public Property Get SoftwarePassword() As String

End Property


'===============================================================================
' Name: Property Let SoftwareCode
' Input:
'   ByVal sCode As String - Software Code
' Output: None
' Purpose: Specifies the software code (product code)
' Remarks: None
'===============================================================================
Public Property Let SoftwareCode(ByVal sCode As String)
Attribute SoftwareCode.VB_Description = "software code (product code).  This is Code1 generated by ALUGEN."

End Property
'===============================================================================
' Name: Property Let SoftwareVersion
' Input:
'   ByVal sVer As String - Version string such as 1.0
' Output: None
' Purpose: Specifies the version of the product being locked.
' Remarks: None
'===============================================================================
Public Property Let SoftwareVersion(ByVal sVer As String)
Attribute SoftwareVersion.VB_Description = "Version of the product being locked."

End Property
'===============================================================================
' Name: Property Get SoftwareVersion
' Input: None
' Output:
'   String - Software Version
' Purpose: Returns the Software Version being used in this instance.
' Remarks: None
'===============================================================================
Public Property Get SoftwareVersion() As String

End Property


'===============================================================================
' Name: Property Let KeyStoreType
' Input:
'   ByVal KeyStore As LicStoreType - Key store type
' Output: None
' Purpose: Specifies the key store type.
' Remarks: None
'===============================================================================
Public Property Let KeyStoreType(ByVal KeyStore As LicStoreType)
Attribute KeyStoreType.VB_Description = "Specifies the key store (e.g. registry or file)."

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
Public Property Let KeyStorePath(ByVal sPath As String)
Attribute KeyStorePath.VB_Description = "Specifies the path where the license store resides."

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
Public Property Get InstallationCode(Optional ByVal user As String = "", Optional Lic As ProductLicense = Nothing) As String

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
Public Property Let AutoRegisterKeyPath(ByVal sPath As String)
Attribute AutoRegisterKeyPath.VB_Description = "Specifies the file path that contains the liberation key for automatic registration upon startup. Upon successful registration, the liberation file WILL be deleted."

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
Public Function LockCode(Optional Lic As ProductLicense = Nothing) As String
Attribute LockCode.VB_Description = "Computes a lock code corresponding to the specified Lock Types, License Class, etc... Optionally, if a product license is specified, then a lock string specific to that license is returned."

End Function

'===============================================================================
' Name: Sub Register
' Input:
'   ByVal LibKey As String - Liberation key
' Output: None
' Purpose: Registers the product using the specified liberation key.
' Remarks: None
'===============================================================================
Public Sub Register(ByVal LibKey As String, Optional user As String = "")
Attribute Register.VB_Description = "Registers the product using the specified liberation key."

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
Public Function Transfer(ByVal InstallCode As String) As String
Attribute Transfer.VB_Description = "Transfers the current license to another computer."

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
Public Sub Init(Optional autoLicString As String)
Attribute Init.VB_Description = "Initializes ActiveLock. This routine must be called before any other routine such as Acquire() or Register()."

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
Public Sub Acquire(Optional strMsg As String)
Attribute Acquire.VB_Description = "Attempts to acquire a valid license token."

End Sub
'===============================================================================
' Name: Sub ResetTrial
' Input: None
' Output: None
' Purpose: Resets the Trial Mode
' Remarks: None
'===============================================================================
Public Sub ResetTrial()

End Sub
'===============================================================================
' Name: Sub KillTrial
' Input: None
' Output: None
' Purpose: Kills the Trial Mode
' Remarks: None
'===============================================================================
Public Sub KillTrial()

End Sub

Public Function GenerateShortSerial(ByVal HDDfirmwareSerial As String) As String

End Function

Public Function GenerateShortKey(ByVal SoftwareCode As String, ByVal SerialNumber As String, ByVal LicenseeAndRegisteredLevel As String, ByVal Expiration As String, ByVal LicType As ActiveLock3.ALLicType, ByVal RegisteredLevel As Long, Optional ByVal MaxUsers As Integer = 1) As String

End Function

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
Public Property Get EventNotifier() As ActiveLockEventNotifier
Attribute EventNotifier.VB_Description = "Returns the Event Notifier object used by ActiveLock to broadcast COM events."

End Property

'===============================================================================
' Name: Property Get UsedDays
' Input: None
' Output:
'   Long - Used days for the license
' Purpose: Returns the number of days this product has been used since its registration.
' Remarks: None
'===============================================================================
Public Property Get UsedDays() As Long
Attribute UsedDays.VB_Description = "Returns the number of days this product has been used since its registration."

End Property
'===============================================================================
' Name: Property Get RegisteredDate
' Input: None
' Output:
'   String - Date on which the product is registered.
' Purpose: Retrieves the registration date.
' Remarks: None
'===============================================================================
Public Property Get RegisteredDate() As String
Attribute RegisteredDate.VB_Description = "Returns the date on which the product was registered."

End Property


'===============================================================================
' Name: Property Get RegisteredUser
' Input: None
' Output:
'   String - Registered user name
' Purpose: Returns the registered user.
' Remarks: None
'===============================================================================
Public Property Get RegisteredUser() As String
Attribute RegisteredUser.VB_Description = "Returns the registered user."

End Property
'===============================================================================
' Name: Property Get ExpirationDate
' Input: None
' Output:
'   String - Date on which the license will expire.
' Purpose: Retrieves the expiration date.
' Remarks: None
'===============================================================================
Public Property Get ExpirationDate() As String
Attribute ExpirationDate.VB_Description = "Returns the expiration date of the current license."

End Property

