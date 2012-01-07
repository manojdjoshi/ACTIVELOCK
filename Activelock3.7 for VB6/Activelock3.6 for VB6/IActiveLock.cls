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
' This project is available from SVN on SourceForge.net under the main project, Activelock !
'
' ProjectPage: http://sourceforge.net/projects/activelock
' WebSite: http://www.activeLockSoftware.com
' DeveloperForums: http://forums.activelocksoftware.com
' ProjectManager: Ismail Alkan - http://activelocksoftware.com/simplemachinesforum/index.php?action=profile;u=1
' ProjectLicense: BSD Open License - http://www.opensource.org/licenses/bsd-license.php
' ProjectPurpose: Copy Protection, Software Locking, Anti Piracy
'
' //////////////////////////////////////////////////////////////////////////////////////////
' *   ActiveLock
' *   Copyright 1998-2002 Nelson Ferraz
' *   Copyright 2003-2009 The ActiveLock Software Group (ASG)
' *   All material is the property of the contributing authors.
' *
' *   Redistribution and use in source and binary forms, with or without
' *   modification, are permitted provided that the following conditions are
' *   met:
' *
' *     [o] Redistributions of source code must retain the above copyright
' *         notice, this list of conditions and the following disclaimer.
' *
' *     [o] Redistributions in binary form must reproduce the above
' *         copyright notice, this list of conditions and the following
' *         disclaimer in the documentation and/or other materials provided
' *         with the distribution.
' *
' *   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' *   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' *   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
' *   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
' *   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
' *   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
' *   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
' *   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' *   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' *   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' *   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' *
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
' @param lockExternalIP Lock to External IP Address
' @param lockHDFirmware Lock to Hard Disk Firmware Serial (HDD Manufacturer's Serial Number)
' @param lockFingerprint   Lock to Fingerprint (Activelock Combination)
' @param lockMemory     Lock to Memory ID
' @param lockCPUID      Lock to CPU ID
' @param lockBaseboardID   Lock to Baseboard Name and Serial Number
' @param lockVideoID    Lock to Video Controller Name and Drive Version Number

'###############################################################
'
'ADDED lockHDFirmware = 256
'Date: 4/21/05
'By: Scott Nelson (Sentax)
Public Enum ALLockTypes
    lockNone = 0
    lockMAC = 1             '8
    lockComp = 2
    lockHD = 4
    lockHDFirmware = 8      '256
    lockWindows = 16        '1
    lockBIOS = 32           '16
    lockMotherboard = 64
    lockIP = 128            '32
    lockExternalIP = 256    '128
    lockFingerprint = 512
    lockMemory = 1024
    lockCPUID = 2048
    lockBaseboardID = 4096
    lockvideoID = 8192
End Enum

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
' @param trialRegistryPerUser        ' Trial information is encrypted and hidden in several files
Public Enum ALTrialHideTypes
    trialSteganography = 1
    trialHiddenFolder = 2
    trialRegistryPerUser = 4
End Enum

'###############################################################
    
' Enum for accessing the Time Servers
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
' @param alsLicenseFileEncrypted    ' Encrypts the license file
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
' Name: Property Let LicenseFileType
' Input:
'    ByVal LicenseFileType As ALLicenseFileType - License File Type object
' Output: None
' Purpose: Interface Property. Specifies the license file type for this instance of ActiveLock.
' Remarks: None
'===============================================================================
Public Property Let LicenseFileType(ByVal LicenseFileType As ALLicenseFileTypes)

End Property

Public Property Get LicenseFileType() As Long

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
' Name: Property Get LastUsed
' Input: None
' Output:
'   String - Last Used date of the license
' Purpose: Last Used date of the license
' Remarks: None
'===============================================================================
Public Property Get LastUsed() As String

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
' Name: Property Let CheckTimeServerForClockTampering
' Input:
'   ByVal iServer As ALTimeServerTypes - Flag being passed to check the time server
' Output: None
' Purpose: Specifies whether a Time Server should be used to check Clock Tampering
' Remarks: None
'===============================================================================
Public Property Let CheckTimeServerForClockTampering(ByVal iServer As ALTimeServerTypes)

End Property

'===============================================================================
' Name: Property Let TrialWarning
' Input:
'   ByVal Value As ALTrialWarningTypes - Trial Warning is either Persistent or Temporary
' Output: None
' Purpose: Specifies whether the Trial Warning is either Persistent or Temporary
' Remarks: None
'===============================================================================
Public Property Let TrialWarning(ByVal Value As ALTrialWarningTypes)

End Property

'===============================================================================
' Name: Property Let AutoRegister
' Input:
'   ByVal iAuto As ALAutoRegisterTypes - Flag to indicate whether auto registration via an ALL file should be performed
' Output: None
' Purpose: Specifies whether auto registration via an ALL file should be performed
' Remarks: None
'===============================================================================
Public Property Let AutoRegister(ByVal iAuto As ALAutoRegisterTypes)

End Property

'===============================================================================
' Name: Property Let CheckSystemFilesForClockTampering
' Input:
'   ByVal iSystem As ALSystemFilesTypes - Flag being passed to check the system files
' Output: None
' Purpose: Specifies whether the system files should be scanned to check Clock Tampering
' Remarks: None
'===============================================================================
Public Property Let CheckSystemFilesForClockTampering(ByVal iSystem As ALSystemFilesTypes)

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
' Name: Sub KillLicense
' Input:
'   ByVal SoftwareNameAndVersion As String - Software Name And Version
'   ByVal LicPath As String - License file path and name
' Output: None
' Purpose: Kills a license using the specified liberation key.
' Remarks: None
'===============================================================================
Public Sub KillLicense(ByVal SoftwareNameAndVersion As String, ByVal LicPath As String)

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
Public Sub Acquire(Optional ByRef strMsg As String = "", Optional ByRef strRemainingTrialDays As String = "", Optional ByRef strRemainingTrialRuns As String = "", Optional ByRef strTrialLength As String = "", Optional ByRef strUsedDays As String = "", Optional ByRef strExpirationDate As String = "", Optional ByRef strRegisteredUser As String = "", Optional ByRef strRegisteredLevel As String = "", Optional ByRef strLicenseClass As String = "", Optional ByRef strMaxCount As String = "", Optional ByRef strLicenseFileType As String = "", Optional ByRef strLicenseType As String = "")
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
' Name: Property Get RemainingTrialDays
' Input: None
' Output:
'   Long - Number of used trial days
' Purpose: Returns the number of used trial days
' Remarks: None
'===============================================================================
Public Property Get RemainingTrialDays() As Long

End Property
'===============================================================================
' Name: Property Get RemainingTrialRuns
' Input: None
' Output:
'   Long - Number of used trial runs
' Purpose: Returns the number of used trial runs
' Remarks: None
'===============================================================================
Public Property Get RemainingTrialRuns() As Long

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

