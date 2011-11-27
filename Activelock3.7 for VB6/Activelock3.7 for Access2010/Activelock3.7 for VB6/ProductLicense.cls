VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProductLicense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
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
' *   Copyright 2003-2012 The ActiveLock Software Group (admin: Ismail Alkan)
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
' Name: ProductLicense
' Purpose: This class encapsulates a product license.  A product license contains
' information such as the registered user, license type, product ID,
' license key, etc...
' Functions:
' Properties:
' Methods:
' Started: 04.21.2005
' Modified: 08.15.2005
'===============================================================================
' @author activelock-admins
' @version 3.0.0
' @date 20050421
'
' * ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
'
'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' <pre>
'   06.16.03 - th2tran  - Created
'   07.07.03 - mecrute  - Updated the header comments for this file.
'   08.03.03 - th2tran  - VBDox&#39;ed this class.
'   10.12.03 - th2tran  - Added VB doc Attributes for showing in VB Object Browser
'                       - Load() now drops CRLF chars from the liberation key.
'                         This allows users to format the key into nice 64-byte
'   11.02.03 - th2tran  - New license properties: License Type, MaxCount
'                       - Removed CLASS_n constants
'                       - Moved LicFlags definition from IActiveLock into here
'   04.17.04 - th2tran  - Don't save product key property--it is now considered a
'                         transient (non-serializable) property.  Saving the product key is unnecessary
'                         and will make our liberation key look (unnecessarily) horrendous.
'   05.13.05 - ialkan   - Modified code to include the Register to Level X feature
' </pre>


'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit

Private mstrLicensee As String
Private mstrRegisteredLevel As String
Private mstrLicenseClass As String
Private mLicType As ALLicType
Private mstrProductKey As String    ' This is a transient property -- TODO: Remove this property
Private mstrProductName As String
Private mstrProductVer As String
Private mstrLicenseKey As String
Private mstrLicenseCode As String
'* Private mstrExpiration As String
'* Private mstrRegisteredDate As String
'* Private mstrLastUsed As String
Private mstrExpiration As Date '*
Private mstrRegisteredDate As Date '*
Private mstrLastUsed As Date '*
Private mstrHash1 As String ' hash of mstrRegisteredDate
Private mnMaxCount As Long ' max number of concurrent users
Private mnLicenseFileType As Long

' License Flags.  Values can be combined (OR&#39;ed) together.
'
' @param alfSingle        Single-user license
' @param alfMulti         Multi-user license
Public Enum LicFlags
    alfSingle = 0
    alfMulti = 1
End Enum

' License Types.  Values are mutually exclusive. i.e. they cannot be OR&#39;ed together.
'
' @param allicNone        No license enforcement
' @param allicPeriodic    License expires after X number of days
' @param allicPermanent   License will never expire
' @param allicTimeLocked  License expires on a particular date
Public Enum ALLicType
    allicNone = 0
    allicPeriodic = 1
    allicPermanent = 2
    allicTimeLocked = 3
End Enum


'===============================================================================
' Name: Property Let RegisteredLevel
' Input:
'   ByVal rLevel As String - Registered license level
' Output: None
' Purpose: [INTERNAL] Specifies the registered level.
' Remarks: !!! WARNING !!! Make sure you know what you're doing when you call this method; otherwise, you run
' the risk of invalidating your existing license.
'===============================================================================
Public Property Let RegisteredLevel(ByVal rLevel As String)
    mstrRegisteredLevel = rLevel
End Property
'===============================================================================
' Name: Property Get RegisteredLevel
' Input: None
' Output:
'   String - Registered license level
' Purpose: Returns the registered level for this license.
' Remarks: None
'===============================================================================
Public Property Get RegisteredLevel() As String
    RegisteredLevel = mstrRegisteredLevel
End Property
'===============================================================================
' Name: Property Let LicenseType
' Input:
'   LicType As ALLicType - License type object
' Output: None
' Purpose: Specifies the license type for this instance of ActiveLock.
' Remarks: None
'===============================================================================
Public Property Let LicenseType(LicType As ALLicType)
Attribute LicenseType.VB_Description = "License Type"
    mLicType = LicType
End Property
'===============================================================================
' Name: Property Let LicenseFileType
' Input:
'   LicFileType As ALLicenseFileTypes - License file type
' Output: None
' Purpose: Specifies the license file type for this instance of ActiveLock.
' Remarks: None
'===============================================================================
Public Property Let LicenseFileType(LicFileType As Long)
    mnLicenseFileType = LicFileType
End Property


'===============================================================================
' Name: Property Get LicenseType
' Input: None
' Output:
'   ALLicType - License type object
' Purpose: Returns the License Type being used in this instance.
' Remarks: None
'===============================================================================
Public Property Get LicenseType() As ALLicType
Attribute LicenseType.VB_Description = "License Type"
    LicenseType = mLicType
End Property

'===============================================================================
' Name: Property Get LicenseFileType
' Input: None
' Output:
'   ALLicenseFileTypes - License file type object
' Purpose: Returns the License File Type being used in this instance.
' Remarks: None
'===============================================================================
Public Property Get LicenseFileType() As Long
    LicenseFileType = mnLicenseFileType
End Property

'===============================================================================
' Name: Property Let ProductName
' Input:
'   ByVal name As String - Product name string
' Output: None
' Purpose: [INTERNAL] Specifies the product name.
' Remarks: None
'===============================================================================
Friend Property Let ProductName(ByVal name As String)
    mstrProductName = name
End Property
'===============================================================================
' Name: Property Get ProductName
' Input: None
' Output:
'   String - Product name string
' Purpose: Returns the product name.
' Remarks: None
'===============================================================================
Public Property Get ProductName() As String
Attribute ProductName.VB_Description = "Returns the product name."
    ProductName = mstrProductName
End Property


'===============================================================================
' Name: Property Let ProductVer
' Input:
'   ByVal Ver As String - Product version string
' Output: None
' Purpose: [INTERNAL] Specifies the product version.
' Remarks: None
'===============================================================================
Friend Property Let ProductVer(ByVal Ver As String)
    mstrProductVer = Ver
End Property
'===============================================================================
' Name: Property Get ProductVer
' Input: None
' Output:
'   String - Product version string
' Purpose: Returns the product version string.
' Remarks: None
'===============================================================================
Public Property Get ProductVer() As String
Attribute ProductVer.VB_Description = "Returns the product version string."
    ProductVer = mstrProductVer
End Property
'===============================================================================
' Name: Property Let ProductKey
' Input:
'   ByVal Key As String - Product key string
' Output: None
' Purpose: Specifies the product key.
' Remarks: !!!WARNING!!! Use this method with caution.  You run the risk of invalidating your existing license
' if you call this method without knowing what you are doing.
'===============================================================================
Public Property Let ProductKey(ByVal Key As String)
Attribute ProductKey.VB_Description = "Product Key."
    mstrProductKey = Key
End Property
'===============================================================================
' Name: Property Get ProductKey
' Input: None
' Output:
'   String - Product Key (aka SoftwareCode)
' Purpose: Returns the product key.
' Remarks: None
'===============================================================================
Public Property Get ProductKey() As String
    ProductKey = mstrProductKey
End Property


'===============================================================================
' Name: Property Let LicenseClass
' Input:
'   ByVal LicClass As String - License class string
' Output: None
' Purpose: [INTERNAL] Specifies the license class string.
' Remarks: None
'===============================================================================
Friend Property Let LicenseClass(ByVal LicClass As String)
    mstrLicenseClass = LicClass
End Property


'===============================================================================
' Name: Property Get LicenseClass
' Input: None
' Output:
'   String - License class string
' Purpose: Returns the license class string.
' Remarks: None
'===============================================================================
Public Property Get LicenseClass() As String
Attribute LicenseClass.VB_Description = "License class."
    LicenseClass = mstrLicenseClass
End Property


'===============================================================================
' Name: Property Let Licensee
' Input:
'   ByVal name As String - Name of the licensed user
' Output: None
' Purpose: [INTERNAL] Specifies the licensed user.
' Remarks: !!! WARNING !!! Make sure you know what you're doing when you call this method; otherwise, you run
' the risk of invalidating your existing license.
'===============================================================================
Public Property Let Licensee(ByVal name As String)
    mstrLicensee = name
End Property
'===============================================================================
' Name: Property Get Licensee
' Input: None
' Output:
'   String - Name of the licensed user
' Purpose: Returns the person or organization registered to this license.
' Remarks: None
'===============================================================================
Public Property Get Licensee() As String
Attribute Licensee.VB_Description = "Person or organization registered to this license."
    Licensee = mstrLicensee
End Property


'===============================================================================
' Name: Property Let LicenseKey
' Input:
'   ByVal Key As String - New license key to be updated.
' Output: None
' Purpose: Updates the License Key.
' Remarks: !!! WARNING !!! Make sure you know what you're doing when you call this method; otherwise, you run
' the risk of invalidating your existing license.
'===============================================================================
Public Property Let LicenseKey(ByVal Key As String)
Attribute LicenseKey.VB_Description = "License key."
    mstrLicenseKey = Key
End Property

'===============================================================================
' Name: Property Let LicenseCode
' Input:
'   ByVal Code As String - New license code to be updated.
' Output: None
' Purpose: Updates the License Code.
'===============================================================================
Public Property Let LicenseCode(ByVal Code As String)
    mstrLicenseCode = Code
End Property

'===============================================================================
' Name: Property Get LicenseKey
' Input: None
' Output:
'   String - License key
' Purpose: Returns the license key.
' Remarks: None
'===============================================================================
Public Property Get LicenseKey() As String
    LicenseKey = mstrLicenseKey
End Property
'===============================================================================
' Name: Property Get LicenseCode
' Input: None
' Output:
'   String - License code
' Purpose: Returns the license code.
' Remarks: None
'===============================================================================
Public Property Get LicenseCode() As String
    LicenseCode = mstrLicenseCode
End Property

'===============================================================================
' Name: Property Let Expiration
' Input:
'   ByVal strDate As Date - Expiration date '*
' Output: None
' Purpose: [INTERNAL] Specifies expiration data.
' Remarks: None
'===============================================================================
Friend Property Let Expiration(ByVal strDate As Date) '* String)
    mstrExpiration = strDate
End Property
'===============================================================================
' Name: Property Get Expiration
' Input: None
' Output:
'   Date - Expiration date '*
' Purpose: Returns the expiration date string in "yyyy/MM/dd" format.
' Remarks: None
'===============================================================================
Public Property Get Expiration() As Date '* String
Attribute Expiration.VB_Description = "Expiration date string in YYYY/MM/DD format."
    Expiration = mstrExpiration
End Property
'===============================================================================
' Name: Property Get RegisteredDate
' Input: None
' Output:
'   Date - Product registration date '*
' Purpose: Returns the date on which the product was registered. '*
' Remarks: None
'===============================================================================
Public Property Get RegisteredDate() As Date '* String
Attribute RegisteredDate.VB_Description = "Date, in YYYY/MM/DD string format, on which the product was registered."
    RegisteredDate = mstrRegisteredDate
End Property
'===============================================================================
' Name: Property Let RegisteredDate
' Input:
'   ByVal strDate As Date - Product registration date '*
' Output: None
' Purpose: [INTERNAL] Specifies the registered date.
' Remarks: None
'===============================================================================
Friend Property Let RegisteredDate(ByVal strDate As Date) '* String)
    mstrRegisteredDate = strDate
End Property


'===============================================================================
' Name: Property Get MaxCount
' Input: None
' Output:
'   Long - Maximum concurrent user count
' Purpose: Returns maximum number of concurrent users
' Remarks: None
'===============================================================================
Public Property Get MaxCount() As Long
    MaxCount = mnMaxCount
End Property

'===============================================================================
' Name: Property Let MaxCount
' Input:
'   nCount As Long - maximum number of concurrent users
' Output: None
' Purpose: Specifies maximum number of concurrent users
' Remarks: None
'===============================================================================
Friend Property Let MaxCount(nCount As Long)
    mnMaxCount = nCount
End Property
'===============================================================================
' Name: Property Get LastUsed
' Input: None
' Output:
'   Date - DateTime string '*
' Purpose: Returns the date when the product was last run.
' Remarks: None
'===============================================================================
Public Property Get LastUsed() As Date '* String
Attribute LastUsed.VB_Description = "The date and time, in YYYY/MM/DD HH:MM:SS format, when the product was last run."
    LastUsed = mstrLastUsed
End Property


'===============================================================================
' Name: Property Let LastUsed
' Input:
'   ByVal strDateTime As Date '*
' Output: None
' Purpose: [INTERNAL] Sets the last used date.
' Remarks: None
'===============================================================================
Friend Property Let LastUsed(ByVal strDateTime As Date) '* String)
    mstrLastUsed = strDateTime
End Property

'===============================================================================
' Name: Property Get Hash1
' Input: None
' Output:
'   String - Hash code
' Purpose: Returns Hash-1 code. Hash-1 code is the encryption hash of the <code>LastUsed</code> property.
' Remarks: None
'===============================================================================
Public Property Get Hash1() As String
Attribute Hash1.VB_Description = "Returns Hash-1 code. Hash-1 code is the encryption hash of the LastUsed property."
    Hash1 = mstrHash1
End Property


'===============================================================================
' Name: Property Let Hash1
' Input:
'   ByVal hcode As String - Hash code
' Output: None
' Purpose: [INTERNAL] Sets the Hash-1 code.
' Remarks: None
'===============================================================================
Friend Property Let Hash1(ByVal hcode As String)
    mstrHash1 = hcode
End Property
'===============================================================================
' Name: Function ToString
' Input: None
' Output:
'   String - Formatted license string
' Purpose: Returns a line-feed delimited string encoding of this object&#39;s properties.
' Remarks: Note: LicenseKey is not included in this string.
'===============================================================================
Public Function ToString() As String
Attribute ToString.VB_Description = "Returns a line-feed delimited string encoding of this object's properties."
'* ToString = ProductName & vbCrLf & _
            ProductVer & vbCrLf & _
            LicenseClass & vbCrLf & _
            LicenseType & vbCrLf & _
            Licensee & vbCrLf & _
            RegisteredLevel & vbCrLf & _
            RegisteredDate & vbCrLf & _
            Expiration & vbCrLf & _
            MaxCount
'* remember to save the dates as doubles!
ToString = ProductName & vbCrLf & _
            ProductVer & vbCrLf & _
            LicenseClass & vbCrLf & _
            LicenseType & vbCrLf & _
            Licensee & vbCrLf & _
            RegisteredLevel & vbCrLf & _
            DateToDblString(RegisteredDate) & vbCrLf & _
            DateToDblString(Expiration) & vbCrLf & _
            MaxCount
End Function
'===============================================================================
' Name: Sub Load
' Input:
'   ByVal strLic As String - Formatted license string, delimited by CrLf characters.
' Output: None
' Purpose: Loads the license from a formatted string created from <a href="ProductLicense.Save.html">Save()</a>.
' Remarks: None
'===============================================================================
Public Sub Load(ByVal strLic As String)
Attribute Load.VB_Description = "Loads the license from a formatted string created by Save() method."
    Dim a() As String
    
    ' First take out all crlf characters
    strLic = Replace(strLic, vbCrLf, "")
    
    ' New in v3.1
    ' Installation code is now appended to the end of the liberation key
    ' because Alugen has the ability to modify it based on
    ' the selected hardware lock keys by the user
    
    ' Split the license key in two parts
    a = Split(strLic, "aLck")
    ' The second part is the new installation code
    strLic = a(0)
    
    ' New in v3.1
    LicenseCode = a(1)
    
    ' base64-decode it
    strLic = modBase64.Base64_Decode(strLic)
    
    Dim arrParts() As String
    arrParts = Split(strLic, vbCrLf)
    ' Initialize appropriate properties
    ProductName = arrParts(0)
    ProductVer = arrParts(1)
    LicenseClass = arrParts(2)
    LicenseType = CLng(arrParts(3))
    Licensee = arrParts(4)
    RegisteredLevel = arrParts(5)
    '* RegisteredDate = arrParts(6)
    '* Expiration = arrParts(7)
    RegisteredDate = DblStringToDate(arrParts(6)) '*
    Expiration = DblStringToDate(arrParts(7)) '*
    RegisteredDate = arrParts(6)
    Expiration = arrParts(7)
    MaxCount = CLng(arrParts(8))
    LicenseKey = arrParts(9)
    
End Sub


'===============================================================================
' Name: Sub Save
' Input:
'   ByRef strOut As String - Formatted license string will be saved into this parameter when the routine returns
' Output: None
' Purpose: Saves the license into a formatted string.
' Remarks: None
'===============================================================================
Public Sub Save(ByRef strOut As String)
Attribute Save.VB_Description = "Saves the license into a formatted string."
    strOut = ToString() & vbCrLf & LicenseKey 'add License Key at the end
    strOut = modBase64.Base64_Encode(strOut)
End Sub

Private Sub Class_Initialize()
mstrRegisteredDate = #1/1/1900#
mstrExpiration = #1/1/1900#
mstrLastUsed = #1/1/1900#
End Sub

