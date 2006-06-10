Option Strict Off
Option Explicit On
<System.Runtime.InteropServices.ProgId("ProductLicense_NET.ProductLicense")> Public Class ProductLicense
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
	' Name: ProductLicense
	' Purpose: This class encapsulates a product license.  A product license contains
	' information such as the registered user, license type, product ID,
	' license key, etc...
	' Functions:
	' Properties:
	' Methods:
	' Started: 04.21.2005
    ' Modified: 03.25.2006
	'===============================================================================
	' @author activelock-admins
    ' @version 3.3.0
    ' @date 03.25.2006

	Private mstrType As String
	Private mstrLicensee As String
	Private mstrRegisteredLevel As String
	Private mstrLicenseClass As String
	Private mLicType As ALLicType
	Private mstrProductKey As String ' This is a transient property -- TODO: Remove this property
	Private mstrProductName As String
	Private mstrProductVer As String
	Private mstrLicenseKey As String
	Private mstrLicenseCode As String
	Private mstrExpiration As String
	Private mstrRegisteredDate As String
	Private mstrLastUsed As String
	Private mstrHash1 As String ' hash of mstrRegisteredDate
	Private mnMaxCount As Integer ' max number of concurrent users
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
	'===============================================================================
	' Name: Property Get RegisteredLevel
	' Input: None
	' Output:
	'   String - Registered license level
	' Purpose: Returns the registered level for this license.
	' Remarks: None
	'===============================================================================
	Public Property RegisteredLevel() As String
		Get
			RegisteredLevel = mstrRegisteredLevel
		End Get
		Set(ByVal Value As String)
			mstrRegisteredLevel = Value
		End Set
	End Property
	'===============================================================================
	' Name: Property Let LicenseType
	' Input:
	'   LicType As ALLicType - License type object
	' Output: None
	' Purpose: Specifies the license type for this instance of ActiveLock.
	' Remarks: None
	'===============================================================================
    '===============================================================================
	' Name: Property Get LicenseType
	' Input: None
	' Output:
	'   ALLicType - License type object
	' Purpose: Returns the License Type being used in this instance.
	' Remarks: None
	'===============================================================================
	Public Property LicenseType() As ALLicType
		Get
			LicenseType = mLicType
		End Get
		Set(ByVal Value As ALLicType)
			mLicType = Value
		End Set
	End Property
    '===============================================================================
	' Name: Property Let ProductName
	' Input:
	'   ByVal name As String - Product name string
	' Output: None
	' Purpose: [INTERNAL] Specifies the product name.
	' Remarks: None
	'===============================================================================
	'===============================================================================
	' Name: Property Get ProductName
	' Input: None
	' Output:
	'   String - Product name string
	' Purpose: Returns the product name.
	' Remarks: None
	'===============================================================================
	Public Property ProductName() As String
		Get
			Return mstrProductName
		End Get
		Set(ByVal Value As String)
			mstrProductName = Value
		End Set
	End Property
    '===============================================================================
	' Name: Property Let ProductVer
	' Input:
	'   ByVal Ver As String - Product version string
	' Output: None
	' Purpose: [INTERNAL] Specifies the product version.
	' Remarks: None
	'===============================================================================
	'===============================================================================
	' Name: Property Get ProductVer
	' Input: None
	' Output:
	'   String - Product version string
	' Purpose: Returns the product version string.
	' Remarks: None
	'===============================================================================
	Public Property ProductVer() As String
		Get
			Return mstrProductVer
		End Get
		Set(ByVal Value As String)
			mstrProductVer = Value
		End Set
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
	'===============================================================================
	' Name: Property Get ProductKey
	' Input: None
	' Output:
	'   String - Product Key (aka SoftwareCode)
	' Purpose: Returns the product key.
	' Remarks: None
	'===============================================================================
	Public Property ProductKey() As String
		Get
			ProductKey = mstrProductKey
		End Get
		Set(ByVal Value As String)
			mstrProductKey = Value
		End Set
	End Property
    '===============================================================================
	' Name: Property Let LicenseClass
	' Input:
	'   ByVal LicClass As String - License class string
	' Output: None
	' Purpose: [INTERNAL] Specifies the license class string.
	' Remarks: None
	'===============================================================================
    '===============================================================================
	' Name: Property Get LicenseClass
	' Input: None
	' Output:
	'   String - License class string
	' Purpose: Returns the license class string.
	' Remarks: None
	'===============================================================================
	Public Property LicenseClass() As String
		Get
			Return mstrLicenseClass
		End Get
		Set(ByVal Value As String)
			mstrLicenseClass = Value
		End Set
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
	'===============================================================================
	' Name: Property Get Licensee
	' Input: None
	' Output:
	'   String - Name of the licensed user
	' Purpose: Returns the person or organization registered to this license.
	' Remarks: None
	'===============================================================================
	Public Property Licensee() As String
		Get
			Licensee = mstrLicensee
		End Get
		Set(ByVal Value As String)
			mstrLicensee = Value
		End Set
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
    '===============================================================================
	' Name: Property Get LicenseKey
	' Input: None
	' Output:
	'   String - License key
	' Purpose: Returns the license key.
	' Remarks: None
	'===============================================================================
	Public Property LicenseKey() As String
		Get
			LicenseKey = mstrLicenseKey
		End Get
		Set(ByVal Value As String)
			mstrLicenseKey = Value
		End Set
	End Property
    '===============================================================================
	' Name: Property Let LicenseCode
	' Input:
	'   ByVal Code As String - New license code to be updated.
	' Output: None
	' Purpose: Updates the License Code.
	'===============================================================================
	'===============================================================================
	' Name: Property Get LicenseCode
	' Input: None
	' Output:
	'   String - License code
	' Purpose: Returns the license code.
	' Remarks: None
	'===============================================================================
	Public Property LicenseCode() As String
		Get
			LicenseCode = mstrLicenseCode
		End Get
		Set(ByVal Value As String)
			mstrLicenseCode = Value
		End Set
	End Property
    '===============================================================================
	' Name: Property Let Expiration
	' Input:
	'   ByVal strDate As String - Expiration date
	' Output: None
	' Purpose: [INTERNAL] Specifies expiration data.
	' Remarks: None
	'===============================================================================
	'===============================================================================
	' Name: Property Get Expiration
	' Input: None
	' Output:
	'   String - Expiration date
    ' Purpose: Returns the expiration date string in YYYY/MM/DD format.
	' Remarks: None
	'===============================================================================
	Public Property Expiration() As String
		Get
			Return mstrExpiration
		End Get
		Set(ByVal Value As String)
			mstrExpiration = Value
		End Set
	End Property
	'===============================================================================
	' Name: Property Get RegisteredDate
	' Input: None
	' Output:
	'   String - Product registration date
    ' Purpose: Returns the date in YYYY/MM/DD format on which the product was registered.
	' Remarks: None
	'===============================================================================
	'===============================================================================
	' Name: Property Let RegisteredDate
	' Input:
	'   ByVal strDate As String - Product registration date
	' Output: None
	' Purpose: [INTERNAL] Specifies the registered date.
	' Remarks: None
	'===============================================================================
	Public Property RegisteredDate() As String
		Get
			Return mstrRegisteredDate
		End Get
		Set(ByVal Value As String)
			mstrRegisteredDate = Value
		End Set
	End Property
    '===============================================================================
	' Name: Property Get MaxCount
	' Input: None
	' Output:
	'   Long - Maximum concurrent user count
	' Purpose: Returns maximum number of concurrent users
	' Remarks: None
	'===============================================================================
	
	'===============================================================================
	' Name: Property Let MaxCount
	' Input:
	'   nCount As Long - maximum number of concurrent users
	' Output: None
	' Purpose: Specifies maximum number of concurrent users
	' Remarks: None
	'===============================================================================
	Public Property MaxCount() As Integer
		Get
			Return mnMaxCount
		End Get
		Set(ByVal Value As Integer)
			mnMaxCount = Value
		End Set
	End Property
	'===============================================================================
	' Name: Property Get LastUsed
	' Input: None
	' Output:
	'   String - DateTiem string
    ' Purpose: Returns the date and time, in YYYY/MM/DD format, when the product was last run.
	' Remarks: None
	'===============================================================================
    '===============================================================================
	' Name: Property Let LastUsed
	' Input:
	'   ByVal strDateTime As String - Date and time string
	' Output: None
	' Purpose: [INTERNAL] Sets the last used date.
	' Remarks: None
	'===============================================================================
	Public Property LastUsed() As String
		Get
			Return mstrLastUsed
		End Get
		Set(ByVal Value As String)
			mstrLastUsed = Value
		End Set
	End Property
    '===============================================================================
	' Name: Property Get Hash1
	' Input: None
	' Output:
	'   String - Hash code
	' Purpose: Returns Hash-1 code. Hash-1 code is the encryption hash of the <code>LastUsed</code> property.
	' Remarks: None
	'===============================================================================
    '===============================================================================
	' Name: Property Let Hash1
	' Input:
	'   ByVal hcode As String - Hash code
	' Output: None
	' Purpose: [INTERNAL] Sets the Hash-1 code.
	' Remarks: None
	'===============================================================================
	Public Property Hash1() As String
		Get
			Return mstrHash1
		End Get
		Set(ByVal Value As String)
			mstrHash1 = Value
		End Set
	End Property
	'===============================================================================
	' Name: Function ToString
	' Input: None
	' Output:
	'   String - Formatted license string
	' Purpose: Returns a line-feed delimited string encoding of this object&#39;s properties.
	' Remarks: Note: LicenseKey is not included in this string.
	'===============================================================================
    Public Function ToString_Renamed() As String
        ToString_Renamed = ProductName & vbCrLf & ProductVer & vbCrLf & LicenseClass & vbCrLf & LicenseType & vbCrLf & Licensee & vbCrLf & RegisteredLevel & vbCrLf & RegisteredDate & vbCrLf & Expiration & vbCrLf & MaxCount
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
        Dim a() As String
        ' First take out all crlf characters
        strLic = Replace_Renamed(strLic, vbCrLf, "")

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
        LicenseType = CInt(arrParts(3))
        Licensee = arrParts(4)
        RegisteredLevel = arrParts(5)
        RegisteredDate = arrParts(6)
        Expiration = arrParts(7)
        MaxCount = CInt(arrParts(8))
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
        strOut = ToString_Renamed() & vbCrLf & LicenseKey 'add License Key at the end
        strOut = modBase64.Base64_Encode(strOut)
    End Sub
End Class