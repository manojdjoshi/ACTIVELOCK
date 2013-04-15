using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
namespace ActiveLock3_6NET
{
[System.Runtime.InteropServices.ProgId("ProductLicense_NET.ProductLicense")]

public class ProductLicense
{
	//*   ActiveLock
	//*   Copyright 1998-2002 Nelson Ferraz
	//*   Copyright 2006 The ActiveLock Software Group (ASG)
	//*   All material is the property of the contributing authors.
	//*
	//*   Redistribution and use in source and binary forms, with or without
	//*   modification, are permitted provided that the following conditions are
	//*   met:
	//*
	//*     [o] Redistributions of source code must retain the above copyright
	//*         notice, this list of conditions and the following disclaimer.
	//*
	//*     [o] Redistributions in binary form must reproduce the above
	//*         copyright notice, this list of conditions and the following
	//*         disclaimer in the documentation and/or other materials provided
	//*         with the distribution.
	//*
	//*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
	//*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
	//*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
	//*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
	//*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
	//*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
	//*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
	//*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
	//*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
	//*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
	//*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
	//*
	//*
	//===============================================================================
	// Name: ProductLicense
	// Purpose: This class encapsulates a product license.  A product license contains
	// information such as the registered user, license type, product ID,
	// license key, etc...
	// Functions:
	// Properties:
	// Methods:
	// Started: 04.21.2005
	// Modified: 03.25.2006
	//===============================================================================
	// @author activelock-admins
	// @version 3.3.0
	// @date 03.25.2006

	private string mstrType;
	private string mstrLicensee;
	private string mstrRegisteredLevel;
	private string mstrLicenseClass;
	private ALLicType mLicType;
		// This is a transient property -- TODO: Remove this property
	private string mstrProductKey;
	private string mstrProductName;
	private string mstrProductVer;
	private string mstrLicenseKey;
	private string mstrLicenseCode;
	private string mstrExpiration;
	private string mstrRegisteredDate;
	private string mstrLastUsed;
		// hash of mstrRegisteredDate
	private string mstrHash1;
		// max number of concurrent users
	private int mnMaxCount;
		// number of days the license was used
	private int mnUsedDays;
		// type of License file; 0 for Plain, 1 for encrypted
	private int mnLicenseFileType;
	//    Private mnUsedLockType As Integer

	// License Flags.  Values can be combined (OR&#39;ed) together.
	//
	// @param alfSingle        Single-user license
	// @param alfMulti         Multi-user license
	public enum LicFlags
	{
		alfSingle = 0,
		alfMulti = 1
	}
	// License Types.  Values are mutually exclusive. i.e. they cannot be OR&#39;ed together.
	//
	// @param allicNone        No license enforcement
	// @param allicPeriodic    License expires after X number of days
	// @param allicPermanent   License will never expire
	// @param allicTimeLocked  License expires on a particular date
	public enum ALLicType
	{
		allicNone = 0,
		allicPeriodic = 1,
		allicPermanent = 2,
		allicTimeLocked = 3
	}
	//===============================================================================
	// Name: Property Let RegisteredLevel
	// Input:
	//   ByVal rLevel As String - Registered license level
	// Output: None
	// Purpose: [INTERNAL] Specifies the registered level.
	// Remarks: !!! WARNING !!! Make sure you know what you're doing when you call this method; otherwise, you run
	// the risk of invalidating your existing license.
	//===============================================================================
	//===============================================================================
	// Name: Property Get RegisteredLevel
	// Input: None
	// Output:
	//   String - Registered license level
	// Purpose: Returns the registered level for this license.
	// Remarks: None
	//===============================================================================
	public string RegisteredLevel {
		get { RegisteredLevel = mstrRegisteredLevel; }
		set { mstrRegisteredLevel = value; }
	}
	//===============================================================================
	// Name: Property Let LicenseType
	// Input:
	//   LicType As ALLicType - License type object
	// Output: None
	// Purpose: Specifies the license type for this instance of ActiveLock.
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Get LicenseType
	// Input: None
	// Output:
	//   ALLicType - License type object
	// Purpose: Returns the License Type being used in this instance.
	// Remarks: None
	//===============================================================================
	public ALLicType LicenseType {
		get { LicenseType = mLicType; }
		set { mLicType = value; }
	}
	//===============================================================================
	// Name: Property Let ProductName
	// Input:
	//   ByVal name As String - Product name string
	// Output: None
	// Purpose: [INTERNAL] Specifies the product name.
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Get ProductName
	// Input: None
	// Output:
	//   String - Product name string
	// Purpose: Returns the product name.
	// Remarks: None
	//===============================================================================
	public string ProductName {
		get { return mstrProductName; }
		set { mstrProductName = value; }
	}
	//===============================================================================
	// Name: Property Let ProductVer
	// Input:
	//   ByVal Ver As String - Product version string
	// Output: None
	// Purpose: [INTERNAL] Specifies the product version.
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Get ProductVer
	// Input: None
	// Output:
	//   String - Product version string
	// Purpose: Returns the product version string.
	// Remarks: None
	//===============================================================================
	public string ProductVer {
		get { return mstrProductVer; }
		set { mstrProductVer = value; }
	}
	//===============================================================================
	// Name: Property Let ProductKey
	// Input:
	//   ByVal Key As String - Product key string
	// Output: None
	// Purpose: Specifies the product key.
	// Remarks: !!!WARNING!!! Use this method with caution.  You run the risk of invalidating your existing license
	// if you call this method without knowing what you are doing.
	//===============================================================================
	//===============================================================================
	// Name: Property Get ProductKey
	// Input: None
	// Output:
	//   String - Product Key (aka SoftwareCode)
	// Purpose: Returns the product key.
	// Remarks: None
	//===============================================================================
	public string ProductKey {
		get { ProductKey = mstrProductKey; }
		set { mstrProductKey = value; }
	}
	//===============================================================================
	// Name: Property Let LicenseClass
	// Input:
	//   ByVal LicClass As String - License class string
	// Output: None
	// Purpose: [INTERNAL] Specifies the license class string.
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Get LicenseClass
	// Input: None
	// Output:
	//   String - License class string
	// Purpose: Returns the license class string.
	// Remarks: None
	//===============================================================================
	public string LicenseClass {
		get { return mstrLicenseClass; }
		set { mstrLicenseClass = value; }
	}
	//===============================================================================
	// Name: Property Let Licensee
	// Input:
	//   ByVal name As String - Name of the licensed user
	// Output: None
	// Purpose: [INTERNAL] Specifies the licensed user.
	// Remarks: !!! WARNING !!! Make sure you know what you're doing when you call this method; otherwise, you run
	// the risk of invalidating your existing license.
	//===============================================================================
	//===============================================================================
	// Name: Property Get Licensee
	// Input: None
	// Output:
	//   String - Name of the licensed user
	// Purpose: Returns the person or organization registered to this license.
	// Remarks: None
	//===============================================================================
	public string Licensee {
		get { Licensee = mstrLicensee; }
		set { mstrLicensee = value; }
	}
	//===============================================================================
	// Name: Property Let LicenseKey
	// Input:
	//   ByVal Key As String - New license key to be updated.
	// Output: None
	// Purpose: Updates the License Key.
	// Remarks: !!! WARNING !!! Make sure you know what you're doing when you call this method; otherwise, you run
	// the risk of invalidating your existing license.
	//===============================================================================
	//===============================================================================
	// Name: Property Get LicenseKey
	// Input: None
	// Output:
	//   String - License key
	// Purpose: Returns the license key.
	// Remarks: None
	//===============================================================================
	public string LicenseKey {
		get { LicenseKey = mstrLicenseKey; }
		set { mstrLicenseKey = value; }
	}
	//===============================================================================
	// Name: Property Let LicenseCode
	// Input:
	//   ByVal Code As String - New license code to be updated.
	// Output: None
	// Purpose: Updates the License Code.
	//===============================================================================
	//===============================================================================
	// Name: Property Get LicenseCode
	// Input: None
	// Output:
	//   String - License code
	// Purpose: Returns the license code.
	// Remarks: None
	//===============================================================================
	public string LicenseCode {
		get { LicenseCode = mstrLicenseCode; }
		set { mstrLicenseCode = value; }
	}
	//===============================================================================
	// Name: Property Let LicenseFileType
	// Input:
	//   ByVal Code As String - License file type: 0 for Plain, 1 for encrypted
	// Output: None
	// Purpose: Updates the License Code.
	//===============================================================================
	//===============================================================================
	// Name: Property Get LicenseFileType
	// Input: None
	// Output:
	//   Integer - License file type: 0 for Plain, 1 for encrypted
	// Purpose: Returns the license file type.
	// Remarks: None
	//===============================================================================
	public int LicenseFileType {
		get { LicenseFileType = mnLicenseFileType; }
		set { mnLicenseFileType = value; }
	}
	//===============================================================================
	// Name: Property Let Expiration
	// Input:
	//   ByVal strDate As String - Expiration date
	// Output: None
	// Purpose: [INTERNAL] Specifies expiration data.
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Get Expiration
	// Input: None
	// Output:
	//   String - Expiration date
	// Purpose: Returns the expiration date string in "yyyy/MM/dd" format.
	// Remarks: None
	//===============================================================================
	public string Expiration {
		get { return mstrExpiration; }
		set { mstrExpiration = value; }
	}
	//===============================================================================
	// Name: Property Get RegisteredDate
	// Input: None
	// Output:
	//   String - Product registration date
	// Purpose: Returns the date in "yyyy/MM/dd" format on which the product was registered.
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Let RegisteredDate
	// Input:
	//   ByVal strDate As String - Product registration date
	// Output: None
	// Purpose: [INTERNAL] Specifies the registered date.
	// Remarks: None
	//===============================================================================
	public string RegisteredDate {
		get { return mstrRegisteredDate; }
		set { mstrRegisteredDate = value; }
	}
	//===============================================================================
	// Name: Property Get MaxCount
	// Input: None
	// Output:
	//   Long - Maximum concurrent user count
	// Purpose: Returns maximum number of concurrent users
	// Remarks: None
	//===============================================================================

	//===============================================================================
	// Name: Property Let MaxCount
	// Input:
	//   nCount As Long - maximum number of concurrent users
	// Output: None
	// Purpose: Specifies maximum number of concurrent users
	// Remarks: None
	//===============================================================================
	public int MaxCount {
		get { return mnMaxCount; }
		set { mnMaxCount = value; }
	}
	//===============================================================================
	// Name: Property Get UsedDays
	// Input: None
	// Output:
	//   Integer - Number of Days the License was used
	// Purpose: Returns Number of Days the License was used
	// Remarks: None
	//===============================================================================

	//===============================================================================
	// Name: Property Let UsedDays
	// Input:
	//   nCount As Integer - Number of Days the License was used
	// Output: None
	// Purpose: Specifies Number of Days the License was used
	// Remarks: None
	//===============================================================================
	public int UsedDays {
		get { return mnUsedDays; }
		set { mnUsedDays = value; }
	}
	//===============================================================================
	// Name: Property Get UsedLockType
	// Input: None
	// Output:
	//   Integer - Used Lock type in Alugen
	// Purpose: Returns Used Lock type in Alugen
	// Remarks: None
	//===============================================================================
	//Public Property UsedLockType() As Integer
	//    Get
	//        Return mnUsedLockType
	//    End Get
	//    Set(ByVal Value As Integer)
	//        mnUsedLockType = Value
	//    End Set
	//End Property

	//===============================================================================
	// Name: Property Get LastUsed
	// Input: None
	// Output:
	//   String - DateTiem string
	// Purpose: Returns the date and time, in "yyyy/MM/dd" format, when the product was last run.
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Let LastUsed
	// Input:
	//   ByVal strDateTime As String - Date and time string
	// Output: None
	// Purpose: [INTERNAL] Sets the last used date.
	// Remarks: None
	//===============================================================================
	public string LastUsed {
		get { return mstrLastUsed; }
		set { mstrLastUsed = value; }
	}
	//===============================================================================
	// Name: Property Get Hash1
	// Input: None
	// Output:
	//   String - Hash code
	// Purpose: Returns Hash-1 code. Hash-1 code is the encryption hash of the <code>LastUsed</code> property.
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Let Hash1
	// Input:
	//   ByVal hcode As String - Hash code
	// Output: None
	// Purpose: [INTERNAL] Sets the Hash-1 code.
	// Remarks: None
	//===============================================================================
	public string Hash1 {
		get { return mstrHash1; }
		set { mstrHash1 = value; }
	}
	//===============================================================================
	// Name: Function ToString
	// Input: None
	// Output:
	//   String - Formatted license string
	// Purpose: Returns a line-feed delimited string encoding of this object&#39;s properties.
	// Remarks: Note: LicenseKey is not included in this string.
	//===============================================================================
	public string ToString_Renamed()
	{
		return ProductName + Constants.vbCrLf + ProductVer + Constants.vbCrLf + LicenseClass + Constants.vbCrLf + LicenseType + Constants.vbCrLf + Licensee + Constants.vbCrLf + RegisteredLevel + Constants.vbCrLf + RegisteredDate + Constants.vbCrLf + Expiration + Constants.vbCrLf + MaxCount;
	}
	//===============================================================================
	// Name: Sub Load
	// Input:
	//   ByVal strLic As String - Formatted license string, delimited by CrLf characters.
	// Output: None
	// Purpose: Loads the license from a formatted string created from <a href="ProductLicense.Save.html">Save()</a>.
	// Remarks: None
	//===============================================================================
	public void Load(string strLic)
	{
		string[] a = null;
		// First take out all crlf characters
		strLic = modTrial.Replace_Renamed(strLic, Constants.vbCrLf, "");

		// New in v3.1
		// Installation code is now appended to the end of the liberation key
		// because Alugen has the ability to modify it based on
		// the selected hardware lock keys by the user

		// Split the license key in two parts
		a = Strings.Split(strLic, "aLck");
		// The second part is the new installation code
		strLic = a[0];

		// New in v3.1
		LicenseCode = a[1];

		// base64-decode it
		strLic = modBase64.Base64_Decode(ref strLic);

		string[] arrParts = null;
		arrParts = Strings.Split(strLic, Constants.vbCrLf);
		// Initialize appropriate properties
		ProductName = arrParts[0];
		ProductVer = arrParts[1];
		LicenseClass = arrParts[2];
		LicenseType = (int)arrParts[3];
		Licensee = arrParts[4];
		RegisteredLevel = arrParts[5];
		RegisteredDate = arrParts[6];
		Expiration = arrParts[7];
		MaxCount = (int)arrParts[8];
		LicenseKey = arrParts[9];
	}
	//===============================================================================
	// Name: Sub Save
	// Input:
	//   ByRef strOut As String - Formatted license string will be saved into this parameter when the routine returns
	// Output: None
	// Purpose: Saves the license into a formatted string.
	// Remarks: None
	//===============================================================================
	public void Save(ref string strOut)
	{
		strOut = ToString_Renamed() + Constants.vbCrLf + LicenseKey;
		//add License Key at the end
		strOut = modBase64.Base64_Encode(ref strOut);
	}
}
}