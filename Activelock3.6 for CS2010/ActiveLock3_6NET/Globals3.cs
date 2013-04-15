using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
using System.Security.Cryptography;
using System.Text;
using System.IO;
using ActiveLock3_6NET;

//Class instancing was changed to public
[System.Runtime.InteropServices.ProgId("Globals_NET.Globals")]
public class Globals
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
	// Name: Globals
	// Purpose: This class contains global object factory and utility methods and constants.
	// <p>It is a global class so its routines in here can be accessed directly
	// from the ActiveLock3 namespace.
	// For example, the <code>NewInstance()</code> function can be accessed via
	// <code>ActiveLock3.NewInstance()</code>.
	// Functions:
	// Properties:
	// Methods:
	// Started: 04.21.2005
	// Modified: 03.24.2006
	//===============================================================================
	// @author activelock-admins
	// @version 3.3.0
	// @date 03.24.2006
	//

	// ActiveLock Error Codes.
	// These error codes are used for <code>Err.Number</code> whenever ActiveLock raises an error.
	//
	// @param alerrOK                    No error. Operation was successful.
	// @param alerrNoLicense             No license available.
	// @param alerrLicenseInvalid        License is invalid.
	// @param alerrLicenseExpired        License has expired.
	// @param alerrLicenseTampered       License has been tampered.
	// @param alerrClockChanged          System clock has been changed.
	// @param alerrWrongIPaddress        Wrong IP Address.
	// @param alerrKeyStoreInvalid       Key Store Provider has not been initialized yet.
	// @param alerrKeyStorePathInvalid   Key Store Path (LIC file path) hasn't been specified.
	// @param alerrFileTampered          ActiveLock DLL file has been tampered.
	// @param alerrNotInitialized        ActiveLock has not been initialized yet.
	// @param alerrNotImplemented        An ActiveLock operation has not been implemented.
	// @param alerrUserNameTooLong       Maximum User Name length of 2000 characters has been exceeded.
	// @param alerrUserNameInvalid       Used User name does not match with the license key.
	// @param alerrInvalidTrialDays      Specified number of Free Trial Days is invalid (possibly <=0).
	// @param alerrInvalidTrialRuns      Specified number of Free Trial Runs is invalid (possibly <=0).
	// @param alerrTrialInvalid          Trial is invalid.
	// @param alerrTrialDaysExpired      Trial Days have expired.
	// @param alerrTrialRunsExpired      Trial Runs have expired.
	// @param alerrNoSoftwareName        Software Name has not been specified.
	// @param alerrNoSoftwareVersion     Software Version has not been specified.
	// @param alerrRSAError              Something went wrong in the RSA routines.
	// @param alerrNoSoftwarePassword    Software Password has not been specified.
	// @param alerrCryptoAPIError        Crypto API error in CryptoAPI class.
	// @param alerrUndefinedSpecialFolder        The special folder used by Activelock is not defined or Virtual folder.
	// @param alerrDateError             There's an error in setting a date used by Activelock.

	public enum ActiveLockErrCodeConstants
	{
		AlerrOK = 0,
		// successful
		
	}
   private uint AlerrNoLicense = 0x80040001;
		// vbObjectError (&H80040000) + 1
		private uint AlerrLicenseInvalid = 0x80040002;
		private uint AlerrLicenseExpired = 0x80040003;
		private uint AlerrLicenseTampered = 0x80040004;
		private uint AlerrClockChanged = 0x80040005;
		private uint AlerrWrongIPaddress = 0x80040006;
		private uint AlerrKeyStoreInvalid = 0x80040010;
		private uint AlerrFileTampered = 0x80040011;
		private uint AlerrNotInitialized = 0x80040012;
		private uint AlerrNotImplemented = 0x80040013;
		private uint AlerrUserNameTooLong = 0x80040014;
        private uint AlerrInvalidTrialDays = 0x80040020;
		private uint AlerrInvalidTrialRuns = 0x80040021;
		private uint AlerrTrialInvalid = 0x80040022;
		private uint AlerrTrialDaysExpired = 0x80040023;
		private uint AlerrTrialRunsExpired = 0x80040024;
		private uint AlerrNoSoftwareName = 0x80040025;
		private uint AlerrNoSoftwareVersion = 0x80040026;
		private uint AlerrRSAError = 0x80040027;
		private uint AlerrKeyStorePathInvalid = 0x80040028;
		private uint AlerrCryptoAPIError = 0x80040029;
		private uint AlerrNoSoftwarePassword = 0x80040030;
		private uint AlerrUndefinedSpecialFolder = 0x80040031;
		private uint AlerrDateError = 0x80040032;
		private uint AlerrInternetConnectionError = 0x80040033;
		private uint AlerrSoftwarePasswordInvalid = 0x80040034;

	private string strCypherText;
	private bool bCypherOn;
	//===============================================================================
	// Name: Function NewInstance
	// Input: None
	// Output: ActiveLock interface.
	// Purpose: Obtains a new instance of an object that implements IActiveLock interface.
	// <p>As of 2.0.5, this method will no longer initialize the instance automatically.
	// Callers will have to call Init() by themselves subsequent to obtaining the instance.
	// Remarks: None
	//===============================================================================
	public _IActiveLock NewInstance()
	{
		_IActiveLock NewInst = null;
        NewInst = new _ActiveLock();
		return NewInst;
	}
	//===============================================================================
	// Name: Function CreateProductLicense
	// Input:
	//   ByVal name As String - Product/Software Name
	//   ByVal Ver As String - Product version
	//   ByVal Code As String - Product/Software Code
	//   ByVal Flags As ActiveLock3.LicFlags - License Flag
	//   ByVal LicType As ActiveLock3.ALLicType - License type
	//   ByVal Licensee As String - Registered party for which the license has been issued
	//   ByVal RegisteredLevel As String - Registered level
	//   ByVal Expiration As String - Expiration date
	//   ByVal LicKey As String - License key
	//   ByVal RegisteredDate As String - Date on which the product is registered
	//   ByVal Hash1 As String - Hash-1 code
	//   ByVal MaxUsers As Integer - Maximum number of users allowed to use this license
	// Output:
	//   ProductLicense - License object
	// Purpose: Instantiates a new ProductLicense object from the specified parameters.
	// <p>If <code>LicType</code> is <i>Permanent</i>, then <code>Expiration</code> date parameter will be ignored.
	// Remarks: None
	//===============================================================================
	public ProductLicense CreateProductLicense(string Name, string Ver, string Code, ProductLicense.LicFlags Flags, ProductLicense.ALLicType LicType, string Licensee, string RegisteredLevel, string Expiration, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute("")]  // ERROR: Optional parameters aren't supported in C#
string LicKey, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute("")]  // ERROR: Optional parameters aren't supported in C#
string RegisteredDate, 
	[System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute("")]  // ERROR: Optional parameters aren't supported in C#
string Hash1, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(1)]  // ERROR: Optional parameters aren't supported in C#
short MaxUsers, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute("")]  // ERROR: Optional parameters aren't supported in C#
string LicCode)
	{
		ProductLicense NewLic = new ProductLicense();
		{
			NewLic.ProductName = Name;
			NewLic.ProductKey = Code;
			NewLic.ProductVer = Ver;
			//If LicType = allicNetwork Then
			//    .LicenseClass = alfMulti
			//Else
			NewLic.LicenseClass = GetClassString(ref Flags);
			//End If
			NewLic.LicenseType = LicType;
			NewLic.Licensee = Licensee;
			NewLic.RegisteredLevel = RegisteredLevel;
			NewLic.MaxCount = MaxUsers;
			// ignore expiration date if license type is "permanent"
			if (LicType != ProductLicense.ALLicType.allicPermanent) {
				NewLic.Expiration = Expiration;
			}
			//IsMissing() was changed to IsNothing()
			if ((LicKey != null)) {
				NewLic.LicenseKey = LicKey;
			}
			//IsMissing() was changed to IsNothing()
			if ((RegisteredDate != null)) {
				NewLic.RegisteredDate = RegisteredDate;
			}
			//IsMissing() was changed to IsNothing()
			if ((Hash1 != null)) {
				NewLic.Hash1 = Hash1;
			}
			// New in v3.1
			// LicenseCode is appended to the end so that we can know
			// Alugen specified the hardware keys, and LockType
			// was not specified by the protected app
			//IsMissing() was changed to IsNothing()
			if ((LicCode != null)) {
				if (!string.IsNullOrEmpty(LicCode)) NewLic.LicenseCode = LicCode; 
			}
		}
		return NewLic;
	}
	//===============================================================================
	// Name: Function GetClassString
	// Input:
	//   ByRef Flags As ActiveLock3.LicFlags - License flag string
	// Output:
	//   String - License flag string
	// Purpose: Gets the license flag string such as MultiUser or Single
	// Remarks: None
	//===============================================================================
	private string GetClassString(ref ProductLicense.LicFlags Flags)
	{
		string functionReturnValue = null;
		// TODO: Decide the class numbers.
		// lockMAC should probably be last,
		// like it is in the enum. (IActivelock.cls)
		if (Flags == ProductLicense.LicFlags.alfMulti) {
			functionReturnValue = "MultiUser";
		}
		// default
		else {
			functionReturnValue = "Single";
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function GetLicTypeString
	// Input:
	//   LicType As ALLicType - License type object
	// Output:
	//   String - License type, such as Period, Permanent, Timed Expiry or None
	// Purpose: Returns a string version of LicType
	// Remarks: None
	//===============================================================================
	private string GetLicTypeString(ref ProductLicense.ALLicType LicType)
	{
		string functionReturnValue = null;
		//TODO: Implement this properly.
		if (LicType == ProductLicense.ALLicType.allicPeriodic) {
			functionReturnValue = "Periodic";
		}
		else if (LicType == ProductLicense.ALLicType.allicPermanent) {
			functionReturnValue = "Permanent";
		}
		else if (LicType == ProductLicense.ALLicType.allicTimeLocked) {
			functionReturnValue = "Timed Expiry";
		}
		// default
		else {
			functionReturnValue = "None";
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function TrimNulls
	// Input:
	//   ByVal str As String - String to be trimmed.
	// Output:
	//   String - Trimmed string.
	// Purpose: Removes Null characters from the string.
	// Remarks: None
	//===============================================================================
	//str was upgraded to str_Renamed
	public string TrimNulls(string str_Renamed)
	{
		return modActiveLock.TrimNulls(ref str_Renamed);
	}
	//===============================================================================
	// Name: Function MD5Hash
	// Input:
	//   ByVal str As String - String to be hashed.
	// Output:
	//   String - Computed hash code.
	// Purpose: Computes an MD5 hash of the specified string.
	// Remarks: None
	//===============================================================================
	//str was upgraded to str_Renamed
	public string MD5Hash(string str_Renamed)
	{
		return modMD5.Hash(ref str_Renamed);
	}
	//===============================================================================
	// Name: Function Base64Encode
	// Input:
	//   ByVal str As String - String to be encoded.
	// Output:
	//   String - Encoded string.
	// Purpose: Encodes a base64-decoded string.
	// Remarks: None
	//===============================================================================
	//str was upgraded to str_Renamed
	public string Base64Encode(string str_Renamed)
	{
		return modBase64.Base64_Encode(ref str_Renamed);
	}
	//===============================================================================
	// Name: Function Base64Decode
	// Input:
	//   ByVal strEncoded As String - String to be decoded.
	// Output:
	//   String - Decoded string.
	// Purpose: Decodes a base64-encoded string.
	// Remarks: None
	//===============================================================================
	public string Base64Decode(string strEncoded)
	{
		return modBase64.Base64_Decode(ref strEncoded);
	}

}
