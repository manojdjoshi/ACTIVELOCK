using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
using System.IO;
using ActiveLock3_6NET;

#region "Copyright"
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
#endregion

/// <summary>
/// This IKeyStoreProvider implementation is used to  maintain the license keys on a file system.
/// </summary>
/// <remarks>Implements IKeyStoreProvider interface.</remarks>
internal class FileKeyStoreProvider : _IKeyStoreProvider
{

	// Started: 21.04.2005
	// Modified: 03.24.2006
	//===============================================================================
	// @author activelock-admins
	// @version 3.3.0
	// @date 03.24.2006

	private string mstrPath;
	private INIFile mINIFile = new INIFile();

	// License File Key names
	private const string KEY_PRODKEY = "ProductKey";
	private const string KEY_PRODNAME = "ProductName";
	private const string KEY_PRODVER = "ProductVersion";
	private const string KEY_LICENSEE = "Licensee";
	private const string KEY_REGISTERED_LEVEL = "RegisteredLevel";
		// Maximum number of users
	private const string KEY_MAXCOUNT = "MaxCount";
	private const string KEY_LICTYPE = "LicenseType";
	private const string KEY_LICCLASS = "LicenseClass";
	private const string KEY_LICKEY = "LicenseKey";
		// New in v3.1
	private const string KEY_LICCODE = "LicenseCode";
	private const string KEY_EXP = "Expiration";
	private const string KEY_REGISTERED_DATE = "RegisteredDate";
		// date and time stamp
	private const string KEY_LASTRUN_DATE = "LastUsed";
		// Hash of LastRunDate
	private const string KEY_LASTRUN_DATE_HASH = "Hash1";

	/// <summary>
	/// Creates an empty file if it doesn't exist
	/// </summary>
	/// <value>String - File path and name</value>
	/// <remarks></remarks>
	private string IKeyStoreProvider_KeyStorePath {
		set {
			if (!modActiveLock.FileExists(value)) {
				// Create an empty file if it doesn't exists
				CreateEmptyFile(value);
			}
			//the file exists, but check to see if it has read-only attribute
			else {
				if ((FileSystem.GetAttr(value) & FileAttribute.ReadOnly) | (FileSystem.GetAttr(value) & FileAttribute.ReadOnly & FileAttribute.Archive)) {
					modActiveLock.Set_Locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid, modTrial.ACTIVELOCKSTRING, modActiveLock.STRLICENSEINVALID);
				}
			}
			mstrPath = value;
			mINIFile.File = mstrPath;
		}
	}
	string _IKeyStoreProvider.KeyStorePath {
		set { IKeyStoreProvider_KeyStorePath = value; }
	}

	/// <summary>
	/// Creates an empty file
	/// </summary>
	/// <param name="sFilePath">String - File path and name</param>
	/// <remarks></remarks>
	private void CreateEmptyFile(string sFilePath)
	{
		int hFile = 0;
		hFile = FreeFile();
		FileSystem.FileOpen(hFile, sFilePath, OpenMode.Output);
		FileSystem.FileClose(hFile);
	}

	// TODO: Perhaps we need to lock the file first.?
	/// <summary>
	/// Write license properties to INI file section
	/// </summary>
	/// <param name="Lic">ProductLicense - Product license object</param>
	/// <param name="mLicenseFileType">IActiveLock.ALLicenseFileTypes - License file type!</param>
	/// <remarks></remarks>
	private void IKeyStoreProvider_Store(ref ProductLicense Lic, IActiveLock.ALLicenseFileTypes mLicenseFileType)
	{
		string[] arrProdVer = null;
		string actualLicensee = null;
		arrProdVer = Lic.Licensee.Split("&&&");
		actualLicensee = arrProdVer[0];

		// Write license properties to INI file section
		mINIFile.Section = Lic.ProductName;

		if (mLicenseFileType == IActiveLock.ALLicenseFileTypes.alsLicenseFileEncrypted) {
			{
				mINIFile.Values(KEY_PRODVER) = modTrial.EncryptString128Bit(Lic.ProductVer, PSWD());
				mINIFile.Values(KEY_LICTYPE) = modTrial.EncryptString128Bit(Lic.LicenseType, PSWD());
				mINIFile.Values(KEY_LICCLASS) = modTrial.EncryptString128Bit(Lic.LicenseClass, PSWD());
				mINIFile.Values(KEY_LICENSEE) = modTrial.EncryptString128Bit(actualLicensee, PSWD());
				mINIFile.Values(KEY_REGISTERED_LEVEL) = modTrial.EncryptString128Bit(Lic.RegisteredLevel, PSWD());
				mINIFile.Values(KEY_MAXCOUNT) = modTrial.EncryptString128Bit((string)Lic.MaxCount, PSWD());
				mINIFile.Values(KEY_LICKEY) = modTrial.EncryptString128Bit(Lic.LicenseKey, PSWD());
				mINIFile.Values(KEY_REGISTERED_DATE) = modTrial.EncryptString128Bit(Lic.RegisteredDate, PSWD());
				mINIFile.Values(KEY_LASTRUN_DATE) = modTrial.EncryptString128Bit(Lic.LastUsed, PSWD());
				mINIFile.Values(KEY_LASTRUN_DATE_HASH) = modTrial.EncryptString128Bit(Lic.Hash1, PSWD());
				mINIFile.Values(KEY_EXP) = modTrial.EncryptString128Bit(Lic.Expiration, PSWD());
				mINIFile.Values(KEY_LICCODE) = modTrial.EncryptString128Bit(Lic.LicenseCode, PSWD());
			}
		}
		else {
			{
				mINIFile.Values(KEY_PRODVER) = Lic.ProductVer;
				mINIFile.Values(KEY_LICTYPE) = Lic.LicenseType;
				mINIFile.Values(KEY_LICCLASS) = Lic.LicenseClass;
				mINIFile.Values(KEY_LICENSEE) = actualLicensee;
				mINIFile.Values(KEY_REGISTERED_LEVEL) = Lic.RegisteredLevel;
				mINIFile.Values(KEY_MAXCOUNT) = (string)Lic.MaxCount;
				mINIFile.Values(KEY_LICKEY) = Lic.LicenseKey;
				mINIFile.Values(KEY_REGISTERED_DATE) = Lic.RegisteredDate;
				mINIFile.Values(KEY_LASTRUN_DATE) = Lic.LastUsed;
				mINIFile.Values(KEY_LASTRUN_DATE_HASH) = Lic.Hash1;
				mINIFile.Values(KEY_EXP) = Lic.Expiration;
				mINIFile.Values(KEY_LICCODE) = Lic.LicenseCode;
			}

		}

		// This was the original idea... to read the file, and encrypt and write again
		// But since this is a stream based operation, it conflicts with the
		// ADS therefore should NOT be used
		//If mLicenseFileType = IActiveLock.ALLicenseFileTypes.alsLicenseFileEncrypted Then
		//    ' Read the LIC file again
		//    Dim stream_reader2 As New IO.StreamReader(mINIFile.File)
		//    Dim fileContents2 As String = stream_reader2.ReadToEnd()
		//    stream_reader2.Close()
		//    ' Encrypt the file and save it
		//    Dim encryptedFile2 As String
		//    encryptedFile2 = EncryptString128Bit(fileContents2, PSWD)
		//    Dim stream_writer2 As New IO.StreamWriter(mINIFile.File, False)
		//    stream_writer2.Write(encryptedFile2)
		//    stream_writer2.Close()
		//End If

	}
	void _IKeyStoreProvider.Store(ref ProductLicense Lic, IActiveLock.ALLicenseFileTypes mLicenseFileType)
	{
		IKeyStoreProvider_Store(Lic, mLicenseFileType);
	}

	/// <summary>
	/// Retrieves the registered license for the specified product.
	/// </summary>
	/// <param name="ProductName">String - Product or application name</param>
	/// <param name="mLicenseFileType">IActiveLock.ALLicenseFileTypes - License file type!</param>
	/// <returns>ProductLicense - Returns the product license object</returns>
	/// <remarks></remarks>

	private ProductLicense IKeyStoreProvider_Retrieve(ref string ProductName, IActiveLock.ALLicenseFileTypes mLicenseFileType)
	{
		ProductLicense functionReturnValue = null;
		functionReturnValue = null;

		mINIFile.Section = ProductName;

		 // ERROR: Not supported in C#: OnErrorStatement

		// No license found
		if (string.IsNullOrEmpty(mINIFile.GetValue(KEY_LICKEY))) return;
 

		// Read the LIC file quickly and decide whether this is a valid LIC file
		string fileText = ActiveLock3_6NET.My.MyProject.Computer.FileSystem.ReadAllText(mINIFile.File);
		string astring = string.Empty;
		if (string.IsNullOrEmpty(fileText)) return;
 
		if (mLicenseFileType == IActiveLock.ALLicenseFileTypes.alsLicenseFileEncrypted) {
			astring = modTrial.DecryptString128Bit(fileText, PSWD());
		}
		else {
			astring = fileText;
		}
		if (astring.ToLower().Contains("productversion") == false) {
			modActiveLock.Set_locale(modActiveLock.regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrKeyStoreInvalid, modTrial.ACTIVELOCKSTRING, modActiveLock.STRKEYSTOREINVALID);
		}

		ProductLicense Lic = new ProductLicense();

		if (mLicenseFileType == IActiveLock.ALLicenseFileTypes.alsLicenseFileEncrypted & mINIFile.GetValue(KEY_LICCLASS) != "Single" & mINIFile.GetValue(KEY_LICCLASS) != "MultiUser") {
			// This was the original idea... to read the file, and encrypt and write again
			// But since this is a stream based operation, it conflicts with the
			// ADS therefore should NOT be used
			//' Read the LIC file
			//Dim stream_reader1 As New IO.StreamReader(mINIFile.File)
			//Dim fileContents1 As String = stream_reader1.ReadToEnd()
			//stream_reader1.Close()
			//' Decrypt the LIC file
			//Dim decryptedFile1 As String
			//decryptedFile1 = DecryptString128Bit(fileContents1, PSWD)
			//' Save the LIC file again
			//Dim stream_writer1 As New IO.StreamWriter(mINIFile.File, False)
			//stream_writer1.Write(decryptedFile1)
			//stream_writer1.Close()

			// Read license properties from INI file section
			{
				Lic.ProductName = ProductName;
				Lic.ProductVer = modTrial.DecryptString128Bit(mINIFile.GetValue(KEY_PRODVER), PSWD());
				Lic.Licensee = modTrial.DecryptString128Bit(mINIFile.GetValue(KEY_LICENSEE), PSWD()) + "&&&" + Lic.ProductName + " (" + Lic.ProductVer + ")";
				Lic.RegisteredLevel = modTrial.DecryptString128Bit(mINIFile.GetValue(KEY_REGISTERED_LEVEL), PSWD());
				Lic.MaxCount = (int)modTrial.DecryptString128Bit(mINIFile.Values(KEY_MAXCOUNT), PSWD());
				Lic.LicenseType = modTrial.DecryptString128Bit(mINIFile.GetValue(KEY_LICTYPE), PSWD());
				Lic.LicenseClass = modTrial.DecryptString128Bit(mINIFile.GetValue(KEY_LICCLASS), PSWD());
				Lic.LicenseKey = modTrial.DecryptString128Bit(mINIFile.GetValue(KEY_LICKEY), PSWD());
				Lic.Expiration = modTrial.DecryptString128Bit(mINIFile.GetValue(KEY_EXP), PSWD());
				Lic.RegisteredDate = modTrial.DecryptString128Bit(mINIFile.Values(KEY_REGISTERED_DATE), PSWD());
				Lic.LastUsed = modTrial.DecryptString128Bit(mINIFile.Values(KEY_LASTRUN_DATE), PSWD());
				Lic.Hash1 = modTrial.DecryptString128Bit(mINIFile.Values(KEY_LASTRUN_DATE_HASH), PSWD());
				Lic.LicenseCode = modTrial.DecryptString128Bit(mINIFile.Values(KEY_LICCODE), PSWD());
			}
		}
		// INI file (LIC) is in PLAIN format
		else {

			// Read license properties from INI file section
			{
				Lic.ProductName = ProductName;
				Lic.ProductVer = mINIFile.GetValue(KEY_PRODVER);
				Lic.Licensee = mINIFile.GetValue(KEY_LICENSEE) + "&&&" + Lic.ProductName + " (" + Lic.ProductVer + ")";
				Lic.RegisteredLevel = mINIFile.GetValue(KEY_REGISTERED_LEVEL);
				Lic.MaxCount = (int)mINIFile.Values(KEY_MAXCOUNT);
				Lic.LicenseType = mINIFile.GetValue(KEY_LICTYPE);
				Lic.LicenseClass = mINIFile.GetValue(KEY_LICCLASS);
				Lic.LicenseKey = mINIFile.GetValue(KEY_LICKEY);
				Lic.Expiration = mINIFile.GetValue(KEY_EXP);
				Lic.RegisteredDate = mINIFile.Values(KEY_REGISTERED_DATE);
				Lic.LastUsed = mINIFile.Values(KEY_LASTRUN_DATE);
				Lic.Hash1 = mINIFile.Values(KEY_LASTRUN_DATE_HASH);
				Lic.LicenseCode = mINIFile.Values(KEY_LICCODE);
				// New in v3.1
			}
		}
		functionReturnValue = Lic;

		return;
		InvalidValue:
		modActiveLock.Set_locale(modActiveLock.regionalSymbol);
		Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrKeyStoreInvalid, modTrial.ACTIVELOCKSTRING, modActiveLock.STRKEYSTOREINVALID);
		return functionReturnValue;
	}
	ProductLicense _IKeyStoreProvider.Retrieve(ref string ProductName, IActiveLock.ALLicenseFileTypes mLicenseFileType)
	{
		return IKeyStoreProvider_Retrieve(ProductName, mLicenseFileType);
	}
}
