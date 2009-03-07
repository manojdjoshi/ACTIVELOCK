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
using System.Security.Cryptography;
using System.Management;
//using System.TimeSpan;
using System.Text;
// For StringBuilder
using System.Runtime.InteropServices;
// For DLL Call

#region "Copyright"
//*   ActiveLock
//*   Copyright 1998-2002 Nelson Ferraz
//*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
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

namespace ActiveLock3_6NET
{

	/// <summary>
	/// <para>This is an implementation of IActiveLock.</para>
	/// <para>It is not public-creatable, and so must only be accessed via ActiveLock.NewInstance() method.</para>
	/// <para>Includes Key generation and validation routines.</para>
	/// </summary>
	/// <remarks>If you want to turn off dll-checksumming, add this compilation flag to the Project Properties (Make tab) AL_DEBUG = 1</remarks>
	internal class ActiveLock:_IActiveLock
	{
		// Started: 21.04.2005
		// Modified: 03.235.2006
		//===============================================================================

		// @author: activelock-admins
		// @version: 3.3.0
		// @date: 03.23.2006

		// Implements the IActiveLock interface.

		Microsoft.VisualBasic.ErrObject Err;
		object blank = null;

		private string mSoftwareName;
		private string mSoftwareVer;
		private string mSoftwarePassword;
		private string mSoftwareCode;
		private string mRegisteredLevel;
		private IActiveLock.ALLockTypes mLockTypes;
		private IActiveLock.ALLicenseKeyTypes mLicenseKeyTypes;
		private IActiveLock.ALLockTypes[] mUsedLockTypes;
		private int mTrialType;
		private int mTrialLength;
		private int mRemainingTrialDays;
		private int mRemainingTrialRuns;
		private IActiveLock.ALTrialHideTypes mTrialHideTypes;
		private _IKeyStoreProvider mKeyStore;
		private string mKeyStorePath;
		private ActiveLockEventNotifier MyNotifier = new ActiveLockEventNotifier();
		private Globals MyGlobals = new Globals();
		private string mLibKeyPath;
		private IActiveLock.ALTimeServerTypes mCheckTimeServerForClockTampering;
		private IActiveLock.ALSystemFilesTypes mChecksystemfilesForClockTampering;
		private IActiveLock.ALLicenseFileTypes mLicenseFileType;
		private IActiveLock.ALAutoRegisterTypes mAutoRegister;
		private IActiveLock.ALTrialWarningTypes mTrialWarning;
		private int mUsedLockType;

		private bool dontValidateLicense;
		/// <summary>
		/// Registry hive used to store Activelock settings.
		/// </summary>
		/// <remarks></remarks>

		private const string AL_REGISTRY_HIVE = "Software\\ActiveLock\\ActiveLock3";
		// Transients

		/// <summary>
		/// flag to indicate that ActiveLock has been initialized
		/// </summary>
		/// <remarks></remarks>
		// flag to indicate that ActiveLock has been initialized
		private bool mfInit;
		[DllImport("kernel32",EntryPoint="GetVolumeInformationA",CharSet=CharSet.Ansi,SetLastError=true,ExactSpelling=true)]
		public static extern int GetVolumeInformation(string lpRootPathName,StringBuilder lpVolumeNameBuffer,int nVolumeNameSize,int lpVolumeSerialNumber,int lpMaximumComponentLength,int lpFileSystemFlags,StringBuilder lpFileSystemNameBuffer,int nFileSystemNameSize);

		/// <summary>
		/// <para>GetVolumeInformation</para>
		/// </summary>
		/// <param name="lpRootPathName">String - A pointer to a string that contains the root directory of the volume to be described.</param>
		/// <param name="lpVolumeNameBuffer">A pointer to a buffer that receives the name of a specified volume. The maximum buffer size is MAX_PATH+1.</param>
		/// <param name="nVolumeNameSize">The length of a volume name buffer, in TCHARs. The maximum buffer size is MAX_PATH+1.</param>
		/// <param name="lpVolumeSerialNumber">A pointer to a variable that receives the volume serial number.</param>
		/// <param name="lpMaximumComponentLength">A pointer to a variable that receives the maximum length, in TCHARs, of a file name component that a specified file system supports.</param>
		/// <param name="lpFileSystemFlags">A pointer to a variable that receives flags associated with the specified file system.</param>
		/// <param name="lpFileSystemNameBuffer">A pointer to a buffer that receives the name of the file system, for example, the FAT file system or the NTFS file system. The maximum buffer size is MAX_PATH+1.</param>
		/// <param name="nFileSystemNameSize">The length of the file system name buffer, in TCHARs. The maximum buffer size is MAX_PATH+1.</param>
		/// <returns>
		/// <para>If all the requested information is retrieved, the return value is nonzero.</para>
		/// <para>If not all the requested information is retrieved, the return value is zero (0). To get extended error information, call GetLastError.</para>
		/// </returns>
		/// <remarks>
		/// <para>See <a href="http://msdn.microsoft.com/en-us/library/aa364993(VS.85).aspx">http://msdn.microsoft.com/en-us/library/aa364993(VS.85).aspx</a></para>
		/// </remarks>

		/// <summary>
		/// IActiveLock_LicenseKeyType - Specifies the ALLicenseKeyTypes type
		/// </summary>
		/// <value>ByVal RHS As ALLicenseKeyTypes - ALLicenseKeyTypes type</value>
		/// <remarks>None</remarks>
		private IActiveLock.ALLicenseKeyTypes IActiveLock_LicenseKeyType
		{
			set { mLicenseKeyTypes = value; }
		}
		IActiveLock.ALLicenseKeyTypes _IActiveLock.LicenseKeyType
		{
			set { IActiveLock_LicenseKeyType = value; }
		}

		/// <summary>
		/// Gets the Registered Level for the license after validating it.
		/// </summary>
		/// <value></value>
		/// <returns>String - License RegisteredLevel</returns>
		/// <remarks>None</remarks>
		private string IActiveLock_RegisteredLevel
		{
			get
			{
				ProductLicense Lic = null;
				Lic = mKeyStore.Retrieve(ref mSoftwareName,mLicenseFileType);
				if(Lic == null)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					
	// Moified th following line...
				 	Err.Raise(Convert.ToInt32(Globals.ActiveLockErrCodeConstants.AlerrNoLicense),modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOLICENSE, blank, blank);
				}
				// Validate the License.
				ValidateLic(ref Lic);
				return Lic.RegisteredLevel;
			}
		}
		string _IActiveLock.RegisteredLevel
		{
			get { return IActiveLock_RegisteredLevel; }
		}

		/// <summary>
		/// Gets the LicenseClass
		/// </summary>
		/// <value></value>
		/// <returns>String - LicenseClass</returns>
		/// <remarks></remarks>
		private string IActiveLock_LicenseClass
		{
			get
			{
				ProductLicense Lic = null;
				Lic = mKeyStore.Retrieve(ref mSoftwareName,mLicenseFileType);
				if(Lic == null)
				{
					
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err.Raise(Convert.ToInt32(Globals.ActiveLockErrCodeConstants.AlerrNoLicense),modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOLICENSE, blank, blank);
				}
				// Validate the License.
				ValidateLic(ref Lic);
				return Lic.LicenseClass;
			}
		}
		string _IActiveLock.LicenseClass
		{
			get { return IActiveLock_LicenseClass; }
		}

		/// <summary>
		/// Gets the Number of Used Trial Days
		/// </summary>
		/// <value></value>
		/// <returns>Integer - License Used Trial Days</returns>
		/// <remarks></remarks>
		private int IActiveLock_RemainingTrialDays
		{
			get { return mRemainingTrialDays; }
		}
		int _IActiveLock.RemainingTrialDays
		{
			get { return IActiveLock_RemainingTrialDays; }
		}

		/// <summary>
		/// Gets the Number of Used Trial Runs
		/// </summary>
		/// <value></value>
		/// <returns>Integer - License Used Trial Runs</returns>
		/// <remarks></remarks>
		private int IActiveLock_RemainingTrialRuns
		{
			get { return mRemainingTrialRuns; }
		}
		int _IActiveLock.RemainingTrialRuns
		{
			get { return IActiveLock_RemainingTrialRuns; }
		}

		/// <summary>
		/// Gets the Number of concurrent users for the networked license
		/// </summary>
		/// <value></value>
		/// <returns>Integer - Number of concurrent users for the networked license</returns>
		/// <remarks></remarks>
		private int IActiveLock_MaxCount
		{
			get
			{
				ProductLicense Lic = null;
				Lic = mKeyStore.Retrieve(ref mSoftwareName,mLicenseFileType);
				if(Lic == null)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err.Raise(Convert.ToInt32(Globals.ActiveLockErrCodeConstants.AlerrNoLicense),modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOLICENSE, blank, blank);
				}
				// Validate the License.
				ValidateLic(ref Lic);
				return Lic.MaxCount;
			}
		}
		int _IActiveLock.MaxCount
		{
			get { return IActiveLock_MaxCount; }
		}

		/// <summary>
		/// <para>IActiveLock Interface implementation</para>
		/// <para>Specifies the liberation key auto file path name</para>
		/// </summary>
		/// <value>ByVal RHS As String - Liberation key file auto path name</value>
		/// <remarks></remarks>
		private string IActiveLock_AutoRegisterKeyPath
		{
			set { mLibKeyPath = value; }
		}
		string _IActiveLock.AutoRegisterKeyPath
		{
			set { IActiveLock_AutoRegisterKeyPath = value; }
		}

		// TODO: ActiveLock.vb - Property AutoRegisterKeyPath - I think this should read Gets, not Sets: japreja
		/// <summary>
		/// Sets the auto register file full path
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
		private string AutoRegisterKeyPath
		{
			get { return mLibKeyPath; }
		}

		/// <summary>
		/// Gets a notification from Activelock
		/// </summary>
		/// <value></value>
		/// <returns>ActiveLockEventNotifier - ???</returns>
		/// <remarks></remarks>
		private ActiveLockEventNotifier IActiveLock_EventNotifier
		{
			get { return MyNotifier; }
		}
		ActiveLockEventNotifier _IActiveLock.EventNotifier
		{
			get { return IActiveLock_EventNotifier; }
		}

		/// <summary>
		/// Gets the license registration date after validating it.
		/// </summary>
		/// <value></value>
		/// <returns>String - License registration date.</returns>
		/// <remarks>This is the date the license was generated by Alugen. NOT the date the license was activated.</remarks>
		private string IActiveLock_RegisteredDate
		{
			get
			{
				ProductLicense Lic = null;
				Lic = mKeyStore.Retrieve(ref mSoftwareName,mLicenseFileType);
				if(Lic == null)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err.Raise(Convert.ToInt32(Globals.ActiveLockErrCodeConstants.AlerrNoLicense),modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOLICENSE, blank, blank);
				}
				// Validate the License.
				ValidateLic(ref Lic);
				return Lic.RegisteredDate;
			}
		}
		string _IActiveLock.RegisteredDate
		{
			get { return IActiveLock_RegisteredDate; }
		}

		/// <summary>
		/// Gets the registered user name after validating the license
		/// </summary>
		/// <value></value>
		/// <returns>String - Registered user name</returns>
		/// <remarks></remarks>
		private string IActiveLock_RegisteredUser
		{
			get
			{
				ProductLicense Lic = null;
				Lic = mKeyStore.Retrieve(ref mSoftwareName,mLicenseFileType);
				if(Lic == null)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err.Raise(Convert.ToInt32(Globals.ActiveLockErrCodeConstants.AlerrNoLicense),modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOLICENSE, blank, blank);
				}
				// Validate the License.
				ValidateLic(ref Lic);
				return Lic.Licensee;
			}
		}
		string _IActiveLock.RegisteredUser
		{
			get { return IActiveLock_RegisteredUser; }
		}

		/// <summary>
		/// Returns the expiration date of the license after validating it
		/// </summary>
		/// <value></value>
		/// <returns>String - Expiration date of the license</returns>
		/// <remarks></remarks>
		private string IActiveLock_ExpirationDate
		{
			get
			{
				ProductLicense Lic = null;
				Lic = mKeyStore.Retrieve(ref mSoftwareName,mLicenseFileType);
				if(Lic == null)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err.Raise(Convert.ToInt32(Globals.ActiveLockErrCodeConstants.AlerrNoLicense),modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOLICENSE, blank, blank);
				}
				// Validate the License.
				ValidateLic(ref Lic);
				return Lic.Expiration;
			}
		}
		string _IActiveLock.ExpirationDate
		{
			get { return IActiveLock_ExpirationDate; }
		}

		/// <summary>
		/// Specifies the license file path name
		/// </summary>
		/// <value>ByVal RHS As String - License file path name</value>
		/// <remarks></remarks>
		private string IActiveLock_KeyStorePath
		{
			set
			{
				if((mKeyStore != null))
				{
					mKeyStore.KeyStorePath = value;
				}
				mKeyStorePath = value;
			}
		}
		string _IActiveLock.KeyStorePath
		{
			set { IActiveLock_KeyStorePath = value; }
		}

		/// <summary>
		/// <para>Specifies the key store type</para>
		/// <para>This version of Activelock does not work with the registry</para>
		/// </summary>
		/// <value>ByVal RHS As LicStoreType - License store type</value>
		/// <remarks>Portions of this (RegistryKeyStoreProvider) not implemented yet</remarks>
		private IActiveLock.LicStoreType IActiveLock_KeyStoreType
		{
			set
			{
				// Instantiate Key Store Provider
				if(value == IActiveLock.LicStoreType.alsFile)
				{
					mKeyStore = new FileKeyStoreProvider();
				}
				else
				{
					// Set mKeyStore = New RegistryKeyStoreProvider
					// TODO: ActiveLock.vb - Property IActiveLock_KeyStoreType - Implement me!
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err.Raise(Convert.ToInt32(Globals.ActiveLockErrCodeConstants.AlerrNotImplemented),modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOTIMPLEMENTED, blank,blank);
				}
				// Set Key Store Path in KeyStoreProvider
				if(!string.IsNullOrEmpty(mKeyStorePath))
				{
					mKeyStore.KeyStorePath = mKeyStorePath;
				}
			}
		}
		IActiveLock.LicStoreType _IActiveLock.KeyStoreType
		{
			set { IActiveLock_KeyStoreType = value; }
		}

		/// <summary>
		/// Gets or Sets the ALLockTypes type
		/// </summary>
		/// <value>ByVal RHS As ALLockTypes - ALLockTypes type</value>
		/// <returns>ALLockTypes - Lock types type</returns>
		/// <remarks></remarks>
		private IActiveLock.ALLockTypes IActiveLock_LockType
		{
			get { return mLockTypes; }
			set { mLockTypes = value; }
		}
		IActiveLock.ALLockTypes _IActiveLock.LockType
		{
			get { return IActiveLock_LockType; }
			set { IActiveLock_LockType = value; }
		}

		/// <summary>
		/// Helper function to build up array of used LockType s
		/// </summary>
		/// <param name="LockType"><para>ByVal LockType As ALLockTypes _ to be added to array.</para><para>ByRef Byref LockTypes() As ALLockTypes - array of used LockTypes being built up.</para></param>
		/// <param name="SizeLT">ByRef SizeLT as Integer - size of array of used LockTypes being built up</param>
		/// <remarks></remarks>
		private void IActiveLock_AddLockCode(IActiveLock.ALLockTypes LockType,ref int SizeLT)
		{
			Array.Resize(ref mUsedLockTypes,SizeLT + 1);
			mUsedLockTypes[SizeLT] = LockType;
			SizeLT = SizeLT + 1;
		}

		/// <summary>
		/// Gets the ALTrialHideTypes type
		/// </summary>
		/// <value></value>
		/// <returns>ALTrialHideTypes - Trial Hide types type</returns>
		/// <remarks></remarks>
		private IActiveLock.ALTrialHideTypes IActiveLock_TrialHideType
		{
			get { return mTrialHideTypes; }
			set { mTrialHideTypes = value; }
		}
		IActiveLock.ALTrialHideTypes _IActiveLock.TrialHideType
		{
			get { return IActiveLock_TrialHideType; }
			set { IActiveLock_TrialHideType = value; }
		}

		/// <summary>
		/// Gets the SoftwareName for the license
		/// </summary>
		/// <value></value>
		/// <returns>String - Software name  for the license</returns>
		/// <remarks></remarks>
		private string IActiveLock_SoftwareName
		{
			get { return mSoftwareName; }
			set { mSoftwareName = value; }
		}
		string _IActiveLock.SoftwareName
		{
			get { return IActiveLock_SoftwareName; }
			set { IActiveLock_SoftwareName = value; }
		}

		/// <summary>
		/// Gets/Sets the SoftwarePassword for the license
		/// </summary>
		/// <value>ByVal RHS As String - Software Password for the license</value>
		/// <returns>String - Software Password for the license</returns>
		/// <remarks></remarks>
		private string IActiveLock_SoftwarePassword
		{
			get { return mSoftwarePassword; }
			set { mSoftwarePassword = value; }
		}
		string _IActiveLock.SoftwarePassword
		{
			get { return IActiveLock_SoftwarePassword; }
			set { IActiveLock_SoftwarePassword = value; }
		}

		/// <summary>
		/// Specifies whether a Time Server should be used to check Clock Tampering
		/// </summary>
		/// <value>ByVal iServer As Integer - Flag being passed to check the time server</value>
		/// <remarks></remarks>
		private IActiveLock.ALTimeServerTypes IActiveLock_CheckTimeServerForClockTampering
		{
			set { mCheckTimeServerForClockTampering = value; }
		}
		IActiveLock.ALTimeServerTypes _IActiveLock.CheckTimeServerForClockTampering
		{
			set { IActiveLock_CheckTimeServerForClockTampering = value; }
		}

		/// <summary>
		/// Specifies whether a Time Server should be used to check Clock Tampering
		/// </summary>
		/// <value>ByVal iServer As Integer - Flag being passed to check the time server</value>
		/// <remarks></remarks>
		private IActiveLock.ALSystemFilesTypes IActiveLock_CheckSystemFilesForClockTampering
		{
			set { mChecksystemfilesForClockTampering = value; }
		}
		IActiveLock.ALSystemFilesTypes _IActiveLock.CheckSystemFilesForClockTampering
		{
			set { IActiveLock_CheckSystemFilesForClockTampering = value; }
		}

		// TODO: ActiveLock.vb - Property IActiveLock_LicenseFileType - Update return value comment!
		/// <summary>
		/// Specifies whether the License File should be encrypted or not
		/// </summary>
		/// <value>ByVal Value As IActiveLock.ALLicenseFileTypes - Flag to indicate the license file will be encrypted or not</value>
		/// <returns></returns>
		/// <remarks></remarks>
		private IActiveLock.ALLicenseFileTypes IActiveLock_LicenseFileType
		{
			get { return mLicenseFileType; }
			set { mLicenseFileType = value; }
		}
		IActiveLock.ALLicenseFileTypes _IActiveLock.LicenseFileType
		{
			get { return IActiveLock_LicenseFileType; }
			set { IActiveLock_LicenseFileType = value; }
		}

		// TODO: ActiveLock.vb - Property IActiveLock_AutoRegister - Update Comment - Not Documented!
		/// <summary>
		/// Not Documented!
		/// </summary>
		/// <value>ALAutoRegisterTypes - ALAutoRegisterType</value>
		/// <remarks></remarks>
		private IActiveLock.ALAutoRegisterTypes IActiveLock_AutoRegister
		{
			set { mAutoRegister = value; }
		}
		IActiveLock.ALAutoRegisterTypes _IActiveLock.AutoRegister
		{
			set { IActiveLock_AutoRegister = value; }
		}

		/// <summary>
		/// Specifies whether the License File should be encrypted or not
		/// </summary>
		/// <value>ByVal Value As IActiveLock.ALTrialWarningTypes - Flag to indicate the license file will be encrypted or not.</value>
		/// <remarks></remarks>
		private IActiveLock.ALTrialWarningTypes IActiveLock_TrialWarning
		{
			set { mTrialWarning = value; }
		}
		IActiveLock.ALTrialWarningTypes _IActiveLock.TrialWarning
		{
			set { IActiveLock_TrialWarning = value; }
		}

		/// <summary>
		/// Gets/Sets the TrialType for the license
		/// </summary>
		/// <value>ByVal Value As IActiveLock.ALTrialTypes</value>
		/// <returns>ALTrialTypes - Trial Type  for the license</returns>
		/// <remarks></remarks>
		private IActiveLock.ALTrialTypes IActiveLock_TrialType
		{
			get { return (IActiveLock.ALTrialTypes)mTrialType; }
			set { mTrialType = (int)value; }
		}
		IActiveLock.ALTrialTypes _IActiveLock.TrialType
		{
			get { return IActiveLock_TrialType; }
			set { IActiveLock_TrialType = value; }
		}

		/// <summary>
		/// Gets/Sets the TrialLength for the license
		/// </summary>
		/// <value></value>
		/// <returns>Integer - Trial Length  for the license</returns>
		/// <remarks></remarks>
		private int IActiveLock_TrialLength
		{
			get { return mTrialLength; }
			set { mTrialLength = value; }
		}
		int _IActiveLock.TrialLength
		{
			get { return IActiveLock_TrialLength; }
			set { IActiveLock_TrialLength = value; }
		}

		/// <summary>
		/// Combines the user name with the lock code and returns it as the installation code
		/// </summary>
		/// <param name="User">Optional - String - User name</param>
		/// <param name="Lic">Optional - ProductLicense - Product License</param>
		/// <value></value>
		/// <returns>String - Installation code</returns>
		/// <remarks></remarks>
		private string IActiveLock_InstallationCode
		{
			get
			{
				//Before we generate the installation code, let's check if this app is using a short key
				string strReq = null;
				string strLock = null;
				string strReq2 = null;
				if(mLicenseKeyTypes == IActiveLock.ALLicenseKeyTypes.alsShortKeyMD5)
				{
					// Shortkeys are no longer using the HDD firmware serial number
					// they are using the Computer Fingerprint after v3.6
					return IActivelock_GenerateShortSerial(modHardware.GetFingerprint());

				}

				else if(mLicenseKeyTypes == IActiveLock.ALLicenseKeyTypes.alsRSA)
				{
					// Generate Request code to Lock

					//Restrict user name to 2000 characters; need more? why?
					if(Strings.Len(User) > 2000)
					{
						modActiveLock.Set_locale(modActiveLock.regionalSymbol);
						Err.Raise(Convert.ToInt32(Globals.ActiveLockErrCodeConstants.AlerrUserNameTooLong),modTrial.ACTIVELOCKSTRING,modActiveLock.STRUSERNAMETOOLONG, blank,blank);
					}

					// New in v3.1
					// Version 3.1 and above of Activelock will append the "+" sign
					// in front of the installation code whenever lockNone is used or
					// lockType is not specified in the protected app.
					// When "+" is not found at the beginning of the installation code,
					// Alugen will not allow users pick the hardware lock method since this
					// corresponds to an installation code which
					// utilizes a hardware lock option specified inside the protected app.
					if(mLockTypes == IActiveLock.ALLockTypes.lockNone)
					{
						strLock = "+" + IActiveLock_LockCode();
					}
					else
					{
						strLock = IActiveLock_LockCode();
					}

					// combine with user name
					strReq = strLock + Constants.vbLf + User;

					// combine with app name and version
					strReq = strReq + "&&&" + IActiveLock_SoftwareName + " (" + IActiveLock_SoftwareVersion + ")";

					// base-64 encode the request
					strReq2 = modBase64.Base64_Encode(ref strReq);
					return strReq2;

					// New in v3.1
					// If there's a license and the LicenseCode exists, then use it
					// LicenseCode is actually the Installation Code modified by Alugen
					// LicenseCode is appended to the end of the lic file so that we can know
					// Alugen specified the hardware keys, and LockType
					// was not specified inside the protected app
					if((Lic != null))
					{
						if(!string.IsNullOrEmpty(Lic.LicenseCode))
						{
							return Lic.LicenseCode;
							if(Strings.Left(IActiveLock_InstallationCode,1) == "+") return Strings.Mid(IActiveLock_InstallationCode,2);
							// We won't do the following in order to maintain backwards compatibility with existing licenses
							// ElseIf Lic.LicenseCode = "" And mLockTypes = lockNone Then
							// Set_locale(regionalSymbol)
							// Err.Raise ActiveLockErrCodeConstants.AlerrLicenseInvalid, ACTIVELOCKSTRING, STRLICENSEINVALID
						}
					}
				}
				return null;
			}
		}
		string _IActiveLock.InstallationCode
		{
			get { return IActiveLock_InstallationCode; }
		}


		/// <summary>
		/// Gets the SoftwareVersion for the license
		/// </summary>
		/// <value></value>
		/// <returns>String - Software version  for the license</returns>
		/// <remarks></remarks>
		private string IActiveLock_SoftwareVersion
		{
			get { return mSoftwareVer; }
			set { mSoftwareVer = value; }
		}
		string _IActiveLock.SoftwareVersion
		{
			get { return IActiveLock_SoftwareVersion; }
			set { IActiveLock_SoftwareVersion = value; }
		}

		/// <summary>
		/// Specifies the SoftwareCode for the license
		/// </summary>
		/// <value>ByVal RHS As String - Software code for the license</value>
		/// <remarks>SoftwareCode is an RSA public key.  This code will be used to verify license keys later on.</remarks>
		private string IActiveLock_SoftwareCode
		{
			set { mSoftwareCode = value; }
		}
		string _IActiveLock.SoftwareCode
		{
			set { IActiveLock_SoftwareCode = value; }
		}

		// TODO: ActiveLock.vb - Property IActiveLock_UsedDays - returns comment not documented!
		/// <summary>
		/// Gets the number of days the license was used after validating it.
		/// </summary>
		/// <value></value>
		/// <returns>Integer - ?</returns>
		/// <remarks></remarks>
		public int IActiveLock_UsedDays
		{
			get
			{
				ProductLicense Lic = null;
				Lic = mKeyStore.Retrieve(ref mSoftwareName,mLicenseFileType);
				if(Lic == null)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOLICENSE);
				}

				// validate the license
				if(dontValidateLicense == false) ValidateLic(ref Lic);

				//IActiveLock_UsedDays = CInt(DateDiff("d", Lic.RegisteredDate, Now.UtcNow))
				DateTime mydate = modTrial.ActiveLockDate(System.DateTime.UtcNow);
				IActiveLock_UsedDays = mydate.Subtract(modTrial.ActiveLockDate((System.DateTime)Lic.RegisteredDate)).Days;
				//CInt(DateDiff("d", CDate(Replace(Lic.RegisteredDate, ".", "-")), Date.UtcNow))
				if(IActiveLock_UsedDays < 0)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
				}
			}
		}
		int _IActiveLock.UsedDays
		{
			get { return IActiveLock_UsedDays; }
		}

		/// <summary>
		/// Gets the Lock Type selected in Alugen.
		/// </summary>
		/// <value></value>
		/// <returns></returns>
		/// <remarks></remarks>
		public int IActiveLock_UsedLockType
		{
			get
			{
				ProductLicense Lic = null;
				Lic = mKeyStore.Retrieve(ref mSoftwareName,mLicenseFileType);

				if(Lic == null)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOLICENSE);
				}

				// validate the license
				if(dontValidateLicense == false) ValidateLic(ref Lic);

				if(Lic.LicenseCode == "Short Key")
				{
					IActiveLock_UsedLockType = 0;
				}
				else
				{
					string usedcode = null;
					if(Strings.Left(Lic.LicenseCode,1) == "+")
					{
						usedcode = modBase64.Base64_Decode(ref Strings.Mid(Lic.LicenseCode,2));
					}
					else
					{
						usedcode = modBase64.Base64_Decode(ref (Lic.LicenseCode));
					}
					int Index = 0;
					int i = 0;
					Index = 0;
					i = 1;
					// Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
					while(i > 0)
					{
						i = Strings.InStr(Index + 1,usedcode,Constants.vbLf);
						if(i > 0) Index = i;
					}

					if(Index <= 0) return;

					// lockcode is from beginning to Index-1
					usedcode = Strings.Left(usedcode,Index - 1);
					if(Strings.Left(usedcode,1) == "+")
					{
						usedcode = Strings.Right(usedcode,usedcode.Length - 1);
					}
					string[] myarray = null;
					int counter = 0;
					myarray = usedcode.Split(Constants.vbLf);
					for(i = 0;i <= Information.UBound(myarray);i++)
					{
						if(myarray[i] != "nokey")
						{
							counter = counter + Math.Pow(2,i);
						}
					}
					IActiveLock_UsedLockType = counter;
				}
			}
		}
		int _IActiveLock.UsedLockType
		{
			get { return IActiveLock_UsedLockType; }
		}

		// TODO: ActiveLock.vb - Sub Class_Initialize_Renamed - Add documentation!
		/// <summary>
		/// Not documented!
		/// </summary>
		/// <remarks>Class_Initialize was upgraded to Class_Initialize_Renamed</remarks>
		private void Class_Initialize_Renamed()
		{
			// Default to alsFile
			IActiveLock_KeyStoreType = IActiveLock.LicStoreType.alsFile;
		}

		// TODO: ActiveLock.vb - Sub New - Add documentation!
		/// <summary>
		/// Not Documented!
		/// </summary>
		/// <remarks></remarks>
		public ActiveLock()
			: base()
		{
			Class_Initialize_Renamed();
		}

		// TODO: ActiveLock.vb - Sub IActiveLock_Init - Update comment for strPath
		/// <summary>
		/// Initalizes Activelock
		/// </summary>
		/// <param name="strPath"></param>
		/// <param name="autoLicString">ByRef autoLicString As String - Returned License Key of AutoRegister is successful.</param>
		/// <remarks>
		/// <para>Performs CRC check on Alcrypto.</para>
		/// <para>Performs auto license registration if the license file is found.</para>
		/// </remarks>
		private void IActiveLock_Init([System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")]  // ERROR: Optional parameters aren't supported in C#
string strPath,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string autoLicString)
		{
			// If running in Debug mode, don't bother with dll authentication
#if AL_DEBUG
	goto Done;
#endif
			// ALL file generatiand software usage on PCs with different cultures 
			// does not work due to the usage of Chr() function in Base64_Decode in
			// modBase64.vb
			// The following is necessary to fix the problem
			ActiveLock3_6NET.My.MyProject.Application.ChangeCulture("en-US");

			// Checksum ALCrypto3NET.dll
			//Const ALCRYPTO_MD5 As String = "54BED793A0E24D3E71706EEC4FA1B0FC"
			//Const ALCRYPTO_MD5$ = "be299ad0f52858fdd9ea3626468dc05c"
			const string ALCRYPTO_MD5 = "6E5C849489281E47A9B4BB8375506D";
			//mod for VB2005'
			string strdata = string.Empty;
			string strMD5 = null;
			string usedFile = null;
			// .NET version of Activelock Init() now supports an optional path string
			// for the Alcrypto3NET.dll
			// This is needed for the cases where the user does not have the luxury of
			// placing this file in the system32 directory

			if(!string.IsNullOrEmpty(strPath))
			{
				usedFile = strPath + "\\alcrypto3NET.dll";
			}
			else
			{
				usedFile = modActiveLock.WinSysDir() + "\\alcrypto3NET.dll";
			}
			if(File.Exists(usedFile) == false)
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrFileTampered,modTrial.ACTIVELOCKSTRING,"Alcrypto3Net.dll could not be found in system32 directory.");
			}
			modActiveLock.ReadFile(usedFile,ref strdata);
			// use the .NET's native MD5 functions instead of our own MD5 hashing routine
			// and instead of ALCrypto's md5_hash() function.
			strMD5 = Strings.UCase(strdata);
			//<--- ReadFile procedure already computes the MD5.Hash


			if(strMD5 != ALCRYPTO_MD5)
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrFileTampered,modTrial.ACTIVELOCKSTRING,modActiveLock.STRFILETAMPERED);
			}
			// Perform automatic license registration
			if(!string.IsNullOrEmpty(AutoRegisterKeyPath) & mAutoRegister == IActiveLock.ALAutoRegisterTypes.alsEnableAutoRegistration)
			{
				DoAutoRegistration(ref autoLicString);
				if(Err().Number != 0) autoLicString = "";
			}
		Done:
			mfInit = true;
		}
		void _IActiveLock.Init([System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")]  // ERROR: Optional parameters aren't supported in C#
string strPath,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string autoLicString)
		{
			IActiveLock_Init(strPath,autoLicString);
		}

		/// <summary>
		/// Checks the specified path to see if the auto registration liberation file is there
		/// </summary>
		/// <param name="strLibKey">strLibKey As String - Returned liberation key if auto register is successful.</param>
		/// <remarks></remarks>

		private void DoAutoRegistration(ref string strLibKey)
		{
			// Don't bother to proceed unless the file is there.
			if(!File.Exists(AutoRegisterKeyPath)) return;


			ReadLibKey(AutoRegisterKeyPath,ref strLibKey);
			IActiveLock_Register(strLibKey);

			// If registration is successful, delete the liberation file so we won't register the same file on next startup
			FileSystem.Kill(AutoRegisterKeyPath);
		}

		/// <summary>
		/// Reads the liberation key from a file
		/// </summary>
		/// <param name="sFileName">ByVal sFileName As String - File name to read the liberation key from.</param>
		/// <param name="strLibKey">ByRef strLibKey As String -  Liberation key returned</param>
		/// <remarks></remarks>
		private void ReadLibKey(string sFileName,ref string strLibKey)
		{
			int hFile = 0;
			hFile = FileSystem.FreeFile();
			FileSystem.FileOpen(hFile,sFileName,OpenMode.Input);
			// ERROR: Not supported in C#: OnErrorStatement

			strLibKey = FileSystem.InputString(hFile,FileSystem.LOF(hFile));
		finally_Renamed:
			FileSystem.FileClose(hFile);
		}

		// TODO: ActiveLock.vb - Sub IActiveLock_Acquire - Update the input paramaters comments!
		/// <summary>
		/// <para>Acquires an Activelock License.</para>
		/// <para>This is the main method that retrieves an Activelock license, validates it, and ends the trial license if it exists.</para>
		/// </summary>
		/// <param name="strMsg"></param>
		/// <param name="strRemainingTrialDays"></param>
		/// <param name="strRemainingTrialRuns"></param>
		/// <param name="strTrialLength"></param>
		/// <param name="strUsedDays"></param>
		/// <param name="strExpirationDate"></param>
		/// <param name="strRegisteredUser"></param>
		/// <param name="strRegisteredLevel"></param>
		/// <param name="strLicenseClass"></param>
		/// <param name="strMaxCount"></param>
		/// <param name="strLicenseFileType"></param>
		/// <param name="strLicenseType"></param>
		/// <param name="strUsedLockType"></param>
		/// <remarks></remarks>
		private void IActiveLock_Acquire([System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strMsg,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strRemainingTrialDays,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strRemainingTrialRuns,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strTrialLength,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strUsedDays,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strExpirationDate,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strRegisteredUser,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strRegisteredLevel,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strLicenseClass,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strMaxCount,
   [System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strLicenseFileType,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strLicenseType,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strUsedLockType)
		{
			bool trialActivated = false;
			string adsText = string.Empty;
			string strStream = string.Empty;
			ProductLicense Lic = null;
			bool trialStatus = false;

			strStream = mSoftwareName + mSoftwareVer + mSoftwarePassword;

			// Get the current date format and save it to regionalSymbol variable
			modActiveLock.Get_locale();
			// Use this trick to temporarily set the date format to "yyyy/MM/dd"
			modActiveLock.Set_locale("");

			//Check the Key Store Provider
			if(mKeyStore == null)
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrKeyStoreInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRKEYSTOREUNINITIALIZED);
			}
			//Check the Key Store Path (LIC file path)
			else if(string.IsNullOrEmpty(mKeyStorePath))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrKeyStorePathInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRKEYSTOREPATHISEMPTY);
			}
			else if(string.IsNullOrEmpty(mSoftwareName))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareName,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOSOFTWARENAME);
			}
			else if(string.IsNullOrEmpty(mSoftwareVer))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareVersion,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOSOFTWAREVERSION);
			}
			else if(string.IsNullOrEmpty(mSoftwarePassword))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwarePassword,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOSOFTWAREPASSWORD);
			}
			else if(specialChar(mSoftwarePassword) | mSoftwarePassword.Length > 255)
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrSoftwarePasswordInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRSOFTWAREPASSWORDINVALID);
			}

			Lic = mKeyStore.Retrieve(ref mSoftwareName,mLicenseFileType);
			if(Lic == null)
			{
				// There's no valid license, so let's see if we can grant this user a "Trial License"
				//No Trial
				if(mTrialType == IActiveLock.ALTrialTypes.trialNone)
				{
					goto noRegistration;
				}

				// ERROR: Not supported in C#: OnErrorStatement

				strMsg = "";
				if(mTrialHideTypes == 0)
				{
					mTrialHideTypes = IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialSteganography;
				}

				if(mCheckTimeServerForClockTampering == IActiveLock.ALTimeServerTypes.alsCheckTimeServer)
				{
					if(modTrial.SystemClockTampered())
					{
						modActiveLock.Set_locale(modActiveLock.regionalSymbol);
						Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged,modTrial.ACTIVELOCKSTRING,modActiveLock.STRCLOCKCHANGED);
					}
				}
				if(mChecksystemfilesForClockTampering == IActiveLock.ALSystemFilesTypes.alsCheckSystemFiles)
				{
					if(modTrial.ClockTampering())
					{
						modActiveLock.Set_locale(modActiveLock.regionalSymbol);
						Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged,modTrial.ACTIVELOCKSTRING,modActiveLock.STRCLOCKCHANGED);
					}
				}

				trialStatus = modTrial.ActivateTrial(mSoftwareName,mSoftwareVer,mTrialType,mTrialLength,mTrialHideTypes,ref strMsg,mSoftwarePassword,mCheckTimeServerForClockTampering,mChecksystemfilesForClockTampering,mTrialWarning,
				ref mRemainingTrialDays,ref mRemainingTrialRuns);
				strRemainingTrialDays = mRemainingTrialDays.ToString();
				strRemainingTrialRuns = mRemainingTrialRuns.ToString();
				strTrialLength = mTrialLength.ToString();
				// Set the locale date format to what we had before; can't leave changed
				modActiveLock.Set_locale((modActiveLock.regionalSymbol));
				if(trialStatus == true)
				{
					return;
				}
				goto continueRegistration;
			noRegistration:

				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoLicense,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOLICENSE);

			}
			//Lic exists therefore we'll check the LIC file ADS
			else
			{

				if(Lic.LicenseType != ProductLicense.ALLicType.allicPermanent)
				{
					if(mCheckTimeServerForClockTampering == IActiveLock.ALTimeServerTypes.alsCheckTimeServer)
					{
						if(modTrial.SystemClockTampered())
						{
							modActiveLock.Set_locale(modActiveLock.regionalSymbol);
							Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged,modTrial.ACTIVELOCKSTRING,modActiveLock.STRCLOCKCHANGED);
						}
					}
					if(mChecksystemfilesForClockTampering == IActiveLock.ALSystemFilesTypes.alsCheckSystemFiles)
					{
						if(modTrial.ClockTampering())
						{
							modActiveLock.Set_locale(modActiveLock.regionalSymbol);
							Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged,modTrial.ACTIVELOCKSTRING,modActiveLock.STRCLOCKCHANGED);
						}
					}
				}

				if(CheckStreamCapability() & Lic.LicenseType != ProductLicense.ALLicType.allicPermanent)
				{
					FileInfo fi = new FileInfo(mKeyStorePath);
					if(fi.Length == 0) goto continueRegistration;
					adsText = ADSFile.Read(mKeyStorePath,strStream);
					if(string.IsNullOrEmpty(adsText))
					{
						modActiveLock.Set_locale(modActiveLock.regionalSymbol);
						Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseTampered,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSETAMPERED);
					}
					DateTime dt1 = Convert.ToDateTime(adsText);
					DateTime dt2 = modTrial.ActiveLockDate(System.DateTime.UtcNow);
					TimeSpan span = dt2.Subtract(dt1);
					if(span.TotalHours < 0)
					{
						modActiveLock.Set_locale(modActiveLock.regionalSymbol);
						Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged,modTrial.ACTIVELOCKSTRING,modActiveLock.STRCLOCKCHANGED);
					}
					int ok = 0;
					ok = ADSFile.Write(modTrial.ActiveLockDate(System.DateTime.UtcNow),mKeyStorePath,strStream);
					goto continueRegistration;
				}
			}
		continueRegistration:

			modActiveLock.Set_locale(modActiveLock.regionalSymbol);
			// Validate license
			ValidateLic(ref Lic);
			// Return all needed properties for faster form loading
			dontValidateLicense = true;
			DateTime mydate = modTrial.ActiveLockDate(System.DateTime.UtcNow);
			strUsedDays = (string)mydate.Subtract(modTrial.ActiveLockDate((System.DateTime)Lic.RegisteredDate)).Days;
			//CInt(DateDiff("d", CDate(Replace(Lic.RegisteredDate, ".", "-")), Date.UtcNow))
			strExpirationDate = Lic.Expiration;
			strRegisteredUser = Lic.Licensee;
			strRegisteredLevel = Lic.RegisteredLevel;
			strLicenseClass = Lic.LicenseClass;
			strMaxCount = Lic.MaxCount.ToString();
			strLicenseFileType = Conversion.Val(IActiveLock_LicenseFileType).ToString();
			strLicenseType = Lic.LicenseType.ToString();
			strUsedLockType = IActiveLock_UsedLockType.ToString();
			dontValidateLicense = false;

		}
		void _IActiveLock.Acquire([System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strMsg,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strRemainingTrialDays,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strRemainingTrialRuns,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strTrialLength,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strUsedDays,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strExpirationDate,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strRegisteredUser,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strRegisteredLevel,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strLicenseClass,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strMaxCount,
   [System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strLicenseFileType,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strLicenseType,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string strUsedLockType)
		{
			IActiveLock_Acquire(strMsg,strRemainingTrialDays,strRemainingTrialRuns,strTrialLength,strUsedDays,strExpirationDate,strRegisteredUser,strRegisteredLevel,strLicenseClass,strMaxCount,
			strLicenseFileType,strLicenseType,strUsedLockType);
		}

		// TODO: ActiveLock.vb - Function CheckStreamCapability - Add Comments - Not Documented!
		/// <summary>
		/// Not Documented!
		/// </summary>
		/// <returns></returns>
		/// <remarks></remarks>
		public bool CheckStreamCapability()
		{
			bool functionReturnValue = false;
			// The following WMI call also works but it seems to be a bit slower than the GetVolumeInformation
			// especially when it checks the A: drive
			// METHOD 1 - WMI
			//Dim mc As New ManagementClass("Win32_LogicalDisk")
			//Dim moc As ManagementObjectCollection = mc.GetInstances()
			//Dim strFileSystem As String = String.Empty
			//Dim mo As ManagementObject
			//For Each mo In moc
			//    If strFileSystem = String.Empty Then ' only return the file system
			//        If mo("Name").ToString = "C:" Then
			//            strFileSystem = mo("FileSystem").ToString
			//            Exit For
			//        End If
			//    End If
			//    mo.Dispose()
			//Next mo
			//If strFileSystem = "NTFS" Then
			//    CheckStreamCapability = True
			//End If

			// METHOD 2 - GetVolumeInformation API
			string lsRootPathName = IO.Directory.GetDirectoryRoot(Application.StartupPath);
			const int MAX_PATH = 260;
			int iSerial = 0;
			int iLength = 0;
			int iFlags = 0;
			StringBuilder sbVol = new StringBuilder(MAX_PATH);
			StringBuilder sbFil = new StringBuilder(MAX_PATH);
			GetVolumeInformation("c:\\",sbVol,MAX_PATH,iSerial,iLength,iFlags,sbFil,MAX_PATH);
			if(sbFil.ToString() == "NTFS")
			{
				functionReturnValue = true;
			}
			return functionReturnValue;

		}

		/// <summary>
		/// <para>Validates the License Key using RSA signature verification.</para>
		/// <para>License key contains the RSA signature of IActiveLock_LockCode.</para>
		/// </summary>
		/// <param name="Lic">Lic As ProductLicense - Product license</param>
		/// <remarks></remarks>
		private void ValidateKey(ref ProductLicense Lic)
		{
			string strPubKey = null;
			string strSig = null;
			string strLic = null;
			string strLicKey = null;

			strPubKey = mSoftwareCode;

			// make sure software code is set
			if(string.IsNullOrEmpty(mSoftwareCode))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNotInitialized,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOSOFTWARECODE);
			}

			strLic = IActiveLock_LockCode(ref Lic);
			strLicKey = Lic.LicenseKey;

			//ALCrypto
			if(Strings.Left(strPubKey,3) != "RSA")
			{
				// decode the license key
				strSig = modBase64.Base64_Decode(ref strLicKey);

				// Print out some info for debugging purposes
				//System.Diagnostics.Debug.WriteLine("Code1: " & strPubKey)
				//System.Diagnostics.Debug.WriteLine("Lic: " & strLic)
				//System.Diagnostics.Debug.WriteLine("Lic hash: " & modMD5.Hash(strLic))
				//System.Diagnostics.Debug.WriteLine("LicKey: " & strLicKey)
				//System.Diagnostics.Debug.WriteLine("Sig: " & strSig)
				//System.Diagnostics.Debug.WriteLine("Verify: " & RSAVerify(strPubKey, strLic, strSig))
				//System.Diagnostics.Debug.WriteLine("====================================================")

				// validate the key
				int rc = 0;
				rc = modActiveLock.RSAVerify(strPubKey,strLic,strSig);
				if(rc != 0)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
				}
			}
			// .NET RSA
			else
			{


				try
				{

					// Verify Signature
					System.Security.Cryptography.RSACryptoServiceProvider rsaCSP = new System.Security.Cryptography.RSACryptoServiceProvider();
					RSAParameters rsaPubParams = default(RSAParameters);
					//stores public key
					string strPublicBlob = null;
					if(modALUGEN.strLeft(strPubKey,6) == "RSA512")
					{
						strPublicBlob = modALUGEN.strRight(strPubKey,Strings.Len(strPubKey) - 6);
					}
					else
					{
						strPublicBlob = modALUGEN.strRight(strPubKey,Strings.Len(strPubKey) - 7);
					}
					rsaCSP.FromXmlString(strPublicBlob);
					rsaPubParams = rsaCSP.ExportParameters(false);
					// import public key params into instance of RSACryptoServiceProvider
					rsaCSP.ImportParameters(rsaPubParams);

					byte[] userData = Encoding.UTF8.GetBytes(strLic);

					byte[] newsignature = null;
					newsignature = Convert.FromBase64String(strLicKey);
					AsymmetricSignatureDeformatter asd = new RSAPKCS1SignatureDeformatter(rsaCSP);
					HashAlgorithm algorithm = new SHA1Managed();
					asd.SetHashAlgorithm(algorithm.ToString());
					byte[] newhashedData = null;
					// a byte array to store hash value
					string newhashedDataString = null;
					newhashedData = algorithm.ComputeHash(userData);
					newhashedDataString = BitConverter.ToString(newhashedData).Replace("-",string.Empty);
					bool verified = false;
					verified = asd.VerifySignature(algorithm,newsignature);
					if(verified)
					{
					}
					//MsgBox("Signature Valid", MsgBoxStyle.Information)
					else
					{
						modActiveLock.Set_locale(modActiveLock.regionalSymbol);
						Err().Raise(AlugenGlobals.alugenErrCodeConstants.alugenProdInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
						//MsgBox("Invalid Signature", MsgBoxStyle.Exclamation)
					}
				}
				catch(Exception ex)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(AlugenGlobals.alugenErrCodeConstants.alugenProdInvalid,modTrial.ACTIVELOCKSTRING,ex.Message);
				}

			}

			// Check if license has not expired
			// but don't do it if there's no expiration date
			if(string.IsNullOrEmpty(Lic.Expiration)) return;

			if(modTrial.ActiveLockDate(System.DateTime.UtcNow) > modTrial.ActiveLockDate((System.DateTime)Lic.Expiration) & Lic.LicenseType != ProductLicense.ALLicType.allicPermanent)
			{
				// ialkan - 9-23-2005 added the following to update and store the license
				// with the new LastUsed property; otherwise setting the clock back next time
				// might bypass the protection mechanism
				// Update last used date
				UpdateLastUsed(ref Lic);
				mKeyStore.Store(ref Lic,mLicenseFileType);
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseExpired,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEEXPIRED);
			}
		}

		/// <summary>
		/// Validates the License Key using the Short Key MD5 verification.
		/// </summary>
		/// <param name="Lic">Lic As ProductLicense - Product license</param>
		/// <param name="user">String - User</param>
		/// <remarks></remarks>

		private void ValidateShortKey(ref ProductLicense Lic,string user)
		{
			clsShortSerial oReg = null;
			clsShortLicenseKey m_Key = null;
			string sKey = null;
			int m_ProdCode = 0;
			string SerialNumber = null;
			System.DateTime ExpireDate = default(System.DateTime);
			short UserData = 0;
			int RegisteredLevel = 0;

			// make sure software code is set
			if(string.IsNullOrEmpty(mSoftwareCode))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNotInitialized,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOSOFTWARECODE);
			}

			//This is a short key
			m_Key = new clsShortLicenseKey();

			m_Key.AddSwapBits(0,0,1,0);
			m_Key.AddSwapBits(0,2,1,1);
			m_Key.AddSwapBits(0,4,2,0);
			m_Key.AddSwapBits(0,5,2,1);
			m_Key.AddSwapBits(2,0,3,0);
			m_Key.AddSwapBits(2,6,3,1);
			m_Key.AddSwapBits(2,7,1,3);

			oReg = new clsShortSerial();
			sKey = oReg.GenerateKey(ref "",ref Strings.Left(mSoftwareCode,Strings.Len(mSoftwareCode) - 2));
			//Do not include the last 2 possible == paddings
			m_ProdCode = (int)Strings.Left(sKey,4);

			// Shortkeys are no longer using the HDD firmware serial number
			// they are using the Computer Fingerprint after v3.6
			SerialNumber = oReg.GenerateKey(ref mSoftwareName + mSoftwareVer + mSoftwarePassword,ref modHardware.GetFingerprint());

			// verify the key is valid
			if(m_Key.ValidateShortKey(Lic.LicenseKey,SerialNumber,user,m_ProdCode,ref ExpireDate,ref UserData,ref RegisteredLevel) == true)
			{
				// After the key is disassembled it fills the output
				// variables with expire date and license counter.
				Lic.LicenseType = (int)(string)modActiveLock.HiByte(UserData);
				Lic.ProductName = mSoftwareName;
				Lic.ProductVer = mSoftwareVer;
				Lic.LicenseClass = ProductLicense.LicFlags.alfSingle;
				//Multi User License will be available with network version
				Lic.Licensee = user;
				if(Lic.RegisteredLevel == 0)
				{
					Lic.RegisteredLevel = (string)RegisteredLevel;
				}
				else if(Lic.RegisteredLevel != (string)RegisteredLevel)
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
				}
				if(string.IsNullOrEmpty(Lic.RegisteredDate))
				{
					Lic.RegisteredDate = modTrial.ActiveLockDate(System.DateTime.UtcNow).ToString("yyyy/MM/dd");
				}
				// ignore expiration date if license type is "permanent"
				if(Lic.LicenseType != ProductLicense.ALLicType.allicPermanent)
				{
					Lic.Expiration = modTrial.ActiveLockDate(ExpireDate).ToString("yyyy/MM/dd");
				}
				Lic.MaxCount = (int)(string)modActiveLock.LoByte(UserData);

				// Finally check if the serial number is Ok
				// Shortkeys are no longer using the HDD firmware serial number
				// they are using the Computer Fingerprint after v3.6
				if(!oReg.IsKeyOK(ref SerialNumber,ref mSoftwareName + mSoftwareVer + mSoftwarePassword,ref modHardware.GetFingerprint()))
				{
					// Something wrong with the serial number used
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
				}
				Lic.LicenseCode = "Short Key";
			}
			//"Key is valid."
			else
			{
				//MsgBox "Invalid license key."
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
			}

			// Check if license has not expired
			// but don't do it if there's no expiration date
			if(string.IsNullOrEmpty(Lic.Expiration)) return;

			System.DateTime dtExp = default(System.DateTime);
			dtExp = modTrial.ActiveLockDate((System.DateTime)Lic.Expiration);
			if(modTrial.ActiveLockDate(System.DateTime.UtcNow) > dtExp & Lic.LicenseType != ProductLicense.ALLicType.allicPermanent)
			{
				// ialkan - 9-23-2005 added the following to update and store the license
				// with the new LastUsed property; otherwise setting the clock back next time
				// might bypass the protection mechanism
				// Update last used date
				UpdateLastUsed(ref Lic);
				mKeyStore.Store(ref Lic,mLicenseFileType);
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseExpired,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEEXPIRED);
			}
		}

		/// <summary>
		/// Validates the entire license (including lastused, etc.)
		/// </summary>
		/// <param name="Lic">ProductLicense - Product License</param>
		/// <remarks></remarks>

		private void ValidateLic(ref ProductLicense Lic)
		{
			// Get the current date format and save it to regionalSymbol variable
			modActiveLock.Get_locale();
			// Use this trick to temporarily set the date format to "yyyy/MM/dd"
			modActiveLock.Set_locale("");

			// make sure we're initialized.
			if(!mfInit)
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNotInitialized,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOTINITIALIZED);
			}

			// validate license key first
			if(Strings.Mid(Lic.LicenseKey,5,1) == "-" & Strings.Mid(Lic.LicenseKey,10,1) == "-" & Strings.Mid(Lic.LicenseKey,15,1) == "-" & Strings.Mid(Lic.LicenseKey,20,1) == "-")
			{
				string[] arrProdVer = null;
				string actualLicensee = null;
				arrProdVer = Lic.Licensee.Split("&&&");
				actualLicensee = arrProdVer[0];
				ValidateShortKey(ref Lic,actualLicensee);
			}
			//ValidateShortKey(Lic, Lic.Licensee)
			//ALCrypto RSA key
			else
			{
				ValidateKey(ref Lic);
			}

			string strEncrypted = null;
			string strHash = null;
			// Validate last run date
			strEncrypted = Lic.LastUsed;
			MyNotifier.Notify("ValidateValue",ref strEncrypted);
			strHash = modMD5.Hash(ref strEncrypted);
			if(strHash != Lic.Hash1)
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseTampered,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSETAMPERED);
			}

			// We will compare several important dates with each other 
			// to see if anything is wrong in the license
			if(Lic.LicenseType != ProductLicense.ALLicType.allicPermanent)
			{
				// Must have NOW>LASTUSED
				if(DateAndTime.DateValue(modTrial.ActiveLockDate(System.DateTime.UtcNow)) < DateAndTime.DateValue(modTrial.ActiveLockDate((System.DateTime)Lic.LastUsed)))
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged,modTrial.ACTIVELOCKSTRING,modActiveLock.STRCLOCKCHANGED);
				}
				// Must have NOW<EXPIRATION
				if(DateAndTime.DateValue(modTrial.ActiveLockDate(System.DateTime.UtcNow)) > DateAndTime.DateValue(modTrial.ActiveLockDate((System.DateTime)Lic.Expiration)))
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseExpired,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEEXPIRED);
				}
				// Must have LASTUSED>=REGISTERED
				if(DateAndTime.DateValue(modTrial.ActiveLockDate(Lic.LastUsed)) < DateAndTime.DateValue(modTrial.ActiveLockDate((System.DateTime)Lic.RegisteredDate)))
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged,modTrial.ACTIVELOCKSTRING,modActiveLock.STRCLOCKCHANGED);
				}
			}
			UpdateLastUsed(ref Lic);
			mKeyStore.Store(ref Lic,mLicenseFileType);
			modActiveLock.Set_locale(modActiveLock.regionalSymbol);
		}

		/// <summary>
		/// Updates LastUsed property with current date stamp.
		/// </summary>
		/// <param name="Lic">ProductLicense - Product License</param>
		/// <remarks></remarks>
		private void UpdateLastUsed(ref ProductLicense Lic)
		{
			// Update license store with LastRunDate
			string strLastUsed = null;
			modActiveLock.Set_locale("");
			strLastUsed = modTrial.ActiveLockDate(System.DateTime.UtcNow);
			strLastUsed = modTrial.ActiveLockDate((System.DateTime)strLastUsed);
			Lic.LastUsed = strLastUsed;
			MyNotifier.Notify("ValidateValue",ref strLastUsed);
			Lic.Hash1 = modMD5.Hash(ref strLastUsed);
		}

		/// <summary>
		/// Registers Activelock license with a given liberation key
		/// </summary>
		/// <param name="LibKey">String - Liberation Key</param>
		/// <param name="user">Optional - String - User</param>
		/// <remarks></remarks>

		private void IActiveLock_Register(string LibKey,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string user)
		{
			ActiveLock3_6NET.ProductLicense Lic = new ActiveLock3_6NET.ProductLicense();
			object varResult = null;
			bool trialStatus = false;

			// Get the current date format and save it to regionalSymbol variable
			modActiveLock.Get_locale();
			// Use this trick to temporarily set the date format to "yyyy/MM/dd"
			modActiveLock.Set_locale("");

			// Check to see if this is a Short License Key
			if(Strings.Mid(LibKey,5,1) == "-" & Strings.Mid(LibKey,10,1) == "-" & Strings.Mid(LibKey,15,1) == "-" & Strings.Mid(LibKey,20,1) == "-")
			{
				Lic.LicenseKey = Strings.UCase(LibKey);
				ValidateShortKey(ref Lic,user);
			}
			// RSA key
			else
			{
				Lic.Load(LibKey);
				// Validate that the license key.
				//   - registered user
				//   - expiry date
				ValidateKey(ref Lic);
			}

			// License was validated successfuly. Check clock tampering for non-permanent licenses.
			if(Lic.LicenseType != ProductLicense.ALLicType.allicPermanent)
			{
				if(mCheckTimeServerForClockTampering == IActiveLock.ALTimeServerTypes.alsCheckTimeServer)
				{
					if(modTrial.SystemClockTampered())
					{
						modActiveLock.Set_locale(modActiveLock.regionalSymbol);
						Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged,modTrial.ACTIVELOCKSTRING,modActiveLock.STRCLOCKCHANGED);
					}
				}
				if(mChecksystemfilesForClockTampering == IActiveLock.ALSystemFilesTypes.alsCheckSystemFiles)
				{
					if(modTrial.ClockTampering())
					{
						modActiveLock.Set_locale(modActiveLock.regionalSymbol);
						Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged,modTrial.ACTIVELOCKSTRING,modActiveLock.STRCLOCKCHANGED);
					}
				}
			}

			// License was validated successfuly.  Store it.
			if(mKeyStore == null)
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrKeyStoreInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRKEYSTOREUNINITIALIZED);
			}

			// Update last used date
			UpdateLastUsed(ref Lic);
			mKeyStore.Store(ref Lic,mLicenseFileType);

			// This works under NTFS and is needed to prevent clock tampering
			if(CheckStreamCapability() & Lic.LicenseType != ProductLicense.ALLicType.allicPermanent)
			{
				// Write the current date and time into the ads
				int ok = 0;
				string strStream = string.Empty;
				strStream = mSoftwareName + mSoftwareVer + mSoftwarePassword;
				ok = ADSFile.Write(modTrial.ActiveLockDate(System.DateTime.UtcNow).ToString(),mKeyStorePath,strStream);
			}

			// Expire all trial licenses
			// ERROR: Not supported in C#: OnErrorStatement

			// Expire the Trial
			if(mTrialType != IActiveLock.ALTrialTypes.trialNone)
			{
				trialStatus = modTrial.ExpireTrial(mSoftwareName,mSoftwareVer,mTrialType,mTrialLength,mTrialHideTypes,mSoftwarePassword);
			}
			modActiveLock.Set_locale(modActiveLock.regionalSymbol);

		}
		void _IActiveLock.Register(string LibKey,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute("")] ref  // ERROR: Optional parameters aren't supported in C#
string user)
		{
			IActiveLock_Register(LibKey,user);
		}

		/// <summary>
		/// Kills a Trial License
		/// </summary>
		/// <remarks></remarks>
		private void IActiveLock_KillTrial()
		{
			// ERROR: Not supported in C#: OnErrorStatement

			//Expire the Trial
			bool trialStatus = false;
			if(string.IsNullOrEmpty(mSoftwareName))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareName,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOSOFTWARENAME);
			}
			else if(string.IsNullOrEmpty(mSoftwareVer))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareVersion,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOSOFTWAREVERSION);
			}
			else if(string.IsNullOrEmpty(mSoftwarePassword))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwarePassword,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOSOFTWAREPASSWORD);
			}
			else
			{
				trialStatus = modTrial.ExpireTrial(mSoftwareName,mSoftwareVer,mTrialType,mTrialLength,mTrialHideTypes,mSoftwarePassword);
			}
		}
		void _IActiveLock.KillTrial()
		{
			IActiveLock_KillTrial();
		}

		// TODO: ActiveLock.vb - Function IActiveLock_GenerateShortKey - Needs Documentation!
		/// <summary>
		/// Not Documented!
		/// </summary>
		/// <param name="SoftwareCode"></param>
		/// <param name="SerialNumber"></param>
		/// <param name="LicenseeAndRegisteredLevel"></param>
		/// <param name="Expiration"></param>
		/// <param name="LicType"></param>
		/// <param name="RegisteredLevel"></param>
		/// <param name="MaxUsers"></param>
		/// <returns></returns>
		/// <remarks></remarks>

		private string IActiveLock_GenerateShortKey(string SoftwareCode,string SerialNumber,string LicenseeAndRegisteredLevel,string Expiration,ProductLicense.ALLicType LicType,int RegisteredLevel,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute(1)]  // ERROR: Optional parameters aren't supported in C#
short MaxUsers)
		{
			string functionReturnValue = null;
			// ERROR: Not supported in C#: OnErrorStatement


			clsShortLicenseKey m_Key = null;
			m_Key = new clsShortLicenseKey();

			m_Key.AddSwapBits(0,0,1,0);
			m_Key.AddSwapBits(0,2,1,1);
			m_Key.AddSwapBits(0,4,2,0);
			m_Key.AddSwapBits(0,5,2,1);
			m_Key.AddSwapBits(2,0,3,0);
			m_Key.AddSwapBits(2,6,3,1);
			m_Key.AddSwapBits(2,7,1,3);

			clsShortSerial oReg = null;
			oReg = new clsShortSerial();
			string sKey = null;
			int m_ProdCode = 0;

			sKey = oReg.GenerateKey(ref "",ref Strings.Left(SoftwareCode,Strings.Len(SoftwareCode) - 2));
			//Do not include the last 2 possible == paddings
			m_ProdCode = (int)Strings.Left(sKey,4);

			// create a new key
			functionReturnValue = m_Key.CreateShortKey(SerialNumber,LicenseeAndRegisteredLevel,m_ProdCode,(System.DateTime)Expiration,modActiveLock.MakeWord((string)MaxUsers,(string)LicType),RegisteredLevel);

			return;
		ErrHandler:
			oReg = null;
			m_Key = null;
			return functionReturnValue;

		}
		string _IActiveLock.GenerateShortKey(string SoftwareCode,string SerialNumber,string LicenseeAndRegisteredLevel,string Expiration,ProductLicense.ALLicType LicType,int RegisteredLevel,[System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute(1)]  // ERROR: Optional parameters aren't supported in C#
short MaxUsers)
		{
			return IActiveLock_GenerateShortKey(SoftwareCode,SerialNumber,LicenseeAndRegisteredLevel,Expiration,LicType,RegisteredLevel,MaxUsers);
		}

		/// <summary>
		/// Resets a Trial License
		/// </summary>
		/// <remarks></remarks>
		private void IActiveLock_ResetTrial()
		{
			// ERROR: Not supported in C#: OnErrorStatement

			//Reset the Trial
			bool trialStatus = false;
			if(string.IsNullOrEmpty(mSoftwareName))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareName,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOSOFTWARENAME);
			}
			else if(string.IsNullOrEmpty(mSoftwareVer))
			{
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNoSoftwareVersion,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOSOFTWAREVERSION);
			}
			else
			{
				trialStatus = modTrial.ResetTrial(mSoftwareName,mSoftwareVer,mTrialType,mTrialLength,mTrialHideTypes,mSoftwarePassword);
			}
		}
		void _IActiveLock.ResetTrial()
		{
			IActiveLock_ResetTrial();
		}

		/// <summary>
		/// Returns the lock code from a given Activelock license
		/// </summary>
		/// <param name="Lic">ProductLicense - Product License</param>
		/// <returns>String - Lock code</returns>
		/// <remarks>v3 includes the new lockHDFirmware option</remarks>
		private string IActiveLock_LockCode(
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute(null)
			] ref ProductLicense Lic)
		{
			string functionReturnValue = null;
			string strLock = string.Empty;
			string noKey = null;
			string userFromInstallCode = null;
			string usedcode = null;
			//Dim tmpLockType As IActiveLock.ALLockTypes
			short j = 0;
			short Index = 0;
			short i = 0;
			string[] a = null;
			string aString = null;
			// we have lockNone ' use temp in case of failure.

			noKey = Strings.Chr(110) + Strings.Chr(111) + Strings.Chr(107) + Strings.Chr(101) + Strings.Chr(121);
			if(Lic == null)
			{
				// New in v3.1
				// Modified this function on 1-13-2006 to append ALL hardware keys
				// to the Installation Code. This way, it will be decided in Alugen which
				// hardware keys will be used to lock the license to
				// If there's already a lock selected in the protected app,
				// such as lockHDfirmware or lockComputer, then Alugen will show
				// these two options and will gray them out (fix these two selections)
				if(mLockTypes == IActiveLock.ALLockTypes.lockNone)
				{
					strLock = "";
					AppendLockString(ref strLock,modHardware.GetMACAddress());
					AppendLockString(ref strLock,modHardware.GetComputerName());
					AppendLockString(ref strLock,modHardware.GetHDSerial());
					AppendLockString(ref strLock,modHardware.GetHDSerialFirmware());
					AppendLockString(ref strLock,modHardware.GetWindowsSerial());
					AppendLockString(ref strLock,modHardware.GetBiosVersion());
					AppendLockString(ref strLock,modHardware.GetMotherboardSerial());
					AppendLockString(ref strLock,modHardware.GetIPaddress());
					AppendLockString(ref strLock,modHardware.GetExternalIP());
					AppendLockString(ref strLock,modHardware.GetFingerprint());
					AppendLockString(ref strLock,modHardware.GetMemoryID());
					AppendLockString(ref strLock,modHardware.GetCPUID());
					AppendLockString(ref strLock,modHardware.GetBaseBoardID());
					AppendLockString(ref strLock,modHardware.GetVideoID());
				}
				else
				{
					//mLockTypes And IActiveLock.ALLockTypes.lockMAC Then
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockMAC))
					{
						AppendLockString(ref strLock,modHardware.GetMACAddress());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					//mLockTypes And IActiveLock.ALLockTypes.lockComp Then
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockComp))
					{
						AppendLockString(ref strLock,modHardware.GetComputerName());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					//mLockTypes And IActiveLock.ALLockTypes.lockHD Then
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockHD))
					{
						AppendLockString(ref strLock,modHardware.GetHDSerial());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					//mLockTypes And IActiveLock.ALLockTypes.lockHDFirmware Then
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockHDFirmware))
					{
						AppendLockString(ref strLock,modHardware.GetHDSerialFirmware());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					//mLockTypes And IActiveLock.ALLockTypes.lockWindows Then
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockWindows))
					{
						AppendLockString(ref strLock,modHardware.GetWindowsSerial());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					//mLockTypes And IActiveLock.ALLockTypes.lockBIOS Then
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockBIOS))
					{
						AppendLockString(ref strLock,modHardware.GetBiosVersion());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					//mLockTypes And IActiveLock.ALLockTypes.lockMotherboard Then
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockMotherboard))
					{
						AppendLockString(ref strLock,modHardware.GetMotherboardSerial());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					//mLockTypes And IActiveLock.ALLockTypes.lockIP Then
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockIP))
					{
						AppendLockString(ref strLock,modHardware.GetIPaddress());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					// new in v3.6
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockExternalIP))
					{
						AppendLockString(ref strLock,modHardware.GetExternalIP());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockFingerprint))
					{
						AppendLockString(ref strLock,modHardware.GetFingerprint());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockMemory))
					{
						AppendLockString(ref strLock,modHardware.GetMemoryID());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockCPUID))
					{
						AppendLockString(ref strLock,modHardware.GetCPUID());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockBaseboardID))
					{
						AppendLockString(ref strLock,modHardware.GetBaseBoardID());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}
					if(modActiveLock.IsNumberIncluded(mLockTypes,IActiveLock.ALLockTypes.lockvideoID))
					{
						AppendLockString(ref strLock,modHardware.GetVideoID());
					}
					else
					{
						AppendLockString(ref strLock,noKey);
					}

				}

				if(Strings.Left(strLock,1) == Constants.vbLf) strLock = Strings.Mid(strLock,2);

				// Append lockcode.
				// Note: The logic here must match the corresponding logic
				//       in ALUGENLib.Generator_GenKey()
				functionReturnValue = strLock;
			}
			else
			{
				// We have a License
				// New in v3.1
				// In such cases when Alugen modifies the Installation Code and sends it
				// back, we need to retrieve in here and process it
				// Modified Installation Code is appended to the end of the Liberation Key
				// The modified Installation Code is also stored in the license file
				// otherwise we'd never know which hardware leys were used to lock the license
				//IActiveLock_LockCode = Lic.ToString_Renamed() & vbLf & strLock

				// Per David Weatherall ' New in v3.3
				//tmpLockType = IActiveLock.ALLockTypes.lockNone ' lockNone = 0 so starting value

				// ERROR: Not supported in C#: ReDimStatement

				// remove all previous
				int SizeLockType = 0;
				// use to build up LockCode.
				SizeLockType = 0;

				if(!string.IsNullOrEmpty(Lic.LicenseCode))
				{
					if(Strings.Left(Lic.LicenseCode,1) == "+")
					{
						usedcode = modBase64.Base64_Decode(ref Strings.Mid(Lic.LicenseCode,2));
						//bLockNone = True ' per David Weatherall
						IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockNone,ref SizeLockType);
						//dw1 build up lockTypes - start with lockNone
					}
					else
					{
						usedcode = modBase64.Base64_Decode(ref (Lic.LicenseCode));
						//bLockNone = False ' per David Weatherall
					}
					a = Strings.Split(usedcode,Constants.vbLf);
					for(j = Information.LBound(a);j <= Information.UBound(a) - 1;j++)
					{
						aString = a[j];
						if(Strings.Left(aString,1) == "+") aString = Strings.Mid(aString,2);
						if(j == Information.LBound(a))
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockMAC,ref SizeLockType);
								if(aString != modHardware.GetMACAddress())
								{
									// Ok MAC address did not match
									// Maybe the laptop owner turned on the wireless connection
									// and it wasn't on when the license was registered
									if(modHardware.CheckMACaddress(aString) == false)
									{
										// we truly failed
										modActiveLock.Set_locale(modActiveLock.regionalSymbol);
										Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
									}
								}
							}
						}
						else if(j == Information.LBound(a) + 1)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockComp,ref SizeLockType);
								if(aString != modHardware.GetComputerName())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 2)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockHD,ref SizeLockType);
								if(aString != modHardware.GetHDSerial())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 3)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockHDFirmware,ref SizeLockType);
								if(aString != modHardware.GetHDSerialFirmware())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 4)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockWindows,ref SizeLockType);
								if(aString != modHardware.GetWindowsSerial())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 5)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockBIOS,ref SizeLockType);
								if(aString != modHardware.GetBiosVersion())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 6)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockMotherboard,ref SizeLockType);
								if(aString != modHardware.GetMotherboardSerial())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 7)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockIP,ref SizeLockType);
								string returnedIP = modHardware.GetIPaddress();
								if(returnedIP == "-1")
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrInternetConnectionError,modTrial.ACTIVELOCKSTRING,modActiveLock.STRINTERNETNOTCONNECTED);
								}
								if(aString != returnedIP)
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrWrongIPaddress,modTrial.ACTIVELOCKSTRING,modActiveLock.STRWRONGIPADDRESS);
								}
							}
						}
						//added v3.6
						else if(j == Information.LBound(a) + 8)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockExternalIP,ref SizeLockType);
								if(aString != modHardware.GetExternalIP())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 9)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockFingerprint,ref SizeLockType);
								if(aString != modHardware.GetFingerprint())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 10)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockMemory,ref SizeLockType);
								if(aString != modHardware.GetMemoryID())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 11)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockCPUID,ref SizeLockType);
								if(aString != modHardware.GetCPUID())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 12)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockBaseboardID,ref SizeLockType);
								if(aString != modHardware.GetBaseBoardID())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}
						else if(j == Information.LBound(a) + 13)
						{
							if(aString != noKey)
							{
								IActiveLock_AddLockCode(IActiveLock.ALLockTypes.lockvideoID,ref SizeLockType);
								if(aString != modHardware.GetVideoID())
								{
									modActiveLock.Set_locale(modActiveLock.regionalSymbol);
									Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
								}
							}
						}

					}

					Index = 0;
					i = 1;
					// Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
					while(i > 0)
					{
						i = Strings.InStr(Index + 1,usedcode,Constants.vbLf);
						if(i > 0) Index = i;
					}
					// user name starts from Index+1 to the end
					userFromInstallCode = Strings.Mid(usedcode,Index + 1);
					// Check to see if this user name matches the one in the liberation key
					if(userFromInstallCode != Lic.Licensee)
					{
						modActiveLock.Set_locale(modActiveLock.regionalSymbol);
						Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
					}

					usedcode = Strings.Mid(usedcode,1,Strings.Len(usedcode) - Strings.Len(userFromInstallCode) - 1);
					functionReturnValue = Lic.ToString_Renamed() + Constants.vbLf + usedcode;
				}
				else
				{
					modActiveLock.Set_locale(modActiveLock.regionalSymbol);
					Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrLicenseInvalid,modTrial.ACTIVELOCKSTRING,modActiveLock.STRLICENSEINVALID);
					return null;
				}
			}
			return functionReturnValue;
		}
		string _IActiveLock.LockCode([System.Runtime.InteropServices.OptionalAttribute,System.Runtime.InteropServices.DefaultParameterValueAttribute(null)] ref  // ERROR: Optional parameters aren't supported in C#
ProductLicense Lic)
		{
			return IActiveLock_LockCode(Lic);
		}

		/// <summary>
		/// Appends the lock string to the given installation code
		/// </summary>
		/// <param name="strLock">String - The lock string to be appended to, returns as an output</param>
		/// <param name="newSubString">String - The string to be appended to the lock string if strLock is empty string</param>
		/// <remarks></remarks>
		private void AppendLockString(ref string strLock,string newSubString)
		{
			if(string.IsNullOrEmpty(strLock))
			{
				strLock = newSubString;
			}
			else
			{
				strLock = strLock + Constants.vbLf + newSubString;
			}
		}

		/// <summary>
		/// Not implemented yet
		/// </summary>
		/// <param name="OtherSoftwareCode">String - Installation code from another machine/software</param>
		/// <returns></returns>
		/// <remarks>Transfers an Activelock license from one machine/software to another</remarks>
		private string IActiveLock_Transfer(string OtherSoftwareCode)
		{
			// TODO: ActiveLock.vb - Function IActiveLock_Transfer - Implement me!
			modActiveLock.Set_locale(modActiveLock.regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNotImplemented,modTrial.ACTIVELOCKSTRING,modActiveLock.STRNOTIMPLEMENTED);
			return null;
		}
		string _IActiveLock.Transfer(string OtherSoftwareCode)
		{
			return IActiveLock_Transfer(OtherSoftwareCode);
		}

		//*******************************************************************************
		// Sub GenerateShortSerial
		//
		// Input:
		// appNameVersionPassword
		// HDDfirmwareSerial
		//
		// DESCRIPTION:
		// Generates a Short Key (Serial Number)
		/// <summary>
		/// Generates a Short Key (Serial Number)
		/// </summary>
		/// <param name="HDDfirmwareSerial"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private string IActivelock_GenerateShortSerial(string HDDfirmwareSerial)
		{
			string functionReturnValue = null;
			clsShortSerial oReg = null;
			string sKey = null;

			oReg = new clsShortSerial();
			sKey = oReg.GenerateKey(ref mSoftwareName + mSoftwareVer + mSoftwarePassword,ref HDDfirmwareSerial);
			functionReturnValue = sKey;
			// If longer serial is used, possible to break up into sections
			//Left(sKey, 4) & "-" & Mid(sKey, 5, 4) & "-" & Mid(sKey, 9, 4) & "-" & Mid(sKey, 13, 4)

			oReg = null;
			return functionReturnValue;
		}
		string _IActiveLock.GenerateShortSerial(string HDDfirmwareSerial)
		{
			return IActivelock_GenerateShortSerial(HDDfirmwareSerial);
		}

		// TODO: ActiveLock.vb - Function specialChar - Add documentation - Not Documented!
		/// <summary>
		/// Not Documented!
		/// </summary>
		/// <param name="s"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		private bool specialChar(string s)
		{
			bool functionReturnValue = false;
			int k = 0;
			s = s + Strings.Space(1);
			//check against null-strings
			for(k = 1;k <= Strings.Len(s);k++)
			{
				switch(Strings.Asc(Strings.Mid(s,k)))
				{
					case 32: // TODO: to 126
						break;
					//continue
					default:
						functionReturnValue = true;
						return;

						break;
				}
			}
			return functionReturnValue;
		}


	}

}