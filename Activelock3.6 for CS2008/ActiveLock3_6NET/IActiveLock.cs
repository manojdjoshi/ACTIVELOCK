// This project was converted to C#, from VB, using https://sourceforge.net/projects/sharpdevelop/
// there are allot of errors, but for the most part, all the code converted...
//
// This project is available from SVN on SourceForge.net under the main project, Activelock !
//
// ProjectPage: https://sourceforge.net/projects/activelock/
// WebSite: http://ActiveLockSoftware.com
// DeveloperForums: http://www.ActiveLockSoftware.com/simplemachinesforum
// ProjectManager: Ismail Alkan - http://ActiveLockSoftware.com/simplemachinesforum/index.php?action=profile;u=1
// ProjectLicense: BSD Open License - http://www.opensource.org/licenses/bsd-license.php
// ProjectPurpose: Software Locking, Anti Piracy
//
// I, wanted to convert this project to C# for several reasons:
//
//      * To make Activelock under Microsoft OS's available to:
//        ** 
//        ** Silverlight http://silverlight.live.com
//        ** XNA http://creators.xna.com
//        ** XBox http://creators.xna.com
//        ** Zune http://creators.xna.com
//        ** .NET Micro
//        **
//
//      * To make Activelock available to Linux and Mac OS's through Mono http://www.mono-project.com/
//        ** Novell-Moonlight https://sourceforge.net/projects/freshmeat_novell-moonlight/
//
////////////////////////////////////////////////////////////////////////////////////////////
//*   ActiveLock
//*   Copyright 1998-2002 Nelson Ferraz
//*   Copyright 2003-2009 The ActiveLock Software Group (ASG)
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

using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

	/// <summary>_IActiveLock - Interface - Implimented by IActiveLock</summary>
	/// <remarks>
	/// <para> - MaintainedBy:</para>
	/// <para> - LastRevisionDate:</para>
	/// <para> - Comments:</para></remarks>
	public interface _IActiveLock
	{
		int RemainingTrialDays{get;}
		int RemainingTrialRuns{get;}
		int MaxCount{get;}
		string RegisteredLevel{get;}
		string LicenseClass{get;}
		IActiveLock.ALLockTypes LockType{get;set;}
		// change this to a method
		IActiveLock.ALLicenseKeyTypes LicenseKeyType{set;}
		int UsedLockType{get;}
		IActiveLock.ALTrialHideTypes TrialHideType{get;set;}
		IActiveLock.ALTrialTypes TrialType{get;set;}
		int TrialLength{get;set;}
		string SoftwareName{get;set;}
		string SoftwarePassword{get;set;}
		// change this to a method
		IActiveLock.ALTimeServerTypes CheckTimeServerForClockTampering{set;}
		// change this to a method
		IActiveLock.ALSystemFilesTypes CheckSystemFilesForClockTampering{set;}
		IActiveLock.ALLicenseFileTypes LicenseFileType{get;set;}
		// change this to a method
		IActiveLock.ALAutoRegisterTypes AutoRegister{set;}
		// change this to a method
		IActiveLock.ALTrialWarningTypes TrialWarning{set;}
		// change this to a method
		string SoftwareCode{set;}
		string SoftwareVersion{get;set;}
		// change this to a method
		IActiveLock.LicStoreType KeyStoreType{set;}
		// change this to a method
		string KeyStorePath{set;}
		//ReadOnly Property InstallationCode(Optional ByVal User As String = vbNullString, Optional ByVal Lic As ProductLicense = Nothing) As String
		string InstallationCode { get; }
		// change this to a method
		string AutoRegisterKeyPath{set;}
		string LockCode(
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute(null)
			] ref ProductLicense Lic);
		void Register(
			string LibKey,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string user);
		string Transfer(
			string InstallCode);
		void Init(
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strPath,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string autoLicString);
		void Acquire(
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strMsg,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strRemainingTrialDays,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strRemainingTrialRuns,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strTrialLength,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strUsedDays,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strExpirationDate,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strRegisteredUser,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strRegisteredLevel,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strLicenseClass,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strMaxCount,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strLicenseFileType,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strLicenseType,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref  string strUsedLockType);
		void ResetTrial();
		void KillTrial();
		string GenerateShortSerial(
			string HDDfirmwareSerial);
		string GenerateShortKey(
			string SoftwareCode,
			string SerialNumber,
			string LicenseeAndRegisteredLevel,
			string Expiration,
			ProductLicense.ALLicType LicType,
			int RegisteredLevel,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute(1)
			] ref short MaxUsers);
		ActiveLockEventNotifier EventNotifier {get;}
		int UsedDays{get;}
		string RegisteredDate{get;}
		string RegisteredUser{get;}
		string ExpirationDate{get;}
	}

	/// <summary>
	/// IActiveLock - Impliments _IActiveLock
	/// </summary>
	/// <remarks>Class instancing was changed to public.</remarks>
	[System.Runtime.InteropServices.ProgId("IActiveLock_NET.IActiveLock")]
	public class IActiveLock:_IActiveLock
	{

		#region "Notes"
		//===============================================================================
		// Name: IActivelock
		// Purpose: This is the main interface into ActiveLock&#39;s functionalities.
		// The user application interacts with ActiveLock primarily through this IActiveLock interface.
		// Typically, the application would obtain an instance of this interface via the
		// <a href="Globals.NewInstance.html">ActiveLock3.NewInstance()</a> accessor method. From there, initialization calls are done,
		// and then various method such as <a href="IActiveLock.Register.html">Register()</a>, <a href="IActiveLock.Acquire.html">Acquire()</a>, etc..., can be used.
		// <p>
		// ActiveLock also sends COM event notifications to the user application whenever it needs help to perform
		// some action, such as license property validation/encryption.  The user application can intercept
		// these events via the ActiveLockEventNotifier object, which can be obtained from
		// <a href="IActiveLock.Get.EventNotifier.html">IActiveLock.EventNotifier</a> property.
		// <p>
		// <b>Important Note</b><br>
		// The user application is strongly advised to perform a checksum on the
		// ActiveLock DLL prior to accessing and interacting with ActiveLock. Using the checksum, you can tell if
		// the DLL has been tampered. Please refer to sample code below on how the checksumming can be done.
		// <p>
		// The sample code fragments below illustrate the typical usage flow between your application and ActiveLock.
		// Please note that the code shown is only for illustration purposes and is not meant to be a complete
		// compilable program. You may have to add variable declarations and function definitions around the code
		// fragments before you can compile it.
		// <p>
		// <pre>
		//   Form1.frm:
		//   ...
		//   Private MyActiveLock As ActiveLock3.IActiveLock
		//   Private WithEvents ActiveLockEventSink As ActiveLockEventNotifier
		//   Private Const AL_CRC& = 123+ &#39; ActiveLock3.dll&#39;s CRC checksum to be used for validation
		//
		//   &#39; This key will be used to set <a href="IActiveLock.Let.SoftwareCode.html">IActiveLock.SoftwareCode</a> property.
		//   &#39; NOTE: This is NOT a complete key (complete key is to long to put in documentation).
		//   &#39; You will generate your own product code using ALUGEN.  This is the <code>VCode</code> generated
		//   by ALUGEN.
		//   Private Const PROD_CODE$ = "AAAAB3NzaC1yc2EAAAABJQAAAIBZnXD4IKfrBH25ekwLWQMs5mJ..."
		//
		//   ....
		//
		//   Private Sub Form_Load()
		//       On Error GoTo ErrHandler
		//       &#39; Obtain an instance of AL. NewInstance() also calls IActiveLock.Init()
		//       Set MyActiveLock = ActiveLock3.NewInstance()
		//       &#39; At this point, either AL has been initialized or an error would have already been raised
		//       &#39; if there were problems (such as ActiveLock3.dll having been tampered).
		//
		//       &#39; Verify AL&#39;s authenticity
		//       &#39; modActiveLock.CRCCheckSumTypeLib() requires a public-creatable object to be passed in
		//       &#39; so that it can determine the Type Library DLL on which to perform the checksum.
		//       &#39; So can&#39;t use MyActiveLock object to authenticate since it is not a public creatable object.
		//       &#39; So we&#39;ll use ActiveLock3.Globals, which is just as good because they are in the same DLL.
		//       Dim crc As Long
		//       crc = CRCCheckSumTypeLib(New ActiveLock3.Globals)
		//       Debug.Print "Hash: " & crc
		//       If crc &lt;&gt; AL_CRC Then
		//           MsgBox "ActiveLock3.dll has been corrupted."
		//           End ' terminate
		//       End If
		//
		//      &#39; Initialize the keystore. We use a File keystore in this case.
		//      MyActiveLock.KeyStoreType = alsFile
		//      MyActiveLock.KeyStorePath = App.path & "\myapp.lic"
		//
		//      &#39; Obtain the EventNotifier so that we can receive notifications from AL.
		//      Set ActiveLockEventSink = MyActiveLock.EventNotifier
		//
		//      &#39; Specify the name of the product that will be locked through AL.
		//      MyActiveLock.SoftwareName = "MyApp"
		//
		//      &#39; Specify our product code.  This code will be used later by ActiveLock to validate license keys.
		//      MyActiveLock.SoftwareCode = PROD_CODE
		//
		//      &#39; Specify product version
		//       MyActiveLock.SoftwareVersion = txtVersion
		//
		//      &#39; Specify Lock Type
		//      MyActiveLock.LockType = lockHD
		//
		//      &#39; Sets path to liberation key file for automatic registration
		//      MyActiveLock.AutoRegisterKeyPath = App.path & "\myapp.alb"
		//
		//      &#39; Initialize the instance
		//      MyActiveLock.Init
		//
		//      &#39; Check registration status by calling Acquire()
		//      &#39; Note: Calling Acquire() may trigger ActiveLockEventNotifier_ValidateValue() event.
		//      &#39; So we should be prepared to handle that.
		//      MyActiveLock.Acquire
		//
		//      &#39; By now, if the product is not registered, then an error would have been raised,
		//      &#39; which means if we get to here, then we're registered.
		//
		//      &#39; Just for fun, print out some registration status info
		//      Debug.Print "Registered User: " & MyActiveLock.RegisteredUser
		//      Debug.Print "Used Days: " & MyActiveLock.UsedDays
		//      Debug.Print "Expiration Date: " & MyActiveLock.ExpirationDate
		//      Exit Sub
		//   ErrHandler:
		//       MsgBox Err.Number & ": " & Err.Description
		//       &#39; End program
		//       End
		//   End Sub
		//   ...
		//   <p>
		//   (Optional) ActiveLock raises this event typically when it needs a value to be encrypted.
		//   We can use any kind of encryption we&#39;d like here, as long as it&#39;s deterministic.
		//   i.e. there&#39;s a one-to-one correspondence between unencrypted value and encrypted value.
		//   NOTE: BlowFish is NOT an example of deterministic encryption so you can&#39;t use it here.
		//   You are allowed to use asymmetric algorithm since you will never be asked to decrypt a value,
		//   only to encrypt.
		//   You don&#39;t have to handle this event if you don&#39;t want to; it just means that the value WILL NOT
		//   be encrypted when it is saved to the keystore.
		//
		//   Private Sub ActiveLockEventSink_ValidateValue(ByVal Value As String, Result As String)
		//       Result = Encrypt(Value)
		//   End Sub
		//<br>
		//   &#39; Roll our own simple-yet-weird encryption routine.
		//   &#39; Must keep in mind that our encryption algorithm must be deterministic.
		//   &#39; In other words, given the same uncrypted string, it must always yield the same encrypted string.
		//   Private Function Encrypt(strData As String) As String
		//       Dim i&, n&
		//       dim sResult$
		//       n = Len(strData)
		//       For i = 1 to n
		//           sResult = sResult & Asc(Mid$(strData, i, 1)) * 7
		//       Next i
		//       Encrypt = sResult
		//   End Function
		//   ...
		//   &#39; Returns the CRC checksum of the ActiveLock3.dll.
		//   Private Property Get ALCRC() As Long
		//       &#39; Don&#39;t just return a single value, but rather compute it using some simple arithmetic
		//       &#39; so that hackers can&#39;t easily find it with a hex editor.
		//       &#39; Of course, the values below will not make up the real checksum. For the most up-to-date
		//       &#39; checksum, please refer to the ActiveLock Release Notes.
		//       ALCRC = 123 + 456
		//   End Property
		// </pre>
		//
		// <p>Generating registration code from the user application to be sent to the vendor in exchange for
		// a liberation key.
		// <pre>
		//    &#39; Generate Installation code
		//    Dim strInstCode As String
		//    strInstCode = MyActiveLock.InstallCode(txtUser)
		//    &#39; strInstCode now contains the request code to be sent to the vendor for activation.
		// </pre>
		//
		// <p>Key Registration functionality - register using a liberation key.
		// <pre>
		//    On Error GoTo ErrHandler
		//    &#39; Register this key
		//    &#39; txtLibKey contains the liberation key entered by the user.
		//    &#39; This key could have be sent via an email to the user or a program that automatically
		//    &#39; requests the key from a registration website.
		//    MyActiveLock.Register txtLibKey
		//    MsgBox "Registration successful!"
		//    Exit Sub
		//ErrHandler:
		//    MsgBox Err.Number & ": " & Err.Description
		// </pre>
		// Remarks:
		// Functions:
		// Properties:
		// Methods:
		// Started: 21.04.2005
		// Modified: 08.05.2005
		//===============================================================================
		// @author activelock-admins
		// @version 3.3.0
		// @date 3-23-2006
		#endregion

		#region "Public Enums"

		/// <summary>License Lock Types.</summary>
		/// <remarks>Values can be combined (OR ed) together.</remarks>
		public enum ALLockTypes
		{
			/// <summary>No locking - not recommended</summary>
			/// <remarks></remarks>
			lockNone=0,
			/// <summary>Lock to Network Interface Card Address</summary>
			/// <remarks></remarks>
			lockMAC=1,
			/// <summary>Lock to Computer Name</summary>
			/// <remarks></remarks>
			lockComp=2,
			/// <summary>Lock to Hard Drive Serial Number (Volume Serial Number)</summary>
			/// <remarks></remarks>
			lockHD=4,
			/// <summary>Lock to Hard Disk Firmware Serial (HDD Manufacturer's Serial Number)</summary>
			/// <remarks></remarks>
			lockHDFirmware=8,
			/// <summary>Lock to Windows Serial Number</summary>
			/// <remarks></remarks>
			lockWindows=16,
			/// <summary>Lock to BIOS Serial Number</summary>
			/// <remarks></remarks>
			lockBIOS=32,
			/// <summary>Lock to Motherboard Serial Number</summary>
			/// <remarks></remarks>
			lockMotherboard=64,
			/// <summary>Lock to Computer Local IP Address</summary>
			/// <remarks></remarks>
			lockIP=128,
			/// <summary>Lock to External IP Address</summary>
			/// <remarks></remarks>
			lockExternalIP=256,
			/// <summary>Lock to Fingerprint (Activelock Combination)</summary>
			/// <remarks></remarks>
			lockFingerprint=512,
			/// <summary>Lock to Memory ID</summary>
			/// <remarks></remarks>
			lockMemory=1024,
			/// <summary>Lock to CPU ID</summary>
			/// <remarks></remarks>
			lockCPUID=2048,
			/// <summary>Lock to Baseboard Name and Serial Number</summary>
			/// <remarks></remarks>
			lockBaseboardID=4096,
			/// <summary>Lock to Video Controller Name and Drive Version Number</summary>
			/// <remarks></remarks>
			lockvideoID=8192
		}

		/// <summary>License Key Type specifies the length/type</summary>
		/// <remarks></remarks>
		public enum ALLicenseKeyTypes
		{
			/// <summary>1024-bit Long keys by RSA via ALCrypto DLL</summary>
			/// <remarks></remarks>
			alsRSA=0,
			/// <summary>Short license keys by MD5</summary>
			/// <remarks></remarks>
			alsShortKeyMD5=1
		}

		/// <summary>License Store Type specifies where to store license keys</summary>
		/// <remarks></remarks>
		public enum LicStoreType
		{
			/// <summary>Store in Windows Registry</summary>
			/// <remarks></remarks>
			alsRegistry=0,
			/// <summary>Store in a license file</summary>
			/// <remarks></remarks>
			alsFile=1
		}

		/// <summary>Products Store Type specifies where to store products infos</summary>
		/// <remarks></remarks>
		public enum ProductsStoreType
		{
			/// <summary>Store in INI file (licenses.ini)</summary>
			/// <remarks></remarks>
			alsINIFile=0,
			/// <summary>Store in XML file (licenses.xml)</summary>
			/// <remarks></remarks>
			alsXMLFile=1,
			/// <summary>Store in MDB file (licenses.mdb)</summary>
			/// <remarks>mdb file should contain a table named products with structure: ID(autonumber), name(text,150), version (text,50), vccode(memo), gcode(memo)</remarks>
			alsMDBFile=2
			// TODO: IActiveLock.vb - Enum ProductStoreType - Store in MSSQL database 'not implemented
			///' <summary>TODO: Store in MSSQL database 'not implemented</summary>
			///' <remarks></remarks>
			//alsMSSQL = 3 'TODO
		}

		/// <summary>Trial Type specifies what kind of Trial Feature is used</summary>
		/// <remarks></remarks>
		public enum ALTrialTypes
		{
			/// <summary>No trial used</summary>
			/// <remarks></remarks>
			trialNone=0,
			/// <summary>Trial by Days</summary>
			/// <remarks></remarks>
			trialDays=1,
			/// <summary>Trial by Runs</summary>
			/// <remarks></remarks>
			trialRuns=2
		}

		/// <summary>Trial Hide Mode Type specifies what kind of Trial Hiding Mode is used</summary>
		/// <remarks>Values can be combined (OR'ed) together.</remarks>
		public enum ALTrialHideTypes
		{
			/// <summary>Trial information is hidden in BMP files</summary>
			/// <remarks></remarks>
			trialSteganography=1,
			/// <summary>Trial information is hidden in a folder which uses a default namespace</summary>
			/// <remarks></remarks>
			trialHiddenFolder=2,
			/// <summary>Trial information is encrypted and hidden in registry (per user)</summary>
			/// <remarks></remarks>
			trialRegistryPerUser=4,
			// TODO: IActiveLock.vb - Enum ALTrialHideYypes - Please update this comment
			/// <summary>Not documented! Please Update!</summary>
			/// <remarks></remarks>
			trialIsolatedStorage=8
		}

		/// <summary>Enum for accessing the Time Server to check Clock Tampering</summary>
		/// <remarks></remarks>
		public enum ALTimeServerTypes
		{
			/// <summary>Skips checking a Time Server</summary>
			/// <remarks></remarks>
			alsDontCheckTimeServer=0,
			/// <summary>Checks a Time Server</summary>
			/// <remarks></remarks>
			alsCheckTimeServer=1
		}

		/// <summary>Enum for scanning the system folders/files to detect clock tampering</summary>
		/// <remarks></remarks>
		public enum ALSystemFilesTypes
		{
			/// <summary>Skips checking system files</summary>
			/// <remarks></remarks>
			alsDontCheckSystemFiles=0,
			/// <summary>Checks system files</summary>
			/// <remarks></remarks>
			alsCheckSystemFiles=1
		}

		/// <summary>Enum for license file encryption</summary>
		/// <remarks></remarks>
		public enum ALLicenseFileTypes
		{
			/// <summary>Encrypts the license file</summary>
			/// <remarks></remarks>
			alsLicenseFilePlain=0,
			/// <summary>Leaves the license file readable</summary>
			/// <remarks></remarks>
			alsLicenseFileEncrypted=1
		}

		/// <summary>Enum for Auto Registeration via ALL files</summary>
		/// <remarks></remarks>
		public enum ALAutoRegisterTypes
		{
			/// <summary>Enables auto license registration</summary>
			/// <remarks></remarks>
			alsEnableAutoRegistration=0,
			/// <summary>Disables auto license registration</summary>
			/// <remarks></remarks>
			alsDisableAutoRegistration=1
		}

		/// <summary>Trial Warning can be persistent or temporary</summary>
		/// <remarks></remarks>
		public enum ALTrialWarningTypes
		{
			/// <summary>Trial Warning is Temporary (1-time only)</summary>
			/// <remarks></remarks>
			trialWarningTemporary=0,
			/// <summary>Trial Warning is Persistent</summary>
			/// <remarks></remarks>
			trialWarningPersistent=1
		}

		#endregion

#region "Public Properties"
// RE: Properties - there are some things not supported in C# that the VB code uses
//     see both VB and C# sections in -> ms-help://MS.VSCC.v90/MS.VSIPCC.v90/MS.W7SDK.1033/MS.W7SDKNET.1033/dv_fxdesignguide/html/652011f3-acfe-470e-bb58-7e5ef09d7374.htm
//

// the following region should be done!  Needs checking...
		#region "Read Only"

		/// <summary>RemainingTrialDays - Read Only - Returns the Number of Used Trial Days.</summary>
		/// <value></value>
		/// <returns>Integer - Number of Used Trial Days</returns>
		/// <remarks>None</remarks>
		public int RemainingTrialDays
		{
			//get { RemainingTrialDays = 0; }
			get { return 0; }
		}

		/// <summary>RemainingTrialRuns - Read Only - Returns the Number of Used Trial Runs.</summary>
		/// <value></value>
		/// <returns>Integer - Number of Used Trial Runs</returns>
		/// <remarks>None</remarks>
		public int RemainingTrialRuns
		{
			//get { RemainingTrialRuns = 0; }
			get { return 0; }
		}

		/// <summary>RegisteredLevel - Read Only - Returns the registered level.</summary>
		/// <value></value>
		/// <returns>String - Registered level</returns>
		/// <remarks>None</remarks>
		public string RegisteredLevel
		{
			//get { RegisteredLevel = string.Empty; }
			get { return string.Empty; }
		}

		/// <summary>MaxCount - Read Only - Returns the Number of concurrent users for the networked license</summary>
		/// <value></value>
		/// <returns>Integer - Number of concurrent users for the networked license</returns>
		/// <remarks>None</remarks>
		public int MaxCount
		{
			//get { MaxCount = 0; }
			get { return 0; }
		}

		/// <summary>LicenseClass - Read Only - Returns the LicenseClass</summary>
		/// <value></value>
		/// <returns>String - LicenseClass</returns>
		/// <remarks>None</remarks>
		public string LicenseClass
		{
			//get { LicenseClass = "Single"; }
			get { return "Single"; }
		}

		/// <summary>UsedLockType - Read Only - Returns the Current Lock Type being used in this instance.</summary>
		/// <value></value>
		/// <returns>ALLockTypes - lock type object corresponding to the current lock type(s) being used</returns>
		/// <remarks>None</remarks>
		public int UsedLockType
		{
			// I defaulted this to -1 : japreja - admin@whatsAvailable.org
			get { return -1; }
		}

		/// <summary>InstallationCode - Read Only - Returns the installation-specific code needed to obtain the liberation key.</summary>
		/// <param name="User">Optional - String - User</param>
		/// <param name="Lic">Optional - ProductLicense - License</param>
		/// <value>ByVal User As String - Optionally tailors the installation code specific to this user.</value>
		/// <returns>String - Installation Code</returns>
		/// <remarks></remarks>
		public string InstallationCode
		{
			//get { InstallationCode = string.Empty; }
			get { return string.Empty; }
		}

		/// <summary>EventNotifier - Read Only - Retrieves the event notifier.</summary>
		/// <value></value>
		/// <returns>ActiveLockEventNotifier - An object that can be used as a COM event source. i.e. can be used in <code>WithEvents</code> statements in VB.</returns>
		/// <remarks>
		/// Client applications uses this Notifier to handle event notifications sent by ActiveLock,
		/// including license property validation and encryption events.
		/// </remarks>
		public ActiveLockEventNotifier EventNotifier
		{
			//get { EventNotifier = null; }
			get { return null; }
		}

		/// <summary>UsedDays - Read Only - Returns the number of days this product has been used since its registration.</summary>
		/// <value></value>
		/// <returns>Long - Used days for the license</returns>
		/// <remarks>None</remarks>
		public int UsedDays
		{
			//get { UsedDays = 0; }
			get { return 0; }
		}

		/// <summary>RegisteredDate - Read Only - Retrieves the registration date.</summary>
		/// <value></value>
		/// <returns>String - Date on which the product is registered.</returns>
		/// <remarks>None</remarks>
		public string RegisteredDate
		{
			//get { RegisteredDate = string.Empty; }
			get { return string.Empty; }
		}

		/// <summary>RegisteredUser - Read Only - Returns the registered user.</summary>
		/// <value></value>
		/// <returns>String - Registered user name</returns>
		/// <remarks>None</remarks>
		public string RegisteredUser
		{
			//get { RegisteredUser = string.Empty; }
			get { return string.Empty; }
		}

		/// <summary>ExpirationDate - Read Only - Retrieves the expiration date.</summary>
		/// <value></value>
		/// <returns>String - Date on which the license will expire.</returns>
		/// <remarks>None</remarks>
		public string ExpirationDate
		{
			//get { ExpirationDate = string.Empty; }
			get { return string.Empty; }
		}

		#endregion

// these should be made into methods/functions
		#region "Write Only"

		/// <summary>
		/// LicenseKeyType - Write Only - Interface Property. Specifies the license key type for this instance of ActiveLock.
		/// </summary>
		/// <value>ByVal LicenseKeyTypes As ALLicenseKeyType - License Key Types object</value>
		/// <remarks>None</remarks>
		public ALLicenseKeyTypes LicenseKeyType
		{
			set { }
		}

		/// <summary>
		/// CheckTimeServerForClockTampering - Write Only - Specifies whether a Time Server should be used to check Clock Tampering
		/// </summary>
		/// <value>ByVal Value As ALTimeServerTypes - Flag to use a Time Server or not</value>
		/// <remarks></remarks>
		public ALTimeServerTypes CheckTimeServerForClockTampering
		{

			set { }
		}

		/// <summary>
		/// CheckSystemFilesForClockTampering - Write Only - Specifies whether the system files should be checked for Clock Tampering
		/// </summary>
		/// <value>ByVal Value As ALSystemFilesTypes - Flag to check system files or not</value>
		/// <remarks></remarks>
		public ALSystemFilesTypes CheckSystemFilesForClockTampering
		{

			set { }
		}

		/// <summary>
		/// AutoRegister - Write Only - Specifies whether the auto register mechanism via an ALL file should be enabled or disabled
		/// </summary>
		/// <value>ByVal Value As ALAutoRegisterTypes - Flag to auto register a license or not</value>
		/// <remarks></remarks>
		public ALAutoRegisterTypes AutoRegister
		{
			set { }
		}

		/// <summary>
		/// TrialWarning - Write Only - Specifies whether the Trial Warning is either Persistent or Temporary
		/// </summary>
		/// <value>ByVal Value As ALTrialWarningTypes - Trial Warning is either Persistent or Temporary</value>
		/// <remarks></remarks>
		public ALTrialWarningTypes TrialWarning
		{

			set { }
		}

		/// <summary>
		/// SoftwareCode - Write Only - Specifies the software code (product code)
		/// </summary>
		/// <value>ByVal sCode As String - Software Code</value>
		/// <remarks></remarks>
		public string SoftwareCode
		{

			set { }
		}

		/// <summary>
		/// KeyStoreType - Write Only - Specifies the key store type.
		/// </summary>
		/// <value>ByVal KeyStore As LicStoreType - Key store type</value>
		/// <remarks></remarks>
		public LicStoreType KeyStoreType
		{

			set { }
		}

		/// <summary>
		/// KeyStorePath - Write Only - Specifies the key store path.
		/// </summary>
		/// <value>ByVal sPath As String - The path to be used for the specified KeyStoreType.</value>
		/// <remarks>
		/// <para>@param sPath - The path to be used for the specified KeyStoreType.</para>
		/// <para>e.g. If <a href="IActiveLock.LicStoreType.html">alsFile</a> is used for <a href="IActiveLock.Let.KeyStoreType.html">KeyStoreType</a>,</para>
		/// <para>then <code>Path</code> specifies the path to the license file.</para>
		/// <para>If <a href="IActiveLock.LicStoreType.html">alsRegistry</a> is used,</para>
		/// <para>the Path specifies the Registry hive where license information is stored.</para>
		/// </remarks>
		public string KeyStorePath
		{

			set { }
		}

		/// <summary>
		/// AutoRegisterKeyPath - Write Only - Specifies the file path that contains the liberation key.
		/// </summary>
		/// <value>ByVal sPath As String - Full path to where the liberation file may reside.</value>
		/// <remarks>
		/// <para>If this file exists, ActiveLock will attempt to register the key automatically during its initialization.</para>
		/// <para>Upon successful registration, the liberation file WILL be deleted.</para>
		/// <para><b>Note</b>: This property is only effective if it is set prior to calling <code>Init</code>.</para>
		/// </remarks>
		public string AutoRegisterKeyPath
		{

			set { }
		}

		#endregion	// end region "Write Only"
//
		#region "Read Write"

		/// <summary>
		/// LockType - Read/Write - Returns the Lock Type being used in this instance.
		/// </summary>
		/// <value>See TODO:</value>
		/// <returns>ALLockTypes - lock type object corresponding to the lock type(s) being used</returns>
		/// <remarks></remarks>
		public ALLockTypes LockType
		{

			get { return LockType; }

			set { LockType = value; }
		}

		/// <summary>
		/// TrialHideType - Read/Write - Returns the Trial Hide Type being used in this instance.
		/// </summary>
		/// <value></value>
		/// <returns>ALTrialHideTypes - trial hide type object corresponding to the trial hide type(s) being used</returns>
		/// <remarks></remarks>
		public ALTrialHideTypes TrialHideType
		{

			get { return TrialHideType; }

			set { TrialHideType = value; }
		}

		/// <summary>
		/// TrialType - Read/Write - Returns the Trial Type being used in this instance.
		/// </summary>
		/// <value></value>
		/// <returns>ALTrialTypes - Trial Type (TrialNone, TrialByDays, TrialByRuns)</returns>
		/// <remarks></remarks>
		public ALTrialTypes TrialType
		{

			get { return TrialType; }

			set { TrialType = value; }
		}

		/// <summary>
		/// TrialLength - Read/Write - Returns the Trial Length being used in this instance.
		/// </summary>
		/// <value></value>
		/// <returns>Integer - Trial Length (Number of Days or Runs)</returns>
		/// <remarks></remarks>
		public int TrialLength
		{

			get { return TrialLength; }

			set { TrialLength = value; }
		}

		/// <summary>
		/// SoftwareName - Read/Write - Returns the Software Name being used in this instance.
		/// </summary>
		/// <value></value>
		/// <returns>String - Software Name</returns>
		/// <remarks></remarks>
		public string SoftwareName
		{
			//get { SoftwareName = string.Empty; }
			get { return SoftwareName;}
			set { }
		}

		/// <summary>
		/// SoftwarePassword - Read/Write - Returns the Software Password being used in this instance.
		/// </summary>
		/// <value></value>
		/// <returns>String - Software Password</returns>
		/// <remarks></remarks>
		public string SoftwarePassword
		{
			get { return SoftwarePassword; }

			set { }
		}

		/// <summary>
		/// LicenseFileType - Read/Write - Specifies whether the system files should be checked for Clock Tampering
		/// </summary>
		/// <value>ByVal Value As ALLicenseFileTypes - Encrypt License File or Leave it Plain</value>
		/// <returns></returns>
		/// <remarks></remarks>
		public ALLicenseFileTypes LicenseFileType
		{

			//LicenseFileType = String.Empty
			get { return LicenseFileType; }
			set { }
		}

		/// <summary>
		/// SoftwareVersion - Read/Write - Returns the Software Version being used in this instance.
		/// </summary>
		/// <value></value>
		/// <returns>String - Software Version</returns>
		/// <remarks></remarks>
		public string SoftwareVersion
		{
			get { return SoftwareVersion; }

			set { }
		}

		#endregion

#endregion // end region "Public Properties"

		#region "Functions"

		/// <summary>
		/// <para>LockCode - Interface Method. Computes a lock code corresponding to the specified Lock Types, License Class, etc.</para>
		/// <para>Optionally, if a product license is specified, then a lock string specific to that license is returned.</para>
		/// </summary>
		/// <param name="Lic">Optional - ByRef Lic As ProductLicense - Product License for which to compute the lock code.</param>
		/// <returns>String - Lock code</returns>
		/// <remarks></remarks>
		public string LockCode(
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute(null)
			] ref ProductLicense Lic)
		{
			return string.Empty;
		}

		// TODO: IActiveLock.vb - Function Transfer - Not Implimented?
		/// <summary>
		/// Transfer - Not Implimented? - Transfers the current license to another computer.
		/// </summary>
		/// <param name="InstallCode">ByVal InstallCode As String - Installation Code generated from the other computer.</param>
		/// <returns>String - The liberation key tailored for the request code generated from the other machine.</returns>
		/// <remarks></remarks>
		public string Transfer(string InstallCode)
		{
			return string.Empty;
		}

		// TODO: IActiveLock.vb - Function GenerateShortSerial - Update Comment
		/// <summary>
		/// GenerateShortSerial ? Undocumented...
		/// </summary>
		/// <param name="HDDfirmwareSerial"></param>
		/// <returns></returns>
		/// <remarks></remarks>
		public string GenerateShortSerial(string HDDfirmwareSerial)
		{
			return null;
		}

		// TODO: IActiveLock.vb - Function GenerateShortKey - Update Comment
		/// <summary>
		/// GenerateShortKey ? Undocumented...
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

		public string GenerateShortKey(
			string SoftwareCode,
			string SerialNumber,
			string LicenseeAndRegisteredLevel,
			string Expiration,
			ProductLicense.ALLicType LicType,
			int RegisteredLevel,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute(1)
			] ref short MaxUsers)
		{
			return null;
		}

		#endregion   // end region "Functions"

		#region "Methods"

		/// <summary>
		/// Register - Registers the product using the specified liberation key.
		/// </summary>
		/// <param name="LibKey">ByVal LibKey As String - Liberation key</param>
		/// <param name="user">Optional - String - User</param>
		/// <remarks></remarks>
		public void Register(
			string LibKey,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string user)
		{
		} // end Register

		/// <summary>
		/// Init - Purpose: Initializes ActiveLock before use. Some of the routines, including <a href="IActiveLock.Acquire.html">Acquire()</a> and <a href="IActiveLock.Register.html">Register()</a> requires <code>Init()</code> to be called first.
		/// </summary>
		/// <param name="strPath">Optional - ?Undocumented!</param>
		/// <param name="autoLicString">Optional - autoLicString As String - license key if autoregister is successful</param>
		/// <remarks></remarks>
		public void Init(
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strPath,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string autoLicString)
		{
		} // end Init

		/// <summary>
		/// <para>Acquires a valid license token.</para>
		/// <para>If no valid license can be found, an appropriate error will be raised, specifying the cause.</para>
		/// </summary>
		/// <param name="strMsg">Optional - ByRef strMsg As String - String returned by Activelock</param>
		/// <param name="strRemainingTrialDays">Optional - ?Undocumented!</param>
		/// <param name="strRemainingTrialRuns">Optional - ?Undocumented!</param>
		/// <param name="strTrialLength">Optional - ?Undocumented!</param>
		/// <param name="strUsedDays">Optional - ?Undocumented!</param>
		/// <param name="strExpirationDate">Optional - ?Undocumented!</param>
		/// <param name="strRegisteredUser">Optional - ?Undocumented!</param>
		/// <param name="strRegisteredLevel">Optional - ?Undocumented!</param>
		/// <param name="strLicenseClass">Optional - ?Undocumented!</param>
		/// <param name="strMaxCount">Optional - ?Undocumented!</param>
		/// <param name="strLicenseFileType">Optional - ?Undocumented!</param>
		/// <param name="strLicenseType">Optional - ?Undocumented!</param>
		/// <param name="strUsedLockType">Optional - ?Undocumented!</param>
		/// <remarks></remarks>
		public void Acquire(
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strMsg,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strRemainingTrialDays,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strRemainingTrialRuns,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strTrialLength,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strUsedDays,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strExpirationDate,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strRegisteredUser,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strRegisteredLevel,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strLicenseClass,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strMaxCount,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strLicenseFileType,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strLicenseType,
			[
				System.Runtime.InteropServices.OptionalAttribute,
				System.Runtime.InteropServices.DefaultParameterValueAttribute("")
			] ref string strUsedLockType)
		{
		} // end Acquire

		/// <summary>
		/// ResetTrial - Resets the Trial Mode
		/// </summary>
		/// <remarks></remarks>
		public void ResetTrial()
		{
		} // end ResetTrial

		/// <summary>
		/// KillTrial - Kills the Trial Mode
		/// </summary>
		/// <remarks></remarks>

		public void KillTrial()
		{
		} // end KillTrial

		#endregion   // end region "Methods"
	} // end class IActiveLock
