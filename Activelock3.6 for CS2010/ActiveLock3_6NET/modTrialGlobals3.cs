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
using System.Security.Cryptography;
using System.Text;
using System.Runtime.InteropServices;
using ActiveLock3_6NET;
static class modTrial
{
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
	//===============================================================================
	// Name: modTrialGlobals
	// Purpose: This module is used by the Trial Period/Runs feature
	// Functions:
	// Properties:
	// Methods:
	// Started: 08.15.2005
	// Modified: 03.25.2006
	//===============================================================================

	public static string EXPIRED_RUNS;
	public static string EXPIRED_DAYS;
	public static string LICENSE_SOFTWARE_NAME;
	public static string LICENSE_SOFTWARE_PASSWORD;
	public static string LICENSE_SOFTWARE_CODE;
	public static string LICENSE_SOFTWARE_VERSION;
	public static string LICENSE_SOFTWARE_LOCKTYPE;
	public static short alockDays;
	public static short alockRuns;
	public static bool trialPeriod;
	public static bool trialRuns;
	public static string TEXTMSG_RUNS;
	public static string TEXTMSG_DAYS;
	public static string TEXTMSG;
	public static string VIDEO;

	public static string OTHERFILE;

	public static string StegInfo;
	public const string HIDDENFOLDER = "SM8YnnHzkjsvBayVJjIexcUpH5+7aO1WosnkqOTm8ZU=";
	public const string EXPIREDDAYS = "ExpiredDays";
	public const string ACTIVELOCKSTRING = "Activelock3";
	public const string INITIALDATE = "01/01/2000";
	public const string TRIALWARNING = "Trial Warning";
	public const string EXPIREDWARNING = "Expired Warning";
	public const string CLSIDSTR = "{645FF040-5081-101B-9F08-00AA002F954E}";

	public const string CHANNELS = "CTDChannels_Version.";
	public const short REG_MULTI_SZ = 7;

	public const short ERROR_MORE_DATA = 234;
	// Windows Security Messages

	const int KEY_ALL_CLASSES = 0xf0063;
	const short ERROR_SUCCESS = 0;
	[DllImport("kernel32", EntryPoint = "lstrlenA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int lstrlen(string lpString);
	[DllImport("advapi32.dll", EntryPoint = "RegSetValueExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegSetValueExString(int hKey, string lpValueName, int Reserved, int dwType, string lpValue, int cbData);
	[DllImport("advapi32.dll", EntryPoint = "RegSetValueExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegSetValueExLong(int hKey, string lpValueName, int Reserved, int dwType, ref int lpValue, int cbData);
	[DllImport("advapi32.dll", EntryPoint = "GetUserNameA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int GetUserName(string lpBuffer, ref int nSize);
	[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegCloseKey(int hKey);
	[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegOpenKey(int hKey, string lpSubKey, ref int phkResult);
	[DllImport("advapi32.dll", EntryPoint = "RegOpenKeyExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegOpenKeyEx(int hKey, string lpSubKey, int ulOptions, int samDesired, ref int phkResult);
	[DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegQueryValueEx(int hKey, string lpValueName, int lpReserved, ref int lpType, ref int lpData, ref int lpcbData);
	[DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegQueryValueExString(int hKey, string lpValueName, int lpReserved, ref int lpType, string lpData, ref int lpcbData);
	[DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegQueryValueExLong(int hKey, string lpValueName, int lpReserved, ref int lpType, ref int lpData, ref int lpcbData);
	[DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegQueryValueExNULL(int hKey, string lpValueName, int lpReserved, ref int lpType, int lpData, ref int lpcbData);
	[DllImport("kernel32", EntryPoint = "ExpandEnvironmentStringsA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int ExpandEnvironmentStrings(string lpSrc, string lpDst, int nSize);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int GlobalLock(int hMem);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int GlobalUnlock(int hMem);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int GlobalAlloc(int wFlags, int dwBytes);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int GlobalFree(int hMem);
	[DllImport("kernel32", EntryPoint = "FindFirstFileA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int FindFirstFile(string lpFileName, ref WIN32_FIND_DATA lpFindFileData);
	[DllImport("kernel32", EntryPoint = "FindNextFileA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int FindNextFile(int hFindFile, ref WIN32_FIND_DATA lpFindFileData);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int FindClose(int hFindFile);
	[DllImport("kernel32", EntryPoint = "SearchPathA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int SearchPath(string lpPath, string lpFileName, string lpExtension, int nBufferLength, string lpBuffer, string lpFilePart);

	// Windows Registry API calls





	private const short MAX_PATH = 260;


	private struct FILETIME
	{
		public int lngLowDateTime;
		public int lngHighDateTime;
	}

	private struct WIN32_FIND_DATA
	{
			// File attributes
		public int lngFileAttributes;
			// Creation time
		public FILETIME ftCreationTime;
			// Last access time
		public FILETIME ftLastAccessTime;
			// Last modified time
		public FILETIME ftLastWriteTime;
			// Size (high word)
		public int lngFileSizeHigh;
			// Size (low word)
		public int lngFileSizeLow;
			// reserved
		public int lngReserved0;
			// reserved
		public int lngReserved1;
		[VBFixedString(modTrial.MAX_PATH), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = modTrial.MAX_PATH)]
			// File name
		public string strFilename;
		[VBFixedString(14), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = 14)]
			// 8.3 name
		public string strAlternate;
	}

	private struct SECURITY_ATTRIBUTES
	{
		public int nLength;
		public object lpSecurityDescriptor;
		public bool bInheritHandle;
	}
	[DllImport("kernel32", EntryPoint = "GetWindowsDirectoryA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetWindowsDirectory(string lpBuffer, int nSize);
	[DllImport("kernel32", EntryPoint = "GetVolumeInformationA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetVolumeInformation(string lpRootPathName, string lpVolumeNameBuffer, int nVolumeNameSize, ref int lpVolumeSerialNumber, ref int lpMaximumComponentLength, ref int lpFileSystemFlags, string lpFileSystemNameBuffer, int nFileSystemNameSize);


	//Some common variables
	public static string sWinPath;
	public static string sLocalKey;

	public static string sMainFile;
	//Error variables
	public static bool bTampered;
	public static bool bAccessDenied;
	public static bool bFatal;

	public static short intProgress;
	[DllImport("kernel32", EntryPoint = "CreateFileA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int CreateFile(string lpFileName, int dwDesiredAccess, int dwShareMode, int lpSecurityAttributes, int dwCreationDisposition, int dwFlagsAndAttributes, int hTemplateFile);
	[DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetWindowDC(int hwnd);
	[DllImport("user32", EntryPoint = "GetClassNameA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetClassName(int hwnd, string lpClassName, int nMaxCount);


	//Spy Scan stuff

	public const short CREATE_NEW = 1;
	public const short CREATE_ALWAYS = 2;
	public const short OPEN_EXISTING = 3;
	public const short OPEN_ALWAYS = 4;

	public const short TRUNCATE_EXISTING = 5;
	private const short FILE_BEGIN = 0;
	private const short FILE_CURRENT = 1;

	private const short FILE_END = 2;
	private const uint FILE_FLAG_WRITE_THROUGH = 0x80000000;
	private const uint FILE_FLAG_OVERLAPPED = 0x40000000;
	private const uint FILE_FLAG_NO_BUFFERING = 0x20000000;
	private const uint FILE_FLAG_RANDOM_ACCESS = 0x10000000;
	private const uint FILE_FLAG_SEQUENTIAL_SCAN = 0x8000000;
	private const uint FILE_FLAG_DELETE_ON_CLOSE = 0x4000000;
	private const uint FILE_FLAG_BACKUP_SEMANTICS = 0x2000000;

	private const int FILE_FLAG_POSIX_SEMANTICS = 0x1000000;
	[DllImport("user32", EntryPoint = "FindWindowA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int FinWin(string lpClassName, string lpWindowName);
	[DllImport("kernel32", EntryPoint = "CreateFileA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int CF(string lpFileName, int dwDesiredAccess, int dwShareMode, ref int lpSecurityAttributes, int dwCreationDisposition, int dwFlagsAndAttributes, int hTemplateFile);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int CloseHandle(int hObject);
	[DllImport("user32.dll", EntryPoint = "FindWindowA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int FindWindow(string lpClassName, string lpWindowName);
	[DllImport("user32", EntryPoint = "PostMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int PostMessage(int hwnd, int wMsg, int wParam, ref int lParam);
	[DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int GetWindowThreadProcessId(int hwnd, ref int lpdwProcessId);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int OpenProcess(int dwDesiredAccess, int bInheritHandle, int dwProcessID);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int WriteProcessMemory(int hProcess, long lpBaseAddress, int lpBuffer, int nSize, ref int lpNumberOfBytesWritten);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int ReadProcessMemory(int hProcess, long lpBaseAddress, long lpBuffer, int nSize, ref int lpNumberOfBytesWritten);
	[DllImport("kernel32", EntryPoint = "FormatMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int FormatMessage(int dwFlags, ref long lpSource, int dwMessageId, int dwLanguageId, string lpBuffer, int nSize, ref int Arguments);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int GetLastError();
	// Called from frmC

	public const short FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x100;

	public const short FORMAT_MESSAGE_FROM_SYSTEM = 0x1000;
	[DllImport("kernel32", EntryPoint = "Process32First", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int ProcessFirst(int hSnapshot, ref PROCESSENTRY32 uProcess);
	[DllImport("kernel32", EntryPoint = "Process32Next", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int ProcessNext(int hSnapshot, ref PROCESSENTRY32 uProcess);
	[DllImport("kernel32", EntryPoint = "CreateToolhelp32Snapshot", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int CreateToolhelpSnapshot(int lFlags, ref int lProcessID);
	//Structure PROCESSENTRY32 may require marshalling attributes to be passed as an argument in this Declare statement
	//Structure PROCESSENTRY32 may require marshalling attributes to be passed as an argument in this Declare statement


	public const int TH32CS_SNAPPROCESS = 2;
	public struct PROCESSENTRY32
	{
		public int dwSize;
		public int cntUsage;
		public int th32ProcessID;
		public int th32DefaultHeapID;
		public int th32ModuleID;
		public int cntThreads;
		public int th32ParentProcessID;
		public int pcPriClassBase;
		public int dwFlags;
		[VBFixedString(260), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = 260)]
		public string szexeFile;
	}

	public const int GENERIC_WRITE = 0x40000000;
	public const uint GENERIC_READ = 0x80000000;
	public const int GENERIC_EXECUTE = 0x20000000;
	public const int GENERIC_ALL = 0x10000000;
	public const short FILE_SHARE_READ = 0x1;
	public const short FILE_SHARE_WRITE = 0x2;
	public const int FILE_SHARE_DELETE = 0x4;
	public const short FILE_ATTRIBUTE_NORMAL = 0x80;
	public const short FILE_ATTRIBUTE_READONLY = 0x1;
	public const short FILE_ATTRIBUTE_HIDDEN = 0x2;
	public const short FILE_ATTRIBUTE_SYSTEM = 0x4;
	public const short FILE_ATTRIBUTE_DIRECTORY = 0x10;
	public const short FILE_ATTRIBUTE_ARCHIVE = 0x20;
	public const short FILE_ATTRIBUTE_TEMPORARY = 0x100;

	public const short FILE_ATTRIBUTE_COMPRESSED = 0x800;

	public const uint EAV = 0xc0000005;
	public static string[] ProcessName = new string[257];
	public static int[] ProcessID = new int[257];
	public static int retVal;
	public static int hFile;
	public static int TimerStart;
	public static int wX;
	public static int wY;
	public static object myHandle;
	//Public Buffer As String
	public static object varchk;
	public static string[] encvar = new string[4001];

	public static bool HAD2HAMMER;
	// Folder Date Stamp API
	private struct SYSTEMTIME
	{
		public short wYear;
		public short wMonth;
		public short wDayOfWeek;
		public short wDay;
		public short wHour;
		public short wMinute;
		public short wSecond;
		public int wMilliseconds;
	}
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetFileTime(int hFile, ref FILETIME lpCreationTime, ref FILETIME lpLastAccessTime, ref FILETIME lpLastWriteTime);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int FileTimeToLocalFileTime(ref FILETIME lpFileTime, ref FILETIME lpLocalFileTime);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int FileTimeToSystemTime(ref FILETIME lpFileTime, ref SYSTEMTIME lpSystemTime);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int SystemTimeToFileTime(ref SYSTEMTIME lpSystemTime, ref FILETIME lpFileTime);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int LocalFileTimeToFileTime(ref FILETIME lpLocalFileTime, ref FILETIME lpFileTime);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int SetFileTime(int hFile, ref FILETIME lpCreationTime, ref FILETIME lpLastAccessTime, ref FILETIME lpLastWriteTime);

	//Structure FILETIME may require marshalling attributes to be passed as an argument in this Declare statement
	//Structure SYSTEMTIME may require marshalling attributes to be passed as an argument in this Declare statement

	//Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Integer, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Integer) As Integer
	//Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hInternetSession As Integer, ByVal sUrl As String, ByVal sHeaders As String, ByVal lHeadersLength As Integer, ByVal lFlags As Integer, ByVal lContext As Integer) As Integer
	//Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Integer, ByVal sBuffer As String, ByVal lNumberOfBytesToRead As Integer, ByRef lNumberOfBytesRead As Integer) As Integer
	//Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Integer) As Integer

	[DllImport("WinInet.dll", EntryPoint = "InternetOpenA", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
	public static IntPtr InternetOpen(string agent, Int32 accessType, string proxyName, string proxyBypass, Int32 flags)
	{
	}


	[DllImport("WinInet.dll", EntryPoint = "InternetOpenUrlA", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
	public static Int32 InternetOpenUrl(IntPtr session, string url, string header, Int32 headerLength, Int32 flags, Int32 context)
	{
	}


	//InternetReadFile
	[DllImport("WinInet.dll", EntryPoint = "InternetReadFile", CharSet = CharSet.Auto, SetLastError = true)]
	public static Int32 InternetReadFile(Int32 handle, 	[MarshalAs(UnmanagedType.LPArray)]
byte[] newBuffer, Int32 bufferLength, ref Int32 bytesRead)
	{
	}

	[DllImport("WinInet.dll", EntryPoint = "InternetCloseHandle", CharSet = CharSet.Ansi, ExactSpelling = true, SetLastError = true)]
	public static Int32 InternetCloseHandle(Int32 hInternet)
	{
	}


	public static string OpenURL(string sUrl)
	{
		string functionReturnValue = null;
		const int INTERNET_ACCESS_TYPE_DIRECT = 1;
		const int INTERNET_FLAG_RELOAD = 0x80000000;
		const string USER_AGENT = "IE";

		Int32 length = 2048;
		IntPtr handle = default(IntPtr);
		Int32 session = default(Int32);
		string header = "Accept: */*" + ControlChars.Cr + ControlChars.Cr;
		byte[] newBuffer = null;
		Int32 bytesRead = default(Int32);
		Int32 response = default(Int32);
		int context = 0;
		int flags = 0;
		string result = null;
		handle = InternetOpen(USER_AGENT, INTERNET_ACCESS_TYPE_DIRECT, Constants.vbNullString, Constants.vbNullString, flags);
		session = InternetOpenUrl(handle, sUrl, header, header.Length, INTERNET_FLAG_RELOAD, context);
		if (session == 0) {
			result = "Error: " + Marshal.GetLastWin32Error();
		}
		else {
			newBuffer = new byte[length];
			response = InternetReadFile(session, newBuffer, length, ref bytesRead);
			if (response == 0) {
				result = "";
				//"Error Reading File: " & Marshal.GetLastWin32Error()
			}
			else {
				// Use appropriate Encoding here to get string from byte array
				result = System.Text.UTF8Encoding.UTF8.GetString(newBuffer);
			}
		}
		functionReturnValue = result;
		InternetCloseHandle(session);
		InternetCloseHandle(handle.ToInt32());
		return result;
		return functionReturnValue;

	}
	//===============================================================================
	// Name: Function DateGoodRegistry
	// Input:
	//   ByRef numDays As Integer - Number of trial days the application is good for
	//   ByRef daysLeft As Integer - Days left in the trial period
	// Output:
	//   Boolean - Returns True if the date is good
	// Purpose: The purpose of this module is to allow you to place a time
	// limit on the unregistered use of your shareware application.
	//       <p>Example:
	//       <br>If DateGoodRegistry(30)=False Then
	//       <br>CrippleApplication
	//       <br>End if
	// Remarks: This module can not be defeated by rolling back the system clock.
	//===============================================================================
	public static bool DateGoodRegistry(ref short numDays, ref short daysLeft)
	{
		bool functionReturnValue = false;
		//Ex: If DateGoodRegistry(30)=False Then
		// CrippleApplication
		// End if
		//Registry Parameters:
		// CRD: Current Run Date
		// LRD: Last Run Date
		// FRD: First Run Date
		System.DateTime TmpCRD = default(System.DateTime);
		System.DateTime TmpLRD = default(System.DateTime);
		System.DateTime TmpFRD = default(System.DateTime);

		 // ERROR: Not supported in C#: OnErrorStatement


		TmpCRD = ActiveLockDate(System.DateTime.UtcNow);
		TmpLRD = (System.DateTime)dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"));
		//1/1/2000
		TmpFRD = (System.DateTime)dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"));
		//1/1/2000
		functionReturnValue = false;

		//If this is the applications first load, write initial settings
		//to the registry
		//1/1/2000
		if (TmpLRD == (System.DateTime)dec2("93.8D.93.8D.96.90.90.90")) {
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2((string)TmpCRD));
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor2", enc2((string)TmpCRD));
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2("0"));
		}
		//Read LRD and FRD from registry
		TmpLRD = (System.DateTime)dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"));
		//1/1/2000
		TmpFRD = (System.DateTime)dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"));
		//1/1/2000

		//System clock rolled back
		if (ActiveLockDate(TmpFRD) > ActiveLockDate(TmpCRD)) {
			functionReturnValue = false;
		}
		//trial expired
		else if (ActiveLockDate(System.DateTime.UtcNow) > ActiveLockDate(TmpFRD).AddDays(numDays)) {
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + EXPIRED_DAYS + "_" + EXPIRED_DAYS + "_" + EXPIRED_RUNS));
			functionReturnValue = false;
		}
		//Everything OK write New LRD date
		else if (ActiveLockDate(TmpCRD) > ActiveLockDate(TmpLRD)) {
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2((string)TmpCRD));
			functionReturnValue = true;
		}
		else if (ActiveLockDate(TmpCRD) == ActiveLockDate(TmpLRD)) {
			functionReturnValue = true;
		}
		else {
			functionReturnValue = false;
		}
		if (functionReturnValue) {
			daysLeft = numDays - ActiveLockDate(System.DateTime.UtcNow).Subtract(ActiveLockDate(TmpFRD)).Days;
		}
		else {
			daysLeft = 0;
		}
		return;
		DateGoodRegistryError:
		daysLeft = 0;
		functionReturnValue = false;
		return;
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function RunsGoodRegistry
	// Input:
	//   ByRef numRuns As Integer - Number of trial days the application is good for
	//   ByRef runsLeft As Integer - Days left in the trial period
	// Output:
	//   Boolean - Returns True if the Runs is good
	// Purpose: The purpose of this module is to allow you to place a time
	// limit on the unregistered use of your shareware application.
	// Remarks: None
	//===============================================================================
	public static bool RunsGoodRegistry(ref short numRuns, ref short runsLeft)
	{
		bool functionReturnValue = false;
		 // ERROR: Not supported in C#: OnErrorStatement

		System.DateTime TmpCRD = default(System.DateTime);
		System.DateTime TmpLRD = default(System.DateTime);
		System.DateTime TmpFRD = default(System.DateTime);

		TmpCRD = (System.DateTime)INITIALDATE;
		TmpLRD = (System.DateTime)dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"));
		//1/1/2000
		TmpFRD = (System.DateTime)dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"));
		//1/1/2000
		runsLeft = Conversion.Int((double)dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2((string)numRuns - 1)))) - 1;
		functionReturnValue = false;

		//If this is the applications first load, write initial settings
		//to the registry
		//1/1/2000
		if (TmpLRD == (System.DateTime)dec2("93.8D.93.8D.96.90.90.90")) {
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2((string)DateTime.Now));
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor2", enc2((string)TmpCRD));
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2((string)numRuns - 1));
		}
		//Read LRD and FRD from registry
		TmpLRD = (System.DateTime)dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"));
		//1/1/2000
		TmpFRD = (System.DateTime)dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor2", "93.8D.93.8D.96.90.90.90"));
		//1/1/2000
		runsLeft = Conversion.Int((double)dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2((string)numRuns - 1))));
		if (TmpLRD == "#12:00:00 AM#") {
			TmpLRD = (System.DateTime)INITIALDATE;
			TmpFRD = (System.DateTime)INITIALDATE;
			runsLeft = numRuns - 1;
		}

		//impossible
		if (runsLeft < 0) {
			functionReturnValue = false;
		}
		//Trial runs expired
		else if (runsLeft > numRuns) {
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + EXPIRED_DAYS + "_" + EXPIRED_DAYS + "_" + EXPIRED_RUNS));
			functionReturnValue = false;
		}
		//Everything OK write the remaining number of runs
		else if (numRuns >= runsLeft) {
			if (TmpLRD == (System.DateTime)INITIALDATE) {
				Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2((string)numRuns - 1));
			}
			else {
				Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor3", enc2((string)runsLeft - 1));
			}
			functionReturnValue = true;
		}
		else {
			functionReturnValue = false;
		}
		if (functionReturnValue) {
		}
		else {
			runsLeft = 0;
		}
		return;
		RunsGoodRegistryError:
		runsLeft = 0;
		functionReturnValue = false;
		return;
		return functionReturnValue;
	}

	public static bool DateGood(short numDays, ref short daysLeft, ref IActiveLock.ALTrialHideTypes TrialHideTypes)
	{
		bool functionReturnValue = false;
		bool use2 = false;
		bool use3 = false;
		bool use4 = false;
		short daysLeft2 = 0;
		short daysLeft3 = 0;
		short daysLeft4 = 0;

		TEXTMSG_DAYS = DecryptString128Bit("sQvYYRLPon5IyH6BQRAUBuCLTq/5VkH3kl7HUwJLZ2M=", PSWD());
		functionReturnValue = false;

		if (TrialSteganographyExists(ref TrialHideTypes)) {
			if (DateGoodSteganography(ref numDays, ref daysLeft2) == false) {
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS);
				//MsgBox "DateGoodSteganography " & daysLeft2
				return;
			}
			use2 = true;
		}
		if (TrialHiddenFolderExists(ref TrialHideTypes)) {
			if (DateGoodHiddenFolder(ref numDays, ref daysLeft3) == false) {
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS);
				//MsgBox "DateGoodHiddenFolder " & daysLeft3
				return;
			}
			use3 = true;
		}
		if (TrialRegistryPerUserExists(ref TrialHideTypes)) {
			if (DateGoodRegistry(ref numDays, ref daysLeft4) == false) {
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS);
				//MsgBox "DateGoodRegistry " & daysLeft4
				return;
			}
			use4 = true;
		}
		//MsgBox "DateGoodSteganography " & daysLeft2
		//MsgBox "DateGoodHiddenFolder " & daysLeft3
		//MsgBox "DateGoodRegistry " & daysLeft4

		functionReturnValue = true;
		if ((use2 & !use3 & !use4)) {
			daysLeft = daysLeft2;
		}
		else if ((!use2 & use3 & !use4)) {
			daysLeft = daysLeft3;
		}
		else if ((!use2 & !use3 & use4)) {
			daysLeft = daysLeft4;
		}
		else if ((use2 & use3 & !use4) & daysLeft2 == daysLeft3) {
			daysLeft = daysLeft2;
		}
		else if ((use2 & !use3 & use4) & daysLeft2 == daysLeft4) {
			daysLeft = daysLeft2;
		}
		else if ((!use2 & use3 & use4) & daysLeft3 == daysLeft4) {
			daysLeft = daysLeft3;
		}
		else if ((use2 & use3 & use4) & daysLeft2 == daysLeft3 & daysLeft2 == daysLeft4) {
			daysLeft = daysLeft2;
		}
		else {
			functionReturnValue = false;
			daysLeft = 0;
		}
		return functionReturnValue;
	}
	public static bool RunsGood(short numRuns, ref short runsLeft, ref IActiveLock.ALTrialHideTypes TrialHideTypes)
	{
		bool functionReturnValue = false;
		bool use2 = false;
		bool use3 = false;
		bool use4 = false;
		short runsLeft2 = 0;
		short runsLeft3 = 0;
		short runsLeft4 = 0;
		TEXTMSG_RUNS = DecryptString128Bit("6urN2+xbgqbLLsOoC4hbGpLT3bnvY3YPGW299cOnqfo=", PSWD());

		functionReturnValue = false;

		if (TrialSteganographyExists(ref TrialHideTypes)) {
			if (RunsGoodSteganography(ref numRuns, ref runsLeft2) == false) {
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS);
				//MsgBox "RunsGoodSteganography " & runsLeft2
				return;
			}
			use2 = true;
		}

		if (TrialHiddenFolderExists(ref TrialHideTypes)) {
			if (RunsGoodHiddenFolder(ref numRuns, ref runsLeft3) == false) {
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS);
				//MsgBox "RunsGoodHiddenFolder " & runsLeft3
				return;
			}
			use3 = true;
		}

		if (TrialRegistryPerUserExists(ref TrialHideTypes)) {
			if (RunsGoodRegistry(ref numRuns, ref runsLeft4) == false) {
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS);
				//MsgBox "RunsGoodRegistry " & runsLeft4
				return;
			}
			use4 = true;
		}

		//MsgBox "RunsGoodSteganography " & runsLeft2
		//MsgBox "RunsGoodHiddenFolder " & runsLeft3
		//MsgBox "RunsGoodRegistry " & runsLeft4

		functionReturnValue = true;
		if ((use2 & !use3 & !use4)) {
			runsLeft = runsLeft2;
		}
		else if ((!use2 & use3 & !use4)) {
			runsLeft = runsLeft3;
		}
		else if ((!use2 & !use3 & use4)) {
			runsLeft = runsLeft4;
		}
		else if ((use2 & use3 & !use4) & runsLeft2 == runsLeft3) {
			runsLeft = runsLeft2;
		}
		else if ((use2 & !use3 & use4) & runsLeft2 == runsLeft4) {
			runsLeft = runsLeft2;
		}
		else if ((!use2 & use3 & use4) & runsLeft3 == runsLeft4) {
			runsLeft = runsLeft3;
		}
		else if ((use2 & use3 & use4) & runsLeft2 == runsLeft3 & runsLeft2 == runsLeft4) {
			runsLeft = runsLeft2;
		}
		else {
			functionReturnValue = false;
			runsLeft = 0;
		}
		return functionReturnValue;
	}

	public static bool DateGoodHiddenFolder(ref short numDays, ref short daysLeft)
	{
		bool functionReturnValue = false;
		//Hidden Folder Parameters:
		// CRD: Current Run Date
		// LRD: Last Run Date
		// FRD: First Run Date
		System.DateTime TmpCRD = default(System.DateTime);
		System.DateTime TmpLRD = default(System.DateTime);
		System.DateTime TmpFRD = default(System.DateTime);
		string strMyString = null;
		string strSource = null;
		short intFF = 0;
		double ok = 0;

		 // ERROR: Not supported in C#: OnErrorStatement

		if (Directory.Exists(modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(myDir(), PSWD())) == false) FileSystem.MkDir(modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(myDir(), PSWD())); 
		strSource = HiddenFolderFunction();

		System.IO.DirectoryInfo checkFile = null;
		string dirPath = null;
		dirPath = modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(HIDDENFOLDER, PSWD());
		checkFile = new System.IO.DirectoryInfo(dirPath);
		System.IO.FileAttributes attributeReader = default(System.IO.FileAttributes);
		attributeReader = checkFile.Attributes;

		if (Directory.Exists(dirPath) == true & (attributeReader & System.IO.FileAttributes.Directory & System.IO.FileAttributes.Hidden & System.IO.FileAttributes.ReadOnly & System.IO.FileAttributes.System) > 0) {
			MinusAttributes();
			//Check to see if our file is there
			if (fileExist(strSource)) {
				FileSystem.SetAttr(strSource, FileAttribute.Normal);
				//    Else
				//        ' User found the file and deleted it; expire
				//        PlusAttributes
				//        DateGoodHiddenFolder = False
				//        Exit Function
			}
		}
		else if (Directory.Exists(dirPath) == true) {
			//Ok, the folder is there with no hidden, system attributes
			//Check to see if our file is there
			if (fileExist(strSource)) {
				FileSystem.SetAttr(strSource, FileAttribute.Normal);
				//    Else
				//        ' User found the file and deleted it; expire
				//        PlusAttributes
				//        DateGoodHiddenFolder = False
				//        Exit Function
			}
		}
		else {
			FileSystem.MkDir(dirPath);
		}

		CreateHdnFile();

		TmpCRD = ActiveLockDate(System.DateTime.UtcNow);
		string[] a = null;
		string aa = null;

		if (fileExist(strSource)) {
			string strContents1 = null;
			StreamReader objReader1 = null;
			objReader1 = new StreamReader(strSource);
			strMyString = objReader1.ReadToEnd();
			objReader1.Close();
			if (Strings.Right(strMyString, 2) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 2); 
			if (Strings.Right(strMyString, 1) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 1); 
			strMyString = DecryptString128Bit(strMyString, PSWD());

			//' Read the file...
			//intFF = FreeFile()
			//FileOpen(intFF, strSource, OpenMode.Input)
			//strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
			//If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
			//If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
			//strMyString = DecryptString128Bit(strMyString, PSWD)
			//FileClose(intFF)

			if (!string.IsNullOrEmpty(strMyString)) {
				a = strMyString.Split("_");
				if (!string.IsNullOrEmpty(a[1])) TmpLRD = (System.DateTime)a[1]; 
				if (!string.IsNullOrEmpty(a[2])) TmpFRD = (System.DateTime)a[2]; 
				if (a[0] != LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD) {
					functionReturnValue = false;
					return;
				}
			}
			if (TmpLRD == "#12:00:00 AM#") TmpLRD = (System.DateTime)INITIALDATE; 
			if (TmpFRD == "#12:00:00 AM#") TmpFRD = (System.DateTime)INITIALDATE; 
		}
		else {
			TmpLRD = (System.DateTime)INITIALDATE;
			TmpFRD = (System.DateTime)INITIALDATE;
		}
		functionReturnValue = false;

		//If this is the applications first load, write initial settings
		//to Hidden Folder
		if (TmpLRD == (System.DateTime)INITIALDATE) {
			TmpLRD = TmpCRD;
			TmpFRD = TmpCRD;
			// Write to the file...
			intFF = FileSystem.FreeFile();
			FileSystem.FileOpen(intFF, strSource, OpenMode.Output);
			FileSystem.PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + TmpLRD + "_" + TmpFRD + "_" + "0", PSWD()));
			FileSystem.FileClose(intFF);
		}
		//Read LRD and FRD from Hidden Folder
		string[] b = null;

		string strContents = null;
		StreamReader objReader = null;
		objReader = new StreamReader(strSource);
		strMyString = objReader.ReadToEnd();
		objReader.Close();
		if (Strings.Right(strMyString, 2) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 2); 
		if (Strings.Right(strMyString, 1) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 1); 
		strMyString = DecryptString128Bit(strMyString, PSWD());

		//' Read the file...
		//intFF = FreeFile()
		//FileOpen(intFF, strSource, OpenMode.Input)
		//strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
		//If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
		//If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
		//strMyString = DecryptString128Bit(strMyString, PSWD)
		//FileClose(intFF)

		b = strMyString.Split("_");
		if (!string.IsNullOrEmpty(b[1])) TmpLRD = (System.DateTime)b[1]; 
		if (!string.IsNullOrEmpty(b[2])) TmpFRD = (System.DateTime)b[2]; 
		if (b[0] != LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD) {
			functionReturnValue = false;
			return;
		}
		if (TmpLRD == "#12:00:00 AM#") TmpLRD = (System.DateTime)INITIALDATE; 
		if (TmpFRD == "#12:00:00 AM#") TmpFRD = (System.DateTime)INITIALDATE; 

		//System clock rolled back
		if (ActiveLockDate(TmpFRD) > ActiveLockDate(TmpCRD)) {
			functionReturnValue = false;
		}
		//trial expired
		else if (ActiveLockDate(System.DateTime.UtcNow) > ActiveLockDate(TmpFRD).AddDays(numDays)) {
			// Write to the file...
			intFF = FileSystem.FreeFile();
			FileSystem.FileOpen(intFF, strSource, OpenMode.Output);
			FileSystem.PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + EXPIRED_DAYS + "_" + EXPIRED_DAYS + "_" + EXPIRED_RUNS, PSWD()));
			FileSystem.FileClose(intFF);
			functionReturnValue = false;
		}
		//Everything OK write New LRD date
		else if (ActiveLockDate(TmpCRD) > ActiveLockDate(TmpLRD)) {
			// Write to the file...
			intFF = FileSystem.FreeFile();
			FileSystem.FileOpen(intFF, strSource, OpenMode.Output);
			FileSystem.PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + TmpCRD + "_" + TmpFRD + "_" + "0", PSWD()));
			FileSystem.FileClose(intFF);
			functionReturnValue = true;
		}
		else if (ActiveLockDate(TmpCRD) == ActiveLockDate(TmpLRD)) {
			functionReturnValue = true;
		}
		else {
			functionReturnValue = false;
		}

		FileSystem.SetAttr(strSource, FileAttribute.ReadOnly + FileAttribute.Hidden + FileAttribute.System);
		PlusAttributes();

		if (functionReturnValue) {
			daysLeft = numDays - ActiveLockDate(System.DateTime.UtcNow).Subtract(ActiveLockDate(TmpFRD)).Days;
		}
		else {
			daysLeft = 0;
		}
		return;
		DateGoodHiddenFolderError:
		return functionReturnValue;


	}
	public static bool RunsGoodHiddenFolder(ref short numRuns, ref short runsLeft)
	{
		bool functionReturnValue = false;
		System.DateTime TmpCRD = default(System.DateTime);
		System.DateTime TmpLRD = default(System.DateTime);
		System.DateTime TmpFRD = default(System.DateTime);
		string strMyString = null;
		string strSource = null;
		short intFF = 0;
		double ok = 0;

		 // ERROR: Not supported in C#: OnErrorStatement

		if (Directory.Exists(modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(myDir(), PSWD())) == false) FileSystem.MkDir(modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(myDir(), PSWD())); 
		strSource = HiddenFolderFunction();

		System.IO.DirectoryInfo checkFile = null;
		string dirPath = null;
		dirPath = modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(HIDDENFOLDER, PSWD());
		checkFile = new System.IO.DirectoryInfo(dirPath);
		System.IO.FileAttributes attributeReader = default(System.IO.FileAttributes);
		attributeReader = checkFile.Attributes;

		if (Directory.Exists(dirPath) == true & (attributeReader & System.IO.FileAttributes.Directory & System.IO.FileAttributes.Hidden & System.IO.FileAttributes.ReadOnly & System.IO.FileAttributes.System) > 0) {
			MinusAttributes();
			//Check to see if our file is there
			if (fileExist(strSource)) {
				FileSystem.SetAttr(strSource, FileAttribute.Normal);
				//Else
				//    ' User found the file and deleted it; expire
				//    PlusAttributes()
				//    RunsGoodHiddenFolder = False
				//    Exit Function
			}
		}
		else if (Directory.Exists(dirPath) == true) {
			//Ok, the folder is there with no hidden, system attributes
			//Check to see if our file is there
			if (fileExist(strSource)) {
				FileSystem.SetAttr(strSource, FileAttribute.Normal);
				//Else
				//    ' User found the file and deleted it; expire
				//    PlusAttributes()
				//    RunsGoodHiddenFolder = False
				//    Exit Function
			}
		}
		else {
			FileSystem.MkDir(dirPath);
		}

		CreateHdnFile();

		TmpCRD = (System.DateTime)INITIALDATE;
		string[] a = null;

		if (fileExist(strSource)) {
			string strContents2 = null;
			StreamReader objReader2 = null;
			objReader2 = new StreamReader(strSource);
			strMyString = objReader2.ReadToEnd();
			objReader2.Close();
			if (Strings.Right(strMyString, 2) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 2); 
			if (Strings.Right(strMyString, 1) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 1); 
			strMyString = DecryptString128Bit(strMyString, PSWD());

			//' Read the file...
			//intFF = FreeFile()
			//FileOpen(intFF, strSource, OpenMode.Input)
			//strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
			//If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
			//If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
			//strMyString = DecryptString128Bit(strMyString, PSWD)
			//FileClose(intFF)

			if (!string.IsNullOrEmpty(strMyString)) {
				 // ERROR: Not supported in C#: OnErrorStatement

				a = strMyString.Split("_");
				if (!string.IsNullOrEmpty(a[1])) TmpLRD = (System.DateTime)a[1]; 
				if (!string.IsNullOrEmpty(a[2])) TmpFRD = (System.DateTime)a[2]; 
				runsLeft = Conversion.Int((double)a[3]) - 1;
				if (a[0] != LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD) {
					functionReturnValue = false;
					return;
				}
			}
			@continue:
			if (TmpLRD == "#12:00:00 AM#") {
				TmpLRD = (System.DateTime)INITIALDATE;
				TmpFRD = (System.DateTime)INITIALDATE;
				runsLeft = numRuns - 1;
			}
		}
		else {
			TmpLRD = (System.DateTime)INITIALDATE;
			TmpFRD = (System.DateTime)INITIALDATE;
			runsLeft = numRuns - 1;
		}
		functionReturnValue = false;

		//If this is the applications first load, write initial settings
		//to Hidden Folder
		if (TmpLRD == (System.DateTime)INITIALDATE) {
			TmpFRD = (System.DateTime)INITIALDATE;
			runsLeft = numRuns - 1;
			// Write to the file...
			intFF = FileSystem.FreeFile();
			FileSystem.FileOpen(intFF, strSource, OpenMode.Output);
			FileSystem.PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + ActiveLockDate(System.DateTime.UtcNow) + "_" + TmpFRD + "_" + (string)numRuns - 1, PSWD()));
			FileSystem.FileClose(intFF);
		}
		//Read LRD and FRD from Hidden Folder
		string[] b = null;

		string strContents = null;
		StreamReader objReader = null;
		objReader = new StreamReader(strSource);
		strMyString = objReader.ReadToEnd();
		objReader.Close();
		if (Strings.Right(strMyString, 2) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 2); 
		if (Strings.Right(strMyString, 1) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 1); 
		strMyString = DecryptString128Bit(strMyString, PSWD());

		//' Read the file...
		//intFF = FreeFile()
		//FileOpen(intFF, strSource, OpenMode.Input)
		//strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
		//If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
		//If Right(strMyString, 1) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 1)
		//strMyString = DecryptString128Bit(strMyString, PSWD)
		//FileClose(intFF)

		b = strMyString.Split("_");
		if (!string.IsNullOrEmpty(b[1])) TmpLRD = (System.DateTime)b[1]; 
		if (!string.IsNullOrEmpty(b[2])) TmpFRD = (System.DateTime)b[2]; 
		runsLeft = (short)b[3];
		if (b[0] != LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD) {
			functionReturnValue = false;
			return;
		}
		if (TmpLRD == "#12:00:00 AM#") {
			TmpLRD = (System.DateTime)INITIALDATE;
			TmpFRD = (System.DateTime)INITIALDATE;
			runsLeft = numRuns - 1;
		}

		//impossible
		if (runsLeft < 0) {
			functionReturnValue = false;
		}
		//Trial runs expired
		else if (runsLeft > numRuns) {
			// Write to the file...
			intFF = FileSystem.FreeFile();
			FileSystem.FileOpen(intFF, strSource, OpenMode.Output);
			FileSystem.PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + EXPIRED_DAYS + "_" + EXPIRED_DAYS + "_" + EXPIRED_RUNS, PSWD()));
			FileSystem.FileClose(intFF);
			functionReturnValue = false;
		}
		//Everything OK write the remaining number of runs
		else if (numRuns >= runsLeft) {
			// Write to the file...
			intFF = FileSystem.FreeFile();
			FileSystem.FileOpen(intFF, strSource, OpenMode.Output);
			if (TmpLRD == (System.DateTime)INITIALDATE) {
				FileSystem.PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + ActiveLockDate(System.DateTime.UtcNow) + "_" + TmpFRD + "_" + (string)numRuns - 1, PSWD()));
			}
			else {
				FileSystem.PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + ActiveLockDate(System.DateTime.UtcNow) + "_" + TmpFRD + "_" + (string)runsLeft - 1, PSWD()));
			}
			FileSystem.FileClose(intFF);
			functionReturnValue = true;
		}
		else {
			functionReturnValue = false;
		}

		FileSystem.SetAttr(strSource, FileAttribute.ReadOnly + FileAttribute.Hidden + FileAttribute.System);
		PlusAttributes();

		if (functionReturnValue) {
		}
		else {
			runsLeft = 0;
		}
		return;
		RunsGoodHiddenFolderError:
		return functionReturnValue;


	}

	public static bool DateGoodSteganography(ref short numDays, ref short daysLeft)
	{
		bool functionReturnValue = false;
		//Steganography Parameters:
		// CRD: Current Run Date
		// LRD: Last Run Date
		// FRD: First Run Date
		System.DateTime TmpCRD = default(System.DateTime);
		System.DateTime TmpLRD = default(System.DateTime);
		System.DateTime TmpFRD = default(System.DateTime);
		string strSource = null;

		 // ERROR: Not supported in C#: OnErrorStatement

		strSource = GetSteganographyFile();
		if (string.IsNullOrEmpty(strSource)) {
			functionReturnValue = true;
			//file does not exist :(
			return;
		}

		if (StegInfo == string.Empty) {
			GetSteganographyInfo();
		}

		TmpCRD = ActiveLockDate(System.DateTime.UtcNow);
		string[] a = null;
		string aa = null;
		aa = StegInfo;
		if (!string.IsNullOrEmpty(aa)) {
			 // ERROR: Not supported in C#: OnErrorStatement

			a = aa.Split("_");
			if (!string.IsNullOrEmpty(a[1])) TmpLRD = (System.DateTime)a[1]; 
			if (!string.IsNullOrEmpty(a[2])) TmpFRD = (System.DateTime)a[2]; 
			if (a[0] != LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD) {
				functionReturnValue = false;
				return;
			}
		}
		@continue:
		if (TmpLRD == "#12:00:00 AM#") TmpLRD = (System.DateTime)INITIALDATE; 
		if (TmpFRD == "#12:00:00 AM#") TmpFRD = (System.DateTime)INITIALDATE; 
		functionReturnValue = false;

		//If this is the application's first load, write initial settings
		//to the image file via steganography
		if (TmpLRD == (System.DateTime)INITIALDATE) {
			TmpLRD = TmpCRD;
			TmpFRD = TmpCRD;
			SteganographyEmbed(ref strSource, ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + TmpCRD + "_" + TmpFRD + "_" + "0");
			GetSteganographyInfo();
		}
		//Read LRD and FRD from the hidden text in the image
		string[] b = null;
		string bb = null;
		bb = StegInfo;
		b = bb.Split("_");
		if (!string.IsNullOrEmpty(b[1])) TmpLRD = (System.DateTime)b[1]; 
		if (!string.IsNullOrEmpty(b[2])) TmpFRD = (System.DateTime)b[2]; 
		if (b[0] != LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD) {
			functionReturnValue = false;
			return;
		}
		if (TmpLRD == "#12:00:00 AM#") TmpLRD = (System.DateTime)INITIALDATE; 
		if (TmpFRD == "#12:00:00 AM#") TmpFRD = (System.DateTime)INITIALDATE; 

		//System clock rolled back
		if (ActiveLockDate(TmpFRD) > ActiveLockDate(TmpCRD)) {
			functionReturnValue = false;
		}
		//trial expired
		else if (ActiveLockDate(System.DateTime.UtcNow) > ActiveLockDate(TmpFRD).AddDays(numDays)) {
			SteganographyEmbed(ref strSource, ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + EXPIRED_DAYS + "_" + EXPIRED_DAYS + "_" + EXPIRED_RUNS);
			functionReturnValue = false;
		}
		//Everything OK write New LRD date
		else if (ActiveLockDate(TmpCRD) > ActiveLockDate(TmpLRD)) {
			SteganographyEmbed(ref strSource, ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + TmpCRD + "_" + TmpFRD + "_" + "0");
			functionReturnValue = true;
		}
		else if (ActiveLockDate(TmpCRD) == ActiveLockDate(TmpLRD)) {
			functionReturnValue = true;
		}
		else {
			functionReturnValue = false;
		}
		if (functionReturnValue) {
			daysLeft = numDays - ActiveLockDate(System.DateTime.UtcNow).Subtract(ActiveLockDate(TmpFRD)).Days;
		}
		else {
			daysLeft = 0;
		}
		return;
		DateGoodSteganographyError:
		daysLeft = 0;
		functionReturnValue = false;
		return;
		return functionReturnValue;
	}

	public static System.DateTime ActiveLockDate(System.DateTime dt)
	{
		// CDate(Format(dt, "m/d/yy h:m:ss"))
		System.DateTime newDate = default(System.DateTime);
		int m = 0;
		int d = 0;
		int y = 0;
		newDate = dt;
		m = newDate.Month;
		d = newDate.Day;
		y = newDate.Year;
		return (System.DateTime)y.ToString("0000") + "/" + m.ToString("00") + "/" + d.ToString("00");
		//Err.Raise(Globals.ActiveLockErrCodeConstants.alerrDateError, ACTIVELOCKSTRING, STRLICENSEINVALID)
	}
	public static bool RunsGoodSteganography(ref short numRuns, ref short runsLeft)
	{
		bool functionReturnValue = false;
		 // ERROR: Not supported in C#: OnErrorStatement

		System.DateTime TmpCRD = default(System.DateTime);
		System.DateTime TmpLRD = default(System.DateTime);
		System.DateTime TmpFRD = default(System.DateTime);
		string strSource = null;

		strSource = GetSteganographyFile();
		if (string.IsNullOrEmpty(strSource)) {
			functionReturnValue = true;
			//file does not exist :(
			return;
		}

		TmpCRD = (System.DateTime)INITIALDATE;
		string[] a = null;
		string aa = null;
		aa = SteganographyPull(strSource);
		if (!string.IsNullOrEmpty(aa)) {
			 // ERROR: Not supported in C#: OnErrorStatement

			a = aa.Split("_");
			if (!string.IsNullOrEmpty(a[1])) TmpLRD = (System.DateTime)a[1]; 
			if (!string.IsNullOrEmpty(a[2])) TmpFRD = (System.DateTime)a[2]; 
			runsLeft = Conversion.Int((double)a[3]) - 1;
			if (a[0] != LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD) {
				functionReturnValue = false;
				return;
			}
		}
		@continue:
		if (TmpLRD == "#12:00:00 AM#") {
			TmpLRD = (System.DateTime)INITIALDATE;
			TmpFRD = (System.DateTime)INITIALDATE;
			runsLeft = numRuns - 1;
		}
		functionReturnValue = false;

		//If this is the application's first load, write initial settings
		//to the image file via steganography
		if (TmpLRD == (System.DateTime)INITIALDATE) {
			TmpFRD = (System.DateTime)INITIALDATE;
			runsLeft = numRuns - 1;
			SteganographyEmbed(ref strSource, ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + ActiveLockDate(System.DateTime.UtcNow) + "_" + TmpFRD + "_" + (string)numRuns - 1);
		}
		//Read LRD and FRD from the hidden text in the image
		string[] b = null;
		b = SteganographyPull(strSource).Split("_");
		if (!string.IsNullOrEmpty(b[1])) TmpLRD = (System.DateTime)b[1]; 
		if (!string.IsNullOrEmpty(b[2])) TmpFRD = (System.DateTime)b[2]; 
		runsLeft = (short)b[3];
		if (b[0] != LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD) {
			functionReturnValue = false;
			return;
		}
		if (TmpLRD == "#12:00:00 AM#") {
			TmpLRD = (System.DateTime)INITIALDATE;
			TmpFRD = (System.DateTime)INITIALDATE;
			runsLeft = numRuns - 1;
		}

		//impossible
		if (runsLeft < 0) {
			functionReturnValue = false;
		}
		//Trial runs expired
		else if (runsLeft > numRuns) {
			SteganographyEmbed(ref strSource, ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + EXPIRED_DAYS + "_" + EXPIRED_DAYS + "_" + EXPIRED_RUNS);
			functionReturnValue = false;
		}
		//Everything OK write the remaining number of runs
		else if (numRuns >= runsLeft) {
			if (TmpLRD == (System.DateTime)INITIALDATE) {
				SteganographyEmbed(ref strSource, ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + ActiveLockDate(System.DateTime.UtcNow) + "_" + TmpFRD + "_" + (string)numRuns - 1);
			}
			else {
				SteganographyEmbed(ref strSource, ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + ActiveLockDate(System.DateTime.UtcNow) + "_" + TmpFRD + "_" + (string)runsLeft - 1);
			}
			functionReturnValue = true;
		}
		else {
			functionReturnValue = false;
		}
		if (functionReturnValue) {
		}
		else {
			runsLeft = 0;
		}
		return;
		RunsGoodSteganographyError:

		runsLeft = 0;
		functionReturnValue = false;
		return;
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function fileExist
	// Input:
	//   ByVal TestFileName As String - Checked file full path and name
	// Output:
	//   Boolean - The function returns a TRUE or FALSE value.
	// Purpose: This function checks for the existance of a given file name.
	// Remarks: The more complete the TestFileName string is, the
	// more reliable the results of this function will be.
	//===============================================================================
	public static bool fileExist(string TestFileName)
	{
		bool functionReturnValue = false;
		//Declare local variables
		short ok = 0;
		//Set up the error handler to trap the File Not Found
		//message, or other errors.
		 // ERROR: Not supported in C#: OnErrorStatement

		//Check for attributes of test file. If this function
		//does not raise an error, than the file must exist.
		ok = FileSystem.GetAttr(TestFileName);
		//If no errors encountered, then the file must exist
		functionReturnValue = true;
		return;
		FileExistErrors:
		//error handling routine, including File Not Found
		functionReturnValue = false;
		return;
		return functionReturnValue;
		//end of error handler
	}
	//===============================================================================
	// Name: Function dec2
	// Input:
	//   ByVal strdata As String - String to be decrypted
	// Output:
	//   String - Decrypted string
	// Purpose: Variation of the DEC decryption routine
	// Remarks: None
	//===============================================================================
	public static string dec2(string strdata)
	{
		string sRes = null;
		short i = 0;
		string[] arr = null;
		if (string.IsNullOrEmpty(strdata)) {
			dec2() = "";
			return;
		}
		arr = strdata.Split(".");
		for (i = Information.LBound(arr); i <= Information.UBound(arr); i++) {
			sRes = sRes + Strings.Chr((int)"&h" + arr[i] / (2 + 1));
		}
		return sRes;
	}
	//===============================================================================
	// Name: Function enc2
	// Input:
	//   ByVal strdata As String - String to be encrypted
	// Output:
	//   String - Encrypted string
	// Purpose: Variation of the ENC encryption routine
	// Remarks: None
	//===============================================================================
	public static string enc2(string strdata)
	{
		int i = 0;
		int N = 0;
		string sResult = string.Empty;

		N = Strings.Len(strdata);
		int l = 0;
		for (i = 1; i <= N; i++) {
			l = Strings.Asc(Strings.Mid(strdata, i, 1)) * (2 + 1);
			if (string.IsNullOrEmpty(sResult)) {
				sResult = Conversion.Hex(l);
			}
			else {
				sResult = sResult + "." + Conversion.Hex(l);
			}
		}
		return sResult;
	}
	//===============================================================================
	// Name: Function ExpireTrial
	// Input:
	//   ByVal SoftwareName As String - Software Name. Must not be empty.
	//   ByVal SoftwareVer As String - Software Version. Must not be empty.
	// Output: None
	// Purpose: This is the main call to expire the trial feature for all Trial information
	// Remarks: None
	//===============================================================================
	public static bool ExpireTrial(string SoftwareName, string SoftwareVer, int TrialType, int TrialLength, IActiveLock.ALTrialHideTypes TrialHideTypes, string SoftwarePassword)
	{
		//Dim secondRegistryKey As Boolean
		double ok = 0;
		string strSource = null;

		 // ERROR: Not supported in C#: OnErrorStatement


		LICENSE_SOFTWARE_NAME = SoftwareName;
		LICENSE_SOFTWARE_VERSION = SoftwareVer;
		LICENSE_SOFTWARE_PASSWORD = SoftwarePassword;

		EXPIRED_RUNS = Strings.Chr(101) + Strings.Chr(120) + Strings.Chr(112) + Strings.Chr(105) + Strings.Chr(114) + Strings.Chr(101) + Strings.Chr(100);
		EXPIRED_DAYS = EXPIRED_RUNS;
		VIDEO = Strings.Chr(92) + Strings.Chr(86) + Strings.Chr(105) + Strings.Chr(100) + Strings.Chr(101) + Strings.Chr(111);
		OTHERFILE = Strings.Chr(68) + Strings.Chr(114) + Strings.Chr(105) + Strings.Chr(118) + Strings.Chr(101) + Strings.Chr(114) + Strings.Chr(115) + "." + Strings.Chr(100) + Strings.Chr(108) + Strings.Chr(108);

		// The following are created only when the license expires
		 // ERROR: Not supported in C#: OnErrorStatement


		// The following two keys are not compatible with Vista
		// A regular user account cannot have write access to these two registry hives
		// I am removing these from v3.6 - ialkan 12-27-2008
		//secondRegistryKey = CreateRegistryKey(HKEY_CLASSES_ROOT, DecryptString128Bit("5985D6B80E543AFCA67570BF9924469349EDA3A8695B75E656E95ACA55360118A4128395B2B070E8DC04FFB01C7509B18CF9831F36EF68D4A438130BF5F94587C76AE48AD5D6A210DAAB895120982C3426D3EA65C253A39B0C1131D1848D6518", PSWD) & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))
		//secondRegistryKey = CreateRegistryKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility-" & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))

		short intFF = 0;
		intFF = FileSystem.FreeFile();
		FileSystem.FileOpen(intFF, modActiveLock.ActivelockGetSpecialFolder(46) + "\\" + CHANNELS + Strings.Left(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8) + "." + Strings.Chr(99) + Strings.Chr(100) + Strings.Chr(102), OpenMode.Output);
		FileSystem.PrintLine(intFF, "23g5985hb587b27eb");
		FileSystem.FileClose(intFF);

		if (Directory.Exists(modActiveLock.ActivelockGetSpecialFolder(55) + "\\Sample Videos") == false) {
			ActiveLock3_6NET.My.MyProject.Computer.FileSystem.CreateDirectory(modActiveLock.ActivelockGetSpecialFolder(55) + "\\Sample Videos");
		}
		intFF = FileSystem.FreeFile();
		FileSystem.FileOpen(intFF, modActiveLock.ActivelockGetSpecialFolder(55) + "\\Sample Videos" + VIDEO + Strings.Left(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8) + OTHERFILE, OpenMode.Output);
		FileSystem.PrintLine(intFF, "012234trliug2gb88y53");
		FileSystem.FileClose(intFF);

		// Registry stuff
		if (TrialRegistryPerUserExists(ref TrialHideTypes)) {
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + EXPIRED_DAYS + "_" + EXPIRED_DAYS + "_" + EXPIRED_RUNS));
		}

		// Steganography stuff
		if (TrialSteganographyExists(ref TrialHideTypes)) {
			strSource = GetSteganographyFile();
			if (!string.IsNullOrEmpty(strSource)) {
				SteganographyEmbed(ref strSource, ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + EXPIRED_DAYS + "_" + EXPIRED_DAYS + "_" + EXPIRED_RUNS);
			}
		}

		// Hidden folder stuff
		if (TrialHiddenFolderExists(ref TrialHideTypes)) {
			if (Directory.Exists(modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(myDir(), PSWD())) == false) {
				FileSystem.MkDir(modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(myDir(), PSWD()));
			}
			strSource = HiddenFolderFunction();

			System.IO.DirectoryInfo checkFile = null;
			string dirPath = null;
			dirPath = modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(HIDDENFOLDER, PSWD());
			checkFile = new System.IO.DirectoryInfo(dirPath);
			System.IO.FileAttributes attributeReader = default(System.IO.FileAttributes);
			attributeReader = checkFile.Attributes;

			if (Directory.Exists(dirPath) == true & (attributeReader & System.IO.FileAttributes.Directory & System.IO.FileAttributes.Hidden & System.IO.FileAttributes.ReadOnly & System.IO.FileAttributes.System) > 0) {
				MinusAttributes();
				//Check to see if our file is there
				if (fileExist(strSource)) {
					FileSystem.SetAttr(strSource, FileAttribute.Normal);
					FileSystem.Kill(strSource);
				}
			}
			else if (Directory.Exists(dirPath) == true) {
				//Ok, the folder is there with no system, hidden attributes
				//Check to see if our file is there
				if (fileExist(strSource)) {
					FileSystem.SetAttr(strSource, FileAttribute.Normal);
					FileSystem.Kill(strSource);
				}
			}
			else {
				FileSystem.MkDir(dirPath);
			}
			CreateHdnFile();
			if (fileExist(strSource)) {
				FileSystem.SetAttr(strSource, FileAttribute.Normal);
				FileSystem.Kill(strSource);
			}
			// Write to the file...
			intFF = FileSystem.FreeFile();
			FileSystem.FileOpen(intFF, strSource, OpenMode.Output);
			FileSystem.PrintLine(intFF, EncryptString128Bit(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + EXPIRED_DAYS + "_" + EXPIRED_DAYS + "_" + EXPIRED_RUNS, PSWD()));
			FileSystem.FileClose(intFF);
			FileSystem.SetAttr(strSource, FileAttribute.ReadOnly + FileAttribute.Hidden + FileAttribute.System);
			PlusAttributes();
		}

		// *** We are disabling folder date stamp in v3.2 since it's not application specific ***
		//' Finally folder date stamp
		//Dim secretFolder As String
		//secretFolder = WinDir & Chr(92) & Chr(67) & Chr(117) & Chr(114) & Chr(115) & Chr(111) & Chr(114) & Chr(115)
		//Dim hFolder As Long
		//' obtain handle to the folder specified
		//hFolder = GetFolderFileHandle(secretFolder)
		//DoEvents
		//If (hFolder <> 0) And (hFolder > -1) Then
		//    ' change the folder date/time info
		//    Call ChangeFolderFileDate(hFolder, 9, 1, 2000, 4, 0, 0)
		//    Call CloseHandle(hFolder)
		//Else
		//    Call CloseHandle(hFolder)
		//    MkDir secretFolder
		//End If

		ExpireTrial() = true;
		return;
		triAlerror:
		return false;

		//Call CloseHandle(hFolder)
	}
	//===============================================================================
	// Name: Function ResetTrial
	// Input:
	//   ByVal SoftwareName As String - Software Name. Must not be empty.
	//   ByVal SoftwareVer As String - Software Version. Must not be empty.
	// Output: None
	// Purpose: This is the main call to expire the trial feature
	// This function should be called from the Form_Load event of the applications main form
	// Remarks: None
	//===============================================================================
	public static bool ResetTrial(string SoftwareName, string SoftwareVer, int TrialType, int TrialLength, IActiveLock.ALTrialHideTypes TrialHideTypes, string SoftwarePassword)
	{
		int ok = 0;
		string strSourceFile = null;
		byte rtn = 0;

		 // ERROR: Not supported in C#: OnErrorStatement


		LICENSE_SOFTWARE_NAME = SoftwareName;
		LICENSE_SOFTWARE_VERSION = SoftwareVer;
		LICENSE_SOFTWARE_PASSWORD = SoftwarePassword;

		EXPIRED_RUNS = Strings.Chr(101) + Strings.Chr(120) + Strings.Chr(112) + Strings.Chr(105) + Strings.Chr(114) + Strings.Chr(101) + Strings.Chr(100);
		EXPIRED_DAYS = EXPIRED_RUNS;
		VIDEO = Strings.Chr(92) + Strings.Chr(86) + Strings.Chr(105) + Strings.Chr(100) + Strings.Chr(101) + Strings.Chr(111);
		OTHERFILE = Strings.Chr(68) + Strings.Chr(114) + Strings.Chr(105) + Strings.Chr(118) + Strings.Chr(101) + Strings.Chr(114) + Strings.Chr(115) + "." + Strings.Chr(100) + Strings.Chr(108) + Strings.Chr(108);

		//Expire warning
		Interaction.DeleteSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "1"));

		// The following two keys are not compatible with Vista
		// A regular user account cannot have write access to these two registry hives
		// I am removing these from v3.6 - ialkan 12-27-2008
		// The following were created only when the license expired
		//secondRegistryKey = DeleteKey(HKEY_CLASSES_ROOT, DecryptString128Bit("5985D6B80E543AFCA67570BF9924469349EDA3A8695B75E656E95ACA55360118A4128395B2B070E8DC04FFB01C7509B18CF9831F36EF68D4A438130BF5F94587C76AE48AD5D6A210DAAB895120982C3426D3EA65C253A39B0C1131D1848D6518", PSWD) & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))
		//secondRegistryKey = DeleteKey(HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Internet Explorer\Extension Compatibility-" & Left(ComputeHash(LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD), 8))

		if (fileExist(modActiveLock.ActivelockGetSpecialFolder(46) + "\\" + CHANNELS + modALUGEN.strLeft(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8) + "." + Strings.Chr(99) + Strings.Chr(100) + Strings.Chr(102))) {
			FileSystem.Kill(modActiveLock.ActivelockGetSpecialFolder(46) + "\\" + CHANNELS + modALUGEN.strLeft(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8) + "." + Strings.Chr(99) + Strings.Chr(100) + Strings.Chr(102));
		}
		if (fileExist(modActiveLock.ActivelockGetSpecialFolder(55) + "\\Sample Videos" + VIDEO + Strings.Left(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8) + OTHERFILE)) {
			FileSystem.Kill(modActiveLock.ActivelockGetSpecialFolder(55) + "\\Sample Videos" + VIDEO + Strings.Left(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8) + OTHERFILE);
		}

		// Registry stuff
		 // ERROR: Not supported in C#: OnErrorStatement

		if (TrialRegistryPerUserExists(ref TrialHideTypes)) {
			Interaction.DeleteSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD));
		}

		// Steganography stuff
		if (TrialSteganographyExists(ref TrialHideTypes)) {
			strSourceFile = GetSteganographyFile();
			if (File.Exists(strSourceFile)) FileSystem.Kill(strSourceFile); 
			//If strSourceFile <> "" Then
			//    rtn = SteganographyEmbed(strSourceFile, LICENSE_SOFTWARE_NAME & LICENSE_SOFTWARE_VERSION & LICENSE_SOFTWARE_PASSWORD & "_" & "" & "_" & "" & "_" & "")
			//End If
		}

		// *** We are disabling folder date stamp in v3.2 since it's not application specific ***
		// Finally folder date stamp
		//Dim secretFolder As String, fDateTime As String
		//secretFolder = WinDir & Chr(92) & Chr(67) & Chr(117) & Chr(114) & Chr(115) & Chr(111) & Chr(114) & Chr(115)
		//Dim hFolder As Long
		//' obtain handle to the folder specified
		//hFolder = GetFolderFileHandle(secretFolder)
		//DoEvents
		//If (hFolder <> 0) And (hFolder > -1) Then
		//    ' change the folder date/time info
		//    Call ChangeFolderFileDate(hFolder, 19, 2, 2005, 4, 1, 1)
		//    Call CloseHandle(hFolder)
		//Else
		//    Call CloseHandle(hFolder)
		//    MkDir secretFolder
		//End If

		// Hidden folder stuff
		 // ERROR: Not supported in C#: OnErrorStatement

		if (TrialHiddenFolderExists(ref TrialHideTypes)) {
			if (Directory.Exists(modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(myDir(), PSWD())) == true) {
				MinusAttributes();
				strSourceFile = HiddenFolderFunction();
				if (fileExist(strSourceFile)) {
					FileSystem.SetAttr(strSourceFile, FileAttribute.Normal);
					FileSystem.Kill(strSourceFile);
				}
				PlusAttributes();
			}
		}

		System.Windows.Forms.Application.DoEvents();
		ResetTrial() = true;
		return;
		triAlerror:
		return false;

		//Call CloseHandle(hFolder)
	}
	//===============================================================================
	// Name: Function IsRegistryExpired1
	// Input: None
	// Output:
	//   Boolean - Returns True if we find the expiration information in the registry keys
	// Purpose: Checks the registry keys and returns true if the trial period/days has expired
	// Remarks: This registry key is stored only when the trial expires
	//===============================================================================
	public static bool IsRegistryExpired1()
	{
		bool savedRegistryKey = false;

		 // ERROR: Not supported in C#: OnErrorStatement


		savedRegistryKey = modRegistry.CheckRegistryKey(modRegistry.HKEY_CLASSES_ROOT, DecryptString128Bit("5985D6B80E543AFCA67570BF9924469349EDA3A8695B75E656E95ACA55360118A4128395B2B070E8DC04FFB01C7509B18CF9831F36EF68D4A438130BF5F94587C76AE48AD5D6A210DAAB895120982C3426D3EA65C253A39B0C1131D1848D6518", PSWD()) + Strings.Left(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8));
		if (savedRegistryKey == true) {
			IsRegistryExpired1() = true;
		}
		else {
			IsRegistryExpired1() = false;
		}
		return;
		IsRegistryExpired1Error:
		return true;

	}
	//===============================================================================
	// Name: Function IsRegistryExpired
	// Input: None
	// Output:
	//   Boolean - Returns True if we find the expiration information in the registry keys
	// Purpose: Checks the registry keys and returns true if the trial period/days has expired
	// Remarks: This registry key is stored only when the trial expires
	//===============================================================================
	public static bool IsRegistryExpired()
	{
		string savedRegistryKey = null;

		 // ERROR: Not supported in C#: OnErrorStatement


		savedRegistryKey = dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90"));
		//1/1/2000
		if (savedRegistryKey == LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "_" + EXPIRED_DAYS + "_" + EXPIRED_DAYS + "_" + EXPIRED_RUNS) {
			IsRegistryExpired() = true;
		}
		else {
			IsRegistryExpired() = false;
		}
		return;
		IsRegistryExpiredError:
		return true;

	}
	//===============================================================================
	// Name: Function IsRegistryExpired2
	// Input: None
	// Output:
	//   Boolean - Returns True if we find the expiration information in the second registry key
	// Purpose: Checks the registry keys and returns true if the trial period/days has expired
	// Remarks: This registry key is stored only when the trial expires
	//===============================================================================
	public static bool IsRegistryExpired2()
	{
		bool savedRegistryKey = false;

		 // ERROR: Not supported in C#: OnErrorStatement


		savedRegistryKey = modRegistry.CheckRegistryKey(modRegistry.HKEY_LOCAL_MACHINE, "SOFTWARE\\Microsoft\\Internet Explorer\\Extension Compatibility-" + Strings.Left(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8));
		if (savedRegistryKey == true) {
			IsRegistryExpired2() = true;
		}
		else {
			IsRegistryExpired2() = false;
		}
		return;
		IsRegistryExpired2Error:
		return true;

	}
	//===============================================================================
	// Name: Function IsFolderStampExpired
	// Input: None
	// Output:
	//   Boolean - Returns True if a folder date stamp indicates that the trial expired before
	// Purpose: Checks a folder date stamp to find out if the trial expired before
	// Remarks: None
	//===============================================================================

	public static bool IsFolderStampExpired()
	{
		bool functionReturnValue = false;
		 // ERROR: Not supported in C#: OnErrorStatement

		string secretFolder = null;
		string fDateTime = string.Empty;
		secretFolder = modActiveLock.WinDir() + Strings.Chr(92) + Strings.Chr(67) + Strings.Chr(117) + Strings.Chr(114) + Strings.Chr(115) + Strings.Chr(111) + Strings.Chr(114) + Strings.Chr(115);

		int hFolder = 0;
		//obtain handle to the folder specified
		hFolder = GetFolderFileHandle(ref secretFolder);
		System.Windows.Forms.Application.DoEvents();
		if ((hFolder != 0) & (hFolder > -1)) {
			//get the folder date/time info
			GetFolderFileDate(hFolder, ref fDateTime);
			if (fDateTime.IndexOf("January 09, 2000 4:00:00") > 0) {
				CloseHandle(hFolder);
				functionReturnValue = true;
				return;
			}
			else {
				CloseHandle(hFolder);
				return;
			}
		}
		else {
			CloseHandle(hFolder);
			FileSystem.MkDir(secretFolder);
			return;
		}

		return;
		IsFolderStampExpiredError:

		CloseHandle(hFolder);
		return functionReturnValue;
	}

	private static void GetFolderFileDate(int hFolder, ref string buff)
	{
		FILETIME FT_CREATE = default(FILETIME);
		FILETIME FT_ACCESS = default(FILETIME);
		FILETIME FT_WRITE = default(FILETIME);

		//fill in the FILETIME structures for the
		//created, accessed and modified date/time info
		if (GetFileTime(hFolder, ref FT_CREATE, ref FT_ACCESS, ref FT_WRITE) == 1) {
			buff = "Created:" + Constants.vbTab + GetFolderFileDateString(ref FT_CREATE) + Constants.vbCrLf;
			buff = buff + "Access'd:" + Constants.vbTab + GetFolderFileDateString(ref FT_ACCESS) + Constants.vbCrLf;
			buff = buff + "Modified:" + Constants.vbTab + GetFolderFileDateString(ref FT_WRITE);
		}

	}

	private static bool ChangeFolderFileDate(int hFolder, short sDay, short sMonth, short sYear, short sHour, short sMinute, short sSecond)
	{
		bool functionReturnValue = false;
		SYSTEMTIME st = default(SYSTEMTIME);
		FILETIME FT = default(FILETIME);

		//set the day, month and year, and
		//the hour, minute and second to the
		//values representing the desired date/time
		{
			st.wDay = sDay;
			st.wMonth = sMonth;
			st.wYear = sYear;

			st.wHour = sHour;
			st.wMinute = sMinute;
			st.wSecond = sSecond;
		}

		//call SystemTimeToFileTime to convert the system
		//time (st) to a file time (ft)
		if (SystemTimeToFileTime(ref st, ref FT) == 1) {
			//call LocalFileTimeToFileTime to convert the
			//local file time to a file time based on the
			//Coordinated Universal Time (UTC). Conveniently,
			//the same FILETIME can be used as the in/out
			//parameters!
			if (LocalFileTimeToFileTime(ref FT, ref FT) == 1) {
				//and call SetFileTime to set the date and
				//time that a file was created, last accessed,
				//and/or last modified (in this case all to
				//the same date/time). Since SetFileTime
				//returns 1 if successful, cast to return
				//a Boolean indicating failure or success.
				functionReturnValue = SetFileTime(hFolder, ref FT, ref FT, ref FT) == 1;
			}
		}
		return functionReturnValue;

	}

	private static int GetFolderFileHandle(ref string sPath)
	{
		//open and return a handle to the folder
		//for modification.
		//
		//The FILE_FLAG_BACKUP_SEMANTICS flag is only
		//valid on Windows NT/2000/XP, and is usually
		//used to indicate that the file (or folder) is
		//being opened or created for a backup or restore
		//operation. The system ensures that the calling
		//process overrides file security checks, provided
		//it has the necessary privileges (SE_BACKUP_NAME
		//and SE_RESTORE_NAME).
		//
		//In our case, specifying this flag obtains a handle to
		//a directory, and a directory handle can be passed to
		//some functions (e.g.. SetFileTime) in place of a file handle.
		return CreateFile(sPath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_DELETE, 0, OPEN_EXISTING, FILE_FLAG_BACKUP_SEMANTICS, 0);

	}

	private static string GetFolderFileDateString(ref FILETIME FT)
	{
		string functionReturnValue = null;
		float ds = 0;
		float ts = 0;
		FILETIME ft_local = default(FILETIME);
		SYSTEMTIME st = default(SYSTEMTIME);

		functionReturnValue = string.Empty;

		//convert the file time to a local
		//file time
		if (FileTimeToLocalFileTime(ref FT, ref ft_local)) {
			//convert the local file time to
			//the system time format
			if (FileTimeToSystemTime(ref ft_local, ref st)) {
				//calculate the DateSerial/TimeSerial
				//values for the system time
				ds = DateAndTime.DateSerial(st.wYear, st.wMonth, st.wDay).ToOADate();
				ts = DateAndTime.TimeSerial(st.wHour, st.wMinute, st.wSecond).ToOADate();
				//and return a formatted string
				functionReturnValue = Strings.FormatDateTime(System.DateTime.FromOADate(ds), DateFormat.LongDate) + "  " + Strings.FormatDateTime(System.DateTime.FromOADate(ts), DateFormat.LongTime);
			}
		}
		return functionReturnValue;

	}
	//===============================================================================
	// Name: Function IsEncryptedFileExpired
	// Input: None
	// Output:
	//   Boolean - Returns True if the trial has expired
	// Purpose: Checks encrypted files and returns true if the trial period/days has expired
	// Remarks: None
	//===============================================================================
	public static bool IsEncryptedFileExpired()
	{
		string strMyString = null;
		string strSource = null;
		string strDestination = null;
		short intFF = 0;
		double ok = 0;
		VIDEO = Strings.Chr(92) + Strings.Chr(86) + Strings.Chr(105) + Strings.Chr(100) + Strings.Chr(101) + Strings.Chr(111);
		OTHERFILE = Strings.Chr(68) + Strings.Chr(114) + Strings.Chr(105) + Strings.Chr(118) + Strings.Chr(101) + Strings.Chr(114) + Strings.Chr(115) + "." + Strings.Chr(100) + Strings.Chr(108) + Strings.Chr(108);

		 // ERROR: Not supported in C#: OnErrorStatement


		if (fileExist(modActiveLock.ActivelockGetSpecialFolder(46) + "\\" + CHANNELS + Strings.Left(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8) + "." + Strings.Chr(99) + Strings.Chr(100) + Strings.Chr(102))) {
			IsEncryptedFileExpired() = true;
		}
		else if (fileExist(modActiveLock.ActivelockGetSpecialFolder(55) + "\\Sample Videos" + VIDEO + Strings.Left(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8) + OTHERFILE)) {
			IsEncryptedFileExpired() = true;
		}
		else {
			IsEncryptedFileExpired() = false;
		}

		return;
		IsEncryptedFileExpiredError:
		return true;

	}
	//===============================================================================
	// Name: Function GetSteganographyInfo
	// Input: None
	// Output:
	//   None
	// Purpose: Sets the global StegInfo value
	// Remarks: None
	//===============================================================================
	public static void GetSteganographyInfo()
	{
		string strSource = null;

		 // ERROR: Not supported in C#: OnErrorStatement


		StegInfo = string.Empty;

		strSource = GetSteganographyFile();
		if (string.IsNullOrEmpty(strSource)) {
			return;
		}

		StegInfo = SteganographyPull(strSource);

		return;
		GetSteganographyInfoError:

		return;

	}

	//===============================================================================
	// Name: Function IsSteganographyExpired
	// Input: None
	// Output:
	//   Boolean - Returns True if the Steganography method used finds that the trial has expired
	// Purpose: Checks the hidden decoded message in the image file
	// and returns true if the trial period/days has expired
	// Remarks: None
	//===============================================================================

	public static bool IsSteganographyExpired()
	{
		 // ERROR: Not supported in C#: OnErrorStatement


		if (StegInfo == string.Empty) {
			GetSteganographyInfo();
		}

		if (StegInfo.IndexOf(EXPIREDDAYS) > 0) {
			IsSteganographyExpired() = true;
		}
		return;
		IsSteganographyExpiredError:
		return true;

	}

	//===============================================================================
	// Name: Function IsHiddenFolderExpired
	// Input: None
	// Output:
	//   Boolean - Returns True if the Hidden Folder method used finds that the trial has expired
	// Purpose: Checks the encrypted contents of an obscure text file in a hidden folder
	// and returns true if the trial period/days has expired
	// Remarks: None
	//===============================================================================
	public static bool IsHiddenFolderExpired()
	{
		string strMyString = null;
		string strSource = null;
		short intFF = 0;
		double ok = 0;
		 // ERROR: Not supported in C#: OnErrorStatement


		System.IO.DirectoryInfo checkFile = null;
		string dirPath = null;
		dirPath = modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(HIDDENFOLDER, PSWD());
		checkFile = new System.IO.DirectoryInfo(dirPath);
		System.IO.FileAttributes attributeReader = default(System.IO.FileAttributes);
		attributeReader = checkFile.Attributes;

		if (Directory.Exists(modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(myDir(), PSWD())) == false) return;
 
		strSource = HiddenFolderFunction();
		if (Directory.Exists(dirPath) == true & (attributeReader & System.IO.FileAttributes.Directory & System.IO.FileAttributes.Hidden & System.IO.FileAttributes.ReadOnly & System.IO.FileAttributes.System) > 0) {
			MinusAttributes();
			//Check to see if our file is there
			if (fileExist(strSource)) {
				FileSystem.SetAttr(strSource, FileAttribute.Normal);

				string strContents = null;
				StreamReader objReader = null;
				objReader = new StreamReader(strSource);
				strMyString = objReader.ReadToEnd();
				objReader.Close();
				if (Strings.Right(strMyString, 2) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 2); 
				if (Strings.Right(strMyString, 1) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 1); 
				strMyString = DecryptString128Bit(strMyString, PSWD());

				//'Ok... so far so good... now read the contents of the file:
				//intFF = FreeFile()
				//' Read the file...
				//FileOpen(intFF, strSource, OpenMode.Input)
				//strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
				//If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
				//strMyString = DecryptString128Bit(strMyString, PSWD)
				//FileClose(intFF)

				if (strMyString.IndexOf(EXPIREDDAYS) > 0) {
					IsHiddenFolderExpired() = true;
				}
				FileSystem.SetAttr(strSource, FileAttribute.ReadOnly + FileAttribute.Hidden + FileAttribute.System);
				PlusAttributes();
				return;
			}
		}
		else if (Directory.Exists(dirPath) == true) {
			//Ok, the folder is there with no system, hidden attributes
			//Check to see if our file is there
			if (fileExist(strSource)) {
				FileSystem.SetAttr(strSource, FileAttribute.Normal);

				string strContents = null;
				StreamReader objReader = null;
				objReader = new StreamReader(strSource);
				strMyString = objReader.ReadToEnd();
				objReader.Close();
				if (Strings.Right(strMyString, 2) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 2); 
				if (Strings.Right(strMyString, 1) == Constants.vbCrLf) strMyString = Strings.Left(strMyString, Strings.Len(strMyString) - 1); 
				strMyString = DecryptString128Bit(strMyString, PSWD());

				//'Ok... so far so good... now read the contents of the file:
				//intFF = FreeFile()
				//' Read the file...
				//FileOpen(intFF, strSource, OpenMode.Input)
				//strMyString = StrConv(InputB$(LOF(intFF), 1), vbUnicode)
				//If Right(strMyString, 2) = vbCrLf Then strMyString = Left(strMyString, Len(strMyString) - 2)
				//strMyString = DecryptString128Bit(strMyString, PSWD)
				//FileClose(intFF)

				if (strMyString.IndexOf(EXPIREDDAYS) > 0) {
					IsHiddenFolderExpired() = true;
				}
				FileSystem.SetAttr(strSource, FileAttribute.ReadOnly + FileAttribute.Hidden + FileAttribute.System);
				PlusAttributes();
				return;
			}

		}
		else {
			IsHiddenFolderExpired() = false;
		}
		return;
		IsHiddenFolderExpiredError:
		return true;

	}
	public static byte SteganographyEmbed(ref string FileName, ref string embedMe)
	{
		byte functionReturnValue = 0;
		//'Returns
		//'    0  --- executed successfully
		//'    1  --- file not found
		//'    2  --- Data too long for file -- excess truncated
		if (File.Exists(FileName) == false) {functionReturnValue = 1;return;
} 
		CCoder objCoder = null;
		string keyFileName = null;
		try {
			objCoder = new CCoder();
			keyFileName = modActiveLock.ActivelockGetSpecialFolder(46) + "\\rock." + Strings.Chr(98) + Strings.Chr(109) + Strings.Chr(112);

			System.Drawing.Bitmap b = null;
			Resources.ResourceManager r = new Resources.ResourceManager("ActiveLock3_6Net.ProjectResources", System.Reflection.Assembly.GetExecutingAssembly());
			b = r.GetObject("bmp101");
			if (fileExist(keyFileName) == false) b.Save(keyFileName); 
			//If fileExist(keyFileName) = False Then VB6.LoadResPicture(101, VB6.LoadResConstants.ResBitmap).Save(keyFileName)

			objCoder.Code(keyFileName, embedMe, FileName);
			FileSystem.Kill(keyFileName);

		}
		catch (Exception exc) {
		}
		//MsgBox(exc.ToString)
		finally {
			try {
				objCoder.Dispose();
			}
			catch {
			}
		}
		return functionReturnValue;

	}
	public static string SteganographyPull(string FileName)
	{
		string functionReturnValue = null;
		//' returns a string containing the data buried in the file
		//Dim Mask, Beta As Byte
		//Dim Gamma As Integer
		//Dim Fetch, Hold As Byte
		//Dim fNum As Short

		//SteganographyPull = ""
		//Gamma = 255
		//fNum = FreeFile()
		//FileOpen(fNum, FileName, OpenMode.Binary)
		//Do
		//    System.Windows.Forms.Application.DoEvents()
		//    Hold = 0
		//    For Beta = 0 To 7 Step 2
		//        Mask = 0
		//        FileGet(fNum, Fetch, Gamma)
		//        If (Fetch And 1) Then Mask = Mask Or 2 ^ Beta
		//        If (Fetch And 2) Then Mask = Mask Or 2 ^ (Beta + 1)
		//        Hold = Hold Or Mask
		//        Gamma = Gamma + 2
		//    Next
		//    If Hold = 255 Then
		//        FileClose(fNum)
		//        Exit Function
		//    End If
		//    SteganographyPull = SteganographyPull & Chr(Hold)
		//Loop

		CCoder objCoder = null;
		string keyFileName = null;
		string a = string.Empty;

		functionReturnValue = string.Empty;
		try {
			objCoder = new CCoder();
			keyFileName = modActiveLock.ActivelockGetSpecialFolder(46) + "\\rock." + Strings.Chr(98) + Strings.Chr(109) + Strings.Chr(112);

			System.Drawing.Bitmap b = null;
			Resources.ResourceManager r = new Resources.ResourceManager("ActiveLock3_6Net.ProjectResources", System.Reflection.Assembly.GetExecutingAssembly());
			b = r.GetObject("bmp101");
			if (fileExist(keyFileName) == false) b.Save(keyFileName); 
			//If fileExist(keyFileName) = False Then VB6.LoadResPicture(101, VB6.LoadResConstants.ResBitmap).Save(keyFileName)

			objCoder.Decode(keyFileName, FileName, ref a);
			functionReturnValue = a;
			FileSystem.Kill(keyFileName);

		}
		catch (Exception exc) {
		}
		//MsgBox(exc.ToString)
		finally {
			try {
				objCoder.Dispose();
			}
			catch {
			}
		}
		return functionReturnValue;

	}
	//===============================================================================
	// Name: Function ActivateTrial
	// Input:
	//   ByVal SoftwareName As String - Software name. Must not be empty.
	//   ByVal SoftwareVer As String - Software version. Must not be empty.
	//   ByVal TrialType As Long - 0 for No Trial, 1 for Trial Days and 2 for Trial Runs.
	//   ByVal TrialLength As Long - Trial Days or Trial Runs depending on TrialType.
	//   ByRef strMsg As String - Message returned by Activelock
	// Output:
	//   Boolean - True if the Trial license is OK
	// Purpose: This function checks the authenticity and validity of the trial period/runs
	// Remarks: This is the main call to activate the trial feature
	//===============================================================================
	public static bool ActivateTrial(string SoftwareName, string SoftwareVer, int TrialType, int TrialLength, IActiveLock.ALTrialHideTypes TrialHideTypes, ref string strMsg, string SoftwarePassword, IActiveLock.ALTimeServerTypes mCheckTimeServerForClockTampering, IActiveLock.ALSystemFilesTypes mChecksystemfilesForClockTampering, IActiveLock.ALTrialWarningTypes mTrialWarning, 
	ref int mRemainingTrialDays, ref int mRemainingTrialRuns)
	{
		bool functionReturnValue = false;
		 // ERROR: Not supported in C#: OnErrorStatement

		string strVal = null;
		short daysLeft = 0;
		short runsLeft = 0;
		short intEXPIREDWARNING = 0;

		EXPIRED_RUNS = Strings.Chr(101) + Strings.Chr(120) + Strings.Chr(112) + Strings.Chr(105) + Strings.Chr(114) + Strings.Chr(101) + Strings.Chr(100) + Strings.Chr(114) + Strings.Chr(117) + Strings.Chr(110) + Strings.Chr(115);
		EXPIRED_DAYS = Strings.Chr(101) + Strings.Chr(120) + Strings.Chr(112) + Strings.Chr(105) + Strings.Chr(114) + Strings.Chr(101) + Strings.Chr(100) + Strings.Chr(100) + Strings.Chr(97) + Strings.Chr(121) + Strings.Chr(115);
		TEXTMSG_DAYS = DecryptString128Bit("sQvYYRLPon5IyH6BQRAUBuCLTq/5VkH3kl7HUwJLZ2M=", PSWD());
		TEXTMSG_RUNS = DecryptString128Bit("6urN2+xbgqbLLsOoC4hbGpLT3bnvY3YPGW299cOnqfo=", PSWD());
		TEXTMSG = DecryptString128Bit("7MKOm40JXXXc6f3svOrlNeFuWWdWD56E3k7FwvTH2i/oyG2MDdVKjLMcQKsNbfXforIQwFlJEEgCMOfWiSI0sw==", PSWD());

		functionReturnValue = false;
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

		LICENSE_SOFTWARE_NAME = SoftwareName;
		LICENSE_SOFTWARE_VERSION = SoftwareVer;
		LICENSE_SOFTWARE_PASSWORD = SoftwarePassword;

		// Set local variables
		if (TrialType == IActiveLock.ALTrialTypes.trialDays) {
			trialPeriod = true;
		}
		else {
			trialRuns = true;
		}
		if (TrialType == IActiveLock.ALTrialTypes.trialDays) {
			alockDays = TrialLength;
		}
		else {
			alockRuns = TrialLength;
		}

		if (alockDays == 0 & trialPeriod == true) {
			modActiveLock.Set_Locale(modActiveLock.regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrInvalidTrialDays, ACTIVELOCKSTRING, modActiveLock.STRINVALIDTRIALDAYS);
		}
		else if (alockRuns == 0 & trialRuns == true) {
			modActiveLock.Set_Locale(modActiveLock.regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrInvalidTrialRuns, ACTIVELOCKSTRING, modActiveLock.STRINVALIDTRIALRUNS);
		}

		strMsg = "";
		intEXPIREDWARNING = Conversion.Int((double)Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "1"), enc2(TRIALWARNING), enc2(EXPIREDWARNING), (string)0));

		 // ERROR: Not supported in C#: OnErrorStatement

		HAD2HAMMER = false;
		GetSystemTime1();

		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;

		// Check to see if any of the hidden signatures say the trial is expired
		// The following two keys are not compatible with Vista
		// A regular user account cannot have write access to these two registry hives
		// I am removing these from v3.6 - ialkan 12-27-2008
		//If IsRegistryExpired1() = True Then
		//    Set_locale(regionalSymbol)
		//    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG)
		//End If
		//If IsRegistryExpired2() = True Then
		//    Set_locale(regionalSymbol)
		//    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG)
		//End If
		if (IsEncryptedFileExpired() == true) {
			modActiveLock.Set_locale(modActiveLock.regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG);
		}
		// *** We are disabling folder date stamp in v3.2 since it's not application specific ***
		// Well... nothing was found
		// Check the last indicator
		//If IsFolderStampExpired() = True Then
		//    Set_locale(regionalSymbol)
		//    Err.Raise -10100, , TEXTMSG
		//End If

		// Must check Registry for Trial
		if (TrialRegistryPerUserExists(ref TrialHideTypes)) {
			if (IsRegistryExpired() == true) {
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG);
			}
		}

		// Must check picture for Trial
		if (TrialSteganographyExists(ref TrialHideTypes)) {
			if (IsSteganographyExpired() == true) {
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG);
			}
		}

		// Must check folder for Trial
		if (TrialHiddenFolderExists(ref TrialHideTypes)) {
			if (IsHiddenFolderExpired() == true) {
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialInvalid, ACTIVELOCKSTRING, TEXTMSG);
			}
		}

		// Nothing bad so far...
		if (trialPeriod) {
			if (!DateGood(alockDays, ref daysLeft, ref TrialHideTypes)) {
				ExpireTrial(SoftwareName, SoftwareVer, TrialType, TrialLength, TrialHideTypes, SoftwarePassword);
				// Trial Period has expired
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialDaysExpired, ACTIVELOCKSTRING, TEXTMSG_DAYS);
			}
			else {
				if (fileExist(GetSteganographyFile()) == false & Directory.Exists(modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(myDir(), PSWD())) == false & dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90")) == dec2("93.8D.93.8D.96.90.90.90")) {
					if (mCheckTimeServerForClockTampering == IActiveLock.ALTimeServerTypes.alsCheckTimeServer) {
						if (SystemClockTampered()) {
							modActiveLock.Set_locale(modActiveLock.regionalSymbol);
							Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, modActiveLock.STRCLOCKCHANGED);
						}
					}
					if (mChecksystemfilesForClockTampering == IActiveLock.ALSystemFilesTypes.alsCheckSystemFiles) {
						if (ClockTampering()) {
							modActiveLock.Set_locale(modActiveLock.regionalSymbol);
							Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, modActiveLock.STRCLOCKCHANGED);
						}
					}
				}
				// So far so good; trial mode seems to be fine
				HAD2HAMMER = false;
				strMsg = "You are running this program in its Trial Period Mode." + Constants.vbCrLf + (string)daysLeft + " day(s) left out of " + (string)alockDays + " day trial.";
				mRemainingTrialDays = daysLeft;
				functionReturnValue = true;
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
				goto exitGracefully;
			}
		}
		else {
			if (!RunsGood(alockRuns, ref runsLeft, ref TrialHideTypes)) {
				ExpireTrial(SoftwareName, SoftwareVer, TrialType, TrialLength, TrialHideTypes, SoftwarePassword);
				// Trial Runs have expired
				modActiveLock.Set_locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrTrialRunsExpired, ACTIVELOCKSTRING, TEXTMSG_RUNS);
			}
			else {
				if (fileExist(GetSteganographyFile()) == false & Directory.Exists(modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(myDir(), PSWD())) == false & dec2(Interaction.GetSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), "param", "factor1", "93.8D.93.8D.96.90.90.90")) == dec2("93.8D.93.8D.96.90.90.90")) {
					if (mCheckTimeServerForClockTampering == IActiveLock.ALTimeServerTypes.alsCheckTimeServer) {
						if (SystemClockTampered()) {
							modActiveLock.Set_locale(modActiveLock.regionalSymbol);
							Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, modActiveLock.STRCLOCKCHANGED);
						}
					}
					if (ClockTampering()) {
						modActiveLock.Set_locale(modActiveLock.regionalSymbol);
						Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrClockChanged, ACTIVELOCKSTRING, modActiveLock.STRCLOCKCHANGED);
					}
				}
				// So far so good; trial mode seems to be fine
				HAD2HAMMER = false;
				strMsg = "You are running this program in its Trial Runs Mode." + Constants.vbCrLf + (string)runsLeft + " run(s) left out of " + (string)alockRuns + " run trial.";
				mRemainingTrialRuns = runsLeft;
				functionReturnValue = true;
				System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
				goto exitGracefully;
			}
		}
		keepChecking:

		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
		ExpireTrial(SoftwareName, SoftwareVer, TrialType, TrialLength, TrialHideTypes, SoftwarePassword);
		if (Err().Number == -10101) {
			strMsg = TEXTMSG_DAYS;
			mRemainingTrialDays = alockDays;
		}
		else if (Err().Number == -10102) {
			strMsg = TEXTMSG_RUNS;
			mRemainingTrialRuns = alockRuns;
		}
		if (intEXPIREDWARNING == 0 & mTrialWarning == IActiveLock.ALTrialWarningTypes.trialWarningPersistent) {
			Interaction.SaveSetting(enc2(LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD + "1"), enc2(TRIALWARNING), enc2(EXPIREDWARNING), (string)-1);
			strMsg = "Free Trial for this application has ended.";
		}
		functionReturnValue = false;
		return;
		NotRegistered:

		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
		if (Err().Number == 1001) {
			Err().Clear();
			 // ERROR: Not supported in C#: ResumeStatement

		}
		strMsg = Err().Description;
		HAD2HAMMER = true;
		functionReturnValue = false;
		return;
		exitGracefully:
		System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
		return;
		return functionReturnValue;

	}
	public static bool ClockTampering()
	{
		bool functionReturnValue = false;
		string t = null;
		string s = null;
		System.DateTime fileDate = default(System.DateTime);
		short i = 0;
		short Count = 0;
		 // ERROR: Not supported in C#: OnErrorStatement


		for (i = 0; i <= 1; i++) {
			switch (i) {
				case 0:
					t = modActiveLock.WinDir() + "\\Prefetch";
					break;
				case 1:
					t = modActiveLock.WinDir() + "\\Temp";
					break;
				//Case 2
				//    t = WinDir() & "\Temp"
				//Case 3
				//    t = WinDir() & "\Applog"
				//Case 4
				//    t = WinDir() & "\Recent"
			}

			Count = 0;
			s = FileSystem.Dir(t + "\\*.*");
			while (!string.IsNullOrEmpty(s)) {
				if (Strings.Left(s, 1) != "$" & Strings.Left(s, 1) != "?") {
					fileDate = FileSystem.FileDateTime(t + "\\" + s);
					long difHours = 0;
					difHours = Math.Abs(((System.DateTime)fileDate.Date.ToString("yyyy/MM/dd")).Subtract((System.DateTime)System.DateTime.UtcNow.ToString("yyyy/MM/dd")).Hours);
					if (difHours > 24) {
						if (Count > 1) {
							functionReturnValue = true;
							return;
						}
						else {
							//Forgiveness for one file only - for now
							Count = Count + 1;
						}
					}
				}
				s = FileSystem.Dir();
			}
		}
		return functionReturnValue;

	}
	private static object GetSteganographyFile()
	{
		object functionReturnValue = null;
		string strSource = null;
		functionReturnValue = null;
		try {
			strSource = modActiveLock.ActivelockGetSpecialFolder(54) + "\\Sample Pictures" + DecryptString128Bit("Qspq9Tu3sG/IE+ugm+o1RQ==", PSWD()) + Strings.Left(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8) + "." + Strings.Chr(98) + Strings.Chr(109) + Strings.Chr(112);
			System.Drawing.Bitmap b = null;
			System.Resources.ResourceManager r = new System.Resources.ResourceManager("ActiveLock3_6Net.ProjectResources", System.Reflection.Assembly.GetExecutingAssembly());
			b = r.GetObject("bmp101");
			if (Directory.Exists(modActiveLock.ActivelockGetSpecialFolder(54) + "\\Sample Pictures") == false) {
				ActiveLock3_6NET.My.MyProject.Computer.FileSystem.CreateDirectory(modActiveLock.ActivelockGetSpecialFolder(54) + "\\Sample Pictures");
			}
			if (fileExist(strSource) == false) {
				b.Save(strSource);
			}
			//If fileExist(strSource) = False Then VB6.LoadResPicture(101, VB6.LoadResConstants.ResBitmap).Save(strSource)
			functionReturnValue = strSource;
		}
		catch (Exception ex) {
			modActiveLock.Set_locale(modActiveLock.regionalSymbol);
			Err().Raise(Err().Number, Err().Source, Err().Description, Err().HelpFile, Err().HelpContext);
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function ReadUntil
	// Input:
	//   ByRef sIn As String - Input string to be read
	//   ByRef sDelim As String - Delimiter string
	//   ByRef bCompare As VbCompareMethod - Optional flag for string comparison
	// Output:
	//   String - The new string read
	// Purpose: Read the contents of a string until the given delimiter is encountered
	// and returns the new string
	// Remarks: None
	//===============================================================================
	public static string ReadUntil(ref string sIn, ref string sDelim, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(CompareMethod.Binary)] ref  // ERROR: Optional parameters aren't supported in C#
CompareMethod bCompare)
	{
		string functionReturnValue = null;
		string nPos = null;
		functionReturnValue = string.Empty;

		nPos = (string)Strings.InStr(1, sIn, sDelim, bCompare);
		if ((double)nPos > 0) {
			functionReturnValue = Strings.Left(sIn, (double)nPos - 1);
			sIn = Strings.Mid(sIn, (double)nPos + Strings.Len(sDelim));
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function InStrRev
	// Input:
	//   ByVal sIn As String - The main string to be searched
	//   ByRef sFind As String - Search string or substring being searched
	//   ByRef nStart As Long - Optional starting character position in the main string
	//   ByRef bCompare As VbCompareMethod - Optional flag for string comparison
	// Output:
	//   Long - The position of the character found. Returns 0 if not found.
	// Purpose: This is the reverse mode operation of the VB InStr function
	// Remarks: None
	//===============================================================================
	//InStrRev was upgraded to InStrRev_Renamed
	public static int InStrRev_Renamed(string sIn, ref string sFind, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(1)] ref  // ERROR: Optional parameters aren't supported in C#
int nStart, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(CompareMethod.Binary)] ref  // ERROR: Optional parameters aren't supported in C#
CompareMethod bCompare)
	{
		int functionReturnValue = 0;
		int nPos = 0;
		sIn = StrReverse_Renamed(sIn);
		sFind = StrReverse_Renamed(sFind);
		nPos = Strings.InStr(nStart, sIn, sFind, bCompare);
		if (nPos == 0) {
			functionReturnValue = 0;
		}
		else {
			functionReturnValue = Strings.Len(sIn) - nPos - Strings.Len(sFind) + 2;
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function StrReverse
	// Input:
	//   ByVal sIn As String - Input string
	// Output:
	//   String - Reversed string
	// Purpose: Reverses a given string
	// Remarks: None
	//===============================================================================
	//StrReverse was upgraded to StrReverse_Renamed
	public static string StrReverse_Renamed(string sIn)
	{
		short nC = 0;
		string sOut = null;
		for (nC = Strings.Len(sIn); nC >= 1; nC += -1) {
			sOut = sOut + Strings.Mid(sIn, nC, 1);
		}
		return sOut;
	}
	//===============================================================================
	// Name: Function Replace
	// Input:
	//   ByRef sIn As String - The main string to be searched
	//   ByRef sFind As String - Search string or substring being searched
	//   ByRef sReplace As String - String to be used as replacement
	//   ByRef nStart As Long - Optional starting character position in the main string
	//   ByRef nCount As Long - Optional number of characters to be searched
	//   ByRef bCompare As VbCompareMethod - Optional flag for string comparison
	// Output:
	//   String - Modified string
	// Purpose: Finds a substring in a given string and replaces it with another string
	// Remarks: None
	//===============================================================================
	//Replace was upgraded to Replace_Renamed
	public static string Replace_Renamed(ref string sIn, ref string sFind, ref string sReplace, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(1)] ref  // ERROR: Optional parameters aren't supported in C#
int nStart, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(-1)] ref  // ERROR: Optional parameters aren't supported in C#
int nCount, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(CompareMethod.Binary)] ref  // ERROR: Optional parameters aren't supported in C#
CompareMethod bCompare)
	{
		int nC = 0;
		short nPos = 0;
		string sOut = null;
		sOut = sIn;
		nPos = Strings.InStr(nStart, sOut, sFind, bCompare);
		if (nPos == 0) goto EndFn; 
		do {
			nC = nC + 1;
			sOut = Strings.Left(sOut, nPos - 1) + sReplace + Strings.Mid(sOut, nPos + Strings.Len(sFind));
			if (nCount != -1 & nC >= nCount) break; // TODO: might not be correct. Was : Exit Do
 
			nPos = Strings.InStr(nStart, sOut, sFind, bCompare);
		}
		while (nPos > 0);
		EndFn:
		return sOut;
	}
	//===============================================================================
	// Name: Function TrimSpaces
	// Input:
	//   ByVal strString As String - Input string to be trimmed
	// Output:
	//   String - Trimmed string
	// Purpose: Removes all spaces from a string
	// Remarks: None
	//===============================================================================
	public static string TrimSpaces(string strString)
	{
		int lngpos = 0;
		while (Strings.InStr(1, strString, " ")) {
			System.Windows.Forms.Application.DoEvents();
			lngpos = Strings.InStr(1, strString, " ");
			strString = Strings.Left(strString, lngpos - 1) + Strings.Right(strString, Strings.Len(strString) - (lngpos + Strings.Len(" ") - 1));
		}
		return strString;
	}
	//===============================================================================
	// Name: Function Scramb
	// Input:
	//   ByVal strString As String - String to be scrambled
	// Output:
	//   String - Scrambled string
	// Purpose: Scrambles a string
	// Remarks: None
	//===============================================================================
	public static string Scramb(string strString)
	{
		short i = 0;
		string even = string.Empty;
		string odd = string.Empty;
		for (i = 1; i <= Strings.Len(strString); i++) {
			if (i % 2 == 0) {
				even = even + Strings.Mid(strString, i, 1);
			}
			else {
				odd = odd + Strings.Mid(strString, i, 1);
			}
		}
		return even + odd;
	}
	//===============================================================================
	// Name: Function dhTrimNull
	// Input:
	//   ByVal strValue As String - Input string
	// Output:
	//   String - Trimmed string
	// Purpose: Removes the leading null in a string
	// Remarks: Useful for API calls
	//===============================================================================
	public static string dhTrimNull(string strValue)
	{
		string functionReturnValue = null;
		short intPos = 0;
		functionReturnValue = string.Empty;

		intPos = Strings.InStr(strValue, Constants.vbNullChar);
		switch (intPos) {
			case 0:
				functionReturnValue = strValue;
				break;
			case 1:
				functionReturnValue = Constants.vbNullString;
				break;
			case  // ERROR: Case labels with binary operators are unsupported : GreaterThan
1:
				functionReturnValue = Strings.Left(strValue, intPos - 1);
				break;
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function Unscramb
	// Input:
	//   ByVal strString As String - Scrambled string
	// Output:
	//   String - Unscrambled string
	// Purpose: Unscrambles a string scrambled by Scramb
	// Remarks: None
	//===============================================================================
	public static string Unscramb(string strString)
	{
		short evenint = 0;
		short x = 0;
		short oddint = 0;
		string odd = null;
		string even = null;
		string fin = string.Empty;

		x = Strings.Len(strString);
		x = Conversion.Int(Strings.Len(strString) / 2);
		//adding this returns the actual number like 1.5 instead of returning 2
		even = Strings.Mid(strString, 1, x);
		odd = Strings.Mid(strString, x + 1);
		for (x = 1; x <= Strings.Len(strString); x++) {
			if (x % 2 == 0) {
				evenint = evenint + 1;
				fin = fin + Strings.Mid(even, evenint, 1);
			}
			else {
				oddint = oddint + 1;
				fin = fin + Strings.Mid(odd, oddint, 1);
			}
		}
		return fin;
	}
	//===============================================================================
	// Name: Function WindowsPath
	// Input: None
	// Output: None
	// Purpose: Gets Windows directory path
	// Remarks: None
	//===============================================================================
	public static string WindowsPath()
	{
		string functionReturnValue = null;
		string sPath = null;
		sPath = Strings.Space(255);
		GetWindowsDirectory(sPath, 255);
		functionReturnValue = dhTrimNull(sPath);
		if (Strings.Right(WindowsPath(), 1) != "\\") {
			functionReturnValue = functionReturnValue + "\\";
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function DoScan
	// Input: None
	// Output:
	//   Boolean - True if a debugger or file monitor is found
	// Purpose: Will scan the memory for common debuggers or File Monitoring software
	// Remarks: None
	//===============================================================================

	public static bool DoScan()
	{
		bool functionReturnValue = false;
		int hFile = 0;
		int retVal = 0;
		string sScan = null;
		// Dim sBuffer As String

		string sRegMonClass = null;
		string sFileMonClass = null;
		//\\We break up the class names to avoid
		//     detection in a hex editor
		sRegMonClass = "R" + "e" + "g" + "m" + "o" + "n" + "C" + "l" + "a" + "s" + "s";
		sFileMonClass = "F" + "i" + "l" + "e" + "M" + "o" + "n" + "C" + "l" + "a" + "s" + "s";
		//\\See if RegMon or FileMon are running

		switch (true) {
			case FindWindow(sRegMonClass, Constants.vbNullString) != 0:
				//Regmon is running...throw an access vio
				//     lation
				functionReturnValue = true;
				return;

				break;
			case FindWindow(sFileMonClass, Constants.vbNullString) != 0:
				//FileMon is running...throw an access vi
				//     olation
				functionReturnValue = true;
				return;

				break;
		}
		//\\So far so good...check for SoftICE in memory
		hFile = CreateFile("\\\\.\\SICE", GENERIC_WRITE | GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);

		if (hFile != -1) {
			// SoftICE is detected.
			retVal = CloseHandle(hFile);
			// Close the file handle
			functionReturnValue = true;
			return;
		}
		else {
			// SoftICE is not found for windows 9x, check for NT.
			hFile = CreateFile("\\\\.\\NTICE", GENERIC_WRITE | GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0);

			if (hFile != -1) {
				// SoftICE is detected.
				retVal = CloseHandle(hFile);
				// Close the file handle
				functionReturnValue = true;
				return;
			}
		}

		sScan = "f" + "i" + "l" + "e" + "W" + "A" + "T" + "C" + "H";

		if (FindWindow(Constants.vbNullString, sScan) != 0) {
			functionReturnValue = true;
			return;
		}

		sScan = "F" + "i" + "l" + "e" + "S" + "p" + "y";
		if (FindWindow(Constants.vbNullString, sScan) != 0) {
			functionReturnValue = true;
			return;
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Sub GetSystemTime
	// Input: None
	// Output: None
	// Purpose: 'This is the main debugger detection routine. Name is misleading !!
	// Remarks: None
	//===============================================================================
	public static void GetSystemTime1()
	{
		//This is the main debugger detection routine.
		string sFc = null;
		string src = null;
		string str1 = null;
		string vernumber = string.Empty;
		short vn2 = 0;
		short vn0 = 0;
		short vn1 = 0;
		short vnx = 0;

		src = DecryptString128Bit("kSPnKBkg9LARQZgo61A3ow==", PSWD());
		//RegmonClass
		sFc = DecryptString128Bit("SWRcv8sZJnnQDZHkS+Ewnw==", PSWD());
		//FileMonClass

		//Check For RegMon
		if (FinWin(src, Constants.vbNullString) != 0) {
			HAD2HAMMER = true;
			return;
		}
		if (FinWin(sFc, Constants.vbNullString) != 0) {
			HAD2HAMMER = true;
			return;
		}

		//Look For Threats via VxD..
		CTV(ref DecryptString128Bit("bVW9MtZbzBzH3f8HeSCO9g==", PSWD()));
		//SICE
		CTV(ref DecryptString128Bit("Jf4nLowTbc8fqW0lTNQ/6Q==", PSWD()));
		//NTICE
		CTV(ref DecryptString128Bit("OO6LYadZGIJ57F7HphxBPw==", PSWD()));
		//SIWDEBUG
		CTV(ref DecryptString128Bit("Bv5XEuseX6OXpMbpZbV+cg==", PSWD()));
		//SIWVID

		//Look For Threats using titles of windows !!!!!!!!!!!!!!!!!!!

		//W32dasm (other than main window)
		CTW(ref DecryptString128Bit("aDHIGwWBTB0qErkuL9yzyd3jTfPbCStMz2Kpd8WgMq8=", PSWD()));
		//Win32Dasm "Goto Code Location (32 Bit)"
		//SoftICE variants
		CTW(ref ActiveLock3_6NET.My.MyProject.Application.Info.DirectoryPath + "\\" + System.IO.Path.GetFileName(Application.ExecutablePath) + ".EXE" + DecryptString128Bit("28u3ww4pEzuawNHeN8kJWscUrblvk5v4Af2gaVLRD0M=", PSWD()));
		//SoftIce; [app_path]+" - Symbolic Loader"
		CTW(ref ActiveLock3_6NET.My.MyProject.Application.Info.DirectoryPath + "\\" + System.IO.Path.GetFileName(Application.ExecutablePath) + ".EXE" + DecryptString128Bit("jk980I+TMZFGIcGEFPSxCJydPCzofApnewH3ORNKRQI=", PSWD()));
		//SoftIce; [app_path]+" - Symbol Loader"
		//CTW(VB6.GetPath & "\" & VB6.GetEXEName() & ".EXE" & DecryptString128Bit("28u3ww4pEzuawNHeN8kJWscUrblvk5v4Af2gaVLRD0M=", PSWD)) 'SoftIce; [app_path]+" - Symbolic Loader"
		//CTW(VB6.GetPath & "\" & VB6.GetEXEName() & ".EXE" & DecryptString128Bit("jk980I+TMZFGIcGEFPSxCJydPCzofApnewH3ORNKRQI=", PSWD)) 'SoftIce; [app_path]+" - Symbol Loader"
		CTW(ref DecryptString128Bit("4WuQVAq9HLlmPg1lsTQN7m9vintF1mbSLUERgq1nONo=", PSWD()));
		//"NuMega SoftICE Symbol Loader"

		//Checks for URSoft W32Dasm app windows versions 0.0x - 12.9x
		str1 = DecryptString128Bit("PhxMXRKD5iLJ6eNYZBKeg3BRdQyRx+fa0S1RijBZRJ4=", PSWD()) + vernumber + DecryptString128Bit("SOsZIVbDRg8v0mcr7pD9w1xrKHCI3xFpof9LET3zFv4=", PSWD());
		for (vn0 = 12; vn0 >= 0; vn0 += -1) {
			for (vn1 = 9; vn1 >= 0; vn1 += -1) {
				for (vn2 = 9; vn2 >= 0; vn2 += -1) {
					vnx = (short)vn1 + vn2;
					vernumber = vn0 + "." + vnx;
					//Check for "URSoft W32Dasm Ver " & vernumber & " Program Disassembler/Debugger"
					CTW(ref str1);
				}
			}
		}

		//Check for step debugging (light check)
		CSD();

		//Check for processes and wipe from 200000 to N amount of bytes in steps of 48
		//(to aggressively screw with the code)
		//RefreshProcessList()
		CFP(ref DecryptString128Bit("OiOALAEiNTL0BDcDFRa5+A==", PSWD()), ref 2000000);
		//Kill "Debuggy By Vanja Fuckar" - Debuggy.exe
		CFP(ref DecryptString128Bit("3idyYtU1s0wY5HvE9AztZQ==", PSWD()), ref 2000000);
		//Kill "OllyDBG" - OLLYDBG.exe
		CFP(ref DecryptString128Bit("2UqaM1jVTgltRHWy59DcyA==", PSWD()), ref 2000000);
		//Kill "ProcDump by G-Rom, Lorian & Stone" - PROCDUMP.exe
		CFP(ref DecryptString128Bit("cXf/bBlwI4BzR2+996P6rw==", PSWD()), ref 2000000);
		//Kill "SoftSnoop by Yoda/f2f" - SoftSnoop.exe
		CFP(ref DecryptString128Bit("kjLerFCcQQ3qcRsJNjb5qA==", PSWD()), ref 2000000);
		//Kill "TimeFix by GodsJiva" - TimeFix.exe
		CFP(ref DecryptString128Bit("lCu+rzLgHvPml0ceQaYnc6Ga/wmZDry4uyL8UmPQVdk=", PSWD()), ref 2000000);
		//Kill "TMR Ripper Studi" - "TMG Ripper Studio.exe"

		//Send the user through a jungle of conditional branches.
		//Hopefully now timefix will be disabled.
		JOC();

		//============ END OF CHECKS ===========
		//Most amateur crackers should have had Win32Dasm shut down by now.
		//If using step-debugging, this app should have given an exception.
		//
		//===== BEFORE RELEASING YOUR EXE =====
		//Use UPX to pack it.  Change the PE header in the file using
		//a hex editor.  (It will stop lamers from being able to use
		//the -d switch with UPX to unpack your program)
		//REMEMBER: Someone will always be able to crack your program!!
		//Delaying crackers is the best you can hope for.

		//Final CRC check on our strings...
		//Remove this following msgbox line if you need to check the CRC
		//and then change the number below to that. This will detect
		//if the user has lamely changed the values
		//we're checking using a hex-editor!!!...
		//--------------------------------------------------------

		if (GC() == 27514) HAD2HAMMER = false; 
		// One final check
		if (DoScan() == true) HAD2HAMMER = true; 

	}
	//===============================================================================
	// Name: Sub CTV
	// Input:
	//   ByRef appid as String
	// Output: None
	// Purpose: Checks threats vxd
	// Remarks: None
	//===============================================================================
	public static void CTV(ref string appid)
	{
		//Check threats vxd
		if (CreateFile("\\\\.\\" + appid, GENERIC_WRITE | GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0) != -1) {
			retVal = CloseHandle(hFile);
			// Close the file handle
			HAD2HAMMER = true;
		}
	}
	//===============================================================================
	// Name: Sub CFP
	// Input:
	//   ByRef procname as String
	//   ByRef hammerrange as Variant
	// Output: None
	// Purpose: [INTERNAL] Process killer
	// Remarks: None
	//===============================================================================
	public static void CFP(ref string procname, ref object hammerrange)
	{
		short xx = 0;
		for (xx = 0; xx <= 256; xx++) {
			if (Strings.LCase(procname) == Strings.LCase(ProcessName[xx])) HAMMERPROCESS(ref (int)ProcessID[xx], ref hammerrange); 
		}

	}
	//===============================================================================
	// Name: Sub HAMMERPROCESS
	// Input:
	//   ByRef PID As Long
	//   ByRef hammertop As Variant
	// Output: None
	// Purpose: Process killer.
	// Remarks: None
	//===============================================================================
	public static void HAMMERPROCESS(ref int PID, ref object hammertop)
	{
		if (!InitProcess(ref PID)) {
			//MsgBox "Failed shutdown"
		}
		int p = 0;
		int addr = 0;
		int l = 0;
		for (p = 20000; p <= hammertop; p += 48) {
			addr = (int)Conversion.Val(((string)p).Trim());
			WriteProcessMemory(myHandle, addr, "6", 1, ref l);
		}
		HAD2HAMMER = true;
	}
	//===============================================================================
	// Name: Sub CTW
	// Input:
	//   ByRef winid As String
	// Output: None
	// Purpose: Process killer. Freezes a window.
	// Remarks: None
	//===============================================================================
	public static void CTW(ref string winid)
	{
		int WID = 0;
		short FLDWIN = 0;
		WID = FindWindow(Constants.vbNullString, winid);
		if (FindWindow(Constants.vbNullString, winid) > 0) {
			//Just sending &H10 closes the window.. but this method freezes it and closes apps where they are usually protected from an external shut down!! ;)
			for (FLDWIN = 0; FLDWIN <= 255; FLDWIN++) {
				PostMessage(WID, FLDWIN, 0, ref 0);
				if (FLDWIN > 16) {
					PostMessage(WID, 0x10, 0, ref 0);
				}
			}
			HAD2HAMMER = true;
		}
	}
	//===============================================================================
	// Name: Function CSD
	// Input: None
	// Output:
	//   Booleans - Does not use the return value
	// Purpose: Process killer. Checks for step debuggers
	// Remarks: None
	//===============================================================================
	public static bool CSD()
	{
		float Timer_start = 0;
		float Timer_time = 0;
		short s = 0;
		//Check for Step Debugger
		Timer_start = DateAndTime.Timer();
		for (s = 1; s <= 25; s++) {
			PSub();
			//Pointless Sub
			PFunction(ref s + Conversion.Int(VBMath.Rnd() * 20));
			//Pointless Function
		}
		Timer_time = DateAndTime.Timer() - Timer_start;

		//Step-debugging Detected...
		if (Timer_time > 1) {
			HAD2HAMMER = true;
		}
	}
	//===============================================================================
	// Name: Sub PSub
	// Input: None
	// Output: None
	// Purpose: Processes some garbage
	// Remarks: None
	//===============================================================================
	public static void PSub()
	{
		object X2 = null;
		object X1 = null;
		object X3 = null;
		//Just some garbage processing...
		System.Windows.Forms.Application.DoEvents();
		X1 = System.Math.Sqrt(65536);
		X2 = Math.Pow(16, 2);
		X3 = X1 - X2;
		X1 = X2 + X3;
		X3 = X2;
	}
	//===============================================================================
	// Name: Function PFunction
	// Input:
	//   ByRef PointlessVariable As Integer - Dummy integer
	// Output: None
	// Purpose: Processes some garbage
	// Remarks: None
	//===============================================================================
	public static object PFunction(ref short PointlessVariable)
	{
		object X2 = null;
		object X1 = null;
		object X3 = null;
		//Just some garbage processing...
		System.Windows.Forms.Application.DoEvents();
		X1 = System.Math.Sqrt(256);
		X2 = Math.Pow(8, 2);
		X3 = X1 + PointlessVariable;
		X1 = X1 + X2 + X3;
		return null;

	}
	//===============================================================================
	// Name: Sub JOC
	// Input: None
	// Output: None
	// Purpose: This is designed to trap programmers stepping through the code
	// <p>Due to time-sensitivity, people stepping through this code
	// will probably find the program ends up closing itself thanks
	// to the timer on the main form.  If the time taken to go through
	// these conditions is too high, Form1's height and width will be set
	// to zero.  The resize event on Form1 detects the abnormal
	// zero -Height And closes the application down.
	// <p>It's basically a more complex version of the 'CSD' routine..
	// <p>To test it.. add a breakpoint on line 5 and step through
	// using Shift+F8... The app will either close, crash or they'll be in
	// an infinite loop.
	// Remarks: None
	//===============================================================================

	public static void JOC()
	{
		//Horrible Sloppy Code but it should help to throw some lamers off..
		//Start off with some fake math, arrays, etc and throw a few pointless encrypted strings in there
		VBMath.Randomize(32);
		object[] JU_C_OR = new object[33];
		object[,] ViT = new object[2, 2];
		object c = null;
		object AMIN = null;
		object tang = null;
		object App_Les = null;
		object yergisiz = null;
		object ang = null;
		object gerinmeoyle = null;
		object PE_ar = null;
		object e = null;
		// Dim TXM, TM, TXD, TXZ As Single

		ViT[0, 0] = "yalnizlikzorseydostum";
		AMIN = 1;
		tang = 12;
		c = 0;
		gerinmeoyle = "eT5XeXeXXeT1MeUX11cU";
		ang = tang - 2;
		App_Les = Conversion.Int(VBMath.Rnd() * 6);
		c = 1 - c;
		PE_ar = Conversion.Int(VBMath.Rnd() * 100) + Strings.Asc(Strings.Mid("saygisizbiradam", 5, 1));
		AMIN = 1 - AMIN;
		JU_C_OR[ang + e] = (int)ViT[AMIN, c];
		yergisiz = PE_ar + JU_C_OR[ang + e] + App_Les + ("yalnizlikzor" + gerinmeoyle);

		//Now a pile of pointless conditions...
		//This is designed to trap programmers stepping through the code
		//
		//Due to time-sensitivity, people stepping through this code
		//will probably find the program ends up closing itself thanks
		//to the timer on the main form.  If the time taken to go through
		//these conditions is too high, Form1's height and width will be set
		//to zero.  The resize event on Form1 detects the abnormal
		//zero -Height And closes the application down.
		//
		//It's basically a more complex version of the 'CSD' routine..

		//To test it.. add a breakpoint on line 5 and step through
		//using Shift+F8... The app will either close, crash or they'll be in
		//an infinite loop.

		//5:
		//        TM = VB.Timer()
		//10:
		//        If AMIN = 0 Then GoTo 30 Else GoTo 70
		//25:     System.Windows.Forms.Application.DoEvents()
		//20:
		//        If PE_ar > AMIN Then GoTo 25 Else GoTo 30
		//30:
		//        If c = 2 Then wX = -800 : GoTo 60 Else GoTo 40
		//40:
		//        c = c + 1
		//        TXD = VB.Timer() : GoTo 60
		//50:
		//        If PE_ar + ang = AMIN Then GoTo 40 Else GoTo 95
		//60:
		//        AMIN = 0 : wY = -2000
		//        If TXD - TM < c Then GoTo 80 Else GoTo 170
		//800:
		//        If frmC.DefInstance.Timer1.Enabled = True Then
		//            TXM = 1
		//        Else
		//            TXM = 0 : GoTo 890
		//        End If
		//70:
		//        AMIN = AMIN + 1 : GoTo 20
		//80:
		//        If AMIN > 1 Then GoTo 20
		//190:    wX = 4800 : wY = 3600 : GoTo 1000 'here we set the window width so that it's no longer 0,0
		//75:
		//        If App_Les > 16 Then
		//            GoTo 190
		//        Else
		//            If App_Les > 256 Then GoTo 140
		//        End If
		//1000:
		//        If VB.Timer() > TM + 2 Then wY = 0 Else Call frmC.DefInstance.frmC_Resize(frmC.DefInstance, New System.EventArgs) : Exit Sub
		//1001:   Call frmC.DefInstance.frmC_Resize(frmC.DefInstance, New System.EventArgs) : Exit Sub
		//90:     GoTo 80
		//95:     wY = 50 : GoTo 800
		//140:    If wY = 360 Then
		//            wY = 0
		//            TM = 20 : GoTo 60
		//        End If
		//125:    GoTo 150
		//890:
		//        wX = wX * TXM
		//        wY = wY * TXM : GoTo 1000
		//120:    GoTo 1000
		//170:
		//        TXZ = TXZ + 1
		//        If TXZ > 50 Then GoTo 135 Else GoTo 175
		//175:    If frmC.DefInstance.Timer1.Enabled = False Then GoTo 135 Else GoTo 170
		//160:    If wX = 30 Then GoTo 170 Else GoTo 20
		//150:    If wX = 0 Then GoTo 170 Else GoTo 130
		//130:    GoTo 95
		//135:    HAD2HAMMER = True
	}
	//===============================================================================
	// Name: Function InitProcess
	// Input:
	//   ByRef PID As Long - Process ID
	// Output:
	//   Variant - Returns true if the process was initialized
	// Purpose: Initializes a process
	// Remarks: None
	//===============================================================================
	public static object InitProcess(ref int PID)
	{
		object functionReturnValue = null;
		int pHandle = 0;
		int myHandle = 0;
		pHandle = OpenProcess(0x1f0fff, false, PID);

		if ((pHandle == 0)) {
			functionReturnValue = false;
			myHandle = 0;
		}
		else {
			functionReturnValue = true;
			myHandle = pHandle;
		}
		return functionReturnValue;

	}
	//===============================================================================
	// Name: Sub RefreshProcessList
	// Input: None
	// Output: None
	// Purpose: Reads Process List and Fills combobox (cboProcess)
	// Remarks: None
	//===============================================================================
	public static void RefreshProcessList()
	{
		//Reads Process List and Fills combobox (cboProcess)

		PROCESSENTRY32 myProcess = null;
		int mySnapshot = 0;
		short xx = 0;

		//first clear our combobox
		myProcess.dwSize = Strings.Len(myProcess);

		//create snapshot
		mySnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, ref 0);

		//clear array
		for (xx = 0; xx <= 256; xx++) {
			ProcessName[xx] = "";
		}

		xx = 0;
		//get first process
		ProcessFirst(mySnapshot, ref myProcess);
		ProcessName[xx] = Strings.Left(myProcess.szexeFile, Strings.InStr(1, myProcess.szexeFile, Strings.Chr(0)) - 1);
		// set exe name
		ProcessID[xx] = myProcess.th32ProcessID;
		// set PID

		//while there are more processes
		while (ProcessNext(mySnapshot, ref myProcess)) {
			xx = xx + 1;
			ProcessName[xx] = Strings.Left(myProcess.szexeFile, Strings.InStr(1, myProcess.szexeFile, Strings.Chr(0)) - 1);
			// set exe name
			ProcessID[xx] = myProcess.th32ProcessID;
			// set PID
		}

	}
	//===============================================================================
	// Name: Function GC
	// Input: None
	// Output:
	//   Variant
	// Purpose: Get CRC of all strings to check if they've been modified
	// Remarks: None
	//===============================================================================
	public static object GC()
	{
		short p = 0;
		short encvars = 0;
		short mycrc = 0;
		//Get CRC of all strings to check if they've been modified
		for (encvars = 0; encvars <= 4000; encvars++) {
			for (p = 1; p <= Strings.Len(encvar[encvars]); p++) {
				mycrc = mycrc + Strings.Asc(Strings.Mid(encvar[encvars], p, 1));
				if (mycrc > 30000) mycrc = mycrc - 30000; 
			}
		}
		return mycrc;
	}
	private static void CreateHdnFile()
	{
		INIFile mINIFile = new INIFile();
		if (fileExist(HiddenFile()) == false) {
			{
				mINIFile.File = HiddenFile();
				mINIFile.Section = ".ShellClassInfo";
				mINIFile.Values("CLSID") = CLSIDSTR;
			}
		}
	}
	private static string HiddenFolderFunction()
	{
		string myFile = null;
		myFile = Strings.Chr(92) + Strings.Chr(82) + Strings.Chr(101) + Strings.Chr(99) + Strings.Chr(121) + Strings.Chr(99) + Strings.Chr(108) + Strings.Chr(101) + Strings.Chr(100) + Strings.Left(modMD5.ComputeHash(ref LICENSE_SOFTWARE_NAME + LICENSE_SOFTWARE_VERSION + LICENSE_SOFTWARE_PASSWORD), 8) + "." + Strings.Chr(108) + Strings.Chr(115) + Strings.Chr(116);
		return modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(HIDDENFOLDER, PSWD()) + myFile;
	}
	private static string HiddenFile()
	{
		string KEYFILE = null;
		KEYFILE = Strings.Chr(92) + Strings.Chr(68) + Strings.Chr(101) + Strings.Chr(115) + Strings.Chr(107) + Strings.Chr(116) + Strings.Chr(111) + Strings.Chr(112) + "." + Strings.Chr(105) + Strings.Chr(110) + Strings.Chr(105);
		return modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(HIDDENFOLDER, PSWD()) + KEYFILE;
	}
	private static void MinusAttributes()
	{
		 // ERROR: Not supported in C#: OnErrorStatement

		double ok = 0;
		//Ok, the folder is there, let's change its attributes
		ok = Interaction.Shell("ATTRIB -h -s -r " + modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(HIDDENFOLDER, PSWD()), AppWinStyle.Hide);
        minusAttributesError;
	}
	private static void PlusAttributes()
	{
		double ok = 0;
		//Ok, the folder is there, let's change its attributes
		ok = Interaction.Shell("ATTRIB +h +s +r " + modActiveLock.ActivelockGetSpecialFolder(46) + DecryptString128Bit(HIDDENFOLDER, PSWD()), AppWinStyle.Hide);
	}

	public static bool SystemClockTampered()
	{
		bool functionReturnValue = false;
		// Section added by Ismail Alkan
		// Access a good time server to see which day it is :)
		// Get the date only... compare with the system clock
		// Die if more than 1 day difference'
		// UTC Time and Date of course...

		// Obviously, for this function to work, there must be a connection to Internet
		if (modActiveLock.IsWebConnected() == false) {
			modActiveLock.Set_locale(modActiveLock.regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrInternetConnectionError, ACTIVELOCKSTRING, modActiveLock.STRINTERNETNOTCONNECTED);
		}

		string ss = null;
		string aa = null;
		string blabla = null;
		short diff = 0;
		short i = 0;
		string[] month1 = null;
		string[] month2 = null;
		functionReturnValue = false;
		month1 = "January;February;March;April;May;June;July;August;September;October;November;December".Split(";");
		month2 = "01;02;03;04;05;06;07;08;09;10;11;12".Split(";");
		ss = OpenURL(DecryptString128Bit("xcZfCWJLnPOl2V5kTJpBOi4ysgfzBj1H3nUzyhxODITU7s9QX0xe23TQ9ue3ypmT", PSWD()));
		//http://www.time.gov/timezone.cgi?UTC/s/0
		if (string.IsNullOrEmpty(ss)) {
			// The fast method above did not do its job
			// Call an existing Daytime class from "The Code Project"
			// to check if the system clock was adjusted
			Daytime systemClock = new Daytime();
			if (Daytime.WindowsClockIncorrect() == true) {
				return true;
			}
		}
		blabla = "</b></font><font size=" + "\"" + "5" + "\"" + " color=" + "\"" + "white" + "\"" + ">";
		i = Strings.InStr(ss, blabla);
		ss = Strings.Mid(ss, i + Strings.Len(blabla));
		i = Strings.InStr(ss, "<br>");
		ss = Strings.Left(ss, i - 1);
		i = Strings.InStr(1, ss, ",");
		ss = Strings.Mid(ss, i + 1);
		ss = Replace_Renamed(ss, ",", " ");
		ss = ss.Trim();
		for (i = 0; i <= 11; i++) {
			if (Strings.InStr(ss, month1[i])) {
				ss = ss.Replace(month1[i], month2[i]);
				break; // TODO: might not be correct. Was : Exit For
			}
		}
		ss = Strings.Format((System.DateTime)ss, "yyyy/MM/dd");
		//"short date")
		aa = System.DateTime.UtcNow.Year.ToString("0000") + "/" + System.DateTime.UtcNow.Month.ToString("00") + "/" + System.DateTime.UtcNow.Day.ToString("00");
		diff = ((System.DateTime)ss).Subtract((System.DateTime)aa).Days;
		if (diff > 1) functionReturnValue = true; 
		return functionReturnValue;

	}
	public static int VarPtr(object o)
	{
		// Undocumented VarPtr in VB.NET !!!
		System.Runtime.InteropServices.GCHandle GC = System.Runtime.InteropServices.GCHandle.Alloc(o, System.Runtime.InteropServices.GCHandleType.Pinned);
		int ret = GC.AddrOfPinnedObject().ToInt32();
		GC.Free();
		return ret;
	}

	public static string EncryptString128Bit(string vstrTextToBeEncrypted, string vstrEncryptionKey)
	{
		byte[] bytValue = null;
		byte[] bytKey = null;
		byte[] bytEncoded = null;
		byte[] bytIV = { 121, 241, 10, 1, 132, 74, 11, 39, 255, 91, 
		45, 78, 14, 211, 22, 62 };
		int intLength = 0;
		int intRemaining = 0;
		MemoryStream objMemoryStream = new MemoryStream();
		CryptoStream objCryptoStream = null;
		RijndaelManaged objRijndaelManaged = null;


		//   **********************************************************************
		//   ******  Strip any null character from string to be encrypted    ******
		//   **********************************************************************

		vstrTextToBeEncrypted = StripNullCharacters(vstrTextToBeEncrypted);

		//   **********************************************************************
		//   ******  Value must be within ASCII range (i.e., no DBCS chars)  ******
		//   **********************************************************************

		bytValue = Encoding.ASCII.GetBytes(vstrTextToBeEncrypted.ToCharArray());

		intLength = Strings.Len(vstrEncryptionKey);

		//   ********************************************************************
		//   ******   Encryption Key must be 256 bits long (32 bytes)      ******
		//   ******   If it is longer than 32 bytes it will be truncated.  ******
		//   ******   If it is shorter than 32 bytes it will be padded     ******
		//   ******   with upper-case Xs.                                  ****** 
		//   ********************************************************************

		if (intLength >= 32) {
			vstrEncryptionKey = Strings.Left(vstrEncryptionKey, 32);
		}
		else {
			intLength = Strings.Len(vstrEncryptionKey);
			intRemaining = 32 - intLength;
			vstrEncryptionKey = vstrEncryptionKey + Strings.StrDup(intRemaining, "X");
		}

		bytKey = Encoding.ASCII.GetBytes(vstrEncryptionKey.ToCharArray());

		objRijndaelManaged = new RijndaelManaged();

		//   ***********************************************************************
		//   ******  Create the encryptor and write value to it after it is   ******
		//   ******  converted into a byte array                              ******
		//   ***********************************************************************

		try {
			objCryptoStream = new CryptoStream(objMemoryStream, objRijndaelManaged.CreateEncryptor(bytKey, bytIV), CryptoStreamMode.Write);
			objCryptoStream.Write(bytValue, 0, bytValue.Length);

			objCryptoStream.FlushFinalBlock();

			bytEncoded = objMemoryStream.ToArray();
			objMemoryStream.Close();
			objCryptoStream.Close();
		}

		catch {
		}

		//   ***********************************************************************
		//   ******   Return encryptes value (converted from  byte Array to   ******
		//   ******   a base64 string).  Base64 is MIME encoding)             ******
		//   ***********************************************************************

		return Convert.ToBase64String(bytEncoded);

	}

	public static string DecryptString128Bit(string vstrStringToBeDecrypted, string vstrDecryptionKey)
	{
		byte[] bytDataToBeDecrypted = null;
		byte[] bytTemp = null;
		byte[] bytIV = { 121, 241, 10, 1, 132, 74, 11, 39, 255, 91, 
		45, 78, 14, 211, 22, 62 };
		RijndaelManaged objRijndaelManaged = new RijndaelManaged();
		MemoryStream objMemoryStream = null;
		CryptoStream objCryptoStream = null;
		byte[] bytDecryptionKey = null;

		int intLength = 0;
		int intRemaining = 0;
		string strReturnString = string.Empty;

		//   *****************************************************************
		//   ******   Convert base64 encrypted value to byte array      ******
		//   *****************************************************************

		bytDataToBeDecrypted = Convert.FromBase64String(vstrStringToBeDecrypted);

		//   ********************************************************************
		//   ******   Encryption Key must be 256 bits long (32 bytes)      ******
		//   ******   If it is longer than 32 bytes it will be truncated.  ******
		//   ******   If it is shorter than 32 bytes it will be padded     ******
		//   ******   with upper-case Xs.                                  ****** 
		//   ********************************************************************

		intLength = Strings.Len(vstrDecryptionKey);

		if (intLength >= 32) {
			vstrDecryptionKey = Strings.Left(vstrDecryptionKey, 32);
		}
		else {
			intLength = Strings.Len(vstrDecryptionKey);
			intRemaining = 32 - intLength;
			vstrDecryptionKey = vstrDecryptionKey + Strings.StrDup(intRemaining, "X");
		}

		bytDecryptionKey = Encoding.ASCII.GetBytes(vstrDecryptionKey.ToCharArray());

		bytTemp = new byte[bytDataToBeDecrypted.Length + 1];

		objMemoryStream = new MemoryStream(bytDataToBeDecrypted);

		//   ***********************************************************************
		//   ******  Create the decryptor and write value to it after it is   ******
		//   ******  converted into a byte array                              ******
		//   ***********************************************************************


		try {
			objCryptoStream = new CryptoStream(objMemoryStream, objRijndaelManaged.CreateDecryptor(bytDecryptionKey, bytIV), CryptoStreamMode.Read);

			objCryptoStream.Read(bytTemp, 0, bytTemp.Length);

			objCryptoStream.FlushFinalBlock();
			objMemoryStream.Close();
			objCryptoStream.Close();

		}

		catch {
		}

		//   *****************************************
		//   ******   Return decypted value     ******
		//   *****************************************

		return StripNullCharacters(Encoding.ASCII.GetString(bytTemp));

	}


	public static string StripNullCharacters(string vstrStringWithNulls)
	{
		int intPosition = 0;
		string strStringWithOutNulls = "";

		if ((vstrStringWithNulls != null)) {
			intPosition = 1;
			strStringWithOutNulls = vstrStringWithNulls;

			while (intPosition > 0) {
				intPosition = Strings.InStr(intPosition, vstrStringWithNulls, Constants.vbNullChar);

				if (intPosition > 0) {
					strStringWithOutNulls = Strings.Left(strStringWithOutNulls, intPosition - 1) + Strings.Right(strStringWithOutNulls, strStringWithOutNulls.Length - intPosition);
				}

				if (intPosition > strStringWithOutNulls.Length) {
					break; // TODO: might not be correct. Was : Exit Do
				}
			}
		}
		return strStringWithOutNulls;

	}
	public static bool TrialRegistryPerUserExists(ref IActiveLock.ALTrialHideTypes TrialHideTypes)
	{
		bool functionReturnValue = false;
		functionReturnValue = false;
		if (TrialHideTypes == IActiveLock.ALTrialHideTypes.trialRegistryPerUser) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialHiddenFolder)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialSteganography)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialIsolatedStorage)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialIsolatedStorage)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialSteganography)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialIsolatedStorage)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialIsolatedStorage | IActiveLock.ALTrialHideTypes.trialHiddenFolder)) {
			functionReturnValue = true;
		}
		return functionReturnValue;
	}
	public static bool TrialHiddenFolderExists(ref IActiveLock.ALTrialHideTypes TrialHideTypes)
	{
		bool functionReturnValue = false;
		functionReturnValue = false;
		if (TrialHideTypes == IActiveLock.ALTrialHideTypes.trialHiddenFolder) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialRegistryPerUser)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialSteganography)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialIsolatedStorage)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialIsolatedStorage)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialSteganography)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialIsolatedStorage)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialIsolatedStorage | IActiveLock.ALTrialHideTypes.trialRegistryPerUser)) {
			functionReturnValue = true;
		}
		return functionReturnValue;
	}
	public static bool TrialSteganographyExists(ref IActiveLock.ALTrialHideTypes TrialHideTypes)
	{
		bool functionReturnValue = false;
		functionReturnValue = false;
		if (TrialHideTypes == IActiveLock.ALTrialHideTypes.trialSteganography) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialRegistryPerUser)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialHiddenFolder)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialIsolatedStorage)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialIsolatedStorage)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialHiddenFolder)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialIsolatedStorage)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialIsolatedStorage | IActiveLock.ALTrialHideTypes.trialRegistryPerUser)) {
			functionReturnValue = true;
		}
		return functionReturnValue;
	}
	public static bool TrialIsolatedStorageExists(ref IActiveLock.ALTrialHideTypes TrialHideTypes)
	{
		bool functionReturnValue = false;
		functionReturnValue = false;
		if (TrialHideTypes == IActiveLock.ALTrialHideTypes.trialIsolatedStorage) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialIsolatedStorage | IActiveLock.ALTrialHideTypes.trialRegistryPerUser)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialIsolatedStorage | IActiveLock.ALTrialHideTypes.trialHiddenFolder)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialIsolatedStorage | IActiveLock.ALTrialHideTypes.trialSteganography)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialIsolatedStorage | IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialSteganography)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialIsolatedStorage | IActiveLock.ALTrialHideTypes.trialRegistryPerUser | IActiveLock.ALTrialHideTypes.trialHiddenFolder)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialIsolatedStorage | IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialSteganography)) {
			functionReturnValue = true;
		}
		else if (TrialHideTypes == (IActiveLock.ALTrialHideTypes.trialIsolatedStorage | IActiveLock.ALTrialHideTypes.trialHiddenFolder | IActiveLock.ALTrialHideTypes.trialSteganography | IActiveLock.ALTrialHideTypes.trialRegistryPerUser)) {
			functionReturnValue = true;
		}
		return functionReturnValue;
	}
	public static string myDir()
	{
		return "mPY+Que6efQvkZsstJlvvw==";
	}
}
