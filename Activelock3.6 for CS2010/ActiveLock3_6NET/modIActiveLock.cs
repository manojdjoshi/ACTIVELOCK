using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

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

 // ERROR: Not supported in C#: OptionDeclaration
using System.IO;
using System.Security.Cryptography;
using System.Text;


/// <summary>
/// <para>This module contains common utility routines that can be shared between
/// ActiveLock and the client application.</para>
/// </summary>
/// <remarks>See the TODO List, within</remarks>
static class modActiveLock
{

	// Started: 04.21.2005
	// Modified: 03.25.2006
	//===============================================================================
	// @author activelock-admins
	// @version 3.3.0
	// @date 03.25.2006
	//
	//* ///////////////////////////////////////////////////////////////////////
	//  /                        MODULE TO DO LIST                            /
	//  ///////////////////////////////////////////////////////////////////////
	//
	// @bug rsa_createkey() sometimes causes crash.  This is due to a bug in
	//      ALCrypto3.dll in which a bad keyset is sometimes generated
	//      (either caused by <code>rsa_generate()</code> or one of <code>rsa_private_key_blob()</code>
	//      and <code>rsa_public_key_blob()</code>--we're not sure which is the culprit yet.
	//      This causes the <code>rsa_createkey()</code> call encryption routines to crash.
	//      The work-around for the time being is to keep regenerating the keyset
	//      until eventually you'll get a valid keyset that no longer causes a crash.
	//      You only need to go through this keyset generation step once.
	//      Once you have a valid keyset, you should store it inside your app for later use.
	//

	public const string STRKEYSTOREINVALID = "A license property contains an invalid value.";
	public const string STRLICENSEEXPIRED = "License expired.";
	public const string STRLICENSEINVALID = "License invalid.";
	public const string STRNOLICENSE = "No valid license.";
	public const string STRLICENSETAMPERED = "License may have been tampered.";
	public const string STRNOTINITIALIZED = "ActiveLock has not been initialized.";
	public const string STRNOTIMPLEMENTED = "Not implemented.";
	public const string STRCLOCKCHANGED = STRLICENSEINVALID + " System clock has been tampered.";
	public const string STRINVALIDTRIALDAYS = "Zero Free Trial days allowed.";
	public const string STRINVALIDTRIALRUNS = "Zero Free Trial runs allowed.";
	public const string STRFILETAMPERED = "Alcrypto3.dll has been tampered.";
	public const string STRKEYSTOREUNINITIALIZED = "Key Store Provider hasn't been initialized yet.";
	public const string STRKEYSTOREPATHISEMPTY = "Key Store Path (LIC file path) not specified.";
	public const string STRNOSOFTWARECODE = "Software code has not been set.";
	public const string STRNOSOFTWARENAME = "Software Name has not been set.";
	public const string STRNOSOFTWAREVERSION = "Software Version has not been set.";
	public const string STRNOSOFTWAREPASSWORD = "Software Password has not been set.";
	public const string STRUSERNAMETOOLONG = "User Name > 2000 characters.";
	public const string STRUSERNAMEINVALID = "User Name invalid.";
	public const string STRRSAERROR = "Internal RSA Error.";
	public const int RETVAL_ON_ERROR = -999;
	public const string STRWRONGIPADDRESS = "Wrong IP Address.";
	public const string STRUNDEFINEDSPECIALFOLDER = "Undefined Special Folder.";
	public const string STRDATEERROR = "Date Error.";
	public const string STRINTERNETNOTCONNECTED = "Internet Connection is Required. Please Connect and Try Again.";

	public const string STRSOFTWAREPASSWORDINVALID = "Password length>255 or invalid characters.";
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int rsa_encrypt(int CryptType, string data, ref int dLen, ref RSAKey ptrKey);
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int rsa_decrypt(int CryptType, string data, ref int dLen, ref RSAKey ptrKey);
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int md5_hash(string inData, int nDataLen, string outData);
	/// <summary>
	/// RSA encrypts the data.
	/// </summary>
	/// <param name="CryptType">0 for public&#59; 1 for private</param>
	/// <param name="data">Data to be encrypted</param>
	/// <param name="dLen">[in/out] Length of data, in bytes. This parameter will contain length of encrypted data when returned.</param>
	/// <param name="ptrKey">Key to be used for encryption</param>
	/// <returns>Integer - ?Not Documented</returns>
	/// <remarks></remarks>

	/// <summary>
	/// RSA decrypts the data.
	/// </summary>
	/// <param name="CryptType">0 for public&#59; 1 for private</param>
	/// <param name="data">Data to be decrypted</param>
	/// <param name="dLen">[in/out] Length of data, in bytes. This parameter will contain length of decrypted data when returned.</param>
	/// <param name="ptrKey">Key to be used for encryption</param>
	/// <returns>Integer - ?Not Documented!</returns>
	/// <remarks></remarks>

	/// <summary>
	/// Computes an MD5 hash from the data.
	/// </summary>
	/// <param name="inData">Data to be hashed</param>
	/// <param name="nDataLen">Length of inData</param>
	/// <param name="outData">[out] 32-byte Computed hash code</param>
	/// <returns>Integer - ?Undocumented!</returns>
	/// <remarks></remarks>

	/// <summary>
	/// ActiveLock Encryption Key
	/// </summary>
	/// <remarks>!!!WARNING!!! It is highly recommended that you change this key for your version of ActiveLock before deploying your app.</remarks>
		// RSA Private Key
	public const string ENCRYPT_KEY = "AAAAgEPRFzhQEF7S91vt2K6kOcEdDDe5BfwNiEL30/+ozTFHc7cZctB8NIlS++ZR//D3AjSMqScjh7xUF/gwvUgGCjiExjj1DF/XWFWnPOCfF8UxYAizCLZ9fdqxb1FRpI5NoW0xxUmvxGjmxKwazIW4P4XVi/+i1Bvh2qQ6ri3whcsNAAAAQQCyWGsbJKO28H2QLYH+enb7ehzwBThqfAeke/Gv1Te95yIAWme71I9aCTTlLsmtIYSk9rNrp3sh9ItD2Re67SE7AAAAQQCAookH1nws1gS2XP9cZTPaZEmFLwuxlSVsLQ5RWmd9cuxpgw5y2gIskbL4c+4oBuj0IDwKtnMrZq7UfV9I5VfVAAAAQQCEnyAuO0ahXH3KhAboop9+tCmRzZInTrDYdMy23xf3PLCLd777dL/Y2Y+zmaH1VO03m6iOog7WLiN4dCL7m+Im";

	/// <summary>
	/// ?Not Documented!
	/// </summary>
	/// <remarks></remarks>
	public const uint MAGICNUMBER_YES = 0xefcdab89;
	/// <summary>
	/// ?Not Documented!
	/// </summary>
	/// <remarks></remarks>

	public const uint MAGICNUMBER_NO = 0x98badcfe;
	/// <summary>
	/// ?Not Documented!
	/// </summary>
	/// <remarks></remarks>
	private const string SERVICE_PROVIDER = "Microsoft Base Cryptographic Provider v1.0";
	/// <summary>
	/// ?Not Documented!
	/// </summary>
	/// <remarks></remarks>
	private const string KEY_CONTAINER = "ActiveLock";
	/// <summary>
	/// ?Not Documented!
	/// </summary>
	/// <remarks></remarks>

	private const int PROV_RSA_FULL = 1;
	/// <summary>
	/// flag to indicate that module initialization has been done
	/// </summary>
	/// <remarks></remarks>
	private static bool fInit;
	[DllImport("kernel32", EntryPoint = "RtlMoveMemory", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern void CopyMem(ref int Destination, ref int source, int length);
	[DllImport("kernel32", EntryPoint = "GetModuleFileNameA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetModuleFileName(int hModule, string lpFileName, int nSize);
	/// <summary>
	/// The RtlMoveMemory routine moves memory either forward or backward, aligned or unaligned, in 4-byte blocks, followed by any remaining bytes.
	/// </summary>
	/// <param name="Destination">Pointer to the destination of the move.</param>
	/// <param name="source">Pointer to the memory to be copied.</param>
	/// <param name="length">Specifies the number of bytes to be copied.</param>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/ms803004.aspx for more Information</remarks>
	/// <summary>
	/// Retrieves the fully-qualified path for the file that contains the specified module. The module must have been loaded by the current process.
	/// </summary>
	/// <param name="hModule">[in, optional] A handle to the loaded module whose path is being requested. If this parameter is NULL, GetModuleFileName retrieves the path of the executable file of the current process.</param>
	/// <param name="lpFileName">[out] A pointer to a buffer that receives the fully-qualified path of the module. If the length of the path is less than the size that the nSize parameter specifies, the function succeeds and the path is returned as a null-terminated string.</param>
	/// <param name="nSize">[in] The size of the lpFilename buffer, in TCHARs.</param>
	/// <returns>If the function succeeds, the return value is the length of the string that is copied
	/// to the buffer, in characters, not including the terminating null character. If the buffer is too
	/// small to hold the module name, the string is truncated to nSize characters including the
	/// terminating null character, the function returns nSize, and the function sets the last error
	/// to ERROR_INSUFFICIENT_BUFFER.</returns>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/ms683197(VS.85).aspx for full Documentation!</remarks>

	// TODO: Check to ensure that the following enum is correct and modify the MapFileAndCheckSum Function, japreja!
	/// <summary>
	/// Not Inplimented - CheckSum return Values for MapFileAndCheckSum
	/// </summary>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/ms680355(VS.85).aspx</remarks>
	public enum CheckSumReturnValues
	{
		// TODO: Impliment Me! 
		/// <summary>
		/// Success!
		/// </summary>
		/// <remarks></remarks>
		CHECKSUM_SUCCESS = 0,
		/// <summary>
		/// Could not open the file.
		/// </summary>
		/// <remarks></remarks>
		CHECKSUM_OPEN_FAILURE = 1,
		/// <summary>
		/// Could not map the file.
		/// </summary>
		/// <remarks></remarks>
		CHECKSUM_MAP_FAILURE = 2,
		/// <summary>
		/// Could not map a view of the file.
		/// </summary>
		/// <remarks></remarks>
		CHECKSUM_MAPVIEW_FAILURE = 3,
		/// <summary>
		/// Could not convert the file name to Unicode.
		/// </summary>
		/// <remarks></remarks>
		CHECKSUM_UNICODE_FAILURE = 4
	}
	[DllImport("imagehlp", EntryPoint = "MapFileAndCheckSumA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int MapFileAndCheckSum(string FileName, ref int HeaderSum, ref int CheckSum);
	[DllImport("SHELL32.DLL", EntryPoint = "SHGetSpecialFolderPathA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern bool SHGetSpecialFolderPath(IntPtr hWnd, string lpszPath, int nFolder, bool fCreate);

	/// <summary>
	/// Computes the checksum of the specified file
	/// </summary>
	/// <param name="FileName">[in] The file name of the file for which the checksum is to be computed.</param>
	/// <param name="HeaderSum">[out] A pointer to a variable that receives the original checksum from the image file, or zero if there is an error.</param>
	/// <param name="CheckSum">[out] A pointer to a variable that receives the computed checksum.</param>
	/// <returns></returns>
	/// <remarks></remarks>

	/// <summary>
	/// <para>Retrieves the path of a special folder, identified by its <a href="http://msdn.microsoft.com/en-us/library/bb762494(VS.85).aspx">CSIDL</a>.</para>
	/// <para>Note: In Windows Vista, these values have been replaced by <a href="http://msdn.microsoft.com/en-us/library/bb762584(VS.85).aspx">KNOWNFOLDERID</a> values.
	/// See that topic for a list of the new constants and their corresponding CSIDL values.  For
	/// convenience, corresponding KNOWNFOLDERID values are also noted here for each CSIDL value.  The
	/// CSIDL system is supported under Windows Vista for compatibility reasons. However, new development
	/// should use KNOWNFOLDERID values rather than CSIDL values.</para>
	/// </summary>
	/// <param name="hWnd">Reserved.</param>
	/// <param name="lpszPath">[out] A pointer to a null-terminated string that receives the drive and path of the specified folder. This buffer must be at least MAX_PATH characters in size.</param>
	/// <param name="nFolder">See notes at http://msdn.microsoft.com/en-us/library/bb762494(VS.85).aspx</param>
	/// <param name="fCreate">[in] Indicates whether the folder should be created if it does not already exist. If this value is nonzero, the folder will be created. If this value is zero, the folder will not be created.</param>
	/// <returns></returns>
	/// <remarks></remarks>

	/// <summary>
	/// Specifies a date and time, using individual members for the month, day, year, weekday, hour, minute, second, and millisecond. The time is either in coordinated universal time (UTC) or local time, depending on the function that is being called.
	/// </summary>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/ms724950(VS.85).aspx for full documentation!</remarks>
	public struct SYSTEMTIME
	{
		/// <summary>
		/// The year. The valid values for this member are 1601 through 30827.
		/// </summary>
		/// <remarks></remarks>
		public short wYear;
		/// <summary>
		/// Month. - The valid values for this member are 1 through 12 corresponding to Jan. through Dec.
		/// </summary>
		/// <remarks></remarks>
		public short wMonth;
		/// <summary>
		/// Day of the week. The valid values for this member are 0 through 6, with Sunday being 0!
		/// </summary>
		/// <remarks></remarks>
		public short wDayOfWeek;
		/// <summary>
		/// The day of the month. The valid values for this member are 1 through 31.
		/// </summary>
		/// <remarks></remarks>
		public short wDay;
		/// <summary>
		/// The hour. The valid values for this member are 0 through 23.
		/// </summary>
		/// <remarks></remarks>
		public short wHour;
		/// <summary>
		/// The minute. The valid values for this member are 0 through 59.
		/// </summary>
		/// <remarks></remarks>
		public short wMinute;
		/// <summary>
		/// The second. The valid values for this member are 0 through 59.
		/// </summary>
		/// <remarks></remarks>
		public short wSecond;
		/// <summary>
		/// The millisecond. The valid values for this member are 0 through 999.
		/// </summary>
		/// <remarks></remarks>
		public short wMilliseconds;
	}

	/// <summary>
	/// Specifies settings for a time zone.
	/// </summary>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/ms725481.aspx for full documentation!</remarks>
	private struct TIME_ZONE_INFORMATION
	{
		/// <summary>
		/// <para>The current bias for local time translation on this computer, in minutes. The bias
		/// is the difference, in minutes, between Coordinated Universal Time (UTC) and local time.
		/// All translations between UTC and local time are based on the following formula:</para>
		/// <para>UTC = local time + bias</para>
		/// <para>This member is required.</para>
		/// </summary>
		/// <remarks></remarks>
			// current offset to GMT
		public int bias;
		/// <summary>
		/// <para>A description for standard time. For example, "EST" could indicate Eastern Standard
		/// Time. The string will be returned unchanged by the <a href="http://msdn.microsoft.com/en-us/library/ms724421(VS.85).aspx">GetTimeZoneInformation</a> function. This
		/// string can be empty.</para>
		/// </summary>
		/// <remarks></remarks>
		[VBFixedArray(64)]
			// unicode string
		public byte[] StandardName;
		/// <summary>
		/// <para>A SYSTEMTIME structure that contains a date and local time when the transition from
		/// daylight saving time to standard time occurs on this operating system. If the time zone does
		/// not support daylight saving time or if the caller needs to disable daylight saving time, the
		/// wMonth member in the SYSTEMTIME structure must be zero. If this date is specified, the
		/// DaylightDate member of this structure must also be specified. Otherwise, the system assumes
		/// the time zone data is invalid and no changes will be applied.</para>
		/// </summary>
		/// <remarks>See http://msdn.microsoft.com/en-us/library/ms725481.aspx for full Documentation!</remarks>
		public SYSTEMTIME StandardDate;
		/// <summary>
		/// <para>The bias value to be used during local time translations that occur during standard
		/// time. This member is ignored if a value for the StandardDate member is not supplied.</para>
		/// <para>This value is added to the value of the Bias member to form the bias used during
		/// standard time. In most time zones, the value of this member is zero.</para>
		/// </summary>
		/// <remarks></remarks>
		public int StandardBias;
		/// <summary>
		/// <para>A description for daylight saving time. For example, "PDT" could indicate Pacific
		/// Daylight Time. The string will be returned unchanged by the GetTimeZoneInformation function.
		/// This string can be empty.</para>
		/// </summary>
		/// <remarks></remarks>
		[VBFixedArray(64)]
		public byte[] DaylightName;
		/// <summary>
		/// <para>A SYSTEMTIME structure that contains a date and local time when the transition from
		/// standard time to daylight saving time occurs on this operating system. If the time zone does
		/// not support daylight saving time or if the caller needs to disable daylight saving time, the
		/// wMonth member in the SYSTEMTIME structure must be zero. If this date is specified, the
		/// StandardDate member in this structure must also be specified. Otherwise, the system assumes
		/// the time zone data is invalid and no changes will be applied.</para>
		/// <para>To select the correct day in the month, set the wYear member to zero, the wHour and
		/// wMinute members to the transition time, the wDayOfWeek member to the appropriate weekday, and
		/// the wDay member to indicate the occurrence of the day of the week within the month (1 to 5,
		/// where 5 indicates the final occurrence during the month if that day of the week does not occur
		/// 5 times).</para>
		/// <para>If the wYear member is not zero, the transition date is absolute; it will only occur one
		/// time. Otherwise, it is a relative date that occurs yearly.</para>
		/// </summary>
		/// <remarks></remarks>
		public SYSTEMTIME DaylightDate;
		/// <summary>
		/// <para>The bias value to be used during local time translations that occur during daylight
		/// saving time. This member is ignored if a value for the DaylightDate member is not supplied.</para>
		/// <para>This value is added to the value of the Bias member to form the bias used during
		/// daylight saving time. In most time zones, the value of this member is –60.</para>
		/// </summary>
		/// <remarks></remarks>

		public int DaylightBias;
		// For more information about the Dynamic DST key, see <a href="http://msdn.microsoft.com/en-us/library/ms724253(VS.85).aspx">DYNAMIC_TIME_ZONE_INFORMATION</a>.

		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

			 // ERROR: Not supported in C#: ReDimStatement

		}
	}

	/// <summary>
	/// ?Not Documented!
	/// </summary>
	/// <remarks></remarks>
	public enum TimeZoneReturn
	{
		TimeZoneCode = 0,
		TimeZoneName = 1,
		UTC_BaseOffset = 2,
		UTC_Offset = 3,
		DST_Active = 4,
		DST_Offset = 5
	}

	#region "For Time Zone Retrieval"

	/// <summary>
	/// The system cannot determine the current time zone. If daylight saving time is not used in the
	/// current time zone, this value is returned because there are no transition dates.
	/// </summary>
	/// <remarks></remarks>
	private const short TIME_ZONE_ID_UNKNOWN = 0;
	/// <summary>
	/// The system is operating in the range covered by the StandardDate member of the
	/// TIME_ZONE_INFORMATION structure.
	/// </summary>
	/// <remarks></remarks>
	private const short TIME_ZONE_ID_STANDARD = 1;
	/// <summary>
	/// If the function fails, the return value is TIME_ZONE_ID_UNKNOWN. To get extended error
	/// information, call GetLastError.
	/// </summary>
	/// <remarks></remarks>
	private const uint TIME_ZONE_ID_INVALID = 0xffffffff;
	/// <summary>
	/// The system is operating in the range covered by the DaylightDate member of the
	/// TIME_ZONE_INFORMATION structure.
	/// </summary>
	/// <remarks></remarks>

	private const short TIME_ZONE_ID_DAYLIGHT = 2;
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern void GetSystemTime(ref SYSTEMTIME lpSystemTime);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetTimeZoneInformation(ref TIME_ZONE_INFORMATION lpTimeZoneInformation);
	/// <summary>
	/// Retrieves the current system date and time. The system time is expressed in Coordinated
	/// Universal Time (UTC).
	/// </summary>
	/// <param name="lpSystemTime">A pointer to a SYSTEMTIME structure to receive the current system date and time. The lpSystemTime parameter must not be NULL. Using NULL will result in an access violation.</param>
	/// <remarks></remarks>
	/// <summary>
	/// Retrieves the current time zone settings. These settings control the translations between
	/// Coordinated Universal Time (UTC) and local time.
	/// </summary>
	/// <param name="lpTimeZoneInformation">A pointer to a <see cref="TIME_ZONE_INFORMATION"/> structure to receive the current settings.</param>
	/// <returns></returns>
	/// <remarks></remarks>

	#endregion

	#region "To Report API errors"

	/// <summary>
	/// The function allocates a buffer large enough to hold the formatted message, and places a pointer
	/// to the allocated buffer at the address specified by lpBuffer. The lpBuffer parameter is a pointer
	/// to an LPTSTR; you must cast the pointer to an LPTSTR (for example, (LPTSTR)&amp;lpBuffer). The nSize
	/// parameter specifies the minimum number of TCHARs to allocate for an output message buffer. The
	/// caller should use the LocalFree function to free the buffer when it is no longer needed.
	/// </summary>
	/// <remarks></remarks>
	private const short FORMAT_MESSAGE_ALLOCATE_BUFFER = 0x100;
	/// <summary>
	/// <para>The Arguments parameter is not a va_list structure, but is a pointer to an array of values
	/// that represent the arguments.</para>
	/// <para>This flag cannot be used with 64-bit integer values. If you are using a 64-bit integer,
	/// you must use the va_list structure.</para>
	/// </summary>
	/// <remarks></remarks>
	private const short FORMAT_MESSAGE_ARGUMENT_ARRAY = 0x2000;
	/// <summary>
	/// <para>The lpSource parameter is a module handle containing the message-table resource(s) to search. If
	/// this lpSource handle is NULL, the current process's application image file will be searched. This
	/// flag cannot be used with FORMAT_MESSAGE_FROM_STRING.</para>
	/// <para>If the module has no message table resource, the function fails with
	/// ERROR_RESOURCE_TYPE_NOT_FOUND.</para>
	/// </summary>
	/// <remarks></remarks>
	private const short FORMAT_MESSAGE_FROM_HMODULE = 0x800;
	/// <summary>
	/// The lpSource parameter is a pointer to a null-terminated string that contains a message 
	/// definition. The message definition may contain insert sequences, just as the message text in a
	/// message table resource may. This flag cannot be used with FORMAT_MESSAGE_FROM_HMODULE or
	/// FORMAT_MESSAGE_FROM_SYSTEM.
	/// </summary>
	/// <remarks></remarks>
	private const short FORMAT_MESSAGE_FROM_STRING = 0x400;
	/// <summary>
	/// <para>The function should search the system message-table resource(s) for the requested message.
	/// If this flag is specified with FORMAT_MESSAGE_FROM_HMODULE, the function searches the system
	/// message table if the message is not found in the module specified by lpSource. This flag cannot
	/// be used with FORMAT_MESSAGE_FROM_STRING.</para>
	/// <para>If this flag is specified, an application can pass the result of the GetLastError function
	/// to retrieve the message text for a system-defined error.</para>
	/// </summary>
	/// <remarks></remarks>
	private const short FORMAT_MESSAGE_FROM_SYSTEM = 0x1000;
	/// <summary>
	/// Insert sequences in the message definition are to be ignored and passed through to the output
	/// buffer unchanged. This flag is useful for fetching a message for later formatting. If this flag
	/// is set, the Arguments parameter is ignored.
	/// </summary>
	/// <remarks></remarks>
	private const short FORMAT_MESSAGE_IGNORE_INSERTS = 0x200;
	/// <summary>
	/// The function ignores regular line breaks in the message definition text. The function stores
	/// hard-coded line breaks in the message definition text into the output buffer. The function
	/// generates no new line breaks.
	/// </summary>
	/// <remarks></remarks>

	private const short FORMAT_MESSAGE_MAX_WIDTH_MASK = 0xff;
	[DllImport("kernel32", EntryPoint = "FormatMessageA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int FormatMessage(int dwFlags, ref long lpSource, int dwMessageId, int dwLanguageId, string lpBuffer, int nSize, ref int Arguments);
	[DllImport("kernel32", EntryPoint = "GetWindowsDirectoryA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int GeneralWinDirApi(string lpBuffer, int nSize);
	[DllImport("kernel32.dll", EntryPoint = "GetSystemDirectoryA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int GetSystemDirectory(string lpBuffer, int nSize);
	[DllImport("kernel32", EntryPoint = "GetLocaleInfoA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetLocaleInfo(int Locale, int LCType, string lpLCData, int cchData);
	[DllImport("kernel32", EntryPoint = "SetLocaleInfoA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern bool SetLocaleInfo(int Locale, int LCType, string lpLCData);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern short GetUserDefaultLCID();
	/// <summary>
	/// Formats a message string. The function requires a message definition as input. The message
	/// definition can come from a buffer passed into the function. It can come from a message table
	/// resource in an already-loaded module. Or the caller can ask the function to search the system's
	/// message table resource(s) for the message definition. The function finds the message definition
	/// in a message table resource based on a message identifier and a language identifier. The function
	/// copies the formatted message text to an output buffer, processing any embedded insert sequences
	/// if requested.
	/// </summary>
	/// <param name="dwFlags">[in] The formatting options, and how to interpret the lpSource parameter.
	/// The low-order byte of dwFlags specifies how the function handles line breaks in the output buffer.
	/// The low-order byte can also specify the maximum width of a formatted output line.</param>
	/// <param name="lpSource">[in, optional] The location of the message definition. The type of this
	/// parameter depends upon the settings in the dwFlags parameter. </param>
	/// <param name="dwMessageId">[in] The message identifier for the requested message. This parameter
	/// is ignored if dwFlags includes FORMAT_MESSAGE_FROM_STRING.</param>
	/// <param name="dwLanguageId">[in] The language identifier for the requested message. This parameter
	/// is ignored if dwFlags includes FORMAT_MESSAGE_FROM_STRING.</param>
	/// <param name="lpBuffer">[out]<para>A pointer to a buffer that receives the null-terminated string
	/// that specifies the formatted message. If dwFlags includes FORMAT_MESSAGE_ALLOCATE_BUFFER, the
	/// function allocates a buffer using the LocalAlloc function, and places the pointer to the buffer
	/// at the address specified in lpBuffer.</para>
	/// <para>This buffer cannot be larger than 64K bytes.</para></param>
	/// <param name="nSize">[in] <para>If the FORMAT_MESSAGE_ALLOCATE_BUFFER flag is not set, this
	/// parameter specifies the size of the output buffer, in TCHARs. If FORMAT_MESSAGE_ALLOCATE_BUFFER
	/// is set, this parameter specifies the minimum number of TCHARs to allocate for an output buffer.</para>
	/// <para>The output buffer cannot be larger than 64K bytes.</para>
	/// </param>
	/// <param name="Arguments">[in, optional] An array of values that are used as insert values in the
	/// formatted message. A %1 in the format string indicates the first value in the Arguments array;
	/// a %2 indicates the second argument; and so on.</param>
	/// <returns>If the function succeeds, the return value is the number of TCHARs stored in the output
	/// buffer, excluding the terminating null character. If the function fails, the return value is zero.
	/// To get extended error information, call GetLastError.</returns>
	/// <remarks></remarks>
	/// <summary>
	/// Retrieves the path of the Windows directory. The Windows directory contains such files as
	/// applications, initialization files, and help files.
	/// </summary>
	/// <param name="lpBuffer">[out] A pointer to a buffer that receives the path. This path does not end
	/// with a backslash unless the Windows directory is the root directory. For example, if the Windows
	/// directory is named Windows on drive C, the path of the Windows directory retrieved by this
	/// function is C:\Windows. If the system was installed in the root directory of drive C, the path
	/// retrieved is C:\.</param>
	/// <param name="nSize">[in] The maximum size of the buffer specified by the lpBuffer parameter, in
	/// TCHARs. This value should be set to MAX_PATH.</param>
	/// <returns><para>If the function succeeds, the return value is the length of the string copied to
	/// the buffer, in TCHARs, not including the terminating null character.</para>
	/// <para>If the length is greater than the size of the buffer, the return value is the size of the
	/// buffer required to hold the path.</para>
	/// </returns>
	/// <remarks>The Windows directory is the directory where an application should store initialization
	/// and help files. If the user is running a shared version of the system, the Windows directory is
	/// guaranteed to be private for each user.</remarks>
	/// <summary>
	/// Retrieves the path of the system directory. The system directory contains system files such as
	/// dynamic-link libraries and drivers.
	/// </summary>
	/// <param name="lpBuffer">[out] A pointer to the buffer to receive the path. This path does not end
	/// with a backslash unless the system directory is the root directory. For example, if the system
	/// directory is named Windows\System on drive C, the path of the system directory retrieved by this
	/// function is C:\Windows\System.</param>
	/// <param name="nSize">[in] The maximum size of the buffer, in TCHARs.</param>
	/// <returns>If the function succeeds, the return value is the length, in TCHARs, of the string copied
	/// to the buffer, not including the terminating null character. If the length is greater than the
	/// size of the buffer, the return value is the size of the buffer required to hold the path,
	/// including the terminating null character.</returns>
	/// <remarks></remarks>

	#endregion

	// The following constants and declares are used to Get/Set Locale Date format
	const short LOCALE_SSHORTDATE = 0x1f;

	public static string regionalSymbol;
	// Internet connection constants
	/// <summary>
	/// Used by Function IsWebConnected
	/// </summary>
	/// <remarks></remarks>

	static string ConnectionQualityString = "Off";
	[DllImport("wininet.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern bool InternetGetConnectedState(ref Int32 lpSFlags, Int32 dwReserved);
	/// <summary>
	/// Retrieves the connected state of the local system.
	/// </summary>
	/// <param name="lpSFlags">[out] Pointer to a variable that receives the connection description. This
	/// parameter may return a valid flag even when the function returns FALSE.</param>
	/// <param name="dwReserved">[in] This parameter is reserved and must be 0.</param>
	/// <returns></returns>
	/// <remarks></remarks>

	/// <summary>
	/// Internet connection states!
	/// </summary>
	/// <remarks></remarks>
	public enum InetConnState
	{
		/// <summary>
		/// Local system uses a modem to connect to the Internet.
		/// </summary>
		/// <remarks></remarks>
		modem = 0x1,
		/// <summary>
		/// Local system uses a local area network to connect to the Internet.
		/// </summary>
		/// <remarks></remarks>
		lan = 0x2,
		/// <summary>
		/// Local system uses a proxy server to connect to the Internet.
		/// </summary>
		/// <remarks></remarks>
		proxy = 0x4,
		/// <summary>
		/// ?Undocumented!
		/// </summary>
		/// <remarks></remarks>
		ras = 0x10,
		/// <summary>
		/// Local system is in offline mode.
		/// </summary>
		/// <remarks></remarks>
		offline = 0x20,
		/// <summary>
		/// Local system has a valid connection to the Internet, but it might or might not be currently
		/// connected.
		/// </summary>
		/// <remarks></remarks>
		configured = 0x40
	}

	/// <summary>
	/// Trims Null characters from the string.
	/// </summary>
	/// <param name="startstr">String - String to be trimmed</param>
	/// <returns>String - Trimmed string</returns>
	/// <remarks></remarks>
	public static string TrimNulls(ref string startstr)
	{
		string functionReturnValue = null;
		short pos = 0;
		pos = Strings.InStr(startstr, Strings.Chr(0));
		if (pos) {
			functionReturnValue = Strings.Trim(Strings.Left(startstr, pos - 1));
		}
		else {
			functionReturnValue = Strings.Trim(startstr);
		}
		return functionReturnValue;
	}

	/// <summary>
	/// Reads a binary file into the sData buffer. Returns the number of bytes read.
	/// </summary>
	/// <param name="sPath">String - Path to the file to be read</param>
	/// <param name="sData">String - Output parameter contains the data that has been read</param>
	/// <returns>Long - Number of bytes read, 0 if no file was read</returns>
	/// <remarks></remarks>

	public static int ReadFile(string sPath, ref string sData)
	{
		int functionReturnValue = 0;
		CRC32 c = new CRC32();
		int crc = 0;

		// CRC32 Hash:
		FileStream f = new FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192);
		crc = c.GetCrc32(ref f);
		f.Close();

		// File size:
		//f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
		//txtSize.Text = String.Format("{0}", f.Length)
		//f.Close()
		//txtCrc32.Text = String.Format("{0:X8}", crc)
		//txtTime.Text = String.Format("{0}", h.ElapsedTime)

		// Run MD5 Hash
		f = new FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192);
		MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
		md5.ComputeHash(f);
		f.Close();

		byte[] hash = md5.Hash;
		StringBuilder buff = new StringBuilder();
		byte hashByte = 0;
		foreach (byte hashByte_loopVariable in hash) {
			hashByte = hashByte_loopVariable;
			buff.Append(string.Format("{0:X1}", hashByte));
		}
		sData = buff.ToString();
		//MD5 String

		// Run SHA-1 Hash
		//f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
		//Dim sha1 As SHA1CryptoServiceProvider = New SHA1CryptoServiceProvider
		//sha1.ComputeHash(f)
		//f.Close()
		//hash = SHA1.Hash
		//buff = New StringBuilder
		//For Each hashByte In hash
		//    buff.Append(String.Format("{0:X1}", hashByte))
		//Next
		//txtSHA1.Text = buff.ToString()

		functionReturnValue = Strings.Len(sData);
		return;
		Hell:
		Set_locale(regionalSymbol);
		Err().Raise(Err().Number, Err().Source, Err().Description, Err().HelpFile, Err().HelpContext);
		return functionReturnValue;
	}

	/// <summary>
	/// [INTERNAL] Call-back routine used by ALCrypto3.dll during key generation process.
	/// </summary>
	/// <param name="param">Long - TBD</param>
	/// <param name="action">Long - Action being performed</param>
	/// <param name="phase">Long - Current phase</param>
	/// <param name="iprogress">Long - Percent complete</param>
	/// <remarks></remarks>
	public static void CryptoProgressUpdate(int param, int action, int phase, int iprogress)
	{
		System.Diagnostics.Debug.WriteLine("Progress Update received " + param + ", action: " + action + ", iprogress: " + iprogress);
	}

	/// <summary>
	/// This is a dummy sub. Used to circumvent the End statement restriction in COM DLLs.
	/// </summary>
	/// <remarks></remarks>
	public static void EndSub()
	{
		//Dummy sub
	}

	/// <summary>
	/// Computes an MD5 hash of the specified file.
	/// </summary>
	/// <param name="strPath">String - File path</param>
	/// <returns>String - MD5 Hash Value</returns>
	/// <remarks></remarks>
	public static string MD5HashFile(string strPath)
	{
		System.Diagnostics.Debug.WriteLine("Hashing file " + strPath);
		System.Diagnostics.Debug.WriteLine("File Date: " + FileSystem.FileDateTime(strPath));
		// read and hash the content
		string sData = string.Empty;
		int nFileLen = 0;
		nFileLen = ReadFile(strPath, ref sData);
		// use the .NET's native MD5 functions instead of our own MD5 hashing routine
		// and instead of ALCrypto's md5_hash() function.
		return Strings.LCase(sData);
		//<--- ReadFile procedure already computes the MD5.Hash
	}

	/// <summary>
	/// Checks if a file exists in the system.
	/// </summary>
	/// <param name="strFile">String - File path and name</param>
	/// <returns>Boolean - True if file exists, False if it doesn't</returns>
	/// <remarks></remarks>
	public static bool FileExists(string strFile)
	{
		bool functionReturnValue = false;
		functionReturnValue = false;
		if (File.Exists(strFile) == true) {
			functionReturnValue = true;
		}
		return functionReturnValue;
	}

	// Name: Function LocalTimeZone
	// Input:
	//   ByVal returnType As TimeZoneReturn - Type of time zone information being requested
	//       UTC_BaseOffset = UTC offset, not including DST <br>
	//       UTC_Offset = UTC offset, including DST if active <br>
	//       DST_Active = True if DST is currently active, otherwise false <br>
	//       DST_Offset = Offset value for DST (generally -60, if in US)
	/// <summary>
	/// Retrieves the local time zone.
	/// </summary>
	/// <param name="returnType">TimeZoneReturn - Type of time zone information being requested</param>
	/// <returns>Variant - Return type varies depending on returnValue parameter.</returns>
	/// <remarks></remarks>
	public static object LocalTimeZone(TimeZoneReturn returnType)
	{
		object functionReturnValue = null;
		int x = 0;
		TIME_ZONE_INFORMATION tzi = null;
		string strName = null;
		bool bDST = false;
		int rc = 0;

		functionReturnValue = null;

		rc = GetTimeZoneInformation(ref tzi);
		switch (rc) {
			// if not daylight assume standard
			case TIME_ZONE_ID_DAYLIGHT:
				strName = System.Text.UnicodeEncoding.Unicode.GetString(tzi.DaylightName);
				// convert to string
				bDST = true;
				break;
			default:
				strName = System.Text.UnicodeEncoding.Unicode.GetString(tzi.StandardName);
				break;
		}

		// name terminates with null
		x = Strings.InStr(strName, Constants.vbNullChar);
		if (x > 0) strName = Strings.Left(strName, x - 1); 

		if (returnType == TimeZoneReturn.DST_Active) {
			functionReturnValue = bDST;
		}

		if (returnType == TimeZoneReturn.TimeZoneName) {
			functionReturnValue = strName;
		}

		if (returnType == TimeZoneReturn.TimeZoneCode) {
			functionReturnValue = Strings.Left(strName, 1);
			x = Strings.InStr(1, strName, " ");
			while (x > 0) {
				functionReturnValue = functionReturnValue + Strings.Mid(strName, x + 1, 1);
				x = Strings.InStr(x + 1, strName, " ");
			}
			functionReturnValue = Strings.Trim(LocalTimeZone());
		}

		if (returnType == TimeZoneReturn.UTC_BaseOffset) {
			functionReturnValue = tzi.bias;
		}

		if (returnType == TimeZoneReturn.DST_Offset) {
			functionReturnValue = tzi.DaylightBias;
		}

		if (returnType == TimeZoneReturn.UTC_Offset) {
			if (tzi.DaylightBias == -60) {
				functionReturnValue = tzi.bias;
			}
			else {
				functionReturnValue = -tzi.bias;
			}
			// Account for Daylight Savings Time
			if (bDST) functionReturnValue = functionReturnValue - 60; 
		}
		return functionReturnValue;
	}

	/// <summary>
	/// Performs RSA signing of <code>strData</code> using the specified key.
	/// </summary>
	/// <param name="strPub">RSA Public key blob</param>
	/// <param name="strPriv">RSA Private key blob</param>
	/// <param name="strdata">String - Data to be signed</param>
	/// <returns>String - Signature string</returns>
	/// <remarks>05.13.05    - alkan  - Removed the modActiveLock references</remarks>
	public static string RSASign(string strPub, string strPriv, string strdata)
	{
		RSAKey KEY = null;
		// create the key from the key blobs
		if (modALUGEN.rsa_createkey(strPub, Strings.Len(strPub), strPriv, Strings.Len(strPriv), ref KEY) == RETVAL_ON_ERROR) {
			Set_locale(regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, modTrial.ACTIVELOCKSTRING, STRRSAERROR);
		}

		// sign the data using the created key
		int sLen = 0;
		if (modALUGEN.rsa_sign(ref KEY, strdata, Strings.Len(strdata), Constants.vbNullString, ref sLen) == RETVAL_ON_ERROR) {
			Set_locale(regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, modTrial.ACTIVELOCKSTRING, STRRSAERROR);
		}
		string strSig = null;
		strSig = new string(Strings.Chr(0), sLen);
		if (modALUGEN.rsa_sign(ref KEY, strdata, Strings.Len(strdata), strSig, ref sLen) == RETVAL_ON_ERROR) {
			Set_locale(regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, modTrial.ACTIVELOCKSTRING, STRRSAERROR);
		}
		// throw away the key
		if (modALUGEN.rsa_freekey(ref KEY) == RETVAL_ON_ERROR) {
			Set_locale(regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, modTrial.ACTIVELOCKSTRING, STRRSAERROR);
		}
		return strSig;
	}

	//===============================================================================
	/// <summary>
	/// Verifies an RSA signature.
	/// </summary>
	/// <param name="strPub">String - Public key blob</param>
	/// <param name="strdata">String - Data to be signed</param>
	/// <param name="strSig">String - Private key blob</param>
	/// <returns>Long - Zero if verification is successful, non-zero otherwise.</returns>
	/// <remarks></remarks>
	public static int RSAVerify(string strPub, string strdata, string strSig)
	{
		RSAKey KEY = null;
		int rc = 0;
		// create the key from the public key blob
		if (modALUGEN.rsa_createkey(strPub, Strings.Len(strPub), Constants.vbNullString, 0, ref KEY) == RETVAL_ON_ERROR) {
			Set_locale(regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, modTrial.ACTIVELOCKSTRING, STRRSAERROR);
		}
		// validate the key
		rc = modALUGEN.rsa_verifysig(ref KEY, strSig, Strings.Len(strSig), strdata, Strings.Len(strdata));
		if (rc == RETVAL_ON_ERROR) {
			Set_locale(regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, modTrial.ACTIVELOCKSTRING, STRRSAERROR);
		}
		// de-allocate memory used by the key
		if (modALUGEN.rsa_freekey(ref KEY) == RETVAL_ON_ERROR) {
			Set_locale(regionalSymbol);
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, modTrial.ACTIVELOCKSTRING, STRRSAERROR);
		}
		return rc;
	}

	/// <summary>
	/// Retrieves the error text for the specified Windows error code
	/// </summary>
	/// <param name="lLastDLLError">Long - Last DLL error as an input</param>
	/// <returns>String - Error message string</returns>
	/// <remarks></remarks>
	public static string WinError(int lLastDLLError)
	{
		string functionReturnValue = null;
		string sBuff = null;
		int lCount = 0;

		functionReturnValue = string.Empty;

		// Return the error message associated with LastDLLError:
		sBuff = new string(Strings.Chr(0), 256);
		lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS, ref 0, lLastDLLError, 0, sBuff, Strings.Len(sBuff), ref 0);
		if (lCount) {
			functionReturnValue = Strings.Left(sBuff, lCount);
		}
		return functionReturnValue;
	}

	/// <summary>
	/// Gets the windows directory
	/// </summary>
	/// <returns>String - Windows directory path</returns>
	/// <remarks></remarks>
	public static string WinDir()
	{
		string WinSysPath = System.Environment.GetFolderPath(Environment.SpecialFolder.System);
		return WinSysPath.Substring(0, WinSysPath.LastIndexOf("\\"));
	}

	/// <summary>
	/// Gets the Windows system directory
	/// </summary>
	/// <returns>String - Windows system directory path</returns>
	/// <remarks></remarks>
	public static string WinSysDir()
	{
		return System.Environment.GetFolderPath(Environment.SpecialFolder.System);
		// or could use WinSysDir = System.Environment.SystemDirectory
	}

	/// <summary>
	/// Checks if a Folder Exists
	/// </summary>
	/// <param name="sFolder">String -  Name of the folder in question</param>
	/// <returns>Boolean - Returns true if the Folder Exists</returns>
	/// <remarks></remarks>
	public static bool FolderExists(string sFolder)
	{
		return Directory.Exists(sFolder);
	}

	/// <summary>
	/// Packs two 8-bit integers into a 16-bit integer.
	/// </summary>
	/// <param name="LoByte">Byte - A byte that is to become the Low byte.</param>
	/// <param name="HiByte">Byte - A byte that is to become the High byte.</param>
	/// <returns></returns>
	/// <remarks></remarks>
	public static short MakeWord(byte LoByte, byte HiByte)
	{
		short functionReturnValue = 0;
		if ((HiByte & 0x80) != 0) {
			functionReturnValue = ((HiByte * 256) + LoByte) | 0xffff0000;
		}
		else {
			functionReturnValue = (HiByte * 256) + LoByte;
		}
		return functionReturnValue;
	}

	/// <summary>
	/// ?Not Documented!
	/// </summary>
	/// <param name="w"></param>
	/// <returns></returns>
	/// <remarks></remarks>
	public static byte HiByte(short w)
	{
		return (w & 0xff00) / 256;
	}

	/// <summary>
	/// ?Not Documented!
	/// </summary>
	/// <param name="w"></param>
	/// <returns></returns>
	/// <remarks></remarks>
	public static byte LoByte(short w)
	{
		return w & 0xff;
	}


	/// <summary>
	/// Converts a local date-time into UTC/GMT date-time
	/// </summary>
	/// <param name="dt">Date - Date-Time input to be converted into UTC Date-Time</param>
	/// <returns>Date - UTC Date-Time</returns>
	/// <remarks></remarks>
	public static System.DateTime UTC(System.DateTime dt)
	{
		//  Returns current UTC date-time.
		return dt.AddMinutes(LocalTimeZone(TimeZoneReturn.UTC_Offset));
	}

	/// <summary>
	/// Retrieves the regional setting
	/// </summary>
	/// <remarks></remarks>
	public static void Get_locale()
	{
		string Symbol = null;
		int iRet1 = 0;
		int iRet2 = 0;
		string lpLCDataVar = string.Empty;
		short Pos = 0;
		int Locale = 0;
		Locale = GetUserDefaultLCID();
		iRet1 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, lpLCDataVar, 0);
		Symbol = new string(Strings.Chr(0), iRet1);
		iRet2 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol, iRet1);
		Pos = Strings.InStr(Symbol, Strings.Chr(0));
		if (Pos > 0) {
			Symbol = Strings.Left(Symbol, Pos - 1);
			if (Symbol != "yyyy/MM/dd") regionalSymbol = Symbol; 
		}
	}

	/// <summary>
	/// Changes the regional setting.
	/// </summary>
	/// <param name="localSymbol"></param>
	/// <remarks></remarks>
	//Change the regional setting
	public static void Set_locale([System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute("")]  // ERROR: Optional parameters aren't supported in C#
string localSymbol)
	{
		string Symbol = null;
		int iRet = 0;
		int Locale = 0;
		Locale = GetUserDefaultLCID();
		//Get user Locale ID
		if (string.IsNullOrEmpty(localSymbol)) {
			Symbol = "yyyy/MM/dd";
			//New character for the locale
		}
		else {
			Symbol = localSymbol;
		}

		iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol);
	}

	/// <summary>
	/// Gets special folders...
	/// </summary>
	/// <param name="CSIDL">See this functions Definition for detailed info...</param>
	/// <returns></returns>
	/// <remarks></remarks>
	public static string ActivelockGetSpecialFolder(long CSIDL)
	{
		string functionReturnValue = null;
		//value  ID                      Result
		//2      CSIDL_PROGRAMS          C:\Documents and Settings\<USERNAME>\Start Menu\Programs
		//5      CSIDL_PERSONAL          C:\Documents and Settings\<USERNAME>\My Documents
		//6      CSIDL_FAVORITES         C:\Documents and Settings\<USERNAME>\Favorites
		//7      CSIDL_STARTUP           C:\Documents and Settings\<USERNAME>\Start Menu\Programs\Startup
		//8      CSIDL_RECENT            C:\Documents and Settings\<USERNAME>\Recent
		//9      CSIDL_SENDTO            C:\Documents and Settings\<USERNAME>\SendTo
		//11     CSIDL_STARTMENU         C:\Documents and Settings\<USERNAME>\Start Menu
		//13     CSIDL_MYMUSIC           C:\Documents and Settings\<USERNAME>\My Documents\My Music
		//14     CSIDL_MYVIDEO           C:\Documents and Settings\<USERNAME>\My Documents\My Video
		//16     CSIDL_DESKTOPDIRECTORY  C:\Documents and Settings\<USERNAME>\Desktop
		//19     CSIDL_NETHOOD           C:\Documents and Settings\<USERNAME>\NetHood
		//20     Virtual - don't use     CSIDL_FONTS             C:\Windows\Fonts
		//21     CSIDL_TEMPLATES         C:\Documents and Settings\<USERNAME>\Templates
		//22     CSIDL_COMMON_STARTMENU  C:\Documents and Settings\All Users\Start Menu
		//23     CSIDL_COMMON_PROGRAMS   C:\Documents and Settings\All Users\Start Menu\Programs
		//24     CSIDL_COMMON_STARTUP    C:\Documents and Settings\All Users\Start Menu\Programs\Startup
		//25     CSIDL_COMMON_DESKTOPDIRECTORY     C:\Documents and Settings\All Users\Desktop
		//26     CSIDL_APPDATA           C:\Documents and Settings\<USERNAME>\Application Data
		//27     CSIDL_PRINTHOOD         C:\Documents and Settings\<USERNAME>\PrintHood
		//28     CSIDL_LOCAL_APPDATA     C:\Documents and Settings\<USERNAME>\Local Settings\Application Data
		//31     CSIDL_COMMON_FAVORITES  C:\Documents and Settings\All Users\Favorites
		//32     CSIDL_INTERNET_CACHE    C:\Documents and Settings\<USERNAME>\Local Settings\Temporary Internet Files
		//33     CSIDL_COOKIES           C:\Documents and Settings\<USERNAME>\Cookies
		//34     CSIDL_HISTORY           C:\Documents and Settings\<USERNAME>\Local Settings\History
		//35     CSIDL_COMMON_APPDATA    C:\Documents and Settings\All Users\Application Data
		//36     CSIDL_WINDOWS           C:\Windows
		//37     CSIDL_SYSTEM            C:\Windows\system32
		//38     CSIDL_PROGRAM_FILES     C:\Program Files
		//39     CSIDL_MYPICTURES        C:\Documents and Settings\<USERNAME>\My Documents\My Pictures
		//40     CSIDL_PROFILE           C:\Documents and Settings\<USERNAME>
		//41     CSIDL_SYSTEMX86         C:\Windows\system32  'x86 system directory on RISC
		//43     CSIDL_PROGRAM_FILES_COMMON     C:\Program Files\Common Files
		//45     CSIDL_COMMON_TEMPLATES  C:\Documents and Settings\All Users\Templates
		//46     CSIDL_COMMON_DOCUMENTS  C:\Documents and Settings\All Users\Documents
		//47     CSIDL_COMMON_ADMINTOOLS C:\Documents and Settings\All Users\Start Menu\Programs\Administrative Tools
		//48     CSIDL_ADMINTOOLS        C:\Documents and Settings\<USERNAME>\Start Menu\Programs\Administrative Tools
		//53     CSIDL_COMMON_MUSIC      C:\Documents and Settings\All Users\My Documents\My Music
		//54     CSIDL_COMMON_PICTURES   C:\Documents and Settings\All Users\My Documents\My Pictures
		//55     CSIDL_COMMON_VIDEO      C:\Documents and Settings\All Users\My Documents\My Video

		// CLSIDs for XP and older
		//CSIDL_DESKTOP = &H0
		//CSIDL_INTERNET = &H1
		//CSIDL_PROGRAMS = &H2
		//CSIDL_CONTROLS = &H3
		//CSIDL_PRINTERS = &H4
		//CSIDL_PERSONAL = &H5
		//CSIDL_FAVORITES = &H6
		//CSIDL_STARTUP = &H7
		//CSIDL_RECENT = &H8
		//CSIDL_SENDTO = &H9
		//CSIDL_BITBUCKET = &HA
		//CSIDL_STARTMENU = &HB
		//CSIDL_MYDOCUMENTS = &HC
		//CSIDL_MYMUSIC = &HD
		//CSIDL_MYVIDEO = &HE
		//CSIDL_DESKTOPDIRECTORY = &H10
		//CSIDL_DRIVES = &H11
		//CSIDL_NETWORK = &H12
		//CSIDL_NETHOOD = &H13
		//CSIDL_FONTS = &H14
		//CSIDL_TEMPLATES = &H15
		//CSIDL_COMMON_STARTMENU = &H16
		//CSIDL_COMMON_PROGRAMS = &H17
		//CSIDL_COMMON_STARTUP = &H18
		//CSIDL_COMMON_DESKTOPDIRECTORY = &H19
		//CSIDL_APPDATA = &H1A
		//CSIDL_PRINTHOOD = &H1B
		//CSIDL_LOCAL_APPDATA = &H1C
		//CSIDL_ALTSTARTUP = &H1D
		//CSIDL_COMMON_ALTSTARTUP = &H1E
		//CSIDL_COMMON_FAVORITES = &H1F
		//CSIDL_INTERNET_CACHE = &H20
		//CSIDL_COOKIES = &H21
		//CSIDL_HISTORY = &H22
		//CSIDL_COMMON_APPDATA = &H23
		//CSIDL_WINDOWS = &H24
		//CSIDL_SYSTEM = &H25
		//CSIDL_PROGRAM_FILES = &H26
		//CSIDL_MYPICTURES = &H27
		//CSIDL_PROFILE = &H28
		//CSIDL_SYSTEMX86 = &H29
		//CSIDL_PROGRAM_FILESX86 = &H2A
		//CSIDL_PROGRAM_FILES_COMMON = &H2B
		//CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
		//CSIDL_COMMON_TEMPLATES = &H2D
		//CSIDL_COMMON_DOCUMENTS = &H2E
		//CSIDL_COMMON_ADMINTOOLS = &H2F
		//CSIDL_ADMINTOOLS = &H30
		//CSIDL_CONNECTIONS = &H31
		//CSIDL_COMMON_MUSIC = &H35
		//CSIDL_COMMON_PICTURES = &H36
		//CSIDL_COMMON_VIDEO = &H37
		//CSIDL_RESOURCES = &H38
		//CSIDL_RESOURCES_LOCALIZED = &H39
		//CSIDL_COMMON_OEM_LINKS = &H3A
		//CSIDL_CDBURN_AREA = &H3B
		//CSIDL_COMPUTERSNEARME = &H3D
		//CSIDL_FLAG_PER_USER_INIT = &H800
		//CSIDL_FLAG_NO_ALIAS = &H1000
		//CSIDL_FLAG_DONT_VERIFY = &H4000
		//CSIDL_FLAG_CREATE = &H8000
		//CSIDL_FLAG_MASK = &HFF00

		//KNOWNFOLDER IDs for Vista
		//Select Case CSIDL

		//    Case CSIDL_NETWORK, CSIDL_COMPUTERSNEARME  ' VIRTUAL
		//        FOLDERID_NetworkFolder = "{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}"
		//        sRfid = FOLDERID_NetworkFolder
		//    Case CSIDL_DRIVES  ' VIRTUAL
		//        FOLDERID_ComputerFolder = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"
		//        sRfid = FOLDERID_ComputerFolder
		//    Case CSIDL_INTERNET  ' VIRTUAL
		//        FOLDERID_InternetFolder = "{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}"
		//        sRfid = FOLDERID_InternetFolder
		//    Case CSIDL_CONTROLS  ' VIRTUAL
		//        FOLDERID_ControlPanelFolder = "{82A74AEB-AEB4-465C-A014-D097EE346D63}"
		//        sRfid = FOLDERID_ControlPanelFolder
		//    Case CSIDL_PRINTERS  ' VIRTUAL
		//        FOLDERID_PrintersFolder = "{76FC4E2D-D6AD-4519-A663-37BD56068185}"
		//        sRfid = FOLDERID_PrintersFolder
		//    Case 101  ' VIRTUAL
		//        FOLDERID_SyncManagerFolder = "{43668BF8-C14E-49B2-97C9-747784D784B7}"
		//        sRfid = FOLDERID_SyncManagerFolder
		//    Case 102  ' VIRTUAL
		//        FOLDERID_SyncSetupFolder = "{0F214138-B1D3-4a90-BBA9-27CBC0C5389A}"
		//        sRfid = FOLDERID_SyncSetupFolder
		//    Case 103  ' VIRTUAL
		//        FOLDERID_ConflictFolder = "{4bfefb45-347d-4006-a5be-ac0cb0567192}"
		//        sRfid = FOLDERID_ConflictFolder
		//    Case 104  ' VIRTUAL
		//        FOLDERID_SyncResultsFolder = "{289a9a43-be44-4057-a41b-587a76d7e7f9}"
		//        sRfid = FOLDERID_SyncResultsFolder
		//    Case CSIDL_BITBUCKET  ' VIRTUAL
		//        FOLDERID_RecycleBinFolder = "{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}"
		//        sRfid = FOLDERID_RecycleBinFolder
		//    Case CSIDL_CONNECTIONS  ' VIRTUAL
		//        FOLDERID_ConnectionsFolder = "{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}"
		//        sRfid = FOLDERID_ConnectionsFolder
		//        ' VISTA - %windir%\Fonts
		//        ' XP - %windir%\Fonts
		//    Case CSIDL_FONTS  ' FIXED
		//        FOLDERID_Fonts = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
		//        sRfid = FOLDERID_Fonts
		//        ' VISTA - %USERPROFILE%\Desktop
		//        ' XP - %USERPROFILE%\Desktop
		//    Case CSIDL_DESKTOP, CSIDL_DESKTOPDIRECTORY  ' PERUSER
		//        FOLDERID_Desktop = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
		//        sRfid = FOLDERID_Desktop
		//        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu\Programs\StartUp
		//        ' XP - %USERPROFILE%\Start Menu\Programs\StartUp
		//    Case CSIDL_STARTUP, CSIDL_ALTSTARTUP  ' PERUSER
		//        FOLDERID_Startup = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"
		//        sRfid = FOLDERID_Startup
		//        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu\Programs
		//        ' XP - %USERPROFILE%\Start Menu\Programs
		//    Case CSIDL_PROGRAMS  ' PERUSER
		//        FOLDERID_Programs = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"
		//        sRfid = FOLDERID_Programs
		//        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu
		//        ' XP - %USERPROFILE%\Start Menu
		//    Case CSIDL_STARTMENU  'PERUSER
		//        FOLDERID_StartMenu = "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"
		//        sRfid = FOLDERID_StartMenu
		//        ' VISTA - %APPDATA%\Microsoft\Windows\Recent
		//        ' XP - %USERPROFILE%\Recent
		//    Case CSIDL_RECENT  ' PERUSER
		//        FOLDERID_Recent = "{AE50C081-EBD2-438A-8655-8A092E34987A}"
		//        sRfid = FOLDERID_Recent
		//        ' VISTA - %APPDATA%\Microsoft\Windows\SendTo
		//        ' XP - %USERPROFILE%\SendTo
		//    Case CSIDL_SENDTO  ' PERUSER
		//        FOLDERID_SendTo = "{8983036C-27C0-404B-8F08-102D10DCFD74}"
		//        sRfid = FOLDERID_SendTo
		//        ' VISTA - %USERPROFILE%\Documents
		//        ' XP - %USERPROFILE%\My Documents
		//    Case CSIDL_MYDOCUMENTS, CSIDL_PERSONAL  ' PERUSER
		//        FOLDERID_Documents = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
		//        sRfid = FOLDERID_Documents
		//        ' VISTA - %USERPROFILE%\Documents
		//        ' XP - %USERPROFILE%\My Documents
		//    Case CSIDL_FAVORITES, CSIDL_COMMON_FAVORITES  ' PERUSER
		//        FOLDERID_Favorites = "{1777F761-68AD-4D8A-87BD-30B759FA33DD}"
		//        sRfid = FOLDERID_Favorites
		//        ' VISTA - %APPDATA%\Microsoft\Windows\Network Shortcuts
		//        ' XP - %USERPROFILE%\NetHood
		//    Case CSIDL_NETHOOD  ' PERUSER
		//        FOLDERID_NetHood = "{C5ABBF53-E17F-4121-8900-86626FC2C973}"
		//        sRfid = FOLDERID_NetHood
		//        ' VISTA - %APPDATA%\Microsoft\Windows\Printer Shortcuts
		//        ' XP - %USERPROFILE%\PrintHood
		//    Case CSIDL_PRINTHOOD  ' PERUSER
		//        FOLDERID_PrintHood = "{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"
		//        sRfid = FOLDERID_PrintHood
		//        ' VISTA - %APPDATA%\Microsoft\Windows\Templates
		//        ' XP - %USERPROFILE%\Templates
		//    Case CSIDL_TEMPLATES  ' PERUSER
		//        FOLDERID_Templates = "{A63293E8-664E-48DB-A079-DF759E0509F7}"
		//        sRfid = FOLDERID_Templates
		//        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\StartUp
		//        ' XP - %ALLUSERSPROFILE%\Start Menu\Programs\StartUp
		//    Case CSIDL_COMMON_STARTUP, CSIDL_COMMON_ALTSTARTUP  ' COMMON
		//        FOLDERID_CommonStartup = "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"
		//        sRfid = FOLDERID_CommonStartup
		//        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs
		//        ' XP - %ALLUSERSPROFILE%\Start Menu\Programs
		//    Case CSIDL_COMMON_PROGRAMS  ' COMMON
		//        FOLDERID_CommonPrograms = "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"
		//        sRfid = FOLDERID_CommonPrograms
		//        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu
		//        ' XP - %ALLUSERSPROFILE%\Start Menu
		//    Case CSIDL_COMMON_STARTMENU  ' COMMON
		//        FOLDERID_CommonStartMenu = "{A4115719-D62E-491D-AA7C-E74B8BE3B067}"
		//        sRfid = FOLDERID_CommonStartMenu
		//        ' VISTA - %PUBLIC%\Desktop
		//        ' XP - %ALLUSERSPROFILE%\Desktop
		//    Case 201  ' COMMON
		//        FOLDERID_PublicDesktop = "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"
		//        sRfid = FOLDERID_PublicDesktop
		//        ' VISTA - %ALLUSERSPROFILE% (%ProgramData%, %SystemDrive%\ProgramData)
		//        ' XP - %ALLUSERSPROFILE%\Application Data
		//    Case CSIDL_COMMON_APPDATA  ' FIXED
		//        FOLDERID_ProgramData = "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"
		//        sRfid = FOLDERID_ProgramData
		//        ' VISTA - %ALLUSERSPROFILE%\Templates
		//        ' XP - %ALLUSERSPROFILE%\Templates
		//    Case CSIDL_COMMON_TEMPLATES  ' COMMON
		//        FOLDERID_CommonTemplates = "{B94237E7-57AC-4347-9151-B08C6C32D1F7}"
		//        sRfid = FOLDERID_CommonTemplates
		//        ' VISTA - %PUBLIC%\Documents
		//        ' XP - %ALLUSERSPROFILE%\Documents
		//    Case CSIDL_COMMON_DOCUMENTS  ' COMMON
		//        FOLDERID_PublicDocuments = "{ED4824AF-DCE4-45A8-81E2-FC7965083634}"
		//        sRfid = FOLDERID_PublicDocuments
		//        ' VISTA - %APPDATA% (%USERPROFILE%\AppData\Roaming)
		//        ' XP - %APPDATA% (%USERPROFILE%\Application Data)
		//    Case CSIDL_APPDATA  ' PERUSER
		//        FOLDERID_RoamingAppData = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
		//        sRfid = FOLDERID_RoamingAppData
		//        ' VISTA - %LOCALAPPDATA% (%USERPROFILE%\AppData\Local)
		//        ' XP - %USERPROFILE%\Local Settings\Application Data
		//    Case CSIDL_LOCAL_APPDATA  ' PERUSER
		//        FOLDERID_LocalAppData = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
		//        sRfid = FOLDERID_LocalAppData
		//        ' VISTA - %USERPROFILE%\AppData\LocalLow
		//        ' XP - NONE
		//    Case 301  ' PERUSER
		//        FOLDERID_LocalAppDataLow = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
		//        sRfid = FOLDERID_LocalAppDataLow
		//        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\Temporary Internet Files
		//        ' XP - %USERPROFILE%\Local Settings\Temporary Internet Files
		//    Case CSIDL_INTERNET_CACHE  ' PERUSER
		//        FOLDERID_InternetCache = "{352481E8-33BE-4251-BA85-6007CAEDCF9D}"
		//        sRfid = FOLDERID_InternetCache
		//        ' VISTA - %APPDATA%\Microsoft\Windows\Cookies
		//        ' XP - %USERPROFILE%\Cookies
		//    Case CSIDL_COOKIES  ' PERUSER
		//        FOLDERID_Cookies = "{2B0F765D-C0E9-4171-908E-08A611B84FF6}"
		//        sRfid = FOLDERID_Cookies
		//        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\History
		//        ' XP - %USERPROFILE%\Local Settings\History
		//    Case CSIDL_HISTORY  ' PERUSER
		//        FOLDERID_History = "{D9DC8A3B-B784-432E-A781-5A1130A75963}"
		//        sRfid = FOLDERID_History
		//        ' VISTA - %windir%\system32
		//        ' XP - %windir%\system32
		//    Case CSIDL_SYSTEM  ' FIXED
		//        FOLDERID_System = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
		//        sRfid = FOLDERID_System
		//        ' VISTA - %windir%\system32
		//        ' XP - %windir%\system32
		//    Case CSIDL_SYSTEMX86  ' FIXED
		//        FOLDERID_SystemX86 = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"
		//        sRfid = FOLDERID_SystemX86
		//        ' VISTA - %windir%
		//        ' XP - %windir%
		//    Case CSIDL_WINDOWS  ' FIXED
		//        FOLDERID_Windows = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
		//        sRfid = FOLDERID_Windows
		//        ' VISTA - %USERPROFILE% (%SystemDrive%\Users\%USERNAME%)
		//        ' XP - %USERPROFILE% (%SystemDrive%\Documents and Settings\%USERNAME%)
		//    Case CSIDL_PROFILE  ' FIXED
		//        FOLDERID_Profile = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
		//        sRfid = FOLDERID_Profile
		//        ' VISTA - %USERPROFILE%\Pictures
		//        ' XP - %USERPROFILE%\My Documents\My Pictures
		//    Case CSIDL_MYPICTURES  ' PERUSER
		//        FOLDERID_Pictures = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
		//        sRfid = FOLDERID_Pictures
		//        ' VISTA - %ProgramFiles% (%SystemDrive%\Program Files)
		//        ' XP - %ProgramFiles% (%SystemDrive%\Program Files)
		//    Case CSIDL_PROGRAM_FILESX86  ' FIXED
		//        FOLDERID_ProgramFilesX86 = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
		//        sRfid = FOLDERID_ProgramFilesX86
		//        ' VISTA - %ProgramFiles%\Common Files
		//        ' XP - %ProgramFiles%\Common Files
		//    Case CSIDL_PROGRAM_FILES_COMMONX86  ' FIXED
		//        FOLDERID_ProgramFilesCommonX86 = "{DE974D24-D9C6-4D3E-BF91-F4455120B917}"
		//        sRfid = FOLDERID_ProgramFilesCommonX86
		//        ' VISTA - %ProgramFiles% (%SystemDrive%\Program Files)
		//        ' XP - %ProgramFiles% (%SystemDrive%\Program Files)
		//    Case 401  ' FIXED
		//        FOLDERID_ProgramFilesX64 = "{6D809377-6AF0-444b-8957-A3773F02200E}"
		//        sRfid = FOLDERID_ProgramFilesX64
		//        ' VISTA - %ProgramFiles%\Common Files
		//        ' XP - %ProgramFiles%\Common Files
		//    Case 402
		//        FOLDERID_ProgramFilesCommonX64 = "{6365D5A7-0F0D-45e5-87F6-0DA56B6A4F7D}"
		//        sRfid = FOLDERID_ProgramFilesCommonX64
		//        ' VISTA - %ProgramFiles% (%SystemDrive%\Program Files)
		//        ' XP - %ProgramFiles% (%SystemDrive%\Program Files)
		//    Case CSIDL_PROGRAM_FILES  ' FIXED
		//        FOLDERID_ProgramFiles = "{905e63b6-c1bf-494e-b29c-65b732d3d21a}"
		//        sRfid = FOLDERID_ProgramFiles
		//        ' VISTA - %ProgramFiles%\Common Files
		//        ' XP - %ProgramFiles%\Common Files
		//    Case CSIDL_PROGRAM_FILES_COMMON  ' FIXED
		//        FOLDERID_ProgramFilesCommon = "{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"
		//        sRfid = FOLDERID_ProgramFilesCommon
		//        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu\Programs\Administrative Tools
		//        ' XP - %USERPROFILE%\Start Menu\Programs\Administrative Tools
		//    Case CSIDL_ADMINTOOLS  ' PERUSER
		//        FOLDERID_AdminTools = "{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"
		//        sRfid = FOLDERID_AdminTools
		//        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Administrative Tools
		//        ' XP - %ALLUSERSPROFILE%\Start Menu\Programs\Administrative Tools
		//    Case CSIDL_COMMON_ADMINTOOLS  ' COMMON
		//        FOLDERID_CommonAdminTools = "{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"
		//        sRfid = FOLDERID_CommonAdminTools
		//        ' VISTA - %USERPROFILE%\Music
		//        ' XP - %USERPROFILE%\My Documents\My Music
		//    Case CSIDL_MYMUSIC  ' PERUSER
		//        FOLDERID_Music = "{4BD8D571-6D19-48D3-BE97-422220080E43}"
		//        sRfid = FOLDERID_Music
		//        ' VISTA - %USERPROFILE%\Videos
		//        ' XP - %USERPROFILE%\My Documents\My Videos
		//    Case CSIDL_MYVIDEO  ' PERUSER
		//        FOLDERID_Videos = "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"
		//        sRfid = FOLDERID_Videos
		//        ' VISTA - %PUBLIC%\Pictures
		//        ' XP - %ALLUSERSPROFILE%\Documents\My Pictures
		//    Case CSIDL_COMMON_PICTURES  ' COMMON
		//        FOLDERID_PublicPictures = "{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"
		//        sRfid = FOLDERID_PublicPictures
		//        ' VISTA - %PUBLIC%\Music
		//        ' XP - %ALLUSERSPROFILE%\Documents\My Music
		//    Case CSIDL_COMMON_MUSIC  ' COMMON
		//        FOLDERID_PublicMusic = "{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"
		//        sRfid = FOLDERID_PublicMusic
		//        'VISTA - %PUBLIC%\Videos
		//        ' XP - %ALLUSERSPROFILE%\Documents\My Videos
		//    Case CSIDL_COMMON_VIDEO  ' COMMON
		//        FOLDERID_PublicVideos = "{2400183A-6185-49FB-A2D8-4A392A602BA3}"
		//        sRfid = FOLDERID_PublicVideos
		//        ' VISTA - %windir%\Resources
		//        ' XP - %windir%\Resources
		//    Case CSIDL_RESOURCES  ' FIXED
		//        FOLDERID_ResourceDir = "{8AD10C31-2ADB-4296-A8F7-E4701232C972}"
		//        sRfid = FOLDERID_ResourceDir
		//        ' VISTA - %windir%\resources\0409 (code page)
		//        ' XP - %windir%\resources\0409 (code page)
		//    Case CSIDL_RESOURCES_LOCALIZED  ' FIXED
		//        FOLDERID_LocalizedResourcesDir = "{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"
		//        sRfid = FOLDERID_LocalizedResourcesDir
		//        ' VISTA - %ALLUSERSPROFILE%\OEM Links
		//        ' XP - %ALLUSERSPROFILE%\OEM Links
		//    Case CSIDL_COMMON_OEM_LINKS  ' COMMON
		//        FOLDERID_CommonOEMLinks = "{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"
		//        sRfid = FOLDERID_CommonOEMLinks
		//        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\Burn\Burn
		//        ' XP - %USERPROFILE%\Local Settings\Application Data\Microsoft\CD Burning
		//    Case CSIDL_CDBURN_AREA  ' PERUSER
		//        FOLDERID_CDBurning = "{9E52AB10-F80D-49DF-ACB8-4330F5687855}"
		//        sRfid = FOLDERID_CDBurning
		//        ' VISTA - %SystemDrive%\Users
		//        ' XP - NONE
		//    Case 501  ' FIXED
		//        FOLDERID_UserProfiles = "{0762D272-C50A-4BB0-A382-697DCD729B80}"
		//        sRfid = FOLDERID_UserProfiles
		//        ' VISTA - %USERPROFILE%\Music\Playlists
		//        ' XP - NONE
		//    Case 502  ' PERUSER
		//        FOLDERID_Playlists = "{DE92C1C7-837F-4F69-A3BB-86E631204A23}"
		//        sRfid = FOLDERID_Playlists
		//        ' VISTA - %PUBLIC%\Music\Sample Playlists
		//        ' XP - NONE
		//    Case 503  ' COMMON
		//        FOLDERID_SamplePlaylists = "{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"
		//        sRfid = FOLDERID_SamplePlaylists
		//        ' VISTA - %PUBLIC%\Music\Sample Music
		//        ' XP - %ALLUSERSPROFILE%\Documents\My Music\Sample Music
		//    Case 504  ' COMMON
		//        FOLDERID_SampleMusic = "{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"
		//        sRfid = FOLDERID_SampleMusic
		//        ' VISTA - %PUBLIC%\Pictures\Sample Pictures
		//        ' XP - %ALLUSERSPROFILE%\Documents\My Pictures\Sample Pictures
		//    Case 505  ' COMMON
		//        FOLDERID_SamplePictures = "{C4900540-2379-4C75-844B-64E6FAF8716B}"
		//        sRfid = FOLDERID_SamplePictures
		//        ' VISTA - %PUBLIC%\Videos\Sample Videos
		//        ' XP - NONE
		//    Case 506  ' COMMON
		//        FOLDERID_SampleVideos = "{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"
		//        sRfid = FOLDERID_SampleVideos
		//        ' VISTA - %USERPROFILE%\Pictures\Slide Shows
		//        ' XP - NONE
		//    Case 507  ' PERUSER
		//        FOLDERID_PhotoAlbums = "{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"
		//        sRfid = FOLDERID_PhotoAlbums
		//        ' VISTA - %PUBLIC% (%SystemDrive%\Users\Public)
		//        ' XP - NONE
		//    Case 508  'FIXED
		//        FOLDERID_Public = "{DFDF76A2-C82A-4D63-906A-5644AC457385}"
		//        sRfid = FOLDERID_Public
		//    Case 509  ' VIRTUAL
		//        FOLDERID_ChangeRemovePrograms = "{df7266ac-9274-4867-8d55-3bd661de872d}"
		//        sRfid = FOLDERID_ChangeRemovePrograms
		//    Case 510  ' VIRTUAL
		//        FOLDERID_AppUpdates = "{a305ce99-f527-492b-8b1a-7e76fa98d6e4}"
		//        sRfid = FOLDERID_AppUpdates
		//    Case 511  ' VIRTUAL
		//        FOLDERID_AddNewPrograms = "{de61d971-5ebc-4f02-a3a9-6c82895e5c04}"
		//        sRfid = FOLDERID_AddNewPrograms
		//        ' VISTA - %USERPROFILE%\Downloads
		//        ' XP - NONE
		//    Case 512  ' PERUSER
		//        FOLDERID_Downloads = "{374DE290-123F-4565-9164-39C4925E467B}"
		//        sRfid = FOLDERID_Downloads
		//        ' VISTA - %PUBLIC%\Downloads
		//        ' XP - NONE
		//    Case 513  ' COMMON
		//        FOLDERID_PublicDownloads = "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}"
		//        sRfid = FOLDERID_PublicDownloads
		//        ' VISTA - %USERPROFILE%\Searches
		//        ' XP - NONE
		//    Case 514  ' PERUSER
		//        FOLDERID_SavedSearches = "{7d1d3a04-debb-4115-95cf-2f29da2920da}"
		//        sRfid = FOLDERID_SavedSearches
		//        'VISTA - %APPDATA%\Microsoft\Internet Explorer\Quick Launch
		//        'XP - %APPDATA%\Microsoft\Internet Explorer\Quick Launch
		//    Case 515  ' PERUSER
		//        FOLDERID_QuickLaunch = "{52a4f021-7b75-48a9-9f6b-4b87a210bc8f}"
		//        sRfid = FOLDERID_QuickLaunch
		//        ' VISTA - %USERPROFILE%\Contacts
		//        ' XP - NONE
		//    Case 516  ' PERUSER
		//        FOLDERID_Contacts = "{56784854-C6CB-462b-8169-88E350ACB882}"
		//        sRfid = FOLDERID_Contacts
		//        ' VISTA -%LOCALAPPDATA%\Microsoft\Windows Sidebar\Gadgets
		//        ' XP - NONE
		//    Case 517  ' PERUSER
		//        FOLDERID_SidebarParts = "{A75D362E-50FC-4fb7-AC2C-A8BEAA314493}"
		//        sRfid = FOLDERID_SidebarParts
		//        ' VISTA - %ProgramFiles%\Windows Sidebar\Gadgets
		//        ' XP - NONE
		//    Case 518  ' COMMON
		//        FOLDERID_SidebarDefaultParts = "{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"
		//        sRfid = FOLDERID_SidebarDefaultParts
		//    Case 519  ' NOT USED
		//        FOLDERID_TreeProperties = "{5b3749ad-b49f-49c1-83eb-15370fbd4882}"
		//        sRfid = FOLDERID_TreeProperties
		//        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\GameExplorer
		//        ' XP - NONE
		//    Case 520  ' COMMON
		//        FOLDERID_PublicGameTasks = "{DEBF2536-E1A8-4c59-B6A2-414586476AEA}"
		//        sRfid = FOLDERID_PublicGameTasks
		//        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\GameExplorer
		//        ' XP - NONE
		//    Case 521  ' PERUSER
		//        FOLDERID_GameTasks = "{054FAE61-4DD8-4787-80B6-090220C4B700}"
		//        sRfid = FOLDERID_GameTasks
		//        ' VISTA - %USERPROFILE%\Saved Games
		//        ' XP - NONE
		//    Case 522  ' PERUSER
		//        FOLDERID_SavedGames = "{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}"
		//        sRfid = FOLDERID_SavedGames
		//    Case 523  ' VIRTUAL
		//        FOLDERID_Games = "{CAC52C1A-B53D-4edc-92D7-6B2E8AC19434}"
		//        sRfid = FOLDERID_Games
		//    Case 524  ' NOT USED
		//        FOLDERID_RecordedTV = "{bd85e001-112e-431e-983b-7b15ac09fff1}"
		//        sRfid = FOLDERID_RecordedTV
		//    Case 525  ' VIRTUAL
		//        FOLDERID_SEARCH_MAPI = "{98ec0e18-2098-4d44-8644-66979315a281}"
		//        sRfid = FOLDERID_SEARCH_MAPI
		//    Case 526  ' VIRTUAL
		//        FOLDERID_SEARCH_CSC = "{ee32e446-31ca-4aba-814f-a5ebd2fd6d5e}"
		//        sRfid = FOLDERID_SEARCH_CSC
		//        ' VISTA - %USERPROFILE%\Links
		//        ' XP - NONE
		//    Case 527  ' PERUSER
		//        FOLDERID_Links = "{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}"
		//        sRfid = FOLDERID_Links
		//    Case 528  ' VIRTUAL
		//        FOLDERID_UsersFiles = "{f3ce0f7c-4901-4acc-8648-d5d44b04ef8f}"
		//        sRfid = FOLDERID_UsersFiles
		//    Case 529  ' VIRTUAL
		//        FOLDERID_SearchHome = "{190337d1-b8ca-4121-a639-6d472d16972a}"
		//        sRfid = FOLDERID_SearchHome
		//        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows Photo Gallery\Original Images
		//        ' XP - NONE
		//    Case 530  ' PERUSER
		//        FOLDERID_OriginalImages = "{2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}"
		//        sRfid = FOLDERID_OriginalImages
		//End Select
		long Ret = 0;
		string Trash = null;
		Trash = Strings.Space(260);
		functionReturnValue = "";
		try {
			Ret = SHGetSpecialFolderPath(0, Trash, CSIDL, false);
			if (Strings.Trim(Trash) != Strings.Chr(0)) {
				Trash = Strings.Left(Trash, Strings.InStr(Trash, Strings.Chr(0)) - 1);
			}
			functionReturnValue = Trash;
			return functionReturnValue;
		}
		catch (Exception ex) {
			Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrUndefinedSpecialFolder, modTrial.ACTIVELOCKSTRING, STRUNDEFINEDSPECIALFOLDER);
		}
		return functionReturnValue;

	}

	/// <summary>
	/// ?Not Documented!
	/// </summary>
	/// <returns></returns>
	/// <remarks></remarks>
	public static string PSWD()
	{
		// Do not modify this unless you change all encrypted strings in the entire project
		return Strings.Chr(109) + Strings.Chr(121) + Strings.Chr(108) + Strings.Chr(111) + Strings.Chr(118) + Strings.Chr(101) + Strings.Chr(97) + Strings.Chr(99) + Strings.Chr(116) + Strings.Chr(105) + Strings.Chr(118) + Strings.Chr(101) + "lock";
	}

	/// <summary>
	/// ?Not Documented!
	/// </summary>
	/// <param name="n1"></param>
	/// <param name="n2"></param>
	/// <returns></returns>
	/// <remarks></remarks>
	public static bool IsNumberIncluded(long n1, long n2)
	{
		bool functionReturnValue = false;
		// n1 = the larger number which may include n2
		// n2 = the number we're checking as "is this a component?"
		string binary1 = string.Empty;
		string binary2 = string.Empty;
		functionReturnValue = false;

		if (n1 < n2) {
			return;
		}
		else if (n1 <= 0 | n2 <= 0) {
			return;
		}
		else if (n1 == n2) {
			functionReturnValue = true;
		}

		else {
			while (!(n1 == 0)) {
				if ((n1 % 2)) {
					binary1 = binary1 + "1";
					//write binary number BACKWARDS
				}
				else {
					binary1 = binary1 + "0";
				}
				n1 = n1 / 2;
			}

			while (!(n2 == 0)) {
				if ((n2 % 2)) {
					binary2 = binary2 + "1";
					//write binary number BACKWARDS
				}
				else {
					binary2 = binary2 + "0";
				}
				n2 = n2 / 2;
			}
			functionReturnValue = (bool)Strings.Mid(binary1, Strings.Len(binary2), 1) == "1";
			if (binary1.Substring(binary2.Length - 1, 1) == "1") functionReturnValue = true; 

		}
		return functionReturnValue;

	}

	/// <summary>
	/// Checks to see if there is a web connection.
	/// </summary>
	/// <returns>True if connected, False otherwise.</returns>
	/// <remarks>This also sets the ConnectionQualityString</remarks>
	public static bool IsWebConnected()
	{
		// Returns True if connection is available

		long lngFlags = 0;

		if (InternetGetConnectedState(ref lngFlags, 0)) {
			return true;
			// True
			if (lngFlags & InetConnState.lan) {
				switch (ConnectionQualityString) {
					case "Good":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Intermittent":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Off":
						//lblConnectStatus.ForeColor = Color.DarkOrange
						//lblConnectStatus.Text = _
						//     "Connection Quality:  Intermittent"
						ConnectionQualityString = "Intermittent";
						break;
				}
			}
			else if (lngFlags & InetConnState.modem) {
				switch (ConnectionQualityString) {
					case "Good":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Intermittent":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Off":
						//lblConnectStatus.ForeColor = Color.DarkOrange
						//lblConnectStatus.Text = _
						//     "Connection Quality:  Intermittent"
						ConnectionQualityString = "Intermittent";
						break;
				}
			}
			else if (lngFlags & InetConnState.configured) {
				switch (ConnectionQualityString) {
					case "Good":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Intermittent":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Off":
						//lblConnectStatus.ForeColor = Color.DarkOrange
						//lblConnectStatus.Text = _
						//     "Connection Quality:  Intermittent"
						ConnectionQualityString = "Intermittent";
						break;
				}
			}
			else if (lngFlags & InetConnState.proxy) {
				switch (ConnectionQualityString) {
					case "Good":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Intermittent":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Off":
						//lblConnectStatus.ForeColor = Color.DarkOrange
						//lblConnectStatus.Text = _
						//     "Connection Quality:  Intermittent"
						ConnectionQualityString = "Intermittent";
						break;
				}
			}
			else if (lngFlags & InetConnState.ras) {
				switch (ConnectionQualityString) {
					case "Good":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Intermittent":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Off":
						//lblConnectStatus.ForeColor = Color.DarkOrange
						//lblConnectStatus.Text = _
						//     "Connection Quality:  Intermittent"
						ConnectionQualityString = "Intermittent";
						break;
				}
			}
			else if (lngFlags & InetConnState.offline) {
				switch (ConnectionQualityString) {
					case "Good":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Intermittent":
						//lblConnectStatus.ForeColor = Color.Green
						//lblConnectStatus.Text = "Connection Quality:  Good"
						ConnectionQualityString = "Good";
						break;
					case "Off":
						//lblConnectStatus.ForeColor = Color.DarkOrange
						//lblConnectStatus.Text = _
						//     "Connection Quality:  Intermittent"
						ConnectionQualityString = "Intermittent";
						break;
				}
			}
		}
		else {
			// False
			switch (ConnectionQualityString) {
				case "Good":
					//lblConnectStatus.ForeColor = Color.DarkOrange
					//lblConnectStatus.Text = "Connection Quality:  Intermittent"
					ConnectionQualityString = "Intermittent";
					break;
				case "Intermittent":
					//lblConnectStatus.ForeColor = Color.Red
					//lblConnectStatus.Text = "Connection Quality:  Off"
					ConnectionQualityString = "Off";
					break;
				case "Off":
					//lblConnectStatus.ForeColor = Color.Red
					//lblConnectStatus.Text = "Connection Quality:  Off"
					ConnectionQualityString = "Off";
					break;
			}
		}

	}

}
