Option Strict Off
Option Explicit On

Imports System.IO
Imports System.Security.Cryptography
Imports System.Text

#Region "Copyright"
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
#End Region

''' <summary>
''' <para>This module contains common utility routines that can be shared between
''' ActiveLock and the client application.</para>
''' </summary>
''' <remarks>See the TODO List, within</remarks>
Module modActiveLock

    ' Started: 04.21.2005
    ' Modified: 03.25.2006
    '===============================================================================
    ' @author activelock-admins
    ' @version 3.3.0
    ' @date 03.25.2006
    '
    '* ///////////////////////////////////////////////////////////////////////
    '  /                        MODULE TO DO LIST                            /
    '  ///////////////////////////////////////////////////////////////////////
    '
    ' @bug rsa_createkey() sometimes causes crash.  This is due to a bug in
    '      ALCrypto3.dll in which a bad keyset is sometimes generated
    '      (either caused by <code>rsa_generate()</code> or one of <code>rsa_private_key_blob()</code>
    '      and <code>rsa_public_key_blob()</code>--we're not sure which is the culprit yet.
    '      This causes the <code>rsa_createkey()</code> call encryption routines to crash.
    '      The work-around for the time being is to keep regenerating the keyset
    '      until eventually you'll get a valid keyset that no longer causes a crash.
    '      You only need to go through this keyset generation step once.
    '      Once you have a valid keyset, you should store it inside your app for later use.
    '

    Public Const STRKEYSTOREINVALID As String = "A license property contains an invalid value."
	Public Const STRLICENSEEXPIRED As String = "License expired."
	Public Const STRLICENSEINVALID As String = "License invalid."
	Public Const STRNOLICENSE As String = "No valid license."
	Public Const STRLICENSETAMPERED As String = "License may have been tampered."
	Public Const STRNOTINITIALIZED As String = "ActiveLock has not been initialized."
	Public Const STRNOTIMPLEMENTED As String = "Not implemented."
	Public Const STRCLOCKCHANGED As String = STRLICENSEINVALID & " System clock has been tampered."
	Public Const STRINVALIDTRIALDAYS As String = "Zero Free Trial days allowed."
	Public Const STRINVALIDTRIALRUNS As String = "Zero Free Trial runs allowed."
	Public Const STRFILETAMPERED As String = "Alcrypto3.dll has been tampered."
	Public Const STRKEYSTOREUNINITIALIZED As String = "Key Store Provider hasn't been initialized yet."
    Public Const STRKEYSTOREPATHISEMPTY As String = "Key Store Path (LIC file path) not specified."
    Public Const STRNOSOFTWARECODE As String = "Software code has not been set."
	Public Const STRNOSOFTWARENAME As String = "Software Name has not been set."
	Public Const STRNOSOFTWAREVERSION As String = "Software Version has not been set."
    Public Const STRNOSOFTWAREPASSWORD As String = "Software Password has not been set."
    Public Const STRUSERNAMETOOLONG As String = "User Name > 2000 characters."
    Public Const STRUSERNAMEINVALID As String = "User Name invalid."
    Public Const STRRSAERROR As String = "Internal RSA Error."
    Public Const RETVAL_ON_ERROR As Integer = -999
    Public Const STRWRONGIPADDRESS As String = "Wrong IP Address."
    Public Const STRUNDEFINEDSPECIALFOLDER As String = "Undefined Special Folder."
    Public Const STRDATEERROR As String = "Date Error."
    Public Const STRINTERNETNOTCONNECTED As String = "Internet Connection is Required. Please Connect and Try Again."
    Public Const STRSOFTWAREPASSWORDINVALID As String = "Password length>255 or invalid characters."
    Public Const STRNOKEYREUSE As String = "You're not allowed to reuse an old license key. Obtain new key."
    Public Const STRNOLICENSEFILE As String = "License does not exist"
    Public Const STREXPIREDPERMANENTLY As String = "This license has expired permanently. Obtain new key."

    ''' <summary>
    ''' RSA encrypts the data.
    ''' </summary>
    ''' <param name="CryptType">0 for public&#59; 1 for private</param>
    ''' <param name="data">Data to be encrypted</param>
    ''' <param name="dLen">[in/out] Length of data, in bytes. This parameter will contain length of encrypted data when returned.</param>
    ''' <param name="ptrKey">Key to be used for encryption</param>
    ''' <returns>Integer - ?Not Documented</returns>
    ''' <remarks></remarks>
    Public Declare Function rsa_encrypt Lib "ALCrypto3NET" (ByVal CryptType As Integer, ByVal data As String, ByRef dLen As Integer, ByRef ptrKey As RSAKey) As Integer

    ''' <summary>
    ''' RSA decrypts the data.
    ''' </summary>
    ''' <param name="CryptType">0 for public&#59; 1 for private</param>
    ''' <param name="data">Data to be decrypted</param>
    ''' <param name="dLen">[in/out] Length of data, in bytes. This parameter will contain length of decrypted data when returned.</param>
    ''' <param name="ptrKey">Key to be used for encryption</param>
    ''' <returns>Integer - ?Not Documented!</returns>
    ''' <remarks></remarks>
    Public Declare Function rsa_decrypt Lib "ALCrypto3NET" (ByVal CryptType As Integer, ByVal data As String, ByRef dLen As Integer, ByRef ptrKey As RSAKey) As Integer
	
    ''' <summary>
    ''' Computes an MD5 hash from the data.
    ''' </summary>
    ''' <param name="inData">Data to be hashed</param>
    ''' <param name="nDataLen">Length of inData</param>
    ''' <param name="outData">[out] 32-byte Computed hash code</param>
    ''' <returns>Integer - ?Undocumented!</returns>
    ''' <remarks></remarks>
    Public Declare Function md5_hash Lib "ALCrypto3NET" (ByVal inData As String, ByVal nDataLen As Integer, ByVal outData As String) As Integer

    ''' <summary>
    ''' ActiveLock Encryption Key
    ''' </summary>
    ''' <remarks>!!!WARNING!!! It is highly recommended that you change this key for your version of ActiveLock before deploying your app.</remarks>
    Public Const ENCRYPT_KEY As String = "AAAAgEPRFzhQEF7S91vt2K6kOcEdDDe5BfwNiEL30/+ozTFHc7cZctB8NIlS++ZR//D3AjSMqScjh7xUF/gwvUgGCjiExjj1DF/XWFWnPOCfF8UxYAizCLZ9fdqxb1FRpI5NoW0xxUmvxGjmxKwazIW4P4XVi/+i1Bvh2qQ6ri3whcsNAAAAQQCyWGsbJKO28H2QLYH+enb7ehzwBThqfAeke/Gv1Te95yIAWme71I9aCTTlLsmtIYSk9rNrp3sh9ItD2Re67SE7AAAAQQCAookH1nws1gS2XP9cZTPaZEmFLwuxlSVsLQ5RWmd9cuxpgw5y2gIskbL4c+4oBuj0IDwKtnMrZq7UfV9I5VfVAAAAQQCEnyAuO0ahXH3KhAboop9+tCmRzZInTrDYdMy23xf3PLCLd777dL/Y2Y+zmaH1VO03m6iOog7WLiN4dCL7m+Im" ' RSA Private Key

    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MAGICNUMBER_YES As Integer = &HEFCDAB89
    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <remarks></remarks>
    Public Const MAGICNUMBER_NO As Integer = &H98BADCFE

    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <remarks></remarks>
    Private Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <remarks></remarks>
    Private Const KEY_CONTAINER As String = "ActiveLock"
    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <remarks></remarks>
    Private Const PROV_RSA_FULL As Integer = 1

    ''' <summary>
    ''' flag to indicate that module initialization has been done
    ''' </summary>
    ''' <remarks></remarks>
    Private fInit As Boolean
    ''' <summary>
    ''' The RtlMoveMemory routine moves memory either forward or backward, aligned or unaligned, in 4-byte blocks, followed by any remaining bytes.
    ''' </summary>
    ''' <param name="Destination">Pointer to the destination of the move.</param>
    ''' <param name="source">Pointer to the memory to be copied.</param>
    ''' <param name="length">Specifies the number of bytes to be copied.</param>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/ms803004.aspx for more Information</remarks>
    Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Integer, ByRef source As Integer, ByVal length As Integer)
    ''' <summary>
    ''' Retrieves the fully-qualified path for the file that contains the specified module. The module must have been loaded by the current process.
    ''' </summary>
    ''' <param name="hModule">[in, optional] A handle to the loaded module whose path is being requested. If this parameter is NULL, GetModuleFileName retrieves the path of the executable file of the current process.</param>
    ''' <param name="lpFileName">[out] A pointer to a buffer that receives the fully-qualified path of the module. If the length of the path is less than the size that the nSize parameter specifies, the function succeeds and the path is returned as a null-terminated string.</param>
    ''' <param name="nSize">[in] The size of the lpFilename buffer, in TCHARs.</param>
    ''' <returns>If the function succeeds, the return value is the length of the string that is copied
    ''' to the buffer, in characters, not including the terminating null character. If the buffer is too
    ''' small to hold the module name, the string is truncated to nSize characters including the
    ''' terminating null character, the function returns nSize, and the function sets the last error
    ''' to ERROR_INSUFFICIENT_BUFFER.</returns>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/ms683197(VS.85).aspx for full Documentation!</remarks>
    Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Integer, ByVal lpFileName As String, ByVal nSize As Integer) As Integer

    ' TODO: Check to ensure that the following enum is correct and modify the MapFileAndCheckSum Function, japreja!
    ''' <summary>
    ''' Not Inplimented - CheckSum return Values for MapFileAndCheckSum
    ''' </summary>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/ms680355(VS.85).aspx</remarks>
    Public Enum CheckSumReturnValues
        ' TODO: Impliment Me! 
        ''' <summary>
        ''' Success!
        ''' </summary>
        ''' <remarks></remarks>
        CHECKSUM_SUCCESS = 0
        ''' <summary>
        ''' Could not open the file.
        ''' </summary>
        ''' <remarks></remarks>
        CHECKSUM_OPEN_FAILURE = 1
        ''' <summary>
        ''' Could not map the file.
        ''' </summary>
        ''' <remarks></remarks>
        CHECKSUM_MAP_FAILURE = 2
        ''' <summary>
        ''' Could not map a view of the file.
        ''' </summary>
        ''' <remarks></remarks>
        CHECKSUM_MAPVIEW_FAILURE = 3
        ''' <summary>
        ''' Could not convert the file name to Unicode.
        ''' </summary>
        ''' <remarks></remarks>
        CHECKSUM_UNICODE_FAILURE = 4
    End Enum

    ''' <summary>
    ''' Computes the checksum of the specified file
    ''' </summary>
    ''' <param name="FileName">[in] The file name of the file for which the checksum is to be computed.</param>
    ''' <param name="HeaderSum">[out] A pointer to a variable that receives the original checksum from the image file, or zero if there is an error.</param>
    ''' <param name="CheckSum">[out] A pointer to a variable that receives the computed checksum.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Declare Function MapFileAndCheckSum Lib "imagehlp" Alias "MapFileAndCheckSumA" (ByVal FileName As String, ByRef HeaderSum As Integer, ByRef CheckSum As Integer) As Integer

    ''' <summary>
    ''' <para>Retrieves the path of a special folder, identified by its <a href="http://msdn.microsoft.com/en-us/library/bb762494(VS.85).aspx">CSIDL</a>.</para>
    ''' <para>Note: In Windows Vista, these values have been replaced by <a href="http://msdn.microsoft.com/en-us/library/bb762584(VS.85).aspx">KNOWNFOLDERID</a> values.
    ''' See that topic for a list of the new constants and their corresponding CSIDL values.  For
    ''' convenience, corresponding KNOWNFOLDERID values are also noted here for each CSIDL value.  The
    ''' CSIDL system is supported under Windows Vista for compatibility reasons. However, new development
    ''' should use KNOWNFOLDERID values rather than CSIDL values.</para>
    ''' </summary>
    ''' <param name="hWnd">Reserved.</param>
    ''' <param name="lpszPath">[out] A pointer to a null-terminated string that receives the drive and path of the specified folder. This buffer must be at least MAX_PATH characters in size.</param>
    ''' <param name="nFolder">See notes at http://msdn.microsoft.com/en-us/library/bb762494(VS.85).aspx</param>
    ''' <param name="fCreate">[in] Indicates whether the folder should be created if it does not already exist. If this value is nonzero, the folder will be created. If this value is zero, the folder will not be created.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As IntPtr, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean

    ''' <summary>
    ''' Specifies a date and time, using individual members for the month, day, year, weekday, hour, minute, second, and millisecond. The time is either in coordinated universal time (UTC) or local time, depending on the function that is being called.
    ''' </summary>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/ms724950(VS.85).aspx for full documentation!</remarks>
    Structure SYSTEMTIME
        ''' <summary>
        ''' The year. The valid values for this member are 1601 through 30827.
        ''' </summary>
        ''' <remarks></remarks>
        Dim wYear As Short
        ''' <summary>
        ''' Month. - The valid values for this member are 1 through 12 corresponding to Jan. through Dec.
        ''' </summary>
        ''' <remarks></remarks>
        Dim wMonth As Short
        ''' <summary>
        ''' Day of the week. The valid values for this member are 0 through 6, with Sunday being 0!
        ''' </summary>
        ''' <remarks></remarks>
        Dim wDayOfWeek As Short
        ''' <summary>
        ''' The day of the month. The valid values for this member are 1 through 31.
        ''' </summary>
        ''' <remarks></remarks>
        Dim wDay As Short
        ''' <summary>
        ''' The hour. The valid values for this member are 0 through 23.
        ''' </summary>
        ''' <remarks></remarks>
        Dim wHour As Short
        ''' <summary>
        ''' The minute. The valid values for this member are 0 through 59.
        ''' </summary>
        ''' <remarks></remarks>
        Dim wMinute As Short
        ''' <summary>
        ''' The second. The valid values for this member are 0 through 59.
        ''' </summary>
        ''' <remarks></remarks>
        Dim wSecond As Short
        ''' <summary>
        ''' The millisecond. The valid values for this member are 0 through 999.
        ''' </summary>
        ''' <remarks></remarks>
        Dim wMilliseconds As Short
    End Structure

    ''' <summary>
    ''' Specifies settings for a time zone.
    ''' </summary>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/ms725481.aspx for full documentation!</remarks>
    Private Structure TIME_ZONE_INFORMATION
        ''' <summary>
        ''' <para>The current bias for local time translation on this computer, in minutes. The bias
        ''' is the difference, in minutes, between Coordinated Universal Time (UTC) and local time.
        ''' All translations between UTC and local time are based on the following formula:</para>
        ''' <para>UTC = local time + bias</para>
        ''' <para>This member is required.</para>
        ''' </summary>
        ''' <remarks></remarks>
        Dim bias As Integer ' current offset to GMT
        ''' <summary>
        ''' <para>A description for standard time. For example, "EST" could indicate Eastern Standard
        ''' Time. The string will be returned unchanged by the <a href="http://msdn.microsoft.com/en-us/library/ms724421(VS.85).aspx">GetTimeZoneInformation</a> function. This
        ''' string can be empty.</para>
        ''' </summary>
        ''' <remarks></remarks>
        <VBFixedArray(64)> Dim StandardName() As Byte ' unicode string
        ''' <summary>
        ''' <para>A SYSTEMTIME structure that contains a date and local time when the transition from
        ''' daylight saving time to standard time occurs on this operating system. If the time zone does
        ''' not support daylight saving time or if the caller needs to disable daylight saving time, the
        ''' wMonth member in the SYSTEMTIME structure must be zero. If this date is specified, the
        ''' DaylightDate member of this structure must also be specified. Otherwise, the system assumes
        ''' the time zone data is invalid and no changes will be applied.</para>
        ''' </summary>
        ''' <remarks>See http://msdn.microsoft.com/en-us/library/ms725481.aspx for full Documentation!</remarks>
        Dim StandardDate As SYSTEMTIME
        ''' <summary>
        ''' <para>The bias value to be used during local time translations that occur during standard
        ''' time. This member is ignored if a value for the StandardDate member is not supplied.</para>
        ''' <para>This value is added to the value of the Bias member to form the bias used during
        ''' standard time. In most time zones, the value of this member is zero.</para>
        ''' </summary>
        ''' <remarks></remarks>
        Dim StandardBias As Integer
        ''' <summary>
        ''' <para>A description for daylight saving time. For example, "PDT" could indicate Pacific
        ''' Daylight Time. The string will be returned unchanged by the GetTimeZoneInformation function.
        ''' This string can be empty.</para>
        ''' </summary>
        ''' <remarks></remarks>
        <VBFixedArray(64)> Dim DaylightName() As Byte
        ''' <summary>
        ''' <para>A SYSTEMTIME structure that contains a date and local time when the transition from
        ''' standard time to daylight saving time occurs on this operating system. If the time zone does
        ''' not support daylight saving time or if the caller needs to disable daylight saving time, the
        ''' wMonth member in the SYSTEMTIME structure must be zero. If this date is specified, the
        ''' StandardDate member in this structure must also be specified. Otherwise, the system assumes
        ''' the time zone data is invalid and no changes will be applied.</para>
        ''' <para>To select the correct day in the month, set the wYear member to zero, the wHour and
        ''' wMinute members to the transition time, the wDayOfWeek member to the appropriate weekday, and
        ''' the wDay member to indicate the occurrence of the day of the week within the month (1 to 5,
        ''' where 5 indicates the final occurrence during the month if that day of the week does not occur
        ''' 5 times).</para>
        ''' <para>If the wYear member is not zero, the transition date is absolute; it will only occur one
        ''' time. Otherwise, it is a relative date that occurs yearly.</para>
        ''' </summary>
        ''' <remarks></remarks>
        Dim DaylightDate As SYSTEMTIME
        ''' <summary>
        ''' <para>The bias value to be used during local time translations that occur during daylight
        ''' saving time. This member is ignored if a value for the DaylightDate member is not supplied.</para>
        ''' <para>This value is added to the value of the Bias member to form the bias used during
        ''' daylight saving time. In most time zones, the value of this member is –60.</para>
        ''' </summary>
        ''' <remarks></remarks>
        Dim DaylightBias As Integer

        ' For more information about the Dynamic DST key, see <a href="http://msdn.microsoft.com/en-us/library/ms724253(VS.85).aspx">DYNAMIC_TIME_ZONE_INFORMATION</a>.

        Public Sub Initialize()
            ReDim StandardName(64)
            ReDim DaylightName(64)
        End Sub
    End Structure

    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum TimeZoneReturn
        TimeZoneCode = 0
        TimeZoneName = 1
        UTC_BaseOffset = 2
        UTC_Offset = 3
        DST_Active = 4
        DST_Offset = 5
    End Enum

#Region "For Time Zone Retrieval"

    ''' <summary>
    ''' The system cannot determine the current time zone. If daylight saving time is not used in the
    ''' current time zone, this value is returned because there are no transition dates.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const TIME_ZONE_ID_UNKNOWN As Short = 0
    ''' <summary>
    ''' The system is operating in the range covered by the StandardDate member of the
    ''' TIME_ZONE_INFORMATION structure.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const TIME_ZONE_ID_STANDARD As Short = 1
    ''' <summary>
    ''' If the function fails, the return value is TIME_ZONE_ID_UNKNOWN. To get extended error
    ''' information, call GetLastError.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const TIME_ZONE_ID_INVALID As Integer = &HFFFFFFFF
    ''' <summary>
    ''' The system is operating in the range covered by the DaylightDate member of the
    ''' TIME_ZONE_INFORMATION structure.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const TIME_ZONE_ID_DAYLIGHT As Short = 2

    ''' <summary>
    ''' Retrieves the current system date and time. The system time is expressed in Coordinated
    ''' Universal Time (UTC).
    ''' </summary>
    ''' <param name="lpSystemTime">A pointer to a SYSTEMTIME structure to receive the current system date and time. The lpSystemTime parameter must not be NULL. Using NULL will result in an access violation.</param>
    ''' <remarks></remarks>
    Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)
    ''' <summary>
    ''' Retrieves the current time zone settings. These settings control the translations between
    ''' Coordinated Universal Time (UTC) and local time.
    ''' </summary>
    ''' <param name="lpTimeZoneInformation">A pointer to a <see cref="TIME_ZONE_INFORMATION"/> structure to receive the current settings.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Declare Function GetTimeZoneInformation Lib "kernel32" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Integer

#End Region

#Region "To Report API errors"

    ''' <summary>
    ''' The function allocates a buffer large enough to hold the formatted message, and places a pointer
    ''' to the allocated buffer at the address specified by lpBuffer. The lpBuffer parameter is a pointer
    ''' to an LPTSTR; you must cast the pointer to an LPTSTR (for example, (LPTSTR)&amp;lpBuffer). The nSize
    ''' parameter specifies the minimum number of TCHARs to allocate for an output message buffer. The
    ''' caller should use the LocalFree function to free the buffer when it is no longer needed.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Short = &H100S
    ''' <summary>
    ''' <para>The Arguments parameter is not a va_list structure, but is a pointer to an array of values
    ''' that represent the arguments.</para>
    ''' <para>This flag cannot be used with 64-bit integer values. If you are using a 64-bit integer,
    ''' you must use the va_list structure.</para>
    ''' </summary>
    ''' <remarks></remarks>
    Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Short = &H2000S
    ''' <summary>
    ''' <para>The lpSource parameter is a module handle containing the message-table resource(s) to search. If
    ''' this lpSource handle is NULL, the current process's application image file will be searched. This
    ''' flag cannot be used with FORMAT_MESSAGE_FROM_STRING.</para>
    ''' <para>If the module has no message table resource, the function fails with
    ''' ERROR_RESOURCE_TYPE_NOT_FOUND.</para>
    ''' </summary>
    ''' <remarks></remarks>
    Private Const FORMAT_MESSAGE_FROM_HMODULE As Short = &H800S
    ''' <summary>
    ''' The lpSource parameter is a pointer to a null-terminated string that contains a message 
    ''' definition. The message definition may contain insert sequences, just as the message text in a
    ''' message table resource may. This flag cannot be used with FORMAT_MESSAGE_FROM_HMODULE or
    ''' FORMAT_MESSAGE_FROM_SYSTEM.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const FORMAT_MESSAGE_FROM_STRING As Short = &H400S
    ''' <summary>
    ''' <para>The function should search the system message-table resource(s) for the requested message.
    ''' If this flag is specified with FORMAT_MESSAGE_FROM_HMODULE, the function searches the system
    ''' message table if the message is not found in the module specified by lpSource. This flag cannot
    ''' be used with FORMAT_MESSAGE_FROM_STRING.</para>
    ''' <para>If this flag is specified, an application can pass the result of the GetLastError function
    ''' to retrieve the message text for a system-defined error.</para>
    ''' </summary>
    ''' <remarks></remarks>
    Private Const FORMAT_MESSAGE_FROM_SYSTEM As Short = &H1000S
    ''' <summary>
    ''' Insert sequences in the message definition are to be ignored and passed through to the output
    ''' buffer unchanged. This flag is useful for fetching a message for later formatting. If this flag
    ''' is set, the Arguments parameter is ignored.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Short = &H200S
    ''' <summary>
    ''' The function ignores regular line breaks in the message definition text. The function stores
    ''' hard-coded line breaks in the message definition text into the output buffer. The function
    ''' generates no new line breaks.
    ''' </summary>
    ''' <remarks></remarks>
    Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Short = &HFFS

    ''' <summary>
    ''' Formats a message string. The function requires a message definition as input. The message
    ''' definition can come from a buffer passed into the function. It can come from a message table
    ''' resource in an already-loaded module. Or the caller can ask the function to search the system's
    ''' message table resource(s) for the message definition. The function finds the message definition
    ''' in a message table resource based on a message identifier and a language identifier. The function
    ''' copies the formatted message text to an output buffer, processing any embedded insert sequences
    ''' if requested.
    ''' </summary>
    ''' <param name="dwFlags">[in] The formatting options, and how to interpret the lpSource parameter.
    ''' The low-order byte of dwFlags specifies how the function handles line breaks in the output buffer.
    ''' The low-order byte can also specify the maximum width of a formatted output line.</param>
    ''' <param name="lpSource">[in, optional] The location of the message definition. The type of this
    ''' parameter depends upon the settings in the dwFlags parameter. </param>
    ''' <param name="dwMessageId">[in] The message identifier for the requested message. This parameter
    ''' is ignored if dwFlags includes FORMAT_MESSAGE_FROM_STRING.</param>
    ''' <param name="dwLanguageId">[in] The language identifier for the requested message. This parameter
    ''' is ignored if dwFlags includes FORMAT_MESSAGE_FROM_STRING.</param>
    ''' <param name="lpBuffer">[out]<para>A pointer to a buffer that receives the null-terminated string
    ''' that specifies the formatted message. If dwFlags includes FORMAT_MESSAGE_ALLOCATE_BUFFER, the
    ''' function allocates a buffer using the LocalAlloc function, and places the pointer to the buffer
    ''' at the address specified in lpBuffer.</para>
    ''' <para>This buffer cannot be larger than 64K bytes.</para></param>
    ''' <param name="nSize">[in] <para>If the FORMAT_MESSAGE_ALLOCATE_BUFFER flag is not set, this
    ''' parameter specifies the size of the output buffer, in TCHARs. If FORMAT_MESSAGE_ALLOCATE_BUFFER
    ''' is set, this parameter specifies the minimum number of TCHARs to allocate for an output buffer.</para>
    ''' <para>The output buffer cannot be larger than 64K bytes.</para>
    ''' </param>
    ''' <param name="Arguments">[in, optional] An array of values that are used as insert values in the
    ''' formatted message. A %1 in the format string indicates the first value in the Arguments array;
    ''' a %2 indicates the second argument; and so on.</param>
    ''' <returns>If the function succeeds, the return value is the number of TCHARs stored in the output
    ''' buffer, excluding the terminating null character. If the function fails, the return value is zero.
    ''' To get extended error information, call GetLastError.</returns>
    ''' <remarks></remarks>
    Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, ByRef lpSource As Long, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As String, ByVal nSize As Integer, ByRef Arguments As Integer) As Integer
    ''' <summary>
    ''' Retrieves the path of the Windows directory. The Windows directory contains such files as
    ''' applications, initialization files, and help files.
    ''' </summary>
    ''' <param name="lpBuffer">[out] A pointer to a buffer that receives the path. This path does not end
    ''' with a backslash unless the Windows directory is the root directory. For example, if the Windows
    ''' directory is named Windows on drive C, the path of the Windows directory retrieved by this
    ''' function is C:\Windows. If the system was installed in the root directory of drive C, the path
    ''' retrieved is C:\.</param>
    ''' <param name="nSize">[in] The maximum size of the buffer specified by the lpBuffer parameter, in
    ''' TCHARs. This value should be set to MAX_PATH.</param>
    ''' <returns><para>If the function succeeds, the return value is the length of the string copied to
    ''' the buffer, in TCHARs, not including the terminating null character.</para>
    ''' <para>If the length is greater than the size of the buffer, the return value is the size of the
    ''' buffer required to hold the path.</para>
    ''' </returns>
    ''' <remarks>The Windows directory is the directory where an application should store initialization
    ''' and help files. If the user is running a shared version of the system, the Windows directory is
    ''' guaranteed to be private for each user.</remarks>
    Public Declare Function GeneralWinDirApi Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    ''' <summary>
    ''' Retrieves the path of the system directory. The system directory contains system files such as
    ''' dynamic-link libraries and drivers.
    ''' </summary>
    ''' <param name="lpBuffer">[out] A pointer to the buffer to receive the path. This path does not end
    ''' with a backslash unless the system directory is the root directory. For example, if the system
    ''' directory is named Windows\System on drive C, the path of the system directory retrieved by this
    ''' function is C:\Windows\System.</param>
    ''' <param name="nSize">[in] The maximum size of the buffer, in TCHARs.</param>
    ''' <returns>If the function succeeds, the return value is the length, in TCHARs, of the string copied
    ''' to the buffer, not including the terminating null character. If the length is greater than the
    ''' size of the buffer, the return value is the size of the buffer required to hold the path,
    ''' including the terminating null character.</returns>
    ''' <remarks></remarks>
    Public Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

#End Region

    ' The following constants and declares are used to Get/Set Locale Date format
    Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
    Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String) As Boolean
    Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Short
    Const LOCALE_SSHORTDATE As Short = &H1FS
    Public regionalSymbol As String

    ' Internet connection constants
    ''' <summary>
    ''' Used by Function IsWebConnected
    ''' </summary>
    ''' <remarks></remarks>
    Dim ConnectionQualityString As String = "Off"

    ''' <summary>
    ''' Retrieves the connected state of the local system.
    ''' </summary>
    ''' <param name="lpSFlags">[out] Pointer to a variable that receives the connection description. This
    ''' parameter may return a valid flag even when the function returns FALSE.</param>
    ''' <param name="dwReserved">[in] This parameter is reserved and must be 0.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Declare Function InternetGetConnectedState _
            Lib "wininet.dll" (ByRef lpSFlags As Int32, _
            ByVal dwReserved As Int32) As Boolean

    ''' <summary>
    ''' Internet connection states!
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum InetConnState
        ''' <summary>
        ''' Local system uses a modem to connect to the Internet.
        ''' </summary>
        ''' <remarks></remarks>
        modem = &H1
        ''' <summary>
        ''' Local system uses a local area network to connect to the Internet.
        ''' </summary>
        ''' <remarks></remarks>
        lan = &H2
        ''' <summary>
        ''' Local system uses a proxy server to connect to the Internet.
        ''' </summary>
        ''' <remarks></remarks>
        proxy = &H4
        ''' <summary>
        ''' ?Undocumented!
        ''' </summary>
        ''' <remarks></remarks>
        ras = &H10
        ''' <summary>
        ''' Local system is in offline mode.
        ''' </summary>
        ''' <remarks></remarks>
        offline = &H20
        ''' <summary>
        ''' Local system has a valid connection to the Internet, but it might or might not be currently
        ''' connected.
        ''' </summary>
        ''' <remarks></remarks>
        configured = &H40
    End Enum

    ''' <summary>
    ''' Trims Null characters from the string.
    ''' </summary>
    ''' <param name="startstr">String - String to be trimmed</param>
    ''' <returns>String - Trimmed string</returns>
    ''' <remarks></remarks>
    Public Function TrimNulls(ByRef startstr As String) As String
        Dim pos As Short
        pos = InStr(startstr, Chr(0))
        If pos Then
            TrimNulls = Trim(Left(startstr, pos - 1))
        Else
            TrimNulls = Trim(startstr)
        End If
    End Function

    ''' <summary>
    ''' Reads a binary file into the sData buffer. Returns the number of bytes read.
    ''' </summary>
    ''' <param name="sPath">String - Path to the file to be read</param>
    ''' <param name="sData">String - Output parameter contains the data that has been read</param>
    ''' <returns>Long - Number of bytes read, 0 if no file was read</returns>
    ''' <remarks></remarks>
    Public Function ReadFile(ByVal sPath As String, ByRef sData As String) As Integer

        Dim c As New CRC32
        Dim crc As Integer = 0

        ' CRC32 Hash:
        Dim f As FileStream = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        crc = c.GetCrc32(f)
        f.Close()

        ' File size:
        'f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        'txtSize.Text = String.Format("{0}", f.Length)
        'f.Close()
        'txtCrc32.Text = String.Format("{0:X8}", crc)
        'txtTime.Text = String.Format("{0}", h.ElapsedTime)

        ' Run MD5 Hash
        f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider
        md5.ComputeHash(f)
        f.Close()

        Dim hash As Byte() = md5.Hash
        Dim buff As StringBuilder = New StringBuilder
        Dim hashByte As Byte
        For Each hashByte In hash
            buff.Append(String.Format("{0:X1}", hashByte))
        Next
        sData = buff.ToString() 'MD5 String

        ' Run SHA-1 Hash
        'f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        'Dim sha1 As SHA1CryptoServiceProvider = New SHA1CryptoServiceProvider
        'sha1.ComputeHash(f)
        'f.Close()
        'hash = SHA1.Hash
        'buff = New StringBuilder
        'For Each hashByte In hash
        '    buff.Append(String.Format("{0:X1}", hashByte))
        'Next
        'txtSHA1.Text = buff.ToString()

        ReadFile = Len(sData)
        Exit Function
Hell:
        Set_locale(regionalSymbol)
        Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End Function

    ''' <summary>
    ''' [INTERNAL] Call-back routine used by ALCrypto3.dll during key generation process.
    ''' </summary>
    ''' <param name="param">Long - TBD</param>
    ''' <param name="action">Long - Action being performed</param>
    ''' <param name="phase">Long - Current phase</param>
    ''' <param name="iprogress">Long - Percent complete</param>
    ''' <remarks></remarks>
    Public Sub CryptoProgressUpdate(ByVal param As Integer, ByVal action As Integer, ByVal phase As Integer, ByVal iprogress As Integer)
        System.Diagnostics.Debug.WriteLine("Progress Update received " & param & ", action: " & action & ", iprogress: " & iprogress)
    End Sub

    ''' <summary>
    ''' This is a dummy sub. Used to circumvent the End statement restriction in COM DLLs.
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub EndSub()
        'Dummy sub
    End Sub

    ''' <summary>
    ''' Computes an MD5 hash of the specified file.
    ''' </summary>
    ''' <param name="strPath">String - File path</param>
    ''' <returns>String - MD5 Hash Value</returns>
    ''' <remarks></remarks>
    Public Function MD5HashFile(ByVal strPath As String) As String
        System.Diagnostics.Debug.WriteLine("Hashing file " & strPath)
        System.Diagnostics.Debug.WriteLine("File Date: " & FileDateTime(strPath))
        ' read and hash the content
        Dim sData As String = String.Empty
        Dim nFileLen As Integer
        nFileLen = ReadFile(strPath, sData)
        ' use the .NET's native MD5 functions instead of our own MD5 hashing routine
        ' and instead of ALCrypto's md5_hash() function.
        MD5HashFile = LCase(sData)    '<--- ReadFile procedure already computes the MD5.Hash
    End Function

    ''' <summary>
    ''' Checks if a file exists in the system.
    ''' </summary>
    ''' <param name="strFile">String - File path and name</param>
    ''' <returns>Boolean - True if file exists, False if it doesn't</returns>
    ''' <remarks></remarks>
    Public Function FileExists(ByVal strFile As String) As Boolean
        FileExists = False
        If File.Exists(strFile) = True Then
            FileExists = True
        End If
    End Function

    ' Name: Function LocalTimeZone
    ' Input:
    '   ByVal returnType As TimeZoneReturn - Type of time zone information being requested
    '       UTC_BaseOffset = UTC offset, not including DST <br>
    '       UTC_Offset = UTC offset, including DST if active <br>
    '       DST_Active = True if DST is currently active, otherwise false <br>
    '       DST_Offset = Offset value for DST (generally -60, if in US)
    ''' <summary>
    ''' Retrieves the local time zone.
    ''' </summary>
    ''' <param name="returnType">TimeZoneReturn - Type of time zone information being requested</param>
    ''' <returns>Variant - Return type varies depending on returnValue parameter.</returns>
    ''' <remarks></remarks>
    Public Function LocalTimeZone(ByVal returnType As TimeZoneReturn) As Object
        Dim x As Integer
        Dim tzi As TIME_ZONE_INFORMATION = Nothing
        Dim strName As String
        Dim bDST As Boolean
        Dim rc As Integer

        LocalTimeZone = Nothing

        rc = GetTimeZoneInformation(tzi)
        Select Case rc
            ' if not daylight assume standard
            Case TIME_ZONE_ID_DAYLIGHT
                strName = System.Text.UnicodeEncoding.Unicode.GetString(tzi.DaylightName) ' convert to string
                bDST = True
            Case Else
                strName = System.Text.UnicodeEncoding.Unicode.GetString(tzi.StandardName)
        End Select

        ' name terminates with null
        x = InStr(strName, vbNullChar)
        If x > 0 Then strName = Left(strName, x - 1)

        If returnType = TimeZoneReturn.DST_Active Then
            LocalTimeZone = bDST
        End If

        If returnType = TimeZoneReturn.TimeZoneName Then
            LocalTimeZone = strName
        End If

        If returnType = TimeZoneReturn.TimeZoneCode Then
            LocalTimeZone = Left(strName, 1)
            x = InStr(1, strName, " ")
            Do While x > 0
                LocalTimeZone = LocalTimeZone & Mid(strName, x + 1, 1)
                x = InStr(x + 1, strName, " ")
            Loop
            LocalTimeZone = Trim(LocalTimeZone)
        End If

        If returnType = TimeZoneReturn.UTC_BaseOffset Then
            LocalTimeZone = tzi.bias
        End If

        If returnType = TimeZoneReturn.DST_Offset Then
            LocalTimeZone = tzi.DaylightBias
        End If

        If returnType = TimeZoneReturn.UTC_Offset Then
            If tzi.DaylightBias = -60 Then
                LocalTimeZone = tzi.bias
            Else
                LocalTimeZone = -tzi.bias
            End If
            ' Account for Daylight Savings Time
            If bDST Then LocalTimeZone = LocalTimeZone - 60
        End If
    End Function

    ''' <summary>
    ''' Performs RSA signing of <code>strData</code> using the specified key.
    ''' </summary>
    ''' <param name="strPub">RSA Public key blob</param>
    ''' <param name="strPriv">RSA Private key blob</param>
    ''' <param name="strdata">String - Data to be signed</param>
    ''' <returns>String - Signature string</returns>
    ''' <remarks>05.13.05    - alkan  - Removed the modActiveLock references</remarks>
    Public Function RSASign(ByVal strPub As String, ByVal strPriv As String, ByVal strdata As String) As String
        Dim KEY As RSAKey = Nothing
        ' create the key from the key blobs
        If rsa_createkey(strPub, Len(strPub), strPriv, Len(strPriv), KEY) = RETVAL_ON_ERROR Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If

        ' sign the data using the created key
        Dim sLen As Integer
        If rsa_sign(KEY, strdata, Len(strdata), vbNullString, sLen) = RETVAL_ON_ERROR Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        Dim strSig As String : strSig = New String(Chr(0), sLen)
        If rsa_sign(KEY, strdata, Len(strdata), strSig, sLen) = RETVAL_ON_ERROR Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        ' throw away the key
        If rsa_freekey(KEY) = RETVAL_ON_ERROR Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        RSASign = strSig
    End Function

    '===============================================================================
    ''' <summary>
    ''' Verifies an RSA signature.
    ''' </summary>
    ''' <param name="strPub">String - Public key blob</param>
    ''' <param name="strdata">String - Data to be signed</param>
    ''' <param name="strSig">String - Private key blob</param>
    ''' <returns>Long - Zero if verification is successful, non-zero otherwise.</returns>
    ''' <remarks></remarks>
    Public Function RSAVerify(ByVal strPub As String, ByVal strdata As String, ByVal strSig As String) As Integer
        Dim KEY As RSAKey = Nothing
        Dim rc As Integer
        ' create the key from the public key blob
        If rsa_createkey(strPub, Len(strPub), vbNullString, 0, KEY) = RETVAL_ON_ERROR Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        ' validate the key
        rc = rsa_verifysig(KEY, strSig, Len(strSig), strdata, Len(strdata))
        If rc = RETVAL_ON_ERROR Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        ' de-allocate memory used by the key
        If rsa_freekey(KEY) = RETVAL_ON_ERROR Then
            Set_locale(regionalSymbol)
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
        End If
        RSAVerify = rc
    End Function

    ''' <summary>
    ''' Retrieves the error text for the specified Windows error code
    ''' </summary>
    ''' <param name="lLastDLLError">Long - Last DLL error as an input</param>
    ''' <returns>String - Error message string</returns>
    ''' <remarks></remarks>
    Public Function WinError(ByVal lLastDLLError As Integer) As String
        Dim sBuff As String
        Dim lCount As Integer

        WinError = String.Empty

        ' Return the error message associated with LastDLLError:
        sBuff = New String(Chr(0), 256)
        lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0, sBuff, Len(sBuff), 0)
        If lCount Then
            WinError = Left(sBuff, lCount)
        End If
    End Function

    ''' <summary>
    ''' Gets the windows directory
    ''' </summary>
    ''' <returns>String - Windows directory path</returns>
    ''' <remarks></remarks>
    Public Function WinDir() As String
        Dim WinSysPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.System)
        WinDir = WinSysPath.Substring(0, WinSysPath.LastIndexOf("\"))
    End Function

    ''' <summary>
    ''' Gets the Windows system directory
    ''' </summary>
    ''' <returns>String - Windows system directory path</returns>
    ''' <remarks></remarks>
    Public Function WinSysDir() As String
        WinSysDir = System.Environment.GetFolderPath(Environment.SpecialFolder.System)
        ' or could use WinSysDir = System.Environment.SystemDirectory
    End Function

    ''' <summary>
    ''' Checks if a Folder Exists
    ''' </summary>
    ''' <param name="sFolder">String -  Name of the folder in question</param>
    ''' <returns>Boolean - Returns true if the Folder Exists</returns>
    ''' <remarks></remarks>
    Public Function FolderExists(ByVal sFolder As String) As Boolean
        FolderExists = Directory.Exists(sFolder)
    End Function

    ''' <summary>
    ''' Packs two 8-bit integers into a 16-bit integer.
    ''' </summary>
    ''' <param name="LoByte">Byte - A byte that is to become the Low byte.</param>
    ''' <param name="HiByte">Byte - A byte that is to become the High byte.</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function MakeWord(ByVal LoByte As Byte, ByVal HiByte As Byte) As Short
        If (HiByte And &H80S) <> 0 Then
            MakeWord = ((HiByte * 256) + LoByte) Or &HFFFF0000
        Else
            MakeWord = (HiByte * 256) + LoByte
        End If
    End Function

    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <param name="w"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function HiByte(ByVal w As Short) As Byte
        HiByte = (w And &HFF00) \ 256
    End Function

    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <param name="w"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoByte(ByVal w As Short) As Byte
        LoByte = w And &HFFS
    End Function


    ''' <summary>
    ''' Converts a local date-time into UTC/GMT date-time
    ''' </summary>
    ''' <param name="dt">Date - Date-Time input to be converted into UTC Date-Time</param>
    ''' <returns>Date - UTC Date-Time</returns>
    ''' <remarks></remarks>
    Public Function UTC(ByVal dt As Date) As Date
        '  Returns current UTC date-time.
        UTC = dt.AddMinutes(LocalTimeZone(TimeZoneReturn.UTC_Offset))
    End Function

    ''' <summary>
    ''' Retrieves the regional setting
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub Get_locale()
        Dim Symbol As String
        Dim iRet1 As Integer
        Dim iRet2 As Integer
        Dim lpLCDataVar As String = String.Empty
        Dim Pos As Short
        Dim Locale As Integer
        Locale = GetUserDefaultLCID()
        iRet1 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, lpLCDataVar, 0)
        Symbol = New String(Chr(0), iRet1)
        iRet2 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol, iRet1)
        Pos = InStr(Symbol, Chr(0))
        If Pos > 0 Then
            Symbol = Left(Symbol, Pos - 1)
            If Symbol <> "yyyy/MM/dd" Then regionalSymbol = Symbol
        End If
    End Sub

    ''' <summary>
    ''' Changes the regional setting.
    ''' </summary>
    ''' <param name="localSymbol"></param>
    ''' <remarks></remarks>
    Public Sub Set_locale(Optional ByVal localSymbol As String = "") 'Change the regional setting
        Dim Symbol As String
        Dim iRet As Integer
        Dim Locale As Integer
        Locale = GetUserDefaultLCID() 'Get user Locale ID
        If localSymbol = "" Then
            Symbol = "yyyy/MM/dd" 'New character for the locale
        Else
            Symbol = localSymbol
        End If

        iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol)
    End Sub

    ''' <summary>
    ''' Gets special folders...
    ''' </summary>
    ''' <param name="CSIDL">See this functions Definition for detailed info...</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ActivelockGetSpecialFolder(ByVal CSIDL As Long) As String
        'value  ID                      Result
        '2      CSIDL_PROGRAMS          C:\Documents and Settings\<USERNAME>\Start Menu\Programs
        '5      CSIDL_PERSONAL          C:\Documents and Settings\<USERNAME>\My Documents
        '6      CSIDL_FAVORITES         C:\Documents and Settings\<USERNAME>\Favorites
        '7      CSIDL_STARTUP           C:\Documents and Settings\<USERNAME>\Start Menu\Programs\Startup
        '8      CSIDL_RECENT            C:\Documents and Settings\<USERNAME>\Recent
        '9      CSIDL_SENDTO            C:\Documents and Settings\<USERNAME>\SendTo
        '11     CSIDL_STARTMENU         C:\Documents and Settings\<USERNAME>\Start Menu
        '13     CSIDL_MYMUSIC           C:\Documents and Settings\<USERNAME>\My Documents\My Music
        '14     CSIDL_MYVIDEO           C:\Documents and Settings\<USERNAME>\My Documents\My Video
        '16     CSIDL_DESKTOPDIRECTORY  C:\Documents and Settings\<USERNAME>\Desktop
        '19     CSIDL_NETHOOD           C:\Documents and Settings\<USERNAME>\NetHood
        '20     Virtual - don't use     CSIDL_FONTS             C:\Windows\Fonts
        '21     CSIDL_TEMPLATES         C:\Documents and Settings\<USERNAME>\Templates
        '22     CSIDL_COMMON_STARTMENU  C:\Documents and Settings\All Users\Start Menu
        '23     CSIDL_COMMON_PROGRAMS   C:\Documents and Settings\All Users\Start Menu\Programs
        '24     CSIDL_COMMON_STARTUP    C:\Documents and Settings\All Users\Start Menu\Programs\Startup
        '25     CSIDL_COMMON_DESKTOPDIRECTORY     C:\Documents and Settings\All Users\Desktop
        '26     CSIDL_APPDATA           C:\Documents and Settings\<USERNAME>\Application Data
        '27     CSIDL_PRINTHOOD         C:\Documents and Settings\<USERNAME>\PrintHood
        '28     CSIDL_LOCAL_APPDATA     C:\Documents and Settings\<USERNAME>\Local Settings\Application Data
        '31     CSIDL_COMMON_FAVORITES  C:\Documents and Settings\All Users\Favorites
        '32     CSIDL_INTERNET_CACHE    C:\Documents and Settings\<USERNAME>\Local Settings\Temporary Internet Files
        '33     CSIDL_COOKIES           C:\Documents and Settings\<USERNAME>\Cookies
        '34     CSIDL_HISTORY           C:\Documents and Settings\<USERNAME>\Local Settings\History
        '35     CSIDL_COMMON_APPDATA    C:\Documents and Settings\All Users\Application Data
        '36     CSIDL_WINDOWS           C:\Windows
        '37     CSIDL_SYSTEM            C:\Windows\system32
        '38     CSIDL_PROGRAM_FILES     C:\Program Files
        '39     CSIDL_MYPICTURES        C:\Documents and Settings\<USERNAME>\My Documents\My Pictures
        '40     CSIDL_PROFILE           C:\Documents and Settings\<USERNAME>
        '41     CSIDL_SYSTEMX86         C:\Windows\system32  'x86 system directory on RISC
        '43     CSIDL_PROGRAM_FILES_COMMON     C:\Program Files\Common Files
        '45     CSIDL_COMMON_TEMPLATES  C:\Documents and Settings\All Users\Templates
        '46     CSIDL_COMMON_DOCUMENTS  C:\Documents and Settings\All Users\Documents
        '47     CSIDL_COMMON_ADMINTOOLS C:\Documents and Settings\All Users\Start Menu\Programs\Administrative Tools
        '48     CSIDL_ADMINTOOLS        C:\Documents and Settings\<USERNAME>\Start Menu\Programs\Administrative Tools
        '53     CSIDL_COMMON_MUSIC      C:\Documents and Settings\All Users\My Documents\My Music
        '54     CSIDL_COMMON_PICTURES   C:\Documents and Settings\All Users\My Documents\My Pictures
        '55     CSIDL_COMMON_VIDEO      C:\Documents and Settings\All Users\My Documents\My Video

        ' CLSIDs for XP and older
        'CSIDL_DESKTOP = &H0
        'CSIDL_INTERNET = &H1
        'CSIDL_PROGRAMS = &H2
        'CSIDL_CONTROLS = &H3
        'CSIDL_PRINTERS = &H4
        'CSIDL_PERSONAL = &H5
        'CSIDL_FAVORITES = &H6
        'CSIDL_STARTUP = &H7
        'CSIDL_RECENT = &H8
        'CSIDL_SENDTO = &H9
        'CSIDL_BITBUCKET = &HA
        'CSIDL_STARTMENU = &HB
        'CSIDL_MYDOCUMENTS = &HC
        'CSIDL_MYMUSIC = &HD
        'CSIDL_MYVIDEO = &HE
        'CSIDL_DESKTOPDIRECTORY = &H10
        'CSIDL_DRIVES = &H11
        'CSIDL_NETWORK = &H12
        'CSIDL_NETHOOD = &H13
        'CSIDL_FONTS = &H14
        'CSIDL_TEMPLATES = &H15
        'CSIDL_COMMON_STARTMENU = &H16
        'CSIDL_COMMON_PROGRAMS = &H17
        'CSIDL_COMMON_STARTUP = &H18
        'CSIDL_COMMON_DESKTOPDIRECTORY = &H19
        'CSIDL_APPDATA = &H1A
        'CSIDL_PRINTHOOD = &H1B
        'CSIDL_LOCAL_APPDATA = &H1C
        'CSIDL_ALTSTARTUP = &H1D
        'CSIDL_COMMON_ALTSTARTUP = &H1E
        'CSIDL_COMMON_FAVORITES = &H1F
        'CSIDL_INTERNET_CACHE = &H20
        'CSIDL_COOKIES = &H21
        'CSIDL_HISTORY = &H22
        'CSIDL_COMMON_APPDATA = &H23
        'CSIDL_WINDOWS = &H24
        'CSIDL_SYSTEM = &H25
        'CSIDL_PROGRAM_FILES = &H26
        'CSIDL_MYPICTURES = &H27
        'CSIDL_PROFILE = &H28
        'CSIDL_SYSTEMX86 = &H29
        'CSIDL_PROGRAM_FILESX86 = &H2A
        'CSIDL_PROGRAM_FILES_COMMON = &H2B
        'CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
        'CSIDL_COMMON_TEMPLATES = &H2D
        'CSIDL_COMMON_DOCUMENTS = &H2E
        'CSIDL_COMMON_ADMINTOOLS = &H2F
        'CSIDL_ADMINTOOLS = &H30
        'CSIDL_CONNECTIONS = &H31
        'CSIDL_COMMON_MUSIC = &H35
        'CSIDL_COMMON_PICTURES = &H36
        'CSIDL_COMMON_VIDEO = &H37
        'CSIDL_RESOURCES = &H38
        'CSIDL_RESOURCES_LOCALIZED = &H39
        'CSIDL_COMMON_OEM_LINKS = &H3A
        'CSIDL_CDBURN_AREA = &H3B
        'CSIDL_COMPUTERSNEARME = &H3D
        'CSIDL_FLAG_PER_USER_INIT = &H800
        'CSIDL_FLAG_NO_ALIAS = &H1000
        'CSIDL_FLAG_DONT_VERIFY = &H4000
        'CSIDL_FLAG_CREATE = &H8000
        'CSIDL_FLAG_MASK = &HFF00

        'KNOWNFOLDER IDs for Vista
        'Select Case CSIDL

        '    Case CSIDL_NETWORK, CSIDL_COMPUTERSNEARME  ' VIRTUAL
        '        FOLDERID_NetworkFolder = "{D20BEEC4-5CA8-4905-AE3B-BF251EA09B53}"
        '        sRfid = FOLDERID_NetworkFolder
        '    Case CSIDL_DRIVES  ' VIRTUAL
        '        FOLDERID_ComputerFolder = "{0AC0837C-BBF8-452A-850D-79D08E667CA7}"
        '        sRfid = FOLDERID_ComputerFolder
        '    Case CSIDL_INTERNET  ' VIRTUAL
        '        FOLDERID_InternetFolder = "{4D9F7874-4E0C-4904-967B-40B0D20C3E4B}"
        '        sRfid = FOLDERID_InternetFolder
        '    Case CSIDL_CONTROLS  ' VIRTUAL
        '        FOLDERID_ControlPanelFolder = "{82A74AEB-AEB4-465C-A014-D097EE346D63}"
        '        sRfid = FOLDERID_ControlPanelFolder
        '    Case CSIDL_PRINTERS  ' VIRTUAL
        '        FOLDERID_PrintersFolder = "{76FC4E2D-D6AD-4519-A663-37BD56068185}"
        '        sRfid = FOLDERID_PrintersFolder
        '    Case 101  ' VIRTUAL
        '        FOLDERID_SyncManagerFolder = "{43668BF8-C14E-49B2-97C9-747784D784B7}"
        '        sRfid = FOLDERID_SyncManagerFolder
        '    Case 102  ' VIRTUAL
        '        FOLDERID_SyncSetupFolder = "{0F214138-B1D3-4a90-BBA9-27CBC0C5389A}"
        '        sRfid = FOLDERID_SyncSetupFolder
        '    Case 103  ' VIRTUAL
        '        FOLDERID_ConflictFolder = "{4bfefb45-347d-4006-a5be-ac0cb0567192}"
        '        sRfid = FOLDERID_ConflictFolder
        '    Case 104  ' VIRTUAL
        '        FOLDERID_SyncResultsFolder = "{289a9a43-be44-4057-a41b-587a76d7e7f9}"
        '        sRfid = FOLDERID_SyncResultsFolder
        '    Case CSIDL_BITBUCKET  ' VIRTUAL
        '        FOLDERID_RecycleBinFolder = "{B7534046-3ECB-4C18-BE4E-64CD4CB7D6AC}"
        '        sRfid = FOLDERID_RecycleBinFolder
        '    Case CSIDL_CONNECTIONS  ' VIRTUAL
        '        FOLDERID_ConnectionsFolder = "{6F0CD92B-2E97-45D1-88FF-B0D186B8DEDD}"
        '        sRfid = FOLDERID_ConnectionsFolder
        '        ' VISTA - %windir%\Fonts
        '        ' XP - %windir%\Fonts
        '    Case CSIDL_FONTS  ' FIXED
        '        FOLDERID_Fonts = "{FD228CB7-AE11-4AE3-864C-16F3910AB8FE}"
        '        sRfid = FOLDERID_Fonts
        '        ' VISTA - %USERPROFILE%\Desktop
        '        ' XP - %USERPROFILE%\Desktop
        '    Case CSIDL_DESKTOP, CSIDL_DESKTOPDIRECTORY  ' PERUSER
        '        FOLDERID_Desktop = "{B4BFCC3A-DB2C-424C-B029-7FE99A87C641}"
        '        sRfid = FOLDERID_Desktop
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu\Programs\StartUp
        '        ' XP - %USERPROFILE%\Start Menu\Programs\StartUp
        '    Case CSIDL_STARTUP, CSIDL_ALTSTARTUP  ' PERUSER
        '        FOLDERID_Startup = "{B97D20BB-F46A-4C97-BA10-5E3608430854}"
        '        sRfid = FOLDERID_Startup
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu\Programs
        '        ' XP - %USERPROFILE%\Start Menu\Programs
        '    Case CSIDL_PROGRAMS  ' PERUSER
        '        FOLDERID_Programs = "{A77F5D77-2E2B-44C3-A6A2-ABA601054A51}"
        '        sRfid = FOLDERID_Programs
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu
        '        ' XP - %USERPROFILE%\Start Menu
        '    Case CSIDL_STARTMENU  'PERUSER
        '        FOLDERID_StartMenu = "{625B53C3-AB48-4EC1-BA1F-A1EF4146FC19}"
        '        sRfid = FOLDERID_StartMenu
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Recent
        '        ' XP - %USERPROFILE%\Recent
        '    Case CSIDL_RECENT  ' PERUSER
        '        FOLDERID_Recent = "{AE50C081-EBD2-438A-8655-8A092E34987A}"
        '        sRfid = FOLDERID_Recent
        '        ' VISTA - %APPDATA%\Microsoft\Windows\SendTo
        '        ' XP - %USERPROFILE%\SendTo
        '    Case CSIDL_SENDTO  ' PERUSER
        '        FOLDERID_SendTo = "{8983036C-27C0-404B-8F08-102D10DCFD74}"
        '        sRfid = FOLDERID_SendTo
        '        ' VISTA - %USERPROFILE%\Documents
        '        ' XP - %USERPROFILE%\My Documents
        '    Case CSIDL_MYDOCUMENTS, CSIDL_PERSONAL  ' PERUSER
        '        FOLDERID_Documents = "{FDD39AD0-238F-46AF-ADB4-6C85480369C7}"
        '        sRfid = FOLDERID_Documents
        '        ' VISTA - %USERPROFILE%\Documents
        '        ' XP - %USERPROFILE%\My Documents
        '    Case CSIDL_FAVORITES, CSIDL_COMMON_FAVORITES  ' PERUSER
        '        FOLDERID_Favorites = "{1777F761-68AD-4D8A-87BD-30B759FA33DD}"
        '        sRfid = FOLDERID_Favorites
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Network Shortcuts
        '        ' XP - %USERPROFILE%\NetHood
        '    Case CSIDL_NETHOOD  ' PERUSER
        '        FOLDERID_NetHood = "{C5ABBF53-E17F-4121-8900-86626FC2C973}"
        '        sRfid = FOLDERID_NetHood
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Printer Shortcuts
        '        ' XP - %USERPROFILE%\PrintHood
        '    Case CSIDL_PRINTHOOD  ' PERUSER
        '        FOLDERID_PrintHood = "{9274BD8D-CFD1-41C3-B35E-B13F55A758F4}"
        '        sRfid = FOLDERID_PrintHood
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Templates
        '        ' XP - %USERPROFILE%\Templates
        '    Case CSIDL_TEMPLATES  ' PERUSER
        '        FOLDERID_Templates = "{A63293E8-664E-48DB-A079-DF759E0509F7}"
        '        sRfid = FOLDERID_Templates
        '        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\StartUp
        '        ' XP - %ALLUSERSPROFILE%\Start Menu\Programs\StartUp
        '    Case CSIDL_COMMON_STARTUP, CSIDL_COMMON_ALTSTARTUP  ' COMMON
        '        FOLDERID_CommonStartup = "{82A5EA35-D9CD-47C5-9629-E15D2F714E6E}"
        '        sRfid = FOLDERID_CommonStartup
        '        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs
        '        ' XP - %ALLUSERSPROFILE%\Start Menu\Programs
        '    Case CSIDL_COMMON_PROGRAMS  ' COMMON
        '        FOLDERID_CommonPrograms = "{0139D44E-6AFE-49F2-8690-3DAFCAE6FFB8}"
        '        sRfid = FOLDERID_CommonPrograms
        '        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu
        '        ' XP - %ALLUSERSPROFILE%\Start Menu
        '    Case CSIDL_COMMON_STARTMENU  ' COMMON
        '        FOLDERID_CommonStartMenu = "{A4115719-D62E-491D-AA7C-E74B8BE3B067}"
        '        sRfid = FOLDERID_CommonStartMenu
        '        ' VISTA - %PUBLIC%\Desktop
        '        ' XP - %ALLUSERSPROFILE%\Desktop
        '    Case 201  ' COMMON
        '        FOLDERID_PublicDesktop = "{C4AA340D-F20F-4863-AFEF-F87EF2E6BA25}"
        '        sRfid = FOLDERID_PublicDesktop
        '        ' VISTA - %ALLUSERSPROFILE% (%ProgramData%, %SystemDrive%\ProgramData)
        '        ' XP - %ALLUSERSPROFILE%\Application Data
        '    Case CSIDL_COMMON_APPDATA  ' FIXED
        '        FOLDERID_ProgramData = "{62AB5D82-FDC1-4DC3-A9DD-070D1D495D97}"
        '        sRfid = FOLDERID_ProgramData
        '        ' VISTA - %ALLUSERSPROFILE%\Templates
        '        ' XP - %ALLUSERSPROFILE%\Templates
        '    Case CSIDL_COMMON_TEMPLATES  ' COMMON
        '        FOLDERID_CommonTemplates = "{B94237E7-57AC-4347-9151-B08C6C32D1F7}"
        '        sRfid = FOLDERID_CommonTemplates
        '        ' VISTA - %PUBLIC%\Documents
        '        ' XP - %ALLUSERSPROFILE%\Documents
        '    Case CSIDL_COMMON_DOCUMENTS  ' COMMON
        '        FOLDERID_PublicDocuments = "{ED4824AF-DCE4-45A8-81E2-FC7965083634}"
        '        sRfid = FOLDERID_PublicDocuments
        '        ' VISTA - %APPDATA% (%USERPROFILE%\AppData\Roaming)
        '        ' XP - %APPDATA% (%USERPROFILE%\Application Data)
        '    Case CSIDL_APPDATA  ' PERUSER
        '        FOLDERID_RoamingAppData = "{3EB685DB-65F9-4CF6-A03A-E3EF65729F3D}"
        '        sRfid = FOLDERID_RoamingAppData
        '        ' VISTA - %LOCALAPPDATA% (%USERPROFILE%\AppData\Local)
        '        ' XP - %USERPROFILE%\Local Settings\Application Data
        '    Case CSIDL_LOCAL_APPDATA  ' PERUSER
        '        FOLDERID_LocalAppData = "{F1B32785-6FBA-4FCF-9D55-7B8E7F157091}"
        '        sRfid = FOLDERID_LocalAppData
        '        ' VISTA - %USERPROFILE%\AppData\LocalLow
        '        ' XP - NONE
        '    Case 301  ' PERUSER
        '        FOLDERID_LocalAppDataLow = "{A520A1A4-1780-4FF6-BD18-167343C5AF16}"
        '        sRfid = FOLDERID_LocalAppDataLow
        '        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\Temporary Internet Files
        '        ' XP - %USERPROFILE%\Local Settings\Temporary Internet Files
        '    Case CSIDL_INTERNET_CACHE  ' PERUSER
        '        FOLDERID_InternetCache = "{352481E8-33BE-4251-BA85-6007CAEDCF9D}"
        '        sRfid = FOLDERID_InternetCache
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Cookies
        '        ' XP - %USERPROFILE%\Cookies
        '    Case CSIDL_COOKIES  ' PERUSER
        '        FOLDERID_Cookies = "{2B0F765D-C0E9-4171-908E-08A611B84FF6}"
        '        sRfid = FOLDERID_Cookies
        '        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\History
        '        ' XP - %USERPROFILE%\Local Settings\History
        '    Case CSIDL_HISTORY  ' PERUSER
        '        FOLDERID_History = "{D9DC8A3B-B784-432E-A781-5A1130A75963}"
        '        sRfid = FOLDERID_History
        '        ' VISTA - %windir%\system32
        '        ' XP - %windir%\system32
        '    Case CSIDL_SYSTEM  ' FIXED
        '        FOLDERID_System = "{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}"
        '        sRfid = FOLDERID_System
        '        ' VISTA - %windir%\system32
        '        ' XP - %windir%\system32
        '    Case CSIDL_SYSTEMX86  ' FIXED
        '        FOLDERID_SystemX86 = "{D65231B0-B2F1-4857-A4CE-A8E7C6EA7D27}"
        '        sRfid = FOLDERID_SystemX86
        '        ' VISTA - %windir%
        '        ' XP - %windir%
        '    Case CSIDL_WINDOWS  ' FIXED
        '        FOLDERID_Windows = "{F38BF404-1D43-42F2-9305-67DE0B28FC23}"
        '        sRfid = FOLDERID_Windows
        '        ' VISTA - %USERPROFILE% (%SystemDrive%\Users\%USERNAME%)
        '        ' XP - %USERPROFILE% (%SystemDrive%\Documents and Settings\%USERNAME%)
        '    Case CSIDL_PROFILE  ' FIXED
        '        FOLDERID_Profile = "{5E6C858F-0E22-4760-9AFE-EA3317B67173}"
        '        sRfid = FOLDERID_Profile
        '        ' VISTA - %USERPROFILE%\Pictures
        '        ' XP - %USERPROFILE%\My Documents\My Pictures
        '    Case CSIDL_MYPICTURES  ' PERUSER
        '        FOLDERID_Pictures = "{33E28130-4E1E-4676-835A-98395C3BC3BB}"
        '        sRfid = FOLDERID_Pictures
        '        ' VISTA - %ProgramFiles% (%SystemDrive%\Program Files)
        '        ' XP - %ProgramFiles% (%SystemDrive%\Program Files)
        '    Case CSIDL_PROGRAM_FILESX86  ' FIXED
        '        FOLDERID_ProgramFilesX86 = "{7C5A40EF-A0FB-4BFC-874A-C0F2E0B9FA8E}"
        '        sRfid = FOLDERID_ProgramFilesX86
        '        ' VISTA - %ProgramFiles%\Common Files
        '        ' XP - %ProgramFiles%\Common Files
        '    Case CSIDL_PROGRAM_FILES_COMMONX86  ' FIXED
        '        FOLDERID_ProgramFilesCommonX86 = "{DE974D24-D9C6-4D3E-BF91-F4455120B917}"
        '        sRfid = FOLDERID_ProgramFilesCommonX86
        '        ' VISTA - %ProgramFiles% (%SystemDrive%\Program Files)
        '        ' XP - %ProgramFiles% (%SystemDrive%\Program Files)
        '    Case 401  ' FIXED
        '        FOLDERID_ProgramFilesX64 = "{6D809377-6AF0-444b-8957-A3773F02200E}"
        '        sRfid = FOLDERID_ProgramFilesX64
        '        ' VISTA - %ProgramFiles%\Common Files
        '        ' XP - %ProgramFiles%\Common Files
        '    Case 402
        '        FOLDERID_ProgramFilesCommonX64 = "{6365D5A7-0F0D-45e5-87F6-0DA56B6A4F7D}"
        '        sRfid = FOLDERID_ProgramFilesCommonX64
        '        ' VISTA - %ProgramFiles% (%SystemDrive%\Program Files)
        '        ' XP - %ProgramFiles% (%SystemDrive%\Program Files)
        '    Case CSIDL_PROGRAM_FILES  ' FIXED
        '        FOLDERID_ProgramFiles = "{905e63b6-c1bf-494e-b29c-65b732d3d21a}"
        '        sRfid = FOLDERID_ProgramFiles
        '        ' VISTA - %ProgramFiles%\Common Files
        '        ' XP - %ProgramFiles%\Common Files
        '    Case CSIDL_PROGRAM_FILES_COMMON  ' FIXED
        '        FOLDERID_ProgramFilesCommon = "{F7F1ED05-9F6D-47A2-AAAE-29D317C6F066}"
        '        sRfid = FOLDERID_ProgramFilesCommon
        '        ' VISTA - %APPDATA%\Microsoft\Windows\Start Menu\Programs\Administrative Tools
        '        ' XP - %USERPROFILE%\Start Menu\Programs\Administrative Tools
        '    Case CSIDL_ADMINTOOLS  ' PERUSER
        '        FOLDERID_AdminTools = "{724EF170-A42D-4FEF-9F26-B60E846FBA4F}"
        '        sRfid = FOLDERID_AdminTools
        '        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\Start Menu\Programs\Administrative Tools
        '        ' XP - %ALLUSERSPROFILE%\Start Menu\Programs\Administrative Tools
        '    Case CSIDL_COMMON_ADMINTOOLS  ' COMMON
        '        FOLDERID_CommonAdminTools = "{D0384E7D-BAC3-4797-8F14-CBA229B392B5}"
        '        sRfid = FOLDERID_CommonAdminTools
        '        ' VISTA - %USERPROFILE%\Music
        '        ' XP - %USERPROFILE%\My Documents\My Music
        '    Case CSIDL_MYMUSIC  ' PERUSER
        '        FOLDERID_Music = "{4BD8D571-6D19-48D3-BE97-422220080E43}"
        '        sRfid = FOLDERID_Music
        '        ' VISTA - %USERPROFILE%\Videos
        '        ' XP - %USERPROFILE%\My Documents\My Videos
        '    Case CSIDL_MYVIDEO  ' PERUSER
        '        FOLDERID_Videos = "{18989B1D-99B5-455B-841C-AB7C74E4DDFC}"
        '        sRfid = FOLDERID_Videos
        '        ' VISTA - %PUBLIC%\Pictures
        '        ' XP - %ALLUSERSPROFILE%\Documents\My Pictures
        '    Case CSIDL_COMMON_PICTURES  ' COMMON
        '        FOLDERID_PublicPictures = "{B6EBFB86-6907-413C-9AF7-4FC2ABF07CC5}"
        '        sRfid = FOLDERID_PublicPictures
        '        ' VISTA - %PUBLIC%\Music
        '        ' XP - %ALLUSERSPROFILE%\Documents\My Music
        '    Case CSIDL_COMMON_MUSIC  ' COMMON
        '        FOLDERID_PublicMusic = "{3214FAB5-9757-4298-BB61-92A9DEAA44FF}"
        '        sRfid = FOLDERID_PublicMusic
        '        'VISTA - %PUBLIC%\Videos
        '        ' XP - %ALLUSERSPROFILE%\Documents\My Videos
        '    Case CSIDL_COMMON_VIDEO  ' COMMON
        '        FOLDERID_PublicVideos = "{2400183A-6185-49FB-A2D8-4A392A602BA3}"
        '        sRfid = FOLDERID_PublicVideos
        '        ' VISTA - %windir%\Resources
        '        ' XP - %windir%\Resources
        '    Case CSIDL_RESOURCES  ' FIXED
        '        FOLDERID_ResourceDir = "{8AD10C31-2ADB-4296-A8F7-E4701232C972}"
        '        sRfid = FOLDERID_ResourceDir
        '        ' VISTA - %windir%\resources\0409 (code page)
        '        ' XP - %windir%\resources\0409 (code page)
        '    Case CSIDL_RESOURCES_LOCALIZED  ' FIXED
        '        FOLDERID_LocalizedResourcesDir = "{2A00375E-224C-49DE-B8D1-440DF7EF3DDC}"
        '        sRfid = FOLDERID_LocalizedResourcesDir
        '        ' VISTA - %ALLUSERSPROFILE%\OEM Links
        '        ' XP - %ALLUSERSPROFILE%\OEM Links
        '    Case CSIDL_COMMON_OEM_LINKS  ' COMMON
        '        FOLDERID_CommonOEMLinks = "{C1BAE2D0-10DF-4334-BEDD-7AA20B227A9D}"
        '        sRfid = FOLDERID_CommonOEMLinks
        '        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\Burn\Burn
        '        ' XP - %USERPROFILE%\Local Settings\Application Data\Microsoft\CD Burning
        '    Case CSIDL_CDBURN_AREA  ' PERUSER
        '        FOLDERID_CDBurning = "{9E52AB10-F80D-49DF-ACB8-4330F5687855}"
        '        sRfid = FOLDERID_CDBurning
        '        ' VISTA - %SystemDrive%\Users
        '        ' XP - NONE
        '    Case 501  ' FIXED
        '        FOLDERID_UserProfiles = "{0762D272-C50A-4BB0-A382-697DCD729B80}"
        '        sRfid = FOLDERID_UserProfiles
        '        ' VISTA - %USERPROFILE%\Music\Playlists
        '        ' XP - NONE
        '    Case 502  ' PERUSER
        '        FOLDERID_Playlists = "{DE92C1C7-837F-4F69-A3BB-86E631204A23}"
        '        sRfid = FOLDERID_Playlists
        '        ' VISTA - %PUBLIC%\Music\Sample Playlists
        '        ' XP - NONE
        '    Case 503  ' COMMON
        '        FOLDERID_SamplePlaylists = "{15CA69B3-30EE-49C1-ACE1-6B5EC372AFB5}"
        '        sRfid = FOLDERID_SamplePlaylists
        '        ' VISTA - %PUBLIC%\Music\Sample Music
        '        ' XP - %ALLUSERSPROFILE%\Documents\My Music\Sample Music
        '    Case 504  ' COMMON
        '        FOLDERID_SampleMusic = "{B250C668-F57D-4EE1-A63C-290EE7D1AA1F}"
        '        sRfid = FOLDERID_SampleMusic
        '        ' VISTA - %PUBLIC%\Pictures\Sample Pictures
        '        ' XP - %ALLUSERSPROFILE%\Documents\My Pictures\Sample Pictures
        '    Case 505  ' COMMON
        '        FOLDERID_SamplePictures = "{C4900540-2379-4C75-844B-64E6FAF8716B}"
        '        sRfid = FOLDERID_SamplePictures
        '        ' VISTA - %PUBLIC%\Videos\Sample Videos
        '        ' XP - NONE
        '    Case 506  ' COMMON
        '        FOLDERID_SampleVideos = "{859EAD94-2E85-48AD-A71A-0969CB56A6CD}"
        '        sRfid = FOLDERID_SampleVideos
        '        ' VISTA - %USERPROFILE%\Pictures\Slide Shows
        '        ' XP - NONE
        '    Case 507  ' PERUSER
        '        FOLDERID_PhotoAlbums = "{69D2CF90-FC33-4FB7-9A0C-EBB0F0FCB43C}"
        '        sRfid = FOLDERID_PhotoAlbums
        '        ' VISTA - %PUBLIC% (%SystemDrive%\Users\Public)
        '        ' XP - NONE
        '    Case 508  'FIXED
        '        FOLDERID_Public = "{DFDF76A2-C82A-4D63-906A-5644AC457385}"
        '        sRfid = FOLDERID_Public
        '    Case 509  ' VIRTUAL
        '        FOLDERID_ChangeRemovePrograms = "{df7266ac-9274-4867-8d55-3bd661de872d}"
        '        sRfid = FOLDERID_ChangeRemovePrograms
        '    Case 510  ' VIRTUAL
        '        FOLDERID_AppUpdates = "{a305ce99-f527-492b-8b1a-7e76fa98d6e4}"
        '        sRfid = FOLDERID_AppUpdates
        '    Case 511  ' VIRTUAL
        '        FOLDERID_AddNewPrograms = "{de61d971-5ebc-4f02-a3a9-6c82895e5c04}"
        '        sRfid = FOLDERID_AddNewPrograms
        '        ' VISTA - %USERPROFILE%\Downloads
        '        ' XP - NONE
        '    Case 512  ' PERUSER
        '        FOLDERID_Downloads = "{374DE290-123F-4565-9164-39C4925E467B}"
        '        sRfid = FOLDERID_Downloads
        '        ' VISTA - %PUBLIC%\Downloads
        '        ' XP - NONE
        '    Case 513  ' COMMON
        '        FOLDERID_PublicDownloads = "{3D644C9B-1FB8-4f30-9B45-F670235F79C0}"
        '        sRfid = FOLDERID_PublicDownloads
        '        ' VISTA - %USERPROFILE%\Searches
        '        ' XP - NONE
        '    Case 514  ' PERUSER
        '        FOLDERID_SavedSearches = "{7d1d3a04-debb-4115-95cf-2f29da2920da}"
        '        sRfid = FOLDERID_SavedSearches
        '        'VISTA - %APPDATA%\Microsoft\Internet Explorer\Quick Launch
        '        'XP - %APPDATA%\Microsoft\Internet Explorer\Quick Launch
        '    Case 515  ' PERUSER
        '        FOLDERID_QuickLaunch = "{52a4f021-7b75-48a9-9f6b-4b87a210bc8f}"
        '        sRfid = FOLDERID_QuickLaunch
        '        ' VISTA - %USERPROFILE%\Contacts
        '        ' XP - NONE
        '    Case 516  ' PERUSER
        '        FOLDERID_Contacts = "{56784854-C6CB-462b-8169-88E350ACB882}"
        '        sRfid = FOLDERID_Contacts
        '        ' VISTA -%LOCALAPPDATA%\Microsoft\Windows Sidebar\Gadgets
        '        ' XP - NONE
        '    Case 517  ' PERUSER
        '        FOLDERID_SidebarParts = "{A75D362E-50FC-4fb7-AC2C-A8BEAA314493}"
        '        sRfid = FOLDERID_SidebarParts
        '        ' VISTA - %ProgramFiles%\Windows Sidebar\Gadgets
        '        ' XP - NONE
        '    Case 518  ' COMMON
        '        FOLDERID_SidebarDefaultParts = "{7B396E54-9EC5-4300-BE0A-2482EBAE1A26}"
        '        sRfid = FOLDERID_SidebarDefaultParts
        '    Case 519  ' NOT USED
        '        FOLDERID_TreeProperties = "{5b3749ad-b49f-49c1-83eb-15370fbd4882}"
        '        sRfid = FOLDERID_TreeProperties
        '        ' VISTA - %ALLUSERSPROFILE%\Microsoft\Windows\GameExplorer
        '        ' XP - NONE
        '    Case 520  ' COMMON
        '        FOLDERID_PublicGameTasks = "{DEBF2536-E1A8-4c59-B6A2-414586476AEA}"
        '        sRfid = FOLDERID_PublicGameTasks
        '        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows\GameExplorer
        '        ' XP - NONE
        '    Case 521  ' PERUSER
        '        FOLDERID_GameTasks = "{054FAE61-4DD8-4787-80B6-090220C4B700}"
        '        sRfid = FOLDERID_GameTasks
        '        ' VISTA - %USERPROFILE%\Saved Games
        '        ' XP - NONE
        '    Case 522  ' PERUSER
        '        FOLDERID_SavedGames = "{4C5C32FF-BB9D-43b0-B5B4-2D72E54EAAA4}"
        '        sRfid = FOLDERID_SavedGames
        '    Case 523  ' VIRTUAL
        '        FOLDERID_Games = "{CAC52C1A-B53D-4edc-92D7-6B2E8AC19434}"
        '        sRfid = FOLDERID_Games
        '    Case 524  ' NOT USED
        '        FOLDERID_RecordedTV = "{bd85e001-112e-431e-983b-7b15ac09fff1}"
        '        sRfid = FOLDERID_RecordedTV
        '    Case 525  ' VIRTUAL
        '        FOLDERID_SEARCH_MAPI = "{98ec0e18-2098-4d44-8644-66979315a281}"
        '        sRfid = FOLDERID_SEARCH_MAPI
        '    Case 526  ' VIRTUAL
        '        FOLDERID_SEARCH_CSC = "{ee32e446-31ca-4aba-814f-a5ebd2fd6d5e}"
        '        sRfid = FOLDERID_SEARCH_CSC
        '        ' VISTA - %USERPROFILE%\Links
        '        ' XP - NONE
        '    Case 527  ' PERUSER
        '        FOLDERID_Links = "{bfb9d5e0-c6a9-404c-b2b2-ae6db6af4968}"
        '        sRfid = FOLDERID_Links
        '    Case 528  ' VIRTUAL
        '        FOLDERID_UsersFiles = "{f3ce0f7c-4901-4acc-8648-d5d44b04ef8f}"
        '        sRfid = FOLDERID_UsersFiles
        '    Case 529  ' VIRTUAL
        '        FOLDERID_SearchHome = "{190337d1-b8ca-4121-a639-6d472d16972a}"
        '        sRfid = FOLDERID_SearchHome
        '        ' VISTA - %LOCALAPPDATA%\Microsoft\Windows Photo Gallery\Original Images
        '        ' XP - NONE
        '    Case 530  ' PERUSER
        '        FOLDERID_OriginalImages = "{2C36C0AA-5812-4b87-BFD0-4CD0DFB19B39}"
        '        sRfid = FOLDERID_OriginalImages
        'End Select
        Dim Ret As Long
        Dim Trash As String
        Trash = Space$(260)
        ActivelockGetSpecialFolder = ""
        Try
            Ret = SHGetSpecialFolderPath(0, Trash, CSIDL, False)
            If Trim$(Trash) <> Chr(0) Then
                Trash = Left$(Trash, InStr(Trash, Chr(0)) - 1)
            End If
            ActivelockGetSpecialFolder = Trash
            Return ActivelockGetSpecialFolder
        Catch ex As Exception
            Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrUndefinedSpecialFolder, ACTIVELOCKSTRING, STRUNDEFINEDSPECIALFOLDER)
        End Try

    End Function

    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PSWD() As String
        ' Do not modify this unless you change all encrypted strings in the entire project
        PSWD = Chr(109) & Chr(121) & Chr(108) & Chr(111) & Chr(118) & Chr(101) & Chr(97) & Chr(99) & Chr(116) & Chr(105) & Chr(118) & Chr(101) & "lock"
    End Function

    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function REGKEY1() As String
        ' Do not modify this unless you change all encrypted strings in the entire project
        REGKEY1 = Chr(65) & Chr(112) & Chr(112) & Chr(69) & Chr(118) & Chr(101) & _
            Chr(110) & Chr(116) & Chr(115) & Chr(92) & Chr(83) & Chr(99) & Chr(104) & _
            Chr(101) & Chr(109) & Chr(101) & Chr(115) & Chr(92) & Chr(65) & Chr(112) & _
            Chr(112) & Chr(115) & Chr(92)
    End Function

    ''' <summary>
    ''' ?Not Documented!
    ''' </summary>
    ''' <param name="n1"></param>
    ''' <param name="n2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function IsNumberIncluded(ByVal n1 As Long, ByVal n2 As Long) As Boolean
        ' n1 = the larger number which may include n2
        ' n2 = the number we're checking as "is this a component?"
        Dim binary1 As String = String.Empty
        Dim binary2 As String = String.Empty
        IsNumberIncluded = False

        If n1 < n2 Then
            Exit Function
        ElseIf n1 <= 0 Or n2 <= 0 Then
            Exit Function
        ElseIf n1 = n2 Then
            IsNumberIncluded = True
        Else

            Do Until n1 = 0
                If (n1 Mod 2) Then
                    binary1 = binary1 & "1"   'write binary number BACKWARDS
                Else
                    binary1 = binary1 & "0"
                End If
                n1 = n1 \ 2
            Loop

            Do Until n2 = 0
                If (n2 Mod 2) Then
                    binary2 = binary2 & "1"   'write binary number BACKWARDS
                Else
                    binary2 = binary2 & "0"
                End If
                n2 = n2 \ 2
            Loop
            IsNumberIncluded = CBool(Mid$(binary1, Len(binary2), 1) = "1")
            If binary1.Substring(binary2.Length - 1, 1) = "1" Then IsNumberIncluded = True

        End If

    End Function

    ''' <summary>
    ''' Checks to see if there is a web connection.
    ''' </summary>
    ''' <returns>True if connected, False otherwise.</returns>
    ''' <remarks>This also sets the ConnectionQualityString</remarks>
    Public Function IsWebConnected() As Boolean
        ' Returns True if connection is available

        Dim lngFlags As Long

        If InternetGetConnectedState(lngFlags, 0) Then
            Return True ' True
            If lngFlags And InetConnState.lan Then
                Select Case ConnectionQualityString
                    Case "Good"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Intermittent"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Off"
                        'lblConnectStatus.ForeColor = Color.DarkOrange
                        'lblConnectStatus.Text = _
                        '     "Connection Quality:  Intermittent"
                        ConnectionQualityString = "Intermittent"
                End Select
            ElseIf lngFlags And InetConnState.modem Then
                Select Case ConnectionQualityString
                    Case "Good"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Intermittent"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Off"
                        'lblConnectStatus.ForeColor = Color.DarkOrange
                        'lblConnectStatus.Text = _
                        '     "Connection Quality:  Intermittent"
                        ConnectionQualityString = "Intermittent"
                End Select
            ElseIf lngFlags And InetConnState.configured Then
                Select Case ConnectionQualityString
                    Case "Good"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Intermittent"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Off"
                        'lblConnectStatus.ForeColor = Color.DarkOrange
                        'lblConnectStatus.Text = _
                        '     "Connection Quality:  Intermittent"
                        ConnectionQualityString = "Intermittent"
                End Select
            ElseIf lngFlags And InetConnState.proxy Then
                Select Case ConnectionQualityString
                    Case "Good"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Intermittent"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Off"
                        'lblConnectStatus.ForeColor = Color.DarkOrange
                        'lblConnectStatus.Text = _
                        '     "Connection Quality:  Intermittent"
                        ConnectionQualityString = "Intermittent"
                End Select
            ElseIf lngFlags And InetConnState.ras Then
                Select Case ConnectionQualityString
                    Case "Good"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Intermittent"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Off"
                        'lblConnectStatus.ForeColor = Color.DarkOrange
                        'lblConnectStatus.Text = _
                        '     "Connection Quality:  Intermittent"
                        ConnectionQualityString = "Intermittent"
                End Select
            ElseIf lngFlags And InetConnState.offline Then
                Select Case ConnectionQualityString
                    Case "Good"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Intermittent"
                        'lblConnectStatus.ForeColor = Color.Green
                        'lblConnectStatus.Text = "Connection Quality:  Good"
                        ConnectionQualityString = "Good"
                    Case "Off"
                        'lblConnectStatus.ForeColor = Color.DarkOrange
                        'lblConnectStatus.Text = _
                        '     "Connection Quality:  Intermittent"
                        ConnectionQualityString = "Intermittent"
                End Select
            End If
        Else
            ' False
            Select Case ConnectionQualityString
                Case "Good"
                    'lblConnectStatus.ForeColor = Color.DarkOrange
                    'lblConnectStatus.Text = "Connection Quality:  Intermittent"
                    ConnectionQualityString = "Intermittent"
                Case "Intermittent"
                    'lblConnectStatus.ForeColor = Color.Red
                    'lblConnectStatus.Text = "Connection Quality:  Off"
                    ConnectionQualityString = "Off"
                Case "Off"
                    'lblConnectStatus.ForeColor = Color.Red
                    'lblConnectStatus.Text = "Connection Quality:  Off"
                    ConnectionQualityString = "Off"
            End Select
        End If

    End Function

End Module