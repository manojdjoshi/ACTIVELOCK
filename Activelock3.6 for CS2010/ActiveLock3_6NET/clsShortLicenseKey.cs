using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

/// <summary>
/// <para>Use to provide license key generation and validation. This class exposes an
/// abstract interface that can be used to implement licensing for all of your
/// commerical and shareware applications.  Keys can be cloaked with a bit
/// swapping technique, and with a private key.  Keys can also be tied to a
/// licensee.</para>
/// </summary>
/// <remarks>See code, within, for more info!</remarks>
 // ERROR: Not supported in C#: OptionDeclaration
internal class clsShortLicenseKey
{

	#region "More Information & Copyright"
	//===============================================================================
	//
	//   LicenseKey Class
	//
	//   Use to provide license key generation and validation. This class exposes an
	//   abstract interface that can be used to implement licensing for all of your
	//   commerical and shareware applications.  Keys can be cloaked with a bit
	//   swapping technique, and with a private key.  Keys can also be tied to a
	//   licensee.
	//
	//   Use the global conditional compile constants (IncludeCreate, IncludeValidate,
	//   and IncludeCheck) to define which members are compiled into your project.
	//   For instance, set IncludeCreate = 0 to exclude it from the client app.
	//
	//   Extra effort is made in the ValidateKey method to so that the entire key is
	//   not held in memory at any time. Keep that in mind if you alter the source.
	//
	//   This implementation breaks a key into the following parts:
	//
	//       1111-2222-3333-4444-5555
	//
	//       1111 = Product code
	//       2222 = Expiration date (days since 1-1-1970)
	//       3333 = Caller definable word (16 bit value)
	//       4444 = CRC for key validation
	//       5555 = CRC for input validation
	//
	//   IMPORTANT: Key generators (no matter how good they are) will NOT thwart a
	//   cracker! Alter the source code to meet your proprietary needs.
	//
	//===============================================================================
	//
	//   Author:             Monte Hansen [monte@killervb.com]
	//   Dependencies:       None
	//   Invitation:         There is an open invitation to comment on this code,
	//                       report bugs or request revisions or enhancements.
	//
	//===============================================================================
	//
	//   ==  Copyright © 1999-2001 by Monte Hansen, All Rights Reserved Worldwide  ==
	//
	//   Monte Hansen  (The Author) grants a royalty-free right to use,  modify,  and
	//   distribute this code  (The Code)  in compiled form,  provided that you agree
	//   that The Author has no warranty,  obligations  or  liability  for  The Code.
	//   You may distribute The Code among peers but may not sell it,  or  distribute
	//   it on any electronic or physical media such  as  floppy  diskettes,  compact
	//   disks, bulletin boards, web sites, and the like, without first obtaining The
	//   Author's consent.
	//
	//   When distributing The Code among peers,  it is respectfully  requested  that
	//   it be distributed as is,  but at no time shall it be distributed without the
	//   copyright notice hereinabove.
	//
	//===============================================================================
	#endregion

	//===============================================================================
	//   Constants
	//===============================================================================
	//UPGRADE_NOTE: Module was upgraded to Module_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'

	private const string Module_Renamed = "clsShortLicenseKey";
	#region "Enums"

	/// <summary>
	/// segments to the license key
	/// </summary>
	/// <remarks></remarks>
	private enum Segments
	{
		// segments to the license key
		iProdCode = 0,
		iExpire = 1,
		iUserData = 2,
		iCRC = 3,
		iCRC2 = 4
	}

	/// <summary>
	/// Undocumented!
	/// </summary>
	/// <remarks></remarks>
	private enum MapFileChecksumErrors
	{
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

	#endregion

	//===============================================================================
	//   Types
	//===============================================================================
	// This structure is used to store a
	// reference to two bits that will be
	// swapped. Each bit can be from a
	// different segment in the key.
	// iCRC2 cannot be swapped since it
	// is a checksum of the first 4
	// segments of the key.
	private struct TBits
	{
		public byte iWord1;
		public byte iBit1;
		public byte iWord2;
		public byte iBit2;
	}

	//===============================================================================
	//   Private Members
	//===============================================================================
	private TBits[] m_Bits;

	private int m_nSwaps;
	[DllImport("kernel32", EntryPoint = "RtlMoveMemory", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern void CopyMemory(int lpDest, int lpSource, int nBytes);
	[DllImport("IMAGEHLP.DLL", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int MapFileAndCheckSumA(string FileName, ref int HeaderSum, ref int CheckSum);
	//===============================================================================
	//   Declares
	//===============================================================================

	// Note: add a project conditional compile argument "IncludeCreate"
	// if the CreateShortKey is to be compiled into the application.
	//#If IncludeCreate = 1 Then
	internal string CreateShortKey(string SerialNumber, string Licensee, int ProductCode, System.DateTime ExpireDate, short UserData, int RegisteredLevel)
	{
		//===============================================================================
		//   CreateShortKey - Creates a new serial number.
		//
		//   SerialNumber    The serial number is generated from the app name, version,
		//                   and password, along with the HDD firmware serial number,
		//                   which makes it unique for the machine running the app.
		//   Licensee        Name of party to whom this license is issued.
		//   ProductCode     A unique number assigned to this product. This is created
		//                   from the app private key and is a 4 digit integer.
		//   ExpireDate      Use this field for time-based serial numbers. This allows
		//                   serial number to be issued that expire in two weeks or at
		//                   the end of the year.
		//   UserData        This field is caller defined. Currently we are using
		//                   the MaxUser and LicType (using a LoByte/HiByte packed field).
		//   RegisteredLevel This is the Registered Level from Alugen. Long only.
		//   RETURNS         A License Key in the form of "233C-3912-00FF-BE49"
		//===============================================================================

		string[] KeySegs = new string[5];
		int i = 0;

		// convert each segment value to a hex string
		KeySegs[Segments.iProdCode] = HexWORD(ProductCode);
		KeySegs[Segments.iExpire] = HexWORD(DateAndTime.DateValue((string)(System.DateTime)ExpireDate.ToString("yyyy/MM/dd")).Subtract(DateAndTime.DateValue((string)1970/01/01 12:00:00 AM)).Days);
		KeySegs[Segments.iUserData] = HexWORD(UserData);

		// Compute CRC against pertinent info.
		KeySegs[Segments.iCRC] = HexWORD(CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(Strings.UCase(Licensee + KeySegs[Segments.iProdCode] + KeySegs[Segments.iExpire] + KeySegs[Segments.iUserData] + SerialNumber))));

		// Perform bit swaps
		for (i = 0; i <= m_nSwaps - 1; i++) {
			SwapBit(ref m_Bits(i), ref KeySegs);
		}

		// Calculate the CRC used to perform
		// simple user input validation.
		KeySegs[Segments.iCRC2] = HexWORD(CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(Strings.UCase(Licensee + KeySegs[Segments.iProdCode] + KeySegs[Segments.iExpire] + KeySegs[Segments.iUserData] + KeySegs[Segments.iCRC]))));

		// Return the key to the caller
		return Strings.UCase(KeySegs[Segments.iProdCode] + "-" + KeySegs[Segments.iExpire] + "-" + KeySegs[Segments.iUserData] + "-" + KeySegs[Segments.iCRC] + "-" + KeySegs[Segments.iCRC2]) + "-" + Strings.StrReverse(HexWORD(RegisteredLevel));

	}
	//#End If

	//#If IncludeCheck = 1 Then
	internal bool TestKey(string LicenseKey, string Licensee)
	{
		bool functionReturnValue = false;
		//===============================================================================
		//   TestKey - Performs a simple CRC test to ensure the key was entered
		//   "correctly". Does NOT validate that the key is VALID. This function allows
		//   the caller to "test" the key input by the user, without having to execute
		//   the key validation code, making it more work for a cracker to generate a
		//   key for your application.
		//===============================================================================

		object KeySegs = null;
		int nCRC = 0;

		 // ERROR: Not supported in C#: OnErrorStatement


		// TODO: don't even call this function if SoftIce was detected!

		if (!SplitKey(ref LicenseKey, ref KeySegs)) goto ExitLabel; 

		// NOTE: Licensee can be omitted from the last checksum
		// if there is no need to bind a serial number to a
		// customer name.
		nCRC = CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(Strings.UCase(Licensee + KeySegs(Segments.iProdCode) + KeySegs(Segments.iExpire) + KeySegs(Segments.iUserData) + KeySegs(Segments.iCRC))));

		// Compare check digits
		functionReturnValue = (nCRC == SegmentValue(KeySegs(Segments.iCRC2)));
		ExitLabel:
		return functionReturnValue;


	}
	//#End If

	//#If IncludeValidate = 1 Then
	/// <summary>
	/// Evaluates the supplied license key and tests that it is valid. We do this by recomputing the checksum and comparing it to the one embedded in the serial number.
	/// </summary>
	/// <param name="LicenseKey">The license number to validate. Liberation Key.</param>
	/// <param name="SerialNumber">A magic string that is application specific. This should be the same as was originally created by the application.</param>
	/// <param name="Licensee">Name of party to whom this license is issued. This should be the same as was used to create the serial number.</param>
	/// <param name="ProductCode">A unique 4 digit number assigned to this product. This should be the same as was used to create the license key.</param>
	/// <param name="ExpireDate">Use this field for time-based serial numbers. This should be the same as was used to create the license key.</param>
	/// <param name="UserData">This field is caller defined. This should be the same as was used to create the license key.</param>
	/// <param name="RegisteredLevel"></param>
	/// <returns>True if the license key checks out, False otherwise.</returns>
	/// <remarks>See code for important notes!</remarks>
	internal bool ValidateShortKey(
		string LicenseKey,
		string SerialNumber,
		string Licensee,
		int ProductCode,
		[
			System.Runtime.InteropServices.OptionalAttribute,
			System.Runtime.InteropServices.DefaultParameterValueAttribute("1970/01/01 12:00:00 AM")
		] ref System.DateTime ExpireDate,
		[
			System.Runtime.InteropServices.OptionalAttribute,
			System.Runtime.InteropServices.DefaultParameterValueAttribute(0)
		] ref short UserData,
		[
			System.Runtime.InteropServices.OptionalAttribute,
			System.Runtime.InteropServices.DefaultParameterValueAttribute(0)
		] ref int RegisteredLevel)
	{
		bool functionReturnValue = false;
		//*******************************************************************************
		//   IMPORTANT       This function is where the most care must be given.
		//                   You should assume that a cracker has seen this code and can
		//                   recognize it from ASM listings, and should be changed.
		//                   - Avoid string compares whenever possible.
		//                   - Pepper lots of JUNK code.
		//                   - Do things in different order (except CRC checks).
		//                   - Do not do things in this routine that are being monitored
		//                     (registry calls, file-system access, phone home w/TCP).
		//                   - Remove the UCase$ statements (just pass serial in ucase).
		//*******************************************************************************

		object KeySegs = null;
		int nCrc1 = 0;
		int nCrc2 = 0;
		int nCrc3 = 0;
		int nCrc4 = 0;
		int i = 0;

		 // ERROR: Not supported in C#: OnErrorStatement


		// TODO: don't even call this function if SoftIce was detected!
		// ----------------------------------------------------------
		// This section of code could raise red flags
		// ----------------------------------------------------------
		RegisteredLevel = SegmentValue(Strings.StrReverse(Strings.Mid(LicenseKey, 26, 4))) - 200;
		LicenseKey = Strings.Mid(LicenseKey, 1, 24);
		if (!SplitKey(ref LicenseKey, ref KeySegs)) goto ExitLabel; 

		// ----------------------------------------------------------
		// TODO: UCase string before it get's here

		// Get CRC used for input validation
		nCrc1 = CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(Strings.UCase(Licensee) + KeySegs(Segments.iProdCode) + KeySegs(Segments.iExpire) + KeySegs(Segments.iUserData) + KeySegs(Segments.iCRC)));

		// Compare check digits
		if ((nCrc1 != SegmentValue(KeySegs(Segments.iCRC2)))) {
			goto ExitLabel;
		}

		// Perform bit swaps (in reverse).
		for (i = m_nSwaps - 1; i >= 0; i += -1) {
			SwapBit(ref m_Bits(i), ref KeySegs);
		}

		// Calculate checksum on the license KeySegs.
		// The LAST thing we want to do is to push a valid
		// serial number on to the stack. This is the first
		// thing a cracker will look for. Instead we will
		// calculate a running checksum on each segment and
		// compare the checksum to the checksum embedded in
		// the key.

		// The supplied product code should be
		// the same as the product code embedded
		// in the key.
		if (ProductCode == SegmentValue(KeySegs(Segments.iProdCode))) {
			// One more check on the check digits before we
			// blow away the value stored in nCrc1.
			if ((SegmentValue(KeySegs(Segments.iCRC2)) == nCrc1)) {
				nCrc1 = CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(Strings.UCase(Licensee)));
			}
		}

		nCrc2 = CRC(ref System.Text.UnicodeEncoding.Unicode.GetBytes(Strings.UCase(KeySegs(Segments.iProdCode))), ref nCrc1);
		nCrc3 = CRC(ref System.Text.UnicodeEncoding.Unicode.GetBytes(Strings.UCase(KeySegs(Segments.iExpire))), ref nCrc2);
		nCrc3 = CRC(ref System.Text.UnicodeEncoding.Unicode.GetBytes(Strings.UCase(KeySegs(Segments.iUserData))), ref nCrc3);
		nCrc4 = CRC(ref System.Text.UnicodeEncoding.Unicode.GetBytes(Strings.UCase(SerialNumber)), ref nCrc3);

		// Return success and fill outputs IF the license KeySegs is valid

		if (nCrc4 == SegmentValue(KeySegs(Segments.iCRC))) {
			// Fill the outputs with expire date and user data.
			ExpireDate = 1970/01/01 12:00:00 AM;
			ExpireDate = ExpireDate.AddDays(SegmentValue(KeySegs(Segments.iExpire)));
			UserData = SegmentValue(KeySegs(Segments.iUserData));

			// IMPORTANT: This is an easy patch point
			// if you use real-time key validation.
			functionReturnValue = true;

		}
		ExitLabel:
		return functionReturnValue;


	}
	//#End If

	//#If IncludeValidate = 1 Or IncludeCreate = 1 Then
	internal void AddSwapBits(int Word1, int Bit1, int Word2, int Bit2)
	{
		//===============================================================================
		//   AddSwapBits - This is used to swap various bits in the serial number. It's
		//   sole purpose is to alter the output serial number.
		//
		//   This process is "played" forwards during the key creation, and in reverse
		//   when validating. This mangling process should be identical for key creation
		//   and validation. Add as many combinations as you like.
		//
		//   Word1/Word2     The words to bit swap. There are 4 words in the serial #.
		//                   This parameter is zero-based.
		//   Bit1/Bit2       The bits to swap. There are 16 bits to each word.
		//                   This parameter is zero-based.
		//
		//   Example: This scenario causes word 3, bit 8 to be swapped with word 1, bit 3
		//
		//       KeyGen.AddSwapBits 1, 3, 3, 8
		//
		//   NOTE:   It is recommended that there be at least 6 combinations in case
		//   the bits being swapped are the same (2 swap bits for words 2, 3 & 4).
		//===============================================================================

		// TODO: don't even call this function if SoftIce was detected!

		// Size array to fit
		if (m_nSwaps == 0) {
			 // ERROR: Not supported in C#: ReDimStatement

		}
		else {
			Array.Resize(ref m_Bits, m_nSwaps + 1);
		}
		m_nSwaps = m_nSwaps + 1;

		// This implementation hardcodes keys that are 8 bytes/4 words
		if (Word1 < 0 | Word1 > 3 | Word2 < 0 | Word2 > 3) {
			modActiveLock.Set_locale(modActiveLock.regionalSymbol);
			Err().Raise(5, Module_Renamed, "Word specification is not within 0-3.");
		}

		// There are only 16 bits to a word.
		if (Bit1 < 0 | Bit1 > 15 | Bit2 < 0 | Bit2 > 15) {
			modActiveLock.Set_locale(modActiveLock.regionalSymbol);
			Err().Raise(5, Module_Renamed, "Bit specification is not within 0-15.");
		}

		// Save the bits to be swapped
		{
			m_Bits(m_nSwaps - 1).iWord1 = Word1;
			m_Bits(m_nSwaps - 1).iBit1 = Bit1;
			m_Bits(m_nSwaps - 1).iWord2 = Word2;
			m_Bits(m_nSwaps - 1).iBit2 = Bit2;
		}

	}
	//#End If

	//#If IncludeValidate = 1 Or IncludeCheck = 1 Then
	private bool SplitKey(ref string LicenseKey, ref object KeySegs)
	{
		bool functionReturnValue = false;
		//===============================================================================
		//   SplitKey - Shared code to massage the input serial number, and slice it into
		//   the required number of segments.
		//===============================================================================

		// ----------------------------------------------------------
		// This section of code could raise red flags
		// ----------------------------------------------------------

		// Sanity check
		if (Strings.InStr(LicenseKey, "-") == 0) goto ExitLabel; 

		// As a courtesy to the user, we convert the
		// letter "O" to the number "0". Users hate
		// serialz that do not have interchangable 0/o's!
		LicenseKey = Strings.Replace(LicenseKey, "o", "0", , , CompareMethod.Text);

		// Splice the KeySegs into 4 segments,
		// exit if wrong # of segments.
		KeySegs = Strings.Split(Strings.UCase(LicenseKey), "-", 5);
		if (Information.UBound(KeySegs) != 4) goto ExitLabel; 

		// ----------------------------------------------------------

		functionReturnValue = true;
		ExitLabel:
		return functionReturnValue;


	}
	//#End If

	private int SegmentValue(string HexString)
	{
		//===============================================================================
		//   Converts a hex string representation into a 4 byte decimal value.
		//===============================================================================

		//Dim Buffer(3) As Byte
		//Dim i As Integer
		//Dim j As Integer

		//' Exit if each byte not represented by a 2 character string
		//If Len(HexString) Mod 2 <> 0 Then Exit Function

		//' Exit if it's larger than a 4 byte value
		//If Len(HexString) > 8 Then Exit Function

		//' NOTE: we populate the byte array in little-endian format
		//For i = Len(HexString) To 1 Step -2
		//    Buffer(j) = CByte("&H" & Mid(HexString, i - 1, 2))
		//    j = j + 1
		//Next i

		//' Return the value
		//CopyMemory(VarPtr(SegmentValue), VarPtr(Buffer(0)), 4)
		return Int32.Parse(HexString, System.Globalization.NumberStyles.HexNumber);

	}


	//#If IncludeCreate = 1 Or IncludeValidate = 1 Then
	private void SwapBit(ref TBits BitList, ref object KeySegs)
	{
		//===============================================================================
		//   SwapBit - Swaps any two bits. The bits can differ as long as they are in
		//   the range of 0 and 15.
		//===============================================================================
		{
			// Essentially, we swap Bit1 with Bit2. We use a bitwise
			// OR operator or a bitwise AND operator depending
			// upon if the subject bit is present. We don't use
			// local variables to avoid synchronizing, especially
			// since we may be doing a bit swap on the same word.
			if ((SegmentValue(KeySegs(BitList.iWord1)) & (Math.Pow(2, BitList.iBit2))) == (Math.Pow(2, BitList.iBit2))) {
				KeySegs(BitList.iWord1) = HexWORD(SegmentValue(KeySegs(BitList.iWord1)) | (Math.Pow(2, BitList.iBit2)), "");
			}
			else {
				KeySegs(BitList.iWord1) = HexWORD(SegmentValue(KeySegs(BitList.iWord1)) & !(Math.Pow(2, BitList.iBit2)), "");
			}

			if ((SegmentValue(KeySegs(BitList.iWord2)) & (Math.Pow(2, BitList.iBit1))) == (Math.Pow(2, BitList.iBit1))) {
				KeySegs(BitList.iWord2) = HexWORD(SegmentValue(KeySegs(BitList.iWord2)) | (Math.Pow(2, BitList.iBit1)), "");
			}
			else {
				KeySegs(BitList.iWord2) = HexWORD(SegmentValue(KeySegs(BitList.iWord2)) & !(Math.Pow(2, BitList.iBit1)), "");
			}

		}

	}
	//#End If

	// Generic helper function
	private int CRC(ref byte[] Buffer, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(0)] ref  // ERROR: Optional parameters aren't supported in C#
int InputCrc)
	{
		int functionReturnValue = 0;
		//===============================================================================
		//   Crc - Returns a 16-bit CRC value for a data block.
		//
		//   Refer to CRC-CCITT compute-on-the-fly implementatations for more info.
		//===============================================================================

		int Bit = 0;
		int i = 0;
		int j = 0;

		 // ERROR: Not supported in C#: OnErrorStatement


		// Derive from a prior CRC value if supplied.
		functionReturnValue = InputCrc;

		// Loop thru entire buffer computing the CRC

		for (i = Information.LBound(Buffer); i <= Information.UBound(Buffer); i++) {
			// Loop thru each of the 8 bits

			for (j = 0; j <= 7; j++) {
				Bit = ((functionReturnValue & 0x8000) == 0x8000) ^ ((Buffer[i] & (Math.Pow(2, j))) == Math.Pow(2, j));

				functionReturnValue = ((short)functionReturnValue & 0x7fff * 2);

				if (Bit != 0) {
					functionReturnValue = functionReturnValue ^ 0x1021;
				}

			}

		}

		return;
		 // ERROR: Not supported in C#: ResumeStatement

		ErrHandler:
		System.Diagnostics.Debug.Assert(0, "");
		return functionReturnValue;

	}

	private string HexWORD(int WORD, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute("")]  // ERROR: Optional parameters aren't supported in C#
string Prefix)
	{
		string functionReturnValue = null;
		//===============================================================================
		//   HexDWORD - Returns a hex string representation of a WORD.
		//
		//   WORD            The 2 byte value to convert to a hex string.
		//   Prefix          A value such as "0x" or "&H".
		//
		//   NOTE:  It's up to the caller to ensure the subject value is a 16-bit number.
		//===============================================================================

		//Dim bytes(1) As Byte
		//Dim i As Integer, str As String

		//CopyMemory(VarPtr(bytes(0)), VarPtr(WORD), 2)

		functionReturnValue = Strings.UCase(System.Convert.ToString(WORD, 16));
		if (Strings.Len(HexWORD()) == 3) {
			functionReturnValue = "0" + functionReturnValue;
		}
		else if (Strings.Len(HexWORD()) == 2) {
			functionReturnValue = "00" + functionReturnValue;
		}
		else if (Strings.Len(HexWORD()) == 1) {
			functionReturnValue = "000" + functionReturnValue;
		}
		return functionReturnValue;

		//str = LCase(System.Convert.ToString(WORD, 16))
		//Dim encoding As New System.Text.ASCIIEncoding
		//bytes = encoding.GetBytes(str)

		//HexWORD = Prefix
		//For i = UBound(bytes) To LBound(bytes, ) Step -1
		//    If Len(Hex(bytes(i))) = 1 Then
		//        HexWORD = HexWORD & "0" & LCase(Hex(bytes(i)))
		//    Else
		//        HexWORD = HexWORD & LCase(Hex(bytes(i)))
		//    End If
		//Next i

	}

	internal bool ExeIsPatched(string FilePath)
	{
		bool functionReturnValue = false;
		//===============================================================================
		//   ExeIsPatched - Tests if the supplied file has been altered by computing a
		//   checksum for the file and comparing it against the checksum in the
		//   executable image.
		//
		//   FileName - Full path to file to check. Caller is responsible for ensuring
		//   that the path exists, and that it is an executable.
		//===============================================================================

		int FileCRC = 0;
		int HdrCRC = 0;
		int ErrorCode = 0;

		// NOTE: Many crackers today are smart enough to
		// update the PE image CRC value. But we check
		// anyhow, just in case. Otherwise, it could be
		// embarrassing if the EXE was patched without
		// updating the PE header.

		ErrorCode = MapFileAndCheckSumA(FilePath, ref HdrCRC, ref FileCRC);

		if (ErrorCode == MapFileChecksumErrors.CHECKSUM_SUCCESS) {

			if (HdrCRC != 0 & HdrCRC != FileCRC) {
				// CRC of file is different than the CRC
				// embedded in the PE image. Try not to
				// let the cracker know that you are on
				// to him. And don't start deleting from
				// their harddrive!
				functionReturnValue = true;

			}

		}
		return functionReturnValue;

	}
}
