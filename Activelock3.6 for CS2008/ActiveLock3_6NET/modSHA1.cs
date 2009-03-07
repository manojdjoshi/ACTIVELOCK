using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
static class modSHA1
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
	// Name: modSHA1
	// Purpose: SHA1 Encryption Module
	// (C) 1998 Ian Lynagh
	// Date Created:         June 16, 1998 - IL
	// Date Last Modified:   July 07, 2003 - MEC
	// Functions:
	// Properties:
	// Methods:
	// Started: 06.16.1998
	// Modified: 03.25.2006
	//===============================================================================
	//===============================================================================
	// Name: Function BigAA1Add
	// Input:
	//   ByVal value1 As String
	//   ByVal value2 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA1Add(string value1, string value2)
	{
		string valueans = string.Empty;
		short loopit = 0;
		short tempnum = 0;

		tempnum = Strings.Len(value1) - Strings.Len(value2);
		if (tempnum < 0) {
			value1 = Strings.Space(System.Math.Abs(tempnum)) + value1;
		}
		else if (tempnum > 0) {
			value2 = Strings.Space(System.Math.Abs(tempnum)) + value2;
		}

		tempnum = 0;
		for (loopit = Strings.Len(value1); loopit >= 1; loopit += -1) {
			tempnum = tempnum + Conversion.Val("&H" + Strings.Mid(value1, loopit, 1)) + Conversion.Val("&H" + Strings.Mid(value2, loopit, 1));
			valueans = Conversion.Hex(tempnum % 16) + valueans;
			tempnum = Conversion.Int(tempnum / 16);
		}
		if (tempnum != 0) {
			valueans = Conversion.Hex(tempnum) + valueans;
		}

		return valueans;

	}
	//===============================================================================
	// Name: Function BigAA1AND
	// Input:
	//   ByVal value1 As String
	//   ByVal value2 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA1AND(string value1, string value2)
	{
		string valueans = string.Empty;
		short loopit = 0;
		short tempnum = 0;

		tempnum = Strings.Len(value1) - Strings.Len(value2);
		if (tempnum < 0) {
			value2 = Strings.Mid(value2, System.Math.Abs(tempnum) + 1);
		}
		else if (tempnum > 0) {
			value1 = Strings.Mid(value1, tempnum + 1);
		}

		for (loopit = 1; loopit <= Strings.Len(value1); loopit++) {
			valueans = valueans + Conversion.Hex(Conversion.Val("&H" + Strings.Mid(value1, loopit, 1)) & Conversion.Val("&H" + Strings.Mid(value2, loopit, 1)));
		}

		return valueans;

	}
	//===============================================================================
	// Name: Function BigAA1Mod32Add
	// Input:
	//   ByVal value1 As String
	//   ByVal value2 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA1Mod32Add(string value1, string value2)
	{
		return Strings.Right(BigAA1Add(value1, value2), 8);
	}
	//===============================================================================
	// Name: Function BigAA1NOT
	// Input:
	//   ByVal value1 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA1NOT(string value1)
	{
		string valueans = string.Empty;
		short loopit = 0;

		value1 = Strings.Right(value1, 8);
		value1 = new string("0", 8 - Strings.Len(value1)) + value1;
		for (loopit = 1; loopit <= 8; loopit++) {
			valueans = valueans + Conversion.Hex(15 ^ Conversion.Val("&H" + Strings.Mid(value1, loopit, 1)));
		}

		return valueans;

	}
	//===============================================================================
	// Name: Function BigAA1OR
	// Input:
	//   ByVal value1 As String
	//   ByVal value2 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA1OR(string value1, string value2)
	{
		string valueans = null;
		short loopit = 0;
		short tempnum = 0;

		tempnum = Strings.Len(value1) - Strings.Len(value2);
		if (tempnum < 0) {
			valueans = Strings.Left(value2, System.Math.Abs(tempnum));
			value2 = Strings.Mid(value2, System.Math.Abs(tempnum) + 1);
		}
		else if (tempnum > 0) {
			valueans = Strings.Left(value1, System.Math.Abs(tempnum));
			value1 = Strings.Mid(value1, tempnum + 1);
		}

		for (loopit = 1; loopit <= Strings.Len(value1); loopit++) {
			valueans = valueans + Conversion.Hex(Conversion.Val("&H" + Strings.Mid(value1, loopit, 1)) | Conversion.Val("&H" + Strings.Mid(value2, loopit, 1)));
		}

		return valueans;

	}
	//===============================================================================
	// Name: Function BigAA1RotLeft
	// Input:
	//   ByVal value1 As String
	//   ByVal rots As Integer
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA1RotLeft(string value1, short rots)
	{
		string tempstr = null;
		short loopinner = 0;
		short loopit = 0;
		short tempnum = 0;

		rots = rots % 32;
		if (rots == 0) {
			BigAA1RotLeft() = value1;
			return;
		}
		value1 = Strings.Right(value1, 8);
		tempstr = new string("0", 8 - Strings.Len(value1)) + value1;
		value1 = "";

		// Convert to binary
		for (loopit = 1; loopit <= 8; loopit++) {
			tempnum = Conversion.Val("&H" + Strings.Mid(tempstr, loopit, 1));
			for (loopinner = 3; loopinner >= 0; loopinner += -1) {
				if (tempnum & Math.Pow(2, loopinner)) {
					value1 = value1 + "1";
				}
				else {
					value1 = value1 + "0";
				}
			}
		}

		tempstr = Strings.Mid(value1, rots + 1) + Strings.Left(value1, rots);

		// And convert back to hex
		value1 = "";
		for (loopit = 0; loopit <= 7; loopit++) {
			tempnum = 0;
			for (loopinner = 0; loopinner <= 3; loopinner++) {
				if (Conversion.Val(Strings.Mid(tempstr, 4 * loopit + loopinner + 1, 1))) {
					tempnum = tempnum + Math.Pow(2, (3 - loopinner));
				}
			}
			value1 = value1 + Conversion.Hex(tempnum);
		}

		return value1;

	}
	//===============================================================================
	// Name: Function BigAA1XOR
	// Input:
	//   ByVal value1 As String
	//   ByVal value2 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA1XOR(string value1, string value2)
	{
		string valueans = string.Empty;
		short loopit = 0;
		short tempnum = 0;

		tempnum = Strings.Len(value1) - Strings.Len(value2);
		if (tempnum < 0) {
			valueans = Strings.Left(value2, System.Math.Abs(tempnum));
			value2 = Strings.Mid(value2, System.Math.Abs(tempnum) + 1);
		}
		else if (tempnum > 0) {
			valueans = Strings.Left(value1, System.Math.Abs(tempnum));
			value1 = Strings.Mid(value1, tempnum + 1);
		}

		for (loopit = 1; loopit <= Strings.Len(value1); loopit++) {
			valueans = valueans + Conversion.Hex(Conversion.Val("&H" + Strings.Mid(value1, loopit, 1)) ^ Conversion.Val("&H" + Strings.Mid(value2, loopit, 1)));
		}

		return valueans;

	}
	//===============================================================================
	// Name: Function BigAA2Add
	// Input:
	//   ByVal value1 As String
	//   ByVal value2 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA2Add(string value1, string value2)
	{
		string valueans = string.Empty;
		string temps1 = null;
		string temps2 = null;
		string tempstr = null;
		int tempnum = 0;

		tempnum = Strings.Len(value1) - Strings.Len(value2);
		if (tempnum < 0) {
			value1 = new string("0", System.Math.Abs(tempnum)) + value1;
		}
		else if (tempnum > 0) {
			value2 = new string("0", tempnum) + value2;
		}

		tempnum = 0;
		while (Strings.Len(value1) > 5) {
			temps1 = Strings.Right(value1, 6);
			temps2 = Strings.Right(value2, 6);
			tempnum = tempnum + Conversion.Val("&H" + temps1 + "&") + Conversion.Val("&H" + temps2 + "&");
			tempstr = Conversion.Hex(tempnum % 16777216);
			valueans = new string("0", 6 - Strings.Len(tempstr)) + tempstr + valueans;
			tempnum = Conversion.Int(tempnum / 16777216);
			value1 = Strings.Left(value1, Strings.Len(value1) - 6);
			value2 = Strings.Left(value2, Strings.Len(value2) - 6);
		}
		tempnum = tempnum + Conversion.Val("&H" + value1 + "&") + Conversion.Val("&H" + value2 + "&");
		valueans = Conversion.Hex(tempnum % 16777216) + valueans;

		return valueans;

	}
	//===============================================================================
	// Name: Function BigAA2AND
	// Input:
	//   ByVal value1 As String
	//   ByVal value2 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA2AND(string value1, string value2)
	{
		string valueans = string.Empty;
		string tempstr = null;
		short tempnum = 0;

		tempnum = Strings.Len(value1) - Strings.Len(value2);
		if (tempnum < 0) {
			value2 = Strings.Mid(value2, System.Math.Abs(tempnum) + 1);
		}
		else if (tempnum > 0) {
			value1 = Strings.Mid(value1, tempnum + 1);
		}

		while (Strings.Len(value1) > 7) {
			tempstr = Conversion.Hex(Conversion.Val("&H" + Strings.Left(value1, 7) + "&") & Conversion.Val("&H" + Strings.Left(value2, 7) + "&"));
			valueans = valueans + new string("0", 7 - Strings.Len(tempstr)) + tempstr;
			value1 = Strings.Mid(value1, 8);
			value2 = Strings.Mid(value2, 8);
		}
		tempnum = Strings.Len(value1);
		tempstr = Conversion.Hex(Conversion.Val("&H" + value1 + "&") & Conversion.Val("&H" + value2 + "&"));
		valueans = valueans + new string("0", tempnum - Strings.Len(tempstr)) + tempstr;

		return valueans;

	}
	//===============================================================================
	// Name: Function BigAA2Mod32Add
	// Input:
	//   ByVal value1 As String
	//   ByVal value2 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA2Mod32Add(string value1, string value2)
	{
		return BigAA1Mod32Add(value1, value2);
	}
	//===============================================================================
	// Name: Function BigAA2NOT
	// Input:
	//   ByVal value1 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA2NOT(string value1)
	{
		string valueans = null;

		value1 = Strings.Right(value1, 8);
		value1 = new string("0", 8 - Strings.Len(value1)) + value1;
		valueans = Conversion.Hex(65535 ^ Conversion.Val("&H" + Strings.Right(value1, 4) + "&"));
		valueans = new string("0", 4 - Strings.Len(valueans)) + valueans;
		valueans = Conversion.Hex(65535 ^ Conversion.Val("&H" + Strings.Left(value1, 4) + "&")) + valueans;
		valueans = new string("0", 8 - Strings.Len(valueans)) + valueans;

		return valueans;

	}
	//===============================================================================
	// Name: Function BigAA2OR
	// Input:
	//   ByVal value1 As String
	//   ByVal value2 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA2OR(string value1, string value2)
	{
		string valueans = string.Empty;
		string tempstr = null;
		short tempnum = 0;

		tempnum = Strings.Len(value1) - Strings.Len(value2);
		if (tempnum < 0) {
			valueans = Strings.Left(value2, System.Math.Abs(tempnum));
			value2 = Strings.Mid(value2, System.Math.Abs(tempnum) + 1);
		}
		else if (tempnum > 0) {
			valueans = Strings.Left(value1, System.Math.Abs(tempnum));
			value1 = Strings.Mid(value1, tempnum + 1);
		}

		while (Strings.Len(value1) > 7) {
			tempstr = Conversion.Hex(Conversion.Val("&H" + Strings.Left(value1, 7) + "&") | Conversion.Val("&H" + Strings.Left(value2, 7) + "&"));
			valueans = valueans + new string("0", 7 - Strings.Len(tempstr)) + tempstr;
			value1 = Strings.Mid(value1, 8);
			value2 = Strings.Mid(value2, 8);
		}
		tempnum = Strings.Len(value1);
		tempstr = Conversion.Hex(Conversion.Val("&H" + value1 + "&") | Conversion.Val("&H" + value2 + "&"));
		valueans = valueans + new string("0", tempnum - Strings.Len(tempstr)) + tempstr;

		return valueans;

	}
	//===============================================================================
	// Name: Function BigAA2RotLeft
	// Input:
	//   ByVal value1 As String
	//   ByVal rots As Integer
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA2RotLeft(string value1, short rots)
	{
		return BigAA1RotLeft(value1, rots);
	}
	//===============================================================================
	// Name: Function BigAA2XOR
	// Input:
	//   ByVal value1 As String
	//   ByVal value2 As String
	// Output: String
	// Purpose: [INTERNAL] SHA function
	// Remarks: None
	//===============================================================================
	public static string BigAA2XOR(string value1, string value2)
	{
		string valueans = string.Empty;
		string tempstr = null;
		short tempnum = 0;

		tempnum = Strings.Len(value1) - Strings.Len(value2);
		if (tempnum < 0) {
			valueans = Strings.Left(value2, System.Math.Abs(tempnum));
			value2 = Strings.Mid(value2, System.Math.Abs(tempnum) + 1);
		}
		else if (tempnum > 0) {
			valueans = Strings.Left(value1, System.Math.Abs(tempnum));
			value1 = Strings.Mid(value1, tempnum + 1);
		}

		while (Strings.Len(value1) > 7) {
			tempstr = Conversion.Hex(Conversion.Val("&H" + Strings.Left(value1, 7) + "&") ^ Conversion.Val("&H" + Strings.Left(value2, 7) + "&"));
			valueans = valueans + new string("0", 7 - Strings.Len(tempstr)) + tempstr;
			value1 = Strings.Mid(value1, 8);
			value2 = Strings.Mid(value2, 8);
		}
		tempnum = Strings.Len(value1);
		tempstr = Conversion.Hex(Conversion.Val("&H" + value1 + "&") ^ Conversion.Val("&H" + value2 + "&"));
		valueans = valueans + new string("0", tempnum - Strings.Len(tempstr)) + tempstr;

		return valueans;

	}
	//===============================================================================
	// Name: Function SHA1AA1Hash
	// Input:
	//   ByVal hashthis As String - Input string to be hashed
	// Output: String
	// Purpose: SHA Hash function
	// Remarks: None
	//===============================================================================
	public static string SHA1AA1Hash(string hashthis)
	{
		string[] buf = new string[5];
		string[] in_ = new string[80];
		short loopouter = 0;
		short tempnum2 = 0;
		short tempnum = 0;
		short loopit = 0;
		short loopinner = 0;
		string d = null;
		string b = null;
		string a = null;
		string c = null;
		string e = null;
		string tempstr = string.Empty;

		// Add padding
		tempnum = 8 * Strings.Len(hashthis);
		hashthis = hashthis + Strings.Chr(128);
		//Add binary 10000000
		tempnum2 = 56 - Strings.Len(hashthis) % 64;
		if (tempnum2 < 0) {
			tempnum2 = 64 + tempnum2;
		}
		hashthis = hashthis + new string(Strings.Chr(0), tempnum2);
		for (loopit = 1; loopit <= 8; loopit++) {
			tempstr = Strings.Chr(tempnum % 256) + tempstr;
			tempnum = tempnum - tempnum % 256;
			tempnum = tempnum / 256;
		}
		hashthis = hashthis + tempstr;

		// Set magic numbers
		buf[0] = "67452301";
		buf[1] = "efcdab89";
		buf[2] = "98badcfe";
		buf[3] = "10325476";
		buf[4] = "c3d2e1f0";

		// For each 512 bit section
		for (loopouter = 0; loopouter <= Strings.Len(hashthis) / 64 - 1; loopouter++) {
			a = buf[0];
			b = buf[1];
			c = buf[2];
			d = buf[3];
			e = buf[4];

			// Get the 512 bits
			for (loopit = 0; loopit <= 15; loopit++) {
				in_[loopit] = "";
				for (loopinner = 4; loopinner >= 1; loopinner += -1) {
					in_[loopit] = Conversion.Hex(Strings.Asc(Strings.Mid(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_[loopit];
					if (Strings.Len(in_[loopit]) % 2) in_[loopit] = "0" + in_[loopit]; 
				}
			}

			for (loopit = 16; loopit <= 79; loopit++) {
				in_[loopit] = BigAA1RotLeft(BigAA1XOR(BigAA1XOR(BigAA1XOR(in_[loopit - 3], in_[loopit - 8]), in_[loopit - 14]), in_[loopit - 16]), 1);
			}

			for (loopit = 0; loopit <= 19; loopit++) {
				tempstr = BigAA1OR(BigAA1AND(b, c), BigAA1AND(BigAA1NOT(b), d));
				tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_[loopit], "5A827999"))));
				e = d;
				d = c;
				c = BigAA1RotLeft(b, 30);
				b = a;
				a = tempstr;
			}

			for (loopit = 20; loopit <= 39; loopit++) {
				tempstr = BigAA1XOR(BigAA1XOR(b, c), d);
				tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_[loopit], "6ED9EBA1"))));
				e = d;
				d = c;
				c = BigAA1RotLeft(b, 30);
				b = a;
				a = tempstr;
			}

			for (loopit = 40; loopit <= 59; loopit++) {
				tempstr = BigAA1OR(BigAA1OR(BigAA1AND(b, c), BigAA1AND(b, d)), BigAA1AND(c, d));
				tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_[loopit], "8F1BBCDC"))));
				e = d;
				d = c;
				c = BigAA1RotLeft(b, 30);
				b = a;
				a = tempstr;
			}

			for (loopit = 60; loopit <= 79; loopit++) {
				tempstr = BigAA1XOR(BigAA1XOR(b, c), d);
				tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_[loopit], "CA62C1D6"))));
				e = d;
				d = c;
				c = BigAA1RotLeft(b, 30);
				b = a;
				a = tempstr;
			}

			buf[0] = BigAA1Mod32Add(buf[0], a);
			buf[1] = BigAA1Mod32Add(buf[1], b);
			buf[2] = BigAA1Mod32Add(buf[2], c);
			buf[3] = BigAA1Mod32Add(buf[3], d);
			buf[4] = BigAA1Mod32Add(buf[4], e);

		}

		// Extract MD5Hash
		hashthis = "";
		for (loopit = 0; loopit <= 4; loopit++) {
			for (loopinner = 0; loopinner <= 3; loopinner++) {
				hashthis = hashthis + Strings.Mid(buf[loopit], 1 + 2 * loopinner, 2);
			}
		}

		// And return it
		return hashthis;

	}
	//===============================================================================
	// Name: Function SHA1AA2Hash
	// Input:
	//   ByVal hashthis As String - Input string to be hashed
	// Output: String
	// Purpose: SHA Hash function
	// Remarks: None
	//===============================================================================
	public static string SHA1AA2Hash(string hashthis)
	{
		string[] buf = new string[5];
		string[] in_ = new string[80];
		short loopouter = 0;
		short tempnum2 = 0;
		short tempnum = 0;
		short loopit = 0;
		short loopinner = 0;
		string d = null;
		string b = null;
		string a = null;
		string c = null;
		string e = null;
		string tempstr = string.Empty;

		// Add padding
		tempnum = 8 * Strings.Len(hashthis);
		hashthis = hashthis + Strings.Chr(128);
		//Add binary 10000000
		tempnum2 = 56 - Strings.Len(hashthis) % 64;
		if (tempnum2 < 0) {
			tempnum2 = 64 + tempnum2;
		}
		hashthis = hashthis + new string(Strings.Chr(0), tempnum2);
		for (loopit = 1; loopit <= 8; loopit++) {
			tempstr = Strings.Chr(tempnum % 256) + tempstr;
			tempnum = tempnum - tempnum % 256;
			tempnum = tempnum / 256;
		}
		hashthis = hashthis + tempstr;

		// Set magic numbers
		buf[0] = "67452301";
		buf[1] = "efcdab89";
		buf[2] = "98badcfe";
		buf[3] = "10325476";
		buf[4] = "c3d2e1f0";

		// For each 512 bit section
		for (loopouter = 0; loopouter <= Strings.Len(hashthis) / 64 - 1; loopouter++) {
			a = buf[0];
			b = buf[1];
			c = buf[2];
			d = buf[3];
			e = buf[4];

			// Get the 512 bits
			for (loopit = 0; loopit <= 15; loopit++) {
				in_[loopit] = "";
				for (loopinner = 4; loopinner >= 1; loopinner += -1) {
					in_[loopit] = Conversion.Hex(Strings.Asc(Strings.Mid(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_[loopit];
					if (Strings.Len(in_[loopit]) % 2) in_[loopit] = "0" + in_[loopit]; 
				}
			}

			for (loopit = 16; loopit <= 79; loopit++) {
				in_[loopit] = BigAA2RotLeft(BigAA2XOR(BigAA2XOR(BigAA2XOR(in_[loopit - 3], in_[loopit - 8]), in_[loopit - 14]), in_[loopit - 16]), 1);
			}

			for (loopit = 0; loopit <= 19; loopit++) {
				tempstr = BigAA2OR(BigAA2AND(b, c), BigAA2AND(BigAA2NOT(b), d));
				tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_[loopit], "5A827999"))));
				e = d;
				d = c;
				c = BigAA2RotLeft(b, 30);
				b = a;
				a = tempstr;
			}

			for (loopit = 20; loopit <= 39; loopit++) {
				tempstr = BigAA2XOR(BigAA2XOR(b, c), d);
				tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_[loopit], "6ED9EBA1"))));
				e = d;
				d = c;
				c = BigAA2RotLeft(b, 30);
				b = a;
				a = tempstr;
			}

			for (loopit = 40; loopit <= 59; loopit++) {
				tempstr = BigAA2OR(BigAA2OR(BigAA2AND(b, c), BigAA2AND(b, d)), BigAA2AND(c, d));
				tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_[loopit], "8F1BBCDC"))));
				e = d;
				d = c;
				c = BigAA2RotLeft(b, 30);
				b = a;
				a = tempstr;
			}

			for (loopit = 60; loopit <= 79; loopit++) {
				tempstr = BigAA2XOR(BigAA2XOR(b, c), d);
				tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_[loopit], "CA62C1D6"))));
				e = d;
				d = c;
				c = BigAA2RotLeft(b, 30);
				b = a;
				a = tempstr;
			}

			buf[0] = BigAA2Mod32Add(buf[0], a);
			buf[1] = BigAA2Mod32Add(buf[1], b);
			buf[2] = BigAA2Mod32Add(buf[2], c);
			buf[3] = BigAA2Mod32Add(buf[3], d);
			buf[4] = BigAA2Mod32Add(buf[4], e);

		}

		// Extract MD5Hash
		hashthis = "";
		for (loopit = 0; loopit <= 4; loopit++) {
			for (loopinner = 0; loopinner <= 3; loopinner++) {
				hashthis = hashthis + Strings.Mid(buf[loopit], 1 + 2 * loopinner, 2);
			}
		}

		// And return it
		return hashthis;

	}
}
