using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
static class modBase64
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
	// Name: modBase64
	// Purpose: This module contains Base-64 encoding and decoding routines.
	// Functions:
	// Properties:
	// Methods:
	// Started: 04.21.2005
	// Modified: 03.25.2006
	//===============================================================================
	// @author activelock-admins
	// @version 3.3.0
	// @date 03.25.2006


	private const string base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
	//===============================================================================
	// Name: Function Base64_Encode
	// Input:
	//   ByRef DecryptedText As String - The decrypted string
	// Output:
	//   String - Base64 encoded string
	// Purpose: Return the Base64 encoded string
	// Remarks: None
	//===============================================================================
	public static string Base64_Encode(ref string DecryptedText)
	{
		int c1 = 0;
		int c2 = 0;
		int c3 = 0;
		short w1 = 0;
		short w2 = 0;
		short w3 = 0;
		short w4 = 0;
		short N = 0;
		string retry = string.Empty;

		for (N = 1; N <= Strings.Len(DecryptedText); N += 3) {
			c1 = Strings.Asc(Strings.Mid(DecryptedText, N, 1));
			c2 = Strings.Asc(Strings.Mid(DecryptedText, N + 1, 1) + Strings.Chr(0));
			c3 = Strings.Asc(Strings.Mid(DecryptedText, N + 2, 1) + Strings.Chr(0));
			w1 = Conversion.Int(c1 / 4);
			w2 = (short)c1 & 3 * 16 + Conversion.Int(c2 / 16);
			if (Strings.Len(DecryptedText) >= N + 1) w3 = (short)c2 & 15 * 4 + Conversion.Int(c3 / 64); 			else w3 = -1; 
			if (Strings.Len(DecryptedText) >= N + 2) w4 = c3 & 63; 			else w4 = -1; 
			retry = retry + mimeencode(ref w1) + mimeencode(ref w2) + mimeencode(ref w3) + mimeencode(ref w4);
		}
		return retry;
	}

	//===============================================================================
	// Name: Function Base64_Decode
	// Input:
	//   ByRef a As String - The string to be decoded
	// Output:
	//   String - Base64 decoded string
	// Purpose: Return the Base64 decoded string
	// Remarks: None
	//===============================================================================
	public static string Base64_Decode(ref string a)
	{
		short w1 = 0;
		short w2 = 0;
		short w3 = 0;
		short w4 = 0;
		short N = 0;
		string retry = string.Empty;

		for (N = 1; N <= Strings.Len(a); N += 4) {
			w1 = mimedecode(ref Strings.Mid(a, N, 1));
			w2 = mimedecode(ref Strings.Mid(a, N + 1, 1));
			w3 = mimedecode(ref Strings.Mid(a, N + 2, 1));
			w4 = mimedecode(ref Strings.Mid(a, N + 3, 1));
			if (w2 >= 0) retry = retry + Strings.Chr((w1 * 4 + Conversion.Int(w2 / 16)) & 255); 
			if (w3 >= 0) retry = retry + Strings.Chr((w2 * 16 + Conversion.Int(w3 / 4)) & 255); 
			if (w4 >= 0) retry = retry + Strings.Chr((w3 * 64 + w4) & 255); 
		}
		return retry;
	}

	//===============================================================================
	// Name: Function mimeencode
	// Input:
	//   ByRef w As Integer - Input integer
	// Output:
	//   String
	// Purpose: Used by the Base64_encode function
	// Remarks: None
	//===============================================================================
	private static string mimeencode(ref short w)
	{
		string functionReturnValue = null;
		if (w >= 0) functionReturnValue = Strings.Mid(base64, w + 1, 1); 		else functionReturnValue = ""; 
		return functionReturnValue;
	}

	//===============================================================================
	// Name: Function mimedecode
	// Input:
	//   ByRef a As String - Input string
	// Output:
	//   Integer
	// Purpose: Used by the Base64_decode function
	// Remarks: None
	//===============================================================================
	private static short mimedecode(ref string a)
	{
		if (Strings.Len(a) == 0) {mimedecode() = -1;return;
} 
		return Strings.InStr(base64, a) - 1;
	}
}
