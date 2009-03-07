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
using System.Text;
using System.Security.Cryptography;
static class modMD5
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
	// Name: modMD5
	// Purpose: MD5 Hashing Module
	// Date Created: 
	// Date Last Modified:
	// Functions:
	// Properties:
	// Methods:
	// Started: 06.16.1998
	// Modified: 03.25.2006
	//===============================================================================
	//===============================================================================
	// Name: Function Hash
	// Input:
	//   ByRef strMessage As String - String to be hashed
	// Output:
	//   String - Hashed string
	// Purpose: MD5 Hash function
	// Remarks: None
	//===============================================================================
	public static string Hash(ref string strMessage)
	{
		//Create an encoding object to ensure the encoding standard for the source text
		UnicodeEncoding Ue = new UnicodeEncoding();
		//Retrieve a byte array based on the source text
		byte[] ByteSourceText = Ue.GetBytes(strMessage);
		//Instantiate an MD5 Provider object
		MD5CryptoServiceProvider Md5 = new MD5CryptoServiceProvider();
		//Compute the hash value from the source
		byte[] ByteHash = Md5.ComputeHash(ByteSourceText);
		//And convert it to String format for return
		return Convert.ToBase64String(ByteHash);
	}
	//===============================================================================
	// Name: Function ComputeHash
	// Input:
	//   ByRef strMessage As String - String to be hashed
	// Output:
	//   String - Hashed string
	// Purpose: MD5 Hash function
	// Remarks: This function is primarily used by the Short Key Function 
	// to hash strings; it matches the hash generated in the VB6 version
	// therefore both .NET and VB6 versions generate the same serial number and key
	//===============================================================================
	public static string ComputeHash(ref string strMessage)
	{
		byte[] hashedDataBytes = null;
		UTF7Encoding encoder = new UTF7Encoding();
		MD5CryptoServiceProvider md5Hasher = new MD5CryptoServiceProvider();
		hashedDataBytes = md5Hasher.ComputeHash(encoder.GetBytes(strMessage));
		string strHash = BitConverter.ToString(hashedDataBytes);
		return strHash.Replace("-", "").ToLower();
	}

}
