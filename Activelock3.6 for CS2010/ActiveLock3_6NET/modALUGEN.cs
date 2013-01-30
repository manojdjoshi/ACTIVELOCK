using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
using System.Runtime.InteropServices;
static class modALUGEN
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
	// Name: modActiveLock
	// Purpose: Module used by ALUGEN.
	// Functions:
	// Properties:
	// Methods:
	// Started: 04.21.2005
	// Modified: 03.25.2006
	//===============================================================================
	// @author activelock-admins
	// @version 3.3.0
	// @date 03.25.2006
	//
	// RSA Key Structure
	// @param bits   Key length in bits
	// @param data   Key data
	[StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi)]
	public struct RSAKey
	{
		// 36-byte structure
		public int bits;
		//<VBFixedArray(32)> Dim data() As Byte
		[MarshalAs(UnmanagedType.ByValArray, SizeConst = 32)]
		public byte[] data;
		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

		}
	}
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int rsa_generate(ref RSAKey ptrKey, int bits, int pfn, int pfnparam);
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int rsa_generate2(ref RSAKey ptrKey, int bits);
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int rsa_public_key_blob(ref RSAKey ptrKey, string blob, ref int blobLen);
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int rsa_private_key_blob(ref RSAKey ptrKey, string blob, ref int blobLen);
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int rsa_createkey(string pub_blob, int pub_len, string priv_blob, int priv_len, ref RSAKey ptrKey);
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int rsa_freekey(ref RSAKey ptrKey);
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int rsa_sign(ref RSAKey ptrKey, string data, int dLen, string sig, ref int sLen);
	[DllImport("ALCrypto3NET", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int rsa_verifysig(ref RSAKey ptrKey, string sig, int sLen, string data, int dLen);

	// Generates an RSA key set
	// @param ptrKey RSA key structure
	// @param bits   key length in bits
	// @param pfn    TBD
	// @param pfnparam    TBD

	// Generates an RSA key set without showing any progress
	// needed for the VB.NET build
	// @param ptrKey RSA key structure
	// @param bits   key length in bits

	// Returns the public key blob for the specified key.
	// @param ptrKey RSA key structure
	// @param blob   [Output] Key bob to be returned
	// @param blobLen    Length of the key blob, in bytes

	// Returns the private key blob for the specified key.
	// @param ptrKey RSA key structure
	// @param blob   [Output] Key bob to be returned
	// @param blobLen    Length of the key blob, in bytes

	// Creates a new RSAKey from the specified key blobs.
	// @param pub_blob   Public key blob
	// @param pub_len    Length of public key blob, in bytes
	// @param priv_blob  Private key blob
	// @param priv_len   Length of private key blob, in bytes
	// @param ptrKey     [out] RSA key to be returned.

	// Release memory allocated by rsa_createkey() to store the key.
	// @param ptrKey     RSA key

	// Signs the data using the specified RSA private key.
	// @param ptrKey Key to be used for signing
	// @param data   Data to be signed
	// @param dLen   Data length
	// @param sig    [out] Signature
	// @param sLen   Signature length
	//
	// Verifies an RSA signature.
	// @param ptrKey Key to be used for signing
	// @param sig    [out] Signature
	// @param sLen   Signature length
	// @param data   Data with which to verify
	// @param dLen   Data length

	public struct PhaseType
	{
		public byte exponential;
		public byte startpoint;
		public byte total;
		public byte param;
		public byte current;
			// if exponential */
		public byte N;
			// if linear */
		public byte Mult;
	}


	public const int MAXPHASE = 5;
	public struct ProgressType
	{
		public int nphases;
		[VBFixedArray(modALUGEN.MAXPHASE - 1)]
		public PhaseType[] phases;
		public byte total;
		public byte divisor;
		public byte range;
		public int hwndProgbar;
		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

		}
	}
	public static string strLeft(string vString, int vLength)
	{
		return vString.Substring(0, vLength);
	}
	public static string strRight(string vString, int vLength)
	{
		return vString.Substring(vString.Length - vLength);
	}

}
