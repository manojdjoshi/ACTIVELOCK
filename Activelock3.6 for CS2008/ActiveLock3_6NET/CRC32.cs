using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
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
// This is v2 of the VB CRC32 algorithm provided by Paul
// (wpsjr1@succeed.net) - much quicker than the nasty
// original version I posted.  Excellent work!

public class CRC32
{
	private int[] crc32Table;

	private const int BUFFER_SIZE = 1024;

	public int GetCrc32(ref System.IO.Stream stream)
	{
		int crc32Result = 0;
		crc32Result = 0xffffffff;

		byte[] buffer = new byte[BUFFER_SIZE + 1];
		int readSize = BUFFER_SIZE;

		int count = stream.Read(buffer, 0, readSize);
		int i = 0;
		int iLookup = 0;
		int tot = 0;
		while ((count > 0)) {
			for (i = 0; i <= count - 1; i++) {
				iLookup = (crc32Result & 0xff) ^ buffer[i];
				crc32Result = ((crc32Result & 0xffffff00) / 0x100) & 0xffffff;
				// nasty shr 8 with vb :/
				crc32Result = crc32Result ^ crc32Table(iLookup);
			}
			count = stream.Read(buffer, 0, readSize);
		}

		return !(crc32Result);

	}

	public CRC32()
	{
		// This is the official polynomial used by CRC32 in PKZip.
		// Often the polynomial is shown reversed (04C11DB7).
		int dwPolynomial = 0xedb88320;
		int i = 0;
		int j = 0;

		 // ERROR: Not supported in C#: ReDimStatement

		int dwCrc = 0;

		for (i = 0; i <= 255; i++) {
			dwCrc = i;
			for (j = 8; j >= 1; j += -1) {
				if ((dwCrc & 1) == 0) {
					dwCrc = Convert.ToInt32(((dwCrc & 0xfffffffe) / 2L) & 0x7fffffff);
					dwCrc = dwCrc ^ dwPolynomial;
				}
				else {
					dwCrc = Convert.ToInt32(((dwCrc & 0xfffffffe) / 2L) & 0x7fffffff);
				}
			}
			crc32Table(i) = dwCrc;
		}
	}

}
