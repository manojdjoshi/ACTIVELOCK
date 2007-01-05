/***********************************************************************
ActiveLock
Copyright 1998-2002 Nelson Ferraz
Copyright 2003-2006 The ActiveLock Software Group (ASG)
All material is the property of the contributing authors.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are
met:

  [o] Redistributions of source code must retain the above copyright
      notice, this list of conditions and the following disclaimer.

  [o] Redistributions in binary form must reproduce the above
      copyright notice, this list of conditions and the following
      disclaimer in the documentation and/or other materials provided
      with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
"AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

***********************************************************************/
using System;
using System.IO;

namespace ALTestApp35NET_CS
{
	public class CRC32
  {
		private const int BUFFER_SIZE = 0x400;
		private int[] crc32Table;

		public CRC32()
    {
			int num1;
      int num2 = -306674912;
      this.crc32Table = new int[0x101];
      int num3 = 0;
  Label_0020:
      num1 = num3;
      int num4 = 8;
      while (true)
      {
        if ((num1 & 1) > 0)
        {
					num1 = (int) ((((long) (num1 & -2)) / 2) & 0x7fffffff);
          num1 ^= num2;
        }
        else
        {
          num1 = (int) ((((long) (num1 & -2)) / 2) & 0x7fffffff);
        }
        num4 += -1;
        if (num4 < 1)
        {
          this.crc32Table[num3] = num1;
          num3++;
          if (num3 > 0xff)
          {
						return;
          }
					goto Label_0020;
        }
      }
		}

    public int GetCrc32(ref Stream stream)
    {
			int num2 = -1;
			byte[] buffer1 = new byte[0x401];
			int num6 = 0x400;
			int num1 = stream.Read(buffer1, 0, num6);
	    
			while (num1 > 0)
			{
				int num8 = num1 - 1;
				for (int num4 = 0; num4 <= num8; num4++)
				{
					int num5 = (num2 & 0xff) ^ buffer1[num4];
					num2 = ((num2 & -256) / 0x100) & 0xffffff;
					num2 ^= this.crc32Table[num5];
				}
				num1 = stream.Read(buffer1, 0, num6);
			}
			return ~num2;
    }

  }
}

