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
using System.Text;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace ALTestApp35NET_CS
{

	internal class modMain
	{
		internal const string PUB_KEY = "2CB.2CB.2CB.2CB.2D6.231.35A.53E.42B.2E1.21B.533.441.226.2F7.2CB.2CB.2CB.2CB.2D6.32E.37B.2CB.2CB.2CB.323.2D6.268.205.2D6.226.339.3BD.4C5.42B.483.226.3BD.391.30D.39C.386.370.441.46D.4AF.34F.4C5.441.53E.457.3C8.4D0.44C.268.4BA.512.210.3D3.23C.4E6.21B.4F1.32E.21B.51D.3B2.231.512.318.226.21B.4DB.23C.4E6.39C.4D0.2F7.3D3.507.2D6.483.2EC.23C.318.302.365.4D0.499.436.35A.2D6.391.386.44C.4D0.2D6.318.32E.30D.3BD.457.441.25D.48E.3A7.483.268.323.391.3B2.210.4D0.34F.252.483.226.339.53E.4BA.48E.478.2E1.4AF.4F1.247.2E1.2F7.4FC.3D3.318.386.533.436.436.483.3D3.512.386.3C8.4A4.457.30D.53E.302.4F1.2CB.2CB.370.268.21B.25D.370.344.35A.231.32E.3D3.4C5.231.3BD.499.2F7.4E6.39C.226.4C5.462.386.247.386.2E1.499.462.478.4AF.528.210.252.210.2D6.39C.268.51D.42B.370.4C5.4DB.4BA.4BA.231.2CB.2D6.25D.4F1.3DE.210.37B.29F.29F";
		internal const string MSGBOXCAPTION = "TestApp 1.0";
		internal const string SOFTWARENAME = "TestApp";

		internal static int Asc(string mystring)
		{
			//Return the character value of the given character
			return System.Text.Encoding.ASCII.GetBytes(mystring)[0];
		}
		
		internal static char Chr(int i)
		{
			//Return the character of the given character value
			return Convert.ToChar(i);
		}

		internal static bool isNumeric(string subject) 
		{ 
			return(System.Text.RegularExpressions.Regex.IsMatch(subject,"^[-0-9]*.[.0-9].[0-9]*$")); 
		} 

		internal static string ReadFile(string FileName)
		{
			StreamReader file = File.OpenText(FileName);
			string s = file.ReadToEnd();
			file.Close();
			return s;
		}

		internal static string Dec(string strdata) 
		{ 
			string[] arr; 
			arr = strdata.Split(Convert.ToChar(".")); 
			StringBuilder mStringBulder = new StringBuilder(null);
			for (int i = 0; i <= arr.Length-1; i++) 
			{ 
				mStringBulder.Append(Chr(int.Parse(arr[i], System.Globalization.NumberStyles.HexNumber)/ 11));
			} 
			return mStringBulder.ToString();
		}


		internal static string Enc(ref string strdata)
		{
			string _enc = null;
			int strDataLenght = strdata.Length;
			for (int Index = 0; Index <= strDataLenght-1; Index++)
			{
				int numValue = Asc(strdata.Substring(Index, 1)) * 11;
				if (_enc.Length == 0)
				{
					_enc = numValue.ToString("X"); 
				}
				else
				{
					_enc = _enc + "." + numValue.ToString("X");
				}
			}
			return _enc;
		}

		internal static string VerifyActiveLockNETdll()
		{
			string crcSumHEX=null;
			string strActiveLockCrcSum=null;
			string strMessage=null;
			CRC32 crcObj = new CRC32();
			int crcSum = 0;
			string activelockFile = Application.StartupPath + @"\Activelock3_5NET.dll";
			if (File.Exists(activelockFile))
			{
				FileStream streamFileActiveLock = new FileStream(activelockFile, FileMode.Open, FileAccess.Read, FileShare.Read, 0x2000);
				Stream streamActiveLock = streamFileActiveLock;
				crcSum = crcObj.GetCrc32(ref streamActiveLock);
				streamFileActiveLock.Close();
				crcSumHEX = string.Format("{0:X8}", crcSum);

				Debug.WriteLine("Hash: " + crcSum.ToString());

				strActiveLockCrcSum = "25D.252.23C.2D6.2D6.2EC.273.302"; //crcSum of Activelock
				if (String.Compare(crcSumHEX, modMain.Dec(strActiveLockCrcSum), true)!=0)
				{
					strMessage = "42B.441.4FC.483.512.457.4A4.4C5.441.499.231.35A.2F7.39C.1FA.44C.4A4.4A4.160.478.42B.4F1.160.436.457.457.4BA.160.441.4C5.4E6.4E6.507.4D0.4FC.457.44C.1FA";
					MessageBox.Show(modMain.Dec(strMessage), modMain.MSGBOXCAPTION, MessageBoxButtons.OK);
					Application.Exit();
				}
				return crcSumHEX;
			}
			strMessage = "42B.441.4FC.483.512.457.4A4.4C5.441.499.231.35A.2F7.39C.1FA.44C.4A4.4A4.160.4BA.4C5.4FC.160.462.4C5.507.4BA.44C";
			MessageBox.Show(modMain.Dec(strMessage), modMain.MSGBOXCAPTION, MessageBoxButtons.OK, MessageBoxIcon.Information);
			Application.Exit();
			return crcSumHEX;
		}

		internal static string AppPath()
		{
			return Application.StartupPath;
		}

		internal static string WindowsRootFolder() 
		{ 
			return Environment.GetEnvironmentVariable("SystemRoot");
		}

		internal static string WindowsSystemFolder() 
		{ 
			return Environment.SystemDirectory;
		}

	}
}

