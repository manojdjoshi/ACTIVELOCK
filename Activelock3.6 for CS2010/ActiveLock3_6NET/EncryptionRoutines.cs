using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Text;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using System.IO;

internal sealed class EncryptionRoutines
{

	#region " Private Instance Members "
	private byte[] bKey;
	private byte[] bIV;
	private bool bInitialised = false;

	private RijndaelManaged rijM = null;
	private string headerString = "CRYPTOR";

	private byte[] headerBytes = new byte[8];
		#endregion
	private bool bCancel = false;
	#region " Public Events And Enums "
	public event ProgressEventHandler Progress;
	public delegate void ProgressEventHandler(int prog);
	public event FinishedEventHandler Finished;
	public delegate void FinishedEventHandler(ReturnType retType);

	public enum ReturnType : int
	{
		Well = 0,
		Badly = 1,
		IncorrectPassword = 2
	}
	#endregion

	public string GenerateHash(string strSource)
	{
		return System.Convert.ToBase64String(new SHA384Managed().ComputeHash(new UnicodeEncoding().GetBytes(strSource)));
	}

	public void Initialise(string sPWH)
	{
		//initialise rijM
		rijM = new RijndaelManaged();
		//derive the key and IV using the 
		//PasswordDeriveBytes class
		Rfc2898DeriveBytes pdb = new Rfc2898DeriveBytes(sPWH, new MD5CryptoServiceProvider().ComputeHash(ConvertStringToBytes(sPWH)));
		//extract the key and IV
		bKey = pdb.GetBytes(32);
		bIV = pdb.GetBytes(16);

		//initialise headerBytes
		headerBytes = ConvertStringToBytes(headerString);

		{
			rijM.Key = bKey;
			//256 bit key
			rijM.IV = bIV;
			//128 bit IV
			rijM.BlockSize = 128;
			//128 bit BlockSize
			rijM.Padding = PaddingMode.PKCS7;
		}

		bInitialised = true;
	}

	public void CancelTransform()
	{
		if (!bInitialised) return; 
		bCancel = true;
	}

	public bool TransformFile(string sInFile, string sOutFile, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(true)]  // ERROR: Optional parameters aren't supported in C#
bool encrypt)
	{
		//make sure that all the initialisation has been completed:
		if (!bInitialised) {if (Finished != null) {
			Finished(ReturnType.Badly);
		}
return false;} 
		if (!IO.File.Exists(sInFile)) {if (Finished != null) {
			Finished(ReturnType.Badly);
		}
return false;} 

		FileStream fsIn = null;
		FileStream fsOut = null;
		CryptoStream encStream = null;
		ReturnType retVal = ReturnType.Badly;
		try {
			//create the input and output streams:
			fsIn = new FileStream(sInFile, FileMode.Open, FileAccess.Read);
			fsOut = new FileStream(sOutFile, FileMode.Create, FileAccess.Write);

			//some helper variables
			byte[] bBuffer = new byte[4097];
			//4KB buffer
			long lBytesRead = 0;
			long lFileSize = fsIn.Length;
			int lBytesToWrite = 0;

			if (encrypt) {
				encStream = new CryptoStream(fsOut, rijM.CreateEncryptor(bKey, bIV), CryptoStreamMode.Write);
				//write the header to the output file for use when decrypting it
				encStream.Write(headerBytes, 0, headerBytes.Length);
				//this is the main encryption routine. it loops over the input data in blocks of 4KB,
				//and writes the encrypted data to disk
				do {
					if (bCancel) break; // TODO: might not be correct. Was : Exit Try
 
					lBytesToWrite = fsIn.Read(bBuffer, 0, 4096);
					if (lBytesToWrite == 0) break; // TODO: might not be correct. Was : Exit Do
 
					encStream.Write(bBuffer, 0, lBytesToWrite);
					lBytesRead += lBytesToWrite;
					if (Progress != null) {
						Progress((int)(lBytesRead / lFileSize) * 100);
					}
				}
				while (true);
				if (Progress != null) {
					Progress(100);
				}
				retVal = ReturnType.Well;
			}
			else {
				encStream = new CryptoStream(fsIn, rijM.CreateDecryptor(bKey, bIV), CryptoStreamMode.Read);

				//read in the header
				byte[] test = new byte[headerBytes.Length + 1];
				encStream.Read(test, 0, headerBytes.Length);

				//check to see if the file header reads correctly.
				//if it doesn't, then close the stream & jump out
				if (ConvertBytesToString(test) != headerString) {
					encStream.Clear();
					encStream = null;
					retVal = ReturnType.IncorrectPassword;
					break; // TODO: might not be correct. Was : Exit Try
				}

				//this is the main decryption routine. it loops over the input data in blocks of 4KB,
				//and writes the decrypted data to disk
				do {
					if (bCancel) {
						//if the cancel flag is set,
						//then jump out
						encStream.Clear();
						encStream = null;
						break; // TODO: might not be correct. Was : Exit Try
					}
					lBytesToWrite = encStream.Read(bBuffer, 0, 4096);
					if (lBytesToWrite == 0) break; // TODO: might not be correct. Was : Exit Do
 
					fsOut.Write(bBuffer, 0, lBytesToWrite);
					lBytesRead += lBytesToWrite;
					if (Progress != null) {
						Progress((int)(lBytesRead / lFileSize) * 100);
					}
				}
				while (true);
				if (Progress != null) {
					Progress(100);
				}
				retVal = ReturnType.Well;
			}
		}
		catch (Exception ex) {
			Console.WriteLine("*****************ERROR*****************");
			Console.WriteLine(ex.ToString());
			Console.WriteLine("****************/ERROR*****************");
		}
		finally {
			//close all I/O streams (encStream first)
			if ((encStream != null)) {
				encStream.Close();
			}
			if ((fsOut != null)) {
				fsOut.Close();
			}
			if ((fsIn != null)) {
				fsIn.Close();
			}
		}
		//only delete the file if the password was bad, and
		//therefore its only an empty file
		if (retVal == ReturnType.IncorrectPassword) {
			IO.File.Delete(sOutFile);
		}
		//raise the Finished event, and then reset bCancel
		if (Finished != null) {
			Finished(retVal);
		}
		bCancel = false;
	}

	public byte[] ConvertStringToBytes(string sString)
	{
		return new UnicodeEncoding().GetBytes(sString);
	}

	public string ConvertBytesToString(byte[] bytes)
	{
		return new UnicodeEncoding().GetString(bytes);
	}
}
