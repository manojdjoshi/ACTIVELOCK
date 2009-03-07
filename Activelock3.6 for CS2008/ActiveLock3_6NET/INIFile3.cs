using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

 // ERROR: Not supported in C#: OptionDeclaration
internal class INIFile
{
	[DllImport("kernel32", EntryPoint = "GetPrivateProfileStringA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetPrivateProfileString(string lpApplicationName, string lpKeyName, string lpDefault, string lpReturnedString, int nSize, string lpFileName);
	[DllImport("kernel32", EntryPoint = "WritePrivateProfileStringA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int WritePrivateProfileString(string lpApplicationName, string lpKeyName, string lpString, string lpFileName);
	[DllImport("kernel32", EntryPoint = "GetPrivateProfileSectionA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetPrivateProfileSection(string lpAppName, string lpReturnedString, int nSize, string lpFileName);
	[DllImport("kernel32", EntryPoint = "WritePrivateProfileSectionA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int WritePrivateProfileSection(string lpAppName, string lpString, string lpFileName);
	[DllImport("kernel32", EntryPoint = "GetPrivateProfileSectionNamesA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetPrivateProfileSectionNames(string lpszReturnBuffer, int nSize, string lpFileName);
	//*   ActiveLock
	//*   Copyright 1998-2002 Nelson Ferraz
	//*   Copyright 2006 The ActiveLock Software Group (ASG)
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
	// Name: INIFile
	// Purpose: Stores and retrieves product keys
	// <p>An "object-oriented" approach to using Windows INI files, with some
	//   useful additions.
	// <p>Klaus H. Probst [kprobst@vbbox.com]
	// Functions:
	// Properties:
	// Methods:
	// Started: 07.07.2003
	// Modified: 03.24.2006
	//===============================================================================

	//  ///////////////////////////////////////////////////////////////////////
	//  / Filename:  INIFile.cls                                              /
	//  / Version:   1.0.0.1                                                  /
	//  / Purpose:   Stores and retrieves product keys                         /
	//  / Klaus H. Probst [kprobst@vbbox.com]                                 /
	//  /                                                                     /
	//  / Date Created:         ???? ??, ???? - KHP                           /
	//  / Date Last Modified:   July 07, 2003 - MEC                           /
	//  /                                                                     /
	//  / This software is released under the license detailed below and is   /
	//  / subject to said license. Neither this header nor the licese below   /
	//  / may be removed from this module.                                    /
	//  ///////////////////////////////////////////////////////////////////////

	//Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long

	private string m_sFileName;
	private string m_sSection;

		// Last error message
	private string m_LastError;

	private const int MAX_BUFFER_SIZE = 4096;
	//===============================================================================
	// Name: Function GetValue
	// Input:
	//   ByVal ValueName As String - Passed key
	//   ByVal Default As String - Default return value to be used if the call fails
	//   ByRef Failed As Boolean - This will be set to False if everything went well, or True of something went wrong
	// Output:
	//   Variant - This will be set to False if everything went well, or True of something went wrong, and the return is
	// the value passed in the Default argument.
	// Purpose: Retrieves a value from the passed key (ValueName) and returns it as a variant
	// (String subtype). This proc is useful if your requirements go above that of the
	// Values Get/Let pair. You can specify a Default return value in case the call fails,
	// and you can pass a Boolean variable in the Fail argument.
	// Remarks: None
	//===============================================================================
	//Default was upgraded to Default_Renamed
	public object GetValue(string ValueName, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute("")]  // ERROR: Optional parameters aren't supported in C#
string Default_Renamed, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(false)] ref  // ERROR: Optional parameters aren't supported in C#
bool Failed)
	{
		object functionReturnValue = null;
		string sBuffer = null;
		int lReturn = 0;

		m_LastError = "";
		// reset error message

		sBuffer = new string(Constants.vbNullChar, MAX_BUFFER_SIZE - 1);
		lReturn = GetPrivateProfileString(m_sSection, ValueName, Default_Renamed, sBuffer, Strings.Len(sBuffer), m_sFileName);

		if (lReturn > 0) {
			Failed = false;
			functionReturnValue = Strings.Left(sBuffer, lReturn);
		}
		else {
			Failed = true;
			m_LastError = modActiveLock.WinError(lReturn);

			functionReturnValue = Default_Renamed;
		}
		return functionReturnValue;

	}
	//===============================================================================
	// Name: Sub AddSection
	// Input:
	//   ByVal name As String - Section name to be added
	//   ByRef SetAsCurrent As Boolean - If True, sets the added section as the current section name
	// Output: None
	// Purpose: Adds an empty section to the current file. If you're wondering
	// why this is here, know that I use it a lot, but I suppose it's
	// of limited value in most cases.
	// Remarks: None
	//===============================================================================
	public void AddSection(string name, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(false)] ref  // ERROR: Optional parameters aren't supported in C#
bool SetAsCurrent)
	{
		 // ERROR: Not supported in C#: OnErrorStatement

		//// to add an empty section, we have to write a dummy
		//// key into it (which automatically creates the section), and
		//// then delete the dummy key. This leaves the section intact
		WritePrivateProfileString(name, "Dummy", "Val", m_sFileName);
		WritePrivateProfileString(name, "Dummy", 0, m_sFileName);

		if (SetAsCurrent == true) m_sSection = name; 
	}
	//===============================================================================
	// Name: Function EnumSectionKeys
	// Input:
	//   ByRef ArrayResult As Variant - Returned array
	// Output:
	//   Integer - Returns 0 if failure
	// Purpose: Enumerates the keys (not the Key-Value pairs) under the current section,
	// copies the array of keys into the ArrayResult argument, and returns the
	// number of keys enumerated.
	// <p>For how the class handles buffer sizes on INI calls, please see the Notes
	// section on the [Declarations] section of the class module.
	// <p>If the method fails, the return will be zero, the ArrayResult argument
	// will be set to Null and you will hit an assert.
	// Remarks: None
	//===============================================================================
	public short EnumSectionKeys(ref object ArrayResult)
	{
		short functionReturnValue = 0;
		string sBuffer = null;
		int lReturn = 0;
		string sValuePair = null;
		short iCounter = 0;
		string[] asEnumKeys = null;
		int lBuffSize = 0;
		short iDelim = 0;
		object avReturn = null;
		short iTries = 0;

		 // ERROR: Not supported in C#: OnErrorStatement

		iCounter = -1;
		//// preset this
		lBuffSize = MAX_BUFFER_SIZE;
		//// start out with 4096 bytes
		do {
			lBuffSize = (lBuffSize * 2);
			lBuffSize = lBuffSize - 1;
			//// otherwise we overflow 32767
			sBuffer = new string(Constants.vbNullChar, lBuffSize);
			lReturn = GetPrivateProfileSection(m_sSection, sBuffer, lBuffSize, m_sFileName);
			if (lReturn == (lBuffSize - 2)) {
				iTries = iTries + 1;
				lBuffSize = lBuffSize + 1;
				//// readjust to pair boundary
			}
			else {
				break; // TODO: might not be correct. Was : Exit Do
			}
			sBuffer = "";
			//// clear if this is not the first time
		}
		while (iTries < 4);

		if (Strings.Len(sBuffer) > 3) {
			//// Trim the resulting buffer
			sBuffer = Strings.Left(sBuffer, lReturn + 1);
			do {
				if (iCounter == -1) {
					iCounter = 0;
					asEnumKeys = new string[1];
				}
				else {
					Array.Resize(ref asEnumKeys, iCounter + 1);
				}

				//pull up to the first Null
				sValuePair = Strings.Mid(sBuffer, 1, Strings.InStr(sBuffer, Constants.vbNullChar) - 1);
				if (Strings.Len(sValuePair) > 0) {
					iDelim = Strings.InStr(sValuePair, "=");
					if (iDelim > 0) {
						asEnumKeys[iCounter] = Strings.Left(sValuePair, iDelim - 1);
					}

				}

				//trim the string to the next null
				sBuffer = Strings.Mid(sBuffer, Strings.InStr(sBuffer, Constants.vbNullChar) + 1);
				iCounter = iCounter + 1;
			}
			while (Strings.Len(sBuffer) > 3);
		}
		finally_Renamed:
		//// See if we have success
		if (iCounter > -1) {
			ArrayResult = asEnumKeys;
			// VB6.CopyArray(asEnumKeys)
		}
		else {
			ArrayResult = System.DBNull.Value;
		}
		functionReturnValue = iCounter;
		return;
		catch_Renamed:
		//// Ah, hello :)
		System.Diagnostics.Debug.Assert(0, "");
		iCounter = -1;
		 // ERROR: Not supported in C#: ResumeStatement

		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function EnumSectionValues
	// Input:
	//   ByRef ArrayResult As Variant - Returned array
	// Output:
	//   Integer - Returns 0 if failure
	// Purpose: This proc will enumerate just the values contained under a given section
	// and will return them in an array.
	// For how the class handles buffer sizes on INI calls, please see the Notes
	// section on the [Declarations] section of the class module.
	// If the method fails, the return will be zero, the ArrayResult argument
	// will be set to Null and you will hit an assert.
	// Remarks: None
	//===============================================================================
	public short EnumSectionValues(ref object ArrayResult)
	{
		short functionReturnValue = 0;

		string sBuffer = null;
		int lReturn = 0;
		string sValuePair = null;
		short iCounter = 0;
		string[] asEnumValues = null;
		int lBuffSize = 0;
		short iDelim = 0;
		short iTries = 0;

		//Debug.Assert 0
		 // ERROR: Not supported in C#: OnErrorStatement

		iCounter = -1;
		lBuffSize = MAX_BUFFER_SIZE;
		//// start out with 4096 bytes
		do {
			lBuffSize = (lBuffSize * 2);
			lBuffSize = lBuffSize - 1;
			//// otherwise we overflow 32767
			sBuffer = new string(Constants.vbNullChar, lBuffSize);
			lReturn = GetPrivateProfileSection(m_sSection, sBuffer, lBuffSize, m_sFileName);
			if (lReturn == (lBuffSize - 2)) {
				iTries = iTries + 1;
				lBuffSize = lBuffSize + 1;
				//// readjust to pair boundary
			}
			else {
				break; // TODO: might not be correct. Was : Exit Do
			}
			sBuffer = "";
			//// clear if this is not the first time
		}
		while (iTries < 4);
		if (((Strings.Len(sBuffer) > 0) & (lReturn != 0))) {
			sBuffer = Strings.Left(sBuffer, lReturn - 1);
			sBuffer = sBuffer + Constants.vbNullChar;
			do {
				if (iCounter == -1) {
					iCounter = 0;
					asEnumValues = new string[1];
				}
				else {
					Array.Resize(ref asEnumValues, iCounter + 1);
				}
				//pull up to the first Null
				sValuePair = Strings.Mid(sBuffer, 1, Strings.InStr(sBuffer, Constants.vbNullChar) - 1);
				if (Strings.Len(sValuePair) > 0) {
					//Parse the string
					iDelim = Strings.InStr(sValuePair, "=");
					if (iDelim > 0) {
						asEnumValues[iCounter] = Strings.Right(sValuePair, Strings.Len(sValuePair) - iDelim);
					}
				}

				//trim the string to the next null
				sBuffer = Strings.Mid(sBuffer, Strings.InStr(sBuffer, Constants.vbNullChar) + 1);
				iCounter = iCounter + 1;
			}
			while (Strings.Len(sBuffer) > 3);
		}
		finally_Renamed:

		//// See if we have success
		if (iCounter > -1) {
			ArrayResult = asEnumValues;
			// VB6.CopyArray(asEnumValues)
		}
		else {
			ArrayResult = System.DBNull.Value;
		}
		functionReturnValue = iCounter;
		return;
		catch_Renamed:
		//// Ah, hello :)
		System.Diagnostics.Debug.Assert(0, "");
		iCounter = -1;
		 // ERROR: Not supported in C#: ResumeStatement

		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function EnumSections
	// Input:
	//   ByRef ArrayResult As Variant - Returned array
	// Output:
	//   Integer - Returns 0 if failure
	// Purpose: This proc will enumerate all the names of the sections of the current INI
	// file and return them in an array.
	// <p>For how the class handles buffer sizes on INI calls, please see the Notes
	// section on the [Declarations] section of the class module.
	// <p>If the method fails, the return will be zero, the ArrayResult argument
	// will be set to Null and you will hit an assert.
	// Remarks: None
	//===============================================================================
	public short EnumSections(ref string[] ArrayResult)
	{
		short functionReturnValue = 0;
		 // ERROR: Not supported in C#: OnErrorStatement

		string sBuffer = null;
		int lReturn = 0;
		string sValue = null;
		short iCounter = 0;
		string[] asEnumSections = null;
		int lBuffSize = 0;
		short iDelim = 0;
		short iTries = 0;

		//Debug.Assert 0
		 // ERROR: Not supported in C#: OnErrorStatement

		iCounter = -1;
		lBuffSize = MAX_BUFFER_SIZE;
		//// start out with 4096 bytes
		do {
			lBuffSize = (lBuffSize * 2);
			lBuffSize = lBuffSize - 1;
			//// otherwise we overflow 32767
			sBuffer = new string(Constants.vbNullChar, lBuffSize);
			lReturn = GetPrivateProfileSectionNames(sBuffer, lBuffSize, m_sFileName);
			if (lReturn == (lBuffSize - 2)) {
				iTries = iTries + 1;
				lBuffSize = lBuffSize + 1;
				//// readjust to pair boundary
			}
			else {
				break; // TODO: might not be correct. Was : Exit Do
			}
			sBuffer = "";
			//// clear if this is not the first time
		}
		while (iTries < 4);

		if (((Strings.Len(sBuffer) > 0) & (lReturn != 0))) {
			sBuffer = Strings.Left(sBuffer, lReturn + 1);
			//// Must allow for an extra NULL
			do {
				if (iCounter == -1) {
					iCounter = 0;
					asEnumSections = new string[1];
				}
				else {
					Array.Resize(ref asEnumSections, iCounter + 1);
				}

				//pull up to the first Null
				sValue = Strings.Mid(sBuffer, 1, Strings.InStr(sBuffer, Constants.vbNullChar) - 1);
				if (Strings.Len(sValue) > 0) {
					asEnumSections[iCounter] = sValue;

				}

				//trim the string to the next null
				sBuffer = Strings.Mid(sBuffer, Strings.InStr(sBuffer, Constants.vbNullChar) + 1);
				iCounter = iCounter + 1;
			}
			while (Strings.Len(sBuffer) > 3);
		}
		else {
			functionReturnValue = 0;
			return;
		}
		finally_Renamed:

		//// See if we have success
		if (iCounter > -1) {
			ArrayResult = asEnumSections;
			// VB6.CopyArray(asEnumSections)
		}
		else {
			ArrayResult = null;
			//System.DBNull.Value
		}

		functionReturnValue = iCounter;
		return;
		catch_Renamed:
		//// Ah, hello :)
		System.Diagnostics.Debug.Assert(0, "");
		iCounter = -1;
		 // ERROR: Not supported in C#: ResumeStatement

		return functionReturnValue;
	}
	//===============================================================================
	// Name: Sub DeleteSection
	// Input:
	//   ByVal SectionName As String - Name of the section to be deleted
	// Output: None
	// Purpose: Deletes a section from the current INI file
	// Remarks: None
	//===============================================================================
	public void DeleteSection(string SectionName)
	{
		WritePrivateProfileString(SectionName, 0, 0, m_sFileName);
	}
	//===============================================================================
	// Name: Sub DeleteKey
	// Input:
	//   ByVal KeyName As String - Name of the key to be deleted
	// Output: None
	// Purpose: Deletes a key (a value pair) from the INI file
	// Remarks: None
	//===============================================================================
	public void DeleteKey(string KeyName)
	{
		WritePrivateProfileString(m_sSection, KeyName, 0, m_sFileName);
	}
	//===============================================================================
	// Name: Function EnumSectionValuePairs
	// Input:
	//   ByRef ArrayResult As Variant - Returned array
	// Output:
	//   Integer - Returns 0 if failure
	// Purpose: This proc will enumerate a given section's Key=Value and place them
	// in the passed array as two different arrays. That is:
	// <p>    Array(Array(Keys),Array(Values))
	// <p>For how the class handles buffer sizes on INI calls, please see the Notes
	// section on the [Declarations] section of the class module.
	// <p>If the method fails, the return will be zero, the ArrayResult argument
	// will be set to Null and you will hit an assert.
	// Remarks: None
	//===============================================================================
	public short EnumSectionValuePairs(ref object ArrayResult)
	{
		short functionReturnValue = 0;
		 // ERROR: Not supported in C#: OnErrorStatement


		string sBuffer = null;
		int lReturn = 0;
		string sValuePair = null;
		short iCounter = 0;
		string[] arrEnumValues = null;
		string[] arrEnumKeys = null;
		int lBuffSize = 0;
		short iDelim = 0;
		short iTries = 0;

		//Debug.Assert 0
		 // ERROR: Not supported in C#: OnErrorStatement

		iCounter = -1;
		lBuffSize = MAX_BUFFER_SIZE;
		//// start out with 4096 bytes
		do {
			lBuffSize = (lBuffSize * 2);
			lBuffSize = lBuffSize - 1;
			//// otherwise we overflow 32767
			sBuffer = new string(Constants.vbNullChar, lBuffSize);
			lReturn = GetPrivateProfileSection(m_sSection, sBuffer, lBuffSize, m_sFileName);
			if (lReturn == (lBuffSize - 2)) {
				iTries = iTries + 1;
				lBuffSize = lBuffSize + 1;
				//// readjust to pair boundary
			}
			else {
				break; // TODO: might not be correct. Was : Exit Do
			}
			sBuffer = "";
			//// clear if this is not the first time
		}
		while (iTries < 4);

		if (((Strings.Len(sBuffer) > 0) & (lReturn != 0))) {
			sBuffer = Strings.Left(sBuffer, lReturn);
			do {
				if (iCounter == -1) {
					iCounter = 0;
					arrEnumValues = new string[1];
					arrEnumKeys = new string[1];
				}
				else {
					//top-redim the array
					Array.Resize(ref arrEnumKeys, iCounter + 1);
					Array.Resize(ref arrEnumValues, iCounter + 1);
				}
				//pull up to the first Null
				sValuePair = Strings.Mid(sBuffer, 1, Strings.InStr(sBuffer, Constants.vbNullChar) - 1);
				if (Strings.Len(sValuePair) > 0) {
					//Parse the string
					iDelim = Strings.InStr(sValuePair, "=");
					if (iDelim > 0) {
						arrEnumKeys[iCounter] = Strings.Left(sValuePair, iDelim - 1);
						arrEnumValues[iCounter] = Strings.Right(sValuePair, Strings.Len(sValuePair) - iDelim);
					}
				}

				//trim the string to the next null
				sBuffer = Strings.Mid(sBuffer, Strings.InStr(sBuffer, Constants.vbNullChar) + 1);
				iCounter = iCounter + 1;
			}
			while (Strings.Len(sBuffer) > 3);
		}
		finally_Renamed:

		//// See if we have success
		if (iCounter > -1) {
			ArrayResult = new object[] { arrEnumKeys, arrEnumValues };
		}
		else {
			ArrayResult = new object[];
			// return empty array
		}
		functionReturnValue = iCounter;
		return;
		catch_Renamed:
		//// Ah, hello :)
		System.Diagnostics.Debug.WriteLine(Err().Number, Err().Description);
		System.Diagnostics.Debug.Assert(0, "");
		iCounter = -1;
		 // ERROR: Not supported in C#: ResumeStatement

		return functionReturnValue;
	}
	//===============================================================================
	// Name: Sub Flush
	// Input: None
	// Output: None
	// Purpose: Flushes the INI file cache for the current file. <p>Note that
	// the INI cache in Win32 (and Win16) is notoriously kranky. This flush
	// is a logical one, not a physical flush. So if you need this type of
	// thing seriously use the registry instead.
	// Remarks: None
	//===============================================================================
	public void Flush()
	{
		WritePrivateProfileString(0, 0, 0, m_sFileName);
	}
	//===============================================================================
	// Name: Property Get Values
	// Input:
	//   ByVal ValueName As String - Key name
	// Output:
	//   Variant - Key value
	// Purpose: Retrieves a value from a key as a variant (string subtype) from the current section.
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Let Values
	// Input:
	//   ByVal ValueName As String - Key name
	//   ByVal rValue As Variant - Key value
	// Output: None
	// Purpose: Sets the value of a key in the current section. Note that the passed value
	// (a variant) is converted to a string before saving. This means that, for
	// example, if you're passing a boolean value and you don't want a '-1' converted
	// to a "True" literal, you'll have to convert the boolean to an integer or
	// long before calling this method.
	// Remarks: None
	//===============================================================================
	public object Values {
		get {
			string szBuffer = null;
			int lReturn = 0;
			string sDefault = null;

			szBuffer = new string(Constants.vbNullChar, MAX_BUFFER_SIZE - 1);
			sDefault = "";
			lReturn = GetPrivateProfileString(m_sSection, ValueName, sDefault, szBuffer, Strings.Len(szBuffer), m_sFileName);

			if (lReturn > 0) {
				Values = Strings.Left(szBuffer, lReturn);
			}
			else {
				Values = sDefault;
			}
		}
		set { WritePrivateProfileString(m_sSection, ValueName, (string)value, m_sFileName); }
	}
	//===============================================================================
	// Name: Property Get File
	// Input: None
	// Output:
	//   String - Current file name
	// Purpose: Returns the current filename
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Let File
	// Input:
	//   ByVal rValue As String - File name used
	// Output: None
	// Purpose: Sets the filename that all methods will use
	// by default. Note that if the passed string is not a FQP,
	// the INI file must be in the Windows search path somewhere.
	// Remarks: None
	//===============================================================================
	public string File {
		get { File = m_sFileName; }
		set {
			// TODO: Check if File exists.
			// If file not found: 
			// Set_locale(regionalSymbol)
			// Err.Raise(vbObjectError, , "File not found: " & rValue)
			m_sFileName = value;
			m_sSection = "";
			//// clear the section
		}
	}
	//===============================================================================
	// Name: Property Get Section
	// Input: None
	// Output:
	//   String - Name of the current section
	// Purpose: Retrieves the name of the current section
	// Remarks: None
	//===============================================================================
	//===============================================================================
	// Name: Property Let Section
	// Input:
	//    ByVal rValue As String - Current section name
	// Output: None
	// Purpose: Sets the name of the current section
	// Remarks: None
	//===============================================================================
	public string Section {
		get { Section = m_sSection; }
		set { m_sSection = value; }
	}
	//===============================================================================
	// Name: Sub WriteSection
	// Input:
	//   ByVal rValue As Variant - variant array
	// Output: None
	// Purpose: Accepts a variant array and writes the contents to
	// the current section. Note that this will overwrite
	// all of the existing key-value pairs under that section.
	// The array must be structured as follows:
	// <p>    rValue(n) = "KeyName = Value"
	// <p>where n is a given index of the array. The "=" literal between the
	// KeyName and Value *must* be present, or the call will fail.
	// Remarks: None
	//===============================================================================
	public void WriteSection(object rValue)
	{
		string sBuffer = string.Empty;
		int a = 0;

		if (!Information.IsArray(rValue)) return;
 

		for (a = Information.LBound(rValue); a <= Information.UBound(rValue); a++) {
			sBuffer = sBuffer + rValue(a) + Constants.vbNullChar;
		}

		sBuffer = sBuffer + new string(Constants.vbNullChar, 2);
		WritePrivateProfileSection(m_sSection, sBuffer, m_sFileName);
	}
}
