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
static class modRegistry
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
	// Name: modRegistryAPIs
	// Purpose: Facilitates Windows registry access
	// Date Last Modified:   July 07, 2003 - MEC
	// Functions:
	// Properties:
	// Methods:
	// Started: 07.07.2003
	// Modified: 03.25.2006
	//===============================================================================

		// Unicode nul terminated string
	public const short REG_SZ = 1;
	public const short REG_EXPAND_SZ = 2;
	public const short REG_BINARY = 3;
	public const short REG_DWORD = 4;

	public const int HKEY_CLASSES_ROOT = 0x80000000;
	public const int HKEY_CURRENT_USER = 0x80000001;
	public const int HKEY_LOCAL_MACHINE = 0x80000002;
	public const int HKEY_USERS = 0x80000003;
	public const int HKEY_PERFORMANCE_DATA = 0x80000004;

	public const int HKEY_CURRENT_CONFIG = 0x80000005;
	public const int HKEY_DYN_DATA = 0x80000006;

	public const short REG_OPTION_NON_VOLATILE = 0;
	public const short REG_CREATED_NEW_KEY = 0x1;
	public const short REG_OPENED_EXISTING_KEY = 0x2;

	public const short KEY_QUERY_VALUE = 0x1;
	public const short KEY_ENUMERATE_SUB_KEYS = 0x8;
	public const short KEY_NOTIFY = 0x10;
	public const int READ_CONTROL = 0x20000;
	public const int STANDARD_RIGHTS_ALL = 0x1f0000;
	public const int STANDARD_RIGHTS_EXECUTE = (READ_CONTROL);
	public const int STANDARD_RIGHTS_READ = (READ_CONTROL);
	public const int STANDARD_RIGHTS_REQUIRED = 0xf0000;
	public const int SYNCHRONIZE = 0x100000;
	public const bool KEY_READ = ((STANDARD_RIGHTS_READ | KEY_QUERY_VALUE | KEY_ENUMERATE_SUB_KEYS | KEY_NOTIFY) & (!SYNCHRONIZE));
	public const short KEY_SET_VALUE = 0x2;
	public const short KEY_CREATE_SUB_KEY = 0x4;
	public const short KEY_CREATE_LINK = 0x20;
	public const int STANDARD_RIGHTS_WRITE = (READ_CONTROL);
	public const bool KEY_WRITE = ((STANDARD_RIGHTS_WRITE | KEY_SET_VALUE | KEY_CREATE_SUB_KEY) & (!SYNCHRONIZE));

	public const bool KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL | KEY_QUERY_VALUE | KEY_SET_VALUE | KEY_CREATE_SUB_KEY | KEY_ENUMERATE_SUB_KEYS | KEY_NOTIFY | KEY_CREATE_LINK) & (!SYNCHRONIZE));
	public const short ERROR_SUCCESS = 0;
	public const short ERROR_ACCESS_DENIED = 5;
	public const short ERROR_NO_MORE_ITEMS = 259;
	public const short ERROR_BADKEY = 1010;
	public const short ERROR_CANTOPEN = 1011;
	public const short ERROR_CANTREAD = 1012;
	public const short ERROR_REGISTRY_CORRUPT = 1015;

	public struct SECURITY_ATTRIBUTES
	{
		public int nLength;
		public int lpSecurityDescriptor;
		public bool bInheritHandle;
	}

	public struct FILETIME
	{
		public int dwLowDateTime;
		public int dwHighDateTime;
	}
	[DllImport("advapi32.dll", EntryPoint = "RegOpenKeyExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegOpenKeyEx(int hKey, string lpSubKey, int ulOptions, int samDesired, ref int phkResult);
	[DllImport("advapi32.dll", EntryPoint = "RegQueryValueExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegQueryValueEx(int hKey, string lpValueName, int lpReserved, ref int lpType, ref int lpData, ref int dwSize);
	[DllImport("advapi32", EntryPoint = "RegCreateKeyExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegCreateKeyEx(int hKey, string lpSubKey, int Reserved, string lpClass, int dwOptions, int samDesired, ref SECURITY_ATTRIBUTES lpSecurityAttributes, ref int phkResult, ref int lpdwDisposition);
	[DllImport("advapi32.dll", EntryPoint = "RegSetValueExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegSetValueEx(int hKey, string lpValueName, int dwReserved, int dwType, ref int lpValue, int dwSize);
	[DllImport("advapi32.dll", EntryPoint = "RegDeleteKeyA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegDeleteKey(int hKey, string lpSubKey);
	[DllImport("advapi32.dll", EntryPoint = "RegDeleteValueA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegDeleteValue(int hKey, string lpValueName);
	[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegCloseKey(int hKey);
	[DllImport("advapi32.dll", EntryPoint = "RegConnectRegistryA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegConnectRegistry(string lpMachineName, int hKey, ref int phkResult);
	[DllImport("advapi32.dll", EntryPoint = "RegCreateKeyA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegCreateKey(int hKey, string lpSubKey, ref int phkResult);
	[DllImport("advapi32.dll", EntryPoint = "RegEnumKeyA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegEnumKey(int hKey, int dwIndex, string lpName, int cbName);
	[DllImport("advapi32.dll", EntryPoint = "RegEnumValueA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegEnumValue(int hKey, int dwIndex, string lpValueName, ref int lpcbValueName, ref int lpReserved, ref int lpType, ref byte lpData, ref int lpcbData);
	[DllImport("advapi32.dll", EntryPoint = "RegEnumKeyExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegEnumKeyEx(int hKey, int dwIndex, string lpName, ref int lpcbName, int lpReserved, string lpClass, ref int lpcbClass, ref FILETIME lpftLastWriteTime);
	[DllImport("advapi32.dll", EntryPoint = "RegLoadKeyA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegLoadKey(int hKey, string lpSubKey, string lpFile);
	[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegNotifyChangeKeyValue(int hKey, int bWatchSubtree, int dwNotifyFilter, int hEvent, int fAsynchronus);
	[DllImport("advapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegOpenKey(int hKey, string lpSubKey, ref int phkResult);
	[DllImport("advapi32.dll", EntryPoint = "RegQueryValueA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int OSRegQueryValue(int hKey, string lpSubKey, string lpValue, ref int lpcbValue);
	[DllImport("advapi32.dll", EntryPoint = "RegReplaceKeyA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegReplaceKey(int hKey, string lpSubKey, string lpNewFile, string lpOldFile);
	[DllImport("advapi32.dll", EntryPoint = "RegRestoreKeyA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegRestoreKey(int hKey, string lpFile, int dwFlags);
	[DllImport("advapi32.dll", EntryPoint = "RegQueryInfoKeyA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int RegQueryInfoKey(int hKey, string lpClass, ref int lpcbClass, int lpReserved, ref int lpcSubKeys, ref int lpcbMaxSubKeyLen, ref int lpcbMaxClassLen, ref int lpcValues, ref int lpcbMaxValueNameLen, ref int lpcbMaxValueLen, 
	ref int lpcbSecurityDescriptor, ref FILETIME lpftLastWriteTime);

	//Structure FILETIME may require marshalling attributes to be passed as an argument in this Declare statement.
	//===============================================================================
	// Name: Function DeleteRegKey
	// Input:
	//   ByRef lngKey As Long - HKEY
	//   ByRef SubKey As String - Sub key name
	// Output:
	//   Long - Return value from the RegDeleteKey function
	// Purpose: Deletes a registry key
	// Remarks: None
	//===============================================================================
	public static int DeleteRegKey(ref int lngKey, ref string SubKey)
	{
		int lngRet = 0;
		lngRet = RegDeleteKey(lngKey, SubKey);
		return lngRet;
	}
	//===============================================================================
	// Name: Function DeleteRegValue
	// Input:
	//   ByRef lngKey As Long - HKEY
	//   ByRef SubKey As String - Sub key name
	//   ByRef ValueName As String - Value name
	// Output:
	//   Long - Return value from the RegDeleteValue function
	// Purpose: Deletes a registry value
	// Remarks: None
	//===============================================================================
	public static int DeleteRegValue(ref int lngKey, ref string SubKey, ref string ValueName)
	{
		int functionReturnValue = 0;
		int lngRet = 0;
		int lngKeyRet = 0;

		lngRet = RegOpenKeyEx(lngKey, SubKey, 0, KEY_WRITE, ref lngKeyRet);
		if (lngRet != ERROR_SUCCESS) return;
 

		lngRet = RegDeleteValue(lngKeyRet, ValueName);
		functionReturnValue = lngRet;
		RegCloseKey(lngKeyRet);
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function WriteRegLong
	// Input:
	//   ByRef lngKey As Long - HKEY
	//   ByRef SubKey As String - Sub key name
	//   ByRef DataName As String - Value name
	//   ByRef DataValue As Long - Key value
	// Output: Long
	// Purpose: Writes a long key value in the registry
	// Remarks: None
	//===============================================================================
	public static int WriteRegLong(ref int lngKey, ref string SubKey, ref string DataName, ref int DataValue)
	{

		SECURITY_ATTRIBUTES SEC = default(SECURITY_ATTRIBUTES);
		int lngKeyRet = 0;
		int lngDis = 0;
		int lngRet = 0;

		lngRet = RegCreateKeyEx(lngKey, SubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ref SEC, ref lngKeyRet, ref lngDis);

		if ((lngRet == ERROR_SUCCESS) | (lngRet == REG_CREATED_NEW_KEY) | (lngRet == REG_OPENED_EXISTING_KEY)) {
			lngRet = RegSetValueEx(lngKeyRet, DataName, 0, REG_DWORD, ref DataValue, 4);
			RegCloseKey(lngKeyRet);
		}
		return lngRet;
	}
	//===============================================================================
	// Name: Function WriteStringValue
	// Input:
	//   ByRef lngKey As Long - HKEY
	//   ByRef SubKey As String - Sub key name
	//   ByRef DataName As String - Value name
	//   ByRef DataValue As String - Key value
	// Output: Long
	// Purpose: Writes a string in the registry
	// Remarks: None
	//===============================================================================
	public static int WriteStringValue(ref int lngKey, ref string SubKey, ref string DataName, ref string DataValue)
	{

		SECURITY_ATTRIBUTES SEC = default(SECURITY_ATTRIBUTES);
		int lngKeyRet = 0;
		int lngDis = 0;
		int lngRet = 0;

		lngRet = RegCreateKeyEx(lngKey, SubKey, 0, Constants.vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, ref SEC, ref lngKeyRet, ref lngDis);
		if (DataValue <= "") DataValue = ""; 
		if ((lngRet == ERROR_SUCCESS) | (lngRet == REG_CREATED_NEW_KEY) | (lngRet == REG_OPENED_EXISTING_KEY)) {
			lngRet = RegSetValueEx(lngKeyRet, DataName, 0, REG_SZ, ref DataValue, Strings.Len(DataValue));
			RegCloseKey(lngKeyRet);
		}
		return lngRet;
	}
	//===============================================================================
	// Name: Function ReadRegVal
	// Input:
	//   ByRef lngKey As Long - HKEY
	//   ByRef SubKey As String - Sub key name
	//   ByRef DataName As String - Value name
	//   ByRef DefaultData As Variant - Default value to be returned
	// Output: None
	// Purpose: Reads a key value from the registry
	// Remarks: None
	//===============================================================================
	public static object ReadRegVal(ref int lngKey, ref string SubKey, ref string DataName, ref object DefaultData)
	{
		object functionReturnValue = null;
		int lngKeyRet = 0;
		int lngData = 0;
		string strdata = null;
		int datatype = 0;
		int DataSize = 0;
		int lngRet = 0;
		functionReturnValue = DefaultData;
		lngRet = RegOpenKeyEx(lngKey, SubKey, 0, KEY_QUERY_VALUE, ref lngKeyRet);
		if (lngRet != ERROR_SUCCESS) return;
 
		lngRet = RegQueryValueEx(lngKeyRet, DataName, 0, ref datatype, ref 0, ref DataSize);
		if (lngRet != ERROR_SUCCESS) {
			RegCloseKey(lngKeyRet);
			return;
		}
		switch (datatype) {
			case REG_SZ:
				strdata = Strings.Space(DataSize + 1);
				lngRet = RegQueryValueEx(lngKeyRet, DataName, 0, ref datatype, ref strdata, ref DataSize);
				if (lngRet == ERROR_SUCCESS) {
					functionReturnValue = (object)StripNulls(Strings.RTrim(strdata));
				}

				break;
			case REG_DWORD:
				lngRet = RegQueryValueEx(lngKeyRet, DataName, 0, ref datatype, ref lngData, ref 4);
				if (lngRet == ERROR_SUCCESS) {
					functionReturnValue = (object)lngData;
				}

				break;
		}
		RegCloseKey(lngKeyRet);
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function GetSubKeys
	// Input:
	//   ByRef strKey As String - Key name
	//   ByRef SubKey As String - Sub key name
	//   ByRef SubKeyCnt As Long - Number of keys
	// Output:
	//   String - Sub key list of a given key (separated by commas)
	// Purpose: Gets subkeys from a given key separated by commas
	// Remarks: None
	//===============================================================================
	public static string GetSubKeys(ref string strKey, ref string SubKey, ref int SubKeyCnt)
	{
		string[] strValues = null;
		string strTemp = string.Empty;
		int lngSub = 0;
		short intCnt = 0;
		int lngRet = 0;
		short intKeyCnt = 0;
		FILETIME FT = default(FILETIME);

		GetSubKeys() = string.Empty;
		lngRet = RegOpenKeyEx((int)strKey, SubKey, 0, KEY_ENUMERATE_SUB_KEYS, ref lngSub);

		if (lngRet != ERROR_SUCCESS) {
			SubKeyCnt = 0;
			return;
		}
		lngRet = RegQueryInfoKey(lngSub, Constants.vbNullString, ref 0, 0, ref SubKeyCnt, ref 65, ref 0, ref 0, ref 0, ref 0, 
		ref 0, ref FT);
		if ((lngRet != ERROR_SUCCESS) | (SubKeyCnt <= 0)) {
			SubKeyCnt = 0;
		}
		strValues = new string[SubKeyCnt];
		for (intCnt = 0; intCnt <= SubKeyCnt - 1; intCnt++) {
			strValues[intCnt] = new string(Strings.Chr(0), 65);
			RegEnumKeyEx(lngSub, intCnt, strValues[intCnt], ref 65, 0, Constants.vbNullString, ref 0, ref FT);
			strValues[intCnt] = StripNulls(strValues[intCnt]);
		}
		RegCloseKey(lngSub);
		for (intKeyCnt = Information.LBound(strValues); intKeyCnt <= Information.UBound(strValues); intKeyCnt++) {
			strTemp = strTemp + strValues[intKeyCnt] + ",";
		}
		return strTemp;
	}
	//===============================================================================
	// Name: Function StripNulls
	// Input:
	//   ByVal s As String - Input string
	// Output:
	//   String - Returned string free of nulls
	// Purpose: Strips nulls in a given string
	// Remarks: None
	//===============================================================================
	public static string StripNulls(string s)
	{
		string functionReturnValue = null;
		short i = 0;
		i = Strings.InStr(s, Strings.Chr(0));
		if (i > 0) {
			functionReturnValue = Strings.Left(s, i - 1);
		}
		else {
			functionReturnValue = s;
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function ParseString
	// Input:
	//   ByRef strIn As String - Input string
	//   ByRef intLoc As Integer - Character location
	//   ByRef strDelimiter As String - String delimiter
	// Output:
	//   String - Parsed string
	// Purpose: String parser
	// Remarks: None
	//===============================================================================
	public static string ParseString(ref string strIn, ref short intLoc, ref string strDelimiter)
	{
		short intPos = 0;
		short intStrt = 0;
		short intStop = 0;
		short intCnt = 0;
		intCnt = intLoc;
		while (intCnt > 0) {
			intStop = intPos;
			intStrt = Strings.InStr(intPos + 1, strIn, Strings.Left(strDelimiter, 1));
			if (intStrt > 0) {
				intPos = intStrt;
				intCnt = intCnt - 1;
			}
			else {
				intPos = Strings.Len(strIn) + 1;
				break; // TODO: might not be correct. Was : Exit Do
			}
		}
		return Strings.Mid(strIn, intStop + 1, intPos - intStop - 1);
	}
	//===============================================================================
	// Name: Sub alSaveSetting
	// Input:
	//   ByRef strRegHive As String, ByRef strRegPath As String, ByRef strAppname As String, ByRef strSection As String, ByRef strKey As String, ByRef vData As Variant
	//   ByRef strRegHive As String - Base registry class
	//   ByRef strRegPath As String - Registry key path under "Software"
	//   ByRef strAppname As String - Application name
	//   ByRef strSection As String - Section name
	//   ByRef strKey As String - Key name
	//   ByRef vDefault As Variant - Key value
	// Output: None
	// Purpose: Saves a key in the registry. Calls the alSaveSettingReg sub.
	// Remarks: None
	//===============================================================================
	public static void alSaveSetting(ref string strRegHive, ref string strRegPath, ref string strAppname, ref string strSection, ref string strKey, ref object vData)
	{
		alSaveSettingReg(ref strRegHive, ref strRegPath, ref strAppname, ref strSection, ref strKey, ref vData);
	}

	//===============================================================================
	// Name: Function alGetSetting
	// Input:
	//   ByRef strRegHive As String - Base registry class
	//   ByRef strRegPath As String - Registry key path under "Software"
	//   ByRef strAppname As String - Application name
	//   ByRef strSection As String - Section name
	//   ByRef strKey As String - Key name
	//   ByRef vDefault As Variant - Key value
	// Output: Variant
	// Purpose: Reads a key value from the registry. Calls the alGetSettingReg function to get a registry value
	// Remarks: None
	//===============================================================================
	public static object alGetSetting(ref string strRegHive, ref string strRegPath, ref string strAppname, ref string strSection, ref string strKey, ref object vDefault)
	{
		return alGetSettingReg(ref strRegHive, ref strRegPath, ref strAppname, ref strSection, ref strKey, ref vDefault);
	}
	//===============================================================================
	// Name: Sub alSaveSettingReg
	// Input:
	//   ByRef strRegHive As String - Base registry class
	//   ByRef strRegPath As String - Registry key path under "Software"
	//   ByRef strAppname As String - Application name
	//   ByRef strSection As String - Section name
	//   ByRef strKey As String - Key name
	//   ByRef vData As Variant - Key value
	// Output: None
	// Purpose: Saves a key in the registry.
	// Remarks: None
	//===============================================================================
	public static void alSaveSettingReg(ref string strRegHive, ref string strRegPath, ref string strAppname, ref string strSection, ref string strKey, ref object vData)
	{
		int lRegistryBase = 0;
		switch (Strings.Left(Strings.UCase(strRegHive), 4)) {
			case "HKLM":
				lRegistryBase = HKEY_LOCAL_MACHINE;
				break;
			case "HKCR":
				lRegistryBase = HKEY_CLASSES_ROOT;
				break;
			default:
				lRegistryBase = HKEY_CURRENT_USER;
				break;
		}
		WriteStringValue(ref lRegistryBase, ref "Software\\" + strRegPath + "\\" + strAppname + "\\" + strSection, ref strKey, ref (string)vData);
	}
	//===============================================================================
	// Name: Function alGetSettingReg
	// Input:
	//   ByRef strRegHive As String - Base registry class
	//   ByRef strRegPath As String - Registry key path under "Software"
	//   ByRef strAppname As String - Application name
	//   ByRef strSection As String - Section name
	//   ByRef strKey As String - Key name
	//   ByRef vDefault As Variant - Key value
	// Output:
	//   Variant - Return value from the ReadRegVal function
	// Purpose: Reads a key value from the registry
	// Remarks: None
	//===============================================================================
	public static object alGetSettingReg(ref string strRegHive, ref string strRegPath, ref string strAppname, ref string strSection, ref string strKey, ref object vDefault)
	{
		int lRegistryBase = 0;
		switch (Strings.Left(Strings.UCase(strRegHive), 4)) {
			case "HKLM":
				lRegistryBase = HKEY_LOCAL_MACHINE;
				break;
			case "HKCR":
				lRegistryBase = HKEY_CLASSES_ROOT;
				break;
			default:
				lRegistryBase = HKEY_CURRENT_USER;
				break;
		}
		return ReadRegVal(ref lRegistryBase, ref "Software\\" + strRegPath + "\\" + strAppname + "\\" + strSection, ref strKey, ref vDefault);
	}
	//===============================================================================
	// Name: Sub Savekey
	// Input:
	//   ByVal hKey As Long - HKEY
	//   ByVal strPath As String - Key Name
	// Output: None
	// Purpose: Saves a key in the registry
	// Remarks: None
	//===============================================================================
	public static void Savekey(ref int hKey, ref string strPath)
	{
		int keyhand = 0;
		int r = 0;
		r = RegCreateKey(hKey, strPath, ref keyhand);
		r = RegCloseKey(keyhand);
	}
	//===============================================================================
	// Name: Function GetString
	// Input:
	//   ByRef hKey As Long - HKEY
	//   ByRef strPath As String - Key Name
	//   ByRef strValue As String - Value Name
	// Output:
	//   String
	// Purpose: Gets a string from the registry
	// Remarks:  EXAMPLE:<br>
	//   text1.text = getstring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String")
	//===============================================================================
	public static string GetString(ref int hKey, ref string strPath, ref string strValue)
	{
		string functionReturnValue = null;
		//EXAMPLE:
		//
		//text1.text = getstring(HKEY_CURRENT_USE
		//     R, "Software\VBW\Registry", "String")
		//
		int keyhand = 0;
		int r = 0;
		int lValueType = 0;
		int lResult = 0;
		string strBuf = null;
		int lDataBufSize = 0;
		short intZeroPos = 0;

		functionReturnValue = string.Empty;

		r = RegOpenKey(hKey, strPath, ref keyhand);
		lResult = RegQueryValueEx(keyhand, strValue, 0, ref lValueType, ref 0, ref lDataBufSize);

		if (lValueType == REG_SZ) {
			strBuf = new string(" ", lDataBufSize);
			lResult = RegQueryValueEx(keyhand, strValue, 0, ref 0, ref strBuf, ref lDataBufSize);
			if (lResult == ERROR_SUCCESS) {
				intZeroPos = Strings.InStr(strBuf, Strings.Chr(0));
				if (intZeroPos > 0) {
					functionReturnValue = Strings.Left(strBuf, intZeroPos - 1);
				}
				else {
					functionReturnValue = strBuf;
				}
			}
		}
		if (string.IsNullOrEmpty(functionReturnValue)) {
			functionReturnValue = "Value or key does not exist";
		}
		return functionReturnValue;

	}
	//===============================================================================
	// Name: Function GetRegistryValue
	// Input:
	//   ByVal hKey As Long - HKEY
	//   ByVal KeyName As String - Key Name
	//   ByVal ValueName As String - Value Name
	//   ByRef defaultValue as Variant - Default value to be returned if the value is missing
	// Output:
	//   Variant - Registry value
	// Purpose: Gets a key value from the registry
	// Remarks: None
	//===============================================================================
	public static object GetRegistryValue(int hKey, string KeyName, string ValueName, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(null)] ref  // ERROR: Optional parameters aren't supported in C#
object defaultValue)
	{
		object functionReturnValue = null;
		int handle = 0;
		int resLong = 0;
		string resString = null;
		byte[] resBinary = null;
		int Length = 0;
		int retVal = 0;
		int valueType = 0;

		// Prepare the default result
		functionReturnValue = ((defaultValue == null) ? null : defaultValue);

		// Open the key, exit if not found.
		if (RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, ref handle)) {
			return;
		}

		// prepare a 1K receiving resBinary
		Length = 1024;
		resBinary = new byte[Length];

		// read the registry key
		retVal = RegQueryValueEx(handle, ValueName, 0, ref valueType, ref resBinary[0], ref Length);
		// if resBinary was too small, try again
		if (retVal == modTrial.ERROR_MORE_DATA) {
			// enlarge the resBinary, and read the value again
			resBinary = new byte[Length];
			retVal = RegQueryValueEx(handle, ValueName, 0, ref valueType, ref resBinary[0], ref Length);
		}

		// return a value corresponding to the value type
		switch (valueType) {
			case REG_DWORD:
				modHardware.CopyMemory(ref resLong, resBinary[0], 4);
				functionReturnValue = resLong;
				break;
			case REG_SZ:
			case REG_EXPAND_SZ:
				// copy everything but the trailing null char
				resString = Strings.Space(Length - 1);
				modHardware.CopyMemory(ref resString, resBinary[0], Length - 1);
				functionReturnValue = resString;
				break;
			case REG_BINARY:
				// resize the result resBinary
				if (Length != Information.UBound(resBinary) + 1) {
					Array.Resize(ref resBinary, Length);
				}

				functionReturnValue = resBinary;
				// VB6.CopyArray(resBinary)
				break;
			case modTrial.REG_MULTI_SZ:
				// copy everything but the 2 trailing null chars
				resString = Strings.Space(Length - 2);
				modHardware.CopyMemory(ref resString, resBinary[0], Length - 2);
				functionReturnValue = resString;
				break;
			default:
				RegCloseKey(handle);
				modActiveLock.Set_Locale(modActiveLock.regionalSymbol);
				Err().Raise(1001, modTrial.ACTIVELOCKSTRING, "Unsupported value type");
				break;
		}

		// close the registry key
		RegCloseKey(handle);
		return functionReturnValue;
	}

	//===============================================================================
	// Name: Function SetRegistryValue
	// Input:
	//   ByVal hKey As Long - HKEY
	//   ByVal KeyName As String - Key Name
	//   ByVal ValueName As String - Value Name
	//   ByRef Value As Variant - Key Value.
	//   Value can be an integer value (REG_DWORD), a string (REG_SZ) or an array of binary (REG_BINARY). Raises an error otherwise.
	// Output:
	//   Boolean - True if successful
	// Purpose: Writes or Creates a Registry value
	// Remarks: Use KeyName = "" for the default value
	//===============================================================================
	public static bool SetRegistryValue(int hKey, string KeyName, string ValueName, ref object Value)
	{
		const int KEY_WRITE = 0x20006;
		//((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or
		// KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

		// Write or Create a Registry value
		// returns True if successful
		//
		// Use KeyName = "" for the default value
		//
		// Value can be an integer value (REG_DWORD), a string (REG_SZ)
		// or an array of binary (REG_BINARY). Raises an error otherwise.
		int handle = 0;
		int lngValue = 0;
		string strValue = null;
		byte[] binValue = null;
		int Length = 0;
		int retVal = 0;

		// Open the key, exit if not found
		if (RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, ref handle)) {
			return;
		}

		// three cases, according to the data type in Value
		switch (Information.VarType(Value)) {
			case VariantType.Short:
			case VariantType.Integer:
				lngValue = Value;
				retVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, ref lngValue, 4);
				break;
			case VariantType.String:
				strValue = Value;
				retVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ref strValue, Strings.Len(strValue));
				break;
			case VariantType.Array + VariantType.Byte:
				binValue = Value;
				Length = Information.UBound(binValue) - Information.LBound(binValue) + 1;
				retVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, ref binValue[Information.LBound(binValue)], Length);
				break;
			default:
				RegCloseKey(handle);
				modActiveLock.Set_Locale(modActiveLock.regionalSymbol);
				Err().Raise(1001, modTrial.ACTIVELOCKSTRING, "Unsupported value type");
				break;
		}

		// Close the key and signal success
		RegCloseKey(handle);
		// signal success if the value was written correctly
		return (retVal == 0);
	}






	//===============================================================================
	// Name: Function SaveString
	// Input:
	//   ByRef hKey As Long - HKEY
	//   ByRef strPath As String - Key Name
	//   ByRef strValue As String - Value Name
	//   ByRef strdata As String - Key Value
	// Output:
	//   Variant - Returns "Success" if successful
	// Purpose: Saves a string in the registry
	// Remarks:  EXAMPLE:<br>
	//   text1.text= savestring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String", text1.text)
	//===============================================================================
	public static object SaveString(ref int hKey, ref string strPath, ref string strValue, ref string strdata)
	{
		object functionReturnValue = null;
		//EXAMPLE:
		//
		//text1.text= savestring(HKEY_CURRENT_USER, "Sof
		//     tware\VBW\Registry", "String", text1.tex
		//     t)
		//
		int keyhand = 0;
		int r = 0;
		r = RegCreateKey(hKey, strPath, ref keyhand);
		r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ref strdata, Strings.Len(strdata));
		r = RegCloseKey(keyhand);
		if (r == 0) {
			functionReturnValue = "Success";
		}
		else {
			functionReturnValue = "Key to Delete Or Key Not Found";
		}
		return functionReturnValue;

	}



	//===============================================================================
	// Name: Function Getdword
	// Input:
	//   ByVal hKey As Long - HKEY
	//   ByVal strPath As String - Key Name
	//   ByVal strValueName As String - Value Name
	// Output:
	//   Variant - Returns the DWORD if successful
	// Purpose: Gets the DWORD of a key from the registry
	// Remarks: EXAMPLE:<br>
	//   text1.text = getdword(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
	//===============================================================================
	public static object Getdword(int hKey, string strPath, string strValueName)
	{
		object functionReturnValue = null;
		//EXAMPLE:
		//
		//text1.text = getdword(HKEY_CURRENT_USER
		//     , "Software\VBW\Registry", "Dword")
		//
		int lResult = 0;
		int lValueType = 0;
		int lBuf = 0;
		int lDataBufSize = 0;
		int r = 0;
		int keyhand = 0;

		functionReturnValue = string.Empty;

		r = RegOpenKey(hKey, strPath, ref keyhand);
		// Get length/data type
		lDataBufSize = 4;
		lResult = RegQueryValueEx(keyhand, strValueName, 0, ref lValueType, ref lBuf, ref lDataBufSize);

		if (lResult == ERROR_SUCCESS) {
			if (lValueType == REG_DWORD) {
				functionReturnValue = lBuf;
			}
		}
		r = RegCloseKey(keyhand);
		if (string.IsNullOrEmpty(functionReturnValue)) {
			functionReturnValue = " Value or key does not exist";
		}
		return functionReturnValue;

	}



	//===============================================================================
	// Name: Function SaveDword
	// Input:
	//   ByVal hKey As Long - HKEY
	//   ByVal strPath As String - Key Name
	//   ByVal strValueName As String - Value Name
	//   ByVal lData As Long - Key Value
	// Output:
	//   Variant - Returns "Success" if successful
	// Purpose: None
	// Remarks: None
	//===============================================================================
	// Saves a DWORD in the registry
	// @param hKey           HKEY
	// @param strPath        Key Name
	// @param strValueName   Value Name
	// @param lData          Value
	// @return               "Success" if success
	//
	public static object SaveDword(int hKey, string strPath, string strValueName, int lData)
	{
		object functionReturnValue = null;
		//EXAMPLE"
		//
		//Text1.text= SaveDword(HKEY_CURRENT_USER, "Soft
		//     ware\VBW\Registry", "Dword", text1.text)
		//
		//
		int lResult = 0;
		int keyhand = 0;
		int r = 0;
		r = RegCreateKey(hKey, strPath, ref keyhand);
		lResult = RegSetValueEx(keyhand, strValueName, 0, REG_DWORD, ref lData, 4);
		//If lResult <> error_success Then Call e
		//     rrlog("SetDWORD", False)
		r = RegCloseKey(keyhand);
		if (r == 0) {
			functionReturnValue = "Success";
		}
		else {
			functionReturnValue = " Failed to save Value";
		}
		return functionReturnValue;

	}

	//===============================================================================
	// Name: Function DeleteKey
	// Input:
	//   ByVal hKey As Long - HKEY
	//   ByVal strKey As String - Key Name
	// Output:
	//   Variant - Returns "Success" if successful
	// Purpose: Deletes a key in the registry
	// Remarks: EXAMPLE:<br>
	//   Call DeleteKey(HKEY_CURRENT_USER, "Software\VBW")
	//===============================================================================
	public static object DeleteKey(int hKey, string strKey)
	{
		object functionReturnValue = null;
		//EXAMPLE:
		//
		//Call DeleteKey(HKEY_CURRENT_USER, "Soft
		//     ware\VBW")
		//
		int r = 0;
		r = RegDeleteKey(hKey, strKey);
		if (r == 0) {
			functionReturnValue = "Success";
		}
		else {
			functionReturnValue = "No! Key to Delete Or Key Not Found";
		}
		return functionReturnValue;

	}

	//===============================================================================
	// Name: Function CheckRegistryKey
	// Input:
	//   ByVal hKey As Long - HKEY
	//   ByVal KeyName As String - Key Name
	// Output:
	//   Boolean - True if the key exists
	// Purpose: Checks a given key in the registry
	// Remarks: None
	//===============================================================================
	public static bool CheckRegistryKey(int hKey, string KeyName)
	{
		bool functionReturnValue = false;
		int handle = 0;
		// Try to open the key
		if (RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, ref handle) == 0) {
			// The key exists
			functionReturnValue = true;
			// Close it before exiting
			RegCloseKey(handle);
		}
		return functionReturnValue;
	}

	//===============================================================================
	// Name: Function CreateRegistryKey
	// Input:
	//   ByVal hKey As Long - HKEY
	//   ByVal KeyName As String - Key Name
	// Output:
	//   Boolean - True if the key already exists, error if unable to create the key
	// Purpose: Creates a key in the registry
	// Remarks: None
	//===============================================================================
	public static bool CreateRegistryKey(int hKey, string KeyName)
	{
		bool functionReturnValue = false;
		const short REG_OPENED_EXISTING_KEY = 0x2;
		int handle = 0;
		int disposition = 0;
		SECURITY_ATTRIBUTES SEC = default(SECURITY_ATTRIBUTES);

		if (RegCreateKeyEx(hKey, KeyName, 0, (string)0, 0, 0, ref SEC, ref handle, ref disposition)) {
			modActiveLock.Set_Locale(modActiveLock.regionalSymbol);
			Err().Raise(1001, modTrial.ACTIVELOCKSTRING, "Unable to create the registry key");
		}
		else {
			// Return True if the key already existed.
			functionReturnValue = (disposition == REG_OPENED_EXISTING_KEY);
			// Close the key.
			RegCloseKey(handle);
		}
		return functionReturnValue;
	}

	//===============================================================================
	// Name: Function DeleteValue
	// Input:
	//   ByVal hKey As Long - HKEY
	//   ByVal strPath As String - Key Name
	//   ByVal strValue As String - Value Name
	// Output:
	//   Variant - Returns "Success" if successful
	// Purpose: Deletes a key value in the registry
	// Remarks: EXAMPLE:<br>
	//   Call DeleteValue(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
	//===============================================================================
	public static object DeleteValue(int hKey, string strPath, string strValue)
	{
		object functionReturnValue = null;
		//EXAMPLE:
		//
		//Call DeleteValue(HKEY_CURRENT_USER, "So
		//     ftware\VBW\Registry", "Dword")
		//
		int keyhand = 0;
		int r = 0;
		r = RegOpenKey(hKey, strPath, ref keyhand);
		r = RegDeleteValue(keyhand, strValue);
		r = RegCloseKey(keyhand);
		if (r == 0) {
			functionReturnValue = "Success";
		}
		else {
			functionReturnValue = "No! Value to Delete";
		}
		return functionReturnValue;

	}
}
