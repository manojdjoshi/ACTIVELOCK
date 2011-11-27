Option Strict Off
Option Explicit On

#Region "Copyright"
' This project is available from SVN on SourceForge.net under the main project, Activelock !
'
' ProjectPage: http://sourceforge.net/projects/activelock
' WebSite: http://www.activeLockSoftware.com
' DeveloperForums: http://forums.activelocksoftware.com
' ProjectManager: Ismail Alkan - http://activelocksoftware.com/simplemachinesforum/index.php?action=profile;u=1
' ProjectLicense: BSD Open License - http://www.opensource.org/licenses/bsd-license.php
' ProjectPurpose: Copy Protection, Software Locking, Anti Piracy
'
' //////////////////////////////////////////////////////////////////////////////////////////
' *   ActiveLock
' *   Copyright 1998-2002 Nelson Ferraz
' *   Copyright 2003-2009 The Activelock - Ismail Alkan (ASG)
' *   All material is the property of the contributing authors.
' *
' *   Redistribution and use in source and binary forms, with or without
' *   modification, are permitted provided that the following conditions are
' *   met:
' *
' *     [o] Redistributions of source code must retain the above copyright
' *         notice, this list of conditions and the following disclaimer.
' *
' *     [o] Redistributions in binary form must reproduce the above
' *         copyright notice, this list of conditions and the following
' *         disclaimer in the documentation and/or other materials provided
' *         with the distribution.
' *
' *   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' *   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' *   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
' *   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
' *   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
' *   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
' *   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
' *   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' *   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' *   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' *   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' *
#End Region

''' <summary>
''' Facilitates Windows registry access
''' </summary>
''' <remarks></remarks>
Module modRegistry

    '===============================================================================
    ' Name: modRegistryAPIs
    ' Purpose: Facilitates Windows registry access
    ' Date Last Modified:   July 07, 2003 - MEC
    ' Functions:
    ' Properties:
    ' Methods:
    ' Started: 07.07.2003
    ' Modified: 03.25.2006
    '===============================================================================

	Public Const REG_SZ As Short = 1 ' Unicode nul terminated string
	Public Const REG_EXPAND_SZ As Short = 2
	Public Const REG_BINARY As Short = 3
	Public Const REG_DWORD As Short = 4
	
	Public Const HKEY_CLASSES_ROOT As Integer = &H80000000
	Public Const HKEY_CURRENT_USER As Integer = &H80000001
	Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
	Public Const HKEY_USERS As Integer = &H80000003
	Public Const HKEY_PERFORMANCE_DATA As Integer = &H80000004
	
	Public Const HKEY_CURRENT_CONFIG As Integer = &H80000005
	Public Const HKEY_DYN_DATA As Integer = &H80000006
	
	Public Const REG_OPTION_NON_VOLATILE As Short = 0
	Public Const REG_CREATED_NEW_KEY As Short = &H1s
	Public Const REG_OPENED_EXISTING_KEY As Short = &H2s
	
	Public Const KEY_QUERY_VALUE As Short = &H1s
	Public Const KEY_ENUMERATE_SUB_KEYS As Short = &H8s
	Public Const KEY_NOTIFY As Short = &H10s
	Public Const READ_CONTROL As Integer = &H20000
	Public Const STANDARD_RIGHTS_ALL As Integer = &H1F0000
	Public Const STANDARD_RIGHTS_EXECUTE As Integer = (READ_CONTROL)
	Public Const STANDARD_RIGHTS_READ As Integer = (READ_CONTROL)
	Public Const STANDARD_RIGHTS_REQUIRED As Integer = &HF0000
	Public Const SYNCHRONIZE As Integer = &H100000
	Public Const KEY_READ As Boolean = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    Public Const KEY_SET_VALUE As Short = &H2S
	Public Const KEY_CREATE_SUB_KEY As Short = &H4s
	Public Const KEY_CREATE_LINK As Short = &H20s
	Public Const STANDARD_RIGHTS_WRITE As Integer = (READ_CONTROL)
	Public Const KEY_WRITE As Boolean = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
	Public Const KEY_ALL_ACCESS As Boolean = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

	Public Const ERROR_SUCCESS As Short = 0
	Public Const ERROR_ACCESS_DENIED As Short = 5
	Public Const ERROR_NO_MORE_ITEMS As Short = 259
	Public Const ERROR_BADKEY As Short = 1010
	Public Const ERROR_CANTOPEN As Short = 1011
	Public Const ERROR_CANTREAD As Short = 1012
	Public Const ERROR_REGISTRY_CORRUPT As Short = 1015
	
	Structure SECURITY_ATTRIBUTES
		Dim nLength As Integer
		Dim lpSecurityDescriptor As Integer
		Dim bInheritHandle As Boolean
	End Structure
	
	Public Structure FILETIME
		Dim dwLowDateTime As Integer
		Dim dwHighDateTime As Integer
	End Structure
	
    Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
    Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Integer, ByRef dwSize As Integer) As Integer
    Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal Reserved As Integer, ByVal lpClass As String, ByVal dwOptions As Integer, ByVal samDesired As Integer, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Integer, ByRef lpdwDisposition As Integer) As Integer
    Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal dwReserved As Integer, ByVal dwType As Integer, ByRef lpValue As Integer, ByVal dwSize As Integer) As Integer
    Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Integer, ByVal lpSubKey As String) As Integer
    Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Integer, ByVal lpValueName As String) As Integer
    Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
    Public Declare Function RegConnectRegistry Lib "advapi32.dll" Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal hKey As Integer, ByRef phkResult As Integer) As Integer
    Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
    Public Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpName As String, ByVal cbName As Integer) As Integer
    Public Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpValueName As String, ByRef lpcbValueName As Integer, ByRef lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Byte, ByRef lpcbData As Integer) As Integer
    Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Integer, ByVal dwIndex As Integer, ByVal lpName As String, ByRef lpcbName As Integer, ByVal lpReserved As Integer, ByVal lpClass As String, ByRef lpcbClass As Integer, ByRef lpftLastWriteTime As FILETIME) As Integer
    Public Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal lpFile As String) As Integer
    Public Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" (ByVal hKey As Integer, ByVal bWatchSubtree As Integer, ByVal dwNotifyFilter As Integer, ByVal hEvent As Integer, ByVal fAsynchronus As Integer) As Integer
    Public Declare Function RegOpenKey Lib "advapi32.dll" (ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
    Public Declare Function OSRegQueryValue Lib "advapi32.dll" Alias "RegQueryValueA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal lpValue As String, ByRef lpcbValue As Integer) As Integer
    Public Declare Function RegReplaceKey Lib "advapi32.dll" Alias "RegReplaceKeyA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal lpNewFile As String, ByVal lpOldFile As String) As Integer
    Public Declare Function RegRestoreKey Lib "advapi32.dll" Alias "RegRestoreKeyA" (ByVal hKey As Integer, ByVal lpFile As String, ByVal dwFlags As Integer) As Integer
    'Structure FILETIME may require marshalling attributes to be passed as an argument in this Declare statement.
    Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Integer, ByVal lpClass As String, ByRef lpcbClass As Integer, ByVal lpReserved As Integer, ByRef lpcSubKeys As Integer, ByRef lpcbMaxSubKeyLen As Integer, ByRef lpcbMaxClassLen As Integer, ByRef lpcValues As Integer, ByRef lpcbMaxValueNameLen As Integer, ByRef lpcbMaxValueLen As Integer, ByRef lpcbSecurityDescriptor As Integer, ByRef lpftLastWriteTime As FILETIME) As Integer
    '===============================================================================
	' Name: Function DeleteRegKey
	' Input:
	'   ByRef lngKey As Long - HKEY
	'   ByRef SubKey As String - Sub key name
	' Output:
	'   Long - Return value from the RegDeleteKey function
	' Purpose: Deletes a registry key
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Deletes a registry key
    ''' </summary>
    ''' <param name="lngKey">HKEY</param>
    ''' <param name="SubKey">Sub key name</param>
    ''' <returns>Return value from the RegDeleteKey function</returns>
    ''' <remarks></remarks>
	Public Function DeleteRegKey(ByRef lngKey As Integer, ByRef SubKey As String) As Integer
		Dim lngRet As Integer
		lngRet = RegDeleteKey(lngKey, SubKey)
		DeleteRegKey = lngRet
	End Function
	'===============================================================================
	' Name: Function DeleteRegValue
	' Input:
	'   ByRef lngKey As Long - HKEY
	'   ByRef SubKey As String - Sub key name
	'   ByRef ValueName As String - Value name
	' Output:
	'   Long - Return value from the RegDeleteValue function
	' Purpose: Deletes a registry value
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Deletes a registry value
    ''' </summary>
    ''' <param name="lngKey">HKEY</param>
    ''' <param name="SubKey">Sub key name</param>
    ''' <param name="ValueName">Value name</param>
    ''' <returns>Return value from the RegDeleteValue function</returns>
    ''' <remarks></remarks>
	Public Function DeleteRegValue(ByRef lngKey As Integer, ByRef SubKey As String, ByRef ValueName As String) As Integer
		Dim lngRet As Integer
		Dim lngKeyRet As Integer
		
		lngRet = RegOpenKeyEx(lngKey, SubKey, 0, KEY_WRITE, lngKeyRet)
		If lngRet <> ERROR_SUCCESS Then Exit Function
		
		lngRet = RegDeleteValue(lngKeyRet, ValueName)
		DeleteRegValue = lngRet
		RegCloseKey(lngKeyRet)
	End Function
    '===============================================================================
	' Name: Function WriteRegLong
	' Input:
	'   ByRef lngKey As Long - HKEY
	'   ByRef SubKey As String - Sub key name
	'   ByRef DataName As String - Value name
	'   ByRef DataValue As Long - Key value
	' Output: Long
	' Purpose: Writes a long key value in the registry
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Writes a long key value in the registry
    ''' </summary>
    ''' <param name="lngKey">HKEY</param>
    ''' <param name="SubKey">Sub key name</param>
    ''' <param name="DataName">Value name</param>
    ''' <param name="DataValue">Key value</param>
    ''' <returns>Long</returns>
    ''' <remarks></remarks>
	Public Function WriteRegLong(ByRef lngKey As Integer, ByRef SubKey As String, ByRef DataName As String, ByRef DataValue As Integer) As Integer
		
		Dim SEC As SECURITY_ATTRIBUTES
		Dim lngKeyRet As Integer
		Dim lngDis As Integer
		Dim lngRet As Integer
		
		lngRet = RegCreateKeyEx(lngKey, SubKey, 0, "", REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SEC, lngKeyRet, lngDis)
		
		If (lngRet = ERROR_SUCCESS) Or (lngRet = REG_CREATED_NEW_KEY) Or (lngRet = REG_OPENED_EXISTING_KEY) Then
			lngRet = RegSetValueEx(lngKeyRet, DataName, 0, REG_DWORD, DataValue, 4)
			RegCloseKey(lngKeyRet)
		End If
		WriteRegLong = lngRet
	End Function
    '===============================================================================
	' Name: Function WriteStringValue
	' Input:
	'   ByRef lngKey As Long - HKEY
	'   ByRef SubKey As String - Sub key name
	'   ByRef DataName As String - Value name
	'   ByRef DataValue As String - Key value
	' Output: Long
	' Purpose: Writes a string in the registry
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Writes a string in the registry
    ''' </summary>
    ''' <param name="lngKey">HKEY</param>
    ''' <param name="SubKey">Sub key name</param>
    ''' <param name="DataName">Value name</param>
    ''' <param name="DataValue">Key value</param>
    ''' <returns>Long</returns>
    ''' <remarks></remarks>
	Public Function WriteStringValue(ByRef lngKey As Integer, ByRef SubKey As String, ByRef DataName As String, ByRef DataValue As String) As Integer
		
		Dim SEC As SECURITY_ATTRIBUTES
		Dim lngKeyRet As Integer
		Dim lngDis As Integer
		Dim lngRet As Integer
		
		lngRet = RegCreateKeyEx(lngKey, SubKey, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, SEC, lngKeyRet, lngDis)
		If DataValue <= "" Then DataValue = ""
		If (lngRet = ERROR_SUCCESS) Or (lngRet = REG_CREATED_NEW_KEY) Or (lngRet = REG_OPENED_EXISTING_KEY) Then
			lngRet = RegSetValueEx(lngKeyRet, DataName, 0, REG_SZ, DataValue, Len(DataValue))
			RegCloseKey(lngKeyRet)
		End If
		WriteStringValue = lngRet
	End Function
    '===============================================================================
	' Name: Function ReadRegVal
	' Input:
	'   ByRef lngKey As Long - HKEY
	'   ByRef SubKey As String - Sub key name
	'   ByRef DataName As String - Value name
	'   ByRef DefaultData As Variant - Default value to be returned
	' Output: None
	' Purpose: Reads a key value from the registry
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Reads a key value from the registry
    ''' </summary>
    ''' <param name="lngKey">HKEY</param>
    ''' <param name="SubKey">Sub key name</param>
    ''' <param name="DataName">Value name</param>
    ''' <param name="DefaultData">Default value to be returned</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
	Public Function ReadRegVal(ByRef lngKey As Integer, ByRef SubKey As String, ByRef DataName As String, ByRef DefaultData As Object) As Object
		Dim lngKeyRet As Integer
		Dim lngData As Integer
		Dim strdata As String
		Dim datatype As Integer
		Dim DataSize As Integer
		Dim lngRet As Integer
        ReadRegVal = DefaultData
		lngRet = RegOpenKeyEx(lngKey, SubKey, 0, KEY_QUERY_VALUE, lngKeyRet)
		If lngRet <> ERROR_SUCCESS Then Exit Function
		lngRet = RegQueryValueEx(lngKeyRet, DataName, 0, datatype, 0, DataSize)
		If lngRet <> ERROR_SUCCESS Then
			RegCloseKey(lngKeyRet)
			Exit Function
		End If
		Select Case datatype
			Case REG_SZ
				strdata = Space(DataSize + 1)
				lngRet = RegQueryValueEx(lngKeyRet, DataName, 0, datatype, strdata, DataSize)
				If lngRet = ERROR_SUCCESS Then
                    ReadRegVal = CObj(StripNulls(RTrim(strdata)))
				End If
			Case REG_DWORD
				lngRet = RegQueryValueEx(lngKeyRet, DataName, 0, datatype, lngData, 4)
				If lngRet = ERROR_SUCCESS Then
                    ReadRegVal = CObj(lngData)
				End If
		End Select
		RegCloseKey(lngKeyRet)
	End Function
    '===============================================================================
	' Name: Function GetSubKeys
	' Input:
	'   ByRef strKey As String - Key name
	'   ByRef SubKey As String - Sub key name
	'   ByRef SubKeyCnt As Long - Number of keys
	' Output:
	'   String - Sub key list of a given key (separated by commas)
	' Purpose: Gets subkeys from a given key separated by commas
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Gets subkeys from a given key separated by commas
    ''' </summary>
    ''' <param name="strKey">Key name</param>
    ''' <param name="SubKey">Sub key name</param>
    ''' <param name="SubKeyCnt">Number of keys</param>
    ''' <returns>Sub key list of a given key (separated by commas)</returns>
    ''' <remarks></remarks>
	Public Function GetSubKeys(ByRef strKey As String, ByRef SubKey As String, ByRef SubKeyCnt As Integer) As String
		Dim strValues() As String
        Dim strTemp As String = String.Empty
		Dim lngSub As Integer
		Dim intCnt As Short
		Dim lngRet As Integer
		Dim intKeyCnt As Short
		Dim FT As FILETIME
		
        GetSubKeys = String.Empty
        lngRet = RegOpenKeyEx(CInt(strKey), SubKey, 0, KEY_ENUMERATE_SUB_KEYS, lngSub)
		
		If lngRet <> ERROR_SUCCESS Then
			SubKeyCnt = 0
			Exit Function
		End If
		lngRet = RegQueryInfoKey(lngSub, vbNullString, 0, 0, SubKeyCnt, 65, 0, 0, 0, 0, 0, FT)
		If (lngRet <> ERROR_SUCCESS) Or (SubKeyCnt <= 0) Then
			SubKeyCnt = 0
		End If
		ReDim strValues(SubKeyCnt - 1)
		For intCnt = 0 To SubKeyCnt - 1
			strValues(intCnt) = New String(Chr(0), 65)
			RegEnumKeyEx(lngSub, intCnt, strValues(intCnt), 65, 0, vbNullString, 0, FT)
			strValues(intCnt) = StripNulls(strValues(intCnt))
		Next intCnt
		RegCloseKey(lngSub)
		For intKeyCnt = LBound(strValues) To UBound(strValues)
			strTemp = strTemp & strValues(intKeyCnt) & ","
		Next intKeyCnt
		GetSubKeys = strTemp
	End Function
    '===============================================================================
	' Name: Function StripNulls
	' Input:
	'   ByVal s As String - Input string
	' Output:
	'   String - Returned string free of nulls
	' Purpose: Strips nulls in a given string
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Strips nulls in a given string
    ''' </summary>
    ''' <param name="s">Input string</param>
    ''' <returns>Returned string free of nulls</returns>
    ''' <remarks></remarks>
	Function StripNulls(ByVal s As String) As String
		Dim i As Short
		i = InStr(s, Chr(0))
		If i > 0 Then
			StripNulls = Left(s, i - 1)
		Else
			StripNulls = s
		End If
	End Function
    '===============================================================================
	' Name: Function ParseString
	' Input:
	'   ByRef strIn As String - Input string
	'   ByRef intLoc As Integer - Character location
	'   ByRef strDelimiter As String - String delimiter
	' Output:
	'   String - Parsed string
	' Purpose: String parser
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' String parser
    ''' </summary>
    ''' <param name="strIn">Input string</param>
    ''' <param name="intLoc">Character location</param>
    ''' <param name="strDelimiter">String delimiter</param>
    ''' <returns>Parsed string</returns>
    ''' <remarks></remarks>
	Public Function ParseString(ByRef strIn As String, ByRef intLoc As Short, ByRef strDelimiter As String) As String
		Dim intPos As Short
		Dim intStrt As Short
		Dim intStop As Short
		Dim intCnt As Short
		intCnt = intLoc
		Do While intCnt > 0
			intStop = intPos
			intStrt = InStr(intPos + 1, strIn, Left(strDelimiter, 1))
			If intStrt > 0 Then
				intPos = intStrt
				intCnt = intCnt - 1
			Else
				intPos = Len(strIn) + 1
				Exit Do
			End If
		Loop 
		ParseString = Mid(strIn, intStop + 1, intPos - intStop - 1)
	End Function
    '===============================================================================
	' Name: Sub alSaveSetting
	' Input:
	'   ByRef strRegHive As String, ByRef strRegPath As String, ByRef strAppname As String, ByRef strSection As String, ByRef strKey As String, ByRef vData As Variant
	'   ByRef strRegHive As String - Base registry class
	'   ByRef strRegPath As String - Registry key path under "Software"
	'   ByRef strAppname As String - Application name
	'   ByRef strSection As String - Section name
	'   ByRef strKey As String - Key name
	'   ByRef vDefault As Variant - Key value
	' Output: None
	' Purpose: Saves a key in the registry. Calls the alSaveSettingReg sub.
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Saves a key in the registry. Calls the alSaveSettingReg sub.
    ''' </summary>
    ''' <param name="strRegHive"></param>
    ''' <param name="strRegPath"></param>
    ''' <param name="strAppname"></param>
    ''' <param name="strSection"></param>
    ''' <param name="strKey"></param>
    ''' <param name="vData"></param>
    ''' <remarks></remarks>
	Public Sub alSaveSetting(ByRef strRegHive As String, ByRef strRegPath As String, ByRef strAppname As String, ByRef strSection As String, ByRef strKey As String, ByRef vData As Object)
		alSaveSettingReg(strRegHive, strRegPath, strAppname, strSection, strKey, vData)
	End Sub
	
	'===============================================================================
	' Name: Function alGetSetting
	' Input:
	'   ByRef strRegHive As String - Base registry class
	'   ByRef strRegPath As String - Registry key path under "Software"
	'   ByRef strAppname As String - Application name
	'   ByRef strSection As String - Section name
	'   ByRef strKey As String - Key name
	'   ByRef vDefault As Variant - Key value
	' Output: Variant
	' Purpose: Reads a key value from the registry. Calls the alGetSettingReg function to get a registry value
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Reads a key value from the registry. Calls the alGetSettingReg function to get a registry value
    ''' </summary>
    ''' <param name="strRegHive">Base registry class</param>
    ''' <param name="strRegPath">Registry key path under "Software"</param>
    ''' <param name="strAppname">Application name</param>
    ''' <param name="strSection">Section name</param>
    ''' <param name="strKey">Key name</param>
    ''' <param name="vDefault">Key value</param>
    ''' <returns>Variant</returns>
    ''' <remarks></remarks>
	Public Function alGetSetting(ByRef strRegHive As String, ByRef strRegPath As String, ByRef strAppname As String, ByRef strSection As String, ByRef strKey As String, ByRef vDefault As Object) As Object
        alGetSetting = alGetSettingReg(strRegHive, strRegPath, strAppname, strSection, strKey, vDefault)
	End Function
    '===============================================================================
	' Name: Sub alSaveSettingReg
	' Input:
	'   ByRef strRegHive As String - Base registry class
	'   ByRef strRegPath As String - Registry key path under "Software"
	'   ByRef strAppname As String - Application name
	'   ByRef strSection As String - Section name
	'   ByRef strKey As String - Key name
	'   ByRef vData As Variant - Key value
	' Output: None
	' Purpose: Saves a key in the registry.
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Saves a key in the registry.
    ''' </summary>
    ''' <param name="strRegHive">Base registry class</param>
    ''' <param name="strRegPath">Registry key path under "Software"</param>
    ''' <param name="strAppname">Application name</param>
    ''' <param name="strSection">Section name</param>
    ''' <param name="strKey">Key name</param>
    ''' <param name="vData">Key value</param>
    ''' <remarks></remarks>
	Public Sub alSaveSettingReg(ByRef strRegHive As String, ByRef strRegPath As String, ByRef strAppname As String, ByRef strSection As String, ByRef strKey As String, ByRef vData As Object)
		Dim lRegistryBase As Integer
		Select Case Left(UCase(strRegHive), 4)
			Case "HKLM"
				lRegistryBase = HKEY_LOCAL_MACHINE
			Case "HKCR"
				lRegistryBase = HKEY_CLASSES_ROOT
			Case Else
				lRegistryBase = HKEY_CURRENT_USER
		End Select
        WriteStringValue(lRegistryBase, "Software\" & strRegPath & "\" & strAppname & "\" & strSection, strKey, CStr(vData))
	End Sub
    '===============================================================================
	' Name: Function alGetSettingReg
	' Input:
	'   ByRef strRegHive As String - Base registry class
	'   ByRef strRegPath As String - Registry key path under "Software"
	'   ByRef strAppname As String - Application name
	'   ByRef strSection As String - Section name
	'   ByRef strKey As String - Key name
	'   ByRef vDefault As Variant - Key value
	' Output:
	'   Variant - Return value from the ReadRegVal function
	' Purpose: Reads a key value from the registry
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Reads a key value from the registry
    ''' </summary>
    ''' <param name="strRegHive">Base registry class</param>
    ''' <param name="strRegPath">Registry key path under "Software"</param>
    ''' <param name="strAppname">Application name</param>
    ''' <param name="strSection">Section name</param>
    ''' <param name="strKey">Key name</param>
    ''' <param name="vDefault">Key value</param>
    ''' <returns>Variant - Return value from the ReadRegVal function</returns>
    ''' <remarks></remarks>
	Public Function alGetSettingReg(ByRef strRegHive As String, ByRef strRegPath As String, ByRef strAppname As String, ByRef strSection As String, ByRef strKey As String, ByRef vDefault As Object) As Object
		Dim lRegistryBase As Integer
		Select Case Left(UCase(strRegHive), 4)
			Case "HKLM"
				lRegistryBase = HKEY_LOCAL_MACHINE
			Case "HKCR"
				lRegistryBase = HKEY_CLASSES_ROOT
			Case Else
				lRegistryBase = HKEY_CURRENT_USER
		End Select
        alGetSettingReg = ReadRegVal(lRegistryBase, "Software\" & strRegPath & "\" & strAppname & "\" & strSection, strKey, vDefault)
	End Function
    '===============================================================================
	' Name: Sub Savekey
	' Input:
	'   ByVal hKey As Long - HKEY
	'   ByVal strPath As String - Key Name
	' Output: None
	' Purpose: Saves a key in the registry
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Saves a key in the registry
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="strPath">Key Name</param>
    ''' <remarks></remarks>
	Public Sub Savekey(ByRef hKey As Integer, ByRef strPath As String)
		Dim keyhand, r As Integer
		r = RegCreateKey(hKey, strPath, keyhand)
		r = RegCloseKey(keyhand)
	End Sub
	'===============================================================================
	' Name: Function GetString
	' Input:
	'   ByRef hKey As Long - HKEY
	'   ByRef strPath As String - Key Name
    '   ByRef strValue As String - Value Name
	' Output:
	'   String
	' Purpose: Gets a string from the registry
	' Remarks:  EXAMPLE:<br>
	'   text1.text = getstring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String")
    '===============================================================================
    ''' <summary>
    ''' Gets a string from the registry
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="strPath">Key Name</param>
    ''' <param name="strValue">Value Name</param>
    ''' <returns>The key as a string.</returns>
    ''' <remarks>see code for an example.</remarks>
	Public Function GetString(ByRef hKey As Integer, ByRef strPath As String, ByRef strValue As String) As String
		'EXAMPLE:
		'
		'text1.text = getstring(HKEY_CURRENT_USE
		'     R, "Software\VBW\Registry", "String")
		'
		Dim keyhand, r As Integer
        Dim lValueType As Integer
		Dim lResult As Integer
		Dim strBuf As String
		Dim lDataBufSize As Integer
		Dim intZeroPos As Short

        GetString = String.Empty

        r = RegOpenKey(hKey, strPath, keyhand)
		lResult = RegQueryValueEx(keyhand, strValue, 0, lValueType, 0, lDataBufSize)
		
		If lValueType = REG_SZ Then
			strBuf = New String(" ", lDataBufSize)
			lResult = RegQueryValueEx(keyhand, strValue, 0, 0, strBuf, lDataBufSize)
			If lResult = ERROR_SUCCESS Then
				intZeroPos = InStr(strBuf, Chr(0))
				If intZeroPos > 0 Then
					GetString = Left(strBuf, intZeroPos - 1)
				Else
					GetString = strBuf
				End If
			End If
		End If
		If GetString = "" Then
			GetString = "Value or key does not exist"
		End If
		
	End Function
    '===============================================================================
	' Name: Function GetRegistryValue
	' Input:
	'   ByVal hKey As Long - HKEY
	'   ByVal KeyName As String - Key Name
	'   ByVal ValueName As String - Value Name
	'   ByRef defaultValue as Variant - Default value to be returned if the value is missing
	' Output:
	'   Variant - Registry value
	' Purpose: Gets a key value from the registry
	' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Gets a key value from the registry
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="KeyName">Key Name</param>
    ''' <param name="ValueName">Value Name</param>
    ''' <param name="defaultValue">Variant - Default value to be returned if the value is missing</param>
    ''' <returns>Variant - Registry value</returns>
    ''' <remarks></remarks>
	Public Function GetRegistryValue(ByVal hKey As Integer, ByVal KeyName As String, ByVal ValueName As String, Optional ByRef defaultValue As Object = Nothing) As Object
		Dim handle As Integer
		Dim resLong As Integer
		Dim resString As String
		Dim resBinary() As Byte
		Dim Length As Integer
		Dim retVal As Integer
		Dim valueType As Integer
		
		' Prepare the default result
        GetRegistryValue = IIf(IsNothing(defaultValue), Nothing, defaultValue)
		
		' Open the key, exit if not found.
		If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
			Exit Function
		End If
		
		' prepare a 1K receiving resBinary
		Length = 1024
		ReDim resBinary(Length - 1)
		
		' read the registry key
		retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
		' if resBinary was too small, try again
		If retVal = ERROR_MORE_DATA Then
			' enlarge the resBinary, and read the value again
			ReDim resBinary(Length - 1)
			retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), Length)
		End If
		
		' return a value corresponding to the value type
		Select Case valueType
			Case REG_DWORD
				CopyMemory(resLong, resBinary(0), 4)
                GetRegistryValue = resLong
			Case REG_SZ, REG_EXPAND_SZ
				' copy everything but the trailing null char
				resString = Space(Length - 1)
				CopyMemory(resString, resBinary(0), Length - 1)
                GetRegistryValue = resString
			Case REG_BINARY
				' resize the result resBinary
				If Length <> UBound(resBinary) + 1 Then
					ReDim Preserve resBinary(Length - 1)
				End If
                GetRegistryValue = resBinary    ' VB6.CopyArray(resBinary)
			Case REG_MULTI_SZ
				' copy everything but the 2 trailing null chars
				resString = Space(Length - 2)
				CopyMemory(resString, resBinary(0), Length - 2)
                GetRegistryValue = resString
			Case Else
				RegCloseKey(handle)
                Change_Culture("")
                Err.Raise(1001, ACTIVELOCKSTRING, "Unsupported value type")
        End Select

        ' close the registry key
        RegCloseKey(handle)
    End Function

    '===============================================================================
    ' Name: Function SetRegistryValue
    ' Input:
    '   ByVal hKey As Long - HKEY
    '   ByVal KeyName As String - Key Name
    '   ByVal ValueName As String - Value Name
    '   ByRef Value As Variant - Key Value.
    '   Value can be an integer value (REG_DWORD), a string (REG_SZ) or an array of binary (REG_BINARY). Raises an error otherwise.
    ' Output:
    '   Boolean - True if successful
    ' Purpose: Writes or Creates a Registry value
    ' Remarks: Use KeyName = "" for the default value
    '===============================================================================
    ''' <summary>
    ''' Writes or Creates a Registry value
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="KeyName">Key Name</param>
    ''' <param name="ValueName">Value Name</param>
    ''' <param name="Value">Key Value.
    ''' <para>Value can be an integer value (REG_DWORD), a string (REG_SZ) or an array of binary (REG_BINARY). Raises an error otherwise.</para></param>
    ''' <returns>True if successful</returns>
    ''' <remarks>Use KeyName = "" for the default value</remarks>
    Public Function SetRegistryValue(ByVal hKey As Integer, ByVal KeyName As String, ByVal ValueName As String, ByRef Value As Object) As Boolean
        Const KEY_WRITE As Integer = &H20006 '((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or
        ' KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

        ' Write or Create a Registry value
        ' returns True if successful
        '
        ' Use KeyName = "" for the default value
        '
        ' Value can be an integer value (REG_DWORD), a string (REG_SZ)
        ' or an array of binary (REG_BINARY). Raises an error otherwise.
        Dim handle As Integer
        Dim lngValue As Integer
        Dim strValue As String
        Dim binValue() As Byte
        Dim Length As Integer
        Dim retVal As Integer

        ' Open the key, exit if not found
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then
            Exit Function
        End If

        ' three cases, according to the data type in Value
        Select Case VarType(Value)
            Case VariantType.Short, VariantType.Integer
                lngValue = Value
                retVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
            Case VariantType.String
                strValue = Value
                retVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, strValue, Len(strValue))
            Case VariantType.Array + VariantType.Byte
                binValue = Value
                Length = UBound(binValue) - LBound(binValue) + 1
                retVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, binValue(LBound(binValue)), Length)
            Case Else
                RegCloseKey(handle)
                Change_Culture("")
                Err.Raise(1001, ACTIVELOCKSTRING, "Unsupported value type")
        End Select

        ' Close the key and signal success
        RegCloseKey(handle)
        ' signal success if the value was written correctly
        SetRegistryValue = (retVal = 0)
    End Function

    '===============================================================================
    ' Name: Function SaveString
    ' Input:
    '   ByRef hKey As Long - HKEY
    '   ByRef strPath As String - Key Name
    '   ByRef strValue As String - Value Name
    '   ByRef strdata As String - Key Value
    ' Output:
    '   Variant - Returns "Success" if successful
    ' Purpose: Saves a string in the registry
    ' Remarks:  EXAMPLE:<br>
    '   text1.text= savestring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String", text1.text)
    '===============================================================================
    ''' <summary>
    ''' Saves a string in the registry
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="strPath">Key Name</param>
    ''' <param name="strValue">Value Name</param>
    ''' <param name="strdata">Key Value</param>
    ''' <returns>Variant - Returns "Success" if successful</returns>
    ''' <remarks>see code for an example</remarks>
    Public Function SaveString(ByRef hKey As Integer, ByRef strPath As String, ByRef strValue As String, ByRef strdata As String) As Boolean
        'EXAMPLE:
        '
        'text1.text= savestring(HKEY_CURRENT_USER, "Sof
        '     tware\VBW\Registry", "String", text1.tex
        '     t)
        '
        Dim keyhand As Integer
        Dim r As Integer
        r = RegCreateKey(hKey, strPath, keyhand)
        r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, strdata, Len(strdata))
        r = RegCloseKey(keyhand)
        If r = 0 Then
            SaveString = True   '"Success"
        Else
            SaveString = False   '"Key to Delete Or Key Not Found"
        End If

    End Function

    '===============================================================================
    ' Name: Function Getdword
    ' Input:
    '   ByVal hKey As Long - HKEY
    '   ByVal strPath As String - Key Name
    '   ByVal strValueName As String - Value Name
    ' Output:
    '   Variant - Returns the DWORD if successful
    ' Purpose: Gets the DWORD of a key from the registry
    ' Remarks: EXAMPLE:<br>
    '   text1.text = getdword(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
    '===============================================================================
    ''' <summary>
    ''' Gets the DWORD of a key from the registry
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="strPath">Key Name</param>
    ''' <param name="strValueName">Value Name</param>
    ''' <returns>Variant - Returns the DWORD if successful</returns>
    ''' <remarks>see code for an example</remarks>
    Function Getdword(ByVal hKey As Integer, ByVal strPath As String, ByVal strValueName As String) As Object
        'EXAMPLE:
        '
        'text1.text = getdword(HKEY_CURRENT_USER
        '     , "Software\VBW\Registry", "Dword")
        '
        Dim lResult As Integer
        Dim lValueType As Integer
        Dim lBuf As Integer
        Dim lDataBufSize As Integer
        Dim r As Integer
        Dim keyhand As Integer

        Getdword = String.Empty

        r = RegOpenKey(hKey, strPath, keyhand)
        ' Get length/data type
        lDataBufSize = 4
        lResult = RegQueryValueEx(keyhand, strValueName, 0, lValueType, lBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            If lValueType = REG_DWORD Then
                Getdword = lBuf
            End If
        End If
        r = RegCloseKey(keyhand)
        If Getdword = "" Then
            Getdword = " Value or key does not exist"
        End If

    End Function

    '===============================================================================
    ' Name: Function SaveDword
    ' Input:
    '   ByVal hKey As Long - HKEY
    '   ByVal strPath As String - Key Name
    '   ByVal strValueName As String - Value Name
    '   ByVal lData As Long - Key Value
    ' Output:
    '   Variant - Returns "Success" if successful
    ' Purpose: None
    ' Remarks: None
    '===============================================================================
    ' Saves a DWORD in the registry
    ' @param hKey           HKEY
    ' @param strPath        Key Name
    ' @param strValueName   Value Name
    ' @param lData          Value
    ' @return               "Success" if success
    '
    ''' <summary>
    ''' Saves a DWORD in the registry
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="strPath">Key Name</param>
    ''' <param name="strValueName">Value Name</param>
    ''' <param name="lData">Key Value</param>
    ''' <returns>Variant - Returns "Success" if successful</returns>
    ''' <remarks>see code for an example</remarks>
    Function SaveDword(ByVal hKey As Integer, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Integer) As Object
        'EXAMPLE"
        '
        'Text1.text= SaveDword(HKEY_CURRENT_USER, "Soft
        '     ware\VBW\Registry", "Dword", text1.text)
        '
        '
        Dim lResult As Integer
        Dim keyhand As Integer
        Dim r As Integer
        r = RegCreateKey(hKey, strPath, keyhand)
        lResult = RegSetValueEx(keyhand, strValueName, 0, REG_DWORD, lData, 4)
        'If lResult <> error_success Then Call e
        '     rrlog("SetDWORD", False)
        r = RegCloseKey(keyhand)
        If r = 0 Then
            SaveDword = "Success"
        Else
            SaveDword = " Failed to save Value"
        End If

    End Function

    '===============================================================================
    ' Name: Function DeleteKey
    ' Input:
    '   ByVal hKey As Long - HKEY
    '   ByVal strKey As String - Key Name
    ' Output:
    '   Variant - Returns "Success" if successful
    ' Purpose: Deletes a key in the registry
    ' Remarks: EXAMPLE:<br>
    '   Call DeleteKey(HKEY_CURRENT_USER, "Software\VBW")
    '===============================================================================
    ''' <summary>
    ''' Deletes a key in the registry
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="strKey">Key Name</param>
    ''' <returns>Variant - Returns "Success" if successful</returns>
    ''' <remarks>see code for an example</remarks>
    Public Function DeleteKey(ByVal hKey As Integer, ByVal strKey As String) As Boolean
        'EXAMPLE:
        '
        'Call DeleteKey(HKEY_CURRENT_USER, "Soft
        '     ware\VBW")
        '
        Dim r As Integer
        r = RegDeleteKey(hKey, strKey)
        If r = 0 Then
            DeleteKey = True '"Success"
        Else
            DeleteKey = False '"No! Key to Delete Or Key Not Found"
        End If

    End Function

    '===============================================================================
    ' Name: Function CheckRegistryKey
    ' Input:
    '   ByVal hKey As Long - HKEY
    '   ByVal KeyName As String - Key Name
    ' Output:
    '   Boolean - True if the key exists
    ' Purpose: Checks a given key in the registry
    ' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Checks a given key in the registry
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="KeyName">Key Name</param>
    ''' <returns>True if the key exists</returns>
    ''' <remarks></remarks>
    Public Function CheckRegistryKey(ByVal hKey As Integer, ByVal KeyName As String) As Boolean
        Dim handle As Integer
        ' Try to open the key
        If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) = 0 Then
            ' The key exists
            CheckRegistryKey = True
            ' Close it before exiting
            RegCloseKey(handle)
        End If
    End Function

    '===============================================================================
    ' Name: Function CreateRegistryKey
    ' Input:
    '   ByVal hKey As Long - HKEY
    '   ByVal KeyName As String - Key Name
    ' Output:
    '   Boolean - True if the key already exists, error if unable to create the key
    ' Purpose: Creates a key in the registry
    ' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Creates a key in the registry
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="KeyName">Key Name</param>
    ''' <returns>True if the key already exists, error if unable to create the key</returns>
    ''' <remarks></remarks>
    Function CreateRegistryKey(ByVal hKey As Integer, ByVal KeyName As String) As Boolean
        Const REG_OPENED_EXISTING_KEY As Short = &H2S
        Dim handle, disposition As Integer
        Dim SEC As SECURITY_ATTRIBUTES

        If RegCreateKeyEx(hKey, KeyName, 0, CStr(0), 0, 0, SEC, handle, disposition) Then
            Change_Culture("")
            Err.Raise(1001, ACTIVELOCKSTRING, "Unable to create the registry key")
        Else
            ' Return True if the key already existed.
            CreateRegistryKey = (disposition = REG_OPENED_EXISTING_KEY)
            ' Close the key.
            RegCloseKey(handle)
        End If
    End Function
	
	'===============================================================================
	' Name: Function DeleteValue
	' Input:
	'   ByVal hKey As Long - HKEY
	'   ByVal strPath As String - Key Name
	'   ByVal strValue As String - Value Name
	' Output:
	'   Variant - Returns "Success" if successful
	' Purpose: Deletes a key value in the registry
	' Remarks: EXAMPLE:<br>
	'   Call DeleteValue(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
    '===============================================================================
    ''' <summary>
    ''' Deletes a key value in the registry
    ''' </summary>
    ''' <param name="hKey">HKEY</param>
    ''' <param name="strPath">Key Name</param>
    ''' <param name="strValue">Value Name</param>
    ''' <returns>Variant - Returns "Success" if successful</returns>
    ''' <remarks>see code for an example.</remarks>
	Public Function DeleteValue(ByVal hKey As Integer, ByVal strPath As String, ByVal strValue As String) As Object
		'EXAMPLE:
		'
		'Call DeleteValue(HKEY_CURRENT_USER, "So
		'     ftware\VBW\Registry", "Dword")
		'
		Dim keyhand, r As Integer
		r = RegOpenKey(hKey, strPath, keyhand)
		r = RegDeleteValue(keyhand, strValue)
		r = RegCloseKey(keyhand)
		If r = 0 Then
            DeleteValue = "Success"
		Else
            DeleteValue = "No! Value to Delete"
		End If
		
	End Function
End Module