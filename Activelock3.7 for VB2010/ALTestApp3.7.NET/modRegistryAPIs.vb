Option Strict Off
Option Explicit On
Module modRegistryAPIs
	'  ///////////////////////////////////////////////////////////////////////
	'  / Filename:  modRegistryAPIs.bas                                      /
	'  / Version:   2.0.0.1                                                  /
	'  / Purpose:   Facilitates Windows registry access                      /
	'  / ActiveLock Software Group (ASG)                                     /
	'  /                                                                     /
	'  / Date Created:         ???? ??, ???? - UNKNOWN                       /
	'  / Date Last Modified:   July 07, 2003 - MEC                           /
	'  /                                                                     /
	'  / This software is released under the license detailed below and is   /
	'  / subject to said license. Neither this header nor the licese below   /
	'  / may be removed from this module.                                    /
	'  ///////////////////////////////////////////////////////////////////////
	'
	'
	'  ///////////////////////////////////////////////////////////////////////
	'  /                        SOFTWARE LICENSE                             /
	'  ///////////////////////////////////////////////////////////////////////
	'
	'   ActiveLock
	'   Copyright 1998-2002 Nelson Ferraz
	'   Copyright 2003 The ActiveLock Software Group (ASG)
	'   All material is the property of the contributing authors.
	'
	'   Redistribution and use in source and binary forms, with or without
	'   modification, are permitted provided that the following conditions are
	'   met:
	'
	'     [o] Redistributions of source code must retain the above copyright
	'         notice, this list of conditions and the following disclaimer.
	'
	'     [o] Redistributions in binary form must reproduce the above
	'         copyright notice, this list of conditions and the following
	'         disclaimer in the documentation and/or other materials provided
	'         with the distribution.
	'
	'   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
	'   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
	'   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
	'   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
	'   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
	'   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
	'   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
	'   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
	'   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
	'   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
	'   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
	'
	'
	'  ///////////////////////////////////////////////////////////////////////
	'  /                        MODULE CHANGE LOG                            /
	'  ///////////////////////////////////////////////////////////////////////
	'
	'   07.07.03 - MEC - Updated the header comments for this file.
	'
	'
	'  ///////////////////////////////////////////////////////////////////////
	'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
	'  ///////////////////////////////////////////////////////////////////////
	
	
	Public Const REG_SZ As Short = 1
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
	Public Const KEY_SET_VALUE As Short = &H2s
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
	
	Public Declare Function RegOpenKeyEx Lib "advapi32.dll"  Alias "RegOpenKeyExA"(ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
    Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As String, ByRef dwSize As Integer) As Integer
    Public Declare Function RegCreateKeyEx Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal Reserved As Integer, ByVal lpClass As String, ByVal dwOptions As Integer, ByVal samDesired As Integer, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByRef phkResult As Integer, ByRef lpdwDisposition As Integer) As Integer
    Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal dwReserved As Integer, ByVal dwType As Integer, ByRef lpValue As String, ByVal dwSize As Integer) As Integer
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
    Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias "RegQueryInfoKeyA" (ByVal hKey As Integer, ByVal lpClass As String, ByRef lpcbClass As Integer, ByVal lpReserved As Integer, ByRef lpcSubKeys As Integer, ByRef lpcbMaxSubKeyLen As Integer, ByRef lpcbMaxClassLen As Integer, ByRef lpcValues As Integer, ByRef lpcbMaxValueNameLen As Integer, ByRef lpcbMaxValueLen As Integer, ByRef lpcbSecurityDescriptor As Integer, ByRef lpftLastWriteTime As FILETIME) As Integer
	
	Public Function DeleteRegKey(ByRef lngKey As Integer, ByRef SubKey As String) As Integer
		Dim lngRet As Integer
		lngRet = RegDeleteKey(lngKey, SubKey)
		DeleteRegKey = lngRet
	End Function
	
	Public Function DeleteRegValue(ByRef lngKey As Integer, ByRef SubKey As String, ByRef ValueName As String) As Integer
		Dim lngRet As Integer
		Dim lngKeyRet As Integer
		
		lngRet = RegOpenKeyEx(lngKey, SubKey, 0, KEY_WRITE, lngKeyRet)
		If lngRet <> ERROR_SUCCESS Then Exit Function
		
		lngRet = RegDeleteValue(lngKeyRet, ValueName)
		DeleteRegValue = lngRet
		RegCloseKey(lngKeyRet)
	End Function
	
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
	
    Public Function ReadRegVal(ByRef lngKey As Integer, ByRef SubKey As String, ByRef DataName As String, ByRef DefaultData As Object) As Object
        'Dim lngKeyRet As Integer
        'Dim lngData As Integer
        'Dim strdata As String
        'Dim Datatype As Integer
        'Dim DataSize As Integer
        'Dim lngRet As Integer
        'ReadRegVal = DefaultData
        'lngRet = RegOpenKeyEx(lngKey, SubKey, 0, KEY_QUERY_VALUE, lngKeyRet)
        'If lngRet <> ERROR_SUCCESS Then Exit Function
        'lngRet = RegQueryValueEx(lngKeyRet, DataName, 0, Datatype, 0, DataSize)
        'If lngRet <> ERROR_SUCCESS Then
        '    RegCloseKey(lngKeyRet)
        '    Exit Function
        'End If
        'Select Case Datatype
        '    Case REG_SZ
        '        strdata = Space(DataSize + 1)
        '        lngRet = RegQueryValueEx(lngKeyRet, DataName, 0, Datatype, strdata, DataSize)
        '        If lngRet = ERROR_SUCCESS Then
        '            ReadRegVal = CObj(StripNulls(RTrim(strdata)))
        '        End If
        '    Case REG_DWORD
        '        lngRet = RegQueryValueEx(lngKeyRet, DataName, 0, Datatype, lngData, 4)
        '        If lngRet = ERROR_SUCCESS Then
        '            ReadRegVal = CObj(lngData)
        '        End If
        'End Select
        'RegCloseKey(lngKeyRet)

        Dim myKey As Microsoft.Win32.RegistryKey
        myKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(SubKey, False)
        ReadRegVal = myKey.GetValue("")
        myKey.Close()

    End Function

    Public Function GetSubKeys(ByRef strKey As String, ByRef SubKey As String, ByRef SubKeyCnt As Integer) As String
        Dim strValues() As String
        Dim strTemp As String = Nothing
        Dim lngSub As Integer
        Dim intCnt As Short
        Dim lngRet As Integer
        Dim intKeyCnt As Short
        Dim FT As FILETIME

        lngRet = RegOpenKeyEx(CInt(strKey), SubKey, 0, KEY_ENUMERATE_SUB_KEYS, lngSub)

        If lngRet <> ERROR_SUCCESS Then
            SubKeyCnt = 0
            Return Nothing
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

    Function StripNulls(ByVal s As String) As String
        Dim i As Short
        i = InStr(s, Chr(0))
        If i > 0 Then
            StripNulls = Left(s, i - 1)
        Else
            StripNulls = s
        End If
    End Function

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

    Public Sub alSaveSetting(ByRef strRegHive As String, ByRef strRegPath As String, ByRef strAppname As String, ByRef strSection As String, ByRef strKey As String, ByRef vData As Object)
        alSaveSettingReg(strRegHive, strRegPath, strAppname, strSection, strKey, vData)
    End Sub

    Public Function alGetSetting(ByRef strRegHive As String, ByRef strRegPath As String, ByRef strAppname As String, ByRef strSection As String, ByRef strKey As String, ByRef vDefault As Object) As Object
        alGetSetting = alGetSettingReg(strRegHive, strRegPath, strAppname, strSection, strKey, vDefault)
    End Function

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
End Module