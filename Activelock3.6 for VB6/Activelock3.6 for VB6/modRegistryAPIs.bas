Attribute VB_Name = "modRegistry"
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
' *   Copyright 2003-2009 The ActiveLock Software Group (ASG)
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
'===============================================================================
' Name: modRegistryAPIs
' Purpose: Facilitates Windows registry access
' Date Last Modified:   July 07, 2003 - MEC
' Functions:
' Properties:
' Methods:
' Started: 07.07.2003
' Modified: 08.15.2005
'===============================================================================
'
''  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
'
'   07.07.03 - MEC - Updated the header comments for this file.
'
'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////

Option Explicit

Public Const REG_SZ = 1 ' Unicode nul terminated string
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4

Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004

Public Const HKEY_CURRENT_CONFIG = &H80000005
Public Const HKEY_DYN_DATA = &H80000006

Public Const REG_OPTION_NON_VOLATILE = 0
Public Const REG_CREATED_NEW_KEY = &H1
Public Const REG_OPENED_EXISTING_KEY = &H2

Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
'Const KEY_READ = &H20019  ' ((READ_CONTROL Or KEY_QUERY_VALUE Or
                          ' KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not
                          ' SYNCHRONIZE))
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_CREATE_LINK = &H20
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))
'Const KEY_ALL_ACCESS = &H3F

Public Const ERROR_SUCCESS = 0&
Public Const ERROR_ACCESS_DENIED = 5&
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const ERROR_BADKEY = 1010&
Public Const ERROR_CANTOPEN = 1011&
Public Const ERROR_CANTREAD = 1012&
Public Const ERROR_REGISTRY_CORRUPT = 1015&

Type SECURITY_ATTRIBUTES
  nLength As Long
  lpSecurityDescriptor As Long
  bInheritHandle As Boolean
End Type

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
"RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long

Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias _
"RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As _
String, ByVal lpReserved As Long, lpType As Long, lpData As Any, _
dwSize As Long) As Long

Public Declare Function RegCreateKeyEx Lib "advapi32" _
Alias "RegCreateKeyExA" (ByVal hKey As Long, _
ByVal lpSubKey As String, ByVal Reserved As Long, _
ByVal lpClass As String, ByVal dwOptions As Long, _
ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, _
phkResult As Long, lpdwDisposition As Long) As Long

Public Declare Function RegSetValueEx Lib "advapi32.dll" _
Alias "RegSetValueExA" (ByVal hKey As Long, _
ByVal lpValueName As String, ByVal dwReserved As Long, _
ByVal dwType As Long, lpValue As Any, ByVal dwSize As Long) As Long

Public Declare Function RegDeleteKey Lib "advapi32.dll" _
Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long

Public Declare Function RegDeleteValue Lib "advapi32.dll" _
Alias "RegDeleteValueA" (ByVal hKey As Long, _
ByVal lpValueName As String) As Long

Public Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long

Public Declare Function RegConnectRegistry Lib "advapi32.dll" _
Alias "RegConnectRegistryA" (ByVal lpMachineName As String, ByVal _
hKey As Long, phkResult As Long) As Long

Public Declare Function RegCreateKey Lib "advapi32.dll" Alias _
"RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
phkResult As Long) As Long

Public Declare Function RegEnumKey Lib "advapi32.dll" Alias _
"RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal _
lpName As String, ByVal cbName As Long) As Long

Public Declare Function RegEnumValue Lib "advapi32.dll" Alias _
"RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal _
lpValueName As String, lpcbValueName As Long, lpReserved As Long, _
lpType As Long, lpData As Byte, lpcbData As Long) As Long

Public Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias _
"RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal _
lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal _
lpClass As String, lpcbClass As Long, lpftLastWriteTime As FILETIME) As Long

Public Declare Function RegLoadKey Lib "advapi32.dll" Alias "RegLoadKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, ByVal lpFile As String) As Long

Public Declare Function RegNotifyChangeKeyValue Lib "advapi32.dll" _
(ByVal hKey As Long, ByVal bWatchSubtree As Long, ByVal dwNotifyFilter _
As Long, ByVal hEvent As Long, ByVal fAsynchronus As Long) As Long

Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
(ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Public Declare Function OSRegQueryValue Lib "advapi32.dll" _
Alias "RegQueryValueA" (ByVal hKey As Long, ByVal lpSubKey As _
String, ByVal lpValue As String, lpcbValue As Long) As Long

Public Declare Function RegReplaceKey Lib "advapi32.dll" Alias _
"RegReplaceKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal lpNewFile As String, ByVal lpOldFile As String) As Long

Public Declare Function RegRestoreKey Lib "advapi32.dll" Alias _
"RegRestoreKeyA" (ByVal hKey As Long, ByVal lpFile As String, _
ByVal dwFlags As Long) As Long

Public Declare Function RegQueryInfoKey Lib "advapi32.dll" Alias _
"RegQueryInfoKeyA" (ByVal hKey As Long, ByVal lpClass As String, _
lpcbClass As Long, ByVal lpReserved As Long, lpcSubKeys As Long, _
lpcbMaxSubKeyLen As Long, lpcbMaxClassLen As Long, lpcValues As Long, _
lpcbMaxValueNameLen As Long, lpcbMaxValueLen As Long, lpcbSecurityDescriptor _
As Long, lpftLastWriteTime As FILETIME) As Long

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
Public Function DeleteRegKey(lngKey As Long, SubKey As String) As Long
    Dim lngRet As Long
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
Public Function DeleteRegValue(lngKey As Long, SubKey As String, ValueName As String) As Long
    Dim lngRet As Long
    Dim lngKeyRet As Long
    
    lngRet = RegOpenKeyEx(lngKey, SubKey, 0, KEY_WRITE, lngKeyRet)
    If lngRet <> ERROR_SUCCESS Then Exit Function
    
    lngRet = RegDeleteValue(lngKeyRet, ValueName)
    DeleteRegValue = lngRet
    RegCloseKey lngKeyRet
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
Public Function WriteRegLong(lngKey As Long, SubKey As String, _
      DataName As String, DataValue As Long) As Long
      
      Dim SEC As SECURITY_ATTRIBUTES
      Dim lngKeyRet As Long
      Dim lngDis As Long
      Dim lngRet As Long
      
      lngRet = RegCreateKeyEx(lngKey, SubKey, 0, "", REG_OPTION_NON_VOLATILE, _
      KEY_ALL_ACCESS, SEC, lngKeyRet, lngDis)
    
      If (lngRet = ERROR_SUCCESS) Or (lngRet = REG_CREATED_NEW_KEY) Or _
      (lngRet = REG_OPENED_EXISTING_KEY) Then
        lngRet = RegSetValueEx(lngKeyRet, DataName, 0&, REG_DWORD, DataValue, 4)
        RegCloseKey lngKeyRet
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
Public Function WriteStringValue(lngKey As Long, SubKey As String, _
    DataName As String, DataValue As String) As Long
    
    Dim SEC As SECURITY_ATTRIBUTES
    Dim lngKeyRet As Long
    Dim lngDis As Long
    Dim lngRet As Long
    
    lngRet = RegCreateKeyEx(lngKey, _
      SubKey, 0, vbNullString, _
      REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, _
      SEC, lngKeyRet, lngDis)
    If DataValue <= "" Then DataValue = ""
    If (lngRet = ERROR_SUCCESS) Or (lngRet = REG_CREATED_NEW_KEY) Or _
    (lngRet = REG_OPENED_EXISTING_KEY) Then
      lngRet = RegSetValueEx(lngKeyRet, DataName, 0&, _
          REG_SZ, ByVal DataValue, Len(DataValue))
      RegCloseKey lngKeyRet
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
Public Function ReadRegVal(lngKey As Long, SubKey As String, _
    DataName As String, DefaultData As Variant) As Variant
    Dim lngKeyRet As Long
    Dim lngData As Long
    Dim strdata As String
    Dim datatype As Long
    Dim DataSize As Long
    Dim lngRet As Long
    ReadRegVal = DefaultData
    lngRet = RegOpenKeyEx(lngKey, SubKey, 0, KEY_QUERY_VALUE, lngKeyRet)
    If lngRet <> ERROR_SUCCESS Then Exit Function
    lngRet = RegQueryValueEx(lngKeyRet, DataName, 0&, datatype, ByVal 0, DataSize)
    If lngRet <> ERROR_SUCCESS Then
      RegCloseKey lngKeyRet
      Exit Function
    End If
    Select Case datatype
      Case REG_SZ
        strdata = Space(DataSize + 1)
        lngRet = RegQueryValueEx(lngKeyRet, DataName, 0&, datatype, ByVal strdata, DataSize)
        If lngRet = ERROR_SUCCESS Then
          ReadRegVal = CVar(StripNulls(RTrim$(strdata)))
        End If
      Case REG_DWORD
        lngRet = RegQueryValueEx(lngKeyRet, DataName, 0&, datatype, lngData, 4)
        If lngRet = ERROR_SUCCESS Then
          ReadRegVal = CVar(lngData)
        End If
      End Select
    RegCloseKey lngKeyRet
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
Public Function GetSubKeys(strKey As String, SubKey As String, ByRef SubKeyCnt As Long) As String
    Dim strValues() As String
    Dim strTemp As String
    Dim lngSub As Long
    Dim intCnt As Integer
    Dim lngRet As Long
    Dim intKeyCnt As Integer
    Dim FT As FILETIME
    
    lngRet = RegOpenKeyEx(strKey, SubKey, 0, KEY_ENUMERATE_SUB_KEYS, lngSub)
    
    If lngRet <> ERROR_SUCCESS Then
      SubKeyCnt = 0
      Exit Function
    End If
    lngRet = RegQueryInfoKey(lngSub, vbNullString, 0, 0, SubKeyCnt, _
    65, 0, 0, 0, 0, 0, FT)
    If (lngRet <> ERROR_SUCCESS) Or (SubKeyCnt <= 0) Then
      SubKeyCnt = 0
    End If
    ReDim strValues(SubKeyCnt - 1)
    For intCnt = 0 To SubKeyCnt - 1
      strValues(intCnt) = String$(65, 0)
      RegEnumKeyEx lngSub, intCnt, strValues(intCnt), 65, 0, vbNullString, 0, FT
      strValues(intCnt) = StripNulls(strValues(intCnt))
    Next intCnt
    RegCloseKey lngSub
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
Function StripNulls(ByVal S As String) As String
    Dim i As Integer
    i = InStr(S, Chr$(0))
    If i > 0 Then
      StripNulls = Left$(S, i - 1)
    Else
      StripNulls = S
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
Public Function ParseString(strIn As String, intLoc As Integer, strDelimiter As String) As String
    Dim intPos As Integer
    Dim intStrt As Integer
    Dim intStop As Integer
    Dim intCnt As Integer
    intCnt = intLoc
    Do While intCnt > 0
      intStop = intPos
      intStrt = InStr(intPos + 1, strIn, Left$(strDelimiter, 1))
        If intStrt > 0 Then
          intPos = intStrt
          intCnt = intCnt - 1
        Else
          intPos = Len(strIn) + 1
          Exit Do
          End If
    Loop
    ParseString = Mid$(strIn, intStop + 1, intPos - intStop - 1)
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
Public Sub alSaveSetting(strRegHive As String, strRegPath As String, strAppname As String, strSection As String, strKey As String, vData As Variant)
    alSaveSettingReg strRegHive, strRegPath, strAppname, strSection, strKey, vData
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
Public Function alGetSetting(strRegHive As String, strRegPath As String, strAppname As String, strSection As String, strKey As String, vDefault As Variant) As Variant
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
Public Sub alSaveSettingReg(strRegHive As String, strRegPath As String, strAppname As String, strSection As String, strKey As String, vData As Variant)
    Dim lRegistryBase As Long
    Select Case Left(UCase(strRegHive), 4)
    Case "HKLM"
      lRegistryBase = HKEY_LOCAL_MACHINE
    Case "HKCR"
      lRegistryBase = HKEY_CLASSES_ROOT
    Case Else
      lRegistryBase = HKEY_CURRENT_USER
    End Select
    WriteStringValue lRegistryBase, "Software\" & strRegPath & "\" & strAppname & "\" & strSection, strKey, CStr(vData)
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
Public Function alGetSettingReg(strRegHive As String, strRegPath As String, strAppname As String, strSection As String, strKey As String, vDefault As Variant) As Variant
    Dim lRegistryBase As Long
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
Public Sub Savekey(hKey As Long, strPath As String)
    Dim keyhand&, r As Long
    r = RegCreateKey(hKey, strPath, keyhand&)
    r = RegCloseKey(keyhand&)
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
Public Function GetString(hKey As Long, strPath As String, strValue As String) As String
    'EXAMPLE:
    '
    'text1.text = getstring(HKEY_CURRENT_USE
    '     R, "Software\VBW\Registry", "String")
    '
    Dim keyhand As Long, r As Long
    Dim datatype As Long, lValueType As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)

    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
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
Public Function GetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, Optional defaultValue As Variant) As Variant
Dim handle As Long
Dim resLong As Long
Dim resString As String
Dim resBinary() As Byte
Dim Length As Long
Dim retVal As Long
Dim valueType As Long

' Prepare the default result
GetRegistryValue = IIf(IsMissing(defaultValue), Empty, defaultValue)

' Open the key, exit if not found.
If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) Then
    Exit Function
End If

' prepare a 1K receiving resBinary
Length = 1024
ReDim resBinary(0 To Length - 1) As Byte

' read the registry key
retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
    Length)
' if resBinary was too small, try again
If retVal = ERROR_MORE_DATA Then
    ' enlarge the resBinary, and read the value again
    ReDim resBinary(0 To Length - 1) As Byte
    retVal = RegQueryValueEx(handle, ValueName, 0, valueType, resBinary(0), _
        Length)
End If

' return a value corresponding to the value type
Select Case valueType
    Case REG_DWORD
        CopyMemory resLong, resBinary(0), 4
        GetRegistryValue = resLong
    Case REG_SZ, REG_EXPAND_SZ
        ' copy everything but the trailing null char
        resString = Space$(Length - 1)
        CopyMemory ByVal resString, resBinary(0), Length - 1
        GetRegistryValue = resString
    Case REG_BINARY
        ' resize the result resBinary
        If Length <> UBound(resBinary) + 1 Then
            ReDim Preserve resBinary(0 To Length - 1) As Byte
        End If
        GetRegistryValue = resBinary()
    Case REG_MULTI_SZ
        ' copy everything but the 2 trailing null chars
        resString = Space$(Length - 2)
        CopyMemory ByVal resString, resBinary(0), Length - 2
        GetRegistryValue = resString
    Case Else
        RegCloseKey handle
        Set_locale regionalSymbol
        Err.Raise 1001, ACTIVELOCKSTRING, "Unsupported value type"
End Select

' close the registry key
RegCloseKey handle
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
Public Function SetRegistryValue(ByVal hKey As Long, ByVal KeyName As String, _
    ByVal ValueName As String, Value As Variant) As Boolean
    Const KEY_WRITE = &H20006  '((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or
                               ' KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
    
    ' Write or Create a Registry value
    ' returns True if successful
    '
    ' Use KeyName = "" for the default value
    '
    ' Value can be an integer value (REG_DWORD), a string (REG_SZ)
    ' or an array of binary (REG_BINARY). Raises an error otherwise.
    Dim handle As Long
    Dim lngValue As Long
    Dim strValue As String
    Dim binValue() As Byte
    Dim Length As Long
    Dim retVal As Long
    
    ' Open the key, exit if not found
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_WRITE, handle) Then
        Exit Function
    End If

    ' three cases, according to the data type in Value
    Select Case VarType(Value)
        Case vbInteger, vbLong
            lngValue = Value
            retVal = RegSetValueEx(handle, ValueName, 0, REG_DWORD, lngValue, 4)
        Case VbString
            strValue = Value
            retVal = RegSetValueEx(handle, ValueName, 0, REG_SZ, ByVal strValue, _
                Len(strValue))
        Case vbArray + vbByte
            binValue = Value
            Length = UBound(binValue) - LBound(binValue) + 1
            retVal = RegSetValueEx(handle, ValueName, 0, REG_BINARY, _
                binValue(LBound(binValue)), Length)
        Case Else
            RegCloseKey handle
            Set_locale regionalSymbol
            Err.Raise 1001, ACTIVELOCKSTRING, "Unsupported value type"
    End Select
    
    ' Close the key and signal success
    RegCloseKey handle
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
Public Function SaveString(hKey As Long, strPath As String, strValue As String, strdata As String) As Boolean
    'EXAMPLE:
    '
    'text1.text= savestring(HKEY_CURRENT_USER, "Sof
    '     tware\VBW\Registry", "String", text1.tex
    '     t)
    '
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
    r = RegCloseKey(keyhand)
    If r = 0 Then
        SaveString = True   '"Success"
    Else
        SaveString = False  '"Key to Delete Or Key Not Found"
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
Function Getdword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String)
    'EXAMPLE:
    '
    'text1.text = getdword(HKEY_CURRENT_USER
    '     , "Software\VBW\Registry", "Dword")
    '
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim r As Long
    Dim keyhand As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    ' Get length/data type
    lDataBufSize = 4
    lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

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
Function SaveDword(ByVal hKey As Long, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    'EXAMPLE"
    '
    'Text1.text= SaveDword(HKEY_CURRENT_USER, "Soft
    '     ware\VBW\Registry", "Dword", text1.text)
    '
    '
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(hKey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
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
Public Function DeleteKey(ByVal hKey As Long, ByVal strKey As String)
    'EXAMPLE:
    '
    'Call DeleteKey(HKEY_CURRENT_USER, "Soft
    '     ware\VBW")
    '
    Dim r As Long
    r = RegDeleteKey(hKey, strKey)
    If r = 0 Then
        DeleteKey = "Success"
    Else
        DeleteKey = "No! Key to Delete Or Key Not Found"
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
Public Function CheckRegistryKey(ByVal hKey As Long, ByVal KeyName As String) As Boolean
Dim handle As Long
    ' Try to open the key
    If RegOpenKeyEx(hKey, KeyName, 0, KEY_READ, handle) = 0 Then
        ' The key exists
        CheckRegistryKey = True
        ' Close it before exiting
        RegCloseKey handle
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
Function CreateRegistryKey(ByVal hKey As Long, ByVal KeyName As String) As Boolean
Const REG_OPENED_EXISTING_KEY = &H2
    Dim handle As Long, disposition As Long
    Dim SEC As SECURITY_ATTRIBUTES
    
    If RegCreateKeyEx(hKey, KeyName, 0, 0, 0, 0, SEC, handle, disposition) Then
        Set_locale regionalSymbol
        Err.Raise 1001, ACTIVELOCKSTRING, "Unable to create the registry key"
    Else
        ' Return True if the key already existed.
        CreateRegistryKey = (disposition = REG_OPENED_EXISTING_KEY)
        ' Close the key.
        RegCloseKey handle
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
Public Function DeleteValue(ByVal hKey As Long, ByVal strPath As String, ByVal strValue As String)
    'EXAMPLE:
    '
    'Call DeleteValue(HKEY_CURRENT_USER, "So
    '     ftware\VBW\Registry", "Dword")
    '
    Dim keyhand As Long, r As Long
    r = RegOpenKey(hKey, strPath, keyhand)
    r = RegDeleteValue(keyhand, strValue)
    r = RegCloseKey(keyhand)
    If r = 0 Then
        DeleteValue = "Success"
    Else
        DeleteValue = "No! Value to Delete"
    End If
    
End Function


