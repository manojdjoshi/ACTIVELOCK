Attribute VB_Name = "modRegistryAPIs"
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

Option Explicit

Public Const REG_SZ = 1
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
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_CREATE_LINK = &H20
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

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

Public Declare Function RegOpenKey Lib "advapi32.dll" _
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

Public Function DeleteRegKey(lngKey As Long, SubKey As String) As Long
  Dim lngRet As Long
  lngRet = RegDeleteKey(lngKey, SubKey)
  DeleteRegKey = lngRet
End Function

Public Function DeleteRegValue(lngKey As Long, SubKey As String, ValueName As String) As Long
  Dim lngRet As Long
  Dim lngKeyRet As Long
  
  lngRet = RegOpenKeyEx(lngKey, SubKey, 0, KEY_WRITE, lngKeyRet)
  If lngRet <> ERROR_SUCCESS Then Exit Function
  
  lngRet = RegDeleteValue(lngKeyRet, ValueName)
  DeleteRegValue = lngRet
  RegCloseKey lngKeyRet
End Function

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

Public Function ReadRegVal(lngKey As Long, SubKey As String, _
  DataName As String, DefaultData As Variant) As Variant
  Dim lngKeyRet As Long
  Dim lngData As Long
  Dim strdata As String
  Dim Datatype As Long
  Dim DataSize As Long
  Dim lngRet As Long
  ReadRegVal = DefaultData
  lngRet = RegOpenKeyEx(lngKey, SubKey, 0, KEY_QUERY_VALUE, lngKeyRet)
  If lngRet <> ERROR_SUCCESS Then Exit Function
  lngRet = RegQueryValueEx(lngKeyRet, DataName, 0&, Datatype, ByVal 0, DataSize)
  If lngRet <> ERROR_SUCCESS Then
    RegCloseKey lngKeyRet
    Exit Function
  End If
  Select Case Datatype
    Case REG_SZ
      strdata = Space(DataSize + 1)
      lngRet = RegQueryValueEx(lngKeyRet, DataName, 0&, Datatype, ByVal strdata, DataSize)
      If lngRet = ERROR_SUCCESS Then
        ReadRegVal = CVar(StripNulls(RTrim$(strdata)))
      End If
    Case REG_DWORD
      lngRet = RegQueryValueEx(lngKeyRet, DataName, 0&, Datatype, lngData, 4)
      If lngRet = ERROR_SUCCESS Then
        ReadRegVal = CVar(lngData)
      End If
    End Select
  RegCloseKey lngKeyRet
End Function

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

Function StripNulls(ByVal s As String) As String
  Dim i As Integer
  i = InStr(s, Chr$(0))
  If i > 0 Then
    StripNulls = Left$(s, i - 1)
  Else
    StripNulls = s
  End If
End Function

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

Public Sub alSaveSetting(strRegHive As String, strRegPath As String, strAppname As String, strSection As String, strKey As String, vData As Variant)
    alSaveSettingReg strRegHive, strRegPath, strAppname, strSection, strKey, vData
End Sub

Public Function alGetSetting(strRegHive As String, strRegPath As String, strAppname As String, strSection As String, strKey As String, vDefault As Variant) As Variant
    alGetSetting = alGetSettingReg(strRegHive, strRegPath, strAppname, strSection, strKey, vDefault)
End Function

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
