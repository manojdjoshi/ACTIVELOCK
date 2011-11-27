Attribute VB_Name = "modActivelockExcelVBA"
Attribute VB_Description = "FW Addin Library"
Option Explicit
Option Compare Text
Public PROJECT_INI_FILENAME As String
Public testString As String
Public thereIsAProblem As Boolean
Public systemEvent As Boolean
Public Const ADDINSTRING As String = "Activelock Excel VBA Protection"
Public Const ADDIN_VERSION_NUMBER As String = "3.7"
Public path As String
'Public sb As clsProgressBar
Public WindowsVersion As String
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
    (ByVal hWnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, _
    ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function KRN32_GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function KRN32_GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function KRN32_WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias _
        "GetDriveTypeA" (ByVal sDrive As String) As Long
Public Function SetProfileString32(sININame As String, sSection As String, sKeyword As String, vsEntry As String) As Long
    On Error GoTo SetProfileString32_ER
    If IsEmpty(sININame) Or IsEmpty(sSection) Or IsEmpty(sKeyword) Then
    SetProfileString32 = 0
    Exit Function
    End If
    SetProfileString32 = KRN32_WritePrivateProfileString(sSection, sKeyword, vsEntry, sININame)
    Exit Function
SetProfileString32_ER:
    SetProfileString32 = 0
End Function

Public Function ProfileString32(sININame As String, sSection As String, sKeyword As String, vsDefault As String) As String
    Dim sReturnValue As String
    Dim nReturnSize As Long
    Dim nValidSize As Long
    On Error GoTo ProfileString32_ER
    If IsEmpty(sININame) Or IsEmpty(sSection) Or IsEmpty(sKeyword) Then
    ProfileString32 = vsDefault
    Exit Function
    End If
    sReturnValue = Space$(512)
    nReturnSize = Len(sReturnValue)
    nValidSize = KRN32_GetPrivateProfileString(sSection, sKeyword, vsDefault, sReturnValue, nReturnSize, sININame)
    ProfileString32 = Left$(sReturnValue, nValidSize)
    Exit Function
ProfileString32_ER:
    ProfileString32 = vsDefault
End Function

Public Function DriveType(sDrive As String) As String
  Dim sDriveName As String
  Const DRIVE_TYPE_UNDETERMINED = 0
  Const DRIVE_ROOT_NOT_EXIST = 1
  Const DRIVE_REMOVABLE = 2
  Const DRIVE_FIXED = 3
  Const DRIVE_REMOTE = 4
  Const DRIVE_CDROM = 5
  Const DRIVE_RAMDISK = 6
  sDriveName = GetDriveType(sDrive & ":\")
  Select Case sDriveName
      Case DRIVE_TYPE_UNDETERMINED
        DriveType = "unrecognized"
      Case DRIVE_ROOT_NOT_EXIST
        DriveType = "unexisting"
      Case DRIVE_CDROM
        DriveType = "cdrom"
      Case DRIVE_FIXED
        DriveType = "harddrive"
      Case DRIVE_RAMDISK
        DriveType = "ram"
      Case DRIVE_REMOTE
        DriveType = "network"
      Case DRIVE_REMOVABLE
        DriveType = "floppy"
    End Select
End Function
Function GetActivelockMessage() As String
    GetActivelockMessage = "What can protect an Excel VBA function better than Activelock? :)"
End Function

Sub Main()

End Sub
Public Function instring(ByVal x As String, ParamArray MyArray()) As Boolean
'Do ANY of a group of sub-strings appear in within the first string?
'Case doesn't count and we don't care WHERE or WHICH
Dim y As Variant    'member of array that holds all arguments except the first
    For Each y In MyArray
    If InStr(1, x, y, 1) > 0 Then 'the "ones" make the comparison case-insensitive
        instring = True
        Exit Function
    End If
    Next y
End Function
