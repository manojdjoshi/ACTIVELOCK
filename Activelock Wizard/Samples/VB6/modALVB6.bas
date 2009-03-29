Attribute VB_Name = "modALVB6"
'********************************************************************************'
'*   ActiveLock                                                                 *'
'*   Copyright 1998-2002 Nelson Ferraz                                          *'
'*   Copyright 2003-2007 The ActiveLock Software Group (ASG)                    *'
'*   All material is the property of the contributing authors.                  *'
'*                                                                              *'
'*   Redistribution and use in source and binary forms, with or without         *'
'*   modification, are permitted provided that the following conditions are     *'
'*   met:                                                                       *'
'*                                                                              *'
'*     [o] Redistributions of source code must retain the above copyright       *'
'*         notice, this list of conditions and the following disclaimer.        *'
'*                                                                              *'
'*     [o] Redistributions in binary form must reproduce the above              *'
'*         copyright notice, this list of conditions and the following          *'
'*         disclaimer in the documentation and/or other materials provided      *'
'*         with the distribution.                                               *'
'*                                                                              *'
'*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS        *'
'*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT          *'
'*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR      *'
'*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT       *'
'*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,      *'
'*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT           *'
'*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,      *'
'*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY      *'
'*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT        *'
'*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE      *'
'*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.       *'
'*                                                                              *'
'********************************************************************************'
'*                                                                              *'
'********************************************************************************'
'*  File     : modALVb6.bas                                                     *'
'*------------------------------------------------------------------------------*'
'*  Object   : implement all functionalities of ActiveLock in your application  *'
'* -----------------------------------------------------------------------------*'
'*  Date     | Author | Revision       | Description                            *'
'* ----------+--------+----------------+----------------------------------------*'
'*  25.10.07 |  WS    | 1.0.1          | Initial revision.                      *'
'* ----------+--------+----------------+----------------------------------------*'
'*  01.01.08 |  WS    | 1.0.2          | added VCode Decoding.                  *'
'* ----------+--------+----------------+----------------------------------------*'
'*  05.01.09 |  ZP    | 1.0.3          | added CRC Check.                       *'
'* ----------+--------+----------------+----------------------------------------*'
'*  28.03.09 |  WS    | 1.0.4          | added Activelock 3.6 Features          *'
'* ----------+--------+----------------+----------------------------------------*'
'********************************************************************************'
'*              This Module was Created By ActiveLock Wizard V1.0.4             *'
'*                     For Use With Activelock VB6 V3.6                         *'
'********************************************************************************'

'********************************************************************************'
'*  Global Variables                                                            *'
'*------------------------------------------------------------------------------*'
'*  Struct  ActivelockValues     Contains All The Activelock Info You Need      *'
'*          MyActiveLock         Interface To The Activelock DLL                *'
'*                                                                              *'
'********************************************************************************'

'********************************************************************************'
'*  Global Functions                                                            *'
'*------------------------------------------------------------------------------*'
'*  InitActivelock()            Initialize The ActiveLock Protection            *'
'*  RegisterTheApplication()    Register The Application                        *'
'*  KillTheLic()                Kill The License                                *'
'*  KillTheTrial()              Kill The Trial                                  *'
'*  ResetTheTrial()             Resets the Trial                                *'
'*                                                                              *'
'********************************************************************************'

'********************************************************************************'
'* General Notes                                                                *'
'*------------------------------------------------------------------------------*'
'* Please Add These References to your Project                                  *'
'*        Activelock3_6                                                         *'
'*                                                                              *'
'********************************************************************************'

Option Explicit
   Type ActivelockValues_t
         RegStatus As String
         UsedDaysOrRuns As String
         ValidTrial As Boolean
         LicenceType As String
         ExpirationDate As String
         RegisteredUser As String
         RegisteredLevel As String
         LicenseClass As String
         MaxCount As Integer
End Type

Private Declare Function MapFileAndCheckSum Lib "imagehlp" Alias "MapFileAndCheckSumA" (ByVal FileName As String, HeaderSum As Long, CheckSum As Long) As Long

' The following declarations are used by the IsDLLAvailable function
' provided by the Activelock user Pinheiro
Dim noTrialThisTime As Boolean
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
'************************
Public Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

'************************

'Windows and System directory API
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public MyActiveLock As ActiveLock3.IActiveLock

Public ActivelockValues As ActivelockValues_t
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim expireTrialLicense As Boolean
Dim strKeyStorePath As String
Dim strAutoRegisterKeyPath As String
Private Const CrcPartValue As Long = 711000
Private Const PUB_KEY As String = "386.391.2CB.21B.210.226.23C.2D6.46D.323.2CB.2CB.2CB.2CB.499.2CB.2CB.2D6.391.3A7.210.2F7.528.2CB.2CB.37B.2CB.2CB.2CB.2F7.2CB.2CB.37B.2E1.2D6.441.51D.339.4D0.273.4D0.21B.4E6.499.37B.478.3DE.533.528.39C.457.386.2F7.478.273.2F7.34F.226.25D.231.51D.2D6.4DB.4D0.365.2EC.391.1D9.339.344.499.457.507.205.528.4BA.205.37B.210.1D9.51D.268.507.478.2D6.37B.462.3DE.4DB.268.462.205.4F1.44C.386.462.344.1D9.391.247.344.323.2D6.323.4FC.365.512.210.4C5.499.30D.23C.226.4AF.339.441.42B.35A.499.210.4AF.386.4C5.391.4C5.32E.499.3A7.370.34F.4E6.386.4D0.23C.231.48E.268.4BA.2CB.339.3B2.2E1.4AF.386.4E6.2EC.25D.4D0.3DE.21B.2EC.4C5.231.507.318.34F.23C.391.48E.533.44C.344.3D3.210.4C5.4AF.457.339.268.512.365.2E1.533.3DE.226.3BD.4A4.44C.3D3.533.210.323.528.51D.46D.37B.48E.35A.4DB.34F.318.344.507.30D.21B.4E6.2E1.46D.21B.2EC.386.23C.1D9.247.46D.29F.29F"

Public Function InitActivelock() As Boolean
   Dim autoRegisterKey As String
   Dim boolAutoRegisterKeyPath As Boolean
   Dim Msg As String
   Dim A() As String
   
   On Error GoTo DLLnotRegistered
   ' Check the existence of necessary files to run this application
   Call CheckForResources("Alcrypto3.dll", "comctl32.ocx", "tabctl32.ocx")
   ' Check if the Activelock3.dll is registered. If not no need to continue.
   If CheckIfDLLIsRegistered = False Then End
   On Error GoTo NotRegistered
   ' Obtain AL instance and initialize its properties
   Set MyActiveLock = ActiveLock3.NewInstance()
   With MyActiveLock
       .SoftwareName = "WalterTest"
       .SoftwareVersion = "1"
       .SoftwarePassword = Chr(99) & Chr(111) & Chr(111) & Chr(108)
       .LicenseKeyType = alsRSA
       .TrialType = trialRuns
       .TrialLength = 10
       .TrialHideType = trialHiddenFolder Or trialRegistryPerUser
       .SoftwareCode = Dec(PUB_KEY)
       .LockType = lockFingerprint
       strAutoRegisterKeyPath = App.Path & "\" & .SoftwareName & ".all"
       .AutoRegister = alsDisableAutoRegistration
       .AutoRegisterKeyPath = strAutoRegisterKeyPath
       If FileExist(strAutoRegisterKeyPath) Then boolAutoRegisterKeyPath = True
       ' use alsCheckTimeServer to enforce time server check for clock tampering detection
       .CheckTimeServerForClockTampering = alsDontCheckTimeServer
       ' use alsCheckSystemFiles to enforce system files scanning for clock tampering detection
       .CheckTimeServerForClockTampering = alsDontCheckSystemFiles
       .LicenseFileType = alsLicenseFileEncrypted
   ' Verify AL's authenticity
   VerifyActiveLockdll
   ' Initialize the keystore. We use a File keystore in this case.
   MyActiveLock.KeyStoreType = alsFile
   ' Path to the license file
   strKeyStorePath = App.Path & "\" & .SoftwareName & ".lic"
   MyActiveLock.KeyStorePath = strKeyStorePath
   End With
   ' Obtain the EventNotifier so that we can receive notifications from AL.
   Set frmMain.ActiveLockEventSink = MyActiveLock.EventNotifier
   ' Initialize AL
   MyActiveLock.Init autoRegisterKey
   If FileExist(strKeyStorePath) And boolAutoRegisterKeyPath = True And autoRegisterKey <> "" Then
       ' This means, an ALL file existed and was used to create a LIC file
       ' Init() method successfully registered the ALL file
       ' and returned the license key
       ' You can process that key here to see if there is any abuse, etc.
       ' ie. whether the key was used before, etc.
   End If
   ' Check registration status
   Dim strMsg As String
   Dim strRemainingTrialDays As String
   Dim strRemainingTrialRuns As String
   Dim strTrialLength As String
   Dim strUsedDays As String
   Dim strExpirationDate As String
   Dim strRegisteredUser As String
   Dim strRegisteredLevel As String
   Dim strLicenseClass As String
   Dim strMaxCount As String
   Dim strLicenseFileType As String
   Dim strLicenseType As String
   Dim strUsedLockType As String
   MyActiveLock.Acquire strMsg, strRemainingTrialDays, strRemainingTrialRuns, strTrialLength, strUsedDays, strExpirationDate, strRegisteredUser, strRegisteredLevel, strLicenseClass, strMaxCount, strLicenseFileType, strLicenseType, strUsedLockType
   If strMsg <> "" Then 'There's a trial
       A = Split(strMsg, vbCrLf)
       ActivelockValues.RegStatus = A(0)
       ActivelockValues.UsedDaysOrRuns = A(1)
       ' You can also get the RemainingTrialDays or RemainingTrialRuns directly by:
       'txtUsedDays.Text = MyActiveLock.RemainingTrialDays Or MyActiveLock.RemainingTrialRuns
       ActivelockValues.ExpirationDate = "No Valid Licence"
       ActivelockValues.RegisteredUser = "No Valid Licence"
       ActivelockValues.RegisteredLevel = "No Valid Licence"
       ActivelockValues.LicenseClass = "No Valid Licence"
       ActivelockValues.ValidTrial = True
       ActivelockValues.LicenceType = "Free Trial"
       InitActivelock = True
       ActivelockValues.LicenceType = "Free Trial"
       InitActivelock = True
       Exit Function
   Else
       'This should never happen (it should be caught by ErrorTrap)
   End If
   ' Uncomment the following to retrieve the usedlocktypes
   ' Dim aa() As ActiveLock3.ALLockTypes
   ' ReDim aa(UBound(MyActiveLock.UsedLockType))
   ' aa = MyActiveLock.UsedLockType
   ' MsgBox aa(0) 'For example, if only lockHDfirmware was used, this will return 256
   ActivelockValues.RegStatus = "Registered"
   ActivelockValues.UsedDaysOrRuns = strUsedDays
   ActivelockValues.ExpirationDate = strExpirationDate
   If ActivelockValues.ExpirationDate = "" Then ActivelockValues.ExpirationDate = "Permanent"  'App has a permanent license
   ActivelockValues.RegisteredUser = strRegisteredUser
   ActivelockValues.RegisteredLevel = strRegisteredLevel
   ' Networked Licenses
   If strLicenseClass = "0" Then  'MultiUser
       ActivelockValues.LicenseClass = "Networked"
   Else
       ActivelockValues.LicenseClass = "Single User"
   End If
   ActivelockValues.MaxCount = strMaxCount
   'Determine The License Type
   Select Case strLicenseType
       Case 3
           ActivelockValues.LicenceType = "Time Limited"
       Case 2
           ActivelockValues.LicenceType = "Permanent"
       Case 1
           ActivelockValues.LicenceType = "Periodic"
   End Select
   InitActivelock = True
   ActivelockValues.ValidTrial = True
   Exit Function
NotRegistered:
   'FunctionalitiesEnabled = False
   If Instring(Err.Description, "no valid license") = False And noTrialThisTime = False Then
       MsgBox Err.Number & ": " & Err.Description
   End If
   ActivelockValues.LicenceType = "None"
   If strMsg <> "" Then
       MsgBox strMsg, vbInformation
   End If
   Exit Function
DLLnotRegistered:
   End
End Function


Function CheckForResources(ParamArray MyArray()) As Boolean
'MyArray is a list of things to check
'These can be DLLs or OCXs

'Files, by default, are searched for in the Windows System Directory
'Exceptions;
'   Begins with a # means it should be in the same directory with the executable
'   Contains the full path (anything with a "\")

'Typical names would be "#aaa.dll", "mydll.dll", "myocx.ocx", "comdlg32.ocx", "mscomctl.ocx", "msflxgrd.ocx"

'If the file has no extension, we;
'     assume it's a DLL, and if it still can't be found
'     assume it's an OCX

On Error GoTo checkForResourcesError
Dim foundIt As Boolean
Dim y As Variant
Dim i As Integer, j As Integer
Dim s As String, systemDir As String, pathName As String

WhereIsDLL ("") 'initialize

systemDir = WindowsSystemDirectory 'Get the Windows system directory
For Each y In MyArray
    foundIt = False
    s = CStr(y)
    
    If Left$(s, 1) = "#" Then
        pathName = App.Path
        s = Mid$(s, 2)
    ElseIf Instring(s, "\") Then
        j = InStrRev(s, "\")
        pathName = Left$(s, j - 1)
        s = Mid$(s, j + 1)
    Else
        pathName = systemDir
    End If
    
    If Instring(s, ".") Then
        If FileExist(pathName & "\" & s) Then foundIt = True
    ElseIf FileExist(pathName & "\" & s & ".DLL") Then
        foundIt = True
    ElseIf FileExist(pathName & "\" & s & ".OCX") Then
        foundIt = True
        s = s & ".OCX" 'this will make the softlocx check easier
    End If
    
    If Not foundIt Then
        MsgBox s & " could not be found in " & pathName & vbCrLf & _
        App.Title & " cannot run without this library file!" & vbCrLf & vbCrLf & "Exiting!", vbCritical, "Missing Resource"
        End
    End If
Next y

CheckForResources = True
Exit Function

checkForResourcesError:
    MsgBox "CheckForResources error", vbCritical, "Error"
    End   'an error kills the program
End Function
' Simple encrypt of a string
Public Function Enc(strdata As String) As String
    Dim i&, n&
    Dim sResult$
    n = Len(strdata)
    Dim l As Long
    For i = 1 To n
        l = Asc(Mid$(strdata, i, 1)) * 11
        If sResult = "" Then
            sResult = Hex(l)
        Else
            sResult = sResult & "." & Hex(l)
        End If
    Next i
    Enc = sResult
End Function

Public Function Dec(strdata As String) As String
    Dim arr() As String
    arr = Split(strdata, ".")
    Dim sRes As String
    Dim i&
    For i = LBound(arr) To UBound(arr)
        sRes = sRes & Chr$(CLng("&h" & arr(i)) / 11)
    Next
    Dec = sRes
End Function

Public Function Encrypt(strdata As String) As String
    Dim i&, n&
    Dim sResult$
    n = Len(strdata)
    For i = 1 To n
        sResult = sResult & Asc(Mid$(strdata, i, 1)) * 7
    Next i
    Encrypt = sResult
End Function

Function FileExist(ByVal TestFileName As String) As Boolean
'This function checks for the existance of a given
'file name. The function returns a TRUE or FALSE value.
'The more complete the TestFileName string is, the
'more reliable the results of this function will be.

'Declare local variables
Dim ok As Integer

'Set up the error handler to trap the File Not Found
'message, or other errors.
On Error GoTo FileExistErrors:

'Check for attributes of test file. If this function
'does not raise an error, than the file must exist.
ok = GetAttr(TestFileName)

'If no errors encountered, then the file must exist
FileExist = True
Exit Function

FileExistErrors:    'error handling routine, including File Not Found
    FileExist = False
    Exit Function 'end of error handler
End Function

Function Instring(ByVal x As String, ParamArray MyArray()) As Boolean
'Do ANY of a group of sub-strings appear in within the first string?
'Case doesn't count and we don't care WHERE or WHICH
Dim y As Variant    'member of array that holds all arguments except the first
    For Each y In MyArray
    If InStr(1, x, y, 1) > 0 Then 'the "ones" make the comparison case-insensitive
        Instring = True
        Exit Function
    End If
    Next y
End Function

Public Function WhereIsDLL(ByVal T As String) As String
'Places where programs look for DLLs
'   1 directory containing the EXE
'   2 current directory
'   3 32 bit system directory   possibly \Windows\system32
'   4 16 bit system directory   possibly \Windows\system
'   5 windows directory         possibly \Windows
'   6 path

'The current directory may be changed in the course of the program
'but current directory -- when the program starts -- is what matters
'so a call should be made to this function early on to "lock" the paths.

'Add a call at the beginning of checkForResources

Static A As Variant
Dim s As String, d As String
Dim EnvString As String, Indx As Integer  ' Declare variables.
Dim i As Integer

On Error Resume Next
i = UBound(A)
If i = 0 Then
    s = App.Path & ";" & CurDir & ";"
    
    d = WindowsSystemDirectory
    s = s & d & ";"
    
    If Right$(d, 2) = "32" Then   'I'm guessing at the name of the 16 bit windows directory (assuming it exists)
        i = Len(d)
        s = s & Left$(d, i - 2) & ";"
    End If
    
    s = s & WindowsDirectory & ";"
    Indx = 1   ' Initialize index to 1.
    Do
        EnvString = Environ(Indx)   ' Get environment variable.
        If StrComp(Left(EnvString, 5), "PATH=", vbTextCompare) = 0 Then ' Check PATH entry.
            s = s & Mid$(EnvString, 6)
            Exit Do
        End If
        Indx = Indx + 1
    Loop Until EnvString = ""
    A = Split(s, ";")
End If

T = Trim(T)
If T = "" Then Exit Function
If Not Instring(Right$(T, 4), ".") Then T = T & ".DLL"   'default extension
For i = 0 To UBound(A)
    If FileExist(A(i) & "\" & T) Then
        WhereIsDLL = A(i)
        Exit Function
    End If
Next i

End Function
Private Function WindowsSystemDirectory() As String

Dim cnt As Long
Dim s As String
Dim dl As Long

cnt = 254
s = String$(254, 0)
dl = GetSystemDirectory(s, cnt)
WindowsSystemDirectory = LooseSpace(Left$(s, cnt))

End Function

Public Function LooseSpace(invoer$) As String
'This routine terminates a string if it detects char 0.

Dim P As Long

P = InStr(invoer$, Chr(0))
If P <> 0 Then
    LooseSpace$ = Left$(invoer$, P - 1)
    Exit Function
End If
LooseSpace$ = invoer$

End Function

Private Function CheckIfDLLIsRegistered() As Boolean
Dim strDllPath As String
Dim Result As Boolean
    
CheckIfDLLIsRegistered = True

strDllPath = GetTypeLibPathFromObject()
Result = IsDLLAvailable(strDllPath)
If Result Then
    ' MsgBox "Activelock3.6.dll is Registered !"
    ' Just quietly proceed
Else
    MsgBox "Activelock3.6.dll is Not Registered!"
    CheckIfDLLIsRegistered = False
End If

End Function

Public Function GetTypeLibPathFromObject() As String
    Dim strDllPath As String
    GetTypeLibPathFromObject = WinSysDir() & "\activelock3.6.dll"
End Function

Private Function IsDLLAvailable(ByVal DllFilename As String) As Boolean
' Code provided by Activelock user Pinheiro
Dim hModule As Long
hModule = LoadLibrary(DllFilename) 'attempt to load DLL
If hModule > 32 Then
    FreeLibrary hModule 'decrement the DLL usage counter
    IsDLLAvailable = True 'Return true
Else
    IsDLLAvailable = False 'Return False
End If
End Function
Public Function CRCCheckSumTypeLib() As Long
    Dim strDllPath As String
    strDllPath = GetTypeLibPathFromObject()
    Dim HeaderSum As Long, RealSum As Long
    MapFileAndCheckSum strDllPath, HeaderSum, RealSum
    CRCCheckSumTypeLib = RealSum
    
End Function

Public Function WinSysDir() As String
    Const FIX_LENGTH% = 4096
    Dim Length As Integer
    Dim Buffer As String * FIX_LENGTH

    Length = GetSystemDirectory(Buffer, FIX_LENGTH - 1)
    WinSysDir = Left$(Buffer, Length)
End Function



' Returns the expected CRC value of ActiveLock3.dll
'
Private Property Get Value() As Long
    Value = CrcPartValue + 121
End Property


Public Function VerifyActiveLockdll() As String
    Dim crc As Long
    crc = CRCCheckSumTypeLib()
    'Debug.Print "Hash: " & crc
    If crc <> Value Then
        ' Encrypted version of "Activelock DLL has been corrupted." If you were running a real application, it should terminate at this point.
        MsgBox Dec("2CB.441.4FC.483.512.457.4A4.4C5.441.499.160.2EC.344.344.160.478.42B.4F1.160.436.457.457.4BA.160.441.4C5.4E6.4E6.507.4D0.4FC.457.44C.1FA"), vbExclamation
        End
    End If
    VerifyActiveLockdll = CStr(crc)
End Function

'not used'Private Function Encrypt(strdata As String) As String
 '   Dim i&, n&
'    Dim sResult$
'    n = Len(strdata)
'    For i = 1 To n
'        sResult = sResult & Asc(Mid$(strdata, i, 1)) * 7
'    Next i
'    Encrypt = sResult
'End Function

Private Function WindowsDirectory() As String
'This function gets the windows directory name
Dim WinPath As String
Dim Temp
WinPath = String(145, Chr(0))
Temp = GetWindowsDirectory(WinPath, 145)
WindowsDirectory = Left(WinPath, InStr(WinPath, Chr(0)) - 1)
End Function

Public Sub ResetTheTrial()
Screen.MousePointer = vbHourglass
MyActiveLock.ResetTrial
MyActiveLock.ResetTrial 'You Have To Call Twice
Screen.MousePointer = vbDefault
MsgBox "Free Trial has been Reset." & vbCrLf & _
"You'll need to restart the application for a new Free Trial.", vbInformation
End Sub

Public Sub KillTheTrial()
Screen.MousePointer = vbHourglass
MyActiveLock.KillTrial
Screen.MousePointer = vbDefault
MsgBox "Free Trial has been Killed." & vbCrLf & _
"There will be no more Free Trial next time you start this application." & vbCrLf & vbCrLf & _
"You must register this application for further use.", vbInformation
End Sub

Public Sub KillTheLic()
    Dim licFile As String
    licFile = App.Path & "\" & MyActiveLock.SoftwareName & ".lic"
    If FileExist(licFile) Then
        If FileLen(licFile) <> 0 Then
            Kill licFile
            MsgBox "Your license has been killed." & vbCrLf & _
                "You need to get a new license for this application if you want to use it.", vbInformation
        Else
            MsgBox "There's no license to kill.", vbInformation
        End If
    Else
        MsgBox "There's no license to kill.", vbInformation
    End If
    frmMain.Form_Load
End Sub


Public Function RegisterTheApplication(ByVal InstallCode As String, Optional ByVal UserName As String = "test") As Boolean
    Dim ok As Boolean, LibKey As String
    On Error GoTo errHandler
    If Mid(InstallCode, 5, 1) = "-" And Mid(InstallCode, 10, 1) = "-" And Mid(InstallCode, 15, 1) = "-" And Mid(InstallCode, 20, 1) = "-" Then
        MyActiveLock.Register InstallCode, UserName 'YOU MUST SPECIFY THE USER NAME WITH SHORT KEYS !!!
    Else    ' ALCRYPTO RSA
        MyActiveLock.Register InstallCode
    End If
    MsgBox Dec("386.457.46D.483.4F1.4FC.4E6.42B.4FC.483.4C5.4BA.160.4F1.507.441.441.457.4F1.4F1.462.507.4A4.16B"), vbInformation ' "Registration successful!"
    RegisterTheApplication = True
    Exit Function
errHandler:
    RegisterTheApplication = False
    'MsgBox Err.Number & ": " & Err.Description
End Function


