VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type ALVariables
    Title As String
    SoftwareName As String
    SoftwareVersion As String
    CheckSum As Long
    RegStatus As String
    UsedDays As Integer
    UsedDaysDesc As String
    NetworkLic As String
    LicenseType As String
    User As String
    RegisteredLevel As String
    Expiration As String
    ExpiryMsg As String
    MaxCount As Integer
End Type

Private varALVar As ALVariables

Public MyActiveLock As ActiveLock3.IActiveLock

Private WithEvents ActiveLockEventSink As ActiveLockEventNotifier
Attribute ActiveLockEventSink.VB_VarHelpID = -1


' Trial mode variables
Dim noTrialThisTime As Boolean 'ialkan - needed for registration while form was loaded via trial
Dim expireTrialLicense As Boolean
Dim strKeyStorePath As String
Dim strAutoRegisterKeyPath As String

'Application name used
Private Const LICENSE_ROOT As String = "ActiveTest"

' The following declarations are used by the IsDLLAvailable function
' provided by the Activelock user Pinheiro
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const MAX_MESSAGE_LENGTH = 512

'Windows and System directory API
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Splash Screen API
Private Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Const GWW_HWNDPARENT = (-8)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Function GetLicenseRoot() As String
    GetLicenseRoot = LICENSE_ROOT
End Function

Public Function GetTrialNoThisTime() As Boolean
    GetTrialNoThisTime = noTrialThisTime
End Function

Public Function SetTrialNoThisTime(anoTrialThisTime As Boolean)
    noTrialThisTime = anoTrialThisTime
End Function

Public Function GetTitle() As String
    GetTitle = varALVar.Title
End Function

Public Function GetVersion() As String
    GetVersion = varALVar.SoftwareVersion
End Function

Public Function GetSoftwareName() As String
    GetSoftwareName = varALVar.SoftwareName
    
End Function

Public Function GetCheckSum() As Long
    GetCheckSum = varALVar.CheckSum
End Function

Public Function GetRegStatus() As String
    GetRegStatus = varALVar.RegStatus
End Function

Public Function GetUsedDays() As Integer
    GetUsedDays = varALVar.UsedDays
End Function

Public Function GetUsedDaysDesc() As String
    GetUsedDaysDesc = varALVar.UsedDaysDesc
End Function
'new in AL 3.5
Public Function GetNetworkLic() As String
    GetNetworkLic = varALVar.NetworkLic
End Function

Public Function GetLicenseType() As String
    GetLicenseType = varALVar.LicenseType
End Function

Public Function GetExpiration() As String
    GetExpiration = varALVar.Expiration
End Function

Public Function GetExpiryMsg() As String
    GetExpiryMsg = varALVar.ExpiryMsg
End Function

Public Function GetUser() As String
    GetUser = varALVar.User
End Function

Public Function GetRegisteredLevel() As String
    GetRegisteredLevel = varALVar.RegisteredLevel
End Function

Public Function SetRegisteredLevel(RegLevel As String)
    varALVar.RegisteredLevel = RegLevel
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

Private Function WindowsDirectory() As String

'This function gets the windows directory name
Dim WinPath As String
Dim Temp
    
WinPath = String(145, Chr(0))
Temp = GetWindowsDirectory(WinPath, 145)
WindowsDirectory = Left(WinPath, InStr(WinPath, Chr(0)) - 1)

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

Public Sub Init()
    Dim autoRegisterKey As String
    Dim boolAutoRegisterKeyPath As Boolean
    Dim Msg As String
    Dim A() As String
    
    On Error GoTo NotRegistered
   
    ' Check the existence of necessary files to run this application
    
    '**** SATYA Comment this during debugging and uncomment while compiling...
'    Call CheckForResources("Alcrypto3.dll", "comctl32.ocx", "tabctl32.ocx")

    ' Check if the Activelock3.dll is registered. If not no need to continue.
    If CheckIfDLLIsRegistered = False Then End
    
    On Error GoTo NotRegistered
    ' Obtain AL instance and initialize its properties
    Set MyActiveLock = ActiveLock3.NewInstance()
    
    With MyActiveLock
        
        .SoftwareName = LICENSE_ROOT
        varALVar.SoftwareName = .SoftwareName
        
        ' Note: Do not use (App.Major & "." & App.Minor & "." & App.Revision)
        ' since the license will fail with version incremented exe builds
        'SATYA
        .SoftwareVersion = "2.0"   ' WARNING *** WARNING *** DO NOT USE App.Major & "." & App.Minor & "." & App.Revision
        varALVar.SoftwareVersion = .SoftwareVersion
                        
        ' New in v3.3
        ' This should be set to protect yourself against ResetTrial abuse
        .SoftwarePassword = Chr(99) & Chr(111) & Chr(111) & Chr(108)
        
        ' New in v3.5
        '.LicenseKeyType = alsShortKeyMD5
        .LicenseKeyType = alsRSA
        
        ' New in v3.1 - Trial Feature
        .TrialType = trialDays
        .TrialLength = 15
        If .TrialType <> trialNone And .TrialLength = 0 Then
            ' Do Nothing
            ' In such cases Activelock automatically generates errors -11001100 or -11001101
            ' to indicate that you're using the trial feature but, trial length was not specified
        End If
        
        ' Uncomment the following statement to use a certain trial data hiding technique
        ' Use OR to combine one or more trial hiding techniques
        ' or don't use this property to use ALL techniques
        .TrialHideType = trialHiddenFolder Or trialRegistryPerUser Or trialSteganography
        
        .SoftwareCode = "RSA1024BgIAAAAkAABSU0ExAAQAAAEAAQCNPM4jUn/wgmJ69OiMxkZYJ+0Fy6gvnc29B1Zwm7xlf3dS+AM629Pd32l4S5xlOjIbrOVs1IRsSrqI/1y45yVPn8m+mWFQR6OQ75PFC6Mg8vvsrOepD244Ll2X5/6u5b5ds8deNTTR0ec/DIhayPFCOkw3YpuHnkUuo5lY+UT4yQ=="
        'ActiveTest 2.0 software code
        '.SoftwareCode = Dec(PUB_KEY)
                
        .LockType = lockNone  'lockWindows 'Or lockComp 'Or lockComp Or lockWindows
        strAutoRegisterKeyPath = App.Path & "\" & LICENSE_ROOT & ".all"
        .AutoRegister = alsEnableAutoRegistration
        .AutoRegisterKeyPath = strAutoRegisterKeyPath
        If FileExist(strAutoRegisterKeyPath) Then boolAutoRegisterKeyPath = True
        
        .CheckTimeServerForClockTampering = alsDontCheckTimeServer      ' use alsCheckTimeServer to enforce time server check for clock tampering detection
        .CheckSystemFilesForClockTampering = alsDontCheckSystemFiles    ' use alsCheckSystemFiles to enforce system files scanning for clock tampering detection
        .LicenseFileType = alsLicenseFileEncrypted
        
    End With

    ' Verify AL's authenticity
    varALVar.CheckSum = modMain.VerifyActiveLockdll()

    ' Initialize the keystore. We use a File keystore in this case.
    MyActiveLock.KeyStoreType = alsFile
    
    ' Path to the license file
    strKeyStorePath = App.Path & "\" & LICENSE_ROOT & ".lic"
    MyActiveLock.KeyStorePath = strKeyStorePath
    
    ' Obtain the EventNotifier so that we can receive notifications from AL.
    Set ActiveLockEventSink = MyActiveLock.EventNotifier
    
    ' Initialize AL
    MyActiveLock.Init autoRegisterKey
    If FileExist(strKeyStorePath) And boolAutoRegisterKeyPath = True And autoRegisterKey <> "" Then
        ' This means, an ALL file existed and was used to create a LIC file
        ' Init() method successfully registered the ALL file
        ' and returned the license key
        ' You can process that key here to see if there is any abuse, etc.
        ' ie. whether the key was used before, etc.
    End If

    ' Initialize other application settings
    varALVar.SoftwareVersion = MyActiveLock.SoftwareVersion
'    cboSpeed = cboSpeed.List(2)

    ' Check registration status
    Dim strMsg As String
        
    
    MyActiveLock.Acquire strMsg
    
    If strMsg <> "" Then 'There's a trial
        A = Split(strMsg, vbCrLf)
        varALVar.RegStatus = A(0)
        varALVar.UsedDaysDesc = A(1)
        varALVar.ExpiryMsg = strMsg
        
        'SATYA : Need to futher Check this.. code....
'        FunctionalitiesEnabled = True
        frmSplash.lblInfo.Caption = vbCrLf & strMsg
        frmSplash.Show
        frmSplash.Refresh
        'Dim rtn As Long 'declare the need variables
        'rtn = SetWindowWord(frmSplash.hWnd, GWW_HWNDPARENT, Me.hWnd) 'let both forms load together
        Sleep 3000 'wait about 3 seconds
        Unload frmSplash
        
        varALVar.LicenseType = "Free Trial"
        frmMain.Show
        Exit Sub
    Else
'        cmdKillTrial.Visible = False
'        cmdResetTrial.Visible = False
    End If
        
    varALVar.RegStatus = "Registered"
    varALVar.UsedDays = MyActiveLock.UsedDays
    varALVar.Expiration = MyActiveLock.ExpirationDate
    
    If varALVar.Expiration = "" Then varALVar.Expiration = "Permanent" 'App has a permanent license
    
    varALVar.User = MyActiveLock.RegisteredUser
    varALVar.RegisteredLevel = MyActiveLock.RegisteredLevel
    
   ' Networked Licenses
    If MyActiveLock.LicenseClass = "MultiUser" Then
        varALVar.NetworkLic = "NetWorked"
'        txtNetworkedLicense.Text = "Networked"
    Else
        varALVar.NetworkLic = "Single User"
'        txtNetworkedLicense.Text = "Single User"
'        txtMaxCount.Visible = False
'        lblConcurrentUsers.Visible = False
    End If
    
    'txtMaxCount.Text = MyActiveLock.MaxCount
    varALVar.MaxCount = MyActiveLock.MaxCount
    
    'Read the license file into a string to determine the license type
    Dim strBuff As String
    Dim fNum As Integer
    fNum = FreeFile
    Open strKeyStorePath For Input As #fNum
    strBuff = Input(LOF(1), 1)
    Close #fNum
    If Instring(strBuff, "LicenseType=3") Then
        varALVar.LicenseType = "Time Limited"
    ElseIf Instring(strBuff, "LicenseType=1") Then
        varALVar.LicenseType = "Periodic"
    ElseIf Instring(strBuff, "LicenseType=2") Then
        varALVar.LicenseType = "Permanent"
    End If
    
    'You can code the below function to enable or disable features...
'    FunctionalitiesEnabled = True
    frmMain.Show
    Exit Sub

NotRegistered:
    'SATYA : I didnot code FunctionalitiesEnabled function,
    'You can code the below function to enable or disable features...
'    FunctionalitiesEnabled = False
    If Instring(Err.Description, "no valid license") = False And noTrialThisTime = False Then
        MsgBox Err.Number & ": " & Err.Description
    End If
    
    varALVar.RegStatus = Err.Description
    varALVar.LicenseType = "None"
        
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation
    End If
        'when no valid license exists, show Registration form
        frmReg.Show vbModal
    Exit Sub

DLLnotRegistered:
    End
End Sub

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
    Dim Y As Variant
    Dim i As Integer, j As Integer
    Dim s As String, systemDir As String, pathName As String

    WhereIsDLL ("") 'initialize

    systemDir = WindowsSystemDirectory 'Get the Windows system directory
    For Each Y In MyArray
        foundIt = False
        s = CStr(Y)
        
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
    Next Y

    CheckForResources = True
Exit Function

checkForResourcesError:
    MsgBox "CheckForResources error", vbCritical, "Error"
    End   'an error kills the program
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

Function WhereIsDLL(ByVal T As String) As String
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

Function FileExist(ByVal TestFileName As String) As Boolean
    'This function checks for the existance of a given
    'file name. The function returns a TRUE or FALSE value.
    'The more complete the TestFileName string is, the
    'more reliable the results of this function will be.

    'Declare local variables
Dim OK As Integer

    'Set up the error handler to trap the File Not Found
    'message, or other errors.
    On Error GoTo FileExistErrors:

    'Check for attributes of test file. If this function
    'does not raise an error, than the file must exist.
    OK = GetAttr(TestFileName)

    'If no errors encountered, then the file must exist
    FileExist = True
Exit Function

FileExistErrors:    'error handling routine, including File Not Found
    FileExist = False
    Exit Function 'end of error handler
End Function

Function Instring(ByVal X As String, ParamArray MyArray()) As Boolean
'Do ANY of a group of sub-strings appear in within the first string?
'Case doesn't count and we don't care WHERE or WHICH
Dim Y As Variant    'member of array that holds all arguments except the first
    For Each Y In MyArray
    If InStr(1, X, Y, 1) > 0 Then 'the "ones" make the comparison case-insensitive
        Instring = True
        Exit Function
    End If
    Next Y
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Key Validation Functionalities
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''
' ActiveLock raises this event typically when it needs a value to be encrypted.
' We can use any kind of encryption we'd like here, as long as it's deterministic.
' i.e. there's a one-to-one correspondence between unencrypted value and encrypted value.
' NOTE: BlowFish is NOT an example of deterministic encryption so you can't use it here.
Private Sub ActiveLockEventSink_ValidateValue(ByRef Value As String)
    Value = Encrypt(Value)
End Sub

Private Function Encrypt(strdata As String) As String
    Dim i&, n&
    Dim sResult$
    n = Len(strdata)
    For i = 1 To n
        sResult = sResult & Asc(Mid$(strdata, i, 1)) * 7
    Next i
    Encrypt = sResult
End Function

Private Function CheckIfDLLIsRegistered() As Boolean
    Dim strDllPath As String
    Dim Result As Boolean
        
    CheckIfDLLIsRegistered = True
    
    strDllPath = GetTypeLibPathFromObject()
    Result = IsDLLAvailable(strDllPath)
    If Result Then
        ' MsgBox "Activelock3.dll is Registered !"
        ' Just quietly proceed
    Else
        MsgBox "Activelock3.dll is Not Registered!"
        CheckIfDLLIsRegistered = False
    End If

End Function


Public Function fResetTrial()
    Screen.MousePointer = vbHourglass
    MyActiveLock.ResetTrial
    MyActiveLock.ResetTrial ' DO NOT REMOVE, NEED TO CALL TWICE
    Screen.MousePointer = vbDefault
    MsgBox "Free Trial has been Reset." & vbCrLf & _
        "You'll need to restart the application for a new Free Trial.", vbInformation
End Function




