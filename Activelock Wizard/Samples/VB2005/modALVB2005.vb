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

'********************************************************************************'
'*  File     : modActiveLockVb2005.vb                                           *'
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
'*  28.03.09 |  WS    | 3.6            | added Activelock 3.6 Features          *'
'* ----------+--------+----------------+----------------------------------------*'
'********************************************************************************'
'*              This Module was Created By ActiveLock Wizard V3.6               *'
'*                     For Use With Activelock VB2005/8 V3.6                    *'
'********************************************************************************'

'********************************************************************************'
'*  Global Variables                                                            *'
'*------------------------------------------------------------------------------*'
'* structure:                                                                   *'
'*           ActivelockValues                                                   *'
'*                                                                              *'
'********************************************************************************'

'********************************************************************************'
'*  Global Functions                                                            *'
'*------------------------------------------------------------------------------*'
'*  InitActivelock()                                                            *'
'*  GetTheInstallationCode()                                                    *'
'*  KillTheLic()                                                                *'
'*  KillTheTrial()                                                              *'
'*  ResetTheTrial()                                                             *'
'*  RegisterTheApplication()                                                    *'
'*                                                                              *'
'********************************************************************************'

'********************************************************************************'
'* General Notes                                                                *'
'*------------------------------------------------------------------------------*'
'* Please Add These References to your Project                                  *'
'*        Activelock3_6Net                                                      *'
'*        Microsoft.Visualbasic.Compatibility                                   *'
'*                                                                              *'
'********************************************************************************'
Option Strict Off
Imports ActiveLock3_6NET
Imports System.IO

Module modActiveLockVb2008
#Region "Structures"
    Public Structure ActivelockValues_t
        Public AppName As String
        Public AppVersion As String
        Public RegStatus As String
        Public UsedDaysOrRuns As String
        Public ValidTrial As Boolean
        Public LicenceType As String
        Public ExpirationDate As String
        Public RegisteredUser As String
        Public RegisteredLevel As String
        Public LicenseClass As String
        Public MaxCount As Integer
        Public Checksum As String
    End Structure
#End Region '"Structures"

#Region "Public Declare"
    Public ActivelockValues As New ActivelockValues_t
#End Region '"Public Declare"

#Region "Local Declare"
    Private Enum LicFlags
        alfSingle = 0
        alfMulti = 1
    End Enum
    Private strKeyStorePath As String
    Private noTrialThisTime As Boolean
    Private MyActiveLock As ActiveLock3_6NET._IActiveLock
    Private WithEvents ActiveLockEventSink As ActiveLock3_6NET.ActiveLockEventNotifier
    Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As IntPtr, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean
    Private Const MyDLLName As String = "ActiveLock3_6Net.dll"
    Private Const CrcDataEnc As String = "25D.268.23C.21B.25D.247.2E1.21B"
    Private Const PUB_KEY As String = "386.391.2CB.21B.210.226.23C.294.386.391.2CB.339.457.533.3B2.42B.4A4.507.457.2AA.294.34F.4C5.44C.507.4A4.507.4F1.2AA.533.318.533.370.4C5.2F7.462.512.318.3C8.23C.4C5.3D3.2CB.4D0.4E6.252.247.2EC.323.2CB.51D.441.4C5.23C.273.3B2.483.436.391.3DE.252.34F.533.457.35A.4BA.499.247.441.323.2F7.302.2CB.436.3D3.39C.46D.302.533.30D.46D.4FC.210.3D3.457.4C5.4FC.34F.4FC.2E1.51D.499.436.226.499.252.512.302.2D6.25D.4F1.4A4.3B2.344.344.344.339.533.205.441.247.318.462.268.478.4FC.323.25D.32E.3DE.4BA.4BA.3BD.4E6.32E.4DB.1D9.302.51D.365.3A7.457.34F.370.4A4.323.499.32E.34F.51D.252.386.268.323.252.370.3C8.34F.2F7.323.44C.533.512.441.462.53E.3A7.3A7.2EC.370.247.2CB.273.318.512.53E.21B.210.2D6.483.4DB.4E6.4E6.2EC.231.21B.1D9.441.478.2D6.4DB.39C.37B.21B.3D3.231.4BA.365.21B.39C.210.2CB.499.528.3B2.30D.441.4E6.3B2.210.29F.294.205.34F.4C5.44C.507.4A4.507.4F1.2AA.294.2F7.528.4D0.4C5.4BA.457.4BA.4FC.2AA.2CB.37B.2CB.2D6.294.205.2F7.528.4D0.4C5.4BA.457.4BA.4FC.2AA.294.205.386.391.2CB.339.457.533.3B2.42B.4A4.507.457.2AA"
#End Region '"Local Declare"

#Region "Global Routines"

    'This Routine Starts The Protection Of Your Software
    Public Function InitActivelock() As Boolean
        Dim MyAL As New ActiveLock3_6NET.Globals
        Dim strAutoRegisterKeyPath As String
        Dim boolAutoRegisterKeyPath As Boolean
        Dim strMsg As String = Nothing
        Dim A() As String
        On Error GoTo NotRegistered
        'CheckForResources("comctl32.ocx", "tabctl32.ocx")
        Call CheckForResources("#ActiveLock3_6Net.dll")
        Dim Ret As Long
        Dim AppfilePath As String
        AppfilePath = Space$(260)
        Ret = SHGetSpecialFolderPath(0, AppfilePath, 46, False) ' 46 is for ...\All Users\Documents folder.
        If AppfilePath.Trim <> Chr(0) Then
            AppfilePath = Strings.Left(AppfilePath, InStr(AppfilePath, Chr(0)) - 1)
        End If
        MyActiveLock = MyAL.NewInstance()
        With MyActiveLock
            .SoftwareName = "WalterApp"
            .SoftwareVersion = "1.0.1"
            .SoftwarePassword = Chr(99) & Chr(111) & Chr(111) & Chr(108)
            .LicenseKeyType = ActiveLock3_6NET.IActiveLock.ALLicenseKeyTypes.alsRSA
            .TrialType = ActiveLock3_6NET.IActiveLock.ALTrialTypes.trialRuns
            .TrialLength = 3
            .TrialHideType = ActiveLock3_6NET.IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or ActiveLock3_6NET.IActiveLock.ALTrialHideTypes.trialSteganography
            .SoftwareCode = Dec(PUB_KEY)
            .LockType = IActiveLock.ALLockTypes.lockFingerprint
            .AutoRegister = ActiveLock3_6NET.IActiveLock.ALAutoRegisterTypes.alsDisableAutoRegistration
            If Directory.Exists(AppfilePath & " \ " & .SoftwareName & .SoftwareVersion) = False Then
                MkDir(AppfilePath & " \ " & .SoftwareName & .SoftwareVersion)
            End If
            strAutoRegisterKeyPath = AppfilePath & "\" & .SoftwareName & .SoftwareVersion & "\" & .SoftwareName & .SoftwareVersion & ".all"
            .AutoRegisterKeyPath = strAutoRegisterKeyPath
            If File.Exists(strAutoRegisterKeyPath) Then boolAutoRegisterKeyPath = True
            .CheckTimeServerForClockTampering = ActiveLock3_6NET.IActiveLock.ALTimeServerTypes.alsDontCheckTimeServer
            .CheckSystemFilesForClockTampering = ActiveLock3_6NET.IActiveLock.ALSystemFilesTypes.alsDontCheckSystemFiles
            .LicenseFileType = ActiveLock3_6NET.IActiveLock.ALLicenseFileTypes.alsLicenseFileEncrypted
            ActivelockValues.Checksum = VerifyActiveLockNETdll()
            .KeyStoreType = ActiveLock3_6NET.IActiveLock.LicStoreType.alsFile
            strKeyStorePath = AppPath() & "\" & .SoftwareName & ".lic"
            .KeyStorePath = strKeyStorePath
            ActiveLockEventSink = .EventNotifier
            'Use the following with ASP.NET applications
            'MyActiveLock.Init(Application.StartupPath & "\bin")
            'Use the following with VB.NET applications
            '.Init()
            'Or if alcrypto3NET.dll is the same directory
            .Init(Application.StartupPath, strKeyStorePath)
            Dim strRemainingTrialDays As String = Nothing
            Dim strRemainingTrialRuns As String = Nothing
            Dim strTrialLength As String = Nothing
            Dim strUsedDays As String = Nothing
            Dim strExpirationDate As String = Nothing
            Dim strRegisteredUser As String = Nothing
            Dim strRegisteredLevel As String = Nothing
            Dim strLicenseClass As String = Nothing
            Dim strMaxCount As String = Nothing
            Dim strLicenseFileType As String = Nothing
            Dim strLicenseType As String = Nothing
            Dim strUsedLockType As String = Nothing
            .Acquire(strMsg, strRemainingTrialDays, strRemainingTrialRuns, strTrialLength, strUsedDays, strExpirationDate, strRegisteredUser, strRegisteredLevel, strLicenseClass, strMaxCount, strLicenseFileType, strLicenseType)
            If strMsg <> "" Then 'There's a trial
                A = Split(strMsg, vbCrLf)
                ActivelockValues.RegStatus = A(0)
                ActivelockValues.UsedDaysOrRuns = A(1)
                ActivelockValues.ValidTrial = True
                ActivelockValues.LicenceType = "Free Trial"
                Return True
            Else
                ActivelockValues.ValidTrial = False
                ActivelockValues.LicenceType = "No Trial"
            End If
            ActivelockValues.LicenceType = "Registered"
            ActivelockValues.UsedDaysOrRuns = strUsedDays
            ActivelockValues.ExpirationDate = strExpirationDate
            If ActivelockValues.ExpirationDate = "" Then ActivelockValues.ExpirationDate = "Permanent"
            ActivelockValues.RegisteredUser = .RegisteredUser
            ActivelockValues.AppName = MyActiveLock.SoftwareName
            ActivelockValues.AppVersion = MyActiveLock.SoftwareVersion
            ActivelockValues.RegisteredLevel = strRegisteredLevel
            ' Networked Licenses
            If strLicenseClass = LicFlags.alfMulti.ToString Then
                ActivelockValues.LicenseClass = "Networked"
            Else 'If strLicenseType = LicFlags.alfSingle Then
                ActivelockValues.LicenseClass = "Single User"
            End If
            ActivelockValues.MaxCount = strMaxCount
            'determine the license type
            If strLicenseType = "allicTimeLocked" Then
                ActivelockValues.LicenceType = "Time Limited"
            ElseIf strLicenseType = "allicPeriodic" Then
                ActivelockValues.LicenceType = "Periodic"
            ElseIf strLicenseType = "allicPermanent" Then
                ActivelockValues.LicenceType = "Permanent"
            Else
                ActivelockValues.LicenceType = "None"
            End If
        End With
        Return True
NotRegistered:
        If Instring(Err.Description, "no valid license") = False And noTrialThisTime = False Then
            MsgBox(Err.Number & ": " & Err.Description)
        End If
        ActivelockValues.RegStatus = Err.Description
        ActivelockValues.LicenceType = "None"
        If strMsg <> "" Then
            MsgBox(strMsg, MsgBoxStyle.Information)
        End If
        Return False
    End Function

    'This Routine Retrieves the Install Code That Is needed To Request The Activation Code
    Public Function GetTheInstallationCode(ByVal Username As String) As String
        Return MyActiveLock.InstallationCode(Username)
    End Function

    'Kills The Current Licence of the program
    Public Sub KillTheLic()
        MsgBox("This feature is not intended to be sent to the end user." & vbCrLf & _
       "Because it removes all traces of a license from a machine." & vbCrLf & vbCrLf & _
       "The best way to terminate a license in an end-user's machine is to" & vbCrLf & _
       "just delete his LIC file and send him a new build of your app with" & vbCrLf & _
       "a new revision number (or upgrade in other terms).", vbInformation)
        Dim licFile As String
        licFile = strKeyStorePath
        Try
            MyActiveLock.KillLicense(MyActiveLock.SoftwareName & MyActiveLock.SoftwareVersion, licFile)

            ' Use the following in an end-user's machine.
            'File.Delete(strKeyStorePath)

            MsgBox("Your license has been killed." & vbCrLf & "You need to get a new license for this application if you want to use it.", MsgBoxStyle.Information)
        Catch ex As Exception
            MsgBox("There's no license to kill.", MsgBoxStyle.Information)
        End Try
    End Sub

    'Kills The Trial Period for the program
    Public Sub KillTheTrial()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        MyActiveLock.KillTrial()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Free Trial has been Killed." & vbCrLf & "There will be no more Free Trial next time you start this application." & vbCrLf & vbCrLf & "You must register this application for further use.", MsgBoxStyle.Information)
    End Sub

    'Resets The Trial Period
    Public Sub ResetTheTrial()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        MyActiveLock.ResetTrial()
        MyActiveLock.ResetTrial() ' DO NOT REMOVE, NEED TO CALL TWICE
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Free Trial has been Reset.", MsgBoxStyle.Information)
    End Sub

    'Registers The Program with the Activation Code
    Public Function RegisterTheApplication(ByVal LibKeyIn As String, ByVal Username As String) As Boolean
        Dim ok As Boolean
        On Error GoTo errHandler
        If Mid(LibKeyIn, 5, 1) = "-" And Mid(LibKeyIn, 10, 1) = "-" And Mid(LibKeyIn, 15, 1) = "-" And Mid(LibKeyIn, 20, 1) = "-" Then
            MyActiveLock.Register(LibKeyIn, Username) 'YOU MUST SPECIFY THE USER NAME WITH SHORT KEYS !!!
        Else ' ALCRYPTO RSA
            MyActiveLock.Register(LibKeyIn)
        End If
        Return True
        Exit Function
errHandler:
        Return False
    End Function

#End Region '"Global Routines"

#Region "Local Routines"

    Private Function VerifyActiveLockNETdll() As String
        ' CRC32 Hash...
        ' I have modified this routine to read the crc32
        ' of the Activelock3NET.dll directly
        ' since the assembly is not a COM object anymore
        ' the method below is very suitable for .NET and more appropriate
        Dim c As New CRC32
        Dim crc As Integer = 0
        Dim fileName As String = AppPath() & "\" & MyDLLName
        If File.Exists(fileName) Then
            Dim f As FileStream = New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            crc = c.GetCrc32(f)
            ' as the most commonly known format
            VerifyActiveLockNETdll = String.Format("{0:X8}", crc)
            f.Close()
            System.Diagnostics.Debug.WriteLine("Hash: " & crc)
            If VerifyActiveLockNETdll <> Dec(CrcDataEnc) Then
                ' Encrypted version of "activelock3NET.dll has been corrupted. If you were running a real application, it should terminate at this point."
                MsgBox(Dec("42B.441.4FC.483.512.457.4A4.4C5.441.499.231.35A.2F7.39C.1FA.44C.4A4.4A4.160.478.42B.4F1.160.436.457.457.4BA.160.441.4C5.4E6.4E6.507.4D0.4FC.457.44C.1FA"))
                End
            End If
        Else
            MsgBox(Dec("42B.441.4FC.483.512.457.4A4.4C5.441.499.231.35A.2F7.39C.1FA.44C.4A4.4A4.160.4BA.4C5.4FC.160.462.4C5.507.4BA.44C"))
            End
        End If
    End Function

    Private Function CheckForResources(ByVal ParamArray MyArray() As Object) As Boolean
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
        Dim y As Object
        Dim i As Short
        Dim j As Integer
        Dim systemDir, s, pathName As String
        WhereIsDLL("") 'initialize
        systemDir = WindowsSystemDirectory() 'Get the Windows system directory
        For Each y In MyArray
            foundIt = False
            s = CStr(y)
            If Strings.Left(s, 1) = "#" Then
                pathName = AppPath()
                s = Mid(s, 2)
            ElseIf Instring(s, "\") Then
                j = InStrRev(s, "\")
                pathName = Strings.Left(s, j - 1)
                s = Mid(s, j + 1)
            Else
                pathName = systemDir
            End If
            If Instring(s, ".") Then
                If File.Exists(pathName & "\" & s) Then foundIt = True
            ElseIf File.Exists(pathName & "\" & s & ".DLL") Then
                foundIt = True
            ElseIf File.Exists(pathName & "\" & s & ".OCX") Then
                foundIt = True
                s = s & ".OCX" 'this will make the softlocx check easier
            End If
            If Not foundIt Then
                MsgBox(s & " could not be found in " & pathName & vbCrLf & System.Reflection.Assembly.GetExecutingAssembly.GetName.Name & " cannot run without this library file!" & vbCrLf & vbCrLf & "Exiting!", MsgBoxStyle.Critical, "Missing Resource")
                End
            End If
        Next y
        CheckForResources = True
        Exit Function
checkForResourcesError:
        MsgBox("CheckForResources error", MsgBoxStyle.Critical, "Error")
        End 'an error kills the program
    End Function

    Private Function Enc(ByRef strdata As String) As String
        Dim i, n As Integer
        Dim sResult As String = Nothing
        n = Len(strdata)
        Dim l As Integer
        For i = 1 To n
            l = Asc(Mid(strdata, i, 1)) * 11
            If sResult = "" Then
                sResult = Hex(l)
            Else
                sResult = sResult & "." & Hex(l)
            End If
        Next i
        Enc = sResult
    End Function

    Private Function Dec(ByRef strdata As String) As String
        Dim arr() As String = Nothing
        arr = Split(strdata, ".")
        Dim sRes As String = Nothing
        Dim i As Integer
        For i = LBound(arr) To UBound(arr)
            sRes = sRes & Chr(CInt("&h" & arr(i)) / 11)
        Next
        Dec = sRes
    End Function

    Private Function AppPath() As String
        Return System.Windows.Forms.Application.StartupPath
    End Function

    Private Function Instring(ByVal x As String, ByVal ParamArray MyArray() As Object) As Boolean
        'Do ANY of a group of sub-strings appear in within the first string?
        'Case doesn't count and we don't care WHERE or WHICH
        Dim y As String 'member of array that holds all arguments except the first
        For Each y In MyArray
            If InStr(1, x, y, CompareMethod.Text) > 0 Then 'the "ones" make the comparison case-insensitive
                Instring = True
                Exit Function
            End If
        Next y
    End Function

    Private Function WhereIsDLL(ByVal T As String) As String
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
        Static A As Object
        Dim s, d As String
        Dim EnvString As String
        Dim Indx As Short ' Declare variables.
        Dim i As Short
        On Error Resume Next
        i = UBound(A)
        If i = 0 Then
            s = AppPath()
            d = WindowsSystemDirectory()
            s = s & ";" & d & ";"
            If Strings.Right(d, 2) = "32" Then 'I'm guessing at the name of the 16 bit windows directory (assuming it exists)
                i = Len(d)
                s = s & Strings.Left(d, i - 2) & ";"
            End If
            s = s & ";" & WindowsDirectory() & ";"
            Indx = 1 ' Initialize index to 1.
            Do
                EnvString = Environ(Indx) ' Get environment variable.
                If StrComp(Strings.Left(EnvString, 5), "PATH=", CompareMethod.Text) = 0 Then ' Check PATH entry.
                    s = s & Mid(EnvString, 6)
                    Exit Do
                End If
                Indx = Indx + 1
            Loop Until EnvString = ""
            A = Split(s, ";")
        End If
        T = Trim(T)
        If T = "" Then Return Nothing
        If Not Instring(Strings.Right(T, 4), ".") Then T = T & ".DLL" 'default extension
        For i = 0 To UBound(A)
            If File.Exists(A(i) & "\" & T) Then
                WhereIsDLL = A(i)
                Exit Function
            End If
        Next i
        Return Nothing
    End Function

    Private Function WindowsSystemDirectory() As String
        Dim cnt As Integer
        Dim s As String
        Dim dl As Integer
        cnt = 254
        s = New String(Chr(0), 254)
        dl = GetSystemDirectory(s, cnt)
        WindowsSystemDirectory = LooseSpace(Strings.Left(s, cnt))
    End Function

    Private Function WindowsDirectory() As String
        'This function gets the windows directory name
        Dim WinPath As String
        Dim Temp As Object
        WinPath = New String(Chr(0), 145)
        Temp = GetWindowsDirectory(WinPath, 145)
        WindowsDirectory = Strings.Left(WinPath, InStr(WinPath, Chr(0)) - 1)
    End Function

    Private Function LooseSpace(ByRef invoer As String) As String
        'This routine terminates a string if it detects char 0.
        Dim P As Integer
        P = InStr(invoer, Chr(0))
        If P <> 0 Then
            LooseSpace = Strings.Left(invoer, P - 1)
            Exit Function
        End If
        LooseSpace = invoer
    End Function

#End Region '"Local Routines"

End Module
