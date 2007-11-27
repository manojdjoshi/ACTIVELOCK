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
'*  25.10.07 |  WS    | 1.0            | Initial revision.                      *'
'* ----------+--------+----------------+----------------------------------------*'
'********************************************************************************'
'*              This Module was Created By ActiveLock Wizard V3.5.5             *'
'*                     For Use With Activelock VB2005 V3.5.5                    *'
'********************************************************************************'

'********************************************************************************'
'*  Global Variables                                                            *'
'*------------------------------------------------------------------------------*'
'* structure:                                                                   *'
'*           ActivelockValues                                                   *'                          *'
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
'*        Activelock3_5Net                                                      *'
'*        Microsoft.Visualbasic.Compatibility                                   *'
'*                                                                              *'
'********************************************************************************'
Option Strict Off
Imports ActiveLock3_5NET
Imports System.IO
Imports VB = Microsoft.VisualBasic.Strings
Imports Microsoft.VisualBasic.Compatibility

Module modActiveLockVb2005
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
    End Structure
#End Region '"Structures"

#Region "Public Declare"
    Public ActivelockValues As New ActivelockValues_t
#End Region '"Public Declare"

#Region "Local Declare"
    Private strKeyStorePath As String
    Private noTrialThisTime As Boolean
    Private MyActiveLock As ActiveLock3_5NET._IActiveLock
    Private WithEvents ActiveLockEventSink As ActiveLock3_5NET.ActiveLockEventNotifier
    Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Private Const PUB_KEY As String = "386.391.2CB.21B.210.226.23C.294.386.391.2CB.339.457.533.3B2.42B.4A4.507.457.2AA.294.34F.4C5.44C.507.4A4.507.4F1.2AA.533.2F7.35A.21B.23C.323.39C.462.48E.46D.4E6.512.318.2D6.441.441.4FC.4E6.3D3.3B2.2CB.370.344.205.462.21B.512.344.483.318.2E1.4C5.436.231.210.1D9.51D.23C.247.478.210.441.4F1.436.3BD.2CB.34F.205.441.436.507.273.34F.3D3.4F1.53E.3DE.39C.210.318.1D9.483.226.231.4F1.391.339.339.46D.4AF.4D0.30D.441.4E6.4DB.226.3B2.21B.4C5.3D3.3BD.365.2D6.344.391.35A.365.42B.499.39C.4BA.512.3D3.3B2.3D3.533.34F.35A.302.205.2F7.323.533.1D9.32E.273.344.2F7.3A7.42B.2D6.4F1.48E.2F7.478.226.226.48E.2EC.46D.39C.365.4C5.370.478.370.3A7.252.1D9.533.3A7.478.499.268.507.273.39C.46D.3BD.457.2EC.391.4D0.512.35A.23C.4C5.436.231.252.210.318.4C5.231.365.3D3.44C.512.2CB.53E.247.3DE.478.370.344.323.205.37B.42B.4AF.499.29F.294.205.34F.4C5.44C.507.4A4.507.4F1.2AA.294.2F7.528.4D0.4C5.4BA.457.4BA.4FC.2AA.2CB.37B.2CB.2D6.294.205.2F7.528.4D0.4C5.4BA.457.4BA.4FC.2AA.294.205.386.391.2CB.339.457.533.3B2.42B.4A4.507.457.2AA"
#End Region '"Local Declare"

#Region "Global Routines"

    'This Routine Starts The Protection Of Your Software
    Public Function InitActivelock() As Boolean
        Dim MyAL As New ActiveLock3_5NET.Globals
        Dim strAutoRegisterKeyPath As String
        Dim boolAutoRegisterKeyPath As Boolean
        Dim strMsg As String = Nothing
        Dim A() As String
        On Error GoTo NotRegistered
        CheckForResources("comctl32.ocx", "tabctl32.ocx")
        MyActiveLock = MyAL.NewInstance()
        With MyActiveLock
            .SoftwareName = "TestApp"
            .SoftwareVersion = "3"
            .SoftwarePassword = Chr(99) & Chr(111) & Chr(111) & Chr(108)
            .LicenseKeyType = ActiveLock3_5NET.IActiveLock.ALLicenseKeyTypes.alsRSA
            .TrialType = ActiveLock3_5NET.IActiveLock.ALTrialTypes.trialDays
            .TrialLength = 10
            .TrialHideType = ActiveLock3_5NET.IActiveLock.ALTrialHideTypes.trialHiddenFolder Or ActiveLock3_5NET.IActiveLock.ALTrialHideTypes.trialRegistry
            .SoftwareCode = Dec(PUB_KEY)
            .LockType = IActiveLock.ALLockTypes.lockBIOS Or IActiveLock.ALLockTypes.lockComp Or IActiveLock.ALLockTypes.lockHD Or IActiveLock.ALLockTypes.lockHDFirmware Or IActiveLock.ALLockTypes.lockIP Or IActiveLock.ALLockTypes.lockMAC Or IActiveLock.ALLockTypes.lockMotherboard Or IActiveLock.ALLockTypes.lockWindows
            .AutoRegister = ActiveLock3_5NET.IActiveLock.ALAutoRegisterTypes.alsEnableAutoRegistration
            strAutoRegisterKeyPath = AppPath() & "\" & .SoftwareName & ".all"
            .AutoRegisterKeyPath = strAutoRegisterKeyPath
            If File.Exists(strAutoRegisterKeyPath) Then boolAutoRegisterKeyPath = True
            .CheckTimeServerForClockTampering = ActiveLock3_5NET.IActiveLock.ALTimeServerTypes.alsDontCheckTimeServer
            .CheckSystemFilesForClockTampering = ActiveLock3_5NET.IActiveLock.ALSystemFilesTypes.alsDontCheckSystemFiles
            .LicenseFileType = ActiveLock3_5NET.IActiveLock.ALLicenseFileTypes.alsLicenseFilePlain
            VerifyActiveLockNETdll()
            .KeyStoreType = ActiveLock3_5NET.IActiveLock.LicStoreType.alsFile
            strKeyStorePath = AppPath() & "\" & .SoftwareName & ".lic"
            .KeyStorePath = strKeyStorePath
            ActiveLockEventSink = .EventNotifier
            'Use the following with ASP.NET applications
            'MyActiveLock.Init(Application.StartupPath & "\bin")
            'Use the following with VB.NET applications
            .Init()
            'Or if alcrypto3NET.dll is the same directory
            '.Init(Application.StartupPath, strKeyStorePath)
            .Acquire(strMsg)
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
            ActivelockValues.UsedDaysOrRuns = .UsedDays.ToString
            ActivelockValues.ExpirationDate = .ExpirationDate
            If ActivelockValues.ExpirationDate = "" Then ActivelockValues.ExpirationDate = "Permanent"
            ActivelockValues.RegisteredUser = .RegisteredUser
            ActivelockValues.RegisteredLevel = .RegisteredLevel
            ' Networked Licenses
            If .LicenseClass = "MultiUser" Then
                ActivelockValues.LicenseClass = "Networked"
            Else
                ActivelockValues.LicenseClass = "Single User"
            End If
            ActivelockValues.MaxCount = .MaxCount
            'Read the license file into a string to determine the license type
            Dim strBuff As String
            Dim fNum As Integer
            fNum = FreeFile()
            FileOpen(fNum, strKeyStorePath, OpenMode.Input)
            strBuff = InputString(1, CType(LOF(1), Integer))
            FileClose(fNum)
            If Instring(strBuff, "LicenseType=3") Then
                ActivelockValues.LicenceType = "Time Limited"
            ElseIf Instring(strBuff, "LicenseType=1") Then
                ActivelockValues.LicenceType = "Periodic"
            ElseIf Instring(strBuff, "LicenseType=2") Then
                ActivelockValues.LicenceType = "Permanent"
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
        Dim licFile As String
        licFile = strKeyStorePath
        If File.Exists(licFile) Then
            If FileLen(licFile) <> 0 Then
                Kill(licFile)
                MsgBox("Your license has been killed." & vbCrLf & "You need to get a new license for this application if you want to use it.", MsgBoxStyle.Information)
            Else
                MsgBox("There's no license to kill.", MsgBoxStyle.Information)
            End If
        Else
            MsgBox("There's no license to kill.", MsgBoxStyle.Information)
        End If
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
        Dim fileName As String = System.Windows.Forms.Application.StartupPath & "\ActiveLock3_5NET.dll"
        If File.Exists(fileName) Then
            Dim f As FileStream = New FileStream(System.Windows.Forms.Application.StartupPath & "\ActiveLock3_5NET.dll", FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            crc = c.GetCrc32(CType(f, Stream))
            ' as the most commonly known format
            VerifyActiveLockNETdll = String.Format("{0:X8}", crc)
            f.Close()
            System.Diagnostics.Debug.WriteLine("Hash: " & crc)
            If VerifyActiveLockNETdll <> Dec("268.25D.21B.2F7.2E1.226.2CB.247") Then
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
            If VB.Left(s, 1) = "#" Then
                pathName = VB6.GetPath
                s = Mid(s, 2)
            ElseIf Instring(s, "\") Then
                j = InStrRev(s, "\")
                pathName = VB.Left(s, j - 1)
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
            s = VB6.GetPath & ";" & CurDir() & ";"
            d = WindowsSystemDirectory()
            s = s & d & ";"
            If VB.Right(d, 2) = "32" Then 'I'm guessing at the name of the 16 bit windows directory (assuming it exists)
                i = Len(d)
                s = s & VB.Left(d, i - 2) & ";"
            End If
            s = s & WindowsDirectory() & ";"
            Indx = 1 ' Initialize index to 1.
            Do
                EnvString = Environ(Indx) ' Get environment variable.
                If StrComp(VB.Left(EnvString, 5), "PATH=", CompareMethod.Text) = 0 Then ' Check PATH entry.
                    s = s & Mid(EnvString, 6)
                    Exit Do
                End If
                Indx = Indx + 1
            Loop Until EnvString = ""
            A = Split(s, ";")
        End If
        T = Trim(T)
        If T = "" Then Return Nothing
        If Not Instring(VB.Right(T, 4), ".") Then T = T & ".DLL" 'default extension
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
        WindowsSystemDirectory = LooseSpace(VB.Left(s, cnt))
    End Function

    Private Function WindowsDirectory() As String
        'This function gets the windows directory name
        Dim WinPath As String
        Dim Temp As Object
        WinPath = New String(Chr(0), 145)
        Temp = GetWindowsDirectory(WinPath, 145)
        WindowsDirectory = VB.Left(WinPath, InStr(WinPath, Chr(0)) - 1)
    End Function

    Private Function LooseSpace(ByRef invoer As String) As String
        'This routine terminates a string if it detects char 0.
        Dim P As Integer
        P = InStr(invoer, Chr(0))
        If P <> 0 Then
            LooseSpace = VB.Left(invoer, P - 1)
            Exit Function
        End If
        LooseSpace = invoer
    End Function

#End Region '"Local Routines"

End Module