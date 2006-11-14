VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALTestApp - ActiveLock3 Test Application"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9720
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   9720
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7740
      Width           =   9720
      _ExtentX        =   17145
      _ExtentY        =   450
      Style           =   1
      SimpleText      =   "Ready"
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7755
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13679
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   16711680
      TabCaption(0)   =   "Registration"
      TabPicture(0)   =   "frmMain.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraRegStatus"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraReg"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Sample App"
      TabPicture(1)   =   "frmMain.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraViewport"
      Tab(1).Control(1)=   "Frame1"
      Tab(1).Control(2)=   "lblTrialInfo"
      Tab(1).Control(3)=   "lblLockStatus2"
      Tab(1).Control(4)=   "lblLockStatus"
      Tab(1).ControlCount=   5
      Begin VB.Frame fraViewport 
         BorderStyle     =   0  'None
         Height          =   525
         Left            =   -74950
         TabIndex        =   30
         Top             =   5510
         Width           =   7850
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   1875
         Left            =   -74850
         TabIndex        =   21
         Top             =   1440
         Width           =   7545
         Begin VB.OptionButton optForm 
            Caption         =   "Option 2"
            Height          =   225
            Index           =   1
            Left            =   3030
            TabIndex        =   27
            Top             =   1245
            Width           =   1395
         End
         Begin VB.OptionButton optForm 
            Caption         =   "Option 1"
            Height          =   225
            Index           =   0
            Left            =   3030
            TabIndex        =   26
            Top             =   990
            Value           =   -1  'True
            Width           =   1395
         End
         Begin VB.ComboBox cboSpeed 
            Height          =   315
            ItemData        =   "frmMain.frx":0D02
            Left            =   5070
            List            =   "frmMain.frx":0D15
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   30
            Width           =   1635
         End
         Begin VB.CheckBox chkPause 
            Caption         =   "Checkbox for Level 3 only"
            Height          =   315
            Left            =   0
            TabIndex        =   24
            Top             =   390
            Width           =   2565
         End
         Begin VB.CheckBox chkFlash 
            Caption         =   "Checkbox for ALL Levels"
            Height          =   495
            Left            =   0
            TabIndex        =   23
            Top             =   690
            Width           =   2610
         End
         Begin VB.CheckBox chkScroll 
            Caption         =   "Checkbox for ALL Levels"
            Height          =   225
            Left            =   0
            TabIndex        =   22
            Top             =   90
            Value           =   1  'Checked
            Width           =   2535
         End
         Begin VB.Label lblHost 
            Caption         =   "Option Buttons for ALL Levels:"
            Height          =   255
            Left            =   3060
            TabIndex        =   29
            Top             =   690
            Width           =   2280
         End
         Begin VB.Label lblSpeed 
            Caption         =   "Activated with Level 4 Only"
            Height          =   255
            Left            =   3030
            TabIndex        =   28
            Top             =   90
            Width           =   2025
         End
      End
      Begin VB.Frame fraReg 
         Caption         =   "Register"
         ForeColor       =   &H00FF0000&
         Height          =   4380
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   9495
         Begin VB.CommandButton cmdPaste 
            Height          =   345
            Left            =   8220
            Picture         =   "frmMain.frx":0D3F
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   2070
            Width           =   345
         End
         Begin VB.CommandButton cmdCopy 
            Height          =   345
            Left            =   8220
            MaskColor       =   &H8000000F&
            Picture         =   "frmMain.frx":1081
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   900
            Width           =   345
         End
         Begin VB.CommandButton cmdKillLicense 
            Caption         =   "&Kill License"
            Height          =   315
            Left            =   8220
            TabIndex        =   43
            ToolTipText     =   "Kill the License"
            Top             =   2880
            Width           =   1095
         End
         Begin VB.TextBox txtUser 
            Height          =   285
            Left            =   1440
            TabIndex        =   7
            Text            =   "Evaluation User"
            Top             =   300
            Width           =   6675
         End
         Begin VB.CommandButton cmdReqGen 
            Caption         =   "&Generate"
            Height          =   315
            Left            =   8220
            TabIndex        =   9
            ToolTipText     =   "Generate Installation Code"
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox txtReqCodeGen 
            BackColor       =   &H80000013&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1080
            Left            =   1440
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   600
            Width           =   6675
         End
         Begin VB.CommandButton cmdRegister 
            Caption         =   "&Register"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8220
            TabIndex        =   11
            ToolTipText     =   "Register the License"
            Top             =   2475
            Width           =   1095
         End
         Begin VB.TextBox txtLibKeyIn 
            BackColor       =   &H80000013&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2565
            Left            =   1440
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   10
            Top             =   1710
            Width           =   6675
         End
         Begin VB.Label Label13 
            Caption         =   "User Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   1335
         End
         Begin VB.Label Label11 
            Caption         =   "Installation Code:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Liberation Key:"
            Height          =   255
            Left            =   135
            TabIndex        =   18
            Top             =   1710
            Width           =   1335
         End
      End
      Begin VB.Frame fraRegStatus 
         Caption         =   "Status"
         ForeColor       =   &H00FF0000&
         Height          =   2745
         Left            =   120
         TabIndex        =   13
         Top             =   420
         Width           =   9495
         Begin VB.TextBox txtMaxCount 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   5355
            Locked          =   -1  'True
            TabIndex        =   51
            Top             =   2340
            Width           =   540
         End
         Begin VB.TextBox txtNetworkedLicense 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   2340
            Width           =   1890
         End
         Begin VB.TextBox txtLicenseType 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   4140
            Locked          =   -1  'True
            TabIndex        =   45
            Top             =   2040
            Width           =   1755
         End
         Begin VB.CommandButton cmdResetTrial 
            Caption         =   "&Reset Trial"
            Height          =   315
            Left            =   8280
            TabIndex        =   42
            ToolTipText     =   "Reset the Free Trial"
            Top             =   1320
            Width           =   1095
         End
         Begin VB.CommandButton cmdKillTrial 
            Caption         =   "&Kill Trial"
            Height          =   315
            Left            =   8280
            TabIndex        =   41
            ToolTipText     =   "End the Free Trial"
            Top             =   1740
            Width           =   1095
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   825
            Left            =   8370
            Picture         =   "frmMain.frx":1393
            ScaleHeight     =   825
            ScaleWidth      =   825
            TabIndex        =   39
            Top             =   210
            Width           =   825
         End
         Begin VB.TextBox txtRegisteredLevel 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   36
            Top             =   1740
            Width           =   4335
         End
         Begin VB.TextBox txtChecksum 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   2040
            Width           =   1215
         End
         Begin VB.TextBox txtVersion 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   2
            Text            =   "1.0"
            Top             =   540
            Width           =   4335
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   1
            Text            =   "TestApp"
            Top             =   240
            Width           =   4335
         End
         Begin VB.TextBox txtExpiration 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   5
            Top             =   1440
            Width           =   4335
         End
         Begin VB.TextBox txtUsedDays 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   1140
            Width           =   4335
         End
         Begin VB.TextBox txtRegStatus 
            BackColor       =   &H80000013&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   840
            Width           =   4335
         End
         Begin VB.Label lblConcurrentUsers 
            Caption         =   "No. of Concurrent Users:"
            Height          =   255
            Left            =   3555
            TabIndex        =   50
            Top             =   2385
            Width           =   1785
         End
         Begin VB.Label Label10 
            Caption         =   "License Class:"
            Height          =   255
            Left            =   135
            TabIndex        =   49
            Top             =   2340
            Width           =   1335
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "License Type:"
            Height          =   255
            Left            =   2940
            TabIndex        =   44
            Top             =   2070
            Width           =   1155
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Caption         =   "Activelock V3"
            ForeColor       =   &H00FF0000&
            Height          =   165
            Left            =   8250
            TabIndex        =   40
            Top             =   1050
            Width           =   1065
         End
         Begin VB.Label Label5 
            Caption         =   "Registered Level:"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   1770
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "DLL Checksum:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   2070
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "App Version:"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   570
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "App Name:"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   270
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Expiry Date:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1470
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Days Used:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1170
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "License Status:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   870
            Width           =   1380
         End
      End
      Begin VB.Label lblTrialInfo 
         Caption         =   "NOTE: All application functionalities are available in Trial Mode."
         Height          =   375
         Left            =   -74910
         TabIndex        =   38
         Top             =   870
         Width           =   4545
      End
      Begin VB.Label lblLockStatus2 
         Caption         =   "Disabled"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -71880
         TabIndex        =   34
         Top             =   480
         Width           =   4515
      End
      Begin VB.Label lblLockStatus 
         Caption         =   "Application Functionalities Are Currently: "
         Height          =   375
         Left            =   -74880
         TabIndex        =   33
         Top             =   480
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
'*   All material is the property of the contributing authors.
'*
'*   Redistribution and use in source and binary forms, with or without
'*   modification, are permitted provided that the following conditions are
'*   met:
'*
'*     [o] Redistributions of source code must retain the above copyright
'*         notice, this list of conditions and the following disclaimer.
'*
'*     [o] Redistributions in binary form must reproduce the above
'*         copyright notice, this list of conditions and the following
'*         disclaimer in the documentation and/or other materials provided
'*         with the distribution.
'*
'*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
'*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
'*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
'*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'*
'*

''
' This test app is used to exercise all functionalities of ActiveLock.
'
' If you're running this from inside VB and would like to bypass dll-checksumming,
' Add the following compilation flag to your Project Properties (Make tab)
'   AL_DEBUG = 1
'
' @author th2tran
' @version 2.0.0
' @date 20030715

'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
'
'   [ ] Re: GetMACAndUserFromRequestCode(), try to move the decoding of the
'       request code inside ActiveLock.  We need to abstract this, if possible,
'       such that the client app doesn't need to understand how the request
'       code was encoded.
'
'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
'
'   07.31.03 - th2tran       - Now performing checksum on ActiveLock2.dll.
'   08.01.03 - wizzardme2000 - LockTypes other than MAC are now supported
'   08.03.03 - th2tran       - Added SoftwareCode generator and usage instructions.
'   09.14.03 - th2tran       - Removed Key Generator functionality. This is now handled
'                              by ALUGEN.
'   10.13.03 - th2tran       - Added simple Encrypt routine to illustrate handling of
'                              ActiveLockEventNotifier.ValidateValue() event.
'   11.02.03 - th2tran       - Store message box messages in encrypted format to elude hex editors.
'                            - txtLibKeyIn is now MultiLine-enabled
'                            - Terminology change: RequestCode is now known as InstallationCode
'   04.17.04 - th2tran       - Added IActiveLock.Init() call--this is now required.
'                            - Set AutoRegisterKeyPath property (new in 2.0.5) to automatically
'                              register liberation file upon startup (if it exists).
'   07.11.04 - th2tran       - Changed liberation file to testapp.all
'                            - Update txtUser upon successful Acquire()
'   07.19.04 - th2tran       - Re-implemented cmdReqGen_Click() to use ActiveLock.InstallationCode()
'   09.17.05 - ialkan        - v3 updates. Major release.
'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////

Option Explicit

Private MyActiveLock As ActiveLock3.IActiveLock
Private WithEvents ActiveLockEventSink As ActiveLockEventNotifier
Attribute ActiveLockEventSink.VB_VarHelpID = -1

' Trial mode variables
Dim noTrialThisTime As Boolean 'ialkan - needed for registration while form was loaded via trial
Dim expireTrialLicense As Boolean
Dim strKeyStorePath As String
Dim strAutoRegisterKeyPath As String

'Application name used
Const LICENSE_ROOT As String = "TestApp"

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

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtReqCodeGen.Text
End Sub

Private Sub cmdKillLicense_Click()
    Dim licFile As String
    licFile = App.Path & "\" & LICENSE_ROOT & ".lic"
    If FileExist(licFile) Then
        If FileLen(licFile) <> 0 Then
            Kill licFile
            MsgBox "Your license has been killed." & vbCrLf & _
                "You need to get a new license for this application if you want to use it.", vbInformation
                txtUsedDays.Text = ""
                txtExpiration.Text = ""
                txtRegisteredLevel.Text = ""
                txtNetworkedLicense.Text = ""
                txtMaxCount.Text = ""
        Else
            MsgBox "There's no license to kill.", vbInformation
        End If
    Else
        MsgBox "There's no license to kill.", vbInformation
    End If
    Form_Load
    cmdResetTrial.Visible = True
End Sub
Private Sub cmdKillTrial_Click()
Screen.MousePointer = vbHourglass
MyActiveLock.KillTrial
Screen.MousePointer = vbDefault
MsgBox "Free Trial has been Killed." & vbCrLf & _
    "There will be no more Free Trial next time you start this application." & vbCrLf & vbCrLf & _
    "You must register this application for further use.", vbInformation
txtRegStatus.Text = "Free Trial has been Killed"
txtUsedDays.Text = ""
txtExpiration.Text = ""
txtRegisteredLevel.Text = ""
txtLicenseType.Text = "None"
txtNetworkedLicense.Text = ""
txtMaxCount.Text = ""

End Sub

Private Sub cmdPaste_Click()
If Clipboard.GetText = txtReqCodeGen.Text Then
    MsgBox "You cannot paste the Installation Code into the Liberation Key field.", vbExclamation
    Exit Sub
End If
txtLibKeyIn.Text = Clipboard.GetText
End Sub

Private Sub cmdResetTrial_Click()
Screen.MousePointer = vbHourglass
MyActiveLock.ResetTrial
MyActiveLock.ResetTrial ' DO NOT REMOVE, NEED TO CALL TWICE
Screen.MousePointer = vbDefault
MsgBox "Free Trial has been Reset." & vbCrLf & _
    "You'll need to restart the application for a new Free Trial.", vbInformation
txtRegStatus.Text = "Free Trial has been Reset"
txtUsedDays.Text = ""
txtExpiration.Text = ""
txtRegisteredLevel.Text = ""
txtLicenseType.Text = "None"
txtNetworkedLicense.Text = ""
txtMaxCount.Text = ""

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ActiveLock Initialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    Dim autoRegisterKey As String
    Dim boolAutoRegisterKeyPath As Boolean
    Dim Msg As String
    Dim A() As String
    
    On Error GoTo NotRegistered
    Me.Caption = "ALTestApp - ActiveLock3 Test Application - v" & App.Major & "." & App.Minor & "." & App.Revision
    
    ' Check the existence of necessary files to run this application
    Call CheckForResources("Alcrypto3.dll", "comctl32.ocx", "tabctl32.ocx")

    ' Check if the Activelock3.dll is registered. If not no need to continue.
    If CheckIfDLLIsRegistered = False Then End
    
    On Error GoTo NotRegistered
    ' Obtain AL instance and initialize its properties
    Set MyActiveLock = ActiveLock3.NewInstance()
    With MyActiveLock
        
        
        .SoftwareName = LICENSE_ROOT
        txtName.Text = .SoftwareName
        
        ' Note: Do not use (App.Major & "." & App.Minor & "." & App.Revision)
        ' since the license will fail with version incremented exe builds
        '.SoftwareVersion = "1.0"   ' WARNING *** WARNING *** DO NOT USE App.Major & "." & App.Minor & "." & App.Revision
        .SoftwareVersion = "2.0"
        txtVersion.Text = .SoftwareVersion
        
        ' New in v3.3
        ' This should be set to protect yourself against ResetTrial abuse
        ' The password is also used by the Short License Key type below
        .SoftwarePassword = Chr(99) & Chr(111) & Chr(111) & Chr(108)
        
        ' New in v3.5
        '.LicenseKeyType = alsShortKeyMD5
        .LicenseKeyType = alsALCryptoRSA
        
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
        .TrialHideType = trialHiddenFolder Or trialRegistry Or trialSteganography
        
        '.SoftwareCode = Dec(PUB_KEY)
        .SoftwareCode = "RSA2048BgIAAAAkAABSU0ExAAgAAAEAAQDPXvRi8OsaXj7Zlf0kHDL9ZZiTFUK02rx0kluIooFQV3hF44+hmnb3O+vLq1DJd5NK9I/tQ9uMTfGGLhjKCXVWYT8SkU3+em0Er2hmZX4EW63kVxFQMFAAYjhN7IFxIMqZJ/2DG+awKlrszorZ8U+fwHKPMomiZDxAzFWmD78H4HC/hxPaG8Tso2AkoScjlYHN8zLh/jeQ6IKaFpIBqLwfLVQF9tCUnsZ8Po6Erk44xqNsH3n0yapqsjdl1jUONT8BOjHtD++7U3lEcyW+n2G4ZCMKCTzDeq7Qddqnz6ubnTUGJvPpXg3DTvfklbpNlcOftfxnr5M+k9MhFzTj2Yvu"
        .LockType = lockNone  'lockWindows 'Or lockComp 'Or lockComp Or lockWindows
        strAutoRegisterKeyPath = App.Path & "\" & LICENSE_ROOT & ".all"
        .AutoRegisterKeyPath = strAutoRegisterKeyPath
        If FileExist(strAutoRegisterKeyPath) Then boolAutoRegisterKeyPath = True
    End With

    ' Verify AL's authenticity
    txtChecksum = modMain.VerifyActiveLockdll()

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
    txtVersion.Text = MyActiveLock.SoftwareVersion
    cboSpeed = cboSpeed.List(2)

    ' Check registration status
    Dim strMsg As String
    MyActiveLock.Acquire strMsg
    
    If strMsg <> "" Then 'There's a trial
        A = Split(strMsg, vbCrLf)
        txtRegStatus.Text = A(0)
        txtUsedDays.Text = A(1)
        FunctionalitiesEnabled = True
        frmSplash.lblInfo.Caption = vbCrLf & strMsg
        frmSplash.Show
        frmSplash.Refresh
        'Dim rtn As Long 'declare the need variables
        'rtn = SetWindowWord(frmSplash.hWnd, GWW_HWNDPARENT, Me.hWnd) 'let both forms load together
        Sleep 3000 'wait about 3 seconds
        Unload frmSplash
        cmdKillTrial.Visible = True
        cmdResetTrial.Visible = True
        txtLicenseType.Text = "Free Trial"
        Exit Sub
    Else
        cmdKillTrial.Visible = False
        cmdResetTrial.Visible = False
    End If
    
    ' Uncomment the following to retrieve the usedlocktypes
'    Dim aa() As ActiveLock3.ALLockTypes
'    ReDim aa(UBound(MyActiveLock.UsedLockType))
'    aa = MyActiveLock.UsedLockType
'    MsgBox aa(0) 'For example, if only lockHDfirmware was used, this will return 256
    
    txtRegStatus.Text = "Registered"
    txtUsedDays.Text = MyActiveLock.UsedDays
    txtExpiration.Text = MyActiveLock.ExpirationDate
    If txtExpiration.Text = "" Then txtExpiration.Text = "Permanent" 'App has a permanent license
    txtUser.Text = MyActiveLock.RegisteredUser
    txtRegisteredLevel.Text = MyActiveLock.RegisteredLevel
    
    ' Networked Licenses
    If MyActiveLock.LicenseClass = "MultiUser" Then
        txtNetworkedLicense.Text = "Networked"
    Else
        txtNetworkedLicense.Text = "Single User"
        txtMaxCount.Visible = False
        lblConcurrentUsers.Visible = False
    End If
    txtMaxCount.Text = MyActiveLock.MaxCount

    'Read the license file into a string to determine the license type
    Dim strBuff As String
    Dim fNum As Integer
    fNum = FreeFile
    Open strKeyStorePath For Input As #fNum
    strBuff = Input(LOF(1), 1)
    Close #fNum
    If Instring(strBuff, "LicenseType=3") Then
        txtLicenseType.Text = "Time Limited"
    ElseIf Instring(strBuff, "LicenseType=1") Then
        txtLicenseType.Text = "Periodic"
    ElseIf Instring(strBuff, "LicenseType=2") Then
        txtLicenseType.Text = "Permanent"
    End If
    FunctionalitiesEnabled = True
    
    Exit Sub

NotRegistered:
    FunctionalitiesEnabled = False
    If Instring(Err.Description, "no valid license") = False And noTrialThisTime = False Then
        MsgBox Err.Number & ": " & Err.Description
    End If
    txtRegStatus.Text = Err.Description
    txtLicenseType.Text = "None"
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation
    End If
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Key Request and Registration Functionalities
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cmdReqGen_Click()
    Dim instCode As String
    ' Generate Request code to Lock
    If MyActiveLock Is Nothing Then
        noTrialThisTime = True
        Form_Load
    End If
    
    If txtRegStatus.Text <> "Registered" Then txtRegStatus.Text = ""
    If Not IsNumeric(txtUsedDays.Text) Then txtUsedDays.Text = ""
    instCode = MyActiveLock.InstallationCode(txtUser.Text)
    If Len(instCode) = 8 Then
        instCode = "You must send all of the following for authorization:" & vbCrLf & _
            "Serial Number: " & instCode & vbCrLf & _
            "Application Name: " & txtName.Text & " - Version " & txtVersion.Text & vbCrLf & _
            "User Name: " & txtUser.Text
    End If
    txtReqCodeGen.Text = instCode
End Sub

Private Sub cmdRegister_Click()
    Dim ok As Boolean, LibKey As String
    On Error GoTo errHandler
    ' Register this key
    LibKey = txtLibKeyIn.Text
    If Mid(LibKey, 5, 1) = "-" And Mid(LibKey, 10, 1) = "-" And Mid(LibKey, 15, 1) = "-" And Mid(LibKey, 20, 1) = "-" Then
        MyActiveLock.Register LibKey, txtUser.Text 'YOU MUST SPECIFY THE USER NAME WITH SHORT KEYS !!!
    Else    ' ALCRYPTO RSA
        MyActiveLock.Register LibKey
    End If
    MsgBox modMain.Dec("386.457.46D.483.4F1.4FC.4E6.42B.4FC.483.4C5.4BA.160.4F1.507.441.441.457.4F1.4F1.462.507.4A4.16B"), vbInformation ' "Registration successful!"
    Unload Me
    Form_Load
    Me.Visible = True
    Exit Sub

errHandler:
    MsgBox Err.Number & ": " & Err.Description
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' The rest of the application's functionalities
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub cboSpeed_Click()

End Sub

Private Property Let FunctionalitiesEnabled(flag As Boolean)
    chkScroll.Enabled = flag
    chkFlash.Enabled = flag
    lblHost.Enabled = flag
    optForm(0).Enabled = flag
    optForm(1).Enabled = flag
    chkPause.Enabled = flag
    lblSpeed.Enabled = flag
    cboSpeed.Enabled = flag
    If flag Then
        If txtRegisteredLevel.Text <> "" Then
            lblLockStatus2.Caption = "Enabled with " & txtRegisteredLevel.Text
            chkPause.Enabled = (InStr(1, txtRegisteredLevel.Text, "Level 3") > 0)
            lblSpeed.Enabled = (InStr(1, txtRegisteredLevel.Text, "Level 4") > 0)
            cboSpeed.Enabled = (InStr(1, txtRegisteredLevel.Text, "Level 4") > 0)
        Else
            lblLockStatus2.Caption = "Enabled with " & txtUsedDays.Text
        End If
    Else
        lblLockStatus2.Caption = "Disabled (Registration Required)"
        chkPause.Enabled = False
        lblSpeed.Enabled = False
        cboSpeed.Enabled = False
    End If
End Property

Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
    'DO NOT ADD THE "END" STATEMENT INTO THIS SUB
    'Form reloads upon registration
End Sub

Private Sub txtLibKeyIn_Change()
    cmdRegister.Enabled = CBool(Trim$(txtLibKeyIn.Text) <> "")
End Sub

Private Sub txtName_Change()
    'MyActiveLock.SoftwareName = txtName
End Sub

Private Sub txtVersion_Change()
    'MyActiveLock.SoftwareVersion = txtVersion
End Sub

Public Sub UpdateStatus(Txt As String)
    sbStatus.SimpleText = Txt
End Sub

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
