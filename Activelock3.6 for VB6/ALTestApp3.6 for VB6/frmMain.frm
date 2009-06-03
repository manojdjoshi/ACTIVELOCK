VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALTestApp - ActiveLock3 Test Application"
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9570
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   11520
   ScaleWidth      =   9570
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   41
      Top             =   11235
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16828
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraRegStatus 
      Caption         =   "Status"
      ForeColor       =   &H00FF0000&
      Height          =   2790
      Left            =   45
      TabIndex        =   24
      Top             =   0
      Width           =   9495
      Begin VB.TextBox txtMaxCount 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   5355
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2340
         Width           =   540
      End
      Begin VB.TextBox txtNetworkedLicense 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   2340
         Width           =   1890
      End
      Begin VB.TextBox txtLicenseType 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   4140
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   2040
         Width           =   1755
      End
      Begin VB.CommandButton cmdResetTrial 
         Caption         =   "&Reset Trial"
         Height          =   315
         Left            =   8280
         TabIndex        =   25
         ToolTipText     =   "Reset the Free Trial"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CommandButton cmdKillTrial 
         Caption         =   "&Kill Trial"
         Height          =   315
         Left            =   8280
         TabIndex        =   15
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
         Picture         =   "frmMain.frx":0CCA
         ScaleHeight     =   825
         ScaleWidth      =   825
         TabIndex        =   23
         Top             =   210
         Width           =   825
      End
      Begin VB.TextBox txtRegisteredLevel 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1740
         Width           =   4335
      End
      Begin VB.TextBox txtChecksum 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtVersion 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "1.0"
         Top             =   540
         Width           =   4335
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "TestApp"
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox txtExpiration 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox txtUsedDays 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1140
         Width           =   4335
      End
      Begin VB.TextBox txtRegStatus 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   4335
      End
      Begin VB.Label lblConcurrentUsers 
         Caption         =   "No. of Concurrent Users:"
         Height          =   255
         Left            =   3555
         TabIndex        =   39
         Top             =   2385
         Width           =   1785
      End
      Begin VB.Label Label10 
         Caption         =   "License Class:"
         Height          =   255
         Left            =   135
         TabIndex        =   38
         Top             =   2340
         Width           =   1335
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "License Type:"
         Height          =   255
         Left            =   2940
         TabIndex        =   37
         Top             =   2070
         Width           =   1155
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Activelock v3"
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   8250
         TabIndex        =   36
         Top             =   1050
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "Registered Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   1770
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "DLL Checksum:"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2070
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "App Version:"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "App Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Expiry Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Days Used:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "License Status:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   870
         Width           =   1380
      End
   End
   Begin VB.Frame fraReg 
      Caption         =   "Register"
      ForeColor       =   &H00FF0000&
      Height          =   5415
      Left            =   45
      TabIndex        =   3
      Top             =   2835
      Width           =   9495
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   360
         Top             =   3015
      End
      Begin VB.CommandButton cmdPaste 
         Height          =   345
         Left            =   8220
         Picture         =   "frmMain.frx":3C52
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2610
         Width           =   345
      End
      Begin VB.CommandButton cmdCopy 
         Height          =   345
         Left            =   8220
         MaskColor       =   &H8000000F&
         Picture         =   "frmMain.frx":3F94
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   900
         Width           =   345
      End
      Begin VB.CommandButton cmdKillLicense 
         Caption         =   "&Kill License"
         Height          =   315
         Left            =   8220
         TabIndex        =   9
         ToolTipText     =   "Kill the License"
         Top             =   3420
         Width           =   1095
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1440
         TabIndex        =   8
         Text            =   "Evaluation User"
         Top             =   300
         Width           =   6675
      End
      Begin VB.CommandButton cmdReqGen 
         Caption         =   "&Generate"
         Height          =   315
         Left            =   8220
         TabIndex        =   7
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
         Height          =   1620
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   600
         Width           =   6675
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "&Register"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8220
         TabIndex        =   5
         ToolTipText     =   "Register the License"
         Top             =   3015
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
         Height          =   3105
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2250
         Width           =   6675
      End
      Begin VB.Label Label13 
         Caption         =   "User Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Installation Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "Liberation Key:"
         Height          =   255
         Left            =   135
         TabIndex        =   12
         Top             =   2250
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   585
      TabIndex        =   0
      Top             =   9810
      Width           =   8265
      Begin VB.OptionButton optForm 
         Caption         =   "Option 2"
         Height          =   225
         Index           =   1
         Left            =   4380
         TabIndex        =   40
         Top             =   1125
         Width           =   1395
      End
      Begin VB.OptionButton optForm 
         Caption         =   "Option 1"
         Height          =   225
         Index           =   0
         Left            =   4380
         TabIndex        =   42
         Top             =   870
         Value           =   -1  'True
         Width           =   1395
      End
      Begin VB.ComboBox cboSpeed 
         Height          =   315
         ItemData        =   "frmMain.frx":42A6
         Left            =   6570
         List            =   "frmMain.frx":42B9
         Style           =   2  'Dropdown List
         TabIndex        =   43
         Top             =   45
         Width           =   1635
      End
      Begin VB.CheckBox chkPause 
         Caption         =   "Checkbox enabled for Level 3 only"
         Height          =   315
         Left            =   0
         TabIndex        =   44
         Top             =   315
         Width           =   2835
      End
      Begin VB.CheckBox chkFlash 
         Caption         =   "Checkbox enabled for ALL Levels"
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   630
         Width           =   2745
      End
      Begin VB.CheckBox chkScroll 
         Caption         =   "Checkbox enabled for ALL Levels"
         Height          =   225
         Left            =   0
         TabIndex        =   1
         Top             =   90
         Value           =   1  'Checked
         Width           =   2760
      End
      Begin VB.Label lblHost 
         Caption         =   "These Option Buttons will be enabled for ALL Levels:"
         Height          =   255
         Left            =   4380
         TabIndex        =   48
         Top             =   585
         Width           =   3765
      End
      Begin VB.Label lblSpeed 
         Caption         =   "This combobox will be enabled for Level 4 only"
         Height          =   255
         Left            =   3150
         TabIndex        =   45
         Top             =   90
         Width           =   3420
      End
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmMain.frx":42E3
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   45
      TabIndex        =   49
      Top             =   8325
      Width           =   9495
   End
   Begin VB.Label lblLockStatus 
      Caption         =   "Application Functionalities Are Currently: "
      Height          =   240
      Left            =   585
      TabIndex        =   46
      Top             =   8955
      Width           =   2895
   End
   Begin VB.Label lblLockStatus2 
      BackStyle       =   0  'Transparent
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
      Height          =   240
      Left            =   3600
      TabIndex        =   47
      Top             =   8955
      Width           =   5865
   End
   Begin VB.Label lblTrialInfo 
      Caption         =   "NOTE: All application functionalities are available in Trial Mode."
      Height          =   255
      Left            =   585
      TabIndex        =   50
      Top             =   9360
      Width           =   4545
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

Private MyActiveLock As ActiveLock3.IActiveLock ' Make sure this is form or project level declaration
Private WithEvents ActiveLockEventSink As ActiveLockEventNotifier
Attribute ActiveLockEventSink.VB_VarHelpID = -1

' Trial mode variables
Dim noTrialThisTime As Boolean 'ialkan - needed for registration while form was loaded via trial
Dim expireTrialLicense As Boolean
Dim strKeyStorePath As String
Dim strAutoRegisterKeyPath As String

'Application name used
Const LICENSE_ROOT As String = "ALVB6Sample"

' Timer count to check the license
Dim timerCount As Long

' The following declarations are used by the IsDLLAvailable function
' provided by the Activelock user Pinheiro
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const MAX_MESSAGE_LENGTH = 512

'Windows and System directory API
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'Splash Screen API
Private Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long
Const GWW_HWNDPARENT = (-8)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean

Private Declare Function GetSystemDefaultLCID Lib "kernel32" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32" _
   Alias "GetLocaleInfoA" _
  (ByVal Locale As Long, _
   ByVal LCType As Long, _
   ByVal lpLCData As String, _
   ByVal cchData As Long) As Long
Const LOCALE_SSHORTDATE As Long = &H1F 'short date format string

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

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtReqCodeGen.Text
End Sub

Private Sub cmdKillLicense_Click()
        
MsgBox "This feature is not intended to be sent to the end user." & vbCrLf & _
    "Because it removes all traces of a license from a machine." & vbCrLf & vbCrLf & _
    "The best way to terminate a license in an end-user's machine is to" & vbCrLf & _
    "just delete his LIC file and send him a new build of your app with" & vbCrLf & _
    "a new revision number (or upgrade in other terms).", vbInformation

' Kill the License File
' Let Activelock handle this
On Error GoTo errorTrap
MyActiveLock.KillLicense MyActiveLock.SoftwareName & MyActiveLock.SoftwareVersion, strKeyStorePath

' Use the following in an end-user's machine.
'Kill strKeyStorePath

MsgBox "Your license has been killed." & vbCrLf & _
    "You need to get a new license for this application if you want to use it.", vbInformation
    
errorTrap:
    txtUsedDays.Text = ""
    txtExpiration.Text = ""
    txtRegisteredLevel.Text = ""
    txtNetworkedLicense.Text = ""
    txtMaxCount.Text = ""
    
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
MsgBox "Free Trial has been Reset." & vbCrLf & "Please restart the application for a new Free Trial." & vbCrLf & vbCrLf & "Note: This feature is provided for the developers only to test their products;" & vbCrLf & "DO NOT provide this feature in your application.", vbInformation
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
    Dim boolAutoRegisterKeyPath As Boolean
    Dim Msg As String
    Dim A() As String
    
    On Error GoTo DLLnotRegistered
    
    ' Form's caption
    Me.Caption = LICENSE_ROOT & " - ActiveLock Test Application v" & App.Major & "." & App.Minor & " for VB6"
    
    ' Check the existence of necessary files to run this application
    ' This is not necessary if you're not using these controls in your app.
    Call CheckForResources("activelock3.6.dll", "mscomctl.ocx")

    ' Check if the Activelock DLL is registered. If not no need to continue.
    ' Since Actvelock DLL is a COM DLL, it must be registered via the
    ' regsvr32 command. Otherwise it's useless.
    ' The following code checks that the file exists and is loadable into memory
    ' You do not have to have this check but is recommended to make sure
    ' your DLL works
    If CheckIfDLLIsRegistered = False Then End
    
    On Error GoTo NotRegistered
    
    ' Set the path to the license file (LIC) and ALL file (if it exists)
    Dim ret As Long
    Dim AppfilePath As String
    AppfilePath = Space$(260)
    ret = SHGetSpecialFolderPath(0, AppfilePath, 46, False) ' 46 is for ...\All Users\Documents folder.
    If Trim(AppfilePath) <> Chr(0) Then
        AppfilePath = Left(AppfilePath, InStr(AppfilePath, Chr(0)) - 1)
    End If
    
    ' Obtain AL instance and initialize its properties
    ' Set a new instance of the Activelock object
    Set MyActiveLock = ActiveLock3.NewInstance()
    
    With MyActiveLock
        
        ' Set the software/application name
        .SoftwareName = LICENSE_ROOT
        txtName.Text = .SoftwareName
        
            ' Set the software/application version number
        ' Note: Do not use (App.Major & "." & App.Minor & "." & App.Revision)
        ' since the license will fail with version incremented exe builds
        ' THE FOLLOWING IS A SAMPLE USAGE
        '.SoftwareVersion = "1.3.2"   ' WARNING *** WARNING *** DO NOT USE App.Major & "." & App.Minor & "." & App.Revision
        .SoftwareVersion = "3.6"
        txtVersion.Text = .SoftwareVersion
        
        ' Set the software/application password
        ' This should be set to protect yourself against ResetTrial abuse
        ' The password is also used by the short keys
        ' Regular licensing does not use this password, but you should still use a password
        ' WARNING: You can not ignore this property. You *must* set a password.
        ' WARNING: Do not use characters outside the range of 32-126 or else you'll get an error
        ' WARNING: Do not use a password that's longer than 255 characters or else you'll get an error
        .SoftwarePassword = Chr(99) & Chr(111) & Chr(111) & Chr(108)
        
        ' Set whether the software/application will use a short key or RSA method
        ' alsRSA covers RSA native classes approach.
        ' RSA classes in .NET allows you to pick from several cipher strengths
        ' alsShortKeyMD5 is for short key protection only
        ' WARNING: Short key licenses use the lockFingerprint by default
        '.LicenseKeyType = alsShortKeyMD5
        .LicenseKeyType = alsRSA
        
        ' Set the Trial Feature properties
        ' If you don't want to use the trial feature in your app, set the TrialType
        ' property to trialNone.
        
        ' Set the trial type property
        ' this is either trialDays, or trialRuns or trialNone.
        .TrialType = trialDays
        
        ' Set the Trial Length property.
        ' This number represents the number of days or the number of runs (whichever is applicable).
        .TrialLength = 15
        If .TrialType <> trialNone And .TrialLength = 0 Then
            ' In such cases Activelock automatically generates errors -11001100 or -11001101
            ' to indicate that you're using the trial feature but, trial length was not specified
        End If
        
        ' Comment the following statement to use a certain trial data hiding technique
        ' Use OR to combine one or more trial hiding techniques
        ' or don't use this property to use ALL techniques
        ' WARNING: trialRegistryPerUser is "Per User"; this means each user trial feature
        ' is controlled that user's own registry hive.
        ' This means initiating a trial with one user does not initiate a trial for another user.
        ' trialHiddenFolder and trialSteganography are for "All Users"
        .TrialHideType = trialHiddenFolder Or trialRegistryPerUser Or trialSteganography
        
        ' Use the following if you'd like to make the trial warning message persistent
        ' until a license is registered.
        ' Expiration of the trial is not so clear to some users without any warnings.
        ' .trialWarningTemporary will show the trial expired warning only once.
        .TrialWarning = trialWarningPersistent
        
        ' Set the Software code
        ' This is the same thing as VCode
        ' Run Alugen first and create a VCode and GCode
        ' for the software name and version number you used above
        ' Then copy and use the VCode as the PUB_KEY here.
        ' It's up to you to encrypt it; just makes it more secure
        ' Enc encodes, Dec decodes the public key (VCode)
        ' Change Enc() and Dec(0 the way you want.
        ' The following is an example of a very long encoded PUB_KEY
'        PUB_KEY = "386.391.2CB.226.210.23C.268.2D6.46D.323.2CB.2CB.2CB.2CB.499.2CB.2CB.2D6.391.3A7.210.2F7.528.2CB.2CB.46D.2CB.2CB.2CB.2F7.2CB.2CB.37B.2D6"
'        PUB_KEY = PUB_KEY & ".21B.507.46D.370.53E.365.2F7.4A4.533.483.499.268.478.37B.483.3DE.457.4BA.533.44C.4E6.4DB.3DE.39C.339.35A.3DE.32E.4DB.3D3.252.478"
'        PUB_KEY = PUB_KEY & ".3C8.318.4D0.4DB.34F.205.2E1.4D0.39C.528.23C.1D9.457.4D0.205.210.2D6.457.3DE.4A4.3C8.4DB.499.441.21B.4E6.25D.51D.4DB.3DE.2F7.2D6.391"
'        PUB_KEY = PUB_KEY & ".30D.23C.51D.478.4AF.25D.507.4A4.436.3D3.51D.4D0.4E6.441.4F1.4D0.3D3.323.34F.42B.2F7.37B.483.302.2E1.44C.4AF.365.4FC.507.4BA.46D"
'        PUB_KEY = PUB_KEY & ".1D9.2CB.35A.365.1D9.42B.37B.478.210.4A4.4C5.2F7.3A7.3B2.2EC.2CB.273.339.247.4BA.370.339.512.483.46D.44C.483.386.441.436.226.25D.3BD"
'        PUB_KEY = PUB_KEY & ".1D9.2D6.483.499.37B.365.4DB.4F1.533.436.483.4DB.210.462.4D0.436.2F7.302.323.53E.34F.34F.4AF.4BA.268.4F1.436.4C5.4C5.344.2CB.499.247"
'        PUB_KEY = PUB_KEY & ".252.391.3A7.3BD.4E6.533.4E6.462.4A4.44C.268.48E.499.4DB.34F.499.3A7.39C.51D.231.3BD.4A4.32E.2CB.528.436.323.46D.2D6.391.4AF.2EC.441"
'        PUB_KEY = PUB_KEY & ".30D.2F7.386.318.365.39C.3BD.386.53E.3D3.3BD.39C.4F1.4DB.23C.483.51D.3A7.273.365.2F7.30D.4DB.4AF.252.273"
'        PUB_KEY = PUB_KEY & ".32E.39C.4D0.247.247.44C.2D6.386.3BD.268.365.528.25D.231.483.37B.3D3.4D0.528.318.344.30D.441.4BA.2D6.3B2.35A.4BA.2CB.3BD.478.252"
'        PUB_KEY = PUB_KEY & ".462.441.231.386.4BA.25D.3A7.44C.25D.2CB.247.30D.44C.46D.4F1.247.51D.370.512.273.1D9.4BA.4AF.391.499.32E.302.34F.323.4C5.30D.512.4D0"
'        PUB_KEY = PUB_KEY & ".1D9.507.323.39C.39C.268.53E.4E6.4C5.268.2F7.4C5.528.4AF.53E.370.436.4D0.3C8.4C5.35A.51D.2CB.3D3.2CB.252.4C5.231.3D3.37B.35A.4BA"
'        PUB_KEY = PUB_KEY & ".499.344.4C5.34F.231.25D.37B.247.39C.1D9.302.210.4E6.30D.4DB.4BA.512.35A.25D.441.339.3B2"
'        .SoftwareCode = Dec(PUB_KEY)
        ' The following practice is not recommended but is used to show how SoftwareCode is set
        ' It is recommended that you encode/decode the PUB_KEY and then set it to SoftwareCode
        .SoftwareCode = "RSA1024BgIAAAAkAABSU0ExAAQAAAEAAQCBcwKp9p1rkQhZyxTeREh9EM273wBqpODS+KLkeu/xn/Q0+w8uhBQfZq8f/sdRfL+S5LIBItOv0okG42mKcaNk0mRoSoJkUPMrRp43j8nAKVCmRrD7pZ1Do3uHM4SjydLY0omeK8vOCyZ2WldYy0IxwgQjNqMHLuG1rCg1DR4+5g=="
        
        ' Set the Hardware keys
        ' In order to pick the keys that you want to lock to in Alugen, use lockNone only
        ' Example: lockWindows Or lockComp
        ' You can combine any lockType(s) using OR as above
        ' If you want to lock to any keys explicitly, combine them using OR
        ' But you won't be able to uncheck/check any of them while in Alugen (too late at that point).
        ' WARNING: Short key licenses use the lockFingerprint by default
        ' WARNING: This has no effect for short key licenses
        .LockType = lockNone
        
        ' USAGE OF .ALL (AUTO-REGISTER) FILES
        ' THESE FILES CARRY THE LICENSE (LIBERATION) KEY FOR AUTO REGISTRATION
        ' Set the .ALL file path if you're using an ALL file.
        ' .ALL is an auto registration file.
        ' You generate .ALL files via Alugen and then send to the users
        ' They put the .ALL file in the directory you specify below
        ' .ALL simply contains the license key
        ' WARNING: .ALL files are deleted after they are used.
        ' It is recommended that you use both SoftwareName and Version in the .ALL filename
        ' since multiple .ALL files might exist in the same directory
        ' If you don't want to use the software name and version number explicitly, use an .ALL
        ' filename that is specific to this application
        ' Be aware that Alugen will not accept creation of ALL files without
        ' the software name and version in the ALL filename.
        ' You can rename the ALL file afterwards, but that name should match the ALL filename used here.
        
        ' .ALL FILE STORAGE OPTION-1
        ' Try to use the simplest accessible solution for both XP and Vista such as C:\TEMP
        If FolderExists("c:\temp") = False Then
            MkDir ("c:\temp")
        End If
        strAutoRegisterKeyPath = "c:\temp\" & .SoftwareName & .SoftwareVersion & ".all"
        
        ' .ALL FILE STORAGE OPTION-2
        ' The following (AppfilePath) is another option for both XP and Vista
        ' but you should be sure that the directory exists.
        ' Since it's also the folder where the LIC file resides, it's not recommended
        'If FolderExists(AppfilePath & "\" & .SoftwareName & .SoftwareVersion) = False Then
        '    MkDir (AppfilePath & "\" & .SoftwareName & .SoftwareVersion)
        'End If
        'strAutoRegisterKeyPath = AppfilePath & "\" & .SoftwareName & .SoftwareVersion & "\" & .SoftwareName & .SoftwareVersion & ".all"
        
        ' .ALL FILE STORAGE OPTION-3
        ' App.Path could be an option for XP, but not for Vista
        'If FolderExists(App.Path & "\" & .SoftwareName & .SoftwareVersion) = False Then
        '    MkDir (App.Path & "\" & .SoftwareName & .SoftwareVersion)
        'End If
        'strAutoRegisterKeyPath = App.Path & "\" & .SoftwareName & .SoftwareVersion & "\" & .SoftwareName & .SoftwareVersion & ".all"
        
        .AutoRegisterKeyPath = strAutoRegisterKeyPath
        If FileExist(strAutoRegisterKeyPath) Then boolAutoRegisterKeyPath = True
        
        ' Set if auto registration will be used.
        ' Auto registration uses the ALL file for license registration.
        .AutoRegister = alsEnableAutoRegistration
        
        ' Set the Time Server check for Clock Tampering
        ' This is optional but highly recommended.
        ' Although Activelock makes every effort to check if the system clock was tampered,
        ' checking a time server is the guaranteed way of knowing the correct UTC time/day.
        ' This feature might add some delay to your apps start-up time.
        .CheckTimeServerForClockTampering = alsDontCheckTimeServer      ' use alsCheckTimeServer to enforce time server check for clock tampering detection
        
        ' Set the system files clock tampering check
        ' This feature might add some delay to your apps start-up time.
        .CheckSystemFilesForClockTampering = alsDontCheckSystemFiles    ' use alsCheckSystemFiles to enforce system files scanning for clock tampering detection
        
        ' Set the license file format; this could be encrypted or plain
        ' Even in a plain file format, certain keys and dates are still encrypted.
        '.LicenseFileType = alsLicenseFileEncrypted
        .LicenseFileType = alsLicenseFilePlain
    
        ' Verify Activelock DLL's authenticity by checking its CRC
        ' This checkes the CRC of the Activelock DLL and compares it with the embedded value
        ' To change the embedded value; find the "CRC <> Value" check in this code,
        ' and change the Value() function to make it the same as the CRC.
        txtChecksum = modMain.VerifyActiveLockdll()
    
        ' Initialize the keystore. We use a File keystore in this case.
        ' The other type alsRegistry is NOT supported.
        .KeyStoreType = alsFile
        
        ' The code below will put the LIC file inside the "...\All Users\Documents" folder
        ' You can hard code this path and put the LIC file anywhere you want
        ' But be careful with limited user accounts in Vista.
        ' It's recommended that you put this file in shared and accessible folders in Vista
        ' It is recommended that you use both SoftwareName and Version in the LIC filename
        ' since multiple LIC files might exist in the same directory
        ' If you don't want to use the software name and version number explicitly, use an LIC
        ' filename that is specific to this application
        If FolderExists(AppfilePath & "\" & .SoftwareName & .SoftwareVersion) = False Then
            MkDir (AppfilePath & "\" & .SoftwareName & .SoftwareVersion)
        End If
        strKeyStorePath = AppfilePath & "\" & .SoftwareName & .SoftwareVersion & "\" & .SoftwareName & .SoftwareVersion & ".lic"
        
        ' App.Path could be an option for XP, but not so for Vista
        'If FolderExists(App.Path & "\" & .SoftwareName & .SoftwareVersion) = False Then
        '    MkDir (App.Path & "\" & .SoftwareName & .SoftwareVersion)
        'End If
        'strKeyStorePath = App.Path & "\" & .SoftwareName & .SoftwareVersion & "\" & .SoftwareName & .SoftwareVersion & ".lic"
        .KeyStorePath = strKeyStorePath
        
        ' Obtain the EventNotifier so that we can receive notifications from AL.
        ' Do Not Change This  - unless you know what this is for.
        Set ActiveLockEventSink = .EventNotifier
        
        ' Initialize Activelock
        ' Init() method below registeres the
        ' application if an ALL file was used.
        .Init strAutoRegisterKeyPath
        If FileExist(strKeyStorePath) And boolAutoRegisterKeyPath = True And strAutoRegisterKeyPath <> "" Then
            ' This means, an ALL file existed and was used to create a LIC file
            ' Init() method successfully registered the ALL file
            ' and returned the license key
            ' You can process that key here to see if there is any abuse, etc.
            ' ie. whether the key was used before, etc.
        End If
    
    End With

    ' Initialize other application settings
    txtVersion.Text = MyActiveLock.SoftwareVersion
    cboSpeed = cboSpeed.List(2)

    ' Check registration status
    ' Acquire() method does both trial and regular licensing
    ' If it generates an error, that means there NO trial, NO license
    ' If no error and returns a string, there's a trial but No license. Parse the string to display a trial message.
    ' If no error and no string returned, you've got a valid license.
    
    ' In case the Acquire method generates an error, so no license and no trial:
    ' If InStr(1, Err.Description, "No valid license") > 0 Or InStr(1, Err.Description, "license invalid") > 0 Then '-2147221503 & -2147221502

    MyActiveLock.Acquire strMsg, strRemainingTrialDays, strRemainingTrialRuns, strTrialLength, strUsedDays, strExpirationDate, strRegisteredUser, strRegisteredLevel, strLicenseClass, strMaxCount, strLicenseFileType, strLicenseType
    ' strMsg is to get the trial status
    ' All other parameters are Optional and you can actually get all of them
    ' using MyActivelock.Property usage, but keep in mind that
    ' doing so will check the license every time making this a time consuming
    ' way of reading those properties
    ' The fastest approach is to use the arguments from Acquire() method.
    If strMsg <> "" Then 'There's a trial
        ' Parse the returned string to display it on a form
        A = Split(strMsg, vbCrLf)
        txtRegStatus.Text = A(0)
        txtUsedDays.Text = A(1)
        
        ' You can also get the RemainingTrialDays or RemainingTrialRuns directly by:
        'txtRemainingTrialDays.Text = MyActiveLock.RemainingTrialDays Or MyActiveLock.RemainingTrialRuns
        ' At this point RemainingTrialDays and RemainingTrialRuns properties are directly accessible
        ' Even if you don't want to show a trial form at this point, you still know the
        ' trial status with one of these two properties (whichever is applicable).

        FunctionalitiesEnabled = True
        
        ' ALTERNATIVE 1
        ' Show a splash form for 3 seconds
        ' The form displays the trial days/run total allowed and remaining
        ' Splash form to display the trial period/run information.
        frmSplash.lblInfo.Caption = vbCrLf & strMsg
        frmSplash.Show
        frmSplash.Refresh
        Sleep 3000 'wait about 3 seconds
        Unload frmSplash
        
        ' ALTERNATIVE 2
        ' Show a splash form that shows two choices, register or try
        ' User must chosse one option for the form to close
        remainingDays = strRemainingTrialDays       'MyActiveLock.RemainingTrialDays
        totalDays = strTrialLength                  'MyActiveLock.TrialLength
        frmSplash1.Show vbModal
        
        cmdKillTrial.Visible = True
        cmdResetTrial.Visible = True
        txtLicenseType.Text = "Free Trial"
        'cmdRegister.Enabled = True
        Exit Sub
    Else
        cmdKillTrial.Visible = False
        cmdResetTrial.Visible = False
    End If
    
    ' You can retrieve the LockTypes set inside Alugen
    ' by accessing the UsedLockType property
    ' Uncomment the following to retrieve the usedlocktypes
'    Dim aa() As ActiveLock3.ALLockTypes
'    ReDim aa(UBound(MyActiveLock.UsedLockType))
'    aa = MyActiveLock.UsedLockType
'    MsgBox aa(0) 'For example, if only lockHDfirmware was used, this will return 256
    
    ' So far so good...
    ' If you are here already, that means you have a valid license.
    ' Set the textboxes in your app accordingly.
    txtRegStatus.Text = "Registered"
    txtUsedDays.Text = strUsedDays                          'MyActiveLock.UsedDays
    
    '* Dim LCID As Long, strSD As String
    'get the system's locale ID. Where the
    'user's locale ID is required, call
    'GetUserLocaleLCID instead.
    '* LCID = GetSystemDefaultLCID()
    '* strSD = GetUserLocaleInfo(LCID, LOCALE_SSHORTDATE)
    '* txtExpiration.Text = Format(strExpirationDate, strSD)                 'MyActiveLock.ExpirationDate
    txtExpiration.Text = strExpirationDate
    If txtExpiration.Text = "" Then txtExpiration.Text = "Permanent" 'App has a permanent license
    
    Dim arrProdVer() As String
    arrProdVer = Split(strRegisteredUser, "&&&") ' Extract the software name and version
    'arrProdVer = Split(MyActiveLock.RegisteredUser, "&&&") ' Extract the software name and version
    txtUser.Text = arrProdVer(0)
    txtRegisteredLevel.Text = strRegisteredLevel            'MyActiveLock.RegisteredLevel
    
    ' Set Networked Licenses if applicable
    'If MyActiveLock.LicenseClass = "MultiUser" Then
    If strLicenseClass = "MultiUser" Then
        txtNetworkedLicense.Text = "Networked"
    Else
        txtNetworkedLicense.Text = "Single User"
        txtMaxCount.Visible = False
        lblConcurrentUsers.Visible = False
    End If
    ' This is for number of concurrent users count in a netwrok license
    txtMaxCount.Text = strMaxCount                          'MyActiveLock.MaxCount

    ' Read the license file into a string to determine the license type
    ' You can read the LicenseType from the LIC file directly
    ' However, the LIC file should be in Plain format for this to work.
    'If MyActiveLock.LicenseFileType = alsLicenseFilePlain Then
    If strLicenseFileType = alsLicenseFilePlain Then
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
    Else
        If strLicenseType = "3" Then
            txtLicenseType.Text = "Time Limited"
        ElseIf strLicenseType = "1" Then
            txtLicenseType.Text = "Periodic"
        ElseIf strLicenseType = "2" Then
            txtLicenseType.Text = "Permanent"
        End If
    End If
    
    FunctionalitiesEnabled = True
    
    ' If your code has reached this point successfully, then you're good.
    ' If not, revisit your code, check and recheck,
    ' watch the video tutorial and if still in doubt,
    ' post a question in the forums
    ' http://www.activelocksoftware.com
    
    ' The following used with a timer control to
    ' check the license periodically
    ' Check the Timer1_Timer event for further details
    timerCount = 5000
    Timer1.Interval = 1
    Timer1.Enabled = True
    'cmdRegister.Enabled = False
   
    Exit Sub

NotRegistered:
    ' There's no valid trial or license - let the user know.
    ' and cripple your application if necessary
    ' or kill it if that's what you want.
    FunctionalitiesEnabled = False
    If Instring(Err.Description, "no valid license") = False And noTrialThisTime = False Then
        MsgBox Err.Number & ": " & Err.Description
    End If
    txtRegStatus.Text = Err.Description
    txtLicenseType.Text = "None"
    txtUser.Text = Current_User_Name
    'cmdRegister.Enabled = True
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation
    End If
    Exit Sub

DLLnotRegistered:
    End

End Sub
Private Function GetUserLocaleInfo(ByVal dwLocaleID As Long, _
                                  ByVal dwLCType As Long) As String

   Dim sReturn As String
   Dim r As Long

  'call the function passing the Locale type
  'variable to retrieve the required size of
  'the string buffer needed
   r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
    
  'if successful..
   If r Then
    
     'pad the buffer with spaces
      sReturn = Space$(r)
       
     'and call again passing the buffer
      r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))
     
     'if successful (r > 0)
      If r Then
      
        'r holds the size of the string
        'including the terminating null
         GetUserLocaleInfo = Left$(sReturn, r - 1)
      
      End If
   
   End If
    
End Function
Function Current_User_Name() As String

' Dimension variables
Dim lpBuff As String * 255
Dim ret As Long, UserName As String

' Get the user name minus any trailing spaces found in the name.
ret = GetUserName(lpBuff, 255)
Current_User_Name = Left(lpBuff, InStr(lpBuff, Chr(0)) - 1)

End Function

Public Function FolderExists(ByVal sFolder As String) As Boolean
Dim sResult As String

On Error Resume Next
sResult = Dir(sFolder, vbDirectory)

On Error GoTo 0
FolderExists = sResult <> ""
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
        pathName = WinSysDir
    End If
    
    If Instring(s, ".") Then
        If FileExist(pathName & "\" & s) Then foundIt = True
    ElseIf FileExist(pathName & "\" & s & ".DLL") Then
        foundIt = True
    ElseIf FileExist(pathName & "\" & s & ".OCX") Then
        foundIt = True
        s = s & ".OCX"
    End If
    
    ' Check the Activelock DLL and copy the app.path version to system directory
    ' if the app.path version is newer
    If s = "activelock" & CStr(App.Major) & "." & CStr(App.Minor) & ".dll" Then
        If FileExist(App.Path & "\" & s) = True Then
            If Get_Latest_File(WinSysDir & "\" & s, App.Path & "\" & s) = App.Path & "\" & s Then
                MsgBox "The DLL in the application folder: " & App.Path & "\" & s & vbCrLf & "is newer then the system32 version: " & WinSysDir & "\" & s & vbCrLf & vbCrLf & "Copy this newer DLL to your system32 folder.", vbInformation
                End
            End If
        End If
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
Function Get_Latest_File(ByVal sFile1 As String, ByVal sFile2 As String) As String
Dim DateFile1 As Date
Dim DateFile2 As Date
DateFile1 = FileDateTime(sFile1)
DateFile2 = FileDateTime(sFile2)
If DateDiff("s", DateFile1, DateFile2) = 0 Then
    Get_Latest_File = sFile1
ElseIf DateDiff("s", DateFile1, DateFile2) < 0 Then
    Get_Latest_File = sFile1
Else
    Get_Latest_File = sFile2
End If
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
    
    d = WinSysDir
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
    If txtUser.Text = "" Then
        MsgBox "User Name field is blank.", vbExclamation
        Exit Sub
    End If
    ' Generate Request code to Lock
    If MyActiveLock Is Nothing Then
        noTrialThisTime = True
        Form_Load
    End If
    
    If txtRegStatus.Text <> "Registered" Then txtRegStatus.Text = ""
    If Not IsNumeric(txtUsedDays.Text) Then txtUsedDays.Text = ""
    Screen.MousePointer = vbHourglass
    instCode = MyActiveLock.InstallationCode(txtUser.Text)
    If Len(instCode) = 8 Then
        instCode = "You must send all of the following for authorization:" & vbCrLf & _
            "Serial Number: " & instCode & vbCrLf & _
            "Application Name: " & txtName.Text & " - Version " & txtVersion.Text & vbCrLf & _
            "User Name: " & txtUser.Text
    End If
    txtReqCodeGen.Text = instCode
    Screen.MousePointer = vbDefault

End Sub

Private Sub cmdRegister_Click()
    Dim ok As Boolean, LibKey As String
    On Error GoTo errHandler
    ' Register this key
    LibKey = txtLibKeyIn.Text
    If Mid(LibKey, 5, 1) = "-" And Mid(LibKey, 10, 1) = "-" And Mid(LibKey, 15, 1) = "-" And Mid(LibKey, 20, 1) = "-" Then
        MyActiveLock.Register LibKey, txtUser.Text 'YOU MUST SPECIFY THE USER NAME WITH SHORT KEYS !!!
    Else    ' RSA
        MyActiveLock.Register LibKey
    End If
    MsgBox modMain.Dec("386.457.46D.483.4F1.4FC.4E6.42B.4FC.483.4C5.4BA.160.4F1.507.441.441.457.4F1.4F1.462.507.4A4.16B"), vbInformation ' "Registration successful!"
    Screen.MousePointer = vbHourglass
    Unload Me
    Set MyActiveLock = Nothing
    strMsg = ""
    remainingDays = 0
    totalDays = 0
    Form_Load
    Me.Visible = True
    Screen.MousePointer = vbDefault
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

Private Sub Timer1_Timer()
' You can use this timer control and this event in your
' own application to check the Activelock license periodically
' If there's anything wrong, this event will end your application
' In addition to the code below, make sure to declare timerCount
' and add the following in your Form_Load:
'    timerCount = 5000
'    Timer1.Interval = 1
'    Timer1.Enabled = True

On Error GoTo errorTrap
timerCount = timerCount - 1
If timerCount = 0 Then
    ' Check the license if there was a license at form_load
    ' and it was not a trial license
    If txtRegStatus.Text = Chr(82) & Chr(101) & Chr(103) & Chr(105) & Chr(115) & Chr(116) & Chr(101) & Chr(114) & Chr(101) & Chr(100) And strMsg = "" Then ' "Registered" and no trial message
        MyActiveLock.Acquire strMsg
        timerCount = 5000
    End If
End If
Exit Sub

errorTrap:
    MsgBox "License error. Quitting.", vbCritical
    End
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
    StatusBar1.SimpleText = Txt
End Sub

Private Function CheckIfDLLIsRegistered() As Boolean
Dim strDllPath As String
Dim Result As Double
On Error GoTo cannotRegister
    
CheckIfDLLIsRegistered = True

strDllPath = GetTypeLibPathFromObject()
Result = Shell("regsvr32.exe /s " & strDllPath)

If Result > 0 Then
    ' MsgBox "Activelock3.dll is Registered !"
    ' Just quietly proceed
Else
    MsgBox "Activelock" & CStr(App.Major) & "." & CStr(App.Minor) & ".dll is Not Registered!"
    CheckIfDLLIsRegistered = False
End If
cannotRegister:
End Function
