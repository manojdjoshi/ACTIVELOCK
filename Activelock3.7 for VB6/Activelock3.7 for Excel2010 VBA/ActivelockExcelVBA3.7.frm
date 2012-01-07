VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmActivelockRegistration 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALTestApp - ActiveLock3 Test Application"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9525
   Icon            =   "ActivelockExcelVBA3.7.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9525
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton but_cont 
      Caption         =   "&Continue"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6450
      TabIndex        =   46
      Top             =   6660
      Width           =   3015
   End
   Begin MSComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   45
      Top             =   7245
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   16748
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraReg 
      Caption         =   "Register"
      ForeColor       =   &H00FF0000&
      Height          =   4065
      Left            =   0
      TabIndex        =   20
      Top             =   2490
      Width           =   9495
      Begin VB.CommandButton cmdPaste 
         Height          =   345
         Left            =   8220
         Picture         =   "ActivelockExcelVBA3.7.frx":0CCA
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1680
         Width           =   345
      End
      Begin VB.CommandButton cmdCopy 
         Height          =   345
         Left            =   8220
         MaskColor       =   &H8000000F&
         Picture         =   "ActivelockExcelVBA3.7.frx":100C
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   900
         Width           =   345
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
         TabIndex        =   26
         Top             =   1380
         Width           =   6675
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "&Register"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8220
         TabIndex        =   25
         ToolTipText     =   "Register the License"
         Top             =   2100
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
         Height          =   765
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   24
         Top             =   600
         Width           =   6675
      End
      Begin VB.CommandButton cmdReqGen 
         Caption         =   "&Generate"
         Height          =   315
         Left            =   8220
         TabIndex        =   23
         ToolTipText     =   "Generate Installation Code"
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1440
         TabIndex        =   22
         Text            =   "Evaluation User"
         Top             =   300
         Width           =   6675
      End
      Begin VB.CommandButton cmdKillLicense 
         Caption         =   "&Kill License"
         Height          =   315
         Left            =   8220
         TabIndex        =   21
         ToolTipText     =   "Kill the License"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Liberation Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Installation Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "User Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   1335
      End
   End
   Begin VB.Frame fraRegStatus 
      Caption         =   "Status"
      ForeColor       =   &H00FF0000&
      Height          =   2475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.Timer continueTimer 
         Interval        =   1000
         Left            =   7470
         Top             =   720
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   8370
         Picture         =   "ActivelockExcelVBA3.7.frx":131E
         ScaleHeight     =   825
         ScaleWidth      =   825
         TabIndex        =   30
         Top             =   210
         Width           =   825
      End
      Begin VB.TextBox txtRegStatus 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox txtUsedDays 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1140
         Width           =   4335
      End
      Begin VB.TextBox txtExpiration 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "ALExcelSample"
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox txtVersion 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "3.7"
         Top             =   540
         Width           =   4335
      End
      Begin VB.TextBox txtChecksum 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtRegisteredLevel 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1740
         Width           =   4335
      End
      Begin VB.CommandButton cmdKillTrial 
         Caption         =   "&Kill Trial"
         Height          =   315
         Left            =   8280
         TabIndex        =   3
         ToolTipText     =   "End the Free Trial"
         Top             =   1740
         Width           =   1095
      End
      Begin VB.CommandButton cmdResetTrial 
         Caption         =   "&Reset Trial"
         Height          =   315
         Left            =   8280
         TabIndex        =   2
         ToolTipText     =   "Reset the Free Trial"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtLicenseType 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   4140
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   2040
         Width           =   1755
      End
      Begin VB.Label Label6 
         Caption         =   "Registered:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   870
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Days Used:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Expiry Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "App Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "App Version:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "DLL Checksum:"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2070
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Registered Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1770
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Activelock v3.7"
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   8130
         TabIndex        =   12
         Top             =   1050
         Width           =   1305
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "License Type:"
         Height          =   255
         Left            =   2940
         TabIndex        =   11
         Top             =   2070
         Width           =   1155
      End
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   300
      Left            =   1980
      TabIndex        =   44
      Top             =   6630
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   529
      _Version        =   393216
      LargeChange     =   1
      Min             =   1
      SelStart        =   1
      Value           =   1
      TextPosition    =   1
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Display (sec)"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1470
      TabIndex        =   43
      Top             =   6630
      Width           =   525
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      Height          =   225
      Index           =   0
      Left            =   2025
      TabIndex        =   42
      Top             =   6975
      Width           =   180
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      Height          =   225
      Index           =   1
      Left            =   2460
      TabIndex        =   41
      Top             =   6975
      Width           =   180
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   225
      Index           =   2
      Left            =   2910
      TabIndex        =   40
      Top             =   6975
      Width           =   180
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   225
      Index           =   3
      Left            =   3330
      TabIndex        =   39
      Top             =   6975
      Width           =   180
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   225
      Index           =   4
      Left            =   3810
      TabIndex        =   38
      Top             =   6975
      Width           =   180
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      Height          =   225
      Index           =   5
      Left            =   4230
      TabIndex        =   37
      Top             =   6975
      Width           =   180
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      Height          =   225
      Index           =   6
      Left            =   4680
      TabIndex        =   36
      Top             =   6975
      Width           =   180
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   225
      Index           =   7
      Left            =   5130
      TabIndex        =   35
      Top             =   6975
      Width           =   180
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   225
      Index           =   8
      Left            =   5550
      TabIndex        =   34
      Top             =   6975
      Width           =   180
   End
   Begin VB.Label lblNumber 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   225
      Index           =   9
      Left            =   6000
      TabIndex        =   33
      Top             =   6975
      Width           =   195
   End
End
Attribute VB_Name = "frmActivelockRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2012 The ActiveLock Software Group (admin: Ismail Alkan)
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

'Application name used
Const LICENSE_ROOT As String = "ALExcelSample"
Dim secondsCount As Integer

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

Private Sub but_cont_Click()
but_cont.Caption = "&Close"
Unload Me
End Sub
Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtReqCodeGen.Text
End Sub

Private Sub cmdKillLicense_Click()
    Dim licFile As String
    licFile = App.path & "\" & LICENSE_ROOT & ".lic"
    If FileExist(licFile) Then
        If FileLen(licFile) <> 0 Then
            Kill licFile
            MsgBox "Your license has been killed." & vbCrLf & _
                "You need to get a new license for this application if you want to use it.", vbInformation
                txtUsedDays.Text = ""
                txtExpiration.Text = ""
                txtRegisteredLevel.Text = ""
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
End Sub

Private Sub continueTimer_Timer()
If thereIsAProblem = True Then Exit Sub
If continueTimer.Interval = 0 Then Exit Sub
'If but_cont.Caption = "&Exit" Then Exit Sub
secondsCount = MaxArg(0, secondsCount - 1)
but_cont.Caption = "Closing in " & CStr(secondsCount) & " sec"
If secondsCount = 0 Then
    but_cont_Click
End If
End Sub

Private Function MaxArg(ByVal x As Variant, ParamArray MyArray()) As Variant  'returns largest one of a variable number of arguements
Dim y As Variant    'member of array that holds all arguments except the first
For Each y In MyArray
If y > x Then x = y
Next y
MaxArg = x  'set return value
End Function

Private Sub Form_Activate()
If secondsCount = 0 Then
    continueTimer.Enabled = False
    Me.Visible = False
End If
End Sub

Private Sub Form_Click()
'    continueTimer.Enabled = False
'    If instring(but_cont.Caption, "continue") = True Then
'        but_cont.Caption = "&Continue"
'    End If
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ActiveLock Initialization
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Form_Load()
    Dim autoRegisterKey As String
    Dim strKeyStorePath As String, strAutoRegisterKeyPath As String
    Dim boolAutoRegisterKeyPath As Boolean
    Dim Msg As String
    Dim A() As String
    
    continueTimer.Interval = 1000
    secondsCount = Val(ProfileString32(LICENSE_ROOT & ".ini", "Startup Options", "Startup Seconds", "10"))
    Slider1.Min = 1
    Slider1.Max = 10
    Slider1.Value = secondsCount
    Label5.Visible = False
    Slider1.Visible = True
    
    On Error GoTo NotRegistered
    Me.Caption = "Excel 2010 - ActiveLock3 VBA Sample for v3.7"
    
    ' Check the existence of necessary files to run this application
    Call CheckForResources("comctl32.ocx", "tabctl32.ocx")
    If thereIsAProblem = True Then Exit Sub
    
    ' Check if the Activelock3.dll is registered. If not no need to continue.
    If CheckIfDLLIsRegistered = False Then
        thereIsAProblem = True
        Exit Sub
    End If
    
    On Error GoTo NotRegistered
    ' Obtain AL instance and initialize its properties
    Set MyActiveLock = ActiveLock3.NewInstance()
    With MyActiveLock
        
        .SoftwareName = LICENSE_ROOT
        txtName.Text = .SoftwareName
        
        ' Note: Do not use (App.Major & "." & App.Minor & "." & App.Revision)
        ' since the license will fail with version incremented exe builds
        .SoftwareVersion = "3.7"   ' WARNING *** WARNING *** DO NOT USE App.Major & "." & App.Minor & "." & App.Revision
        txtVersion.Text = .SoftwareVersion
        
        ' New in v3.3
        ' This should be set to protect yourself against ResetTrial abuse
        .SoftwarePassword = Chr(99) & Chr(111) & Chr(111) & Chr(108)
        
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
        
        .SoftwareCode = PUB_KEY
        '.LockType = lockHDFirmware Or lockComp 'Or lockComp Or lockWindows 'Change this to lockNone if just want to lock to user name
        strAutoRegisterKeyPath = App.path & "\" & LICENSE_ROOT & ".all"
        .AutoRegisterKeyPath = strAutoRegisterKeyPath
        If FileExist(strAutoRegisterKeyPath) Then boolAutoRegisterKeyPath = True
    End With

    ' Verify AL's authenticity
    txtChecksum = modMain.VerifyActiveLockdll()
    If thereIsAProblem = True Then Exit Sub
    
    ' Initialize the keystore. We use a File keystore in this case.
    MyActiveLock.KeyStoreType = alsFile
    
    ' Path to the license file
    strKeyStorePath = App.path & "\" & LICENSE_ROOT & ".lic"
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

    ' Check registration status
    Dim strMsg As String
    MyActiveLock.Acquire strMsg
    If strMsg <> "" Then 'There's a trial
        A = Split(strMsg, vbCrLf)
        txtRegStatus.Text = A(0)
        txtUsedDays.Text = A(1)
        thereIsAProblem = False
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
    
    txtRegStatus.Text = "Registered"
    txtUsedDays.Text = MyActiveLock.UsedDays
    txtExpiration.Text = MyActiveLock.ExpirationDate
    If txtExpiration.Text = "" Then txtExpiration.Text = "Permanent" 'App has a permanent license
    txtUser.Text = MyActiveLock.RegisteredUser
    txtRegisteredLevel.Text = MyActiveLock.RegisteredLevel
    'Read the license file into a string to determine the license type
    Dim strBuff As String
    Dim fNum As Integer
    fNum = FreeFile
    Open strKeyStorePath For Input As #fNum
    strBuff = Input(LOF(1), 1)
    Close #fNum
    If instring(strBuff, "LicenseType=3") Then
        txtLicenseType.Text = "Time Limited"
    ElseIf instring(strBuff, "LicenseType=1") Then
        txtLicenseType.Text = "Periodic"
    ElseIf instring(strBuff, "LicenseType=2") Then
        txtLicenseType.Text = "Permanent"
    End If
    thereIsAProblem = False
    Exit Sub

NotRegistered:
    thereIsAProblem = True
    If instring(Err.Description, "no valid license") = False And noTrialThisTime = False Then
        MsgBox Err.Number & ": " & Err.Description
    End If
    txtRegStatus.Text = Err.Description
    txtLicenseType.Text = "None"
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation
    End If
    but_cont.Caption = "&Close"
    continueTimer.Interval = 0
    Exit Sub

DLLnotRegistered:
    thereIsAProblem = True
    Exit Sub

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
        pathName = App.path
        s = Mid$(s, 2)
    ElseIf instring(s, "\") Then
        j = InStrRev(s, "\")
        pathName = Left$(s, j - 1)
        s = Mid$(s, j + 1)
    Else
        pathName = systemDir
    End If
    
    If instring(s, ".") Then
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
        thereIsAProblem = True
        Exit Function
    End If
Next y

CheckForResources = True
Exit Function

checkForResourcesError:
    MsgBox "CheckForResources error", vbCritical, "Error"
    thereIsAProblem = True
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
    s = App.path & ";" & CurDir & ";"
    
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
If Not instring(Right$(T, 4), ".") Then T = T & ".DLL"   'default extension
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
Function instring(ByVal x As String, ParamArray MyArray()) As Boolean
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
    ' Generate Request code to Lock
    If MyActiveLock Is Nothing Then
        noTrialThisTime = True
        Form_Load
    End If
    If txtRegStatus.Text <> "Registered" Then txtRegStatus.Text = ""
    If Not IsNumeric(txtUsedDays.Text) Then txtUsedDays.Text = ""
    txtReqCodeGen.Text = MyActiveLock.InstallationCode(txtUser)
End Sub

Private Sub cmdRegister_Click()
    Dim ok As Boolean
    Dim strKeyStorePath As String
    On Error GoTo errHandler
    ' Register this key
    MyActiveLock.Register txtLibKeyIn
    MsgBox modMain.Dec("386.457.46D.483.4F1.4FC.4E6.42B.4FC.483.4C5.4BA.160.4F1.507.441.441.457.4F1.4F1.462.507.4A4.16B"), vbInformation ' "Registration successful!"
    txtRegStatus.Text = "Registered"
    txtUsedDays.Text = MyActiveLock.UsedDays
    txtExpiration.Text = MyActiveLock.ExpirationDate
    If txtExpiration.Text = "" Then txtExpiration.Text = "Permanent" 'App has a permanent license
    txtUser.Text = MyActiveLock.RegisteredUser
    txtRegisteredLevel.Text = MyActiveLock.RegisteredLevel
    'Read the license file into a string to determine the license type
    Dim strBuff As String
    Dim fNum As Integer
    fNum = FreeFile
    strKeyStorePath = App.path & "\" & LICENSE_ROOT & ".lic"
    Open strKeyStorePath For Input As #fNum
    strBuff = Input(LOF(1), 1)
    Close #fNum
    If instring(strBuff, "LicenseType=3") Then
        txtLicenseType.Text = "Time Limited"
    ElseIf instring(strBuff, "LicenseType=1") Then
        txtLicenseType.Text = "Periodic"
    ElseIf instring(strBuff, "LicenseType=2") Then
        txtLicenseType.Text = "Permanent"
    End If
    thereIsAProblem = False
    Me.Visible = True
    continueTimer.Interval = 1000
    Exit Sub

errHandler:
    MsgBox Err.Number & ": " & Err.Description
End Sub
Private Sub Form_Unload(Cancel As Integer)
Me.Visible = False
secondsCount = 0
'but_cont.Caption = "&Exit"
continueTimer.Interval = 0
continueTimer.Enabled = False
Set frmActivelockRegistration = Nothing
'DO NOT ADD THE "END" STATEMENT INTO THIS SUB
'Form reloads upon registration
End Sub

Private Sub Slider1_Click()
    secondsCount = Slider1.Value
    Dim m As Long
    m = SetProfileString32(LICENSE_ROOT & ".ini", "Startup Options", "Startup Seconds", CStr(secondsCount))
End Sub

Private Sub Slider1_Scroll()
    secondsCount = Slider1.Value
    Dim m As Long
    m = SetProfileString32(LICENSE_ROOT & ".ini", "Startup Options", "Startup Seconds", CStr(secondsCount))
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
