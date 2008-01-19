VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9435
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Sample App Controls"
      Height          =   975
      Left            =   3840
      TabIndex        =   34
      Top             =   4440
      Width           =   5535
      Begin VB.CommandButton Command3 
         Caption         =   "Active For USA ONLY"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3600
         TabIndex        =   37
         Top             =   240
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Active For Any Registered"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1800
         TabIndex        =   36
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Active For ALL"
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "For Debug"
      Height          =   1695
      Left            =   120
      TabIndex        =   20
      Top             =   3720
      Width           =   3615
      Begin VB.CommandButton cmdKillLic 
         Caption         =   "Kill The Licence"
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   1200
         Width           =   3135
      End
      Begin VB.CommandButton cmdKillTrial 
         Caption         =   "Kill The Trial"
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   720
         Width           =   3135
      End
      Begin VB.CommandButton cmdResetTrial 
         Caption         =   "Reset The Trial"
         Height          =   375
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   3135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Register"
      Height          =   4335
      Left            =   3840
      TabIndex        =   19
      Top             =   120
      Width           =   5535
      Begin VB.CommandButton cmdGenerateCode 
         Caption         =   "Go"
         Height          =   375
         Left            =   5040
         TabIndex        =   33
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Register"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   3720
         Width           =   5295
      End
      Begin VB.CommandButton cmdCopy 
         Height          =   345
         Left            =   5040
         MaskColor       =   &H8000000F&
         Picture         =   "Form1.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   720
         Width           =   345
      End
      Begin VB.CommandButton cmdPaste 
         Height          =   345
         Left            =   5040
         Picture         =   "Form1.frx":031E
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   1800
         Width           =   345
      End
      Begin VB.TextBox txtLibKeyIn 
         Height          =   1815
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   1800
         Width           =   3495
      End
      Begin VB.TextBox txtReqCodeGen 
         Height          =   855
         Left            =   1440
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtUsername 
         Height          =   285
         Left            =   1440
         TabIndex        =   25
         Text            =   "Evaluation User"
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label12 
         Caption         =   "Liberation Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Installation Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Username:"
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ActiveLock Info"
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.TextBox Text6 
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   2760
         Width           =   1815
      End
      Begin VB.TextBox Text9 
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "RegStatus"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "UsedDaysOrRuns"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label4 
         Caption         =   "ValidTrial"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "LicenceType"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "ExpirationDate"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label7 
         Caption         =   "RegisteredUser"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label8 
         Caption         =   "RegisteredLevel"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "LicenseClass"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label10 
         Caption         =   "MaxCount"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3120
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents ActiveLockEventSink As ActiveLockEventNotifier
Attribute ActiveLockEventSink.VB_VarHelpID = -1

Private Sub cmdCopy_Click()
If Len(txtReqCodeGen.Text) = 0 Then
txtReqCodeGen.Text = MyActiveLock.InstallationCode(txtUsername.Text)
End If
Clipboard.Clear
Clipboard.SetText txtReqCodeGen.Text
End Sub

Private Sub cmdGenerateCode_Click()
 txtReqCodeGen.Text = MyActiveLock.InstallationCode(txtUsername.Text)
End Sub

Private Sub cmdKillLic_Click()
KillTheLic
End Sub

Private Sub cmdKillTrial_Click()
KillTheTrial
End Sub

Private Sub cmdPaste_Click()
If Clipboard.GetText = txtReqCodeGen.Text Then
    MsgBox "You cannot paste the Installation Code into the Liberation Key field.", vbExclamation
    Exit Sub
End If
txtLibKeyIn.Text = Clipboard.GetText
End Sub

Private Sub cmdResetTrial_Click()
ResetTheTrial
End Sub



Private Sub Command4_Click()
Dim GoodRegister As Boolean
GoodRegister = RegisterTheApplication(txtLibKeyIn.Text, txtUsername.Text)
If GoodRegister = True Then
    Unload frmMain
    Form_Load
    Me.Visible = True
Else
MsgBox "Error Registering Your Application"
End If
End Sub

Public Sub Form_Load()
Dim CanRun As Boolean
CanRun = InitActivelock
If CanRun = True Then
'Valid Trial Or Valid License
Else
'Trial Or License Expired or Tampered
'Disable everything and ask to register
End If
      Text1.Text = ActivelockValues.RegStatus
      Text2.Text = ActivelockValues.UsedDaysOrRuns
      Text3.Text = ActivelockValues.ValidTrial
      Text4.Text = ActivelockValues.LicenceType
      Text5.Text = ActivelockValues.ExpirationDate
      Text6.Text = ActivelockValues.RegisteredUser
      Text7.Text = ActivelockValues.RegisteredLevel
      Text8.Text = ActivelockValues.LicenseClass
      Text9.Text = ActivelockValues.MaxCount
End Sub

Public Sub ActiveLockEventSink_ValidateValue(ByRef Value As String)
   ' Value = modALVB6.Encrypt(Value)
    Value = Encrypt(Value)
    
End Sub

