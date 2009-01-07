VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   5955
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Splash1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   510
      Left            =   5310
      TabIndex        =   8
      Top             =   5355
      Width           =   1860
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Height          =   510
      Left            =   2790
      TabIndex        =   7
      Top             =   5355
      Width           =   1860
   End
   Begin VB.CommandButton cmdTry 
      Caption         =   "Try"
      Height          =   510
      Left            =   180
      TabIndex        =   6
      Top             =   5355
      Width           =   1860
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   135
      TabIndex        =   5
      Top             =   4905
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   4725
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Label lblInfo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2610
         TabIndex        =   9
         Top             =   4410
         Width           =   4320
      End
      Begin VB.Image imgLogo 
         Height          =   4500
         Left            =   45
         Picture         =   "Splash1.frx":000C
         Top             =   135
         Width           =   2460
      End
      Begin VB.Label lblWarning 
         Caption         =   "Warning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   1
         Top             =   3660
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5985
         TabIndex        =   2
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2610
         TabIndex        =   4
         Top             =   1140
         Width           =   1365
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Activelock VB6 Sample"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2610
         TabIndex        =   3
         Top             =   705
         Width           =   3960
      End
   End
End
Attribute VB_Name = "frmSplash1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdRegister_Click()
    Unload Me
End Sub

Private Sub cmdTry_Click()
    Unload Me
End Sub


Private Sub Form_Load()
lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
lblProductName.Caption = App.Title
If strMsg <> "" Then
    ProgressBar1.Value = (1 - remainingDays / totalDays) * 100
    If remainingDays > 0 Then lblInfo.Caption = CStr(totalDays - remainingDays) & " days used out of " & CStr(totalDays) & " trial days"
End If
 
End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub
