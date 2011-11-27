VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   """"""
   ClientHeight    =   2250
   ClientLeft      =   2625
   ClientTop       =   1845
   ClientWidth     =   3735
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   2250
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Activelock V3"
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   1350
      TabIndex        =   1
      Top             =   1980
      Width           =   1065
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   1440
      Picture         =   "Splash.frx":4CC3
      Top             =   1125
      Width           =   825
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2250
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub SSPanel2_Click()

End Sub


Private Sub Form_Load()
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'center the form on the screen
End Sub

