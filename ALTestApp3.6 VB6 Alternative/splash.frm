VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H80000014&
   BorderStyle     =   0  'None
   Caption         =   """"""
   ClientHeight    =   3870
   ClientLeft      =   2625
   ClientTop       =   1845
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3870
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton OK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5700
      TabIndex        =   3
      Top             =   180
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Your Product."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   585
      Index           =   1
      Left            =   2100
      TabIndex        =   10
      Top             =   420
      Width           =   2625
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   420
      Picture         =   "splash.frx":0000
      Top             =   420
      Width           =   825
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "forum.activelocksoftware.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   4290
      TabIndex        =   9
      Top             =   2280
      Width           =   2325
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "For Technical Support:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   8
      Top             =   2310
      Width           =   1995
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "www.activelocksoftware.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   4290
      TabIndex        =   7
      Top             =   1920
      Width           =   2205
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "For more information visit:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   6
      Top             =   1920
      Width           =   1995
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2007-2008 Active Lock SG. All Rights reserved."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2220
      TabIndex        =   5
      Top             =   1380
      Width           =   4635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Your sample product description...."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2220
      TabIndex        =   4
      Top             =   960
      Width           =   3645
   End
   Begin VB.Label LabelVer 
      BackStyle       =   0  'Transparent
      Caption         =   "version"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   4770
      TabIndex        =   2
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Activelock V3"
      ForeColor       =   &H00FF0000&
      Height          =   165
      Left            =   120
      TabIndex        =   1
      Top             =   2130
      Width           =   1065
   End
   Begin VB.Label lblInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Trial Message."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   60
      TabIndex        =   0
      Top             =   2970
      Width           =   6825
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000F&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   3885
      Left            =   0
      Top             =   0
      Width           =   7155
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2 'center the form on the screen
    LabelVer.Caption = App.Major & "." & App.Minor
    lblInfo.Caption = vbCrLf & objReg.GetExpiryMsg
End Sub

Private Sub OK_Click()
    Unload Me
End Sub
