VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FTest 
   Caption         =   "Test Registration"
   ClientHeight    =   2145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   ScaleHeight     =   2145
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtModulus 
      Height          =   330
      Left            =   5445
      TabIndex        =   18
      Text            =   "211"
      Top             =   360
      Width           =   780
   End
   Begin VB.TextBox txtEnigma 
      Height          =   330
      Left            =   3285
      TabIndex        =   13
      Text            =   "Encrypted Name"
      Top             =   720
      Width           =   2945
   End
   Begin VB.TextBox txtEncryptKey 
      Height          =   315
      Left            =   3285
      TabIndex        =   10
      Text            =   "FailSafe Systems"
      Top             =   360
      Width           =   2130
   End
   Begin MSComCtl2.DTPicker dtpExpire 
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "yyy MMMM dd"
      Format          =   20381699
      CurrentDate     =   36916.9999884259
   End
   Begin VB.TextBox txtFullKey 
      Height          =   315
      Left            =   3285
      TabIndex        =   8
      Top             =   1080
      Width           =   2940
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Text            =   "Text to Encrypt"
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "&Validate Key"
      Height          =   315
      Left            =   3285
      TabIndex        =   5
      Top             =   1440
      Width           =   2595
   End
   Begin VB.TextBox txtPart4 
      Height          =   315
      Left            =   2655
      TabIndex        =   4
      Top             =   1080
      Width           =   600
   End
   Begin VB.TextBox txtPart3 
      Height          =   315
      Left            =   1802
      TabIndex        =   3
      Top             =   1080
      Width           =   600
   End
   Begin VB.TextBox txtPart2 
      Height          =   315
      Left            =   951
      TabIndex        =   2
      Top             =   1080
      Width           =   600
   End
   Begin VB.TextBox txtPart1 
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   1095
      Width           =   700
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate Key"
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Width           =   2595
   End
   Begin VB.Label lblModulus 
      Caption         =   "Modulus"
      Height          =   315
      Left            =   5490
      TabIndex        =   17
      Top             =   45
      Width           =   735
   End
   Begin VB.Label lblH1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2475
      TabIndex        =   16
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label lblH1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1620
      TabIndex        =   15
      Top             =   1080
      Width           =   105
   End
   Begin VB.Label lblH1 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   765
      TabIndex        =   14
      Top             =   1080
      Width           =   100
   End
   Begin VB.Label lblEncryptKey 
      Caption         =   "Encryption Key"
      Height          =   315
      Left            =   3285
      TabIndex        =   12
      Top             =   45
      Width           =   2130
   End
   Begin VB.Label lblName 
      Caption         =   "Name of Licencee"
      Height          =   315
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Width           =   3120
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   6270
   End
End
Attribute VB_Name = "FTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const PRIME_GEN = 511
Private Sub cmdGenerate_Click()
    Dim sKey            As String
    If oCDKey.IsPrime(txtModulus.Text) Then
        If txtName.Text <> "" Then
            lblStatus.Caption = "Generating Key..."
            oCDKey.Modulus = txtModulus.Text
            sKey = oCDKey.MakeKeyCode(txtName.Text, FormatDateTime(dtpExpire.Value, vbShortDate), txtEncryptKey.Text)
            txtPart1.Text = Mid(sKey, 1, 5)
            txtPart2.Text = Mid(sKey, 6, 4)
            txtPart3.Text = Mid(sKey, 10, 4)
            txtPart4.Text = Mid(sKey, 14, 4)
            txtFullKey.Text = Mid(sKey, 1, 5) & "-" & Mid(sKey, 6, 4) & "-" & Mid(sKey, 10, 4) & "-" & Mid(sKey, 14, 4)
            oCrypto.Keyword = txtEncryptKey.Text
            txtEnigma.Text = oCrypto.Encrypt(txtName.Text)
            lblStatus.Caption = "Key Generated using Modulus " & oCDKey.Modulus
        Else
            lblStatus.Caption = "Supply a NAME first!"
        End If
    Else
        lblStatus.Caption = "Use a PRIME Number for the Modulus!"
    End If
End Sub

Private Sub cmdTest_Click()
    Dim sKey            As String
    Dim lStatus         As CDStatus
    If oCDKey.IsPrime(txtModulus.Text) Then
        lblStatus.Caption = "Generating Key..."
        oCDKey.Modulus = txtModulus.Text
        sKey = Trim(txtPart1.Text) & Trim(txtPart2.Text) & Trim(txtPart3.Text) & Trim(txtPart4.Text)
        If txtName.Text <> "" Then
            lblStatus.Caption = "Checking Validity of Key..."
            oCrypto.Keyword = txtEncryptKey.Text
            txtEnigma.Text = oCrypto.Decrypt(txtEnigma.Text)
            If oCrypto.ErrorMsg <> "" Then txtEnigma.Text = oCrypto.Encrypted
            lStatus = oCDKey.Status(txtName.Text, sKey, FormatDateTime(dtpExpire.Value, vbShortDate), txtEncryptKey.Text)
            MsgBox txtName.Text
            MsgBox sKey
            MsgBox FormatDateTime(dtpExpire.Value, vbShortDate)
            MsgBox txtEncryptKey.Text
            If lStatus = CDTampered Then
                lblStatus.Caption = "!!! Key has Tampered with !!!"
            ElseIf lStatus = CDExpired Then
                lblStatus.Caption = "!!! Key has Expired !!!"
            Else
                lblStatus.Caption = "Key is Valid (using Modulus " & oCDKey.Modulus & ")"
            End If
        Else
            lblStatus.Caption = "Supply a NAME first!"
        End If
    Else
        lblStatus.Caption = "Use a PRIME Number for the Modulus!"
    End If
End Sub

Private Sub DTPicker1_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)

End Sub

Private Sub Form_Load()
    Set oCDKey = New CDkey
    'txtName.Text = "FailSafe Systems"
    With dtpExpire
        .MinDate = "01 January 2001"
        .MaxDate = "13 October 2173"
        .Value = Now + 1
    End With
    'txtExpireDate.Text = FormatDateTime(CDate(99999), vbLongDate)
End Sub

