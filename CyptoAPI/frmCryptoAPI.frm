VERSION 5.00
Begin VB.Form frmCryptoAPI 
   Caption         =   "CryptoAPI Demo - RSA Signing/Verification by Ismail"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9270
   LinkTopic       =   "FrmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   9270
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optStrength 
      Caption         =   "512-bit"
      Height          =   240
      Index           =   4
      Left            =   6885
      TabIndex        =   23
      Top             =   1170
      Width           =   1005
   End
   Begin VB.OptionButton optStrength 
      Caption         =   "1024-bit"
      Height          =   240
      Index           =   3
      Left            =   5805
      TabIndex        =   22
      Top             =   1170
      Value           =   -1  'True
      Width           =   1005
   End
   Begin VB.OptionButton optStrength 
      Caption         =   "1536-bit"
      Height          =   240
      Index           =   2
      Left            =   4725
      TabIndex        =   21
      Top             =   1170
      Width           =   1005
   End
   Begin VB.OptionButton optStrength 
      Caption         =   "2048-bit"
      Height          =   240
      Index           =   1
      Left            =   3645
      TabIndex        =   20
      Top             =   1170
      Width           =   1005
   End
   Begin VB.OptionButton optStrength 
      Caption         =   "4096-bit"
      Height          =   240
      Index           =   0
      Left            =   2565
      TabIndex        =   19
      Top             =   1170
      Width           =   1005
   End
   Begin VB.CommandButton cmdBtnClear 
      Caption         =   "Clear"
      Height          =   315
      Left            =   8220
      TabIndex        =   4
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Sign"
      Height          =   315
      Index           =   4
      Left            =   45
      TabIndex        =   5
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Generate Signature Key Pair"
      Height          =   315
      Index           =   1
      Left            =   45
      TabIndex        =   9
      Top             =   1125
      Width           =   2370
   End
   Begin VB.Frame Frame2 
      Height          =   2040
      Left            =   45
      TabIndex        =   11
      Top             =   1440
      Width           =   9195
      Begin VB.TextBox txtPublicKey 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "frmCryptoAPI.frx":0000
         Top             =   1485
         Width           =   9000
      End
      Begin VB.TextBox txtPrivateKey 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "frmCryptoAPI.frx":00CB
         Top             =   360
         Width           =   9000
      End
      Begin VB.Label Label3 
         Caption         =   "Private Key"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label Label2 
         Caption         =   "Public Key"
         Height          =   195
         Left            =   90
         TabIndex        =   14
         Top             =   1260
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Encrypt"
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   7155
      TabIndex        =   10
      Top             =   4320
      Width           =   975
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Validate"
      Height          =   315
      Index           =   5
      Left            =   1125
      TabIndex        =   8
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   3090
      Left            =   45
      TabIndex        =   6
      Top             =   4590
      Width           =   9195
      Begin VB.TextBox txtResult 
         Height          =   975
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   405
         Width           =   9000
      End
      Begin VB.TextBox txtSignedText 
         Height          =   1380
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmCryptoAPI.frx":03EC
         Top             =   1620
         Width           =   9000
      End
      Begin VB.Label Label5 
         Caption         =   "Plain Text"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   180
         Width           =   1050
      End
      Begin VB.Label Label4 
         Caption         =   "Signed Text"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   1395
         Width           =   1050
      End
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Destroy"
      Height          =   315
      Index           =   2
      Left            =   6030
      TabIndex        =   2
      Top             =   135
      Width           =   975
   End
   Begin VB.CommandButton cmdBtn 
      Caption         =   "Create"
      Height          =   315
      Index           =   0
      Left            =   4950
      TabIndex        =   1
      Top             =   135
      Width           =   975
   End
   Begin VB.TextBox txtContainer 
      Height          =   285
      Left            =   1920
      TabIndex        =   0
      Text            =   "ALVB6Sample3.6"
      Top             =   135
      Width           =   2910
   End
   Begin VB.Label Label1 
      Caption         =   "Key Container Name:"
      Height          =   255
      Left            =   270
      TabIndex        =   3
      Top             =   135
      Width           =   1575
   End
End
Attribute VB_Name = "frmCryptoAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cryptSession As clsCryptoAPI
Dim strCypherText As String
Dim bCypherOn As Boolean

Private Sub cmdBtnClear_Click()
' clear text box
bCypherOn = False
Me.txtResult = vbNullString
strCypherText = vbNullString
'Me.cmdBtn(3).Caption = "Encrypt"
'Me.frameCtl.Caption = "Plain Text"

End Sub

Private Sub Form_Load()

If App.PrevInstance = True Then
    MsgBox "Plurality is not allowed.", vbOKOnly, "I'm already running..."
    Unload Me
End If
Me.cmdBtn(0).Enabled = True
Set cryptSession = New clsCryptoAPI
If cryptSession.VerifyContainer(Me.txtContainer.Text, False) Then
    Me.cmdBtn(0).Enabled = False
    Me.cmdBtn(1).Enabled = True
    Me.cmdBtn(2).Enabled = True
End If

txtResult.Text = "ALVB6Sample" & vbLf & "3.6" & vbLf & "Single" & vbLf & "2" & vbLf & "test&&&ALVB6Sample - 3.6" _
    & vbLf & "Limited A" & vbLf & "2009/03/26" & vbLf & "5" & vbLf & "+nokey" & vbLf & _
    "nokey" & vbLf & "nokey" & vbLf & "nokey" & vbLf & "nokey" & vbLf & "nokey" & vbLf & _
    "nokey" & vbLf & "nokey" & vbLf & "nokey" & vbLf & "nokey" & vbLf & "nokey" & vbLf & _
    "nokey" & vbLf & "nokey" & vbLf & "nokey"

End Sub

Private Sub cmdBtn_Click(Index As Integer)
Dim sKeyDongle As String, lX As Long, modulus As Long

    Select Case Index
    Case 0               'create container
        Set cryptSession = New clsCryptoAPI
        If cryptSession.SessionStart(Me.txtContainer.Text, True) Then
            cryptSession.Generate False 'Signature Key
            cryptSession.ImportPrivateKey True, False
            cryptSession.ExportPrivateKey True, False
            cryptSession.ImportPublicKey False
            cryptSession.ExportPublicKey False
            Me.cmdBtn(0).Enabled = False
            Me.cmdBtn(1).Enabled = True
            Me.cmdBtn(2).Enabled = True
            'Me.cmdBtn(3).Enabled = False
            cryptSession.SessionEnd
        Else
            MsgBox "This System Unable to Perform Cryptographic Functions", vbOKOnly, "Cryptographic Process Failure..."
            Set cryptSession = Nothing
            Unload Me
            Exit Sub
        End If
        Set cryptSession = Nothing
    
    Case 1            'generate exchange or signature key
        Set cryptSession = New clsCryptoAPI
        If cryptSession.SessionStart(Me.txtContainer.Text) Then
            If optStrength(0).Value = True Then
                modulus = 4096
            ElseIf optStrength(1).Value = True Then
                modulus = 2048
            ElseIf optStrength(2).Value = True Then
                modulus = 1536
            ElseIf optStrength(3).Value = True Then
                modulus = 1024
            ElseIf optStrength(4).Value = True Then
                modulus = 512
            End If
            Screen.MousePointer = vbHourglass
            cryptSession.Generate False, modulus 'Signature Key
            Screen.MousePointer = vbDefault
            cryptSession.ExportPrivateKey True, False
            cryptSession.ExportPublicKey False
            txtPrivateKey.Text = publicPrivateKeyBlob
            txtPublicKey.Text = publicPublicKeyBlob
            'Me.cmdBtn(3).Enabled = False
            cryptSession.SessionEnd
        Else
            MsgBox "This System Unable to Perform Cryptographic Functions", vbOKOnly, "Cryptographic Process Failure..."
            Set cryptSession = Nothing
            Unload Me
            Exit Sub
        End If
        Set cryptSession = Nothing
    
    Case 2            'destroy container
        Set cryptSession = New clsCryptoAPI
        If cryptSession.SessionStart(Me.txtContainer.Text) Then
            cryptSession.DestroyContainer Me.txtContainer.Text
            cryptSession.SessionEnd
            Me.cmdBtn(0).Enabled = True
            Me.cmdBtn(1).Enabled = False
            Me.cmdBtn(2).Enabled = False
            'Me.cmdBtn(3).Enabled = False
            'Me.cmdBtn(3).Caption = "Encrypt"
            'Me.frameCtl.Caption = "Plain Text"
            bCypherOn = False
            Me.txtResult = vbNullString
            Me.txtSignedText.Text = vbNullString
            strCypherText = vbNullString
            'Me.txtResult.Enabled = False
        Else
            MsgBox "This System Unable to Perform Cryptographic Functions", vbOKOnly, "Cryptographic Process Failure..."
            Set cryptSession = Nothing
            Unload Me
            Exit Sub
        End If
        Set cryptSession = Nothing
    
    Case 3                'encrypt/decrypt
        Dim iChr As Integer
        Set cryptSession = New clsCryptoAPI
        If cryptSession.SessionStart(Me.txtContainer.Text) Then
            If Not bCypherOn Then
                bCypherOn = True
                strCypherText = cryptSession.EncryptString(txtResult.Text)
                Me.txtResult = vbNullString
                Me.cmdBtn(3).Caption = "Decrypt"
                'Me.frameCtl.Caption = "Cypher Text"
                Me.MousePointer = vbHourglass
                Me.Enabled = False
                sKeyDongle = strCypherText
                For lX = 0 To 31
                    sKeyDongle = Replace(sKeyDongle, Chr(lX), Chr(lX + 32))
                Next lX
                Me.txtResult = sKeyDongle
                Me.MousePointer = vbNormal
                Me.Enabled = True
            Else
                Me.txtResult = cryptSession.DecryptString(strCypherText)
                If Me.txtResult.Text <> vbNullString Then
                    Me.cmdBtn(3).Caption = "Encrypt"
                    'Me.frameCtl.Caption = "Plain Text"
                End If
                strCypherText = vbNullString
                bCypherOn = False
            End If
            cryptSession.SessionEnd
        Else
            MsgBox "Crypto Error", vbOKOnly, "OK to continue"
        End If
        Set cryptSession = Nothing
    
    Case 4              'sign the key
        If txtResult.Text = "" Then
            MsgBox "There's no string to sign !", vbInformation
            Exit Sub
        End If
'        Dim b64 As clsCryptoAPIBase64
'        Set b64 = New clsCryptoAPIBase64
        Set cryptSession = New clsCryptoAPI
        If cryptSession.VerifyContainer(Me.txtContainer.Text, False) = True Then
            cryptSession.DestroyContainer Me.txtContainer.Text
            cryptSession.SessionEnd
        End If
        If cryptSession.SessionStart(Me.txtContainer.Text, True) Then
            cryptSession.Generate False 'Signature Key
            cryptSession.ImportPrivateKey True, False
            'cryptSession.ExportPrivateKey True, False
            'cryptSession.ImportPublicKey False
            'cryptSession.ExportPublicKey False
            
            Dim inputText As String
            inputText = txtResult.Text
            Call cryptSession.SignString(inputText, True)
            txtSignedText.Text = inputText
        End If
    
    Case 5              'validate the key
        Set cryptSession = New clsCryptoAPI
        If cryptSession.VerifyContainer(Me.txtContainer.Text, False) = True Then
            cryptSession.DestroyContainer Me.txtContainer.Text
            cryptSession.SessionEnd
        End If
        If cryptSession.SessionStart(Me.txtContainer.Text, True) Then
            cryptSession.Generate False 'Signature Key
            cryptSession.ImportPublicKey False
            
            inputText = txtSignedText.Text
            If inputText = "" Then
                MsgBox "No string to validate !", vbInformation
                Exit Sub
            End If
            If cryptSession.ValidateString(inputText) = True Then
                MsgBox "Valid signature !!", vbInformation
            Else
                MsgBox "Invalid signature !!", vbInformation
            End If
        End If
    
    End Select
End Sub

Private Sub txtContainer_Change()
Dim sKeyDongle As String, sKeyDonglePrivate As String

    strCypherText = vbNullString
    Me.cmdBtn(0).Enabled = False        'create container
    Me.cmdBtn(1).Enabled = False        'generate exchange key
    Me.cmdBtn(2).Enabled = False        'destroy container
    'Me.cmdBtn(3).Enabled = False        'encrypt/decrypt text
    'Me.cmdBtn(3).Caption = "Encrypt"
    'Me.frameCtl.Caption = "Plain Text"
    'Me.txtResult.Enabled = False
    'Me.txtResult = vbNullString
    
    If Me.txtContainer.Text <> vbNullString Then
        Me.cmdBtn(0).Enabled = True
        Set cryptSession = New clsCryptoAPI
        If cryptSession.VerifyContainer(Me.txtContainer.Text, False) Then
            Me.cmdBtn(0).Enabled = False
            Me.cmdBtn(1).Enabled = True
            Me.cmdBtn(2).Enabled = True
            
            Dim publicKey As String, privateKey As String
            'privateKey = "UzMwODoHAgAAACQAAPh4QW25OhuOEKDIksffuUuJYQTsMeu/lKoemvJz8Oj56ShbVvou69rbAybVHXbhKnhrDnS9wjuIqOt9MWNmzld4G2tr33ll/8aQ9MaOZT/0G7hDmOVFPtl3L/3OtbGR9nj5oSoGlNWTrsgYNa9fV2M5kD4tVnheVqa55pf23K2iN40nufeAd2Oqnk7p8JHx58nITbCkQ1Q1W1zGm4z6jIktn0sPdgGuxmI5Rnx0onBtVp9XMIkwJgyzWvWUqrUKG4v2bu98Vmms3+3o9WOJDNYbZB1MeaGfBA+B6lUonRATh7jfefSh/8WCKxnu7JJPtp23rM64CpclIhjq2s+YAYhUPe9J/u+lx0hE61psS9rlAhBFu8nqOPVLiPuJSu+p3K5kjF59O5u1Sp9nxA"
            'publicKey = "Uzg3fDg0OgYCAAAAJAAAUlNBMQACAAABAAEAp6+IziBn6pKjSMh+PBAvUUtiSbUqJPUpcBJWdSkgeYwfJo2FBE0PuuIqtS3vFUK9PLnGsQ1k7/0F/RnaGzhC6A"
            'privateKey = Base64_Decode(privateKey)
            'publicKey = Base64_Decode(publicKey)
            'privateBlob = cryptSession.StringToByteArray(privateKey, True, False)
            'publicBlob = cryptSession.StringToByteArray(publicKey, True, False)
            'cryptSession.ValuePrivateKey = privateKey
            'cryptSession.ValuePublicKey = publicKey
            'privateKey = cryptSession.ByteArrayToString(privateBlob)
            'publicKey = cryptSession.ByteArrayToString(publicBlob)
            
            'cryptSession.ImportPublicKey False
            'cryptSession.ImportPrivateKey True, False
            
            cryptSession.Generate False
            cryptSession.ExportPublicKey False
            cryptSession.ExportPrivateKey True, False
            'sKeyDongle = cryptSession.ValuePublicKey
            'sKeyDonglePrivate = cryptSession.ValuePrivateKey
            
            'publicKey = Base64_Encode(sKeyDongle)
            'privateKey = Base64_Encode(sKeyDonglePrivate)
            
            cryptSession.SessionEnd
        End If
        Set cryptSession = Nothing
    End If
End Sub

Private Sub txtContainer_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case 13
        cmdBtn_Click 0
    Case 8 To 10, 32, 44, 160, 163, 169, 174
        Exit Sub
    Case 0 To 47, 58 To 63, 91 To 96, 123 To 255
        KeyAscii = 0
    End Select
End Sub

Private Sub txtResult_Change()

    If bCypherOn Then Exit Sub
    If Me.txtResult.Text <> vbNullString Then
        'Me.cmdBtn(3).Enabled = True
    Else
        'Me.cmdBtn(3).Enabled = False
    End If
End Sub

Private Sub txtResult_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
    Case 3, 8 To 10, 13, 22, 24, 32, 44, 160, 163, 169, 174
        Exit Sub
    Case 0 To 31, 127 To 255
        KeyAscii = 0
    End Select
End Sub

