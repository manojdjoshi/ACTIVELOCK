VERSION 5.00
Begin VB.Form frmReg 
   Caption         =   "Registration...."
   ClientHeight    =   7080
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9645
   Icon            =   "frmReg.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7080
   ScaleWidth      =   9645
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraRegStatus 
      Caption         =   "Status"
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   60
      TabIndex        =   12
      Top             =   90
      Width           =   9495
      Begin VB.TextBox txtRegStatus 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox txtUsedDays 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1140
         Width           =   4335
      End
      Begin VB.TextBox txtExpiration 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1440
         Width           =   4335
      End
      Begin VB.TextBox txtName 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "TestApp"
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox txtVersion 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "1.0"
         Top             =   540
         Width           =   4335
      End
      Begin VB.TextBox txtChecksum 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   2040
         Width           =   1215
      End
      Begin VB.TextBox txtRegisteredLevel 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1740
         Width           =   4335
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   8370
         Picture         =   "frmReg.frx":0442
         ScaleHeight     =   825
         ScaleWidth      =   825
         TabIndex        =   16
         Top             =   210
         Width           =   825
      End
      Begin VB.CommandButton cmdKillTrial 
         Caption         =   "&Kill Trial"
         Height          =   315
         Left            =   8280
         TabIndex        =   15
         ToolTipText     =   "End the Free Trial"
         Top             =   1740
         Width           =   1125
      End
      Begin VB.CommandButton cmdResetTrial 
         Caption         =   "&Reset Trial"
         Height          =   315
         Left            =   8280
         TabIndex        =   14
         ToolTipText     =   "Reset the Free Trial"
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtLicenseType 
         BackColor       =   &H80000013&
         Height          =   285
         Left            =   4140
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   2040
         Width           =   1755
      End
      Begin VB.Label Label10 
         Caption         =   "Warning: Remove the Reset Trial and Kill Trial button for your actual release or production systems."
         Height          =   1005
         Left            =   6300
         TabIndex        =   33
         Top             =   1380
         Width           =   1755
      End
      Begin VB.Label Label6 
         Caption         =   "Registered:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   870
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Days Used:"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label8 
         Caption         =   "Expiry Date:"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1470
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "App Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   270
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "App Version:"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   570
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "DLL Checksum:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   2070
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Registered Level:"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   1770
         Width           =   1335
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         Caption         =   "Activelock V3"
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   8250
         TabIndex        =   25
         Top             =   1050
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "License Type:"
         Height          =   255
         Left            =   2940
         TabIndex        =   24
         Top             =   2070
         Width           =   1155
      End
   End
   Begin VB.Frame fraReg 
      Caption         =   "Register"
      ForeColor       =   &H00FF0000&
      Height          =   4155
      Left            =   60
      TabIndex        =   0
      Top             =   2820
      Width           =   9495
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
         TabIndex        =   8
         Top             =   1380
         Width           =   6675
      End
      Begin VB.CommandButton cmdRegister 
         Caption         =   "&Register"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8220
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   600
         Width           =   6675
      End
      Begin VB.CommandButton cmdReqGen 
         Caption         =   "&Generate"
         Height          =   315
         Left            =   8220
         TabIndex        =   5
         ToolTipText     =   "Generate Installation Code"
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox txtUser 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "Evaluation User"
         Top             =   300
         Width           =   6675
      End
      Begin VB.CommandButton cmdKillLicense 
         Caption         =   "&Kill License"
         Height          =   315
         Left            =   8220
         TabIndex        =   3
         ToolTipText     =   "Kill the License"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdCopy 
         Height          =   345
         Left            =   8220
         MaskColor       =   &H8000000F&
         Picture         =   "frmReg.frx":33CA
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   900
         Width           =   345
      End
      Begin VB.CommandButton cmdPaste 
         Height          =   345
         Left            =   8220
         Picture         =   "frmReg.frx":36DC
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Width           =   345
      End
      Begin VB.Label Label4 
         Caption         =   "Liberation Key:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1380
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Installation Code:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "User Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtReqCodeGen.Text
End Sub

Private Sub cmdKillLicense_Click()
    Dim licFile As String
    licFile = App.Path & "\" & objReg.GetLicenseRoot & ".lic"
    
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

Private Sub cmdPaste_Click()
If Clipboard.GetText = txtReqCodeGen.Text Then
    MsgBox "You cannot paste the Installation Code into the Liberation Key field.", vbExclamation
    Exit Sub
End If
txtLibKeyIn.Text = Clipboard.GetText
End Sub

Private Sub txtLibKeyIn_Change()
    cmdRegister.Enabled = CBool(Trim$(txtLibKeyIn.Text) <> "")
End Sub

Private Sub txtName_Change()
    'MyActiveLock.SoftwareName = txtName
End Sub

Private Sub cmdRegister_Click()
Dim OK As Boolean
    On Error GoTo errHandler
    ' Register this key
    objReg.MyActiveLock.Register txtLibKeyIn
    MsgBox modMain.Dec("386.457.46D.483.4F1.4FC.4E6.42B.4FC.483.4C5.4BA.160.4F1.507.441.441.457.4F1.4F1.462.507.4A4.16B"), vbInformation ' "Registration successful!"
    
    Unload Me
    frmMain.Show
    
    Exit Sub

errHandler:
    MsgBox Err.Number & ": " & Err.Description
End Sub

Private Sub cmdReqGen_Click()
    ' Generate Request code to Lock

    If objReg.MyActiveLock Is Nothing Then
        objReg.SetTrialNoThisTime True
'        Form_Load
    End If
'    If FileExist(strKeyStorePath) And FileLen(strKeyStorePath) <> 0 Then
'        MsgBox "There's an active license file already." & vbCrLf & _
'            "The LockTypes are acquired from the types used in the license file." & vbCrLf & vbCrLf & _
'            "If you don't want this, you must kill the existing license first.", vbInformation
'    End If
    If txtRegStatus.Text <> "Registered" Then txtRegStatus.Text = ""
    If Not IsNumeric(txtUsedDays.Text) Then txtUsedDays.Text = ""
    txtReqCodeGen.Text = objReg.MyActiveLock.InstallationCode(txtUser)
    
End Sub

Private Sub cmdResetTrial_Click()
    objReg.fResetTrial
End Sub

'Private Type ALVariables
'    Title As String
'    SoftwareName As String
'    SoftwareVersion As String
'    CheckSum As Long
'    RegStatus As String
'    UsedDays As Integer
'    LicenseType As String
'    Expiration As String
'    User As String
'    RegisteredLevel As String
'End Type

Private Sub Form_Load()
    If Not objReg Is Nothing Then
        'What to do here....? Initiate it again..?
        'Leave this for Later
    End If
    
    txtName.Text = objReg.GetSoftwareName
    txtVersion.Text = objReg.GetVersion
    txtChecksum.Text = objReg.GetCheckSum
    txtUser.Text = objReg.GetUser
    txtLicenseType.Text = objReg.GetLicenseType
    txtRegStatus.Text = objReg.GetRegStatus
    txtUsedDays.Text = objReg.GetUsedDaysDesc
End Sub



Function FileExist(ByVal TestFileName As String) As Boolean
'This function checks for the existance of a given
'file name. The function returns a TRUE or FALSE value.
'The more complete the TestFileName string is, the
'more reliable the results of this function will be.

'Declare local variables
Dim OK As Integer

'Set up the error handler to trap the File Not Found
'message, or other errors.
On Error GoTo FileExistErrors:

'Check for attributes of test file. If this function
'does not raise an error, than the file must exist.
OK = GetAttr(TestFileName)

'If no errors encountered, then the file must exist
FileExist = True
Exit Function

FileExistErrors:    'error handling routine, including File Not Found
    FileExist = False
    Exit Function 'end of error handler
End Function
