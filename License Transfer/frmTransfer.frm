VERSION 5.00
Begin VB.Form frmTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "License"
   ClientHeight    =   5490
   ClientLeft      =   1695
   ClientTop       =   2805
   ClientWidth     =   8040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   8040
   Begin VB.Frame Frame1 
      Caption         =   "License OUT (Step 1 of 2)"
      Height          =   3735
      Left            =   3720
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Height          =   315
         Left            =   3240
         TabIndex        =   9
         Top             =   3180
         Width           =   855
      End
      Begin VB.TextBox txtPath 
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   3180
         Width           =   3015
      End
      Begin VB.Label Label2 
         Caption         =   "Transfer Out "
         ForeColor       =   &H00FF0000&
         Height          =   3315
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   3915
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   2
      Left            =   5880
      TabIndex        =   5
      Top             =   4080
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next >"
      Height          =   375
      Index           =   1
      Left            =   3960
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continue Transfer Later"
      Height          =   375
      Index           =   0
      Left            =   3960
      TabIndex        =   3
      Top             =   4680
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transfer license INTO this computer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   1
      Left            =   2160
      TabIndex        =   1
      Top             =   1920
      Width           =   4035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Transfer license OUT of this computer"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Index           =   0
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   4035
   End
   Begin VB.Image ImageIn 
      BorderStyle     =   1  'Fixed Single
      Height          =   5100
      Left            =   120
      Picture         =   "FRMTRA~1.frx":0000
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Image ImageOut 
      BorderStyle     =   1  'Fixed Single
      Height          =   5100
      Left            =   120
      Picture         =   "FRMTRA~1.frx":2921
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00FF0000&
      Height          =   795
      Left            =   2400
      TabIndex        =   2
      Top             =   3600
      Width           =   4395
   End
End
Attribute VB_Name = "frmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mode As Integer
Dim selectedDirectory As String

Private Sub cmdBrowse_Click()
If txtPath.Text = "" Then txtPath.Text = CurDir
Call BrowseFolders(frmTransfer, txtPath, "Location of license-file")
selectedDirectory = txtPath
Command2(1).Enabled = directoryExists(selectedDirectory)
End Sub

Function directoryExists(ByVal Path As String) As Boolean
'Returns TRUE or FALSE if the named directory exists
On Error GoTo errorDirectoryExists
If Path = "" Then Exit Function
If Dir(Path, vbDirectory) <> "" Then directoryExists = True
Exit Function

errorDirectoryExists:
directoryExists = False

End Function

Private Sub Command1_Click(index As Integer)
If index = 0 Then 'transfer out
    mode = 0
    ImageOut.Visible = True
Else
    mode = 2
    ImageIn.Visible = True
End If
Label1.Visible = False
Command2(1).Visible = True
Command2(2).Visible = True
Command1(0).Visible = False
Command1(1).Visible = False
Command2_Click (1) 'next
End Sub


Private Sub Command2_Click(index As Integer)
Dim i As Integer
On Error Resume Next
Select Case index
Case 0 'continue later
    Unload Me
Case 1 'next
    If InStr(1, Command2(1).caption, "Finish", vbTextCompare) > 0 Then Unload Me
    mode = mode + 1
    Command2(0).Visible = CBool(mode = 4)
    With Frame1
        .Left = 3720                    'Command1(0).Left
        .Top = Command1(0).Top - 90
        '.Height = ImageOut.Height       '+ 90
        '.Width = Command2(2).Left + Command2(2).Width - .Left
        Select Case mode
        Case 1
            .caption = "License OUT (Step 1 of 2)"
            Label2.caption = "Transfer Out enables you to transfer a license from this program to an unlicensed copy on another computer" & vbCrLf & vbCrLf & _
            "To begin, run the unlicensed copy on the remote computer, select ""Transfer License IN"", and follow the on-screen instructions" & vbCrLf & vbCrLf & _
            "The license file can be transferred between computers on a removable disc, a USB drive, or through the Network" & vbCrLf & vbCrLf & _
            "Select the directory where you'll store the license file, and then press ""Next"""
        Case 2
            .caption = "License OUT (Step 2 of 2)"
             Label2.caption = "Congratulations!" & vbCrLf & vbCrLf & _
            "The license has been successfully transferred to " & vbCrLf & _
            selectedDirectory & vbCrLf & vbCrLf & _
            "Now, go to the OTHER computer to complete the transfer process.  If the license is on removable media (not on the Network) take it along with you" & vbCrLf & vbCrLf & "Press ""Finish"" to continue"
        Case 3
            .caption = "License IN (Step 1 of 3)"
            Label2.caption = "Transfer In enables you to transfer a license to this program from a licensed copy on another computer" & vbCrLf & vbCrLf & _
            "For this process you will need access to a licensed copy of this program on another computer and a place to store the license file during the transfer process" & vbCrLf & vbCrLf & _
            "This can be a removable disc, a USB drive, or a Network directory accessible by both computers" & vbCrLf & vbCrLf & _
            "Select the directory where you'll store the license file, and then press ""Next"""
        Case 4
            .caption = "License IN (Step 2 of 3)"
            Label2.caption = "Now the license needs to be copied to:" & vbCrLf & selectedDirectory & vbCrLf & vbCrLf & _
            "1. If the license is on removable media (i.e. not on the        Network), take it to the computer with the licensed         copy of this program" & vbCrLf & vbCrLf & _
            "2. Run the licensed copy of this program and select           ""Transfer Out""" & vbCrLf & vbCrLf & _
            "3. Come back to this computer (bring along the                  removable media if necessary) and press ""Next""" & vbCrLf & vbCrLf & _
            "If there is a significant delay between transfers (e.g, if you are moving the license between your home and work computers) you can suspend the transfer process and resume it later.  Press ""Continue Transfer Later"""
        Case 5
            .caption = "License IN (Step 3 of 3)"
            Label2.caption = "Congratulations!" & vbCrLf & vbCrLf & _
            "The license has been successfully transferred" & vbCrLf & vbCrLf & _
            "Press ""Finish"" to continue"
        End Select
        .Visible = True
    End With
    cmdBrowse.Visible = CBool(mode = 1) Or CBool(mode = 3)
    txtPath.Visible = cmdBrowse.Visible
    
    If mode = 2 Or mode = 5 Then
        Command2(1).caption = "Finish"
        Command2(2).Visible = False
    End If
    Command2(1).Enabled = directoryExists(selectedDirectory)
    'Command2(1).Enabled = True 'only while debugging
Case 2 'cancel
    Unload Me
End Select
End Sub
Private Sub Form_Initialize()
If selectedDirectory = "" Then selectedDirectory = "A:\"
End Sub

Private Sub Form_Load()
With Me
    .Width = 8160

    .Height = 5895
    .Left = Screen.Width / 2 - .Width / 2
    .Top = Screen.Height / 2 - .Height / 2
End With

Label1.caption = "Questions?" & vbCrLf & vbCrLf & "Call Ismail Alkan (Foster Wheeler North America) at (908)713-2477 or e-mail Ismail_Alkan@FWEC.COM"




End Sub


Private Sub Form_Unload(Cancel As Integer)
End
End Sub


