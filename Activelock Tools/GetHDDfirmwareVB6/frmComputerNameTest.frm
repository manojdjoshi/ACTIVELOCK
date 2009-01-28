VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Activelock HDD Firmware Serial Number Detection Routines"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   ScaleHeight     =   4155
   ScaleWidth      =   7380
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   6240
      Picture         =   "frmComputerNameTest.frx":0000
      ScaleHeight     =   735
      ScaleWidth      =   975
      TabIndex        =   11
      Top             =   3120
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   330
      Left            =   3600
      TabIndex        =   9
      Top             =   2040
      Width           =   3000
   End
   Begin VB.TextBox Text4 
      Height          =   330
      Left            =   3600
      TabIndex        =   7
      Top             =   1620
      Width           =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get the HDD Firmware Serial Number"
      Height          =   960
      Left            =   1440
      TabIndex        =   6
      Top             =   3000
      Width           =   4425
   End
   Begin VB.TextBox Text3 
      Height          =   330
      Left            =   3600
      TabIndex        =   5
      Top             =   1200
      Width           =   3000
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   3600
      TabIndex        =   3
      Top             =   810
      Width           =   3000
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   3600
      TabIndex        =   1
      Top             =   405
      Width           =   3000
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "VB6 Version"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   2640
      Width           =   3255
   End
   Begin VB.Label Label6 
      Caption         =   "VB6 API with WMI (Win32_DiskDrive)"
      Height          =   330
      Left            =   360
      TabIndex        =   10
      Top             =   2040
      Width           =   3255
   End
   Begin VB.Label Label4 
      Caption         =   "VB6 API with SCSI Modified (NOT in AL)"
      Height          =   330
      Left            =   360
      TabIndex        =   8
      Top             =   1620
      Width           =   3255
   End
   Begin VB.Label Label3 
      Caption         =   "VB6 API with SCSI Back Door"
      Height          =   330
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "VB6 with ALCrypto (WinSim DISKID32)"
      Height          =   330
      Left            =   360
      TabIndex        =   2
      Top             =   810
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "VB6 with SMART VxD"
      Height          =   330
      Left            =   360
      TabIndex        =   0
      Top             =   405
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, _
                ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, _
                lpOutBuffer As Any, ByVal nOutBufferSize As Long, _
                lpBytesReturned As Long, lpOverlapped As Any) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, _
                ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
                ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, _
                ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Private Type XBUFFER
     Buff(559) As Byte
End Type
Public Function GetSerial(ByVal iController As Integer, Optional ByVal isMaster As Boolean = True) As String
    Dim s As String
    Dim i As Integer
    Dim h As Long
    Dim o As Long
    Dim rV As String
    
    'From:
    'http://discuss.develop.com/archives/wa.exe?A2=ind0309a&L=advanced-dotnet&D=0&T=0&P=3760
 
    rV = ""
    GetSerial = ""
    
    Dim retVal As Integer
    h = CreateFile("\\.\Scsi" & iController & ":", &H80000000 Or &H40000000, &H1 Or &H2, ByVal 0&, 3, 0, 0)
     If (h <> -1) Then
     
        Dim b As XBUFFER
        Dim b1 As XBUFFER
        For i = 0 To 559  '556?
        b.Buff(i) = 0
        b1.Buff(i) = 0
        Next i
        
        b.Buff(0) = 28:     b.Buff(4) = 83:     b.Buff(5) = 67:     b.Buff(6) = 83:     b.Buff(7) = 73
        b.Buff(8) = 68:     b.Buff(9) = 73:     b.Buff(10) = 83:    b.Buff(11) = 75:    b.Buff(12) = 16
        b.Buff(13) = 39:    b.Buff(16) = 1:     b.Buff(17) = 5:     b.Buff(18) = 27:    b.Buff(24) = 20 '17?
        b.Buff(25) = 2:     b.Buff(38) = 236 '&HEC
        
        Select Case isMaster
            Case True
                b.Buff(40) = 0
                'rV = rV & "Controller = " & iController & ", Master Drive" & vbCrLf
            Case False
                b.Buff(40) = 1
                'rV = rV & "Controller = " & iController & ", Slave Drive" & vbCrLf
        End Select
        
        Dim u As Long
        
        'u = DeviceIoControl(h, 315400, b, 63, b1, 560, o, 0)
        u = DeviceIoControl(h, 315400, b, Len(b), b1, Len(b1), o, 0)
            
        If (u) Then
            'HDD Serial Number
            For retVal = 64 To 83
            s = s & Chr(b1.Buff(retVal))
            Next
            rV = rV & Swap(s)   '"HDD Firmware Serial Number:     " & Swap(s) & vbCrLf
            
            'HDD Model Number
            s = ""
            For retVal = 98 To 137
            s = s & Chr(b1.Buff(retVal))
            Next
            'rV = rV & "HDD Model Number:               " & Swap(s) & vbCrLf
            
            'HDD Controller Revision Number
            s = ""
            For retVal = 90 To 97
            s = s & Chr(b1.Buff(retVal))
            Next
            'rV = rV & "HDD Controller Revision Number: " & Swap(s) & vbCrLf
            'rV = rV & "===========================================================" & vbCrLf
            
            GetSerial = rV
        End If
      End If
    CloseHandle h
    
End Function

Private Function Swap(ByVal os As String) As String
Dim i As Integer
Dim j As Integer
Dim ds As String
Dim ts1 As String
Dim ts2 As String

ds = ""
For i = 0 To Len(os) / 2
ts1 = Mid$(os, 2 * i + 1, 1)
ts2 = Mid$(os, 2 * i + 2, 1)
ds = ds & ts2 & ts1
Next i

Swap = Trim(LTrim(ds))
End Function



Private Sub Command1_Click()
Text1.Text = GetHDSerialFirmware(0)
Text2.Text = GetHDSerialFirmware(1)
Text3.Text = GetHDSerialFirmware(2)
Dim i As Integer
For i = 0 To 15
    Text4.Text = GetSerial(i, True)
    If Text4.Text <> "" Then Exit For
Next i
'Text5.Text = GetHDSerialFirmware(3)
Text6.Text = GetHDSerialFirmware(4)

End Sub

