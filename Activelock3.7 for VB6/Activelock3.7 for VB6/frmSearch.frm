VERSION 5.00
Begin VB.Form frmSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Results"
   ClientHeight    =   2385
   ClientLeft      =   8040
   ClientTop       =   8070
   ClientWidth     =   5235
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2385
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   720
      TabIndex        =   10
      Top             =   120
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search direction"
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   2415
      Begin VB.OptionButton Option1 
         Caption         =   "Downward from top of file"
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   2200
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Downward from current line"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   2250
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Upward from current line"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   2200
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Upward from end of file"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   2200
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   4080
      TabIndex        =   3
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Find"
      Height          =   375
      Index           =   0
      Left            =   4080
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CheckBox Check1 
      Caption         =   " Ignore case (e.g. ABC = abc)"
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   1200
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.Label lblMessage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   2640
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Find..."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'API call for force a window to remain on top of all other windows
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Sub Combo1_Change()
Command1(0).Enabled = Trim(Combo1) <> ""
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Command1(0).Enabled Then Command1_Click (0) 'enter
If KeyAscii = 27 Then Command1_Click (1) 'escape
End Sub


Private Sub Command1_Click(Index As Integer)
Dim startAt As Integer, endAt As Integer, increment As Integer, i As Integer, j As Integer
Dim firstVisible As Integer, lastVisible As Integer, numberVisible As Integer
Dim found As Boolean
Dim searchString As String

numberVisible = frmAlugenDatabase.List1.Height / frmAlugenDatabase.TextHeight("X") - 1 'window height, lines

lblMessage = ""
If Index = 1 Then   'cancel
    frmSearch.Hide
    Exit Sub
End If
searchString = Combo1
found = False
    For i = 0 To Combo1.ListCount - 1   'check for duplicates
    If Combo1.List(i) = searchString Then
        found = True
        Exit For
    End If
    Next i
If Not found Then Combo1.AddItem searchString
If Combo1.ListCount > 12 Then Combo1.RemoveItem (0) 'remove oldest item
Combo1.SetFocus 'snap focus back to combo box so user doesn't have to

If Check1 = vbChecked Then searchString = UCase(searchString)

If Option1(0).Value Then    'search downward, beginning at start of file
    startAt = 0
    endAt = frmAlugenDatabase.List1.ListCount - 1
    increment = 1
ElseIf Option1(1).Value Then    'search downward, beginning at next line
    startAt = frmAlugenDatabase.List1.ListIndex + 1
    endAt = frmAlugenDatabase.List1.ListCount - 1
    increment = 1
ElseIf Option1(2).Value Then    'search upwards, beginning at previous line
    startAt = frmAlugenDatabase.List1.ListIndex - 1
    endAt = 0
    increment = -1
ElseIf Option1(3).Value Then    'search upwards, beginning at last line
    startAt = frmAlugenDatabase.List1.ListCount - 1
    endAt = 0
    increment = -1
End If

found = False
    For i = startAt To endAt Step increment
    If Check1 = vbChecked Then 'case doesn't matter
        If InStr(UCase(frmAlugenDatabase.List1.List(i)), searchString) > 0 Then found = True
    Else    'case matter
        If InStr(frmAlugenDatabase.List1.List(i), searchString) > 0 Then found = True
    End If
    
    If found Then
        frmAlugenDatabase.List1.Selected(frmAlugenDatabase.List1.ListIndex) = False 'turn off old highlight
        lblMessage = "Found on line " & i + 1 & " of " & frmAlugenDatabase.List1.ListCount
        firstVisible = frmAlugenDatabase.List1.TopIndex
        lastVisible = firstVisible + frmAlugenDatabase.List1.Height / frmAlugenDatabase.TextHeight("X") - 1
        If i >= firstVisible And i <= lastVisible Then  'line shows on screen
        Else
            j = i - numberVisible / 2
            If j < 0 Then j = 0
            frmAlugenDatabase.List1.TopIndex = j
        End If
        frmAlugenDatabase.List1.Selected(i) = True 'turn on new highlight
        
        If Option1(0) Then Option1(1).Value = True  'change next search
        If Option1(3) Then Option1(2).Value = True
        Exit Sub
    End If
    Next i
lblMessage = "Search string not found!"
End Sub

Private Sub Form_Load()
' Set some constant values (from WINAPI.TXT).
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40
Dim leftEdge As Integer, topEdge As Integer, formHeight As Integer, formWidth As Integer
leftEdge = (frmAlugenDatabase.Left + frmAlugenDatabase.Width / 2 - Me.Width / 2) / Screen.TwipsPerPixelX
topEdge = (frmAlugenDatabase.Top + frmAlugenDatabase.Height - Me.Height - 700) / Screen.TwipsPerPixelY
formWidth = 5355 / Screen.TwipsPerPixelX
formHeight = 2790 / Screen.TwipsPerPixelY

' Turn on the TopMost attribute.
SetWindowPos hWnd, conHwndTopmost, leftEdge, topEdge, formWidth, formHeight, conSwpNoActivate Or conSwpShowWindow

Command1(0).Enabled = Trim(Combo1) <> ""
frmSearch.Show
Combo1.SetFocus 'put initial focus on combo box
End Sub




