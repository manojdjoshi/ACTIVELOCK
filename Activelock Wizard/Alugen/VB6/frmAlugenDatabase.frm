VERSION 5.00
Begin VB.Form frmAlugenDatabase 
   BorderStyle     =   0  'None
   Caption         =   "ActiveLock3 Universal GENerator - License Database"
   ClientHeight    =   8040
   ClientLeft      =   2115
   ClientTop       =   2115
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   9735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H00FFFFC0&
      Height          =   2010
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuSort 
         Caption         =   "&Sort by..."
         Begin VB.Menu mnuKey 
            Caption         =   "mnuKey"
            Index           =   0
         End
      End
      Begin VB.Menu mnuFind 
         Caption         =   "&Find"
      End
   End
End
Attribute VB_Name = "frmAlugenDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const LICENSES_FILE = "authorizations.txt"
Dim sortOptions As Variant
Private Type storageType
    D(0 To 10) As String
End Type
Dim storage() As storageType
Dim systemEvent As Boolean
Public Sub ArchiveLicense(productName As String, _
    userName As String, _
    registrationDate As String, expiresAfter As String, licenseType As String, _
    registeredLevel As String, installationCode As String, liberationCode As String, _
    lockTypes As String)
Dim fileNumber As Integer, i As Integer, j As Integer
Dim counter As Integer, attb As Integer
Dim s As String

If fileExist(App.path & "\" & LICENSES_FILE) Then
    attb = GetAttr(App.path & "\" & LICENSES_FILE)
    If attb <> vbArchive Then Call SetAttr(App.path & "\" & LICENSES_FILE, vbArchive)
    fileNumber = FreeFile
    Open App.path & "\" & LICENSES_FILE For Input As #fileNumber
        s = Input(LOF(fileNumber), #fileNumber)
    Close fileNumber
End If

fileNumber = FreeFile
Open App.path & "\" & LICENSES_FILE For Output As #fileNumber
Print #fileNumber, s
If s <> "" Then Print #fileNumber, "~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~"
systemEvent = True
Print #fileNumber, mnuKey(0).Caption & ": " & productName
Print #fileNumber, mnuKey(1).Caption & ": " & userName
Print #fileNumber, mnuKey(2).Caption & ": " & registrationDate
Print #fileNumber, mnuKey(3).Caption & ": " & expiresAfter
Print #fileNumber, mnuKey(4).Caption & ": " & licenseType
Print #fileNumber, mnuKey(5).Caption & ": " & registeredLevel
Print #fileNumber, mnuKey(6).Caption & ": " & installationCode
Print #fileNumber, mnuKey(7).Caption & ": " & liberationCode
Print #fileNumber, mnuKey(8).Caption & ": " & lockTypes
systemEvent = False

Close fileNumber
Screen.MousePointer = vbDefault
End Sub
Function inString(ByVal x As String, ParamArray MyArray()) As Boolean
'Do ANY of a group of sub-strings appear in within the first string?
'Case doesn't count and we don't care WHERE or WHICH
Dim Y As Variant    'member of array that holds all arguments except the first
    For Each Y In MyArray
    If InStr(1, x, Y, 1) > 0 Then 'the "ones" make the comparison case-insensitive
        inString = True
        Exit Function
    End If
    Next Y
End Function

Function directoryExists(ByVal path As String) As Boolean

'Returns TRUE or FALSE if the named directory exists
On Error GoTo errorDirectoryExists

If Dir(path, vbDirectory) <> "" Then directoryExists = True
Exit Function

errorDirectoryExists:
directoryExists = False

End Function
Public Function fileExist(ByVal TestFileName As String) As Boolean
'This function checks for the existance of a given
'file name. The function returns a TRUE or FALSE value.
'The more complete the TestFileName string is, the
'more reliable the results of this function will be.

'Declare local variables
Dim ok As Integer


'Set up the error handler to trap the File Not Found
'message, or other errors.
On Error GoTo FileExistErrors:


'Check for attributes of test file. If this function
'does not raise an error, than the file must exist.
ok = GetAttr(TestFileName)

'If no errors encountered, then the file must exist
fileExist = True
Exit Function

FileExistErrors:    'error handling routine, including File Not Found
fileExist = False
Exit Function 'end of error handler
End Function

Private Sub Form_Load()
Dim i As Integer, fileNumber As Integer, counter As Integer, j As Integer
Dim s As String
'On Error Resume Next
Me.Width = 11470
Me.Height = 9705

With List1
    .Width = Me.ScaleWidth - 2 * .Left
    .Height = Me.ScaleHeight - 30
End With
sortOptions = Split("Program Name and Version:User Name:Registration Date:Expires After:License Type:Registered Level:Installation/Site Code:Liberation/Unlock Code:Lock Types", ":")
For i = 0 To UBound(sortOptions)
    If i > 0 Then
        Load mnuKey(i)
        mnuKey(i).Visible = True
    End If
    mnuKey(i).Caption = sortOptions(i)
Next i

'Center Me
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmAlugenDatabase = Nothing
End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If List1.SelCount = 1 Then List1.Selected(List1.ListIndex) = False
End Sub


Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuFind_Click()
If Not frmSearch.Visible Then frmSearch.Show vbModal
End Sub

Private Sub mnuKey_Click(Index As Integer)
Dim i As Integer, j As Integer
Dim printed As Boolean
If frmSearch.Visible Then frmSearch.Hide

Screen.MousePointer = vbHourglass
Do
    storage(0).D(0) = "!!!!"
    For i = 2 To UBound(storage)
        If StrComp(storage(i).D(Index), storage(i - 1).D(Index), vbTextCompare) < 0 Then
            storage(0) = storage(i)
            storage(i) = storage(i - 1)
            storage(i - 1) = storage(0)
        End If
    Next i
Loop Until storage(0).D(0) = "!!!!"

Dim LBHS As New CLBHScroll
LBHS.Init List1
List1.Clear

For i = 1 To UBound(storage)
    printed = False
    With storage(i)
            For j = 0 To UBound(.D)
            If .D(j) <> "" Then
                printed = True
                LBHS.AddItem .D(j)
            End If
            Next j
    End With
    If printed Then LBHS.AddItem ""
Next i
Screen.MousePointer = vbDefault
Me.Caption = "Authorizations sorted by " & mnuKey(Index).Caption
End Sub


Private Sub mnuPrint_Click()
Dim i As Integer
With List1
    If .ListCount = 0 Then Exit Sub
    Screen.MousePointer = vbHourglass
        For i = 0 To .ListCount - 1
            If .SelCount = 0 Or .Selected(i) Then Printer.Print
        Next i
    Printer.EndDoc
End With
Screen.MousePointer = vbDefault
End Sub

Public Sub ShowArchive()
Dim i As Integer, fileNumber As Integer, counter As Integer, j As Integer
Dim s As String
ReDim storage(1)
counter = 1

If fileExist(App.path & "\" & LICENSES_FILE) = False Then
    ' Create an empty authorizations.txt file if it doesn't exists
    Dim hFile As Long
    hFile = FreeFile
    Open App.path & "\" & LICENSES_FILE For Output As #hFile
    Close #hFile
End If

fileNumber = FreeFile
Open App.path & "\" & LICENSES_FILE For Input As #fileNumber
    Do While Not EOF(fileNumber)
    Line Input #fileNumber, s
    'Debug.Print s
    s = Trim(s)
    If inString(s, "~~~~~~~~") Then
        counter = counter + 1
        If counter > UBound(storage) Then ReDim Preserve storage(counter)
    ElseIf s = "" Then
        'skip
    Else
        For i = 0 To UBound(sortOptions)
            j = Len(sortOptions(i))
            If StrComp(sortOptions(i), Left$(s, j), vbTextCompare) = 0 Then
                    Do While inString(s, ": ")
                    s = Replace(s, ": ", ":")
                    Loop
                s = Replace(s, Left$(s, j) & ":", Left$(s, j) & ": ")
                storage(counter).D(i) = s
                Exit For
            End If
        Next i
    End If
    Loop
Close #fileNumber

mnuKey_Click (0)

End Sub
