VERSION 5.00
Begin VB.Form frmLevelManager 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Level Manager"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   4605
   Icon            =   "frmLevelManager.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstRegisteredLevel 
      Height          =   4155
      ItemData        =   "frmLevelManager.frx":000C
      Left            =   135
      List            =   "frmLevelManager.frx":000E
      TabIndex        =   3
      Top             =   405
      Width           =   4350
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3330
      TabIndex        =   2
      Top             =   4815
      Width           =   1140
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1350
      TabIndex        =   1
      Top             =   4815
      Width           =   1140
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   135
      TabIndex        =   0
      Top             =   4815
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "List of Registered Level:"
      Height          =   240
      Left            =   135
      TabIndex        =   4
      Top             =   90
      Width           =   3345
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000016&
      X1              =   135
      X2              =   4470
      Y1              =   4695
      Y2              =   4695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      X1              =   135
      X2              =   4470
      Y1              =   4680
      Y2              =   4680
   End
End
Attribute VB_Name = "frmLevelManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : frmLevelManager
'    Project    : ALUGEN3
'
'    Description: For organize your personal Registered Level
'
'    Modified   : By Kirtaph On 2006-02-16
'--------------------------------------------------------------------------------
'</CSCC>

Option Explicit

Private Sub cmdAdd_Click()

    With dlgRegisteredLevel
        .ItemData = 0
        .Show vbModal

        If Not .Annullato Then
            lstRegisteredLevel.AddItem .Description
            lstRegisteredLevel.ItemData(lstRegisteredLevel.NewIndex) = .ItemData
        End If

    End With

    Unload dlgRegisteredLevel
End Sub

Private Sub cmdClose_Click()
    SaveComboBox strRegisteredLevelDBName, lstRegisteredLevel, True
    Me.Hide
End Sub

Private Sub cmdRemove_Click()
    On Error Resume Next
    lstRegisteredLevel.RemoveItem lstRegisteredLevel.ListIndex
    lstRegisteredLevel.ListIndex = 0
    lstRegisteredLevel.SetFocus
End Sub

Private Sub Form_Load()
    LoadComboBox strRegisteredLevelDBName, lstRegisteredLevel, True
    lstRegisteredLevel.ListIndex = 0
End Sub

Private Sub lstRegisteredLevel_DblClick()

    If lstRegisteredLevel.ListCount > 0 Then

        With dlgRegisteredLevel
            .ItemData = lstRegisteredLevel.ItemData(lstRegisteredLevel.ListIndex)
            .Description = lstRegisteredLevel.List(lstRegisteredLevel.ListIndex)
            .Show vbModal

            If Not .Annullato Then
                lstRegisteredLevel.List(lstRegisteredLevel.ListIndex) = .Description
                lstRegisteredLevel.ItemData(lstRegisteredLevel.ListIndex) = .ItemData
            End If

        End With

        Unload dlgRegisteredLevel
    End If

End Sub
