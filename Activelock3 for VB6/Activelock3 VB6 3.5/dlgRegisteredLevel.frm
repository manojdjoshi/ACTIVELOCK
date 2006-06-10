VERSION 5.00
Begin VB.Form dlgRegisteredLevel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registered Level Properties"
   ClientHeight    =   1470
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   5925
   Icon            =   "dlgRegisteredLevel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1470
   ScaleWidth      =   5925
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3375
      TabIndex        =   5
      Top             =   990
      Width           =   1140
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4635
      TabIndex        =   4
      Top             =   990
      Width           =   1140
   End
   Begin VB.TextBox txtItemData 
      Height          =   330
      Left            =   4635
      TabIndex        =   3
      Top             =   315
      Width           =   1140
   End
   Begin VB.TextBox txtDescription 
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   315
      Width           =   4245
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000016&
      X1              =   135
      X2              =   5760
      Y1              =   870
      Y2              =   870
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000015&
      X1              =   135
      X2              =   5715
      Y1              =   855
      Y2              =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Item Data:"
      Height          =   240
      Left            =   4635
      TabIndex        =   2
      Top             =   90
      Width           =   1140
   End
   Begin VB.Label Label1 
      Caption         =   "Description:"
      Height          =   240
      Left            =   135
      TabIndex        =   0
      Top             =   90
      Width           =   2355
   End
End
Attribute VB_Name = "dlgRegisteredLevel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<CSCC>
'--------------------------------------------------------------------------------
'    Component  : dlgRegisteredLevel
'    Project    : ALUGEN3
'
'    Description: For Add and Modify a Registered Level
'
'    Modified   : By Kirtaph on 2006-02-16
'--------------------------------------------------------------------------------
'</CSCC>

Option Explicit

Private Const NumberOfControl As Integer = 2

Private blnAnnullato As Boolean

Private blnCanEnabled(NumberOfControl - 1) As Boolean

Private m_strDescription As String

Private m_intItemData As Integer

Public Property Get ItemData() As Integer
    If Val(txtItemData) > 32767 Then
        MsgBox "ItemData cannot be greater than 32767.", vbExclamation
        Exit Sub
    End If
    ItemData = Val(txtItemData)
End Property

Public Property Let ItemData(ByVal intValue As Integer)
    m_intItemData = intValue
End Property

Public Property Get Description() As String
    Description = txtDescription
End Property

Public Property Let Description(ByVal strValue As String)
    m_strDescription = strValue
End Property

Public Property Get Annullato() As Boolean
    Annullato = blnAnnullato
End Property

Private Function CanEnable() As Boolean

    Dim inti As Integer

    CanEnable = True

    For inti = 0 To UBound(blnCanEnabled)

        CanEnable = CanEnable And blnCanEnabled(inti)

    Next

End Function

Private Function InitializeEnable()

    Dim inti As Integer

    For inti = 0 To UBound(blnCanEnabled)

        blnCanEnabled(inti) = False

    Next

End Function

Private Sub cmdCancel_Click()

    blnAnnullato = True
    Me.Hide

End Sub

Private Sub cmdOk_Click()

    blnAnnullato = False
    Me.Hide

End Sub

Private Sub Form_Load()
    InitializeEnable
    txtItemData = m_intItemData
    txtDescription = m_strDescription
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, _
                             UnloadMode As Integer)

    blnAnnullato = True
    'Set myClassVector = Nothing

End Sub

Private Sub txtDescription_Change()

    Dim intControlNumber As Integer
    intControlNumber = 0

    If txtDescription.Text = "" Then
     
        blnCanEnabled(intControlNumber) = False

    Else
       
        blnCanEnabled(intControlNumber) = True

    End If

    cmdOk.Enabled = CanEnable

End Sub

Private Sub txtItemData_Change()

    Dim intControlNumber As Integer
    intControlNumber = 1
    If txtItemData.Text = "" Then
        blnCanEnabled(intControlNumber) = False
    Else
        blnCanEnabled(intControlNumber) = True
    End If
    cmdOk.Enabled = CanEnable

End Sub
