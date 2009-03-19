VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDocument 
   Caption         =   "frmDocument"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   Begin RichTextLib.RichTextBox rtfText 
      Height          =   2000
      Left            =   100
      TabIndex        =   0
      Top             =   100
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3519
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmDocument.frx":0000
   End
End
Attribute VB_Name = "frmDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub rtfText_SelChange()
    frmMain.tbToolBar.Buttons("Bold").Value = IIf(rtfText.SelBold, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Italic").Value = IIf(rtfText.SelItalic, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Underline").Value = IIf(rtfText.SelUnderline, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Align Left").Value = IIf(rtfText.SelAlignment = rtfLeft, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Center").Value = IIf(rtfText.SelAlignment = rtfCenter, tbrPressed, tbrUnpressed)
    frmMain.tbToolBar.Buttons("Align Right").Value = IIf(rtfText.SelAlignment = rtfRight, tbrPressed, tbrUnpressed)
End Sub


Private Sub Form_Load()
    Form_Resize
End Sub


Private Sub Form_Resize()
    On Error Resume Next
    rtfText.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 200
    rtfText.RightMargin = rtfText.Width - 400
End Sub

