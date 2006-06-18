VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProductsStorage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Products Storage"
   ClientHeight    =   2220
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frmProductsStorage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2220
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   60
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5460
      TabIndex        =   6
      Top             =   1860
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   4050
      TabIndex        =   5
      Top             =   1860
      Width           =   1365
   End
   Begin VB.Frame frameProductsStorageType 
      Caption         =   "Storage Type"
      Height          =   1725
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   6705
      Begin VB.TextBox txtConnectionString 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1860
         TabIndex        =   11
         Top             =   1320
         Width           =   4755
      End
      Begin VB.CommandButton cmdBrowseForStorageFile 
         Caption         =   "..."
         Height          =   315
         Left            =   6330
         TabIndex        =   9
         Top             =   420
         Width           =   285
      End
      Begin VB.TextBox txtStorageFile 
         Height          =   315
         Left            =   1860
         TabIndex        =   7
         Top             =   420
         Width           =   4455
      End
      Begin VB.OptionButton optStorageMSSQL 
         Caption         =   "MSSQL database"
         Enabled         =   0   'False
         Height          =   345
         Left            =   210
         TabIndex        =   4
         Top             =   1230
         Width           =   1605
      End
      Begin VB.OptionButton optStorageMDBFile 
         Caption         =   "MDB file"
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   900
         Width           =   1605
      End
      Begin VB.OptionButton optStorageXmlFile 
         Caption         =   "XML file"
         Height          =   345
         Left            =   210
         TabIndex        =   2
         Top             =   570
         Width           =   1485
      End
      Begin VB.OptionButton optStorageIniFile 
         Caption         =   "INI file"
         Height          =   315
         Left            =   210
         TabIndex        =   1
         Top             =   270
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Connection String:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1860
         TabIndex        =   10
         Top             =   1080
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Storage file:"
         Height          =   255
         Left            =   1860
         TabIndex        =   8
         Top             =   210
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmProductsStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowseForStorageFile_Click()
  'browse for storage file
  On Error GoTo cmdBrowseForStorageFile_Click_Error
  CommonDialog1.FileName = txtStorageFile.Text
  CommonDialog1.Flags = cdlOFNExplorer Or cdlOFNCreatePrompt Or cdlOFNLongNames Or cdlOFNOverwritePrompt Or cdlOFNPathMustExist
  CommonDialog1.DialogTitle = "Select products storage file"
  CommonDialog1.CancelError = True
  
  On Error GoTo cmdBrowseForStorageFile_ShowOpen_Error
  CommonDialog1.ShowOpen
  txtStorageFile.Text = CommonDialog1.FileName

cmdBrowseForStorageFile_ShowOpen_Error:

  On Error GoTo 0
  Exit Sub

cmdBrowseForStorageFile_Click_Error:

  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdBrowseForStorageFile_Click of Form frmProductsStorage"
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  SaveSettings
  frmMain.Form_Load
  Unload Me
End Sub

Private Sub Form_Load()
  LoadSettings
End Sub

Private Sub LoadSettings()
On Error GoTo LoadSettings_Error
  txtStorageFile.Text = frmMain.mProductsStoragePath
  Select Case frmMain.mProductsStoreType
    Case alsINIFile
      optStorageIniFile.Value = True
    Case alsXMLFile
      optStorageXmlFile.Value = True
    Case alsMDBFile
      optStorageMDBFile.Value = True
    'Case alsMSSQL
  End Select
  
  On Error GoTo 0
  Exit Sub

LoadSettings_Error:

  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadSettings of Form frmProductsStorage"
End Sub

Private Sub SaveSettings()
  On Error GoTo SaveSettings_Error
  frmMain.mProductsStoragePath = txtStorageFile.Text
  frmMain.mProductsStoreType = IIf(optStorageIniFile.Value = True, alsINIFile, _
    IIf(optStorageXmlFile.Value = True, alsXMLFile, _
    IIf(optStorageMDBFile.Value = True, alsMDBFile, _
    alsINIFile)))
  
  On Error GoTo 0
  Exit Sub

SaveSettings_Error:

  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveSettings of Form frmProductsStorage"
End Sub
