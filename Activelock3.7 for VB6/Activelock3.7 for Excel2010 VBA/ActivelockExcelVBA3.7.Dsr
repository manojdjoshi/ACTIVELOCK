VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} ActivelockVBADesigner 
   ClientHeight    =   6285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8250
   _ExtentX        =   14552
   _ExtentY        =   11086
   _Version        =   393216
   Description     =   $"ActivelockExcelVBA3.7.dsx":0000
   DisplayName     =   "Activelock v3.3 Sample for Excel VBA"
   AppName         =   "Microsoft Excel"
   AppVer          =   "Microsoft Excel 9.0"
   LoadName        =   "Load on demand"
   LoadBehavior    =   9
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel"
End
Attribute VB_Name = "ActivelockVBADesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
On Error Resume Next
thereIsAProblem = False
'ChangePath
frmActivelockRegistration.Show vbModal
If thereIsAProblem = False Then AddInInst.Object = Me
'The following is necessary to activate the ActivelockExcelVBA3.7.DLL
Dim rtn As Boolean
rtn = GetActivelockMessage()
'This is an example of how the forms could be reached from Excel
'Public Property Get DLLfrmSplash()
'    Set DLLfrmSplash = frmSplash
'End Property
Unload frmActivelockRegistration
Set frmActivelockRegistration = Nothing
End Sub
Function GetActivelockMessage_wrap()
GetActivelockMessage_wrap = GetActivelockMessage()
End Function

Public Property Get DLLfrmActivelockRegistration()
    Set DLLfrmActivelockRegistration = frmActivelockRegistration
End Property

