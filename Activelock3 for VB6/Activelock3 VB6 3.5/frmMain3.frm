VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ALUGEN - ActiveLock3 Universal GENerator"
   ClientHeight    =   8190
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9735
   Icon            =   "frmMain3.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   35
      Top             =   7875
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   17119
            Text            =   "Ready"
            TextSave        =   "Ready"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7875
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   13891
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Pro&duct Code Generator"
      TabPicture(0)   =   "frmMain3.frx":0CCA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label17"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "gridProds"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraProdNew"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdRemove"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Picture1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdValidate"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "&License Key Generator"
      TabPicture(1)   =   "frmMain3.frx":0CE6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdViewArchive"
      Tab(1).Control(1)=   "cmbRegisteredLevel"
      Tab(1).Control(2)=   "cmbProds"
      Tab(1).Control(3)=   "frmKeyGen"
      Tab(1).Control(4)=   "cmdViewLevel"
      Tab(1).Control(5)=   "Label15"
      Tab(1).Control(6)=   "Label8"
      Tab(1).ControlCount=   7
      Begin VB.CommandButton cmdValidate 
         Caption         =   "&Validate"
         Height          =   315
         Left            =   8520
         TabIndex        =   49
         Top             =   4080
         Width           =   1005
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   825
         Left            =   8730
         Picture         =   "frmMain3.frx":0D02
         ScaleHeight     =   825
         ScaleWidth      =   825
         TabIndex        =   47
         Top             =   5820
         Width           =   825
      End
      Begin VB.CommandButton cmdViewArchive 
         Caption         =   "&View License Database"
         Height          =   315
         Left            =   -73560
         TabIndex        =   46
         ToolTipText     =   "View License Archive"
         Top             =   7425
         Width           =   2115
      End
      Begin VB.ComboBox cmbRegisteredLevel 
         Height          =   315
         Left            =   -68340
         Style           =   2  'Dropdown List
         TabIndex        =   42
         Top             =   480
         Width           =   2625
      End
      Begin VB.Frame Frame1 
         Caption         =   "Product Codes for the Selected Product"
         ForeColor       =   &H00FF0000&
         Height          =   1095
         Left            =   120
         TabIndex        =   36
         Top             =   2760
         Width           =   8295
         Begin VB.TextBox txtCode2_2 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   690
            Width           =   7425
         End
         Begin VB.TextBox txtCode1_2 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   720
            Locked          =   -1  'True
            TabIndex        =   37
            ToolTipText     =   "Use this code to set ActiveLock's SoftwareCode property within your application."
            Top             =   330
            Width           =   7425
         End
         Begin VB.Label Label14 
            Caption         =   "GCode"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label13 
            Caption         =   "VCode"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   615
         End
      End
      Begin VB.CommandButton cmdRemove 
         Caption         =   "&Remove"
         Enabled         =   0   'False
         Height          =   315
         Left            =   8520
         TabIndex        =   11
         Top             =   4500
         Width           =   1000
      End
      Begin VB.ComboBox cmbProds 
         Height          =   315
         Left            =   -73560
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   480
         Width           =   3615
      End
      Begin VB.Frame frmKeyGen 
         BorderStyle     =   0  'None
         Height          =   6495
         Left            =   -74880
         TabIndex        =   13
         Top             =   840
         Width           =   9495
         Begin VB.CheckBox chkLockIP 
            Caption         =   "Lock to IP Address"
            Height          =   195
            Left            =   1305
            TabIndex        =   61
            Top             =   3375
            Width           =   7095
         End
         Begin VB.CheckBox chkLockMotherboard 
            Caption         =   "Lock to Motherboard Serial"
            Height          =   195
            Left            =   1305
            TabIndex        =   60
            Top             =   3105
            Width           =   7095
         End
         Begin VB.CheckBox chkLockBIOS 
            Caption         =   "Lock to BIOS Version"
            Height          =   195
            Left            =   1305
            TabIndex        =   59
            Top             =   2835
            Width           =   7095
         End
         Begin VB.CheckBox chkLockWindows 
            Caption         =   "Lock to Windows Serial"
            Height          =   195
            Left            =   1305
            TabIndex        =   58
            Top             =   2565
            Width           =   7095
         End
         Begin VB.CheckBox chkLockHDfirmware 
            Caption         =   "Lock to HDD Firmware Serial"
            Height          =   195
            Left            =   1305
            TabIndex        =   57
            Top             =   2295
            Width           =   7095
         End
         Begin VB.CheckBox chkLockHD 
            Caption         =   "Lock to HDD Volume Serial"
            Height          =   195
            Left            =   1305
            TabIndex        =   56
            Top             =   2025
            Width           =   7095
         End
         Begin VB.CheckBox chkLockComputer 
            Caption         =   "Lock to Computer Name"
            Height          =   195
            Left            =   1305
            TabIndex        =   55
            Top             =   1755
            Width           =   7095
         End
         Begin VB.CheckBox chkLockMACaddress 
            Caption         =   "Lock to MAC Address"
            Height          =   195
            Left            =   1305
            TabIndex        =   54
            Top             =   1485
            Width           =   7095
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   825
            Left            =   120
            Picture         =   "frmMain3.frx":3C8A
            ScaleHeight     =   825
            ScaleWidth      =   825
            TabIndex        =   52
            Top             =   4575
            Width           =   825
         End
         Begin VB.CheckBox chkItemData 
            Caption         =   "Use Item Data for Code"
            Height          =   330
            Left            =   6540
            TabIndex        =   51
            Top             =   0
            Width           =   2760
         End
         Begin VB.CommandButton cmdCopy 
            Height          =   345
            Left            =   8520
            MaskColor       =   &H8000000F&
            Picture         =   "frmMain3.frx":6C12
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   4095
            Width           =   345
         End
         Begin VB.CommandButton cmdPaste 
            Height          =   345
            Left            =   8520
            Picture         =   "frmMain3.frx":6D5C
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   720
            Width           =   345
         End
         Begin VB.TextBox txtUser 
            Height          =   315
            Left            =   1320
            TabIndex        =   23
            Top             =   1080
            Width           =   7095
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   8040
            TabIndex        =   29
            ToolTipText     =   "Generate liberation key for the above request code (which should not be blank)."
            Top             =   6135
            Width           =   375
         End
         Begin VB.TextBox txtLibFile 
            Height          =   315
            Left            =   1320
            TabIndex        =   28
            Top             =   6135
            Width           =   6735
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8520
            TabIndex        =   30
            ToolTipText     =   "Generate liberation key for the above request code (which should not be blank)."
            Top             =   6135
            Width           =   975
         End
         Begin VB.ComboBox cmbLicType 
            Height          =   315
            ItemData        =   "frmMain3.frx":709E
            Left            =   1320
            List            =   "frmMain3.frx":70AB
            Style           =   2  'Dropdown List
            TabIndex        =   17
            Top             =   0
            Width           =   3615
         End
         Begin VB.TextBox txtDays 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            Height          =   315
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   19
            Text            =   "30"
            Top             =   375
            Width           =   1755
         End
         Begin VB.TextBox txtReqCodeIn 
            Height          =   315
            Left            =   1320
            TabIndex        =   21
            Top             =   720
            Width           =   7095
         End
         Begin VB.TextBox txtLibKey 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   2355
            Left            =   1320
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Text            =   "frmMain3.frx":70D1
            Top             =   3735
            Width           =   7095
         End
         Begin VB.CommandButton cmdKeyGen 
            Caption         =   "&Generate"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8520
            TabIndex        =   26
            ToolTipText     =   "Generate liberation key for the above request code (which should not be blank)."
            Top             =   3735
            Width           =   975
         End
         Begin MSComDlg.CommonDialog CommonDlg 
            Left            =   9000
            Top             =   720
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.Label Label18 
            Caption         =   "Note: IP address may be Dynamic!"
            Height          =   390
            Left            =   0
            TabIndex        =   62
            Top             =   3330
            Width           =   1335
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Caption         =   "Activelock V3"
            ForeColor       =   &H00FF0000&
            Height          =   165
            Left            =   0
            TabIndex        =   53
            Top             =   5415
            Width           =   1065
         End
         Begin VB.Label Label11 
            Caption         =   "User Name:"
            Height          =   252
            Left            =   0
            TabIndex        =   22
            Top             =   1111
            Width           =   1332
         End
         Begin VB.Label Label5 
            Caption         =   "Liberation &File:"
            Height          =   255
            Left            =   0
            TabIndex        =   27
            Top             =   6165
            Width           =   1335
         End
         Begin VB.Label lblExpiry 
            Caption         =   "&Expires After:"
            Height          =   255
            Left            =   0
            TabIndex        =   18
            Top             =   405
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "License &Type:"
            Height          =   255
            Left            =   0
            TabIndex        =   16
            Top             =   30
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Installation C&ode:"
            Height          =   255
            Left            =   0
            TabIndex        =   20
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label Label12 
            Caption         =   "Liberation &Key:"
            Height          =   255
            Left            =   0
            TabIndex        =   24
            Top             =   3795
            Width           =   1335
         End
         Begin VB.Label lblDays 
            Caption         =   "days"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3120
            TabIndex        =   31
            Top             =   430
            Width           =   1935
         End
      End
      Begin VB.Frame fraProdNew 
         Height          =   2295
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   9495
         Begin VB.CommandButton cmdProductsStorage 
            Caption         =   "Products st&orage ..."
            Height          =   345
            Left            =   7770
            TabIndex        =   63
            Top             =   360
            Width           =   1575
         End
         Begin VB.CommandButton cmdCopyGCode 
            Height          =   345
            Left            =   7920
            Picture         =   "frmMain3.frx":7114
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   1910
            Width           =   345
         End
         Begin VB.CommandButton cmdCopyVCode 
            Height          =   345
            Left            =   4680
            Picture         =   "frmMain3.frx":725E
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   1910
            Width           =   345
         End
         Begin VB.CommandButton cmdCodeGen 
            Caption         =   "&Generate"
            Enabled         =   0   'False
            Height          =   315
            Left            =   8400
            TabIndex        =   8
            Top             =   1320
            Width           =   1000
         End
         Begin VB.TextBox txtCode2 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   5040
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   7
            Top             =   1320
            Width           =   3225
         End
         Begin VB.TextBox txtCode1 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1320
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   6
            ToolTipText     =   "Use this code to set ActiveLock's SoftwareCode property within your application."
            Top             =   1320
            Width           =   3705
         End
         Begin VB.TextBox txtName 
            Height          =   315
            Left            =   1320
            TabIndex        =   2
            Top             =   360
            Width           =   3705
         End
         Begin VB.TextBox txtVer 
            Height          =   315
            Left            =   1320
            TabIndex        =   4
            Top             =   720
            Width           =   1545
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add To Product List"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1320
            TabIndex        =   9
            Top             =   1920
            Width           =   1845
         End
         Begin VB.Label Label10 
            Caption         =   "GCode (for the new application)"
            Height          =   255
            Left            =   5040
            TabIndex        =   34
            Top             =   1080
            Width           =   3135
         End
         Begin VB.Label Label9 
            Caption         =   "VCode (for the new application)"
            Height          =   255
            Left            =   1320
            TabIndex        =   33
            Top             =   1080
            Width           =   3615
         End
         Begin VB.Label Label4 
            Caption         =   "&Code:"
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "&Name:"
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label3 
            Caption         =   "&Version:"
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   720
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid gridProds 
         Height          =   3255
         Left            =   120
         TabIndex        =   10
         Top             =   4080
         Width           =   8265
         _ExtentX        =   14579
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         BackColorFixed  =   -2147483626
         ForeColorSel    =   -2147483643
         BackColorBkg    =   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483633
         AllowBigSelection=   0   'False
         FocusRect       =   0
         GridLines       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   $"frmMain3.frx":73A8
      End
      Begin VB.Image cmdViewLevel 
         Height          =   240
         Left            =   -65640
         MouseIcon       =   "frmMain3.frx":7468
         MousePointer    =   99  'Custom
         Picture         =   "frmMain3.frx":7772
         Top             =   480
         Width           =   240
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Activelock V3"
         ForeColor       =   &H00FF0000&
         Height          =   165
         Left            =   8610
         TabIndex        =   48
         Top             =   6660
         Width           =   1065
      End
      Begin VB.Label Label15 
         Caption         =   "Registered Level:"
         Height          =   255
         Left            =   -69600
         TabIndex        =   43
         Top             =   510
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   "&Product List:"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   3840
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "&Product:"
         Height          =   255
         Left            =   -74880
         TabIndex        =   14
         Top             =   510
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
'*   All material is the property of the contributing authors.
'*
'*   Redistribution and use in source and binary forms, with or without
'*   modification, are permitted provided that the following conditions are
'*   met:
'*
'*     [o] Redistributions of source code must retain the above copyright
'*         notice, this list of conditions and the following disclaimer.
'*
'*     [o] Redistributions in binary form must reproduce the above
'*         copyright notice, this list of conditions and the following
'*         disclaimer in the documentation and/or other materials provided
'*         with the distribution.
'*
'*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
'*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
'*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
'*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'*
'*

''
' ALUGEN Executable Main Form
'
' @author th2tran@users.sourceforge.net
' @version 3.0.0
' @date 20050421
'
'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
'
' @todo TODO1
'
'

'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' @history
' <pre>
' 04.21.05 - sentax         -Updated frmMain to new version.  Fixes bugs of overwriting codes in product.ini
'                           -frmMain update written by ialkin
' 04.21.05 - sentax         -Upgraded to version 3
' 08.15.03 - th2tran        - Created
' 09.21.03 - th2tran        - Finished first pass of implementation.
'                             ALUGEN is now able to create product codes and liberation keys.
' 10.13.03 - th2tran        - Renamed Code1 and Code2 to VCode and GCode respectively.
'                           - Added accessor keys to all controls.
'                           - Split Liberation Key into 64-byte chunks to make it look nicer.
'                             Need handle this on ActiveLock's side. i.e. ignore vbCrLf characters
'                             in the liberation key.
'                           - Don't process button click when not in the right tab.
'                             e.g. when in Key Generator tab and accessor key for cmdRemove is pressed,
'                             just ignore the event.
' 11.02.03 - th2tran        - Removed License Class from ALUGEN.  This setting was totally unnecessary.
'                           - Reworked KeyGen logic a bit, due to IALUGenerator_GenKey() change.
' 12.xx.03 - th2tran        - Users uncovered another bug in alcrypto.dll that I've yet to figure out the
'                             cause so, for now, I've put a product code validation function to let users test
'                             for valid keys before accepting.
' 25.02.04 - th2tran        - Save license key directly to a license file.  We can then give the license file
'                             to the user, who will simply place it in their application directory.
'                             The application will automatically recognize the license file the next time it starts up.
' 17.04.04 - th2tran        - Saving to license file will not be safe because we can't update the LastUsed property
'                             (only the user app can do that). So instead, we  save the liberation key to a file
'                             and provide the ability automatically register upon initialization, which is being handled
'                             in the current IActiveLock implementation.
' 28.04.04 - th2tran        - Normalize date format to YYYY/MM/DD to handle different regional settings.
'                             This fixes the "type mismatch error" encountered by IActiveLock_Register() when
'                             it tried to call CDate() on a date string of the format yyyy.mm.dd.
' 11.07.04 - th2tran        - Changed liberation file extension from .alb to .all
'                           - Added User Name text box to allow liberation key generation for lockToNone settings
'                             without requiring Installation Code to be entered.
'                           - Apparently, having the Code Generate button enabled for an added product was
'                             confusing to the users. So: Disable the Code Generate button after a product
'                             has been added to the list. i.e. only allow codeset generation once per product.
' 05.13.05 - ialkan         - In an effort to merge activelock.dll, alutil.dll and alugen.dll, all Alugenlib references
'                             in this form have been replaced with Activelock3
' </pre>

'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit
Private GeneratorInstance As IALUGenerator
Private fDisableNotifications As Boolean
Private ActiveLock As ActiveLock3.IActiveLock
Public mKeyStoreType As LicStoreType
Public mProductsStoreType As ProductsStoreType
Public mProductsStoragePath As String
Private blnIsFirstLaunch As Boolean

' Hardware keys from the Installation Code
Dim MACaddress As String, ComputerName As String
Dim VolumeSerial As String, FirmwareSerial As String
Dim WindowsSerial As String, BIOSserial As String
Dim MotherboardSerial As String, IPaddress As String
Dim systemEvent As Boolean

'Windows and System directory API
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'Declare API calls needed to read/write to INI files
Private Declare Function KRN32_GetProfileInt Lib "kernel32" Alias "GetProfileIntA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal nDefault As Long) As Long
Private Declare Function KRN32_GetProfileString Lib "kernel32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Private Declare Function KRN32_WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
Private Declare Function KRN32_GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function KRN32_GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function KRN32_WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Dim PROJECT_INI_FILENAME As String

Public Function ProfileString32(sININame As String, sSection As String, sKeyword As String, vsDefault As String) As String
'
'   This routine will get a string from an INI file. The user must pass the
'   following information:
'
'   Input Values
'       sININame   - The name of the INI file (with path if desired)
'       sSection   - The section head to search for
'       sKeyword   - The keyword within the section to return
'       vsDefault   - The default value to return if nothing is found
'
'   Output Value
'       A string with the value found (max length = 512)
'
    Dim sReturnValue As String
    Dim nReturnSize As Long
    Dim nValidSize As Long

    On Error GoTo ProfileString32_ER

    If IsEmpty(sININame) Or IsEmpty(sSection) Or IsEmpty(sKeyword) Then
        ProfileString32 = vsDefault
        Exit Function
    End If

    sReturnValue = Space$(512)
    nReturnSize = Len(sReturnValue)

    If Right$(UCase$(sININame), 7) = "WIN.INI" Then
        nValidSize = KRN32_GetProfileString(sSection, sKeyword, vsDefault, sReturnValue, nReturnSize)
    Else
        nValidSize = KRN32_GetPrivateProfileString(sSection, sKeyword, vsDefault, sReturnValue, nReturnSize, sININame)
    End If

    ProfileString32 = Left$(sReturnValue, nValidSize)
    Exit Function

ProfileString32_ER:
    ProfileString32 = vsDefault
    Exit Function
End Function

'===============================================================================
' Name: Sub AppendLockString
' Input:
'   ByRef strLock As String - The lock string to be appended to, returns as an output
'   ByVal newSubString As String - The string to be appended to the lock string if strLock is empty string
' Output:
'   Appended lock string and installation code
' Purpose: Appends the lock string to the given installation code
' Remarks: None
'===============================================================================
Private Sub AppendLockString(ByRef strLock As String, ByVal newSubString As String)
    If strLock = "" Then
        strLock = newSubString
    Else
        strLock = strLock & vbLf & newSubString
    End If
End Sub

Private Sub gridProds_CellDataChanged(Row As Long, Col As Long, OldValue As Variant, NewValue As Variant)
    Dim nRow As Integer
    nRow = Row
    On Error Resume Next
    If Col = 1 Then
        ' Will never happen - Do nothing
    ElseIf Col = 2 Then
        '@todo Anything needs be done here? Update the INI file with the changed Code1 may be?
    ElseIf Col = 3 Then
        '@todo Anything needs be done here? Update the INI file with the changed Code2 may be?
    End If
    Debug.Print "gridProds_CellDataChanged"
End Sub


Private Function ReconstructedInstallationCode() As String
    Dim strLock As String, strReq As String
    Dim noKey As String
    noKey = Chr(110) & Chr(111) & Chr(107) & Chr(101) & Chr(121)
    
    If Me.chkLockMACaddress.Value = vbChecked Then
        AppendLockString strLock, MACaddress
    Else
        AppendLockString strLock, noKey
    End If
    If Me.chkLockComputer.Value = vbChecked Then
        AppendLockString strLock, ComputerName
    Else
        AppendLockString strLock, noKey
    End If
    If Me.chkLockHD.Value = vbChecked Then
        AppendLockString strLock, VolumeSerial
    Else
        AppendLockString strLock, noKey
    End If
    If Me.chkLockHDfirmware.Value = vbChecked Then
        AppendLockString strLock, FirmwareSerial
    Else
        AppendLockString strLock, noKey
    End If
    If Me.chkLockWindows.Value = vbChecked Then
        AppendLockString strLock, WindowsSerial
    Else
        AppendLockString strLock, noKey
    End If
    If Me.chkLockBIOS.Value = vbChecked Then
        AppendLockString strLock, BIOSserial
    Else
        AppendLockString strLock, noKey
    End If
    If Me.chkLockMotherboard.Value = vbChecked Then
        AppendLockString strLock, MotherboardSerial
    Else
        AppendLockString strLock, noKey
    End If
    If Me.chkLockIP.Value = vbChecked Then
        AppendLockString strLock, IPaddress
    Else
        AppendLockString strLock, noKey
    End If

    'If strLock = Nothing Then Exit Function
    If Left(strLock, 1) = vbLf Then strLock = Mid(strLock, 2)

    Dim Index As Integer, i As Integer
    Dim strInstCode As String
    strInstCode = ActiveLock3.Base64Decode(txtReqCodeIn)
    If strInstCode = "" Then Exit Function
    If Left(strInstCode, 1) = "+" Then
        strInstCode = Mid(strInstCode, 2)
    End If
    Index = 0
    i = 1
    ' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
    Do While i > 0
        i = InStr(Index + 1, strInstCode, vbLf)
        If i > 0 Then Index = i
    Loop
    ' user name starts from Index+1 to the end
    Dim user As String
    user = Mid$(strInstCode, Index + 1)
    
    ' combine with user name
    strReq = strLock & vbLf & user
    
    ' base-64 encode the request
    Dim strReq2 As String
    strReq2 = modBase64.Base64_Encode("+" & strReq)
    ReconstructedInstallationCode = strReq2

End Function

Private Sub chkLockBIOS_Click()
If systemEvent Then Exit Sub
systemEvent = True
txtReqCodeIn.Text = ReconstructedInstallationCode
systemEvent = False
End Sub

Private Sub chkLockComputer_Click()
If systemEvent Then Exit Sub
systemEvent = True
txtReqCodeIn.Text = ReconstructedInstallationCode
systemEvent = False
End Sub

Private Sub chkLockHD_Click()
If systemEvent Then Exit Sub
systemEvent = True
txtReqCodeIn.Text = ReconstructedInstallationCode
systemEvent = False
End Sub


Private Sub chkLockHDfirmware_Click()
If systemEvent Then Exit Sub
systemEvent = True
txtReqCodeIn.Text = ReconstructedInstallationCode
systemEvent = False
End Sub


Private Sub chkLockIP_Click()
If systemEvent Then Exit Sub
systemEvent = True
txtReqCodeIn.Text = ReconstructedInstallationCode
systemEvent = False
End Sub

Private Sub chkLockMACaddress_Click()
If systemEvent Then Exit Sub
systemEvent = True
txtReqCodeIn.Text = ReconstructedInstallationCode
systemEvent = False
End Sub

Private Sub chkLockMotherboard_Click()
If systemEvent Then Exit Sub
systemEvent = True
txtReqCodeIn.Text = ReconstructedInstallationCode
systemEvent = False
End Sub

Private Sub chkLockWindows_Click()
If systemEvent Then Exit Sub
systemEvent = True
txtReqCodeIn.Text = ReconstructedInstallationCode
systemEvent = False
End Sub

Private Sub cmbLicType_Click()
    ' enable the days edit box
    If cmbLicType = "Periodic" Or cmbLicType = "Time Locked" Then
        txtDays.Locked = False
        txtDays.BackColor = &H80000005
        txtDays.ForeColor = vbBlack
    Else
        txtDays.Locked = True
        txtDays.BackColor = &H8000000F
        txtDays.ForeColor = &H8000000F
    End If
    If cmbLicType = "Time Locked" Then
        lblExpiry = "&Expires on Date:"
        txtDays = ActiveLockDateFormat(Now() + 30)
        lblDays = "YYYY/MM/DD"
    Else
        lblExpiry = "&Expires after:"
        txtDays = "30"
        lblDays = "Day(s)"
    End If
End Sub

''
' Normalizes date format to yyyy/mm/dd
'
Private Function ActiveLockDateFormat(dt As Date) As String
    'ActiveLockDateFormat = Year(dt) & "/" & Month(dt) & "/" & Day(dt)
    ActiveLockDateFormat = Format(UTC(dt), "YYYY/MM/DD")
End Function

Private Sub cmbProds_Click()
    UpdateKeyGenButtonStatus
End Sub

Private Sub cmdAdd_Click()
    If SSTab1.Tab <> 0 Then Exit Sub ' our tab not active - do nothing
    AddRow txtName, txtVer, txtCode1, txtCode2
    cmdAdd.Enabled = False ' disallow repeated clicking of Add button
    gridProds_EnterCell
    txtName.SetFocus
    
    'begin -- Added code by Ismail
    txtCode1.Text = ""
    txtCode2.Text = ""
    cmdValidate.Enabled = True
    'end -- Added code by Ismail

End Sub
Private Sub cmdBrowse_Click()
    On Error GoTo ErrHandler
    Dim arrProdVer() As String
    arrProdVer = Split(cmbProds, "-")
    Dim strName As String
    strName = Trim$(arrProdVer(0))
    With CommonDlg
        .InitDir = Dir(txtLibFile.Text)
        .Filter = "ALL Files (*.ALL)|*.ALL"
        .Flags = cdlOFNExplorer Or cdlOFNShareAware Or cdlOFNNoChangeDir
        .CancelError = True
        .FileName = strName
        .ShowOpen
        txtLibFile.Text = .FileName
    End With

    Exit Sub
ErrHandler:
    ' no change
End Sub

Private Sub cmdCodeGen_Click()
    If SSTab1.Tab <> 0 Then Exit Sub ' our tab not active - do nothing

    Screen.MousePointer = vbHourglass
    fDisableNotifications = True
    txtCode1 = ""
    txtCode2 = ""
    fDisableNotifications = False
    Enabled = False
    DoEvents
    On Error GoTo Done
    Dim KEY As RSAKey
    Dim progress As ProgressType
    ' generate the key
    If modALUGEN.rsa_generate(KEY, 1024, AddressOf CryptoProgressUpdate, VarPtr(progress)) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
    'if modALUGEN.rsa_generate2(Key, 1024)= RETVAL_ON_ERROR Then
    '    Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    'End If
    ' extract private and public key blobs
    Dim strBlob As String
    Dim blobLen As Long
    If rsa_public_key_blob(KEY, vbNullString, blobLen) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
    If blobLen > 0 Then
        strBlob = String(blobLen, 0)
        If rsa_public_key_blob(KEY, strBlob, blobLen) = RETVAL_ON_ERROR Then
            Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
        End If
        Debug.Print "Public blob: " & strBlob
        txtCode1 = strBlob
    End If

    If modALUGEN.rsa_private_key_blob(KEY, vbNullString, blobLen) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
    If blobLen > 0 Then
        strBlob = String(blobLen, 0)
        If modALUGEN.rsa_private_key_blob(KEY, strBlob, blobLen) = RETVAL_ON_ERROR Then
            Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
        End If
        Debug.Print "Private blob: " & strBlob
        txtCode2 = strBlob
    End If
    ' done with the key - throw it away
    If modALUGEN.rsa_freekey(KEY) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If

    ' Test generated key for correctness by recreating it from the blobs
    ' Note:
    ' ====
    ' Due to an outstanding bug in ALCrypto.dll, sometimes this step will crash the app because
    ' the generated keyset is bad.
    ' The work-around for the time being is to keep regenerating the keyset until eventually
    ' you'll get a valid keyset that no longer crashes.
    Dim strdata$: strdata = "This is a test string to be encrypted."
    If modALUGEN.rsa_createkey(txtCode1, Len(txtCode1), txtCode2, Len(txtCode2), KEY) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
    ' It worked! We're all set to go.
    If modALUGEN.rsa_freekey(KEY) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
Done:
    Screen.MousePointer = vbDefault
    Enabled = True
End Sub

Private Sub cmdCopy_Click()
Clipboard.Clear
Clipboard.SetText txtLibKey.Text
End Sub
Private Sub cmdCopyGCode_Click()
Clipboard.Clear
Clipboard.SetText txtCode2_2.Text
End Sub

Private Sub cmdCopyVCode_Click()
Clipboard.Clear
Clipboard.SetText txtCode1_2.Text
End Sub

Private Sub cmdDecryptInstCode_Click()
Dim strInstCode As String
Dim userName As String
Dim Index As Integer, i As Integer
Dim a() As String

If txtReqCodeIn.Text = "" Then Exit Sub
strInstCode = ActiveLock3.Base64Decode(txtReqCodeIn.Text)
Index = 0
i = 1
' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
Do While i > 0
    i = InStr(Index + 1, strInstCode, vbLf)
    If i > 0 Then Index = i
Loop
' user name starts from Index+1 to the end
userName = Mid$(strInstCode, Index + 1)
a = Split(strInstCode, vbLf)
Dim aString As String
For i = LBound(a) To UBound(a) - 1
    aString = aString & a(i) & vbCrLf
Next i

MsgBox "User Name: " & userName & vbCrLf & vbCrLf & _
    "Since the lockType is not known at this point," & vbCrLf & _
    "the decrypted code below might include:" & vbCrLf & _
    "    the HD Volume Serial Number," & vbCrLf & _
    "    and/or the HD Firmware (Manufacturer) Serial Number," & vbCrLf & _
    "    and/or the Windows Serial Number," & vbCrLf & _
    "    and/or the Computer Name," & vbCrLf & _
    "    and/or the MAC Address:" & vbCrLf & vbCrLf & aString, vbInformation, "Decrypted Installation Code"

End Sub

' Generate liberation key
Private Sub cmdKeyGen_Click()
    If SSTab1.Tab <> 1 Then Exit Sub ' our tab not active - do nothing
    
    If chkLockMACaddress.Value = vbUnchecked And chkLockComputer.Value = vbUnchecked And _
        chkLockHD.Value = vbUnchecked And chkLockHDfirmware.Value = vbUnchecked And _
        chkLockWindows.Value = vbUnchecked And chkLockBIOS.Value = vbUnchecked And _
        chkLockMotherboard.Value = vbUnchecked And chkLockIP.Value = vbUnchecked Then
        MsgBox "Warning: You did not select any hardware keys to lock the license.", vbExclamation
    End If

    ' get product and version
    Screen.MousePointer = vbHourglass
    UpdateStatus "Generating license key..."
    On Error GoTo ErrHandler
    Dim arrProdVer() As String
    arrProdVer = Split(cmbProds, "-")
    Dim strName$, strVer$
    strName = Trim$(arrProdVer(0))
    strVer = Trim$(arrProdVer(1))
    With ActiveLock
        .SoftwareName = strName
        .SoftwareVersion = strVer
    End With

    Dim varLicType As ActiveLock3.ALLicType

    If cmbLicType = "Periodic" Then
        varLicType = allicPeriodic
    ElseIf cmbLicType = "Permanent" Then
        varLicType = allicPermanent
    ElseIf cmbLicType = "Time Locked" Then
        varLicType = allicTimeLocked
    Else
        varLicType = allicNone
    End If

    Dim strExpire$
    strExpire = GetExpirationDate()
    Dim Lic As ActiveLock3.ProductLicense
'    Set Lic = ActiveLock3.CreateProductLicense(strName, "Code", strVer, alfSingle, varLicType, "Licensee", strExpire, "LicKey", "RegDate", "Hash1")
    ' Create a product license object without the product key or license key
    
    Dim strRegDate As String
    'strRegDate = ActiveLockDateFormat(Now())
    strRegDate = Format(UTC(Now()), "YYYY/MM/DD")
     '<Modified by: kirtaph at 2/16/2006-13.03.47 on machine: KIRTAPHPC>
    Set Lic = ActiveLock3.CreateProductLicense(strName, strVer, "", alfSingle, varLicType, "", IIf(chkItemData.Value = vbUnchecked, cmbRegisteredLevel.List(cmbRegisteredLevel.ListIndex), cmbRegisteredLevel.ItemData(cmbRegisteredLevel.ListIndex)), strExpire, , strRegDate)
     '</Modified by: kirtaph at 2/16/2006-13.03.47 on machine: KIRTAPHPC>
    
    Dim strLibKey As String
    ' Pass it to IALUGenerator to generate the key
     '<Modified by: kirtaph at 2/16/2006-13.04.30 on machine: KIRTAPHPC>
    strLibKey = GeneratorInstance.GenKey(Lic, txtReqCodeIn, IIf(chkItemData.Value = vbUnchecked, cmbRegisteredLevel.List(cmbRegisteredLevel.ListIndex), cmbRegisteredLevel.ItemData(cmbRegisteredLevel.ListIndex)))
     '</Modified by: kirtaph at 2/16/2006-13.04.30 on machine: KIRTAPHPC>
    txtLibKey.Text = Make64ByteChunks(strLibKey & "aLck" & txtReqCodeIn.Text)
    txtLibFile.Text = App.path & "\" & strName & ".all"
    
    Screen.MousePointer = vbNormal
    UpdateStatus "Ready"
    
    If MsgBox("Would you like to save the new license in the License Database?", vbYesNo + vbQuestion) = vbYes Then
        Load frmAlugenDatabase
        Dim lockTypesString As String
        lockTypesString = ""
        If chkLockMACaddress.Value = vbChecked Then
            lockTypesString = lockTypesString & "MAC Address"
        End If
        If chkLockComputer.Value = vbChecked Then
            If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
            lockTypesString = lockTypesString & "Computer Name"
        End If
        If chkLockHD.Value = vbChecked Then
            If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
            lockTypesString = lockTypesString & "HDD Volume Serial"
        End If
        If chkLockHDfirmware.Value = vbChecked Then
            If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
            lockTypesString = lockTypesString & "HDD Firmware Serial"
        End If
        If chkLockWindows.Value = vbChecked Then
            If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
            lockTypesString = lockTypesString & "Windows Serial"
        End If
        If chkLockBIOS.Value = vbChecked Then
            If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
            lockTypesString = lockTypesString & "BIOS Serial"
        End If
        If chkLockMotherboard.Value = vbChecked Then
            If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
            lockTypesString = lockTypesString & "Motherboard Serial"
        End If
        If chkLockIP.Value = vbChecked Then
            If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
            lockTypesString = lockTypesString & "IP Address"
        End If
        Call frmAlugenDatabase.ArchiveLicense(cmbProds.Text, _
            txtUser.Text, _
            strRegDate, strExpire, cmbLicType, _
            cmbRegisteredLevel.Text, txtReqCodeIn.Text, txtLibKey.Text, lockTypesString)
        Unload frmAlugenDatabase
        Set frmAlugenDatabase = Nothing
    End If

    Exit Sub
ErrHandler:
    UpdateStatus "Error: " + Err.Description
    Screen.MousePointer = vbNormal
End Sub


''
' Breaks a long string into chunks of 64-byte lines.
'
Private Function Make64ByteChunks(strdata As String) As String
    Dim i As Long
    Dim Count As Long
    Count = Len(strdata)
    Dim sResult As String
    sResult = Left$(strdata, 64)
    i = 65
    While i <= Count
        sResult = sResult & vbCrLf & Mid$(strdata, i, 64)
        i = i + 64
    Wend
    Make64ByteChunks = sResult
End Function

Private Function GetExpirationDate() As String
    If cmbLicType = "Time Locked" Then
        GetExpirationDate = txtDays
    Else
        GetExpirationDate = ActiveLockDateFormat(Now + CInt(txtDays))
    End If
End Function

Private Sub cmdPaste_Click()
txtReqCodeIn.Text = Clipboard.GetText
End Sub

Private Sub cmdProductsStorage_Click()
  On Error GoTo cmdProductsStorage_Click_Error
      
  Dim mPSform As frmProductsStorage
  Set mPSform = New frmProductsStorage
  mPSform.Show 1, Me

  On Error GoTo 0
  Exit Sub

cmdProductsStorage_Click_Error:

  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdProductsStorage_Click of Form frmMain"
End Sub

Private Sub cmdRemove_Click()
    If SSTab1.Tab <> 0 Then Exit Sub ' our tab not active - do nothing
    Dim strName As String
    Dim SelStart%, SelEnd%
    Dim fEnableAdd As Boolean
    fEnableAdd = False

    With gridProds
        If .Row < .RowSel Then
            SelStart = .Row
            SelEnd = .RowSel
        Else
            SelStart = .RowSel
            SelEnd = .Row
        End If
        Dim i%
        For i = SelEnd To SelStart Step -1
            strName = .TextMatrix(i, 0)
            If txtName = strName Then
                fEnableAdd = True
            End If
            RemoveRow i
        Next
        .RowSel = .Row ' negate current selections
    End With

    ' Enable Add button if we're removing the variable currently being edited.
    If fEnableAdd Then
        cmdAdd.Enabled = True
    End If
    If gridProds.Rows = 1 Then
        ' no (selectable) rows left (just the header)
        cmdRemove.Enabled = False
    End If
    gridProds_EnterCell
    
    'begin -- Added code by Ismail
    txtCode1.Text = ""
    txtCode2.Text = ""
    cmdValidate.Enabled = True
    cmbProds.RemoveItem i
    'end -- Added code by Ismail

End Sub

Private Sub cmdSave_Click()
    UpdateStatus "Saving liberation key to file..."
    ' save the liberation key
    SaveLiberationKey txtLibKey, txtLibFile
    UpdateStatus "Liberation key saved."
End Sub

Private Sub SaveLiberationKey(ByVal sLibKey As String, ByVal sFileName As String)
    Dim hFile As Long
    hFile = FreeFile
    Open sFileName For Output As #hFile
        Print #hFile, sLibKey
    Close #hFile
End Sub

Private Sub cmdValidate_Click()
    Screen.MousePointer = vbHourglass
    If txtCode1_2 = "" And txtCode2_2 = "" Then
        UpdateStatus "GCode and VCode fields are blank.  Nothing to validate."
        Exit Sub ' nothing to validate
    End If
    ' Validate to keyset to make sure it's valid.
    UpdateStatus "Validating keyset..."
    Dim KEY As RSAKey
    Dim strdata$: strdata = "This is a test string to be signed."
    Dim strSig$
    If modALUGEN.rsa_createkey(txtCode1_2, Len(txtCode1_2), txtCode2_2, Len(txtCode2_2), KEY) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
    ' sign it
    strSig = RSASign(txtCode1_2, txtCode2_2, strdata)
    Dim rc&
    rc = RSAVerify(txtCode1_2, strdata, strSig)
    If rc = 0 Then
        UpdateStatus gridProds.TextMatrix(gridProds.Row, 0) & " (" + gridProds.TextMatrix(gridProds.Row, 1) + ") validated successfully."
    Else
        UpdateStatus gridProds.TextMatrix(gridProds.Row, 0) & " (" + gridProds.TextMatrix(gridProds.Row, 1) + ") GCode-VCode mismatch!"
    End If
    ' It worked! We're all set to go.
    If modALUGEN.rsa_freekey(KEY) = RETVAL_ON_ERROR Then
        Err.Raise ActiveLockErrCodeConstants.alerrRSAError, ACTIVELOCKSTRING, STRRSAERROR
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub UpdateStatus(Msg As String)
    sbStatus.Panels(1) = Msg
End Sub

Private Sub cmdViewArchive_Click()
Const LICENSES_FILE = "authorizations.txt"
If fileExist(App.path & "\" & LICENSES_FILE) = False Then
    MsgBox "License archive file " & """" & LICENSES_FILE & """" & " does not exist.", vbInformation
    Exit Sub
End If
Load frmAlugenDatabase
frmAlugenDatabase.ShowArchive
frmAlugenDatabase.Show
End Sub

Private Sub cmdViewLevel_Click()
    With frmLevelManager
        .Show vbModal
        cmbRegisteredLevel.Clear
        LoadComboBox strRegisteredLevelDBName, cmbRegisteredLevel, True
        cmbRegisteredLevel.ListIndex = 0
    End With

    Unload frmLevelManager
End Sub

Private Sub Form_Initialize()
  'on first start.. initialize variables
  blnIsFirstLaunch = True
  'with default values
  mKeyStoreType = alsFile 'alsRegistry or alsFile
  mProductsStoreType = alsINIFile 'alsINIFile - for ini file, alsXMLFile for xml file, alsMDBFile for MDB file
  Select Case mProductsStoreType
    Case alsINIFile
      mProductsStoragePath = App.path & "\licenses.ini"
    Case alsXMLFile
      mProductsStoragePath = App.path & "\licenses.xml" 'for XML store
    Case alsMDBFile
      mProductsStoragePath = App.path & "\licenses.mdb" 'for MDB store
    'Case alsMSSQL '-not implemented yet
      'mProductsStoragePath =
  End Select
End Sub

Public Sub Form_Load()
    
    '<Added by: kirtaph at: 2/16/2006-13.05.40 on machine: KIRTAPHPC>
    strRegisteredLevelDBName = AddBackSlash(App.path) & "RegisteredLevelDB.dat"
    '</Added by: kirtaph at: 2/16/2006-13.05.40 on machine: KIRTAPHPC>
    
    ' Check the existence of necessary files to run this application
    Call CheckForResources("Alcrypto3.dll", "comdlg32.ocx", "msflxgrd.ocx", "comctl32.ocx", "tabctl32.ocx")

    '<Modified by: kirtaph at 2/16/2006-13.06.25 on machine: KIRTAPHPC>
    If Not fileExist(strRegisteredLevelDBName) Then
    
        With cmbRegisteredLevel
            .Clear
            .AddItem "Limited A"
            .AddItem "Limited B"
            .AddItem "Limited C"
            .AddItem "Limited D"
            .AddItem "Limited E"
            .AddItem "No Print/Save"
            .AddItem "Educational A"
            .AddItem "Educational B"
            .AddItem "Educational C"
            .AddItem "Educational D"
            .AddItem "Educational E"
            .AddItem "Level 1"
            .AddItem "Level 2"
            .AddItem "Level 3"
            .AddItem "Level 4"
            .AddItem "Light Version"
            .AddItem "Pro Version"
            .AddItem "Enterprise Version"
            .AddItem "Demo Only"
            .AddItem "Full Version-Europe"
            .AddItem "Full Version-Africa"
            .AddItem "Full Version-Asia"
            .AddItem "Full Version-USA"
            .AddItem "Full Version-International"
            .ListIndex = 0
            SaveComboBox strRegisteredLevelDBName, cmbRegisteredLevel, True
        End With

    Else
        LoadComboBox strRegisteredLevelDBName, cmbRegisteredLevel, True
        cmbRegisteredLevel.ListIndex = 0
    End If
    '</Modified by: kirtaph at 2/16/2006-13.06.25 on machine: KIRTAPHPC>
    
    'load form settings
    LoadFormSetting
    
    'initialize ActiveLock instances
    InitActiveLock
    
    ' Initialize GUI
    txtLibKey = ""
    InitUI
    If cmbProds.ListCount > 0 Then
      cmbProds.ListIndex = 0
    End If
    
    'Assume originally the app is not using LockNone as the LockType
    txtUser.Enabled = False
    txtUser.Locked = True
    txtUser.BackColor = vbButtonFace

    Me.Caption = "ALUGEN - ActiveLock3 Universal GENerator - v" & App.Major & "." & App.Minor & "." & App.Revision
    Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

Private Sub LoadFormSetting()
  'Read the program INI file to retrieve control settings
  On Error GoTo LoadFormSetting_Error
  
  If Not blnIsFirstLaunch Then Exit Sub
  
  PROJECT_INI_FILENAME = WinDir() & "\Alugen3.ini"
  On Error Resume Next
  SSTab1.Tab = Val(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "TabNumber", 0))
  cmbProds.ListIndex = Val(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cmbProds", 0))
  cmbLicType.ListIndex = Val(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cmbLicType", 1))
  cmbRegisteredLevel.ListIndex = Val(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cmbRegisteredLevel", 0))
  chkItemData.Value = Val(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkItemData", False))
  mKeyStoreType = Val(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "KeyStoreType", 1))
  mProductsStoreType = Val(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoreType", 0))
  mProductsStoragePath = ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoragePath", App.path & "\licenses.ini")
  If Not fileExist(mProductsStoragePath) And mProductsStoreType = alsMDBFile Then
    mProductsStoreType = alsINIFile
    mProductsStoragePath = App.path & "\licenses.ini"
  End If

  blnIsFirstLaunch = False
  
  On Error GoTo 0
  Exit Sub

LoadFormSetting_Error:

  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure LoadFormSetting of Form frmMain"
End Sub


Private Sub InitActiveLock()
  On Error GoTo InitForm_Error
  ' Initialize AL
  Set ActiveLock = ActiveLock3.NewInstance()
  ActiveLock.KeyStoreType = mKeyStoreType
  ActiveLock.Init
  
  ' Initialize Generator
  Set GeneratorInstance = ActiveLock3.GeneratorInstance(mProductsStoreType) 'alsINIFile - for ini file, alsXMLFile for xml file, alsMDBFile for MDB file
  GeneratorInstance.StoragePath = mProductsStoragePath
   
  On Error GoTo 0
  Exit Sub

InitForm_Error:

  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure InitForm of Form frmMain"
End Sub

Function CheckForResources(ParamArray MyArray()) As Boolean
'MyArray is a list of things to check
'These can be DLLs or OCXs

'Files, by default, are searched for in the Windows System Directory
'Exceptions;
'   Begins with a # means it should be in the same directory with the executable
'   Contains the full path (anything with a "\")

'Typical names would be "#aaa.dll", "mydll.dll", "myocx.ocx", "comdlg32.ocx", "mscomctl.ocx", "msflxgrd.ocx"

'If the file has no extension, we;
'     assume it's a DLL, and if it still can't be found
'     assume it's an OCX

On Error GoTo checkForResourcesError
Dim foundIt As Boolean
Dim Y As Variant
Dim i As Integer, j As Integer
Dim s As String, systemDir As String, pathName As String

WhereIsDLL ("") 'initialize

systemDir = WinSysDir 'Get the Windows system directory
For Each Y In MyArray
    foundIt = False
    s = CStr(Y)
    
    If Left$(s, 1) = "#" Then
        pathName = App.path
        s = Mid$(s, 2)
    ElseIf inString(s, "\") Then
        j = InStrRev(s, "\")
        pathName = Left$(s, j - 1)
        s = Mid$(s, j + 1)
    Else
        pathName = systemDir
    End If
    
    If inString(s, ".") Then
        If fileExist(pathName & "\" & s) Then foundIt = True
    ElseIf fileExist(pathName & "\" & s & ".DLL") Then
        foundIt = True
    ElseIf fileExist(pathName & "\" & s & ".OCX") Then
        foundIt = True
        s = s & ".OCX" 'this will make the softlocx check easier
    End If
    
    If Not foundIt Then
        MsgBox s & " could not be found in " & pathName & vbCrLf & _
        App.Title & " cannot run without this library file!" & vbCrLf & vbCrLf & "Exiting!", vbCritical, "Missing Resource"
        End
    End If
Next Y

CheckForResources = True
Exit Function

checkForResourcesError:
    MsgBox "CheckForResources error", vbCritical, "Error"
    End   'an error kills the program
End Function
Function WhereIsDLL(ByVal T As String) As String
'Places where programs look for DLLs
'   1 directory containing the EXE
'   2 current directory
'   3 32 bit system directory   possibly \Windows\system32
'   4 16 bit system directory   possibly \Windows\system
'   5 windows directory         possibly \Windows
'   6 path

'The current directory may be changed in the course of the program
'but current directory -- when the program starts -- is what matters
'so a call should be made to this function early on to "lock" the paths.

'Add a call at the beginning of checkForResources

Static a As Variant
Dim s As String, D As String
Dim EnvString As String, Indx As Integer  ' Declare variables.
Dim i As Integer

On Error Resume Next
i = UBound(a)
If i = 0 Then
    s = App.path & ";" & CurDir & ";"
    
    D = WinSysDir
    s = s & D & ";"
    
    If Right$(D, 2) = "32" Then   'I'm guessing at the name of the 16 bit windows directory (assuming it exists)
        i = Len(D)
        s = s & Left$(D, i - 2) & ";"
    End If
    
    s = s & WinDir & ";"
    Indx = 1   ' Initialize index to 1.
    Do
        EnvString = Environ(Indx)   ' Get environment variable.
        If StrComp(Left(EnvString, 5), "PATH=", vbTextCompare) = 0 Then ' Check PATH entry.
            s = s & Mid$(EnvString, 6)
            Exit Do
        End If
        Indx = Indx + 1
    Loop Until EnvString = ""
    a = Split(s, ";")
End If

T = Trim(T)
If T = "" Then Exit Function
If Not inString(Right$(T, 4), ".") Then T = T & ".DLL"   'default extension
For i = 0 To UBound(a)
    If fileExist(a(i) & "\" & T) Then
        WhereIsDLL = a(i)
        Exit Function
    End If
Next i

End Function
Function fileExist(ByVal TestFileName As String) As Boolean
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

Function inString(ByVal X As String, ParamArray MyArray()) As Boolean
'Do ANY of a group of sub-strings appear in within the first string?
'Case doesn't count and we don't care WHERE or WHICH
Dim Y As Variant    'member of array that holds all arguments except the first
    For Each Y In MyArray
    If InStr(1, X, Y, 1) > 0 Then 'the "ones" make the comparison case-insensitive
        inString = True
        Exit Function
    End If
    Next Y
End Function

Public Function LooseSpace(invoer$) As String
'This routine terminates a string if it detects char 0.

Dim P As Long

P = InStr(invoer$, Chr(0))
If P <> 0 Then
    LooseSpace$ = Left$(invoer$, P - 1)
    Exit Function
End If
LooseSpace$ = invoer$

End Function

Private Sub Form_Unload(Cancel As Integer)
  SaveFormSettings
End Sub

Private Sub SaveFormSettings()
  'save form settings
  On Error GoTo SaveFormSettings_Error
  Dim mnReturnValue As Long
  mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "TabNumber", CStr(SSTab1.Tab))
  mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cmbProds", CStr(cmbProds.ListIndex))
  mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cmbLicType", CStr(cmbLicType.ListIndex))
  mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cmbRegisteredLevel", CStr(cmbRegisteredLevel.ListIndex))
  mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkItemData", CStr(chkItemData.Value))
  mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "KeyStoreType", CStr(mKeyStoreType))
  mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoreType", CStr(mProductsStoreType))
  mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoragePath", CStr(mProductsStoragePath))

 On Error GoTo 0
 Exit Sub

SaveFormSettings_Error:

  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure SaveFormSettings of Form frmMain"
End Sub

Public Function SetProfileString32(sININame As String, sSection As String, sKeyword As String, vsEntry As String) As Long
'
'   This routine will write a string to an INI file. The user must pass the
'   following information:
'
'   Input Values
'       sININame   - The name of the INI file (with path if desired)
'       sSection   - The section head to search for
'       sKeyword   - The keyword within the section to return
'       vsEntry     - The value to write for the keyword
'
'   Output Value
'       non-zero if successful, zero if failure.
'
    On Error GoTo SetProfileString32_ER

If IsEmpty(sININame) Or IsEmpty(sSection) Or IsEmpty(sKeyword) Then
    SetProfileString32 = 0
    Exit Function
End If

If Right$(UCase$(sININame), 7) = "WIN.INI" Then
    SetProfileString32 = KRN32_WriteProfileString(sSection, sKeyword, vsEntry)
Else
    SetProfileString32 = KRN32_WritePrivateProfileString(sSection, sKeyword, vsEntry, sININame)
End If
Exit Function

SetProfileString32_ER:
    SetProfileString32 = 0
    Exit Function

End Function

Private Sub gridProds_Click()
    gridProds_EnterCell
    
    'begin -- added code by Ismail
    txtCode1.Text = ""
    txtCode2.Text = ""
    cmdValidate.Enabled = True
    'end -- added code by Ismail
    
End Sub

Private Sub gridProds_EnterCell()
Debug.Print "EnterCell"
    If gridProds.Row < 1 Then
        ' heading row is selected.  Make it look like it's not.
        gridProds.BackColorSel = gridProds.BackColorFixed
        gridProds.ForeColorSel = gridProds.ForeColor
        Exit Sub ' skip the header row
    Else
        gridProds.BackColorSel = &H8000000D ' navy blue
        gridProds.ForeColorSel = &H80000005 ' white
    End If
    fDisableNotifications = True
    txtName = gridProds.TextMatrix(gridProds.Row, 0)
    txtVer = gridProds.TextMatrix(gridProds.Row, 1)
    
    'begin -- Modified code by Ismail
    txtCode1_2 = gridProds.TextMatrix(gridProds.Row, 2)
    txtCode2_2 = gridProds.TextMatrix(gridProds.Row, 3)
    'end -- Modified code by Ismail

    cmdRemove.Enabled = True
    fDisableNotifications = False
End Sub

''
' Add a Product Row to the GUI.
' If fUpdateStore is True, then product info is also saved to the store.
'
Private Sub AddRow(Name As String, Ver As String, Code1 As String, Code2 As String, Optional fUpdateStore As Boolean = True)
    ' Update the view
    With gridProds
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = Name
        .TextMatrix(.Rows - 1, 1) = Ver
        .TextMatrix(.Rows - 1, 2) = Code1
        .TextMatrix(.Rows - 1, 3) = Code2
        .Row = .Rows - 1
        .Col = 0
        .ColSel = 3
        ' .SetFocus ' show selected row
    End With
    ' Call Activelock3.IALUGenerator to add product
    Dim ProdInfo As ActiveLock3.ProductInfo
    Set ProdInfo = ActiveLock3.CreateProductInfo(Name, Ver, Code1, Code2)
    If fUpdateStore Then
        Call GeneratorInstance.SaveProduct(ProdInfo)
    End If
    cmbProds.AddItem Name & " - " & Ver
    cmdRemove.Enabled = True
End Sub

Private Sub RemoveRow(Row As Integer)
    Dim strName$, strVer$
    strName = gridProds.TextMatrix(Row, 0)
    strVer = gridProds.TextMatrix(Row, 1)
    ' Update the view
    If gridProds.Rows = 2 Then
        gridProds.Rows = 1
    Else
        gridProds.RemoveItem Row
    End If
    ' Call Activelock3.IALUGenerator to remove product
    GeneratorInstance.DeleteProduct strName, strVer
End Sub

Private Sub UpdateRow(Row As Integer, Code1 As String, Code2 As String)
    Dim strName$, strVer$
    strName = gridProds.TextMatrix(Row, 0)
    strVer = gridProds.TextMatrix(Row, 1)
    With gridProds
        .TextMatrix(Row, 2) = Code1
        .TextMatrix(Row, 3) = Code2
    End With
    ' update the store file
    GeneratorInstance.SaveProduct ActiveLock3.CreateProductInfo(strName, strVer, Code1, Code2)
End Sub

Private Sub gridProds_RowColChange()
Debug.Print "RowColChange!!"
End Sub

Private Sub txtCode1_2_GotFocus()
txtCode1_2.SelStart = 0
txtCode1_2.SelLength = Len(txtCode1_2.Text)
End Sub


Private Sub txtCode1_GotFocus()
txtCode1.SelStart = 0
txtCode1.SelLength = Len(txtCode1.Text)
End Sub

Private Sub txtCode2_2_GotFocus()
txtCode2_2.SelStart = 0
txtCode2_2.SelLength = Len(txtCode2_2.Text)
End Sub


Private Sub txtCode2_GotFocus()
txtCode2.SelStart = 0
txtCode2.SelLength = Len(txtCode2.Text)
End Sub

Private Sub txtLibKey_Change()
    cmdSave.Enabled = CBool(Len(txtLibKey) > 0)
End Sub

Private Sub txtLibKey_GotFocus()
txtLibKey.SelStart = 0          ' Start selection at beginning.
txtLibKey.SelLength = Len(txtLibKey.Text)  ' Length of text in Text1.
End Sub

Private Sub txtName_Change()
    UpdateCodeGenButtonStatus
    UpdateAddButtonStatus
End Sub

Private Sub txtReqCodeIn_Change()
'<Modified by: kirtaph at 2/16/2006-13.08.40 on machine: KIRTAPHPC>
If Len(txtReqCodeIn) > 0 Then

    If systemEvent Then Exit Sub
    
    UpdateKeyGenButtonStatus
    If fDisableNotifications Then Exit Sub
    
    fDisableNotifications = True
    txtUser = GetUserFromInstallCode(txtReqCodeIn)
    fDisableNotifications = False
    
    systemEvent = True
    If chkLockMACaddress.Enabled = True Then chkLockMACaddress.Value = vbChecked
    If chkLockComputer.Enabled = True Then chkLockComputer.Value = vbChecked
    If chkLockHD.Enabled = True Then chkLockHD.Value = vbChecked
    If chkLockHDfirmware.Enabled = True Then chkLockHDfirmware.Value = vbChecked
    If chkLockWindows.Enabled = True Then chkLockWindows.Value = vbChecked
    If chkLockBIOS.Enabled = True Then chkLockBIOS.Value = vbChecked
    If chkLockMotherboard.Enabled = True Then chkLockMotherboard.Value = vbChecked
    If chkLockIP.Enabled = True Then chkLockIP.Value = vbChecked
    systemEvent = False

Else
    fDisableNotifications = True
    chkLockComputer.Enabled = True
    chkLockComputer.Caption = "Lock to Computer Name"
    chkLockHD.Enabled = True
    chkLockHD.Caption = "Lock to HDD Volume Serial"
    chkLockHDfirmware.Enabled = True
    chkLockHDfirmware.Caption = "Lock to HDD Firmware Serial"
    chkLockMACaddress.Enabled = True
    chkLockMACaddress.Caption = "Lock to MAC Address"
    chkLockWindows.Enabled = True
    chkLockWindows.Caption = "Lock to Windows Serial"
    chkLockBIOS.Enabled = True
    chkLockBIOS.Caption = "Lock to BIOS Serial"
    chkLockMotherboard.Enabled = True
    chkLockMotherboard.Caption = "Lock to Motherboard Serial"
    chkLockIP.Enabled = True
    chkLockIP.Caption = "Lock to IP Address"
    txtUser.Text = ""
    fDisableNotifications = False
End If

'</Modified by: kirtaph at 2/16/2006-13.08.40 on machine: KIRTAPHPC>

End Sub

Private Sub txtUser_Change()
    If fDisableNotifications Then Exit Sub
    fDisableNotifications = True
    txtReqCodeIn = ActiveLock.installationCode(txtUser)
    fDisableNotifications = False
End Sub


Private Sub UpdateKeyGenButtonStatus()
    If txtReqCodeIn = "" Then
        cmdKeyGen.Enabled = False
    Else
        If cmbProds <> "" Then
            cmdKeyGen.Enabled = True
        End If
    End If
End Sub

Private Sub txtVer_Change()
    ' New product - will be processed by Add command
    UpdateCodeGenButtonStatus
    UpdateAddButtonStatus
End Sub

Private Sub txtCode1_Change()
    UpdateAddButtonStatus
    If fDisableNotifications Then Exit Sub

'begin -- modified by Ismail
'    With gridProds
'        If .Rows = 1 Then Exit Sub
'        ' update current row
'        UpdateRow .Row, txtCode1, txtCode2
'    End With
'end -- modified by Ismail

End Sub

Private Sub txtCode2_Change()
    If fDisableNotifications Then Exit Sub
    
'begin -- modified by Ismail
'    With gridProds
'        If .Rows = 1 Then Exit Sub
'        ' update current row
'        UpdateRow .Row, txtCode1, txtCode2
'    End With
'end -- modified by Ismail

End Sub

Private Sub UpdateCodeGenButtonStatus()
    If txtName = "" Or txtVer = "" Then
        cmdCodeGen.Enabled = False
    ElseIf CheckDuplicate(txtName, txtVer) Then
        cmdCodeGen.Enabled = False
    Else
        cmdCodeGen.Enabled = True
    End If
End Sub

Private Sub UpdateAddButtonStatus()
    If txtName = "" Or txtVer = "" Or txtCode1 = "" Then
        cmdAdd.Enabled = False
    ElseIf CheckDuplicate(txtName, txtVer) Then
        cmdAdd.Enabled = False
    Else
        cmdAdd.Enabled = True
    End If
End Sub

''
' Validate and enable Add/Change button as appropriate
'
Private Function CheckDuplicate(Name As String, Ver As String) As Boolean
    CheckDuplicate = False
    Dim i%
    For i = 0 To gridProds.Rows - 1
        If gridProds.TextMatrix(i, 0) = Name Then
            If gridProds.TextMatrix(i, 1) = Ver Then
                CheckDuplicate = True
                Exit Function
            End If
        End If
    Next
End Function

Private Function CheckDuplicateVCode(VCode As String) As Boolean
    CheckDuplicateVCode = False
    Dim i%
    For i = 0 To gridProds.Rows - 1
        If gridProds.TextMatrix(i, 2) = VCode Then
            CheckDuplicateVCode = True
        End If
    Next
End Function


''
' Initialize the GUI with the proper grid headings and alignments
'
Private Sub InitUI()
    ' Init Default license class
    cmbLicType = "Periodic"
    With gridProds
        .Clear
        .Rows = 1
        .FormatString = "Name                             |Version          | VCode                                          |GCode                                                                                  "
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
    End With
    cmbProds.Clear
    
    ' Populate Product List on Product Code Generator tab
    ' and Key Gen tab with product info from licenses.ini
    Dim arrProdInfos() As ProductInfo
    arrProdInfos = GeneratorInstance.RetrieveProducts()
    If IsArrayEmpty(arrProdInfos) Then Exit Sub

    Dim i As Integer
    For i = LBound(arrProdInfos) To UBound(arrProdInfos)
        PopulateUI arrProdInfos(i)
    Next
    gridProds_Click
End Sub
Private Function IsArrayEmpty(arrVar As Variant) As Boolean
    IsArrayEmpty = True
    On Error GoTo ErrHandler
    Dim lb As Long
    lb = UBound(arrVar, 1) ' this will raise an error if the array is empty
    IsArrayEmpty = False ' If we managed to get to here, then it's not empty
    Exit Function
ErrHandler:
    ' return false
End Function

Private Sub PopulateUI(ProdInfo As ActiveLock3.ProductInfo)
    With ProdInfo
        AddRow .Name, .Version, .VCode, .GCode, False
    End With
End Sub

''
' Retrieves lock string and user info from the request string
'
Private Function GetUserFromInstallCode(ByVal strInstCode As String) As String
Dim a() As String
Dim Index As Integer, i As Integer
Dim aString As String
Dim usedLockNone As Boolean
Dim noKey As String
noKey = Chr(110) & Chr(111) & Chr(107) & Chr(101) & Chr(121)

If strInstCode = "" Then Exit Function
strInstCode = ActiveLock3.Base64Decode(strInstCode)
If Left(strInstCode, 1) = "+" Then
    strInstCode = Mid(strInstCode, 2)
    usedLockNone = True
End If
Index = 0
i = 1
' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
Do While i > 0
    i = InStr(Index + 1, strInstCode, vbLf)
    If i > 0 Then Index = i
Loop
' user name starts from Index+1 to the end
GetUserFromInstallCode = Mid$(strInstCode, Index + 1)

systemEvent = True
chkLockMACaddress.Value = vbUnchecked
chkLockMACaddress.Enabled = True
chkLockComputer.Value = vbUnchecked
chkLockComputer.Enabled = True
chkLockHD.Value = vbUnchecked
chkLockHD.Enabled = True
chkLockHDfirmware.Value = vbUnchecked
chkLockHDfirmware.Enabled = True
chkLockWindows.Value = vbUnchecked
chkLockWindows.Enabled = True
' Added in v3.3
chkLockBIOS.Value = vbUnchecked
chkLockBIOS.Enabled = True
chkLockMotherboard.Value = vbUnchecked
chkLockMotherboard.Enabled = True
chkLockIP.Value = vbUnchecked
chkLockIP.Enabled = True

a = Split(strInstCode, vbLf)
If usedLockNone = True Then
    For i = LBound(a) To UBound(a) - 1
        aString = a(i)  'aString & A(i) & vbCrLf
        If i = LBound(a) Then
            MACaddress = aString
            chkLockMACaddress.Caption = "Lock to MAC Address:                           " & MACaddress
        ElseIf i = LBound(a) + 1 Then
            ComputerName = aString
            chkLockComputer.Caption = "Lock to Computer Name:                       " & ComputerName
        ElseIf i = LBound(a) + 2 Then
            VolumeSerial = aString
            chkLockHD.Caption = "Lock to HDD Volume Serial Number:     " & VolumeSerial
        ElseIf i = LBound(a) + 3 Then
            FirmwareSerial = aString
            chkLockHDfirmware.Caption = "Lock to HDD Firmware Serial Number:  " & FirmwareSerial
        ElseIf i = LBound(a) + 4 Then
            WindowsSerial = aString
            chkLockWindows.Caption = "Lock to Windows Serial Number:           " & WindowsSerial
        ElseIf i = LBound(a) + 5 Then
            BIOSserial = aString
            chkLockBIOS.Caption = "Lock to BIOS Serial Number:                 " & BIOSserial
        ElseIf i = LBound(a) + 6 Then
            MotherboardSerial = aString
            chkLockMotherboard.Caption = "Lock to Motherboard Serial Number:     " & MotherboardSerial
        ElseIf i = LBound(a) + 7 Then
            IPaddress = aString
            chkLockIP.Caption = "Lock to IP Address:                               " & IPaddress
        End If
    Next i
Else '"+" was not used, therefore one or more lockTypes were specified in the application
    chkLockMACaddress.Enabled = False
    chkLockHD.Enabled = False
    chkLockHDfirmware.Enabled = False
    chkLockWindows.Enabled = False
    chkLockComputer.Enabled = False
    chkLockBIOS.Enabled = False
    chkLockMotherboard.Enabled = False
    chkLockIP.Enabled = False

    For i = LBound(a) To UBound(a) - 1
        aString = a(i)  'aString & A(i) & vbCrLf
        
        If i = LBound(a) And aString <> noKey Then
            MACaddress = aString
            chkLockMACaddress.Caption = "Lock to MAC Address:                           " & MACaddress
            chkLockMACaddress.Value = vbChecked
        ElseIf i = (LBound(a) + 1) And aString <> noKey Then
            ComputerName = aString
            chkLockComputer.Caption = "Lock to Computer Name:                       " & ComputerName
            chkLockComputer.Value = vbChecked
        ElseIf i = (LBound(a) + 2) And aString <> noKey Then
            VolumeSerial = aString
            chkLockHD.Caption = "Lock to HDD Volume Serial Number:     " & VolumeSerial
            chkLockHD.Value = vbChecked
        ElseIf i = (LBound(a) + 3) And aString <> noKey Then
            FirmwareSerial = aString
            chkLockHDfirmware.Caption = "Lock to HDD Firmware Serial Number:  " & FirmwareSerial
            chkLockHDfirmware.Value = vbChecked
        ElseIf i = (LBound(a) + 4) And aString <> noKey Then
            WindowsSerial = aString
            chkLockWindows.Caption = "Lock to Windows Serial Number:           " & WindowsSerial
            chkLockWindows.Value = vbChecked
        ElseIf i = (LBound(a) + 5) And aString <> noKey Then
            BIOSserial = aString
            chkLockBIOS.Caption = "Lock to BIOS Serial Number:                 " & BIOSserial
            chkLockBIOS.Value = vbChecked
        ElseIf i = (LBound(a) + 6) And aString <> noKey Then
            MotherboardSerial = aString
            chkLockMotherboard.Caption = "Lock to Motherboard Serial Number:     " & MotherboardSerial
            chkLockMotherboard.Value = vbChecked
        ElseIf i = (LBound(a) + 7) And aString <> noKey Then
            IPaddress = aString
            chkLockIP.Caption = "Lock to IP Address:                               " & IPaddress
            chkLockIP.Value = vbChecked
        End If
        
    Next i
End If

If VolumeSerial = "Not Available" Or VolumeSerial = "0000-0000" Then
    chkLockHD.Enabled = False
    chkLockHD.Value = vbUnchecked
End If
If MACaddress = "00 00 00 00 00 00" Or MACaddress = "00-00-00-00-00-00" Or MACaddress = "" Or MACaddress = "Not Available" Then
    chkLockMACaddress.Enabled = False
    chkLockMACaddress.Value = vbUnchecked
End If
If FirmwareSerial = "Not Available" Then
    chkLockHDfirmware.Enabled = False
    chkLockHDfirmware.Value = vbUnchecked
End If
If BIOSserial = "Not Available" Then
    chkLockBIOS.Enabled = False
    chkLockBIOS.Value = vbUnchecked
End If
If MotherboardSerial = "Not Available" Then
    chkLockMotherboard.Enabled = False
    chkLockMotherboard.Value = vbUnchecked
End If
If IPaddress = "Not Available" Then
    chkLockIP.Enabled = False
    chkLockIP.Value = vbUnchecked
End If

systemEvent = False

End Function

