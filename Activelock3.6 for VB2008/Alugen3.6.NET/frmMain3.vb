'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2008 The ActiveLock Software Group (ASG)
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
Option Strict On
Option Explicit On
Imports System
Imports System.IO
Imports System.Text
Imports System.Windows.Forms
Imports ActiveLock3_6NET
Imports System.Security.Cryptography
Imports System.text.regularexpressions

Friend Class frmMain
    Inherits System.Windows.Forms.Form
    ' Precompiled base64 regular expression
    Public Shared ReadOnly ExpBase64 As Regex = New Regex("^[a-zA-Z0-9+/=]{4,}$", RegexOptions.Compiled)
    Dim cboProducts_Array() As String


#Region "Private variables "
    'listview sorter class
    Private lvwColumnSorter As ListViewColumnSorter
    Private MyGlobals As New Globals
    'ActiveLock objects
    Public GeneratorInstance As _IALUGenerator
    Public ActiveLock As _IActiveLock
    Public fDisableNotifications As Boolean

    Public mKeyStoreType As IActiveLock.LicStoreType
    Public mProductsStoreType As IActiveLock.ProductsStoreType
    Public mProductsStoragePath As String
    Private blnIsFirstLaunch As Boolean

    ' Hardware keys from the Installation Code
    Private MACaddress, ComputerName As String
    Private VolumeSerial, FirmwareSerial As String
    Private WindowsSerial, BIOSserial As String
    Private MotherboardSerial, IPaddress As String
    Private ExternalIP, Fingerprint As String
    Private Memory, CPUID As String
    Private BaseboardID, VideoID As String
    Private systemEvent As Boolean

    Private PROJECT_INI_FILENAME As String

    Public Shared resxList As New Collections.Specialized.ListDictionary
    Friend WithEvents cmdStartWizard As System.Windows.Forms.Button
    Friend WithEvents optStrength4 As System.Windows.Forms.RadioButton
    Friend WithEvents lblLockMacAddress As System.Windows.Forms.Label
    Friend WithEvents lblLockMotherboard As System.Windows.Forms.Label
    Friend WithEvents lblLockBIOS As System.Windows.Forms.Label
    Friend WithEvents lblLockWindows As System.Windows.Forms.Label
    Friend WithEvents lblLockHDfirmware As System.Windows.Forms.Label
    Friend WithEvents lblLockHD As System.Windows.Forms.Label
    Friend WithEvents lblLockComputer As System.Windows.Forms.Label
    Friend WithEvents lblLockIP As System.Windows.Forms.Label
    Friend WithEvents lblLockCPUID As System.Windows.Forms.Label
    Friend WithEvents chkLockCPUID As System.Windows.Forms.CheckBox
    Friend WithEvents lblLockMemory As System.Windows.Forms.Label
    Friend WithEvents chkLockMemory As System.Windows.Forms.CheckBox
    Friend WithEvents lblLockFingerprint As System.Windows.Forms.Label
    Friend WithEvents chkLockFingerprint As System.Windows.Forms.CheckBox
    Friend WithEvents lblLockExternalIP As System.Windows.Forms.Label
    Friend WithEvents chkLockExternalIP As System.Windows.Forms.CheckBox
    Friend WithEvents lblLockVideoID As System.Windows.Forms.Label
    Friend WithEvents chkLockVideoID As System.Windows.Forms.CheckBox
    Friend WithEvents lblLockBaseboardID As System.Windows.Forms.Label
    Friend WithEvents chkLockBaseboardID As System.Windows.Forms.CheckBox
    Friend WithEvents lblKeyStrength As System.Windows.Forms.Label

    Private printPreviewDialog1 As New PrintPreviewDialog

#End Region

#Region "Windows Form Designer generated code "
    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'on first start.. initialize variables
        blnIsFirstLaunch = True
        'with default values
        mKeyStoreType = IActiveLock.LicStoreType.alsFile 'alsRegistry or alsFile
        mProductsStoreType = IActiveLock.ProductsStoreType.alsINIFile 'alsINIFile - for ini file, alsXMLFile for xml file, alsMDBFile for MDB file
        Select Case mProductsStoreType
            Case IActiveLock.ProductsStoreType.alsINIFile
                mProductsStoragePath = modALUGEN.AppPath & "\licenses.ini"
            Case IActiveLock.ProductsStoreType.alsXMLFile
                mProductsStoragePath = modALUGEN.AppPath & "\licenses.xml" 'for XML store
            Case IActiveLock.ProductsStoreType.alsMDBFile
                mProductsStoragePath = modALUGEN.AppPath & "\licenses.mdb" 'for MDB store
                'Case alsMSSQL '-not implemented yet
                'mProductsStoragePath =
        End Select

        'load resources
        LoadResources()

        ' Create an instance of a ListView column sorter and assign it 
        ' to the ListView control.
        lvwColumnSorter = New ListViewColumnSorter
        Me.lstvwProducts.ListViewItemSorter = lvwColumnSorter

    End Sub
    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
            If Not components Is Nothing Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(Disposing)
    End Sub
    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer
    Friend ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmdCopyGCode As System.Windows.Forms.Button
    Friend WithEvents cmdCopyVCode As System.Windows.Forms.Button
    Friend WithEvents cmdCodeGen As System.Windows.Forms.Button
    Friend WithEvents txtName As System.Windows.Forms.TextBox
    Friend WithEvents txtVer As System.Windows.Forms.TextBox
    Friend WithEvents cmdAdd As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents fraProdNew As System.Windows.Forms.GroupBox
    Friend WithEvents cmdRemove As System.Windows.Forms.Button
    Friend WithEvents cmdValidate As System.Windows.Forms.Button
    Friend WithEvents _SSTab1_TabPage0 As System.Windows.Forms.TabPage
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents chkLockIP As System.Windows.Forms.CheckBox
    Friend WithEvents chkLockMotherboard As System.Windows.Forms.CheckBox
    Friend WithEvents chkLockBIOS As System.Windows.Forms.CheckBox
    Friend WithEvents chkLockWindows As System.Windows.Forms.CheckBox
    Friend WithEvents chkLockHDfirmware As System.Windows.Forms.CheckBox
    Friend WithEvents chkLockHD As System.Windows.Forms.CheckBox
    Friend WithEvents chkLockComputer As System.Windows.Forms.CheckBox
    Friend WithEvents chkLockMACaddress As System.Windows.Forms.CheckBox
    Friend WithEvents chkItemData As System.Windows.Forms.CheckBox
    Friend WithEvents cmdCopy As System.Windows.Forms.Button
    Friend WithEvents cmdPaste As System.Windows.Forms.Button
    Friend WithEvents txtUser As System.Windows.Forms.TextBox
    Friend WithEvents cmdBrowse As System.Windows.Forms.Button
    Friend WithEvents cmdSave As System.Windows.Forms.Button
    Friend WithEvents txtDays As System.Windows.Forms.TextBox
    Friend WithEvents cmdKeyGen As System.Windows.Forms.Button
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents lblExpiry As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents lblDays As System.Windows.Forms.Label
    Friend WithEvents frmKeyGen As System.Windows.Forms.Panel
    Friend WithEvents cmdViewArchive As System.Windows.Forms.Button
    Friend WithEvents _SSTab1_TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents SSTab1 As System.Windows.Forms.TabControl
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents lstvwProducts As System.Windows.Forms.ListView
    Friend WithEvents sbStatus As System.Windows.Forms.StatusBar
    Friend WithEvents mainStatusBarPanel As System.Windows.Forms.StatusBarPanel
    Friend WithEvents saveDlg As System.Windows.Forms.SaveFileDialog
    Friend WithEvents cmdViewLevel As System.Windows.Forms.Button
    Friend WithEvents grpCodes As System.Windows.Forms.GroupBox
    Friend WithEvents grpProductsList As System.Windows.Forms.GroupBox
    Friend WithEvents picALBanner As System.Windows.Forms.PictureBox
    Friend WithEvents picALBanner2 As System.Windows.Forms.PictureBox
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Friend WithEvents lblGCode As System.Windows.Forms.Label
    Friend WithEvents lblVCode As System.Windows.Forms.Label
    Friend WithEvents lblLicenseKey As System.Windows.Forms.Label
    Friend WithEvents dtpExpireDate As DateTimePicker
    Friend WithEvents txtVCode As System.Windows.Forms.TextBox
    Friend WithEvents txtGCode As System.Windows.Forms.TextBox
    Friend WithEvents txtInstallCode As System.Windows.Forms.TextBox
    Friend WithEvents txtLicenseKey As System.Windows.Forms.TextBox
    Friend WithEvents txtLicenseFile As System.Windows.Forms.TextBox
    Friend WithEvents cboProducts As System.Windows.Forms.ComboBox
    Friend WithEvents lnkActivelockSoftwareGroup As System.Windows.Forms.LinkLabel
    Friend WithEvents cmdPrintLicenseKey As System.Windows.Forms.Button
    Friend WithEvents cboRegisteredLevel As System.Windows.Forms.ComboBox
    Friend WithEvents cboLicType As System.Windows.Forms.ComboBox
    Friend WithEvents cmdEmailLicenseKey As System.Windows.Forms.Button
    Friend WithEvents cmdProductsStorage As System.Windows.Forms.Button
    Friend WithEvents chkNetworkedLicense As System.Windows.Forms.CheckBox
    Friend WithEvents txtMaxCount As System.Windows.Forms.TextBox
    Friend WithEvents lblConcurrentUsers As System.Windows.Forms.Label
    Friend WithEvents optStrength0 As System.Windows.Forms.RadioButton
    Friend WithEvents optStrength1 As System.Windows.Forms.RadioButton
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents optStrength2 As System.Windows.Forms.RadioButton
    Friend WithEvents optStrength5 As System.Windows.Forms.RadioButton
    Friend WithEvents optStrength3 As System.Windows.Forms.RadioButton
    Friend WithEvents cmdValidate2 As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdCheckAll As System.Windows.Forms.Button
    Friend WithEvents cmdUncheckAll As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.txtVCode = New System.Windows.Forms.TextBox
        Me.cmdBrowse = New System.Windows.Forms.Button
        Me.cmdSave = New System.Windows.Forms.Button
        Me.cmdKeyGen = New System.Windows.Forms.Button
        Me.cmdViewArchive = New System.Windows.Forms.Button
        Me.lstvwProducts = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.picALBanner2 = New System.Windows.Forms.PictureBox
        Me.cmdCopyVCode = New System.Windows.Forms.Button
        Me.cmdCopyGCode = New System.Windows.Forms.Button
        Me.cmdCodeGen = New System.Windows.Forms.Button
        Me.txtName = New System.Windows.Forms.TextBox
        Me.txtVer = New System.Windows.Forms.TextBox
        Me.cmdAdd = New System.Windows.Forms.Button
        Me.cmdValidate = New System.Windows.Forms.Button
        Me.cmdRemove = New System.Windows.Forms.Button
        Me.picALBanner = New System.Windows.Forms.PictureBox
        Me.cmdCopy = New System.Windows.Forms.Button
        Me.cmdPaste = New System.Windows.Forms.Button
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.txtLicenseFile = New System.Windows.Forms.TextBox
        Me.cboLicType = New System.Windows.Forms.ComboBox
        Me.txtInstallCode = New System.Windows.Forms.TextBox
        Me.txtLicenseKey = New System.Windows.Forms.TextBox
        Me.cboRegisteredLevel = New System.Windows.Forms.ComboBox
        Me.cboProducts = New System.Windows.Forms.ComboBox
        Me.cmdViewLevel = New System.Windows.Forms.Button
        Me.cmdPrintLicenseKey = New System.Windows.Forms.Button
        Me.cmdEmailLicenseKey = New System.Windows.Forms.Button
        Me.cmdProductsStorage = New System.Windows.Forms.Button
        Me.txtMaxCount = New System.Windows.Forms.TextBox
        Me.cmdValidate2 = New System.Windows.Forms.Button
        Me.SSTab1 = New System.Windows.Forms.TabControl
        Me._SSTab1_TabPage0 = New System.Windows.Forms.TabPage
        Me.grpProductsList = New System.Windows.Forms.GroupBox
        Me.fraProdNew = New System.Windows.Forms.GroupBox
        Me.optStrength4 = New System.Windows.Forms.RadioButton
        Me.cmdStartWizard = New System.Windows.Forms.Button
        Me.optStrength3 = New System.Windows.Forms.RadioButton
        Me.optStrength5 = New System.Windows.Forms.RadioButton
        Me.optStrength2 = New System.Windows.Forms.RadioButton
        Me.Label1 = New System.Windows.Forms.Label
        Me.optStrength1 = New System.Windows.Forms.RadioButton
        Me.optStrength0 = New System.Windows.Forms.RadioButton
        Me.grpCodes = New System.Windows.Forms.GroupBox
        Me.txtGCode = New System.Windows.Forms.TextBox
        Me.lblGCode = New System.Windows.Forms.Label
        Me.lblVCode = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me._SSTab1_TabPage1 = New System.Windows.Forms.TabPage
        Me.frmKeyGen = New System.Windows.Forms.Panel
        Me.lblKeyStrength = New System.Windows.Forms.Label
        Me.lblLockVideoID = New System.Windows.Forms.Label
        Me.chkLockVideoID = New System.Windows.Forms.CheckBox
        Me.lblLockBaseboardID = New System.Windows.Forms.Label
        Me.chkLockBaseboardID = New System.Windows.Forms.CheckBox
        Me.lblLockCPUID = New System.Windows.Forms.Label
        Me.chkLockCPUID = New System.Windows.Forms.CheckBox
        Me.lblLockMemory = New System.Windows.Forms.Label
        Me.chkLockMemory = New System.Windows.Forms.CheckBox
        Me.lblLockFingerprint = New System.Windows.Forms.Label
        Me.chkLockFingerprint = New System.Windows.Forms.CheckBox
        Me.lblLockExternalIP = New System.Windows.Forms.Label
        Me.chkLockExternalIP = New System.Windows.Forms.CheckBox
        Me.lblLockIP = New System.Windows.Forms.Label
        Me.lblLockMotherboard = New System.Windows.Forms.Label
        Me.lblLockBIOS = New System.Windows.Forms.Label
        Me.lblLockWindows = New System.Windows.Forms.Label
        Me.lblLockHDfirmware = New System.Windows.Forms.Label
        Me.lblLockHD = New System.Windows.Forms.Label
        Me.lblLockComputer = New System.Windows.Forms.Label
        Me.lblLockMacAddress = New System.Windows.Forms.Label
        Me.cmdUncheckAll = New System.Windows.Forms.Button
        Me.cmdCheckAll = New System.Windows.Forms.Button
        Me.lblConcurrentUsers = New System.Windows.Forms.Label
        Me.chkNetworkedLicense = New System.Windows.Forms.CheckBox
        Me.dtpExpireDate = New System.Windows.Forms.DateTimePicker
        Me.chkLockIP = New System.Windows.Forms.CheckBox
        Me.chkLockMotherboard = New System.Windows.Forms.CheckBox
        Me.chkLockBIOS = New System.Windows.Forms.CheckBox
        Me.chkLockWindows = New System.Windows.Forms.CheckBox
        Me.chkLockHDfirmware = New System.Windows.Forms.CheckBox
        Me.chkLockHD = New System.Windows.Forms.CheckBox
        Me.chkLockComputer = New System.Windows.Forms.CheckBox
        Me.chkLockMACaddress = New System.Windows.Forms.CheckBox
        Me.chkItemData = New System.Windows.Forms.CheckBox
        Me.txtDays = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.lblExpiry = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.lblLicenseKey = New System.Windows.Forms.Label
        Me.lblDays = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label15 = New System.Windows.Forms.Label
        Me.sbStatus = New System.Windows.Forms.StatusBar
        Me.mainStatusBarPanel = New System.Windows.Forms.StatusBarPanel
        Me.saveDlg = New System.Windows.Forms.SaveFileDialog
        Me.lnkActivelockSoftwareGroup = New System.Windows.Forms.LinkLabel
        CType(Me.picALBanner2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picALBanner, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SSTab1.SuspendLayout()
        Me._SSTab1_TabPage0.SuspendLayout()
        Me.grpProductsList.SuspendLayout()
        Me.fraProdNew.SuspendLayout()
        Me.grpCodes.SuspendLayout()
        Me._SSTab1_TabPage1.SuspendLayout()
        Me.frmKeyGen.SuspendLayout()
        CType(Me.mainStatusBarPanel, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtVCode
        '
        Me.txtVCode.AcceptsReturn = True
        Me.txtVCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVCode.BackColor = System.Drawing.SystemColors.Control
        Me.txtVCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVCode.Location = New System.Drawing.Point(84, 20)
        Me.txtVCode.MaxLength = 0
        Me.txtVCode.Multiline = True
        Me.txtVCode.Name = "txtVCode"
        Me.txtVCode.ReadOnly = True
        Me.txtVCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVCode.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtVCode.Size = New System.Drawing.Size(565, 52)
        Me.txtVCode.TabIndex = 2
        Me.ToolTip1.SetToolTip(Me.txtVCode, "Use this code to set ActiveLock's SoftwareCode property within your application.")
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdBrowse.BackColor = System.Drawing.SystemColors.Control
        Me.cmdBrowse.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdBrowse.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdBrowse.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdBrowse.Location = New System.Drawing.Point(576, 602)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdBrowse.Size = New System.Drawing.Size(21, 22)
        Me.cmdBrowse.TabIndex = 32
        Me.ToolTip1.SetToolTip(Me.cmdBrowse, "Browse for license file.")
        Me.cmdBrowse.UseVisualStyleBackColor = False
        '
        'cmdSave
        '
        Me.cmdSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSave.BackColor = System.Drawing.SystemColors.Control
        Me.cmdSave.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdSave.Enabled = False
        Me.cmdSave.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdSave.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdSave.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdSave.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdSave.Location = New System.Drawing.Point(598, 602)
        Me.cmdSave.Name = "cmdSave"
        Me.cmdSave.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdSave.Size = New System.Drawing.Size(65, 22)
        Me.cmdSave.TabIndex = 33
        Me.cmdSave.Text = "&Save"
        Me.cmdSave.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdSave, "Save into file generated license key for the above request code (which should not" & _
                " be blank).")
        Me.cmdSave.UseVisualStyleBackColor = False
        '
        'cmdKeyGen
        '
        Me.cmdKeyGen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdKeyGen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdKeyGen.Enabled = False
        Me.cmdKeyGen.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdKeyGen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdKeyGen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdKeyGen.ImageAlign = System.Drawing.ContentAlignment.BottomLeft
        Me.cmdKeyGen.Location = New System.Drawing.Point(1, 415)
        Me.cmdKeyGen.Name = "cmdKeyGen"
        Me.cmdKeyGen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdKeyGen.Size = New System.Drawing.Size(80, 24)
        Me.cmdKeyGen.TabIndex = 26
        Me.cmdKeyGen.Text = "&Generate"
        Me.cmdKeyGen.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdKeyGen, "Generate license key for the above request code (which should not be blank).")
        Me.cmdKeyGen.UseVisualStyleBackColor = False
        '
        'cmdViewArchive
        '
        Me.cmdViewArchive.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdViewArchive.BackColor = System.Drawing.SystemColors.Control
        Me.cmdViewArchive.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdViewArchive.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdViewArchive.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdViewArchive.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdViewArchive.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdViewArchive.Location = New System.Drawing.Point(3, 551)
        Me.cmdViewArchive.Name = "cmdViewArchive"
        Me.cmdViewArchive.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdViewArchive.Size = New System.Drawing.Size(79, 44)
        Me.cmdViewArchive.TabIndex = 29
        Me.cmdViewArchive.Text = "&View Lic Database"
        Me.cmdViewArchive.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdViewArchive, "View License Archive")
        Me.cmdViewArchive.UseVisualStyleBackColor = False
        '
        'lstvwProducts
        '
        Me.lstvwProducts.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lstvwProducts.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
        Me.lstvwProducts.Dock = System.Windows.Forms.DockStyle.Fill
        Me.lstvwProducts.FullRowSelect = True
        Me.lstvwProducts.GridLines = True
        Me.lstvwProducts.HideSelection = False
        Me.lstvwProducts.Location = New System.Drawing.Point(3, 16)
        Me.lstvwProducts.MultiSelect = False
        Me.lstvwProducts.Name = "lstvwProducts"
        Me.lstvwProducts.Size = New System.Drawing.Size(657, 347)
        Me.lstvwProducts.Sorting = System.Windows.Forms.SortOrder.Ascending
        Me.lstvwProducts.TabIndex = 10
        Me.ToolTip1.SetToolTip(Me.lstvwProducts, "Products list")
        Me.lstvwProducts.UseCompatibleStateImageBehavior = False
        Me.lstvwProducts.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Name"
        Me.ColumnHeader1.Width = 140
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Version"
        Me.ColumnHeader2.Width = 80
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "VCode"
        Me.ColumnHeader3.Width = 120
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "GCode"
        Me.ColumnHeader4.Width = 120
        '
        'picALBanner2
        '
        Me.picALBanner2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.picALBanner2.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.picALBanner2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picALBanner2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.picALBanner2.Location = New System.Drawing.Point(585, 24)
        Me.picALBanner2.Name = "picALBanner2"
        Me.picALBanner2.Size = New System.Drawing.Size(74, 38)
        Me.picALBanner2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.picALBanner2.TabIndex = 64
        Me.picALBanner2.TabStop = False
        Me.ToolTip1.SetToolTip(Me.picALBanner2, "www.activelocksoftware.com")
        '
        'cmdCopyVCode
        '
        Me.cmdCopyVCode.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopyVCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCopyVCode.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdCopyVCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopyVCode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopyVCode.Location = New System.Drawing.Point(60, 48)
        Me.cmdCopyVCode.Name = "cmdCopyVCode"
        Me.cmdCopyVCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCopyVCode.Size = New System.Drawing.Size(23, 23)
        Me.cmdCopyVCode.TabIndex = 1
        Me.cmdCopyVCode.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdCopyVCode, "Copy VCode into clipboard")
        Me.cmdCopyVCode.UseVisualStyleBackColor = False
        '
        'cmdCopyGCode
        '
        Me.cmdCopyGCode.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopyGCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCopyGCode.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdCopyGCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopyGCode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopyGCode.Location = New System.Drawing.Point(60, 100)
        Me.cmdCopyGCode.Name = "cmdCopyGCode"
        Me.cmdCopyGCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCopyGCode.Size = New System.Drawing.Size(23, 23)
        Me.cmdCopyGCode.TabIndex = 4
        Me.cmdCopyGCode.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdCopyGCode, "Copy GCode into clipboard")
        Me.cmdCopyGCode.UseVisualStyleBackColor = False
        '
        'cmdCodeGen
        '
        Me.cmdCodeGen.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdCodeGen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCodeGen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCodeGen.Enabled = False
        Me.cmdCodeGen.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdCodeGen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCodeGen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCodeGen.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCodeGen.Location = New System.Drawing.Point(289, 36)
        Me.cmdCodeGen.Name = "cmdCodeGen"
        Me.cmdCodeGen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCodeGen.Size = New System.Drawing.Size(78, 23)
        Me.cmdCodeGen.TabIndex = 4
        Me.cmdCodeGen.Text = "&Generate"
        Me.cmdCodeGen.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdCodeGen, "Generate new codes")
        Me.cmdCodeGen.UseVisualStyleBackColor = False
        '
        'txtName
        '
        Me.txtName.AcceptsReturn = True
        Me.txtName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtName.BackColor = System.Drawing.SystemColors.Window
        Me.txtName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtName.Location = New System.Drawing.Point(86, 14)
        Me.txtName.MaxLength = 0
        Me.txtName.Name = "txtName"
        Me.txtName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtName.Size = New System.Drawing.Size(280, 20)
        Me.txtName.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.txtName, "Product Name")
        '
        'txtVer
        '
        Me.txtVer.AcceptsReturn = True
        Me.txtVer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtVer.BackColor = System.Drawing.SystemColors.Window
        Me.txtVer.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVer.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVer.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVer.Location = New System.Drawing.Point(86, 38)
        Me.txtVer.MaxLength = 0
        Me.txtVer.Name = "txtVer"
        Me.txtVer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVer.Size = New System.Drawing.Size(199, 20)
        Me.txtVer.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.txtVer, "Product Version")
        '
        'cmdAdd
        '
        Me.cmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
        Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdAdd.Enabled = False
        Me.cmdAdd.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdAdd.Location = New System.Drawing.Point(7, 228)
        Me.cmdAdd.Name = "cmdAdd"
        Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdAdd.Size = New System.Drawing.Size(128, 23)
        Me.cmdAdd.TabIndex = 7
        Me.cmdAdd.Text = "&Add to product list"
        Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdAdd, "Add to product list")
        Me.cmdAdd.UseVisualStyleBackColor = False
        '
        'cmdValidate
        '
        Me.cmdValidate.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdValidate.BackColor = System.Drawing.SystemColors.Control
        Me.cmdValidate.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdValidate.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdValidate.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdValidate.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdValidate.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdValidate.Location = New System.Drawing.Point(369, 36)
        Me.cmdValidate.Name = "cmdValidate"
        Me.cmdValidate.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdValidate.Size = New System.Drawing.Size(78, 23)
        Me.cmdValidate.TabIndex = 5
        Me.cmdValidate.Text = "&Validate"
        Me.cmdValidate.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdValidate, "Validate product codes")
        Me.cmdValidate.UseVisualStyleBackColor = False
        '
        'cmdRemove
        '
        Me.cmdRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdRemove.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRemove.Enabled = False
        Me.cmdRemove.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdRemove.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRemove.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRemove.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdRemove.Location = New System.Drawing.Point(139, 228)
        Me.cmdRemove.Name = "cmdRemove"
        Me.cmdRemove.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRemove.Size = New System.Drawing.Size(154, 23)
        Me.cmdRemove.TabIndex = 8
        Me.cmdRemove.Text = "&Remove from product list"
        Me.cmdRemove.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdRemove, "Remove from product list")
        Me.cmdRemove.UseVisualStyleBackColor = False
        '
        'picALBanner
        '
        Me.picALBanner.BackColor = System.Drawing.SystemColors.ActiveCaptionText
        Me.picALBanner.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picALBanner.Cursor = System.Windows.Forms.Cursors.Hand
        Me.picALBanner.Location = New System.Drawing.Point(4, 144)
        Me.picALBanner.Name = "picALBanner"
        Me.picALBanner.Size = New System.Drawing.Size(74, 38)
        Me.picALBanner.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.picALBanner.TabIndex = 63
        Me.picALBanner.TabStop = False
        Me.ToolTip1.SetToolTip(Me.picALBanner, "www.activelocksoftware.com")
        '
        'cmdCopy
        '
        Me.cmdCopy.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopy.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCopy.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdCopy.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopy.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopy.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdCopy.Location = New System.Drawing.Point(1, 441)
        Me.cmdCopy.Name = "cmdCopy"
        Me.cmdCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCopy.Size = New System.Drawing.Size(80, 24)
        Me.cmdCopy.TabIndex = 28
        Me.cmdCopy.Text = "Copy  key"
        Me.cmdCopy.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdCopy, "Copy License Key into clipboard")
        Me.cmdCopy.UseVisualStyleBackColor = False
        '
        'cmdPaste
        '
        Me.cmdPaste.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdPaste.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPaste.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPaste.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdPaste.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPaste.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPaste.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPaste.Location = New System.Drawing.Point(579, 97)
        Me.cmdPaste.Name = "cmdPaste"
        Me.cmdPaste.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPaste.Size = New System.Drawing.Size(86, 24)
        Me.cmdPaste.TabIndex = 14
        Me.cmdPaste.Text = "Paste code"
        Me.cmdPaste.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdPaste, "Paste installation code from clipboard")
        Me.cmdPaste.UseVisualStyleBackColor = False
        '
        'txtUser
        '
        Me.txtUser.AcceptsReturn = True
        Me.txtUser.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtUser.BackColor = System.Drawing.SystemColors.Window
        Me.txtUser.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUser.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUser.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUser.Location = New System.Drawing.Point(88, 121)
        Me.txtUser.MaxLength = 0
        Me.txtUser.Name = "txtUser"
        Me.txtUser.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUser.Size = New System.Drawing.Size(578, 20)
        Me.txtUser.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.txtUser, "Here will apear user name based on the instalation code")
        '
        'txtLicenseFile
        '
        Me.txtLicenseFile.AcceptsReturn = True
        Me.txtLicenseFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLicenseFile.BackColor = System.Drawing.SystemColors.Window
        Me.txtLicenseFile.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLicenseFile.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLicenseFile.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLicenseFile.Location = New System.Drawing.Point(86, 603)
        Me.txtLicenseFile.MaxLength = 0
        Me.txtLicenseFile.Name = "txtLicenseFile"
        Me.txtLicenseFile.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLicenseFile.Size = New System.Drawing.Size(487, 20)
        Me.txtLicenseFile.TabIndex = 31
        Me.ToolTip1.SetToolTip(Me.txtLicenseFile, "Enter or select license file")
        '
        'cboLicType
        '
        Me.cboLicType.BackColor = System.Drawing.SystemColors.Window
        Me.cboLicType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboLicType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboLicType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboLicType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboLicType.Items.AddRange(New Object() {"Time Locked", "Periodic", "Permanent"})
        Me.cboLicType.Location = New System.Drawing.Point(88, 28)
        Me.cboLicType.Name = "cboLicType"
        Me.cboLicType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboLicType.Size = New System.Drawing.Size(234, 22)
        Me.cboLicType.TabIndex = 6
        Me.ToolTip1.SetToolTip(Me.cboLicType, "Select license type")
        '
        'txtInstallCode
        '
        Me.txtInstallCode.AcceptsReturn = True
        Me.txtInstallCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtInstallCode.BackColor = System.Drawing.SystemColors.Window
        Me.txtInstallCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtInstallCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtInstallCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtInstallCode.Location = New System.Drawing.Point(88, 99)
        Me.txtInstallCode.MaxLength = 0
        Me.txtInstallCode.Name = "txtInstallCode"
        Me.txtInstallCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtInstallCode.Size = New System.Drawing.Size(488, 20)
        Me.txtInstallCode.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.txtInstallCode, "Enter here installation code")
        '
        'txtLicenseKey
        '
        Me.txtLicenseKey.AcceptsReturn = True
        Me.txtLicenseKey.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtLicenseKey.BackColor = System.Drawing.SystemColors.Control
        Me.txtLicenseKey.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLicenseKey.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLicenseKey.ForeColor = System.Drawing.Color.Blue
        Me.txtLicenseKey.Location = New System.Drawing.Point(86, 399)
        Me.txtLicenseKey.MaxLength = 0
        Me.txtLicenseKey.Multiline = True
        Me.txtLicenseKey.Name = "txtLicenseKey"
        Me.txtLicenseKey.ReadOnly = True
        Me.txtLicenseKey.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLicenseKey.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtLicenseKey.Size = New System.Drawing.Size(580, 198)
        Me.txtLicenseKey.TabIndex = 27
        Me.txtLicenseKey.Text = "1234567890123456789012345678901234567890123456789012345678901234" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.ToolTip1.SetToolTip(Me.txtLicenseKey, "License Key")
        '
        'cboRegisteredLevel
        '
        Me.cboRegisteredLevel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cboRegisteredLevel.BackColor = System.Drawing.SystemColors.Window
        Me.cboRegisteredLevel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboRegisteredLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboRegisteredLevel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboRegisteredLevel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboRegisteredLevel.Location = New System.Drawing.Point(418, 25)
        Me.cboRegisteredLevel.Name = "cboRegisteredLevel"
        Me.cboRegisteredLevel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboRegisteredLevel.Size = New System.Drawing.Size(223, 22)
        Me.cboRegisteredLevel.TabIndex = 3
        Me.ToolTip1.SetToolTip(Me.cboRegisteredLevel, "Select desired registration level")
        '
        'cboProducts
        '
        Me.cboProducts.BackColor = System.Drawing.SystemColors.Window
        Me.cboProducts.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboProducts.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboProducts.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboProducts.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboProducts.Location = New System.Drawing.Point(88, 4)
        Me.cboProducts.Name = "cboProducts"
        Me.cboProducts.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboProducts.Size = New System.Drawing.Size(234, 22)
        Me.cboProducts.TabIndex = 1
        Me.ToolTip1.SetToolTip(Me.cboProducts, "Select product from list")
        '
        'cmdViewLevel
        '
        Me.cmdViewLevel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdViewLevel.BackColor = System.Drawing.SystemColors.Control
        Me.cmdViewLevel.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdViewLevel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdViewLevel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdViewLevel.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdViewLevel.Location = New System.Drawing.Point(643, 25)
        Me.cmdViewLevel.Name = "cmdViewLevel"
        Me.cmdViewLevel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdViewLevel.Size = New System.Drawing.Size(22, 22)
        Me.cmdViewLevel.TabIndex = 4
        Me.cmdViewLevel.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.ToolTip1.SetToolTip(Me.cmdViewLevel, "Manage registered levels")
        Me.cmdViewLevel.UseVisualStyleBackColor = False
        '
        'cmdPrintLicenseKey
        '
        Me.cmdPrintLicenseKey.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPrintLicenseKey.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPrintLicenseKey.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdPrintLicenseKey.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPrintLicenseKey.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPrintLicenseKey.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdPrintLicenseKey.Location = New System.Drawing.Point(1, 467)
        Me.cmdPrintLicenseKey.Name = "cmdPrintLicenseKey"
        Me.cmdPrintLicenseKey.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPrintLicenseKey.Size = New System.Drawing.Size(80, 24)
        Me.cmdPrintLicenseKey.TabIndex = 64
        Me.cmdPrintLicenseKey.Text = "Print  key"
        Me.cmdPrintLicenseKey.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdPrintLicenseKey, "Print License Key")
        Me.cmdPrintLicenseKey.UseVisualStyleBackColor = False
        '
        'cmdEmailLicenseKey
        '
        Me.cmdEmailLicenseKey.BackColor = System.Drawing.SystemColors.Control
        Me.cmdEmailLicenseKey.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdEmailLicenseKey.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdEmailLicenseKey.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdEmailLicenseKey.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdEmailLicenseKey.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdEmailLicenseKey.Location = New System.Drawing.Point(1, 493)
        Me.cmdEmailLicenseKey.Name = "cmdEmailLicenseKey"
        Me.cmdEmailLicenseKey.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdEmailLicenseKey.Size = New System.Drawing.Size(80, 24)
        Me.cmdEmailLicenseKey.TabIndex = 65
        Me.cmdEmailLicenseKey.Text = "Email  key"
        Me.cmdEmailLicenseKey.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdEmailLicenseKey, "Email License Key")
        Me.cmdEmailLicenseKey.UseVisualStyleBackColor = False
        '
        'cmdProductsStorage
        '
        Me.cmdProductsStorage.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdProductsStorage.BackColor = System.Drawing.SystemColors.Control
        Me.cmdProductsStorage.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdProductsStorage.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdProductsStorage.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdProductsStorage.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdProductsStorage.Location = New System.Drawing.Point(519, 228)
        Me.cmdProductsStorage.Name = "cmdProductsStorage"
        Me.cmdProductsStorage.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdProductsStorage.Size = New System.Drawing.Size(140, 23)
        Me.cmdProductsStorage.TabIndex = 65
        Me.cmdProductsStorage.Text = "Products st&orage ..."
        Me.cmdProductsStorage.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdProductsStorage, "Select product storage")
        Me.cmdProductsStorage.UseVisualStyleBackColor = False
        '
        'txtMaxCount
        '
        Me.txtMaxCount.AcceptsReturn = True
        Me.txtMaxCount.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMaxCount.BackColor = System.Drawing.SystemColors.Window
        Me.txtMaxCount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMaxCount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxCount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMaxCount.Location = New System.Drawing.Point(323, 76)
        Me.txtMaxCount.MaxLength = 2
        Me.txtMaxCount.Name = "txtMaxCount"
        Me.txtMaxCount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMaxCount.Size = New System.Drawing.Size(84, 20)
        Me.txtMaxCount.TabIndex = 68
        Me.txtMaxCount.Text = "5"
        Me.ToolTip1.SetToolTip(Me.txtMaxCount, "Enter or select license file")
        Me.txtMaxCount.Visible = False
        '
        'cmdValidate2
        '
        Me.cmdValidate2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdValidate2.BackColor = System.Drawing.SystemColors.Control
        Me.cmdValidate2.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdValidate2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdValidate2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdValidate2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdValidate2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdValidate2.Location = New System.Drawing.Point(449, 36)
        Me.cmdValidate2.Name = "cmdValidate2"
        Me.cmdValidate2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdValidate2.Size = New System.Drawing.Size(78, 23)
        Me.cmdValidate2.TabIndex = 73
        Me.cmdValidate2.Text = "&Validate2"
        Me.cmdValidate2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.ToolTip1.SetToolTip(Me.cmdValidate2, "Validate product codes")
        Me.cmdValidate2.UseVisualStyleBackColor = False
        Me.cmdValidate2.Visible = False
        '
        'SSTab1
        '
        Me.SSTab1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage0)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage1)
        Me.SSTab1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SSTab1.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTab1.Location = New System.Drawing.Point(0, 0)
        Me.SSTab1.Name = "SSTab1"
        Me.SSTab1.SelectedIndex = 1
        Me.SSTab1.Size = New System.Drawing.Size(679, 656)
        Me.SSTab1.TabIndex = 0
        '
        '_SSTab1_TabPage0
        '
        Me._SSTab1_TabPage0.Controls.Add(Me.grpProductsList)
        Me._SSTab1_TabPage0.Controls.Add(Me.fraProdNew)
        Me._SSTab1_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage0.Name = "_SSTab1_TabPage0"
        Me._SSTab1_TabPage0.Size = New System.Drawing.Size(671, 631)
        Me._SSTab1_TabPage0.TabIndex = 0
        Me._SSTab1_TabPage0.Text = "Product Code Generator"
        '
        'grpProductsList
        '
        Me.grpProductsList.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpProductsList.Controls.Add(Me.lstvwProducts)
        Me.grpProductsList.Location = New System.Drawing.Point(4, 262)
        Me.grpProductsList.Name = "grpProductsList"
        Me.grpProductsList.Size = New System.Drawing.Size(663, 366)
        Me.grpProductsList.TabIndex = 1
        Me.grpProductsList.TabStop = False
        Me.grpProductsList.Text = " Products list "
        '
        'fraProdNew
        '
        Me.fraProdNew.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.fraProdNew.BackColor = System.Drawing.SystemColors.Control
        Me.fraProdNew.Controls.Add(Me.optStrength4)
        Me.fraProdNew.Controls.Add(Me.cmdStartWizard)
        Me.fraProdNew.Controls.Add(Me.cmdValidate2)
        Me.fraProdNew.Controls.Add(Me.optStrength3)
        Me.fraProdNew.Controls.Add(Me.optStrength5)
        Me.fraProdNew.Controls.Add(Me.optStrength2)
        Me.fraProdNew.Controls.Add(Me.Label1)
        Me.fraProdNew.Controls.Add(Me.optStrength1)
        Me.fraProdNew.Controls.Add(Me.optStrength0)
        Me.fraProdNew.Controls.Add(Me.cmdProductsStorage)
        Me.fraProdNew.Controls.Add(Me.picALBanner2)
        Me.fraProdNew.Controls.Add(Me.grpCodes)
        Me.fraProdNew.Controls.Add(Me.cmdCodeGen)
        Me.fraProdNew.Controls.Add(Me.txtName)
        Me.fraProdNew.Controls.Add(Me.txtVer)
        Me.fraProdNew.Controls.Add(Me.cmdAdd)
        Me.fraProdNew.Controls.Add(Me.Label2)
        Me.fraProdNew.Controls.Add(Me.Label3)
        Me.fraProdNew.Controls.Add(Me.cmdValidate)
        Me.fraProdNew.Controls.Add(Me.cmdRemove)
        Me.fraProdNew.Controls.Add(Me.GroupBox1)
        Me.fraProdNew.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraProdNew.ForeColor = System.Drawing.SystemColors.ControlText
        Me.fraProdNew.Location = New System.Drawing.Point(0, 4)
        Me.fraProdNew.Name = "fraProdNew"
        Me.fraProdNew.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraProdNew.Size = New System.Drawing.Size(666, 256)
        Me.fraProdNew.TabIndex = 0
        Me.fraProdNew.TabStop = False
        Me.fraProdNew.Text = " Product details "
        '
        'optStrength4
        '
        Me.optStrength4.Location = New System.Drawing.Point(404, 76)
        Me.optStrength4.Name = "optStrength4"
        Me.optStrength4.Size = New System.Drawing.Size(65, 20)
        Me.optStrength4.TabIndex = 76
        Me.optStrength4.Text = "1024-bit"
        '
        'cmdStartWizard
        '
        Me.cmdStartWizard.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdStartWizard.BackColor = System.Drawing.SystemColors.Control
        Me.cmdStartWizard.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdStartWizard.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.cmdStartWizard.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdStartWizard.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdStartWizard.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdStartWizard.Location = New System.Drawing.Point(369, 12)
        Me.cmdStartWizard.Name = "cmdStartWizard"
        Me.cmdStartWizard.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdStartWizard.Size = New System.Drawing.Size(78, 23)
        Me.cmdStartWizard.TabIndex = 75
        Me.cmdStartWizard.Text = "&Wizard"
        Me.cmdStartWizard.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdStartWizard.UseVisualStyleBackColor = False
        '
        'optStrength3
        '
        Me.optStrength3.Location = New System.Drawing.Point(339, 76)
        Me.optStrength3.Name = "optStrength3"
        Me.optStrength3.Size = New System.Drawing.Size(65, 20)
        Me.optStrength3.TabIndex = 72
        Me.optStrength3.Text = "1536-bit"
        '
        'optStrength5
        '
        Me.optStrength5.Location = New System.Drawing.Point(469, 76)
        Me.optStrength5.Name = "optStrength5"
        Me.optStrength5.Size = New System.Drawing.Size(62, 20)
        Me.optStrength5.TabIndex = 71
        Me.optStrength5.Text = "512-bit"
        '
        'optStrength2
        '
        Me.optStrength2.Location = New System.Drawing.Point(274, 76)
        Me.optStrength2.Name = "optStrength2"
        Me.optStrength2.Size = New System.Drawing.Size(65, 20)
        Me.optStrength2.TabIndex = 69
        Me.optStrength2.Text = "2048-bit"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(32, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(55, 18)
        Me.Label1.TabIndex = 68
        Me.Label1.Text = "Strength:"
        '
        'optStrength1
        '
        Me.optStrength1.Location = New System.Drawing.Point(208, 76)
        Me.optStrength1.Name = "optStrength1"
        Me.optStrength1.Size = New System.Drawing.Size(66, 20)
        Me.optStrength1.TabIndex = 67
        Me.optStrength1.Text = "4096-bit"
        '
        'optStrength0
        '
        Me.optStrength0.Checked = True
        Me.optStrength0.Location = New System.Drawing.Point(88, 76)
        Me.optStrength0.Name = "optStrength0"
        Me.optStrength0.Size = New System.Drawing.Size(116, 20)
        Me.optStrength0.TabIndex = 66
        Me.optStrength0.TabStop = True
        Me.optStrength0.Text = "ALCrypto-1024-bit"
        '
        'grpCodes
        '
        Me.grpCodes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpCodes.Controls.Add(Me.txtGCode)
        Me.grpCodes.Controls.Add(Me.txtVCode)
        Me.grpCodes.Controls.Add(Me.lblGCode)
        Me.grpCodes.Controls.Add(Me.lblVCode)
        Me.grpCodes.Controls.Add(Me.cmdCopyVCode)
        Me.grpCodes.Controls.Add(Me.cmdCopyGCode)
        Me.grpCodes.Location = New System.Drawing.Point(4, 102)
        Me.grpCodes.Name = "grpCodes"
        Me.grpCodes.Size = New System.Drawing.Size(655, 123)
        Me.grpCodes.TabIndex = 6
        Me.grpCodes.TabStop = False
        Me.grpCodes.Text = " Codes "
        '
        'txtGCode
        '
        Me.txtGCode.AcceptsReturn = True
        Me.txtGCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtGCode.BackColor = System.Drawing.SystemColors.Control
        Me.txtGCode.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtGCode.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtGCode.Location = New System.Drawing.Point(84, 74)
        Me.txtGCode.MaxLength = 0
        Me.txtGCode.Multiline = True
        Me.txtGCode.Name = "txtGCode"
        Me.txtGCode.ReadOnly = True
        Me.txtGCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtGCode.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtGCode.Size = New System.Drawing.Size(565, 50)
        Me.txtGCode.TabIndex = 5
        '
        'lblGCode
        '
        Me.lblGCode.BackColor = System.Drawing.SystemColors.Control
        Me.lblGCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblGCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblGCode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblGCode.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblGCode.Location = New System.Drawing.Point(6, 74)
        Me.lblGCode.Name = "lblGCode"
        Me.lblGCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblGCode.Size = New System.Drawing.Size(78, 26)
        Me.lblGCode.TabIndex = 3
        Me.lblGCode.Text = "GCode (PRV_KEY)"
        Me.lblGCode.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'lblVCode
        '
        Me.lblVCode.BackColor = System.Drawing.SystemColors.Control
        Me.lblVCode.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblVCode.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblVCode.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblVCode.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblVCode.Location = New System.Drawing.Point(6, 20)
        Me.lblVCode.Name = "lblVCode"
        Me.lblVCode.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblVCode.Size = New System.Drawing.Size(78, 28)
        Me.lblVCode.TabIndex = 0
        Me.lblVCode.Text = "VCode (PUB_KEY)"
        Me.lblVCode.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(30, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(40, 18)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "&Name:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(30, 40)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(52, 18)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "V&ersion:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Font = New System.Drawing.Font("Arial", 7.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(204, 65)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(328, 36)
        Me.GroupBox1.TabIndex = 74
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = ".NET CLR RSA"
        '
        '_SSTab1_TabPage1
        '
        Me._SSTab1_TabPage1.Controls.Add(Me.frmKeyGen)
        Me._SSTab1_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage1.Name = "_SSTab1_TabPage1"
        Me._SSTab1_TabPage1.Size = New System.Drawing.Size(671, 630)
        Me._SSTab1_TabPage1.TabIndex = 1
        Me._SSTab1_TabPage1.Text = "License Key Generator"
        '
        'frmKeyGen
        '
        Me.frmKeyGen.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.frmKeyGen.BackColor = System.Drawing.SystemColors.Control
        Me.frmKeyGen.Controls.Add(Me.lblKeyStrength)
        Me.frmKeyGen.Controls.Add(Me.cboRegisteredLevel)
        Me.frmKeyGen.Controls.Add(Me.lblLockVideoID)
        Me.frmKeyGen.Controls.Add(Me.chkLockVideoID)
        Me.frmKeyGen.Controls.Add(Me.lblLockBaseboardID)
        Me.frmKeyGen.Controls.Add(Me.chkLockBaseboardID)
        Me.frmKeyGen.Controls.Add(Me.lblLockCPUID)
        Me.frmKeyGen.Controls.Add(Me.chkLockCPUID)
        Me.frmKeyGen.Controls.Add(Me.lblLockMemory)
        Me.frmKeyGen.Controls.Add(Me.chkLockMemory)
        Me.frmKeyGen.Controls.Add(Me.lblLockFingerprint)
        Me.frmKeyGen.Controls.Add(Me.chkLockFingerprint)
        Me.frmKeyGen.Controls.Add(Me.lblLockExternalIP)
        Me.frmKeyGen.Controls.Add(Me.chkLockExternalIP)
        Me.frmKeyGen.Controls.Add(Me.lblLockIP)
        Me.frmKeyGen.Controls.Add(Me.lblLockMotherboard)
        Me.frmKeyGen.Controls.Add(Me.lblLockBIOS)
        Me.frmKeyGen.Controls.Add(Me.lblLockWindows)
        Me.frmKeyGen.Controls.Add(Me.lblLockHDfirmware)
        Me.frmKeyGen.Controls.Add(Me.lblLockHD)
        Me.frmKeyGen.Controls.Add(Me.lblLockComputer)
        Me.frmKeyGen.Controls.Add(Me.lblLockMacAddress)
        Me.frmKeyGen.Controls.Add(Me.cmdUncheckAll)
        Me.frmKeyGen.Controls.Add(Me.cmdCheckAll)
        Me.frmKeyGen.Controls.Add(Me.lblConcurrentUsers)
        Me.frmKeyGen.Controls.Add(Me.txtMaxCount)
        Me.frmKeyGen.Controls.Add(Me.chkNetworkedLicense)
        Me.frmKeyGen.Controls.Add(Me.cmdEmailLicenseKey)
        Me.frmKeyGen.Controls.Add(Me.cmdPrintLicenseKey)
        Me.frmKeyGen.Controls.Add(Me.dtpExpireDate)
        Me.frmKeyGen.Controls.Add(Me.picALBanner)
        Me.frmKeyGen.Controls.Add(Me.chkLockIP)
        Me.frmKeyGen.Controls.Add(Me.chkLockMotherboard)
        Me.frmKeyGen.Controls.Add(Me.chkLockBIOS)
        Me.frmKeyGen.Controls.Add(Me.chkLockWindows)
        Me.frmKeyGen.Controls.Add(Me.chkLockHDfirmware)
        Me.frmKeyGen.Controls.Add(Me.chkLockHD)
        Me.frmKeyGen.Controls.Add(Me.chkLockComputer)
        Me.frmKeyGen.Controls.Add(Me.chkLockMACaddress)
        Me.frmKeyGen.Controls.Add(Me.chkItemData)
        Me.frmKeyGen.Controls.Add(Me.cmdCopy)
        Me.frmKeyGen.Controls.Add(Me.cmdPaste)
        Me.frmKeyGen.Controls.Add(Me.txtUser)
        Me.frmKeyGen.Controls.Add(Me.cmdBrowse)
        Me.frmKeyGen.Controls.Add(Me.txtLicenseFile)
        Me.frmKeyGen.Controls.Add(Me.cmdSave)
        Me.frmKeyGen.Controls.Add(Me.cboLicType)
        Me.frmKeyGen.Controls.Add(Me.txtDays)
        Me.frmKeyGen.Controls.Add(Me.txtInstallCode)
        Me.frmKeyGen.Controls.Add(Me.txtLicenseKey)
        Me.frmKeyGen.Controls.Add(Me.cmdKeyGen)
        Me.frmKeyGen.Controls.Add(Me.Label11)
        Me.frmKeyGen.Controls.Add(Me.Label5)
        Me.frmKeyGen.Controls.Add(Me.lblExpiry)
        Me.frmKeyGen.Controls.Add(Me.Label6)
        Me.frmKeyGen.Controls.Add(Me.Label7)
        Me.frmKeyGen.Controls.Add(Me.lblLicenseKey)
        Me.frmKeyGen.Controls.Add(Me.lblDays)
        Me.frmKeyGen.Controls.Add(Me.cmdViewArchive)
        Me.frmKeyGen.Controls.Add(Me.Label8)
        Me.frmKeyGen.Controls.Add(Me.Label15)
        Me.frmKeyGen.Controls.Add(Me.cboProducts)
        Me.frmKeyGen.Controls.Add(Me.cmdViewLevel)
        Me.frmKeyGen.Cursor = System.Windows.Forms.Cursors.Default
        Me.frmKeyGen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.frmKeyGen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.frmKeyGen.Location = New System.Drawing.Point(2, 0)
        Me.frmKeyGen.Name = "frmKeyGen"
        Me.frmKeyGen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.frmKeyGen.Size = New System.Drawing.Size(666, 627)
        Me.frmKeyGen.TabIndex = 0
        '
        'lblKeyStrength
        '
        Me.lblKeyStrength.BackColor = System.Drawing.SystemColors.Control
        Me.lblKeyStrength.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblKeyStrength.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblKeyStrength.ForeColor = System.Drawing.Color.Blue
        Me.lblKeyStrength.Location = New System.Drawing.Point(328, 7)
        Me.lblKeyStrength.Name = "lblKeyStrength"
        Me.lblKeyStrength.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblKeyStrength.Size = New System.Drawing.Size(338, 17)
        Me.lblKeyStrength.TabIndex = 92
        '
        'lblLockVideoID
        '
        Me.lblLockVideoID.Location = New System.Drawing.Point(293, 378)
        Me.lblLockVideoID.Name = "lblLockVideoID"
        Me.lblLockVideoID.Size = New System.Drawing.Size(368, 18)
        Me.lblLockVideoID.TabIndex = 91
        '
        'chkLockVideoID
        '
        Me.chkLockVideoID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockVideoID.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockVideoID.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockVideoID.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockVideoID.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockVideoID.Location = New System.Drawing.Point(87, 378)
        Me.chkLockVideoID.Name = "chkLockVideoID"
        Me.chkLockVideoID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockVideoID.Size = New System.Drawing.Size(279, 18)
        Me.chkLockVideoID.TabIndex = 90
        Me.chkLockVideoID.Text = "Lock to Video ID"
        Me.chkLockVideoID.UseVisualStyleBackColor = False
        '
        'lblLockBaseboardID
        '
        Me.lblLockBaseboardID.Location = New System.Drawing.Point(293, 360)
        Me.lblLockBaseboardID.Name = "lblLockBaseboardID"
        Me.lblLockBaseboardID.Size = New System.Drawing.Size(368, 18)
        Me.lblLockBaseboardID.TabIndex = 89
        '
        'chkLockBaseboardID
        '
        Me.chkLockBaseboardID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockBaseboardID.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockBaseboardID.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockBaseboardID.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockBaseboardID.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockBaseboardID.Location = New System.Drawing.Point(87, 360)
        Me.chkLockBaseboardID.Name = "chkLockBaseboardID"
        Me.chkLockBaseboardID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockBaseboardID.Size = New System.Drawing.Size(279, 18)
        Me.chkLockBaseboardID.TabIndex = 88
        Me.chkLockBaseboardID.Text = "Lock to Baseboard ID"
        Me.chkLockBaseboardID.UseVisualStyleBackColor = False
        '
        'lblLockCPUID
        '
        Me.lblLockCPUID.Location = New System.Drawing.Point(293, 342)
        Me.lblLockCPUID.Name = "lblLockCPUID"
        Me.lblLockCPUID.Size = New System.Drawing.Size(368, 18)
        Me.lblLockCPUID.TabIndex = 87
        '
        'chkLockCPUID
        '
        Me.chkLockCPUID.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockCPUID.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockCPUID.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockCPUID.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockCPUID.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockCPUID.Location = New System.Drawing.Point(87, 342)
        Me.chkLockCPUID.Name = "chkLockCPUID"
        Me.chkLockCPUID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockCPUID.Size = New System.Drawing.Size(279, 18)
        Me.chkLockCPUID.TabIndex = 86
        Me.chkLockCPUID.Text = "Lock to CPU ID"
        Me.chkLockCPUID.UseVisualStyleBackColor = False
        '
        'lblLockMemory
        '
        Me.lblLockMemory.Location = New System.Drawing.Point(293, 324)
        Me.lblLockMemory.Name = "lblLockMemory"
        Me.lblLockMemory.Size = New System.Drawing.Size(368, 18)
        Me.lblLockMemory.TabIndex = 85
        '
        'chkLockMemory
        '
        Me.chkLockMemory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockMemory.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockMemory.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockMemory.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockMemory.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockMemory.Location = New System.Drawing.Point(87, 324)
        Me.chkLockMemory.Name = "chkLockMemory"
        Me.chkLockMemory.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockMemory.Size = New System.Drawing.Size(279, 18)
        Me.chkLockMemory.TabIndex = 84
        Me.chkLockMemory.Text = "Lock to Memory ID"
        Me.chkLockMemory.UseVisualStyleBackColor = False
        '
        'lblLockFingerprint
        '
        Me.lblLockFingerprint.Location = New System.Drawing.Point(293, 306)
        Me.lblLockFingerprint.Name = "lblLockFingerprint"
        Me.lblLockFingerprint.Size = New System.Drawing.Size(368, 18)
        Me.lblLockFingerprint.TabIndex = 83
        '
        'chkLockFingerprint
        '
        Me.chkLockFingerprint.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockFingerprint.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockFingerprint.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockFingerprint.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockFingerprint.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockFingerprint.Location = New System.Drawing.Point(87, 306)
        Me.chkLockFingerprint.Name = "chkLockFingerprint"
        Me.chkLockFingerprint.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockFingerprint.Size = New System.Drawing.Size(279, 18)
        Me.chkLockFingerprint.TabIndex = 82
        Me.chkLockFingerprint.Text = "Lock to Computer Fingerprint"
        Me.chkLockFingerprint.UseVisualStyleBackColor = False
        '
        'lblLockExternalIP
        '
        Me.lblLockExternalIP.Location = New System.Drawing.Point(293, 288)
        Me.lblLockExternalIP.Name = "lblLockExternalIP"
        Me.lblLockExternalIP.Size = New System.Drawing.Size(368, 18)
        Me.lblLockExternalIP.TabIndex = 81
        '
        'chkLockExternalIP
        '
        Me.chkLockExternalIP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockExternalIP.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockExternalIP.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockExternalIP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockExternalIP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockExternalIP.Location = New System.Drawing.Point(87, 288)
        Me.chkLockExternalIP.Name = "chkLockExternalIP"
        Me.chkLockExternalIP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockExternalIP.Size = New System.Drawing.Size(279, 18)
        Me.chkLockExternalIP.TabIndex = 80
        Me.chkLockExternalIP.Text = "Lock to External IP Address"
        Me.chkLockExternalIP.UseVisualStyleBackColor = False
        '
        'lblLockIP
        '
        Me.lblLockIP.Location = New System.Drawing.Point(293, 270)
        Me.lblLockIP.Name = "lblLockIP"
        Me.lblLockIP.Size = New System.Drawing.Size(368, 18)
        Me.lblLockIP.TabIndex = 79
        '
        'lblLockMotherboard
        '
        Me.lblLockMotherboard.Location = New System.Drawing.Point(293, 252)
        Me.lblLockMotherboard.Name = "lblLockMotherboard"
        Me.lblLockMotherboard.Size = New System.Drawing.Size(368, 18)
        Me.lblLockMotherboard.TabIndex = 78
        '
        'lblLockBIOS
        '
        Me.lblLockBIOS.Location = New System.Drawing.Point(293, 234)
        Me.lblLockBIOS.Name = "lblLockBIOS"
        Me.lblLockBIOS.Size = New System.Drawing.Size(368, 18)
        Me.lblLockBIOS.TabIndex = 77
        '
        'lblLockWindows
        '
        Me.lblLockWindows.Location = New System.Drawing.Point(293, 216)
        Me.lblLockWindows.Name = "lblLockWindows"
        Me.lblLockWindows.Size = New System.Drawing.Size(368, 18)
        Me.lblLockWindows.TabIndex = 76
        '
        'lblLockHDfirmware
        '
        Me.lblLockHDfirmware.Location = New System.Drawing.Point(293, 198)
        Me.lblLockHDfirmware.Name = "lblLockHDfirmware"
        Me.lblLockHDfirmware.Size = New System.Drawing.Size(368, 18)
        Me.lblLockHDfirmware.TabIndex = 75
        '
        'lblLockHD
        '
        Me.lblLockHD.Location = New System.Drawing.Point(293, 180)
        Me.lblLockHD.Name = "lblLockHD"
        Me.lblLockHD.Size = New System.Drawing.Size(368, 18)
        Me.lblLockHD.TabIndex = 74
        '
        'lblLockComputer
        '
        Me.lblLockComputer.Location = New System.Drawing.Point(293, 162)
        Me.lblLockComputer.Name = "lblLockComputer"
        Me.lblLockComputer.Size = New System.Drawing.Size(368, 18)
        Me.lblLockComputer.TabIndex = 73
        '
        'lblLockMacAddress
        '
        Me.lblLockMacAddress.Location = New System.Drawing.Point(293, 143)
        Me.lblLockMacAddress.Name = "lblLockMacAddress"
        Me.lblLockMacAddress.Size = New System.Drawing.Size(368, 18)
        Me.lblLockMacAddress.TabIndex = 72
        '
        'cmdUncheckAll
        '
        Me.cmdUncheckAll.Location = New System.Drawing.Point(4, 214)
        Me.cmdUncheckAll.Name = "cmdUncheckAll"
        Me.cmdUncheckAll.Size = New System.Drawing.Size(72, 20)
        Me.cmdUncheckAll.TabIndex = 71
        Me.cmdUncheckAll.Text = "Uncheck All"
        '
        'cmdCheckAll
        '
        Me.cmdCheckAll.Location = New System.Drawing.Point(4, 188)
        Me.cmdCheckAll.Name = "cmdCheckAll"
        Me.cmdCheckAll.Size = New System.Drawing.Size(72, 20)
        Me.cmdCheckAll.TabIndex = 70
        Me.cmdCheckAll.Text = "Check All"
        '
        'lblConcurrentUsers
        '
        Me.lblConcurrentUsers.BackColor = System.Drawing.SystemColors.Control
        Me.lblConcurrentUsers.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblConcurrentUsers.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConcurrentUsers.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblConcurrentUsers.Location = New System.Drawing.Point(222, 79)
        Me.lblConcurrentUsers.Name = "lblConcurrentUsers"
        Me.lblConcurrentUsers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblConcurrentUsers.Size = New System.Drawing.Size(100, 17)
        Me.lblConcurrentUsers.TabIndex = 69
        Me.lblConcurrentUsers.Text = "Concurrent Users"
        Me.lblConcurrentUsers.Visible = False
        '
        'chkNetworkedLicense
        '
        Me.chkNetworkedLicense.BackColor = System.Drawing.SystemColors.Control
        Me.chkNetworkedLicense.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkNetworkedLicense.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkNetworkedLicense.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkNetworkedLicense.Location = New System.Drawing.Point(88, 76)
        Me.chkNetworkedLicense.Name = "chkNetworkedLicense"
        Me.chkNetworkedLicense.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkNetworkedLicense.Size = New System.Drawing.Size(122, 22)
        Me.chkNetworkedLicense.TabIndex = 67
        Me.chkNetworkedLicense.Text = "Networked License"
        Me.chkNetworkedLicense.UseVisualStyleBackColor = False
        '
        'dtpExpireDate
        '
        '* Me.dtpExpireDate.CustomFormat = "yyyy/MM/dd"
        '* Me.dtpExpireDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpExpireDate.CustomFormat = ""
        Me.dtpExpireDate.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.dtpExpireDate.Location = New System.Drawing.Point(88, 52)
        Me.dtpExpireDate.Name = "dtpExpireDate"
        Me.dtpExpireDate.Size = New System.Drawing.Size(118, 20)
        Me.dtpExpireDate.TabIndex = 10
        Me.dtpExpireDate.Value = New Date(2006, 5, 25, 20, 28, 52, 803)
        Me.dtpExpireDate.Visible = False
        '
        'chkLockIP
        '
        Me.chkLockIP.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockIP.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockIP.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockIP.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockIP.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockIP.Location = New System.Drawing.Point(87, 270)
        Me.chkLockIP.Name = "chkLockIP"
        Me.chkLockIP.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockIP.Size = New System.Drawing.Size(279, 18)
        Me.chkLockIP.TabIndex = 24
        Me.chkLockIP.Text = "Lock to Local IP Address"
        Me.chkLockIP.UseVisualStyleBackColor = False
        '
        'chkLockMotherboard
        '
        Me.chkLockMotherboard.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockMotherboard.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockMotherboard.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockMotherboard.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockMotherboard.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockMotherboard.Location = New System.Drawing.Point(87, 252)
        Me.chkLockMotherboard.Name = "chkLockMotherboard"
        Me.chkLockMotherboard.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockMotherboard.Size = New System.Drawing.Size(279, 18)
        Me.chkLockMotherboard.TabIndex = 23
        Me.chkLockMotherboard.Text = "Lock to Motherboard Serial"
        Me.chkLockMotherboard.UseVisualStyleBackColor = False
        '
        'chkLockBIOS
        '
        Me.chkLockBIOS.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockBIOS.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockBIOS.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockBIOS.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockBIOS.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockBIOS.Location = New System.Drawing.Point(87, 234)
        Me.chkLockBIOS.Name = "chkLockBIOS"
        Me.chkLockBIOS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockBIOS.Size = New System.Drawing.Size(279, 18)
        Me.chkLockBIOS.TabIndex = 22
        Me.chkLockBIOS.Text = "Lock to BIOS Version"
        Me.chkLockBIOS.UseVisualStyleBackColor = False
        '
        'chkLockWindows
        '
        Me.chkLockWindows.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockWindows.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockWindows.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockWindows.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockWindows.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockWindows.Location = New System.Drawing.Point(87, 216)
        Me.chkLockWindows.Name = "chkLockWindows"
        Me.chkLockWindows.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockWindows.Size = New System.Drawing.Size(279, 18)
        Me.chkLockWindows.TabIndex = 21
        Me.chkLockWindows.Text = "Lock to Windows Serial"
        Me.chkLockWindows.UseVisualStyleBackColor = False
        '
        'chkLockHDfirmware
        '
        Me.chkLockHDfirmware.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockHDfirmware.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockHDfirmware.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockHDfirmware.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockHDfirmware.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockHDfirmware.Location = New System.Drawing.Point(87, 198)
        Me.chkLockHDfirmware.Name = "chkLockHDfirmware"
        Me.chkLockHDfirmware.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockHDfirmware.Size = New System.Drawing.Size(279, 18)
        Me.chkLockHDfirmware.TabIndex = 20
        Me.chkLockHDfirmware.Text = "Lock to HDD Firmware Serial"
        Me.chkLockHDfirmware.UseVisualStyleBackColor = False
        '
        'chkLockHD
        '
        Me.chkLockHD.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockHD.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockHD.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockHD.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockHD.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockHD.Location = New System.Drawing.Point(87, 180)
        Me.chkLockHD.Name = "chkLockHD"
        Me.chkLockHD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockHD.Size = New System.Drawing.Size(279, 18)
        Me.chkLockHD.TabIndex = 19
        Me.chkLockHD.Text = "Lock to HDD Volume Serial"
        Me.chkLockHD.UseVisualStyleBackColor = False
        '
        'chkLockComputer
        '
        Me.chkLockComputer.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockComputer.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockComputer.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockComputer.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockComputer.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockComputer.Location = New System.Drawing.Point(87, 162)
        Me.chkLockComputer.Name = "chkLockComputer"
        Me.chkLockComputer.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockComputer.Size = New System.Drawing.Size(279, 18)
        Me.chkLockComputer.TabIndex = 18
        Me.chkLockComputer.Text = "Lock to Computer Name"
        Me.chkLockComputer.UseVisualStyleBackColor = False
        '
        'chkLockMACaddress
        '
        Me.chkLockMACaddress.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkLockMACaddress.BackColor = System.Drawing.SystemColors.Control
        Me.chkLockMACaddress.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkLockMACaddress.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkLockMACaddress.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkLockMACaddress.Location = New System.Drawing.Point(87, 143)
        Me.chkLockMACaddress.Name = "chkLockMACaddress"
        Me.chkLockMACaddress.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkLockMACaddress.Size = New System.Drawing.Size(279, 18)
        Me.chkLockMACaddress.TabIndex = 17
        Me.chkLockMACaddress.Text = "Lock to MAC Address"
        Me.chkLockMACaddress.UseVisualStyleBackColor = False
        '
        'chkItemData
        '
        Me.chkItemData.BackColor = System.Drawing.SystemColors.Control
        Me.chkItemData.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkItemData.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkItemData.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkItemData.Location = New System.Drawing.Point(418, 49)
        Me.chkItemData.Name = "chkItemData"
        Me.chkItemData.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkItemData.Size = New System.Drawing.Size(146, 22)
        Me.chkItemData.TabIndex = 7
        Me.chkItemData.Text = "Use Item Data for Code"
        Me.chkItemData.UseVisualStyleBackColor = False
        '
        'txtDays
        '
        Me.txtDays.AcceptsReturn = True
        Me.txtDays.BackColor = System.Drawing.SystemColors.Control
        Me.txtDays.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDays.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDays.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDays.Location = New System.Drawing.Point(88, 52)
        Me.txtDays.MaxLength = 0
        Me.txtDays.Name = "txtDays"
        Me.txtDays.ReadOnly = True
        Me.txtDays.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDays.Size = New System.Drawing.Size(117, 20)
        Me.txtDays.TabIndex = 9
        Me.txtDays.Text = "365"
        Me.txtDays.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(3, 123)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(89, 17)
        Me.Label11.TabIndex = 15
        Me.Label11.Text = "User Name:"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(2, 605)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(78, 17)
        Me.Label5.TabIndex = 30
        Me.Label5.Text = "License &file:"
        '
        'lblExpiry
        '
        Me.lblExpiry.BackColor = System.Drawing.SystemColors.Control
        Me.lblExpiry.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblExpiry.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExpiry.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblExpiry.Location = New System.Drawing.Point(2, 54)
        Me.lblExpiry.Name = "lblExpiry"
        Me.lblExpiry.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblExpiry.Size = New System.Drawing.Size(89, 17)
        Me.lblExpiry.TabIndex = 8
        Me.lblExpiry.Text = "&Expires After:"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(2, 30)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(89, 17)
        Me.Label6.TabIndex = 5
        Me.Label6.Text = "License &Type:"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(3, 101)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(89, 17)
        Me.Label7.TabIndex = 12
        Me.Label7.Text = "Installation C&ode:"
        '
        'lblLicenseKey
        '
        Me.lblLicenseKey.BackColor = System.Drawing.SystemColors.Control
        Me.lblLicenseKey.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLicenseKey.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLicenseKey.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLicenseKey.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblLicenseKey.Location = New System.Drawing.Point(1, 351)
        Me.lblLicenseKey.Name = "lblLicenseKey"
        Me.lblLicenseKey.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLicenseKey.Size = New System.Drawing.Size(82, 59)
        Me.lblLicenseKey.TabIndex = 25
        Me.lblLicenseKey.Text = "License &Key:"
        Me.lblLicenseKey.TextAlign = System.Drawing.ContentAlignment.BottomLeft
        '
        'lblDays
        '
        Me.lblDays.BackColor = System.Drawing.SystemColors.Control
        Me.lblDays.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDays.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDays.Location = New System.Drawing.Point(208, 56)
        Me.lblDays.Name = "lblDays"
        Me.lblDays.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDays.Size = New System.Drawing.Size(114, 17)
        Me.lblDays.TabIndex = 11
        Me.lblDays.Text = "days"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(2, 6)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(65, 17)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "&Product:"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.SystemColors.Control
        Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label15.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label15.Location = New System.Drawing.Point(326, 27)
        Me.Label15.Name = "Label15"
        Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label15.Size = New System.Drawing.Size(94, 17)
        Me.Label15.TabIndex = 2
        Me.Label15.Text = "Registered Level:"
        '
        'sbStatus
        '
        Me.sbStatus.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.sbStatus.Location = New System.Drawing.Point(0, 657)
        Me.sbStatus.Name = "sbStatus"
        Me.sbStatus.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.mainStatusBarPanel})
        Me.sbStatus.ShowPanels = True
        Me.sbStatus.Size = New System.Drawing.Size(679, 22)
        Me.sbStatus.TabIndex = 1
        Me.sbStatus.Text = "Ready ..."
        '
        'mainStatusBarPanel
        '
        Me.mainStatusBarPanel.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.mainStatusBarPanel.Name = "mainStatusBarPanel"
        Me.mainStatusBarPanel.Text = "Ready ..."
        Me.mainStatusBarPanel.Width = 662
        '
        'lnkActivelockSoftwareGroup
        '
        Me.lnkActivelockSoftwareGroup.ActiveLinkColor = System.Drawing.Color.Green
        Me.lnkActivelockSoftwareGroup.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lnkActivelockSoftwareGroup.Location = New System.Drawing.Point(515, 4)
        Me.lnkActivelockSoftwareGroup.Name = "lnkActivelockSoftwareGroup"
        Me.lnkActivelockSoftwareGroup.Size = New System.Drawing.Size(156, 12)
        Me.lnkActivelockSoftwareGroup.TabIndex = 66
        Me.lnkActivelockSoftwareGroup.TabStop = True
        Me.lnkActivelockSoftwareGroup.Text = "www.activelocksoftware.com"
        Me.lnkActivelockSoftwareGroup.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lnkActivelockSoftwareGroup.VisitedLinkColor = System.Drawing.Color.Blue
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(679, 679)
        Me.Controls.Add(Me.lnkActivelockSoftwareGroup)
        Me.Controls.Add(Me.sbStatus)
        Me.Controls.Add(Me.SSTab1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MinimumSize = New System.Drawing.Size(588, 576)
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ALUGEN - ActiveLock3 Universal GENerator"
        CType(Me.picALBanner2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picALBanner, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SSTab1.ResumeLayout(False)
        Me._SSTab1_TabPage0.ResumeLayout(False)
        Me.grpProductsList.ResumeLayout(False)
        Me.fraProdNew.ResumeLayout(False)
        Me.fraProdNew.PerformLayout()
        Me.grpCodes.ResumeLayout(False)
        Me.grpCodes.PerformLayout()
        Me._SSTab1_TabPage1.ResumeLayout(False)
        Me.frmKeyGen.ResumeLayout(False)
        Me.frmKeyGen.PerformLayout()
        CType(Me.mainStatusBarPanel, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region


#Region "Methods"

    Private Sub LoadResources()
        Dim resxReader As Resources.ResXResourceReader = New Resources.ResXResourceReader("Alugen3NET.resx")
        Dim resxDE As IDictionaryEnumerator = resxReader.GetEnumerator
        For Each dd As DictionaryEntry In resxReader
            resxList.Add(dd.Key, dd.Value)
        Next
    End Sub

    Private Sub LoadImages()
        'load buttons images
        cmdAdd.Image = CType(resxList("add"), Image)
        cmdRemove.Image = CType(resxList("delete"), Image)
        cmdCodeGen.Image = CType(resxList("generate"), Image)
        cmdKeyGen.Image = CType(resxList("generate"), Image)
        cmdValidate.Image = CType(resxList("validate"), Image)
        cmdCopyVCode.Image = CType(resxList("copy"), Image)
        cmdCopyGCode.Image = CType(resxList("copy"), Image)
        cmdCopy.Image = CType(resxList("copy"), Image)
        cmdViewLevel.Image = CType(resxList("preview"), Image)
        cmdPaste.Image = CType(resxList("paste"), Image)
        cmdSave.Image = CType(resxList("save"), Image)
        cmdViewArchive.Image = CType(resxList("viewdatabase"), Image)
        cmdBrowse.Image = CType(resxList("folder"), Image)
        cmdPrintLicenseKey.Image = CType(resxList("print"), Image)
        cmdEmailLicenseKey.Image = CType(resxList("email"), Image)
        cmdProductsStorage.Image = CType(resxList("database"), Image)
        cmdStartWizard.Image = CType(resxList("wizard"), Image)
        'lbl's
        lblVCode.Image = CType(resxList("keys"), Image)
        lblGCode.Image = CType(resxList("keys"), Image)
        lblLicenseKey.Image = CType(resxList("KeyLock"), Image)
        'AL banners
        picALBanner.Image = CType(resxList("I_Trust_AL_small"), Image)
        picALBanner2.Image = CType(resxList("I_Trust_AL_small"), Image)
    End Sub

    Private Sub AppendLockString(ByRef strLock As String, ByVal newSubString As String)
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

        If strLock = "" Then
            strLock = newSubString
        Else
            strLock = strLock & vbLf & newSubString
        End If
    End Sub

    Private Function ReconstructedInstallationCode() As String
        Dim strLock As String = Nothing
        Dim strReq As String
        Dim noKey As String
        noKey = Chr(110) & Chr(111) & Chr(107) & Chr(101) & Chr(121)

        If Me.chkLockMACaddress.CheckState = CheckState.Checked Then
            AppendLockString(strLock, MACaddress)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockComputer.CheckState = CheckState.Checked Then
            AppendLockString(strLock, ComputerName)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockHD.CheckState = CheckState.Checked Then
            AppendLockString(strLock, VolumeSerial)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockHDfirmware.CheckState = CheckState.Checked Then
            AppendLockString(strLock, FirmwareSerial)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockWindows.CheckState = CheckState.Checked Then
            AppendLockString(strLock, WindowsSerial)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockBIOS.CheckState = CheckState.Checked Then
            AppendLockString(strLock, BIOSserial)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockMotherboard.CheckState = CheckState.Checked Then
            AppendLockString(strLock, MotherboardSerial)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockIP.CheckState = CheckState.Checked Then
            AppendLockString(strLock, IPaddress)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockExternalIP.CheckState = CheckState.Checked Then
            AppendLockString(strLock, ExternalIP)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockFingerprint.CheckState = CheckState.Checked Then
            AppendLockString(strLock, Fingerprint)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockMemory.CheckState = CheckState.Checked Then
            AppendLockString(strLock, Memory)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockCPUID.CheckState = CheckState.Checked Then
            AppendLockString(strLock, CPUID)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockBaseboardID.CheckState = CheckState.Checked Then
            AppendLockString(strLock, BaseboardID)
        Else
            AppendLockString(strLock, noKey)
        End If
        If Me.chkLockVideoID.CheckState = CheckState.Checked Then
            AppendLockString(strLock, VideoID)
        Else
            AppendLockString(strLock, noKey)
        End If

        If Not strLock Is Nothing _
          AndAlso strLock.Substring(0, 1) = vbLf Then
            strLock = strLock.Substring(2)
        End If

        Dim Index, i As Integer
        Dim strInstCode As String
        strInstCode = ActiveLock3Globals_definst.Base64Decode(txtInstallCode.Text)

        If strInstCode = "" Then Return Nothing

        If Not strInstCode Is Nothing _
          AndAlso strInstCode.Substring(0, 1) = "+" Then
            strInstCode = strInstCode.Substring(2)
        End If
        Dim arrProdVer() As String
        arrProdVer = Split(strInstCode, "&&&") ' Extract the software name and version
        strInstCode = arrProdVer(0)
        Index = 0
        i = 1
        ' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
        Do While i > 0
            i = strInstCode.IndexOf(vbLf, Index) 'InStr(Index + 1, strInstCode, vbLf)
            If i > 0 Then Index = i + 1
        Loop
        ' user name starts from Index+1 to the end
        Dim user As String
        user = strInstCode.Substring(Index)

        ' combine with user name
        strReq = strLock & vbLf & user

        ' combine with app name and version
        strReq = strReq & "&&&" & cboProducts.Text

        ' base-64 encode the request
        Dim strReq2 As String
        strReq2 = ActiveLock3Globals_definst.Base64Encode("+" & strReq)
        ReconstructedInstallationCode = strReq2

    End Function

    Private Function GetUserSoftwareNameandVersionfromInstallCode(ByVal strInstCode As String) As String
        On Error GoTo noInfo
        If strInstCode = "" Then Return String.Empty
        strInstCode = ActiveLock3Globals_definst.Base64Decode(txtInstallCode.Text)
        Dim arrProdVer() As String
        arrProdVer = Split(strInstCode, "&&&")
        GetUserSoftwareNameandVersionfromInstallCode = Trim$(arrProdVer(1))
noInfo:
    End Function

    Private Sub UpdateKeyGenButtonStatus()
        If txtInstallCode.Text = "" Then
            cmdKeyGen.Enabled = False
        Else
            If cboProducts.SelectedIndex >= 0 Then
                cmdKeyGen.Enabled = True
            End If
        End If
    End Sub

    Private Function Make64ByteChunks(ByRef strdata As String) As String
        ' Breaks a long string into chunks of 64-byte lines.
        Dim i As Integer
        Dim Count As Integer
        Dim strNew64Chunk As String
        Dim sResult As String = ""

        Count = strdata.Length
        For i = 0 To Count Step 64
            If i + 64 > Count Then
                strNew64Chunk = strdata.Substring(i)

            Else
                strNew64Chunk = strdata.Substring(i, 64)
            End If
            If sResult.Length > 0 Then
                sResult = sResult & strNew64Chunk
                'sResult = sResult & vbCrLf & strNew64Chunk
            Else
                sResult = sResult & strNew64Chunk
            End If
        Next

        Make64ByteChunks = sResult
    End Function
    '* Convenience function coverts Expire date to string
    Private Function GetExpirationString() As String '*
        Return GetExpirationDate().ToString '*
    End Function '*


    Private Function GetExpirationDate() As Date
        If cboLicType.Text = "Time Locked" Then
            'GetExpirationDate = txtDays.Text
            '* GetExpirationDate = CType(dtpExpireDate.Value, DateTime).ToString("yyyy/MM/dd")
            Return CType(dtpExpireDate.Value, DateTime) '*
        Else
            GetExpirationDate = Date.UtcNow.AddDays(CShort(txtDays.Text)) '*.ToString("yyyy/MM/dd")
        End If
    End Function


    Private Sub SaveLicenseKey(ByVal sLibKey As String, ByVal sFileName As String)
        Dim hFile As Integer
        hFile = FreeFile()
        FileOpen(hFile, sFileName, OpenMode.Output)
        PrintLine(hFile, sLibKey)
        FileClose(hFile)
    End Sub

    Private Sub UpdateStatus(ByRef Msg As String)
        'write status on fist status bar panel
        sbStatus.Panels(0).Text = Msg
    End Sub

    Function CheckForResources(ByVal ParamArray MyArray() As Object) As Boolean
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

        Try

            Dim foundIt As Boolean
            Dim Y As Object
            Dim j As Integer
            Dim systemDir, s, pathName As String

            WhereIsDLL("") 'initialize

            systemDir = WinSysDir() 'Get the Windows system directory
            For Each Y In MyArray
                foundIt = False
                s = CStr(Y)

                If s.Substring(0, 1) = "#" Then
                    pathName = Application.StartupPath
                    s = s.Substring(1)
                ElseIf s.IndexOf("\") > 0 Then
                    j = s.LastIndexOf("\") 'InStrRev(s, "\")
                    '!!!!!!!!!!!!!! TODO ?
                    'pathName = s.Substring(s, j - 1)
                    pathName = s.Substring(0, j - 1)
                    s = s.Substring(j + 1)
                Else
                    pathName = systemDir
                End If

                If s.IndexOf(".") > 0 Then
                    If File.Exists(pathName & "\" & s) Then foundIt = True
                ElseIf File.Exists(pathName & "\" & s & ".DLL") Then
                    foundIt = True
                ElseIf File.Exists(pathName & "\" & s & ".OCX") Then
                    foundIt = True
                    s = s & ".OCX" 'this will make the softlocx check easier
                End If

                If Not foundIt Then
                    MsgBox(s & " could not be found in " & pathName & vbCrLf & System.Reflection.Assembly.GetExecutingAssembly.GetName.Name & " cannot run without this library file!" & vbCrLf & vbCrLf & "Exiting!", MsgBoxStyle.Critical, "Missing Resource")
                    End
                End If
            Next Y

            CheckForResources = True
        Catch ex As Exception
            MessageBox.Show(ex.Message, ACTIVELOCKSTRING, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try


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

        If T = "" Then
            WhereIsDLL = ""
            Exit Function
        End If

        Static a As String() 'Object
        Dim s, D As String
        Dim EnvString As String
        Dim Indx As Integer
        Dim i As Integer

        On Error Resume Next
        i = UBound(a)
        If i = 0 Then
            s = Environment.GetEnvironmentVariable("PATH") & ";" & CurDir() & ";"

            D = WinSysDir()
            s = s & D & ";"

            If D.Substring(D.Length - 2) = "32" Then 'I'm guessing at the name of the 16 bit windows directory (assuming it exists)
                i = D.Length
                '!!!!!!!! TODO ?
                's = s & D.Substring(D, i - 2) & ";"
                s = s & D.Substring(0, i - 2) & ";"
            End If

            s = s & WinDir() & ";"
            Indx = 1 ' Initialize index to 1.
            Do
                EnvString = Environ(Indx) ' Get environment variable.
                If StrComp(EnvString.Substring(0, 5), "PATH=", CompareMethod.Text) = 0 Then ' Check PATH entry.
                    s = s & EnvString.Substring(6)
                    Exit Do
                End If
                Indx += 1
            Loop Until EnvString = ""
            a = Split(s, ";")
        End If

        T = Trim(T)
        If T = "" Then Return Nothing
        If Not T.Substring(T.Length - 4).IndexOf(".") > 0 Then T = T & ".DLL" 'default extension
        For i = 0 To UBound(a)
            If File.Exists(a(i) & "\" & T) Then
                WhereIsDLL = a(i)
                Exit Function
            End If
        Next i
        Return Nothing
    End Function

    Private Sub AddRow(ByRef productName As String, ByRef productVer As String, ByRef productVCode As String, ByRef productGCode As String, Optional ByRef fUpdateStore As Boolean = True)
        ' Add a Product Row to the GUI.
        ' If fUpdateStore is True, then product info is also saved to the store.
        ' Call Activelock3.IALUGenerator to add product
        If fUpdateStore Then
            Dim ProdInfo As ProductInfo
            ProdInfo = ActiveLock3AlugenGlobals_definst.CreateProductInfo(productName, productVer, productVCode, productGCode)
            Call GeneratorInstance.SaveProduct(ProdInfo)
        End If
        ' Update the view
        Dim itemProductInfo As New ProductInfoItem(productName, productVer)
        cboProducts.Items.Add(itemProductInfo)
        ReDim Preserve cboProducts_Array(cboProducts.Items.Count - 1)
        If productVCode.Contains("RSA512") Then
            cboProducts_Array(cboProducts.Items.Count - 1) = "512"
        ElseIf productVCode.Contains("RSA1024") Then
            cboProducts_Array(cboProducts.Items.Count - 1) = "1024"
        ElseIf productVCode.Contains("RSA1536") Then
            cboProducts_Array(cboProducts.Items.Count - 1) = "1536"
        ElseIf productVCode.Contains("RSA2048") Then
            cboProducts_Array(cboProducts.Items.Count - 1) = "2048"
        ElseIf productVCode.Contains("RSA4096") Then
            cboProducts_Array(cboProducts.Items.Count - 1) = "4096"
        Else ' ALCrypto 1024-bit
            cboProducts_Array(cboProducts.Items.Count - 1) = "0"
        End If

        Dim mainListItem As New ListViewItem(productName)
        mainListItem.SubItems.Add(productVer)
        mainListItem.SubItems.Add(productVCode)
        mainListItem.SubItems.Add(productGCode)
        lstvwProducts.BeginUpdate()
        lstvwProducts.Items.Add(mainListItem)
        lstvwProducts.EndUpdate()
        mainListItem.Selected = True
        cmdRemove.Enabled = True
    End Sub

    Private Sub UpdateCodeGenButtonStatus()
        If txtName.Text = "" Or txtVer.Text = "" Then
            cmdCodeGen.Enabled = False
        ElseIf CheckDuplicate(txtName.Text, txtVer.Text) Then
            cmdCodeGen.Enabled = False
        Else
            cmdCodeGen.Enabled = True
        End If
    End Sub

    Private Sub UpdateAddButtonStatus()
        If txtName.Text = "" Or txtVer.Text = "" Or txtVCode.Text = "" Then
            cmdAdd.Enabled = False
        ElseIf CheckDuplicate(txtName.Text, txtVer.Text) Then
            cmdAdd.Enabled = False
        Else
            cmdAdd.Enabled = True
        End If
    End Sub

    Private Function CheckDuplicate(ByRef productName As String, ByRef productVer As String) As Boolean
        CheckDuplicate = False
        Dim i As Integer
        For i = 0 To lstvwProducts.Items.Count - 1
            If lstvwProducts.Items(i).Text = productName _
              AndAlso lstvwProducts.Items(i).SubItems(1).Text = productVer Then
                If Not fDisableNotifications Then
                    txtVCode.Text = lstvwProducts.Items(i).SubItems(2).Text
                    txtGCode.Text = lstvwProducts.Items(i).SubItems(3).Text
                End If
                CheckDuplicate = True
                Exit Function
            End If
        Next
        If Not fDisableNotifications Then
            txtVCode.Text = ""
            txtGCode.Text = ""
        End If
    End Function

    Private Sub InitUI()
        ' Initialize the GUI

        txtLicenseKey.Text = ""
        cboLicType.Text = "Periodic"

        cboProducts.DisplayMember = "ProductNameVersion"
        cboProducts.ValueMember = "ProductNameVersion"

        lstvwProducts.Items.Clear()
        cboProducts.Items.Clear()

        ' Populate Product List on Product Code Generator tab
        ' and Key Gen tab with product info from licenses.ini
        Dim arrProdInfos() As ProductInfo
        arrProdInfos = GeneratorInstance.RetrieveProducts()

        If IsArrayEmpty(arrProdInfos) Then Exit Sub

        Dim i As Integer
        For i = LBound(arrProdInfos) To UBound(arrProdInfos)
            PopulateUI(CType(arrProdInfos(i), ProductInfo))
        Next
        lstvwProducts.Items(0).Selected = True
    End Sub

    Private Function IsArrayEmpty(ByRef arrVar As ProductInfo()) As Boolean
        IsArrayEmpty = True
        Try
            Dim lb As Integer
            lb = UBound(arrVar, 1) ' this will raise an error if the array is empty
            IsArrayEmpty = False ' If we managed to get to here, then it's not empty
        Catch ex As Exception
            'MessageBox.Show(ex.Message, ACTIVELOCKSTRING, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub PopulateUI(ByRef ProdInfo As ProductInfo)
        With ProdInfo
            AddRow(.Name, .Version, .VCode, .GCode, False)
        End With
    End Sub

    Private Function GetUserFromInstallCode(ByVal strInstCode As String) As String
        ' Retrieves lock string and user info from the request string
        '
        Dim a() As String
        Dim i As Integer
        Dim aString As String
        Dim usedLockNone As Boolean
        Dim noKey As String
        noKey = Chr(110) & Chr(111) & Chr(107) & Chr(101) & Chr(121)

        If strInstCode = "" Then Return Nothing
        strInstCode = ActiveLock3Globals_definst.Base64Decode(strInstCode)

        If strInstCode <> "" Then
            If strInstCode.Substring(0, 1) = "+" Then
                strInstCode = strInstCode.Substring(1)
                usedLockNone = True
            End If
        End If
        Dim arrProdVer() As String
        arrProdVer = Split(strInstCode, "&&&") ' Extract the software name and version
        strInstCode = arrProdVer(0)

        systemEvent = True
        'clean checkboxes
        chkLockMACaddress.Enabled = True
        chkLockComputer.Enabled = True
        chkLockHD.Enabled = True
        chkLockHDfirmware.Enabled = True
        chkLockWindows.Enabled = True
        chkLockBIOS.Enabled = True
        chkLockMotherboard.Enabled = True
        chkLockIP.Enabled = True
        chkLockExternalIP.Enabled = True
        chkLockFingerprint.Enabled = True
        chkLockMemory.Enabled = True
        chkLockCPUID.Enabled = True
        chkLockBaseboardID.Enabled = True
        chkLockVideoID.Enabled = True

        a = Split(strInstCode, vbLf)
        If usedLockNone = True Then
            For i = LBound(a) To UBound(a) - 1
                aString = a(i)
                If i = LBound(a) Then
                    MACaddress = aString
                    lblLockMacAddress.Text = MACaddress
                ElseIf i = LBound(a) + 1 Then
                    ComputerName = aString
                    lblLockComputer.Text = ComputerName
                ElseIf i = LBound(a) + 2 Then
                    VolumeSerial = aString
                    lblLockHD.Text = VolumeSerial
                ElseIf i = LBound(a) + 3 Then
                    FirmwareSerial = aString
                    lblLockHDfirmware.Text = FirmwareSerial
                ElseIf i = LBound(a) + 4 Then
                    WindowsSerial = aString
                    lblLockWindows.Text = WindowsSerial
                ElseIf i = LBound(a) + 5 Then
                    BIOSserial = aString
                    lblLockBIOS.Text = BIOSserial
                ElseIf i = LBound(a) + 6 Then
                    MotherboardSerial = aString
                    lblLockMotherboard.Text = MotherboardSerial
                ElseIf i = LBound(a) + 7 Then
                    IPaddress = aString
                    lblLockIP.Text = IPaddress
                ElseIf i = LBound(a) + 8 Then
                    ExternalIP = aString
                    lblLockExternalIP.Text = ExternalIP
                ElseIf i = LBound(a) + 9 Then
                    Fingerprint = aString
                    lblLockFingerprint.Text = Fingerprint
                ElseIf i = LBound(a) + 10 Then
                    Memory = aString
                    lblLockMemory.Text = Memory
                ElseIf i = LBound(a) + 11 Then
                    CPUID = aString
                    lblLockCPUID.Text = CPUID
                ElseIf i = LBound(a) + 12 Then
                    BaseboardID = aString
                    lblLockBaseboardID.Text = BaseboardID
                ElseIf i = LBound(a) + 13 Then
                    VideoID = aString
                    lblLockVideoID.Text = VideoID
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
            chkLockExternalIP.Enabled = False
            chkLockFingerprint.Enabled = False
            chkLockMemory.Enabled = False
            chkLockCPUID.Enabled = False
            chkLockBaseboardID.Enabled = False
            chkLockVideoID.Enabled = False

            chkLockMACaddress.CheckState = CheckState.Unchecked
            chkLockHD.CheckState = CheckState.Unchecked
            chkLockHDfirmware.CheckState = CheckState.Unchecked
            chkLockWindows.CheckState = CheckState.Unchecked
            chkLockComputer.CheckState = CheckState.Unchecked
            chkLockBIOS.CheckState = CheckState.Unchecked
            chkLockMotherboard.CheckState = CheckState.Unchecked
            chkLockIP.CheckState = CheckState.Unchecked
            chkLockExternalIP.CheckState = CheckState.Unchecked
            chkLockFingerprint.CheckState = CheckState.Unchecked
            chkLockMemory.CheckState = CheckState.Unchecked
            chkLockCPUID.CheckState = CheckState.Unchecked
            chkLockBaseboardID.CheckState = CheckState.Unchecked
            chkLockVideoID.CheckState = CheckState.Unchecked

            For i = LBound(a) To UBound(a) - 1
                aString = a(i)
                If i = LBound(a) And aString <> noKey Then
                    MACaddress = aString
                    lblLockMacAddress.Text = MACaddress
                    chkLockMACaddress.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 1) And aString <> noKey Then
                    ComputerName = aString
                    lblLockComputer.Text = ComputerName
                    chkLockComputer.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 2) And aString <> noKey Then
                    VolumeSerial = aString
                    lblLockHD.Text = VolumeSerial
                    chkLockHD.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 3) And aString <> noKey Then
                    FirmwareSerial = aString
                    lblLockHDfirmware.Text = FirmwareSerial
                    chkLockHDfirmware.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 4) And aString <> noKey Then
                    WindowsSerial = aString
                    lblLockWindows.Text = WindowsSerial
                    chkLockWindows.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 5) And aString <> noKey Then
                    BIOSserial = aString
                    lblLockBIOS.Text = BIOSserial
                    chkLockBIOS.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 6) And aString <> noKey Then
                    MotherboardSerial = aString
                    lblLockMotherboard.Text = MotherboardSerial
                    chkLockMotherboard.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 7) And aString <> noKey Then
                    IPaddress = aString
                    lblLockIP.Text = IPaddress
                    chkLockIP.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 8) And aString <> noKey Then
                    ExternalIP = aString
                    lblLockExternalIP.Text = ExternalIP
                    chkLockExternalIP.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 9) And aString <> noKey Then
                    Fingerprint = aString
                    lblLockFingerprint.Text = Fingerprint
                    chkLockFingerprint.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 10) And aString <> noKey Then
                    Memory = aString
                    lblLockMemory.Text = Memory
                    chkLockMemory.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 11) And aString <> noKey Then
                    CPUID = aString
                    lblLockCPUID.Text = CPUID
                    chkLockCPUID.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 12) And aString <> noKey Then
                    BaseboardID = aString
                    lblLockBaseboardID.Text = BaseboardID
                    chkLockBaseboardID.CheckState = CheckState.Checked
                ElseIf i = (LBound(a) + 13) And aString <> noKey Then
                    VideoID = aString
                    lblLockVideoID.Text = VideoID
                    chkLockVideoID.CheckState = CheckState.Checked
                End If

            Next i
        End If

        If VolumeSerial = "Not Available" Or VolumeSerial = "0000-0000" Then
            chkLockHD.Enabled = False
            chkLockHD.CheckState = CheckState.Unchecked
        End If
        If MACaddress = "00 00 00 00 00 00" Or MACaddress = "00-00-00-00-00-00" Or MACaddress = "" Or MACaddress = "Not Available" Then
            chkLockMACaddress.Enabled = False
            chkLockMACaddress.CheckState = CheckState.Unchecked
        End If
        If FirmwareSerial = "Not Available" Then
            chkLockHDfirmware.Enabled = False
            chkLockHDfirmware.CheckState = CheckState.Unchecked
        End If
        If BIOSserial = "Not Available" Then
            chkLockBIOS.Enabled = False
            chkLockBIOS.CheckState = CheckState.Unchecked
        End If
        If MotherboardSerial = "Not Available" Then
            chkLockMotherboard.Enabled = False
            chkLockMotherboard.CheckState = CheckState.Unchecked
        End If
        If IPaddress = "Not Available" Then
            chkLockIP.Enabled = False
            chkLockIP.CheckState = CheckState.Unchecked
        End If
        If ExternalIP = "Not Available" Then
            chkLockExternalIP.Enabled = False
            chkLockExternalIP.CheckState = CheckState.Unchecked
        End If
        If Fingerprint = "Not Available" Then
            chkLockFingerprint.Enabled = False
            chkLockFingerprint.CheckState = CheckState.Unchecked
        End If
        If Memory = "Not Available" Then
            chkLockMemory.Enabled = False
            chkLockMemory.CheckState = CheckState.Unchecked
        End If
        If CPUID = "Not Available" Then
            chkLockCPUID.Enabled = False
            chkLockCPUID.CheckState = CheckState.Unchecked
        End If
        If BaseboardID = "Not Available" Then
            chkLockBaseboardID.Enabled = False
            chkLockBaseboardID.CheckState = CheckState.Unchecked
        End If
        If VideoID = "Not Available" Then
            chkLockVideoID.Enabled = False
            chkLockVideoID.CheckState = CheckState.Unchecked
        End If

        GetUserFromInstallCode = a(a.Length - 1)
        systemEvent = False

    End Function

    Private Sub SelectOnEnterTextBox(ByRef sender As Object)
        'select all text in a textbox
        CType(sender, TextBox).SelectionStart = 0
        CType(sender, TextBox).SelectionLength = CType(sender, TextBox).Text.Length
    End Sub

#End Region


#Region "Events"

    Friend Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load

        'load images for buttons, labels, pictureboxes
        LoadImages()

        'initialize RegisteredLevels
        strRegisteredLevelDBName = AddBackSlash(Application.StartupPath) & "RegisteredLevelDB.dat"

        ' Check the existence of necessary files to run this application
        Call CheckForResources("#Alcrypto3NET.dll")

        'load RegisteredLevels
        If Not File.Exists(strRegisteredLevelDBName) Then

            With cboRegisteredLevel
                .Items.Clear()
                .Items.Add(New Mylist("Limited A", 0))
                .Items.Add(New Mylist("Limited B", 1))
                .Items.Add(New Mylist("Limited C", 2))
                .Items.Add(New Mylist("Limited D", 3))
                .Items.Add(New Mylist("Limited E", 4))
                .Items.Add(New Mylist("No Print/Save", 5))
                .Items.Add(New Mylist("Educational A", 6))
                .Items.Add(New Mylist("Educational B", 7))
                .Items.Add(New Mylist("Educational C", 8))
                .Items.Add(New Mylist("Educational D", 9))
                .Items.Add(New Mylist("Educational E", 10))
                .Items.Add(New Mylist("Level 1", 11))
                .Items.Add(New Mylist("Level 2", 12))
                .Items.Add(New Mylist("Level 3", 13))
                .Items.Add(New Mylist("Level 4", 14))
                .Items.Add(New Mylist("Light Version", 15))
                .Items.Add(New Mylist("Pro Version", 16))
                .Items.Add(New Mylist("Enterprise Version", 17))
                .Items.Add(New Mylist("Demo Only", 18))
                .Items.Add(New Mylist("Full Version-Europe", 19))
                .Items.Add(New Mylist("Full Version-Africa", 20))
                .Items.Add(New Mylist("Full Version-Asia", 21))
                .Items.Add(New Mylist("Full Version-USA", 22))
                .Items.Add(New Mylist("Full Version-International", 23))
                .SelectedIndex = 0
                SaveComboBox(strRegisteredLevelDBName, cboRegisteredLevel.Items.GetEnumerator, True)
            End With
        Else
            LoadComboBox(strRegisteredLevelDBName, cboRegisteredLevel, True)
            'cboRegisteredLevel.SelectedIndex = 0
        End If

        'load form settings
        LoadFormSetting()

        ' Initialize AL
        InitActiveLock()

        ' Initialize GUI
        InitUI()

        'Assume that the application LockType is not LOckNone only
        txtUser.Enabled = False
        txtUser.ReadOnly = True
        txtUser.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)

        Dim sIndex As Integer = Convert.ToInt32(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboProducts", CStr(0)))
        If sIndex > cboProducts.Items.Count Then
            cboProducts.SelectedIndex = cboProducts.Items.Count
        ElseIf sIndex < 0 Then
            cboProducts.SelectedIndex = 0
        Else
            cboProducts.SelectedIndex = sIndex - 1
        End If
        txtDays.Text = ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "txtDays", CStr(365))
        Me.Text = "Alugen - ActiveLock v" & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & " Build (" & My.Application.Info.Version.Build & ") Key Generator for VB2008"
        CheckIfWizardPresent()
    End Sub
    Private Sub CheckIfWizardPresent()
        Dim fileName As String = System.Windows.Forms.Application.StartupPath & "\Activelock Wizard.exe"
        If File.Exists(fileName) Then
            cmdStartWizard.Visible = True
        Else
            cmdStartWizard.Visible = False
        End If
    End Sub
    Private Sub LoadFormSetting()
        'Read the program INI file to retrieve control settings
        On Error GoTo LoadFormSetting_Error

        If Not blnIsFirstLaunch Then Exit Sub

        On Error Resume Next
        'Read the program INI file to retrieve control settings
        PROJECT_INI_FILENAME = AppPath() & "\" & Application.ProductName & ".ini"

        SSTab1.SelectedIndex = Convert.ToInt32(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "TabNumber", CStr(0)))
        'cboProducts.SelectedIndex = Convert.ToInt32(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboProducts", CStr(0)))
        cboLicType.SelectedIndex = Convert.ToInt32(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboLicType", CStr(1)))
        cboRegisteredLevel.SelectedIndex = Convert.ToInt32(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboRegisteredLevel", CStr(0)))
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkItemData", CStr(0)) = "Unchecked" Then
            chkItemData.CheckState = CheckState.Unchecked
        Else
            chkItemData.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockBIOS", CStr(0)) = "Unchecked" Then
            chkLockBIOS.CheckState = CheckState.Unchecked
        Else
            chkLockBIOS.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockComputer", CStr(0)) = "Unchecked" Then
            chkLockComputer.CheckState = CheckState.Unchecked
        Else
            chkLockComputer.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockHD", CStr(0)) = "Unchecked" Then
            chkLockHD.CheckState = CheckState.Unchecked
        Else
            chkLockHD.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockHDfirmware", CStr(0)) = "Unchecked" Then
            chkLockHDfirmware.CheckState = CheckState.Unchecked
        Else
            chkLockHDfirmware.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockIP", CStr(0)) = "Unchecked" Then
            chkLockIP.CheckState = CheckState.Unchecked
        Else
            chkLockIP.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMACaddress", CStr(0)) = "Unchecked" Then
            chkLockMACaddress.CheckState = CheckState.Unchecked
        Else
            chkLockMACaddress.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMotherboard", CStr(0)) = "Unchecked" Then
            chkLockMotherboard.CheckState = CheckState.Unchecked
        Else
            chkLockMotherboard.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockWindows", CStr(0)) = "Unchecked" Then
            chkLockWindows.CheckState = CheckState.Unchecked
        Else
            chkLockWindows.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockExternalIP", CStr(0)) = "Unchecked" Then
            chkLockExternalIP.CheckState = CheckState.Unchecked
        Else
            chkLockExternalIP.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockFingerprint", CStr(0)) = "Unchecked" Then
            chkLockFingerprint.CheckState = CheckState.Unchecked
        Else
            chkLockFingerprint.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMemory", CStr(0)) = "Unchecked" Then
            chkLockMemory.CheckState = CheckState.Unchecked
        Else
            chkLockMemory.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockCPUID", CStr(0)) = "Unchecked" Then
            chkLockCPUID.CheckState = CheckState.Unchecked
        Else
            chkLockCPUID.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockBaseboardID", CStr(0)) = "Unchecked" Then
            chkLockBaseboardID.CheckState = CheckState.Unchecked
        Else
            chkLockBaseboardID.CheckState = CheckState.Checked
        End If
        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockVideoID", CStr(0)) = "Unchecked" Then
            chkLockVideoID.CheckState = CheckState.Unchecked
        Else
            chkLockVideoID.CheckState = CheckState.Checked
        End If

        If ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkNetworkedLicense", CStr(0)) = "Unchecked" Then
            chkNetworkedLicense.CheckState = CheckState.Unchecked
        Else
            chkNetworkedLicense.CheckState = CheckState.Checked
        End If

        txtMaxCount.Text = ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "txtMaxCount", CStr(5))
        'txtDays.Text = ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "txtDays", CStr(365))

        optStrength0.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength0", "1"))
        optStrength1.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength1", "0"))
        optStrength2.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength2", "0"))
        optStrength3.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength3", "0"))
        optStrength4.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength4", "0"))
        optStrength5.Checked = CBool(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optStrength5", "0"))

        mKeyStoreType = CType(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "KeyStoreType", CStr(1)), IActiveLock.LicStoreType)
        mProductsStoreType = CType(ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoreType", CStr(0)), IActiveLock.ProductsStoreType)
        mProductsStoragePath = ProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoragePath", modALUGEN.AppPath & "\licenses.ini")
        If Not File.Exists(mProductsStoragePath) And mProductsStoreType = IActiveLock.ProductsStoreType.alsMDBFile Then
            mProductsStoreType = IActiveLock.ProductsStoreType.alsINIFile
            mProductsStoragePath = modALUGEN.AppPath & "\licenses.ini"
        End If

        blnIsFirstLaunch = False

        On Error GoTo 0
        Exit Sub

LoadFormSetting_Error:

        MessageBox.Show("Error " & Err.Number & " (" & Err.Description & ") in procedure LoadFormSetting of Form frmMain", modALUGEN.ACTIVELOCKSTRING)

    End Sub

    Private Sub InitActiveLock()
        On Error GoTo InitForm_Error
        ActiveLock = ActiveLock3Globals_definst.NewInstance()
        ActiveLock.KeyStoreType = mKeyStoreType

        Dim MyAL As New Globals
        Dim MyGen As New AlugenGlobals

        'Use the following for ASP.NET applications
        'ActiveLock.Init(Application.StartupPath & "\bin")
        'Use the following for the VB.NET applications
        ActiveLock.Init(Application.StartupPath)

        ' Initialize Generator
        GeneratorInstance = MyGen.GeneratorInstance(mProductsStoreType)
        If File.Exists(mProductsStoragePath) = False Then
            Select Case mainForm.mProductsStoreType
                Case IActiveLock.ProductsStoreType.alsINIFile
                    mProductsStoragePath = modALUGEN.AppPath & "\licenses.ini"
                Case IActiveLock.ProductsStoreType.alsMDBFile
                    mProductsStoragePath = modALUGEN.AppPath & "\licenses.mdb"
                Case IActiveLock.ProductsStoreType.alsXMLFile
                    mProductsStoragePath = modALUGEN.AppPath & "\licenses.xml"
            End Select
        End If
        GeneratorInstance.StoragePath = mProductsStoragePath

        On Error GoTo 0
        Exit Sub

InitForm_Error:

        MessageBox.Show("Error " & Err.Number & " (" & Err.Description & ") in procedure InitForm of Form frmMain", modALUGEN.ACTIVELOCKSTRING)
    End Sub

    Private Sub frmMain_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
        'save form settings
        SaveFormSettings()
    End Sub

    Private Sub SaveFormSettings()
        'save form settings
        On Error GoTo SaveFormSettings_Error
        Dim mnReturnValue As Long
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "TabNumber", CStr(SSTab1.SelectedIndex))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboProducts", CStr(cboProducts.SelectedIndex))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboLicType", CStr(cboLicType.SelectedIndex))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "cboRegisteredLevel", CStr(cboRegisteredLevel.SelectedIndex))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkItemData", chkItemData.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkNetworkedLicense", chkNetworkedLicense.CheckState.ToString)

        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "KeyStoreType", CStr(mKeyStoreType))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoreType", CStr(mProductsStoreType))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "ProductsStoragePath", CStr(mProductsStoragePath))

        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockBIOS", chkLockBIOS.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockComputer", chkLockComputer.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockHD", chkLockHD.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockHDfirmware", chkLockHDfirmware.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockIP", chkLockIP.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMACaddress", chkLockMACaddress.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMotherboard", chkLockMotherboard.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockWindows", chkLockWindows.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockExternalIP", chkLockExternalIP.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockFingerpoint", chkLockFingerprint.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockMemory", chkLockMemory.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockCPUID", chkLockCPUID.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockBaseboardID", chkLockBaseboardID.CheckState.ToString)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "chkLockVideoID", chkLockVideoID.CheckState.ToString)

        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "txtMaxCount", txtMaxCount.Text)
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "txtDays", txtDays.Text)

        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength0", CStr(optStrength0.Checked))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength1", CStr(optStrength1.Checked))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength2", CStr(optStrength2.Checked))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength3", CStr(optStrength3.Checked))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength4", CStr(optStrength4.Checked))
        mnReturnValue = SetProfileString32(PROJECT_INI_FILENAME, "Startup Options", "optstrength5", CStr(optStrength5.Checked))

        On Error GoTo 0
        Exit Sub

SaveFormSettings_Error:

        MessageBox.Show("Error " & Err.Number & " (" & Err.Description & ") in procedure SaveFormSettings of Form frmMain", modALUGEN.ACTIVELOCKSTRING)
    End Sub

    Private Sub chkLockBIOS_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockBIOS.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockComputer_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockComputer.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockHD_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockHD.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockHDfirmware_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockHDfirmware.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockIP_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockIP.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockMACaddress_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockMACaddress.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockMotherboard_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockMotherboard.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockWindows_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockWindows.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub
    Private Sub chkLockExternalIP_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockExternalIP.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub
    Private Sub chkLockFingerpoint_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockFingerprint.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub
    Private Sub chkLockMemory_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockMemory.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub
    Private Sub chkLockCPUID_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockCPUID.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub
    Private Sub chkLockBaseboardID_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockBaseboardID.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub
    Private Sub chkLockVideoID_CheckStateChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles chkLockVideoID.CheckStateChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub cboLicType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboLicType.SelectedIndexChanged
        ' enable the days edit box
        If cboLicType.Text = "Periodic" Or cboLicType.Text = "Time Locked" Then
            txtDays.ReadOnly = False
            txtDays.BackColor = System.Drawing.ColorTranslator.FromOle(&H80000005)
            txtDays.ForeColor = System.Drawing.Color.Black
        Else
            txtDays.ReadOnly = True
            txtDays.BackColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
            txtDays.ForeColor = System.Drawing.ColorTranslator.FromOle(&H8000000F)
        End If
        If cboLicType.Text = "Time Locked" Then
            lblExpiry.Text = "&Expires on Date:"
            txtDays.Text = Date.Now.AddDays(365).ToString("yyyy/MM/dd")
            lblDays.Text = "yyyy/MM/dd"
            txtDays.Visible = False
            dtpExpireDate.Visible = True
            'walter'wrongdate'dtpExpireDate.Value = Now.UtcNow.AddDays(30)
            dtpExpireDate.Value = Date.Now.AddDays(30)
        Else
            lblExpiry.Text = "&Expires after:"
            txtDays.Text = "365"
            lblDays.Text = "Day(s)"
            txtDays.Visible = True
            dtpExpireDate.Visible = False
        End If
    End Sub

    Private Sub cboProducts_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboProducts.SelectedIndexChanged
        'product selected from products combo - update the controls
        UpdateKeyGenButtonStatus()

        If cboProducts_Array(cboProducts.SelectedIndex) = "512" Then
            lblKeyStrength.Text = "[Key Strength: CryptoAPI RSA 512-bit]"
        ElseIf cboProducts_Array(cboProducts.SelectedIndex) = "1024" Then
            lblKeyStrength.Text = "[Key Strength: CryptoAPI RSA 1024-bit]"
        ElseIf cboProducts_Array(cboProducts.SelectedIndex) = "1536" Then
            lblKeyStrength.Text = "[Key Strength: CryptoAPI RSA 1536-bit]"
        ElseIf cboProducts_Array(cboProducts.SelectedIndex) = "2048" Then
            lblKeyStrength.Text = "[Key Strength: CryptoAPI RSA 2048-bit]"
        ElseIf cboProducts_Array(cboProducts.SelectedIndex) = "4096" Then
            lblKeyStrength.Text = "[Key Strength: CryptoAPI RSA 4096-bit]"
        ElseIf cboProducts_Array(cboProducts.SelectedIndex) = "0" Then
            lblKeyStrength.Text = "[Key Strength: ALCrypto 1024-bit]"
        End If

    End Sub

    Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
        If SSTab1.SelectedIndex <> 0 Then Exit Sub ' our tab not active - do nothing
        cmdAdd.Enabled = False ' disallow repeated clicking of Add button
        AddRow(txtName.Text, txtVer.Text, txtVCode.Text, txtGCode.Text)

        UpdateStatus(String.Format("Product '{0}({1})' was added successfully.", txtName.Text, txtVer.Text))
        txtName.Focus()
        cmdValidate.Enabled = True
    End Sub
    Private Sub cmdBrowse_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdBrowse.Click
        Dim itemProductInfo As ProductInfoItem = CType(cboProducts.SelectedItem, ProductInfoItem)
        Dim strName As String = itemProductInfo.ProductName
        Dim strVersion As String = itemProductInfo.ProductVersion
        Try
            With saveDlg
                .InitialDirectory = Dir(txtLicenseFile.Text)
                .Filter = "ALL Files (*.ALL)|*.ALL"
                .FileName = strName & strVersion
                .ShowDialog()
                txtLicenseFile.Text = .FileName
            End With
        Catch ex As Exception
            MessageBox.Show(ex.Message, ACTIVELOCKSTRING, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub cmdCodeGen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCodeGen.Click
        If SSTab1.SelectedIndex <> 0 Then Exit Sub ' our tab not active - do nothing
        If txtName.Text.Contains("-") Then
            MsgBox("Dash '-' is not allowed in Product Name.", vbInformation)
            Exit Sub
        End If
        If txtVer.Text.Contains("-") Then
            MsgBox("Dash '-' is not allowed in Product Version.", vbInformation)
            Exit Sub
        End If

        Cursor = Cursors.WaitCursor
        UpdateStatus("Generating product codes! Please wait ...")
        fDisableNotifications = True
        txtVCode.Text = ""
        txtGCode.Text = ""
        fDisableNotifications = False
        Enabled = False

        Application.DoEvents()

        Try
            ' ALCrypto DLL with 1024-bit strength
            If optStrength0.Checked = True Then
                Dim KEY As RSAKey
                ReDim KEY.data(32)
                'Dim progress As ProgressType
                ' generate the key
                'VarPtr function is not supported in VB.NET
                'VB6 equivalent function is used instead - ialkan
                'Adding a delegate for AddressOf CryptoProgressUpdate did not work
                'for now. Modified alcrypto3NET.dll to create a second generate function
                'rsa_generate2 that does not deal with progress monitoring

                ' Get the current date format and save it to regionalSymbol variable
                '* Get_Locale()
                ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
                '* Set_locale("")

                If modALUGEN.rsa_generate2(KEY, 1024) = RETVAL_ON_ERROR Then
                    '* Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                End If
                ' extract private and public key blobs
                Dim strBlob As String
                Dim blobLen As Integer
                If rsa_public_key_blob(KEY, vbNullString, blobLen) = RETVAL_ON_ERROR Then
                    '* Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                End If

                If blobLen > 0 Then
                    strBlob = New String(Chr(0), blobLen)
                    If rsa_public_key_blob(KEY, strBlob, blobLen) = RETVAL_ON_ERROR Then
                        '* Set_locale(regionalSymbol)
                        Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                    End If

                    System.Diagnostics.Debug.WriteLine("Public blob: " & strBlob)
                    txtVCode.Text = strBlob
                End If

                If modALUGEN.rsa_private_key_blob(KEY, vbNullString, blobLen) = RETVAL_ON_ERROR Then
                    '* Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                End If

                If blobLen > 0 Then
                    strBlob = New String(Chr(0), blobLen)
                    If modALUGEN.rsa_private_key_blob(KEY, strBlob, blobLen) = RETVAL_ON_ERROR Then
                        '* Set_locale(regionalSymbol)
                        Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                    End If

                    System.Diagnostics.Debug.WriteLine("Private blob: " & strBlob)
                    txtGCode.Text = strBlob
                End If
                ' done with the key - throw it away
                If modALUGEN.rsa_freekey(KEY) = RETVAL_ON_ERROR Then
                    '* Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                End If

                ' Test generated key for correctness by recreating it from the blobs
                ' Note:
                ' ====
                ' Due to an outstanding bug in ALCrypto.dll, sometimes this step will crash the app because
                ' the generated keyset is bad.
                ' The work-around for the time being is to keep regenerating the keyset until eventually
                ' you'll get a valid keyset that no longer crashes.
                Dim strdata As String : strdata = "This is a test string to be encrypted."
                If modALUGEN.rsa_createkey(txtVCode.Text, txtVCode.Text.Length, txtGCode.Text, txtGCode.Text.Length, KEY) = RETVAL_ON_ERROR Then
                    '* Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                End If

                ' It worked! We're all set to go.
                If modALUGEN.rsa_freekey(KEY) = RETVAL_ON_ERROR Then
                    '* Set_locale(regionalSymbol)
                    Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
                End If

            Else  ' CryptoAPI - RSA with given strength

                Dim strPublicBlob As String, strPrivateBlob As String
                Dim imodulus As Integer

                If optStrength1.Checked = True Then
                    imodulus = 4096
                ElseIf optStrength2.Checked = True Then
                    imodulus = 2048
                ElseIf optStrength3.Checked = True Then
                    imodulus = 1536
                ElseIf optStrength4.Checked = True Then
                    imodulus = 1024
                ElseIf optStrength5.Checked = True Then
                    imodulus = 512
                End If

                'create new instance of RSACryptoServiceProvider
                'Dim cspParam As CspParameters = New CspParameters
                'cspParam.Flags = CspProviderFlags.UseMachineKeyStore
                'cspParam.KeyContainerName = txtName.Text & txtVer.Text
                'cspParam.KeyNumber = 2 'signature key pair
                ''Set the CSP Provider Type PROV_RSA_FULL
                'cspParam.ProviderType = 1
                ''Set the CSP Provider Name
                'cspParam.ProviderName = "Microsoft Base Cryptographic Provider v1.0"

                'create new instance of RSACryptoServiceProvider
                Dim rsaCSP As New System.Security.Cryptography.RSACryptoServiceProvider(imodulus)   ', cspParam)

                'Generate public and private key data and allowing their exporting.
                'True to include private parameters; otherwise, false

                Dim rsaPubParams As RSAParameters       'stores public key
                Dim rsaPrivateParams As RSAParameters   'stores private key
                rsaPrivateParams = rsaCSP.ExportParameters(True)
                rsaPubParams = rsaCSP.ExportParameters(False)

                strPrivateBlob = rsaCSP.ToXmlString(True)
                strPublicBlob = rsaCSP.ToXmlString(False)

                'ok = ActiveLock3Globals_definst.ContainerChange(txtName.Text & txtVer.Text)
                'ok = ActiveLock3Globals_definst.CryptoAPIAction(1, txtName.Text & txtVer.Text, "", "", strPublicBlob, strPrivateBlob, modulus)
                txtVCode.Text = "RSA" & imodulus & strPublicBlob
                txtGCode.Text = "RSA" & imodulus & strPrivateBlob

                rsaPubParams = Nothing
                rsaPrivateParams = Nothing
                rsaCSP = Nothing
            End If
            '* Set_locale(regionalSymbol)

        Catch ex As Exception
            '* Set_locale(regionalSymbol)
            MessageBox.Show(ex.Message, ACTIVELOCKSTRING, MessageBoxButtons.OK, MessageBoxIcon.Error)
        Finally
            'update controls
            fDisableNotifications = True
            UpdateAddButtonStatus()
            UpdateStatus("Product codes generated successfully.")
            Cursor = Cursors.Default
            Enabled = True
            fDisableNotifications = False
            '* Set_locale(regionalSymbol)
        End Try
    End Sub

    Private Sub cmdCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopy.Click
        Dim aDataObject As New DataObject
        aDataObject.SetData(DataFormats.Text, txtLicenseKey.Text)
        Clipboard.SetDataObject(aDataObject)
    End Sub
    Private Sub cmdCopyGCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopyGCode.Click
        Dim aDataObject As New DataObject
        aDataObject.SetData(DataFormats.Text, txtGCode.Text)
        Clipboard.SetDataObject(aDataObject)
    End Sub

    Private Sub cmdCopyVCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopyVCode.Click
        Dim aDataObject As New DataObject
        aDataObject.SetData(DataFormats.Text, txtVCode.Text)
        Clipboard.SetDataObject(aDataObject)
    End Sub

    ' Generate license key
    Private Sub cmdKeyGen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdKeyGen.Click
        Dim usedVCode As String = Nothing
        Dim licFlag As ActiveLock3_6NET.ProductLicense.LicFlags, maximumUsers As Short

        If txtInstallCode.Text.Length < 8 Then Exit Sub

        If SSTab1.SelectedIndex <> 1 Then Exit Sub ' our tab not active - do nothing
        ' Get the current date format and save it to regionalSymbol variable
        '* Get_Locale()
        ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
        '* Set_locale("")

        If cboLicType.Text = "Time Locked" Then
            ' Check to see if there's a valid expiration date
            If CDate(CType(dtpExpireDate.Value, DateTime).ToString("yyyy/MM/dd")) < CDate(Format(Date.Now, "yyyy/MM/dd")) Then
                '* Set_locale(regionalSymbol)
                MsgBox("Entered date occurs in the past.", vbExclamation)
                Exit Sub
            End If
        End If

        If txtInstallCode.Text.Length <> 8 Then  'Not a Short Key License
            If chkLockMACaddress.CheckState = CheckState.Unchecked _
              And chkLockComputer.CheckState = CheckState.Unchecked _
              And chkLockHD.CheckState = CheckState.Unchecked _
              And chkLockHDfirmware.CheckState = CheckState.Unchecked _
              And chkLockWindows.CheckState = CheckState.Unchecked _
              And chkLockBIOS.CheckState = CheckState.Unchecked _
              And chkLockMotherboard.CheckState = CheckState.Unchecked _
              And chkLockExternalIP.CheckState = CheckState.Unchecked _
              And chkLockFingerprint.CheckState = CheckState.Unchecked _
              And chkLockMemory.CheckState = CheckState.Unchecked _
              And chkLockCPUID.CheckState = CheckState.Unchecked _
              And chkLockBaseboardID.CheckState = CheckState.Unchecked _
              And chkLockVideoID.CheckState = CheckState.Unchecked _
              And chkLockIP.CheckState = CheckState.Unchecked Then
                MsgBox("Warning: You did not select any hardware keys to lock the license." & vbCrLf & "This license will be machine independent. License will be locked to the username only !!!", MsgBoxStyle.Exclamation)
            End If
        End If

        systemEvent = True
        If Len(txtInstallCode.Text) <> 8 Then txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False

        ' get product and version
        Cursor = Cursors.WaitCursor
        UpdateStatus("Generating license key...")

        Try
            Dim itemProductInfo As ProductInfoItem = CType(cboProducts.SelectedItem, ProductInfoItem)
            Dim strName, strVer As String
            strName = itemProductInfo.ProductName
            strVer = itemProductInfo.ProductVersion
            With ActiveLock
                .SoftwareName = strName
                .SoftwareVersion = strVer
            End With

            Dim varLicType As ProductLicense.ALLicType

            If cboLicType.Text = "Periodic" Then
                varLicType = ProductLicense.ALLicType.allicPeriodic
            ElseIf cboLicType.Text = "Permanent" Then
                varLicType = ProductLicense.ALLicType.allicPermanent
            ElseIf cboLicType.Text = "Time Locked" Then
                varLicType = ProductLicense.ALLicType.allicTimeLocked
            Else
                varLicType = ProductLicense.ALLicType.allicNone
            End If
            '* get the dates ready, first we need strings for the expiration and the registration
            Dim strExpire As String
            '* strExpire = GetExpirationDate()
            strExpire = GetExpirationString()
            Dim strRegDate As String
            strRegDate = Date.UtcNow.ToString '*("yyyy/MM/dd")
            Dim Lic As ProductLicense

            'generate license object
            Dim selRegLevel As Mylist = CType(cboRegisteredLevel.SelectedItem, Mylist)
            Dim selRelLevelType As String
            If chkItemData.CheckState = CheckState.Unchecked Then
                selRelLevelType = selRegLevel.Name
            Else
                selRelLevelType = selRegLevel.ItemData.ToString
            End If

            'Take care of the networked licenses
            If chkNetworkedLicense.CheckState = CheckState.Checked Then
                licFlag = ProductLicense.LicFlags.alfMulti
            Else
                licFlag = ProductLicense.LicFlags.alfSingle
            End If
            maximumUsers = CShort(txtMaxCount.Text)
            '* create the license object passing the expiration and registration date
            Lic = ActiveLock3Globals_definst.CreateProductLicense(strName, strVer, "", _
                      licFlag, varLicType, "", _
                      selRelLevelType, _
                      GetExpirationDate, , Date.UtcNow, , maximumUsers)
            '* strExpire, , strRegDate, , maximumUsers)

            Dim strLibKey As String, i As Integer
            If Len(txtInstallCode.Text) = 8 Then  'Short Key License
                For i = 0 To lstvwProducts.Items.Count
                    If strName = lstvwProducts.Items(i).Text And strVer = lstvwProducts.Items(i).SubItems(1).Text Then
                        usedVCode = lstvwProducts.Items(i).SubItems(2).Text
                        Exit For
                    End If
                Next
                '* this generates the key we send to the person (short key)
                '* strLibKey = ActiveLock.GenerateShortKey(usedVCode, txtInstallCode.Text, Trim(txtUser.Text), strExpire, varLicType, cboRegisteredLevel.SelectedIndex + 200, maximumUsers)
                strLibKey = ActiveLock.GenerateShortKey(usedVCode, txtInstallCode.Text, _
                Trim(txtUser.Text), GetExpirationString(), varLicType, _
                cboRegisteredLevel.SelectedIndex + 200, maximumUsers)
                txtLicenseKey.Text = strLibKey
            Else 'ALCrypto License Key
                ' Pass it to IALUGenerator to generate the key
                Dim selectedRegisteredLevel As String
                Dim mList As Mylist
                mList = CType(cboRegisteredLevel.Items(cboRegisteredLevel.SelectedIndex), Mylist)
                If chkItemData.CheckState = CheckState.Unchecked Then
                    selectedRegisteredLevel = mList.Name
                Else
                    selectedRegisteredLevel = mList.ItemData.ToString
                End If
                strLibKey = GeneratorInstance.GenKey(Lic, txtInstallCode.Text, selectedRegisteredLevel)
                'split license key into 64byte chunks
                txtLicenseKey.Text = Make64ByteChunks(strLibKey & "aLck" & txtInstallCode.Text)
                'update license file path
                txtLicenseFile.Text = Application.StartupPath & "\" & strName & strVer & ".all"
            End If

            Cursor = Cursors.Default
            UpdateStatus("Ready")

            'add license to database
            Dim lockTypesString As String
            Dim frmAlugenDatabase As New frmAlugenDb
            If MessageBox.Show("Would you like to save the new license in the License Database?", ACTIVELOCKSTRING, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1) = Windows.Forms.DialogResult.Yes Then
                lockTypesString = ""
                If chkLockMACaddress.CheckState = CheckState.Checked Then
                    lockTypesString = lockTypesString & "MAC Address"
                End If
                If chkLockComputer.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Computer Name"
                End If
                If chkLockHD.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "HDD Volume Serial"
                End If
                If chkLockHDfirmware.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "HDD Firmware Serial"
                End If
                If chkLockWindows.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Windows Serial"
                End If
                If chkLockBIOS.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "BIOS Version"
                End If
                If chkLockMotherboard.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Motherboard Serial"
                End If
                If chkLockIP.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Local IP Address"
                End If
                If chkLockExternalIP.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "External IP Address"
                End If
                If chkLockFingerprint.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Computer Fingerprint"
                End If
                If chkLockMemory.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Memory ID"
                End If
                If chkLockCPUID.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "CPU ID"
                End If
                If chkLockBaseboardID.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Baseboard ID"
                End If
                If chkLockVideoID.CheckState = CheckState.Checked Then
                    If lockTypesString <> "" Then lockTypesString = lockTypesString & "+"
                    lockTypesString = lockTypesString & "Video ID"
                End If

                'add license to database
                Call frmAlugenDatabase.ArchiveLicense(strName, strVer, txtUser.Text.Trim, strRegDate, strExpire, cboLicType.Text, lockTypesString, cboRegisteredLevel.Text, txtInstallCode.Text, txtLicenseKey.Text)

            End If
            Label5.Visible = True
            txtLicenseFile.Visible = True
            cmdBrowse.Visible = True
            cmdSave.Visible = True
            '* Set_locale(regionalSymbol)
        Catch ex As Exception
            '* Set_locale(regionalSymbol)
            UpdateStatus("Error: " & ex.Message)
        Finally
            '* Set_locale(regionalSymbol)
            Cursor = Cursors.Default
        End Try
    End Sub

    Private Sub cmdPaste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPaste.Click

        If txtInstallCode.Text.Length < 8 Then GoTo continueHere

        If txtInstallCode.Text.Substring(0, 8).ToLower = "you must" Then 'short key license
            Dim arrProdVer() As String
            arrProdVer = Split(txtInstallCode.Text, vbLf)
            systemEvent = True
            txtInstallCode.Text = (arrProdVer(1).Substring(15, 8)).Trim
            txtUser.Text = (arrProdVer(3).Substring(11, arrProdVer(3).Length - 11)).Trim
            systemEvent = False
            HandleInstallationCode()
        Else
continueHere:
            If Clipboard.GetDataObject.GetDataPresent(DataFormats.Text) Then
                txtInstallCode.Text = CType(Clipboard.GetDataObject.GetData(DataFormats.Text), String)
                UpdateKeyGenButtonStatus()
                HandleInstallationCode()
            End If
        End If
        txtLicenseKey.Text = String.Empty
    End Sub

    Private Sub cmdRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRemove.Click
        If SSTab1.SelectedIndex <> 0 Then Exit Sub ' our tab not active - do nothing
        'walter'unused'Dim SelStart As Short
        'walter'unused'Dim SelEnd As Short
        Dim fEnableAdd As Boolean
        fEnableAdd = False

        Dim strName As String = lstvwProducts.SelectedItems(0).Text
        Dim strVer As String = lstvwProducts.SelectedItems(0).SubItems(1).Text
        Dim selItem As String = strName & " - " & strVer
        'delete from INI File
        GeneratorInstance.DeleteProduct(strName, strVer)
        'remove from products list
        lstvwProducts.SelectedItems(0).Remove()

        ' Enable Add button if we're removing the variable currently being edited.
        If fEnableAdd Then
            cmdAdd.Enabled = True
        End If
        If lstvwProducts.Items.Count = 0 Then
            ' no (selectable) rows left (just the header)
            cmdRemove.Enabled = False
        End If

        For Each itemProduct As ProductInfoItem In cboProducts.Items
            If itemProduct.ProductName = strName _
            AndAlso itemProduct.ProductVersion = strVer Then
                cboProducts.Items.Remove(itemProduct)
                Exit For
            End If
        Next

        txtVCode.Text = ""
        txtGCode.Text = ""
        cmdValidate.Enabled = True

    End Sub

    Private Sub cmdSave_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdSave.Click
        If txtInstallCode.Text.Length = 8 Then
            MsgBox("ALL files are not used for Short-Key licenses.", vbInformation)
            Exit Sub
        End If

        Dim itemProductInfo As ProductInfoItem = CType(cboProducts.SelectedItem, ProductInfoItem)
        Dim strName As String = itemProductInfo.ProductName
        strName = strName.Replace("(", "").Replace(")", "").Trim
        Dim strVersion As String = itemProductInfo.ProductVersion
        If txtLicenseFile.Text.Contains(strName & strVersion & ".all") = False Then
            MsgBox("The saved ALL file name should contain " & "'" & cboProducts.Text & ".all'.", vbExclamation)
            Exit Sub
        End If
        UpdateStatus("Saving license key to file...")
        ' save the license key
        SaveLicenseKey(txtLicenseKey.Text, txtLicenseFile.Text)
        UpdateStatus("License key saved.")
    End Sub

    Private Sub cmdValidate_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdValidate.Click
        Dim KEY As RSAKey = Nothing
        Dim strdata As String
        Dim strSig As String, rc As Integer

        Cursor = Cursors.WaitCursor
        If txtVCode.Text = "" And txtGCode.Text = "" Then
            UpdateStatus("GCode and VCode fields are blank.  Nothing to validate.")
            Exit Sub ' nothing to validate
        End If

        strdata = "I love Activelock"
        'strdata = "TestApp" & vbCrLf & "3" & vbCrLf & "Single" & vbCrLf & "1" & vbCrLf & "Evaluation User" & vbCrLf & "0" & vbCrLf & "2006/11/22" & vbCrLf & "2006/12/22" & vbCrLf & "5" & vbLf & "+00 10 18 09 71 85" & vbCrLf & "MYSWEETBABY" & vbCrLf & "5CA9-4B2A" & vbCrLf & "3JT26AA0" & vbCrLf & "55274-OEM-0011903-00102" & vbCrLf & "DELL   - 7" & vbCrLf & "BFWB741" & vbCrLf & "192.168.0.1"

        ' Get the current date format and save it to regionalSymbol variable
        '* Get_Locale()
        ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
        '* Set_locale("")

        ' ALCrypto DLL with 1024-bit strength
        If strLeft(txtVCode.Text, 3) <> "RSA" Then
            ' Validate to keyset to make sure it's valid.
            UpdateStatus("Validating keyset...")
            rc = modALUGEN.rsa_createkey(txtVCode.Text, txtVCode.Text.Length, txtGCode.Text, txtGCode.Text.Length, KEY)
            If rc = RETVAL_ON_ERROR Then
                '* Set_locale(regionalSymbol)
                MessageBox.Show("Code not valid! " & vbCrLf & STRRSAERROR, ACTIVELOCKSTRING, MessageBoxButtons.OK, MessageBoxIcon.Error)
                UpdateStatus(txtName.Text & " (" & txtVer.Text & ") " & STRRSAERROR)
                Exit Sub
            End If

            ' sign it
            strSig = RSASign(txtVCode.Text, txtGCode.Text, strdata)
            rc = RSAVerify(txtVCode.Text, strdata, strSig)
            If rc = 0 Then
                UpdateStatus(txtName.Text & " (" & txtVer.Text & ") validated successfully.")
            Else
                UpdateStatus(txtName.Text & " (" & txtVer.Text & ") GCode-VCode mismatch!")
            End If
            ' It worked! We're all set to go.
            If modALUGEN.rsa_freekey(KEY) = RETVAL_ON_ERROR Then
                '* Set_locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrRSAError, ACTIVELOCKSTRING, STRRSAERROR)
            End If
        Else  '.NET RSA

            Try
                ' ------------------ begin Message from Ismail ------------------
                ' This code block is used to Encrypt/Sign and then Validate/Decrypt
                ' a small size text.
                ' This code uses "I love Activelock" to validate the given Public/Private key pair
                ' If you try to do the same with a much longer string, these routines will fail
                ' with a "Bad Length" error
                ' Increasing the cpher strength (say from 1024 to 2048-bits) will allow you to
                ' run this code with much longer data strings
                ' Activelock DLL uses a different algorithm to sign/validate
                ' This section is functional, but more than that it's provided here
                ' as the entire solution for a typical RSA signing/validation algorithm
                ' ------------------ end Message from Ismail ------------------

                Dim rsaCSP As New System.Security.Cryptography.RSACryptoServiceProvider
                Dim rsaPubParams As RSAParameters 'stores public key
                Dim strPublicBlob, strPrivateBlob As String

                strPublicBlob = txtVCode.Text
                strPrivateBlob = txtGCode.Text

                ' ENCRYPT PLAIN TEXT USING THE PUBLIC KEY
                ' Convert the data string to a byte array.
                Dim toEncrypt As Byte()     ' Holds message in bytes
                Dim enc As New UTF8Encoding ' new instance of Unicode8 instance
                Dim encrypted As Byte() ' holds encrypted data
                Dim encryptedPlainText As String

                If strLeft(txtVCode.Text, 6) = "RSA512" Then
                    strPublicBlob = strRight(txtVCode.Text, Len(txtVCode.Text) - 6)
                Else
                    strPublicBlob = strRight(txtVCode.Text, Len(txtVCode.Text) - 7)
                End If
                rsaCSP.FromXmlString(strPublicBlob)
                rsaPubParams = rsaCSP.ExportParameters(False)
                ' import public key params into instance of RSACryptoServiceProvider
                rsaCSP.ImportParameters(rsaPubParams)
                toEncrypt = enc.GetBytes(strdata)


                '' The following Encrypt method works for long and short strings
                '' =============================== ACTIVATE FOR LONG AND SHORT STRINGS =============================
                ''The RSA algorithm works on individual blocks of unencoded bytes. In this case, the
                ''maximum is 58 bytes. Therefore, we are required to break up the text into blocks and
                ''encrypt each one individually. Each encrypted block will give us an output of 128 bytes.
                ''If we do not break up the blocks in this manner, we will throw a "key not valid for use
                ''in specified state" exception

                ''Get the size of the final block
                'Const RSA_BLOCKSIZE As Integer = 58
                'Dim lastBlockLength As Integer = toEncrypt.Length Mod RSA_BLOCKSIZE
                'Dim blockCount As Integer = CType(Math.Floor(toEncrypt.Length / RSA_BLOCKSIZE), Integer) ' CType not necessary in VB 2008
                'Dim hasLastBlock As Boolean = False
                'If Not lastBlockLength.Equals(0) Then
                '    'We need to create a final block for the remaining characters
                '    blockCount += 1
                '    hasLastBlock = True
                'End If

                ''Initialize the result buffer
                'Dim result() As Byte = New Byte() {}

                ''Initialize the RSA Service Provider with the public key
                ''rsaCSP.FromXmlString(strPublicBlob) 'This was taken care of already

                ''Break the text into blocks and work on each block individually
                'For blockIndex As Integer = 0 To blockCount - 1
                '    Dim thisBlockLength As Integer

                '    'If this is the last block and we have a remainder, then set the length accordingly
                '    If blockCount.Equals(blockIndex + 1) AndAlso hasLastBlock Then
                '        thisBlockLength = lastBlockLength
                '    Else
                '        thisBlockLength = RSA_BLOCKSIZE
                '    End If
                '    Dim startChar As Integer = blockIndex * RSA_BLOCKSIZE

                '    'Define the block that we will be working on
                '    Dim currentBlock(thisBlockLength - 1) As Byte
                '    Array.Copy(toEncrypt, startChar, currentBlock, 0, thisBlockLength)

                '    'Encrypt the current block and append it to the result stream
                '    Dim encryptedBlock() As Byte = rsaCSP.Encrypt(currentBlock, False)
                '    Dim originalResultLength As Integer = result.Length

                '    ReDim Preserve result(originalResultLength + encryptedBlock.Length) ' This is for VB 2008
                '    'Array.Resize(result, originalResultLength + encryptedBlock.Length)

                '    encryptedBlock.CopyTo(result, originalResultLength)
                'Next

                'encrypted = result
                ' =============================================================================================

                ' The following Encrypt method works only for short strings
                encrypted = rsaCSP.Encrypt(toEncrypt, False)
                encryptedPlainText = Convert.ToBase64String(encrypted) ' convert to base64/Radix output

                ' HASH AND SIGN THE SIGNATURE
                ' GENERATE SIGNATURE BLOCK USING SENDER'S PRIVATE KEY
                Dim signatureBlock As String
                ' Hash the encrypted data and generate a signature block on the hash
                ' using the sender's private key. (Signature Block)
                ' create new instance of SHA1 hash algorithm to compute hash
                Dim hash As New SHA1Managed
                Dim hashedData() As Byte ' a byte array to store hash value
                If strLeft(txtGCode.Text, 6) = "RSA512" Then
                    strPrivateBlob = strRight(txtGCode.Text, Len(txtGCode.Text) - 6)
                Else
                    strPrivateBlob = strRight(txtGCode.Text, Len(txtGCode.Text) - 7)
                End If
                ' import private key params into instance of RSACryptoServiceProvider
                rsaCSP.FromXmlString(strPrivateBlob)
                Dim rsaPrivateParams As RSAParameters 'stores private key
                rsaPrivateParams = rsaCSP.ExportParameters(True)
                rsaCSP.ImportParameters(rsaPrivateParams)
                ' compute hash with algorithm specified as here we have SHA1
                hashedData = hash.ComputeHash(encrypted)
                ' Sign Data using private key and  OID is simple name of the algorithm for which to get the object identifier (OID)
                Dim signature As Byte() ' holds signatures
                signature = rsaCSP.SignHash(hashedData, CryptoConfig.MapNameToOID("SHA1"))
                ' convert to base64/Radix output
                signatureBlock = Convert.ToBase64String(signature)

                ' VERIFY SIGNATURE BLOCK USING THE SENDER'S PUBLIC KEY
                ' VALIDATE THE STRING WITH THE PUBLIC/PRIVATE KEY PAIR
                ' Verify the signature is authentic using the sender's public key(decrypt Signature block)
                Dim myencrypted() As Byte
                Dim mysignature() As Byte
                myencrypted = Convert.FromBase64String(encryptedPlainText)
                mysignature = Convert.FromBase64String(signatureBlock)
                ' create new instance of SHA1 hash algorithm to compute hash
                Dim sha1hash As New SHA1Managed
                Dim sha1hashedData() As Byte ' a byte array to store hash value
                ' import  public key params into instance of RSACryptoServiceProvider
                rsaCSP.ImportParameters(rsaPubParams)
                ' compute hash with algorithm specified as here we have SHA1
                sha1hashedData = sha1hash.ComputeHash(myencrypted)
                ' Sign Data using public key and  OID is simple name of the algorithm for which to get the object identifier (OID)
                Dim verified As Boolean
                verified = rsaCSP.VerifyHash(sha1hashedData, CryptoConfig.MapNameToOID("SHA1"), mysignature)
                If verified Then
                    UpdateStatus(txtName.Text & " (" & txtVer.Text & ") validated successfully.")
                    'MsgBox("Signature Valid", MsgBoxStyle.Information)
                Else
                    UpdateStatus(txtName.Text & " (" & txtVer.Text & ") GCode-VCode mismatch!")
                    'MsgBox("Invalid Signature", MsgBoxStyle.Exclamation)
                End If

                ' THE FOLLOWING CODE BLOCK IS USED TO RETRIEVE THE ORIGINAL
                ' STRING strData BUT IS NOT NEEDED FOR THE VALIDATION PROCESS
                ' IT'S BEEN SHOWN HERE FOR DEMONSTRATION PURPOSES
                ' This works for short strings only
                Dim newencrypted() As Byte
                newencrypted = Convert.FromBase64String(encryptedPlainText)
                Dim fromEncrypt() As Byte ' a byte array to store decrypted bytes
                Dim roundTrip As String ' holds original message
                ' import  private key params into instance of RSACryptoServiceProvider
                rsaCSP.ImportParameters(rsaPrivateParams)


                '' The following Decrypt method works for long and short strings
                '' It's currently not functioning correctly
                '' =============================== ACTIVATE FOR LONG AND SHORT STRINGS =============================
                ''When we encrypt a string using RSA, it works on individual blocks of up to
                ''58 bytes. Each block generates an output of 128 encrypted bytes. Therefore, to
                ''decrypt the message, we need to break the encrypted stream into individual
                ''chunks of 128 bytes and decrypt them individually

                ''Determine how many bytes are in the encrypted stream. The input is in hex format,
                ''so we have to divide it by 2
                'Const RSA_DECRYPTBLOCKSIZE As Integer = 128
                'Dim maxBytes As Integer = CType(encryptedPlainText.Length / 2, Integer)  ' CType not necessary in VB 2008

                ''Ensure that the length of the encrypted stream is divisible by 128
                'If Not (maxBytes Mod RSA_DECRYPTBLOCKSIZE).Equals(0) Then
                '    Throw New System.Security.Cryptography.CryptographicException("Encrypted text is an invalid length")
                'End If

                ''Calculate the number of blocks we will have to work on
                'Dim blockCount2 As Integer = CType(maxBytes / RSA_DECRYPTBLOCKSIZE, Integer)

                ''Initialize the result buffer
                'Dim result2() As Byte = New Byte() {}

                ''rsaCSP.FromXmlString(strPrivateBlob) ' This was done already

                ''Iterate through each block and decrypt it
                'For blockIndex As Integer = 0 To blockCount2 - 1
                '    'Get the current block to work on
                '    Dim currentBlockHex As String = encryptedPlainText.Substring(blockIndex * (RSA_DECRYPTBLOCKSIZE * 2), RSA_DECRYPTBLOCKSIZE * 2)
                '    Dim currentBlockBytes As Byte() = HexToBytes(currentBlockHex)

                '    'Decrypt the current block and append it to the result stream
                '    Dim currentBlockDecrypted() As Byte = rsaCSP.Decrypt(currentBlockBytes, False)
                '    Dim originalResultLength As Integer = result2.Length

                '    ReDim Preserve result2(originalResultLength + currentBlockDecrypted.Length)
                '    'Array.Resize(result, originalResultLength + currentBlockDecrypted.Length) ' This is for VB 2008

                '    currentBlockDecrypted.CopyTo(result2, originalResultLength)
                'Next
                'fromEncrypt = result2
                ' =============================================================================================


                ' The following Decrypt works for short strings only
                'store decrypted data into byte array
                fromEncrypt = rsaCSP.Decrypt(newencrypted, False)

                'convert bytes to string
                roundTrip = enc.GetString(fromEncrypt)
                If roundTrip <> strdata Then
                    UpdateStatus(txtName.Text & " (" & txtVer.Text & ") GCode-VCode mismatch!")
                End If

                'Release any resources held by the RSA Service Provider
                rsaCSP.Clear()
                '* Set_locale(regionalSymbol)

            Catch ex As Exception
                '* Set_locale(regionalSymbol)
                UpdateStatus(ex.Message)
            End Try

        End If

        Cursor = Cursors.Default
        Exit Sub

exitValidate:
        '* Set_locale(regionalSymbol)
        UpdateStatus(txtName.Text & " (" & txtVer.Text & ") GCode-VCode mismatch!")
        Cursor = Cursors.Default
    End Sub
    '********************************************************
    '* HexToBytes: Converts a hex-encoded string to a
    '*             byte array
    '********************************************************
    Private Shared Function HexToBytes(ByVal Hex As String) As Byte()
        Dim numBytes As Integer = CType(Hex.Length / 2, Integer)  ' CType not necessary in VB 2008
        Dim bytes(numBytes - 1) As Byte
        For n As Integer = 0 To numBytes - 1
            Dim hexByte As String = Hex.Substring(n * 2, 2)
            bytes(n) = CType(Integer.Parse(hexByte, Globalization.NumberStyles.HexNumber), Byte) ' CType not necessary with VB 2008
        Next
        Return bytes
    End Function
    '********************************************************
    '* BytesToHex: Converts a byte array to a hex-encoded
    '*             string
    '********************************************************
    Private Shared Function BytesToHex(ByVal bytes() As Byte) As String
        Dim hex As New StringBuilder
        For n As Integer = 0 To bytes.Length - 1
            hex.AppendFormat("{0:X2}", bytes(n))
        Next
        Return hex.ToString
    End Function

    Private Sub cmdViewArchive_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdViewArchive.Click
        Dim lic As New frmAlugenDb
        lic.Show()
    End Sub

    Private Sub cmdViewLevel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdViewLevel.Click
        Dim mfrmLevelManager As New frmLevelManager
        With mfrmLevelManager
            .ShowDialog()
            cboRegisteredLevel.Items.Clear()
            LoadComboBox(strRegisteredLevelDBName, cboRegisteredLevel, True)
            cboRegisteredLevel.SelectedIndex = 0
        End With
        mfrmLevelManager.Close()
        mfrmLevelManager = Nothing
    End Sub

    Private Sub lstvwProducts_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lstvwProducts.SelectedIndexChanged
        If lstvwProducts.SelectedItems.Count > 0 Then
            Dim selListItem As ListViewItem = lstvwProducts.SelectedItems(0)
            txtName.Text = selListItem.Text
            txtVer.Text = selListItem.SubItems(1).Text

            txtVCode.Text = selListItem.SubItems(2).Text
            txtGCode.Text = selListItem.SubItems(3).Text

            cmdRemove.Enabled = True
            cmdValidate.Enabled = True
        End If
    End Sub

    Private Sub lstvwProducts_ColumnClick(ByVal sender As Object, ByVal e As ColumnClickEventArgs) Handles lstvwProducts.ColumnClick
        'Determine if the clicked column is already the column that is 
        ' being sorted.
        If (e.Column = lvwColumnSorter.SortColumn) Then
            ' Reverse the current sort direction for this column.
            If (lvwColumnSorter.Order = SortOrder.Ascending) Then
                lvwColumnSorter.Order = SortOrder.Descending
            Else
                lvwColumnSorter.Order = SortOrder.Ascending
            End If
        Else
            ' Set the column number that is to be sorted; default to ascending.
            lvwColumnSorter.SortColumn = e.Column
            lvwColumnSorter.Order = SortOrder.Ascending
        End If

        ' Perform the sort with these new sort options.
        Me.lstvwProducts.Sort()
    End Sub

    Private Sub picALBanner_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picALBanner.Click
        'navigate to www.activelocksoftware.com
        System.Diagnostics.Process.Start(ACTIVELOCKSOFTWAREWEB)
    End Sub

    Private Sub picALBanner2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picALBanner2.Click
        'navigate to www.activelocksoftware.com
        picALBanner_Click(sender, e)
    End Sub

    Private Sub txtLicenseKey_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLicenseKey.TextChanged
        cmdSave.Enabled = CBool(txtLicenseKey.Text.Length > 0)
    End Sub

    Private Sub txtName_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtName.KeyPress
        If Char.IsControl(e.KeyChar) = False Then
            If Char.IsLetter(e.KeyChar) = False And Char.IsNumber(e.KeyChar) = False Then
                If e.KeyChar <> "." And e.KeyChar <> "_" And e.KeyChar <> " " Then
                    e.Handled = True
                End If
            Else
            End If
        End If
    End Sub

    Private Sub txtName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtName.TextChanged
        fDisableNotifications = False
        UpdateCodeGenButtonStatus()
        UpdateAddButtonStatus()
    End Sub

    Private Sub txtVer_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtVer.KeyPress
        If Char.IsControl(e.KeyChar) = False Then
            If Char.IsLetter(e.KeyChar) = False And Char.IsNumber(e.KeyChar) = False Then
                If e.KeyChar <> "." And e.KeyChar <> "_" And e.KeyChar <> " " Then
                    e.Handled = True
                End If
            Else
            End If
        End If
    End Sub

    Private Sub txtVer_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVer.TextChanged
        ' New product - will be processed by Add command
        fDisableNotifications = False
        UpdateCodeGenButtonStatus()
        UpdateAddButtonStatus()
    End Sub

    Private Sub txtUser_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtUser.TextChanged
        If systemEvent Then Exit Sub
        If fDisableNotifications Then Exit Sub
        fDisableNotifications = True
        If Len(txtInstallCode.Text) <> 8 Then txtInstallCode.Text = ActiveLock.InstallationCode(Trim(txtUser.Text))
        fDisableNotifications = False
    End Sub

    Private Sub txtInstallCode_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtInstallCode.TextChanged
        HandleInstallationCode()
    End Sub

    Private Sub txtDays_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDays.Enter
        SelectOnEnterTextBox(sender)
    End Sub

    Private Sub txtInstallCode_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtInstallCode.Enter
        SelectOnEnterTextBox(sender)
    End Sub

    Private Sub txtVCode_Enter(ByVal sender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVCode.Enter
        SelectOnEnterTextBox(sender)
    End Sub

    Private Sub txtGCode_Enter(ByVal sender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtGCode.Enter
        SelectOnEnterTextBox(sender)
    End Sub

    Private Sub txtLicenseKey_Enter(ByVal sender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLicenseKey.Enter
        SelectOnEnterTextBox(sender)
    End Sub

    Private Sub txtLicenseFile_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLicenseFile.Enter
        SelectOnEnterTextBox(sender)
    End Sub

    Private Sub txtName_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtName.Enter
        SelectOnEnterTextBox(sender)
    End Sub

    Private Sub txtVer_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtVer.Enter
        SelectOnEnterTextBox(sender)
    End Sub

    Private Sub lnkActivelockSoftwareGroup_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkActivelockSoftwareGroup.LinkClicked
        picALBanner_Click(sender, e)
    End Sub

    Private Sub cmdPrintLicenseKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrintLicenseKey.Click
        Dim daPrintDocument As New daReport.DaPrintDocument
        Dim hashParameters As New Hashtable
        Dim selProduct As ProductInfoItem = CType(cboProducts.SelectedItem, ProductInfoItem)
        Dim itemRegLevel As Mylist = CType(cboRegisteredLevel.SelectedItem, Mylist)

        'set .xml file for printing
        daPrintDocument.setXML("reports\repLicenseKey.xml")

        'build parameters
        hashParameters.Add("pProductName", selProduct.ProductName)
        hashParameters.Add("pProductVersion", selProduct.ProductVersion)
        hashParameters.Add("pRegisteredLevel", itemRegLevel.Name)
        hashParameters.Add("pLicenseType", CType(cboLicType.SelectedItem, String))
        hashParameters.Add("pRegisteredDate", "")
        hashParameters.Add("pExpireDate", "")
        hashParameters.Add("pInstallCode", txtInstallCode.Text)
        hashParameters.Add("pUserName", txtUser.Text)
        hashParameters.Add("pLicenseKey", txtLicenseKey.Text)

        'setting parameters
        daPrintDocument.SetParameters(hashParameters)
        'daPrintDocument.DocumentName = "License key"

        'print preview
        printPreviewDialog1.Icon = CType(resxList("report_ico"), Icon)
        printPreviewDialog1.Text = daPrintDocument.DocumentName
        printPreviewDialog1.Document = daPrintDocument
        printPreviewDialog1.PrintPreviewControl.Zoom = 1.0
        printPreviewDialog1.WindowState = FormWindowState.Maximized
        printPreviewDialog1.ShowDialog(Me)
    End Sub

    Private Sub cmdEmailLicenseKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdEmailLicenseKey.Click
        Dim mailToString As String
        Dim emailAddress As String = "user@company.com"
        Dim strSubject As String
        Dim strBodyMessage As String
        Dim selProduct As ProductInfoItem = CType(cboProducts.SelectedItem, ProductInfoItem)
        Dim itemRegLevel As Mylist = CType(cboRegisteredLevel.SelectedItem, Mylist)
        Dim strNewLine As String = "%0D%0A"

        strSubject = String.Format("License key for application {0} ({1}), user [{2}]", selProduct.ProductName, selProduct.ProductVersion, txtUser.Text)
        strBodyMessage = strNewLine & String.Format("Install code:" & strNewLine & "{0}", txtInstallCode.Text)
        strBodyMessage = strBodyMessage & strNewLine & strNewLine & String.Format("License key:" & strNewLine & "{0}", txtLicenseKey.Text.Replace(Chr(13), strNewLine))

        'final constructor
        mailToString = String.Format("mailto:{0}?subject={1}&body={2}", emailAddress, strSubject, strBodyMessage)

        'launch default email client
        System.Diagnostics.Process.Start(mailToString)
    End Sub
#End Region

    Private Sub cmdProductsStorage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdProductsStorage.Click
        Dim myProductsStorageForm As New frmProductsStorage
        myProductsStorageForm.ShowDialog(Me)
    End Sub

    Private Sub chkNetworkedLicense_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkNetworkedLicense.CheckedChanged
        If chkNetworkedLicense.CheckState = CheckState.Checked Then
            lblConcurrentUsers.Visible = True
            txtMaxCount.Visible = True
        Else
            lblConcurrentUsers.Visible = False
            txtMaxCount.Visible = False
        End If

    End Sub

    Private Sub cmdValidate2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdValidate2.Click

        ' ------------------ begin Message from Ismail ------------------
        ' This code block is used to Sign and then Validate any size text
        ' If you try to do the same with the typical RSA and with long strings, 
        ' routines will fail with a "Bad Length" error
        ' I am providing this sample here to show a second type of sign/verify scheme
        ' Although RSA is usually intended for signing/verifying short keys,
        ' the routine below with sign/verify any length string
        ' This 2nd type of validation button is hidden and is intended for 
        ' developers to test and learn.
        ' Note that no facility exists to retrieve the original data.
        ' Similar code can be found under SourceForge in project NCrypto
        ' ------------------ end Message from Ismail ------------------

        ' Get the current date format and save it to regionalSymbol variable
        '* Get_Locale()
        ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
        '* Set_locale("")

        If strLeft(txtVCode.Text, 3) = "RSA" Then

            Try
                Dim strData As String
                strData = "TestApp" & vbCrLf & "3" & vbCrLf & "Single" & vbCrLf & "1" & vbCrLf & "Evaluation User" & vbCrLf & "0" & vbCrLf & "2006/11/22" & vbCrLf & "2006/12/22" & vbCrLf & "5" & vbLf & "+00 10 18 09 71 85" & vbCrLf & "MYSWEETBABY" & vbCrLf & "5CA9-4B2A" & vbCrLf & "3JT26AA0" & vbCrLf & "55274-OEM-0011903-00102" & vbCrLf & "DELL   - 7" & vbCrLf & "BFWB741" & vbCrLf & "192.168.0.1"
                Dim rsaCSP As New System.Security.Cryptography.RSACryptoServiceProvider
                Dim rsaPubParams As RSAParameters 'stores public key
                Dim strPublicBlob, strPrivateBlob As String

                strPublicBlob = txtVCode.Text
                strPrivateBlob = txtGCode.Text

                If strLeft(txtGCode.Text, 6) = "RSA512" Then
                    strPrivateBlob = strRight(txtGCode.Text, Len(txtGCode.Text) - 6)
                Else
                    strPrivateBlob = strRight(txtGCode.Text, Len(txtGCode.Text) - 7)
                End If
                ' import private key params into instance of RSACryptoServiceProvider
                rsaCSP.FromXmlString(strPrivateBlob)
                Dim rsaPrivateParams As RSAParameters 'stores private key
                rsaPrivateParams = rsaCSP.ExportParameters(True)
                rsaCSP.ImportParameters(rsaPrivateParams)

                Dim userData As Byte() = Encoding.UTF8.GetBytes(strData)
                Dim asf As AsymmetricSignatureFormatter = New RSAPKCS1SignatureFormatter(rsaCSP)
                Dim algorithm As HashAlgorithm = New SHA1Managed
                asf.SetHashAlgorithm(algorithm.ToString)
                Dim myhashedData() As Byte ' a byte array to store hash value
                Dim myhashedDataString As String
                myhashedData = algorithm.ComputeHash(userData)
                myhashedDataString = BitConverter.ToString(myhashedData).Replace("-", String.Empty)
                Dim mysignature As Byte() ' holds signatures
                mysignature = asf.CreateSignature(algorithm)
                Dim mySignatureBlock As String
                mySignatureBlock = Convert.ToBase64String(mysignature)

                ' Verify Signature
                If strLeft(txtVCode.Text, 6) = "RSA512" Then
                    strPublicBlob = strRight(txtVCode.Text, Len(txtVCode.Text) - 6)
                Else
                    strPublicBlob = strRight(txtVCode.Text, Len(txtVCode.Text) - 7)
                End If
                rsaCSP.FromXmlString(strPublicBlob)
                rsaPubParams = rsaCSP.ExportParameters(False)
                ' import public key params into instance of RSACryptoServiceProvider
                rsaCSP.ImportParameters(rsaPubParams)

                ' Also could use the following to check if the string is a base64 string
                If ExpBase64.IsMatch(mySignatureBlock) Then
                End If

                Dim newsignature() As Byte
                newsignature = Convert.FromBase64String(mySignatureBlock)
                Dim asd As AsymmetricSignatureDeformatter = New RSAPKCS1SignatureDeformatter(rsaCSP)
                asd.SetHashAlgorithm(algorithm.ToString)
                Dim newhashedData() As Byte ' a byte array to store hash value
                Dim newhashedDataString As String
                newhashedData = algorithm.ComputeHash(userData)
                newhashedDataString = BitConverter.ToString(newhashedData).Replace("-", String.Empty)
                Dim verified As Boolean
                verified = asd.VerifySignature(algorithm, newsignature)
                If verified Then
                    UpdateStatus(txtName.Text & " (" & txtVer.Text & ") validated successfully.")
                    'MsgBox("Signature Valid", MsgBoxStyle.Information)
                Else
                    UpdateStatus(txtName.Text & " (" & txtVer.Text & ") GCode-VCode mismatch!")
                    'MsgBox("Invalid Signature", MsgBoxStyle.Exclamation)
                End If

                'Release any resources held by the RSA Service Provider
                rsaCSP.Clear()
                '* Set_locale(regionalSymbol)

            Catch ex As Exception
                '* Set_locale(regionalSymbol)
                UpdateStatus(ex.Message)
            End Try
        End If

    End Sub

    Private Sub chkLockIP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockIP.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
        If chkLockIP.Checked Then
            MsgBox("Warning: Use Local IP addresses cautiously since they may not be static.", MsgBoxStyle.Exclamation, "Static IP Address Warning")
        End If
    End Sub
    Private Sub chkLockexternalIP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockExternalIP.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
        If chkLockExternalIP.Checked Then
            MsgBox("Warning: Use External IP addresses cautiously since they may not be static.", MsgBoxStyle.Exclamation, "Static IP Address Warning")
        End If
    End Sub

    Private Sub cmdCheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCheckAll.Click
        chkLockMACaddress.CheckState = CheckState.Checked
        chkLockComputer.CheckState = CheckState.Checked
        chkLockHD.CheckState = CheckState.Checked
        chkLockHDfirmware.CheckState = CheckState.Checked
        chkLockWindows.CheckState = CheckState.Checked
        chkLockBIOS.CheckState = CheckState.Checked
        chkLockMotherboard.CheckState = CheckState.Checked
        chkLockIP.CheckState = CheckState.Checked
        chkLockExternalIP.CheckState = CheckState.Checked
        chkLockFingerprint.CheckState = CheckState.Checked
        chkLockMemory.CheckState = CheckState.Checked
        chkLockCPUID.CheckState = CheckState.Checked
        chkLockBaseboardID.CheckState = CheckState.Checked
        chkLockVideoID.CheckState = CheckState.Checked
    End Sub

    Private Sub cmdUncheckAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUncheckAll.Click
        chkLockMACaddress.CheckState = CheckState.Unchecked
        chkLockComputer.CheckState = CheckState.Unchecked
        chkLockHD.CheckState = CheckState.Unchecked
        chkLockHDfirmware.CheckState = CheckState.Unchecked
        chkLockWindows.CheckState = CheckState.Unchecked
        chkLockBIOS.CheckState = CheckState.Unchecked
        chkLockMotherboard.CheckState = CheckState.Unchecked
        chkLockIP.CheckState = CheckState.Unchecked
        chkLockExternalIP.CheckState = CheckState.Unchecked
        chkLockFingerprint.CheckState = CheckState.Unchecked
        chkLockMemory.CheckState = CheckState.Unchecked
        chkLockCPUID.CheckState = CheckState.Unchecked
        chkLockBaseboardID.CheckState = CheckState.Unchecked
        chkLockVideoID.CheckState = CheckState.Unchecked
    End Sub

    Private Sub chkLockMACaddress_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockMACaddress.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockComputer_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockComputer.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockHD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockHD.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockHDfirmware_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockHDfirmware.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockWindows_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockWindows.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockBIOS_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockBIOS.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockMotherboard_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockMotherboard.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub cmdStartWizard_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStartWizard.Click
        Dim Arguments As String = Nothing
        Try
            Arguments = Chr(34) & "AppName=" & txtName.Text & Chr(34)
            Arguments = Arguments & " "
            Arguments = Arguments & Chr(34) & "AppVersion=" & txtVer.Text & Chr(34)
            Arguments = Arguments & " "
            Arguments = Arguments & Chr(34) & "PUB_KEY=" & txtVCode.Text & Chr(34)
            Shell("Activelock Wizard.exe " & Arguments, AppWinStyle.NormalFocus) 'Make Sure The Wizard Is the Alugen Directory
        Catch ex As Exception
            cmdStartWizard.Visible = False
        End Try

    End Sub

    Private Sub chkLockVideoID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockVideoID.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockMemory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockMemory.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockCPUID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockCPUID.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockBaseboardID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockBaseboardID.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub

    Private Sub chkLockFingerprint_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLockFingerprint.CheckedChanged
        If systemEvent Then Exit Sub
        systemEvent = True
        txtInstallCode.Text = ReconstructedInstallationCode()
        systemEvent = False
    End Sub
    Public Sub HandleInstallationCode()
        Dim installNameandVersion As String
        Dim i As Integer, success As Boolean

        If systemEvent Then Exit Sub
        If txtInstallCode.Text.Length < 8 Then Exit Sub

        If txtInstallCode.Text.Substring(0, 8).ToLower = "you must" Then 'short key license
            Dim arrProdVer() As String
            arrProdVer = Split(txtInstallCode.Text, vbLf)
            systemEvent = True
            txtInstallCode.Text = (arrProdVer(1).Substring(15, 8)).Trim
            txtUser.Text = (arrProdVer(3).Substring(11, arrProdVer(3).Length - 11)).Trim
            systemEvent = False
        End If

        If txtInstallCode.Text.Length = 8 Then 'Short key authorization is much simpler
            UpdateKeyGenButtonStatus()
            If fDisableNotifications Then Exit Sub

            chkLockMACaddress.Visible = False
            chkLockComputer.Visible = False
            chkLockHD.Visible = False
            chkLockHDfirmware.Visible = False
            chkLockWindows.Visible = False
            chkLockBIOS.Visible = False
            chkLockMotherboard.Visible = False
            chkLockIP.Visible = False
            chkLockExternalIP.Visible = False
            chkLockFingerprint.Visible = False
            chkLockMemory.Visible = False
            chkLockCPUID.Visible = False
            chkLockBaseboardID.Visible = False
            chkLockVideoID.Visible = False

            'txtUser.Text = ""
            txtUser.Enabled = True
            txtUser.ReadOnly = False
            txtUser.BackColor = Color.White

            Label5.Visible = False
            txtLicenseFile.Visible = False
            cmdBrowse.Visible = False
            cmdSave.Visible = False
            Exit Sub

        Else 'ALCrypto

            chkLockMACaddress.Visible = True
            chkLockComputer.Visible = True
            chkLockHD.Visible = True
            chkLockHDfirmware.Visible = True
            chkLockWindows.Visible = True
            chkLockBIOS.Visible = True
            chkLockMotherboard.Visible = True
            chkLockIP.Visible = True
            chkLockExternalIP.Visible = True
            chkLockFingerprint.Visible = True
            chkLockMemory.Visible = True
            chkLockCPUID.Visible = True
            chkLockBaseboardID.Visible = True
            chkLockVideoID.Visible = True
            txtUser.Enabled = False
            txtUser.ReadOnly = True
            txtUser.BackColor = System.Drawing.SystemColors.Control

            Label5.Visible = True
            txtLicenseFile.Visible = True
            cmdBrowse.Visible = True
            cmdSave.Visible = True

            If Len(txtInstallCode.Text) > 0 Then
                If systemEvent Then Exit Sub
                UpdateKeyGenButtonStatus()
                If fDisableNotifications Then Exit Sub

                fDisableNotifications = True
                txtUser.Text = GetUserFromInstallCode(txtInstallCode.Text)
                fDisableNotifications = False

                installNameandVersion = GetUserSoftwareNameandVersionfromInstallCode(txtInstallCode.Text)
                For i = 0 To cboProducts.Items.Count - 1
                    cboProducts.SelectedIndex = i
                    If installNameandVersion = cboProducts.Text Then
                        success = True
                        Exit For
                    End If
                Next i
                If Not success Then
                    MsgBox("There's no matching Software Name and Version Number for this Installation Code.", MsgBoxStyle.Exclamation)
                    cmdKeyGen.Enabled = False
                End If
            Else
                fDisableNotifications = True
                chkLockComputer.Enabled = True
                chkLockComputer.Text = "Lock to Computer Name"
                chkLockHD.Enabled = True
                chkLockHD.Text = "Lock to HDD Volume Serial"
                chkLockHDfirmware.Enabled = True
                chkLockHDfirmware.Text = "Lock to HDD Firmware Serial"
                chkLockMACaddress.Enabled = True
                chkLockMACaddress.Text = "Lock to MAC Address"
                chkLockWindows.Enabled = True
                chkLockWindows.Text = "Lock to Windows Serial"
                chkLockBIOS.Enabled = True
                chkLockBIOS.Text = "Lock to BIOS Version"
                chkLockMotherboard.Enabled = True
                chkLockMotherboard.Text = "Lock to Motherboard Serial"
                chkLockIP.Enabled = True
                chkLockIP.Text = "Lock to Local IP Address"
                chkLockExternalIP.Enabled = True
                chkLockExternalIP.Text = "Lock to External IP Address"
                chkLockFingerprint.Enabled = True
                chkLockFingerprint.Text = "Lock to Computer Fingerprint [VB.NET]"
                chkLockMemory.Enabled = True
                chkLockMemory.Text = "Lock to Memory"
                chkLockCPUID.Enabled = True
                chkLockCPUID.Text = "Lock to CPU ID"
                chkLockBaseboardID.Enabled = True
                chkLockBaseboardID.Text = "Lock to Baseboard ID"
                chkLockVideoID.Enabled = True
                chkLockVideoID.Text = "Lock to Video Controller ID"
                txtUser.Text = ""
                fDisableNotifications = False
            End If
        End If

    End Sub
End Class