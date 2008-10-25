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
' ****************************************
' * Utility per la gestione del database *
' * dove memorizzare le licenze generate *
' *      Salvatore La Porta              *
' *     mjlapo23@hotmail.com             *
' ****************************************
Option Strict On
Option Explicit On

Imports System.Data
Imports System.Data.OleDb
Imports System.Windows.Forms
Imports System.Reflection

Public Class frmAlugenDb
  Inherits System.Windows.Forms.Form

#Region "Private members"
  Private conn As OleDbConnection
    Private stringaconn As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=licenses.mdb"
  Private LicensesTableMapping As String = "Licenses"
  Private DefaultQueryTemplate As String = "SELECT * FROM License WHERE LockType Like @locktype AND username Like @username AND progname Like @progname " & _
                    " AND progver Like @progver AND LicType Like @selcboLicType AND RegLevel Like @selcboRegisteredLevel"

  Private gridHelper As GridLayoutHelper
  Private printPreviewDialog1 As New PrintPreviewDialog
#End Region

#Region "Windows Form Designer generated code "

  Public Sub New()
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
  End Sub

  'Form overrides dispose to clean up the component list.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Required by the Windows Form Designer
  Private components As System.ComponentModel.IContainer
  Friend WithEvents menudg As System.Windows.Forms.ContextMenu
  Friend WithEvents menudg_delete As System.Windows.Forms.MenuItem
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents Label5 As System.Windows.Forms.Label
  Friend WithEvents txtusername As System.Windows.Forms.TextBox
  Friend WithEvents txtlocktype As System.Windows.Forms.TextBox
  Friend WithEvents Label6 As System.Windows.Forms.Label
  Friend WithEvents Label15 As System.Windows.Forms.Label
  Friend WithEvents Label3 As System.Windows.Forms.Label
  Friend WithEvents Label4 As System.Windows.Forms.Label
  Friend WithEvents Label9 As System.Windows.Forms.Label
  Friend WithEvents menudg_dettagli As System.Windows.Forms.MenuItem
  Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
  Friend WithEvents cmdReset As System.Windows.Forms.Button
  Friend WithEvents Label8 As System.Windows.Forms.Label
  Friend WithEvents cmdFilter As System.Windows.Forms.Button
  Friend WithEvents dtpRegDateFrom As NullableDateTimePicker
  Friend WithEvents dtpRegDateTo As NullableDateTimePicker
  Friend WithEvents dtpExpDateFrom As NullableDateTimePicker
  Friend WithEvents dtpExpDateTo As NullableDateTimePicker
  Friend WithEvents cboProductName As System.Windows.Forms.ComboBox
  Friend WithEvents cboLicType As System.Windows.Forms.ComboBox
  Friend WithEvents cboRegisteredLevel As System.Windows.Forms.ComboBox
  Friend WithEvents grdLicenses As System.Windows.Forms.DataGrid
  Friend WithEvents MainMenu As System.Windows.Forms.MainMenu
  Friend WithEvents MenuItem_File As System.Windows.Forms.MenuItem
  Friend WithEvents MenuItem_Exit As System.Windows.Forms.MenuItem
  Friend WithEvents cmdPrintLicensesList As System.Windows.Forms.Button
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAlugenDb))
    Me.menudg = New System.Windows.Forms.ContextMenu
    Me.menudg_dettagli = New System.Windows.Forms.MenuItem
    Me.MenuItem3 = New System.Windows.Forms.MenuItem
    Me.menudg_delete = New System.Windows.Forms.MenuItem
    Me.MainMenu = New System.Windows.Forms.MainMenu
    Me.MenuItem_File = New System.Windows.Forms.MenuItem
    Me.MenuItem_Exit = New System.Windows.Forms.MenuItem
    Me.grdLicenses = New System.Windows.Forms.DataGrid
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.cmdPrintLicensesList = New System.Windows.Forms.Button
        Me.dtpExpDateTo = New ALUGEN3_6NET.NullableDateTimePicker
        Me.dtpExpDateFrom = New ALUGEN3_6NET.NullableDateTimePicker
        Me.dtpRegDateTo = New ALUGEN3_6NET.NullableDateTimePicker
        Me.dtpRegDateFrom = New ALUGEN3_6NET.NullableDateTimePicker
    Me.cmdReset = New System.Windows.Forms.Button
    Me.cmdFilter = New System.Windows.Forms.Button
    Me.Label9 = New System.Windows.Forms.Label
    Me.Label8 = New System.Windows.Forms.Label
    Me.Label4 = New System.Windows.Forms.Label
    Me.Label3 = New System.Windows.Forms.Label
    Me.Label15 = New System.Windows.Forms.Label
    Me.cboRegisteredLevel = New System.Windows.Forms.ComboBox
    Me.cboLicType = New System.Windows.Forms.ComboBox
    Me.Label6 = New System.Windows.Forms.Label
    Me.Label5 = New System.Windows.Forms.Label
    Me.txtlocktype = New System.Windows.Forms.TextBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.Label1 = New System.Windows.Forms.Label
    Me.cboProductName = New System.Windows.Forms.ComboBox
    Me.txtusername = New System.Windows.Forms.TextBox
    CType(Me.grdLicenses, System.ComponentModel.ISupportInitialize).BeginInit()
    Me.GroupBox1.SuspendLayout()
    Me.SuspendLayout()
    '
    'menudg
    '
    Me.menudg.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.menudg_dettagli, Me.MenuItem3, Me.menudg_delete})
    '
    'menudg_dettagli
    '
    Me.menudg_dettagli.Index = 0
    Me.menudg_dettagli.Text = "Details"
    '
    'MenuItem3
    '
    Me.MenuItem3.Index = 1
    Me.MenuItem3.Text = "-"
    '
    'menudg_delete
    '
    Me.menudg_delete.Index = 2
    Me.menudg_delete.Text = "Delete"
    '
    'MainMenu
    '
    Me.MainMenu.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem_File})
    '
    'MenuItem_File
    '
    Me.MenuItem_File.Index = 0
    Me.MenuItem_File.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem_Exit})
    Me.MenuItem_File.Text = "File"
    '
    'MenuItem_Exit
    '
    Me.MenuItem_Exit.Index = 0
    Me.MenuItem_Exit.Text = "Exit"
    '
    'grdLicenses
    '
    Me.grdLicenses.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.grdLicenses.ContextMenu = Me.menudg
    Me.grdLicenses.DataMember = ""
    Me.grdLicenses.HeaderForeColor = System.Drawing.SystemColors.ControlText
    Me.grdLicenses.Location = New System.Drawing.Point(0, 120)
    Me.grdLicenses.Name = "grdLicenses"
    Me.grdLicenses.ReadOnly = True
    Me.grdLicenses.Size = New System.Drawing.Size(664, 288)
    Me.grdLicenses.TabIndex = 1
    '
    'GroupBox1
    '
    Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox1.Controls.Add(Me.cmdPrintLicensesList)
    Me.GroupBox1.Controls.Add(Me.dtpExpDateTo)
    Me.GroupBox1.Controls.Add(Me.dtpExpDateFrom)
    Me.GroupBox1.Controls.Add(Me.dtpRegDateTo)
    Me.GroupBox1.Controls.Add(Me.dtpRegDateFrom)
    Me.GroupBox1.Controls.Add(Me.cmdReset)
    Me.GroupBox1.Controls.Add(Me.cmdFilter)
    Me.GroupBox1.Controls.Add(Me.Label9)
    Me.GroupBox1.Controls.Add(Me.Label8)
    Me.GroupBox1.Controls.Add(Me.Label4)
    Me.GroupBox1.Controls.Add(Me.Label3)
    Me.GroupBox1.Controls.Add(Me.Label15)
    Me.GroupBox1.Controls.Add(Me.cboRegisteredLevel)
    Me.GroupBox1.Controls.Add(Me.cboLicType)
    Me.GroupBox1.Controls.Add(Me.Label6)
    Me.GroupBox1.Controls.Add(Me.Label5)
    Me.GroupBox1.Controls.Add(Me.txtlocktype)
    Me.GroupBox1.Controls.Add(Me.Label2)
    Me.GroupBox1.Controls.Add(Me.Label1)
    Me.GroupBox1.Controls.Add(Me.cboProductName)
    Me.GroupBox1.Controls.Add(Me.txtusername)
    Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(664, 118)
    Me.GroupBox1.TabIndex = 0
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = " Filters "
    '
    'cmdPrintLicensesList
    '
    Me.cmdPrintLicensesList.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdPrintLicensesList.BackColor = System.Drawing.SystemColors.Control
    Me.cmdPrintLicensesList.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdPrintLicensesList.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdPrintLicensesList.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdPrintLicensesList.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdPrintLicensesList.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdPrintLicensesList.Location = New System.Drawing.Point(516, 88)
    Me.cmdPrintLicensesList.Name = "cmdPrintLicensesList"
    Me.cmdPrintLicensesList.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdPrintLicensesList.Size = New System.Drawing.Size(72, 23)
    Me.cmdPrintLicensesList.TabIndex = 65
    Me.cmdPrintLicensesList.Text = "Print  list"
    Me.cmdPrintLicensesList.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    '
    'dtpExpDateTo
    '
    Me.dtpExpDateTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dtpExpDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Short
    Me.dtpExpDateTo.Location = New System.Drawing.Point(560, 40)
    Me.dtpExpDateTo.Name = "dtpExpDateTo"
    Me.dtpExpDateTo.NullValue = " (none)"
    Me.dtpExpDateTo.Size = New System.Drawing.Size(96, 20)
    Me.dtpExpDateTo.TabIndex = 23
    Me.dtpExpDateTo.Value = New Date(2006, 5, 24, 23, 39, 50, 850)
    '
    'dtpExpDateFrom
    '
    Me.dtpExpDateFrom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dtpExpDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short
    Me.dtpExpDateFrom.Location = New System.Drawing.Point(424, 40)
    Me.dtpExpDateFrom.Name = "dtpExpDateFrom"
    Me.dtpExpDateFrom.NullValue = " (none)"
    Me.dtpExpDateFrom.Size = New System.Drawing.Size(96, 20)
    Me.dtpExpDateFrom.TabIndex = 22
    Me.dtpExpDateFrom.Value = New Date(2006, 5, 24, 23, 39, 50, 850)
    '
    'dtpRegDateTo
    '
    Me.dtpRegDateTo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dtpRegDateTo.Format = System.Windows.Forms.DateTimePickerFormat.Short
    Me.dtpRegDateTo.Location = New System.Drawing.Point(560, 16)
    Me.dtpRegDateTo.Name = "dtpRegDateTo"
    Me.dtpRegDateTo.NullValue = " (none)"
    Me.dtpRegDateTo.Size = New System.Drawing.Size(96, 20)
    Me.dtpRegDateTo.TabIndex = 21
    Me.dtpRegDateTo.Value = New Date(2006, 5, 24, 23, 39, 50, 850)
    '
    'dtpRegDateFrom
    '
    Me.dtpRegDateFrom.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.dtpRegDateFrom.Format = System.Windows.Forms.DateTimePickerFormat.Short
    Me.dtpRegDateFrom.Location = New System.Drawing.Point(424, 16)
    Me.dtpRegDateFrom.Name = "dtpRegDateFrom"
    Me.dtpRegDateFrom.NullValue = " (none)"
    Me.dtpRegDateFrom.Size = New System.Drawing.Size(96, 20)
    Me.dtpRegDateFrom.TabIndex = 11
    Me.dtpRegDateFrom.Value = New Date(2006, 5, 24, 23, 39, 50, 850)
    '
    'cmdReset
    '
    Me.cmdReset.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdReset.BackColor = System.Drawing.SystemColors.Control
    Me.cmdReset.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdReset.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdReset.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdReset.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdReset.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdReset.Location = New System.Drawing.Point(592, 88)
    Me.cmdReset.Name = "cmdReset"
    Me.cmdReset.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdReset.Size = New System.Drawing.Size(64, 23)
    Me.cmdReset.TabIndex = 1
    Me.cmdReset.Text = "&Reset"
    Me.cmdReset.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    '
    'cmdFilter
    '
    Me.cmdFilter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdFilter.BackColor = System.Drawing.SystemColors.Control
    Me.cmdFilter.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdFilter.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdFilter.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdFilter.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdFilter.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdFilter.Location = New System.Drawing.Point(448, 88)
    Me.cmdFilter.Name = "cmdFilter"
    Me.cmdFilter.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdFilter.Size = New System.Drawing.Size(64, 23)
    Me.cmdFilter.TabIndex = 0
    Me.cmdFilter.Text = "&Filter"
    Me.cmdFilter.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    '
    'Label9
    '
    Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label9.BackColor = System.Drawing.SystemColors.Control
    Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
    Me.Label9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
    Me.Label9.Location = New System.Drawing.Point(536, 40)
    Me.Label9.Name = "Label9"
    Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label9.Size = New System.Drawing.Size(16, 15)
    Me.Label9.TabIndex = 17
    Me.Label9.Text = "to"
    '
    'Label8
    '
    Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label8.BackColor = System.Drawing.SystemColors.Control
    Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
    Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
    Me.Label8.Location = New System.Drawing.Point(536, 16)
    Me.Label8.Name = "Label8"
    Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label8.Size = New System.Drawing.Size(16, 15)
    Me.Label8.TabIndex = 13
    Me.Label8.Text = "to"
    '
    'Label4
    '
    Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label4.BackColor = System.Drawing.SystemColors.Control
    Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
    Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
    Me.Label4.Location = New System.Drawing.Point(304, 40)
    Me.Label4.Name = "Label4"
    Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label4.Size = New System.Drawing.Size(112, 17)
    Me.Label4.TabIndex = 15
    Me.Label4.Text = "Expiration date: from"
    '
    'Label3
    '
    Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label3.BackColor = System.Drawing.SystemColors.Control
    Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
    Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
    Me.Label3.Location = New System.Drawing.Point(304, 16)
    Me.Label3.Name = "Label3"
    Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label3.Size = New System.Drawing.Size(120, 17)
    Me.Label3.TabIndex = 8
    Me.Label3.Text = "Registration date: from"
    '
    'Label15
    '
    Me.Label15.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Label15.BackColor = System.Drawing.SystemColors.Control
    Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
    Me.Label15.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label15.ForeColor = System.Drawing.SystemColors.ControlText
    Me.Label15.Location = New System.Drawing.Point(304, 64)
    Me.Label15.Name = "Label15"
    Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label15.Size = New System.Drawing.Size(90, 17)
    Me.Label15.TabIndex = 19
    Me.Label15.Text = "Registered Level:"
    '
    'cboRegisteredLevel
    '
    Me.cboRegisteredLevel.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboRegisteredLevel.BackColor = System.Drawing.SystemColors.Window
    Me.cboRegisteredLevel.Cursor = System.Windows.Forms.Cursors.Default
    Me.cboRegisteredLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboRegisteredLevel.ForeColor = System.Drawing.SystemColors.WindowText
    Me.cboRegisteredLevel.Location = New System.Drawing.Point(424, 64)
    Me.cboRegisteredLevel.Name = "cboRegisteredLevel"
    Me.cboRegisteredLevel.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cboRegisteredLevel.Size = New System.Drawing.Size(232, 21)
    Me.cboRegisteredLevel.TabIndex = 20
    '
    'cboLicType
    '
    Me.cboLicType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboLicType.BackColor = System.Drawing.SystemColors.Window
    Me.cboLicType.Cursor = System.Windows.Forms.Cursors.Default
    Me.cboLicType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
    Me.cboLicType.ForeColor = System.Drawing.SystemColors.WindowText
    Me.cboLicType.Location = New System.Drawing.Point(80, 64)
    Me.cboLicType.Name = "cboLicType"
    Me.cboLicType.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cboLicType.Size = New System.Drawing.Size(216, 21)
    Me.cboLicType.TabIndex = 5
    '
    'Label6
    '
    Me.Label6.BackColor = System.Drawing.SystemColors.Control
    Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
    Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
    Me.Label6.Location = New System.Drawing.Point(8, 64)
    Me.Label6.Name = "Label6"
    Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label6.Size = New System.Drawing.Size(72, 17)
    Me.Label6.TabIndex = 4
    Me.Label6.Text = "License &Type:"
    '
    'Label5
    '
    Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.Label5.Location = New System.Drawing.Point(8, 88)
    Me.Label5.Name = "Label5"
    Me.Label5.Size = New System.Drawing.Size(72, 17)
    Me.Label5.TabIndex = 6
    Me.Label5.Text = "Lock Type"
    '
    'txtlocktype
    '
    Me.txtlocktype.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtlocktype.Location = New System.Drawing.Point(80, 88)
    Me.txtlocktype.Name = "txtlocktype"
    Me.txtlocktype.Size = New System.Drawing.Size(216, 20)
    Me.txtlocktype.TabIndex = 7
    Me.txtlocktype.Text = ""
    '
    'Label2
    '
    Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.Label2.Location = New System.Drawing.Point(8, 40)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(72, 17)
    Me.Label2.TabIndex = 2
    Me.Label2.Text = "User Name"
    '
    'Label1
    '
    Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.Label1.Location = New System.Drawing.Point(8, 16)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(72, 17)
    Me.Label1.TabIndex = 0
    Me.Label1.Text = "Program"
    '
    'cboProductName
    '
    Me.cboProductName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cboProductName.Location = New System.Drawing.Point(80, 16)
    Me.cboProductName.Name = "cboProductName"
    Me.cboProductName.Size = New System.Drawing.Size(216, 21)
    Me.cboProductName.TabIndex = 1
    '
    'txtusername
    '
    Me.txtusername.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtusername.Location = New System.Drawing.Point(80, 40)
    Me.txtusername.Name = "txtusername"
    Me.txtusername.Size = New System.Drawing.Size(216, 20)
    Me.txtusername.TabIndex = 3
    Me.txtusername.Text = ""
    '
    'frmAlugenDb
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.ClientSize = New System.Drawing.Size(664, 409)
    Me.Controls.Add(Me.GroupBox1)
    Me.Controls.Add(Me.grdLicenses)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Menu = Me.MainMenu
    Me.MinimumSize = New System.Drawing.Size(672, 456)
    Me.Name = "frmAlugenDb"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Alugen License Database"
    CType(Me.grdLicenses, System.ComponentModel.ISupportInitialize).EndInit()
    Me.GroupBox1.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

#Region "Events"
  Private Sub frmAlugenDb_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    'setting connection to database
    InitConnection()

    'load images for buttons
    LoadImages()

    'populate combos
    PopulateCboProductName()
    PopulateCboRegLevel()
    PopulateCboLicType()

    'reset all filters by default
    cmdReset_Click(sender, e)

  End Sub

  Private Sub menudg_delete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menudg_delete.Click
    Dim cm As New OleDbCommand
        cm.Connection = conn
    cm.CommandText = "DELETE * FROM license WHERE id=@id"

    cm.Parameters.Clear()
        cm.Parameters.AddWithValue("@id", Me.grdLicenses.Item(Me.grdLicenses.CurrentRowIndex, 0))
    Try
      If conn Is Nothing Then InitConnection()
      conn.Open()
      cm.ExecuteNonQuery()
    Catch ex As Exception
      MessageBox.Show(ex.Message, ACTIVELOCKSTRING, MessageBoxButtons.OK, MessageBoxIcon.Error)
    Finally
      conn.Close()
    End Try

    'refresh data
    LoadData()

  End Sub

  Private Sub menuprinc_exit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem_Exit.Click
    Me.Close()
  End Sub

  Private Sub cmdFilter_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdFilter.Click
    'refresh data
    LoadData()
  End Sub

  Private Sub cmdReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdReset.Click
    'formatting datagrid
    FormatGrdLicensesTableStyle()

    'clean values
    Me.txtlocktype.Text = ""
    Me.txtusername.Text = ""
    Me.cboProductName.SelectedIndex = 0
    Me.cboLicType.SelectedIndex = 0
    Me.cboRegisteredLevel.SelectedIndex = 0
    Me.dtpRegDateFrom.Value = DBNull.Value
    Me.dtpRegDateTo.Value = DBNull.Value
    Me.dtpExpDateFrom.Value = DBNull.Value
    Me.dtpExpDateTo.Value = DBNull.Value

    'refresh data
    LoadData()
  End Sub

  Private Sub menudg_dettagli_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles menudg_dettagli.Click
    Dim currrowindex As Integer
    If Me.grdLicenses.VisibleRowCount <> 0 Then
      currrowindex = Me.grdLicenses.CurrentRowIndex
      Dim licenseDetails As New frmAlugendb_details(CType(Me.grdLicenses.Item(currrowindex, 1), String), _
              CType(Me.grdLicenses.Item(currrowindex, 2), String), CType(Me.grdLicenses.Item(currrowindex, 3), String), _
              CType(Me.grdLicenses.Item(currrowindex, 4), String), CType(Me.grdLicenses.Item(currrowindex, 5), String), _
              CType(Me.grdLicenses.Item(currrowindex, 6), String), CType(Me.grdLicenses.Item(currrowindex, 7), String), _
              CType(Me.grdLicenses.Item(currrowindex, 8), String), CType(Me.grdLicenses.Item(currrowindex, 9), String), _
              CType(Me.grdLicenses.Item(currrowindex, 10), String))
      licenseDetails.ShowDialog()
    End If
  End Sub

  Private Sub grdLicenses_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grdLicenses.DoubleClick
    Dim currrowindex As Integer
    If Me.grdLicenses.VisibleRowCount <> 0 Then
      currrowindex = Me.grdLicenses.CurrentRowIndex
      Dim licenseDetails As New frmAlugendb_details(CType(Me.grdLicenses.Item(currrowindex, 1), String), _
                CType(Me.grdLicenses.Item(currrowindex, 2), String), CType(Me.grdLicenses.Item(currrowindex, 3), String), _
                CType(Me.grdLicenses.Item(currrowindex, 4), String), CType(Me.grdLicenses.Item(currrowindex, 5), String), _
                CType(Me.grdLicenses.Item(currrowindex, 6), String), CType(Me.grdLicenses.Item(currrowindex, 7), String), _
                CType(Me.grdLicenses.Item(currrowindex, 8), String), CType(Me.grdLicenses.Item(currrowindex, 9), String), _
                CType(Me.grdLicenses.Item(currrowindex, 10), String))
      licenseDetails.ShowDialog()
    End If
  End Sub

  Private Sub cmdPrintLicensesList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrintLicensesList.Click
    Dim daPrintDocument As New daReport.DaPrintDocument
    Dim hashParameters As New Hashtable
    'Dim selProduct As ProductInfoItem = cboProducts.SelectedItem
    'Dim itemRegLevel As Mylist = cboRegisteredLevel.SelectedItem

    'set .xml file for printing
    daPrintDocument.setXML("reports\repLicensesList.xml")

    ''build parameters
    'hashParameters.Add("pProductName", selProduct.ProductName)
    'hashParameters.Add("pProductVersion", selProduct.ProductVersion)
    'hashParameters.Add("pRegisteredLevel", itemRegLevel.Name)
    'hashParameters.Add("pLicenseType", CType(cboLicType.SelectedItem, String))
    'hashParameters.Add("pRegisteredDate", "")
    'hashParameters.Add("pExpireDate", "")
    'hashParameters.Add("pInstallCode", txtInstallCode.Text)
    'hashParameters.Add("pUserName", txtUser.Text)
    'hashParameters.Add("pLicenseKey", txtLicenseKey.Text)

    ''setting parameters
    'daPrintDocument.SetParameters(hashParameters)
    ''daPrintDocument.DocumentName = "Licenses list"
    Dim xds As DataTable = CType(grdLicenses.DataSource, DataTable)
    daPrintDocument.AddData(xds)

    'print preview
        printPreviewDialog1.Icon = CType(frmMain.resxList("report_ico"), Icon)
    printPreviewDialog1.Text = daPrintDocument.DocumentName
    printPreviewDialog1.Document = daPrintDocument
    printPreviewDialog1.PrintPreviewControl.Zoom = 1.0
    printPreviewDialog1.WindowState = FormWindowState.Maximized
    printPreviewDialog1.ShowDialog(Me)
  End Sub

#End Region


#Region "Methods"

  Private Sub FormatGrdLicensesTableStyle()
    grdLicenses.RowHeaderWidth = 15
    grdLicenses.CaptionVisible = False

    Dim TableStylegrdLicenses As New DataGridTableStyle
    Dim scarto As Integer = 10

    TableStylegrdLicenses.MappingName = LicensesTableMapping

    Dim idcol As New DataGridTextBoxColumn
    idcol.MappingName = "id"
    idcol.HeaderText = ""
    idcol.Width = 0
    idcol.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(idcol)

    Dim Progname As New DataGridTextBoxColumn
    Progname.MappingName = "Progname"
    Progname.HeaderText = "Program"
    Progname.Width = CType((Me.grdLicenses.Width - Me.grdLicenses.RowHeaderWidth - scarto) / 100 * 10, Integer)
    Progname.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(Progname)

    Dim Ver As New DataGridTextBoxColumn
    Ver.MappingName = "progver"
    Ver.HeaderText = "Ver"
    Ver.Width = CType((Me.grdLicenses.Width - Me.grdLicenses.RowHeaderWidth - scarto) / 100 * 3, Integer)
    Ver.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(Ver)

    Dim RegDate As New DataGridTextBoxColumn
    RegDate.MappingName = "RegDate"
    RegDate.HeaderText = "Registered"
    RegDate.Width = CType((Me.grdLicenses.Width - Me.grdLicenses.RowHeaderWidth - scarto) / 100 * 9, Integer)
    RegDate.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(RegDate)

    Dim ExpDate As New DataGridTextBoxColumn
    ExpDate.MappingName = "ExpDate"
    ExpDate.HeaderText = "Expire"
    ExpDate.Width = CType((Me.grdLicenses.Width - Me.grdLicenses.RowHeaderWidth - scarto) / 100 * 9, Integer)
    ExpDate.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(ExpDate)

    Dim LicType As New DataGridTextBoxColumn
    LicType.MappingName = "LicType"
    LicType.HeaderText = "LicType"
    LicType.Width = CType((Me.grdLicenses.Width - Me.grdLicenses.RowHeaderWidth - scarto) / 100 * 7, Integer)
    LicType.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(LicType)

    Dim LockType As New DataGridTextBoxColumn
    LockType.MappingName = "LockType"
    LockType.HeaderText = "LockType"
    LockType.Width = CType((Me.grdLicenses.Width - Me.grdLicenses.RowHeaderWidth - scarto) / 100 * 5, Integer)
    LockType.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(LockType)

    Dim RegLevel As New DataGridTextBoxColumn
    RegLevel.MappingName = "RegLevel"
    RegLevel.HeaderText = "Level"
    RegLevel.Width = CType((Me.grdLicenses.Width - Me.grdLicenses.RowHeaderWidth - scarto) / 100 * 7, Integer)
    RegLevel.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(RegLevel)

    Dim InstCode As New DataGridTextBoxColumn
    InstCode.MappingName = "InstCode"
    InstCode.HeaderText = "InstCode"
    InstCode.Width = CType((Me.grdLicenses.Width - Me.grdLicenses.RowHeaderWidth - scarto) / 100 * 13, Integer)
    InstCode.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(InstCode)

    Dim UserName As New DataGridTextBoxColumn
    UserName.MappingName = "UserName"
    UserName.HeaderText = "UserName"
    UserName.Width = CType((Me.grdLicenses.Width - Me.grdLicenses.RowHeaderWidth - scarto) / 100 * 10, Integer)
    UserName.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(UserName)

    Dim LibCode As New DataGridTextBoxColumn
    LibCode.MappingName = "LibCode"
    LibCode.HeaderText = "LibCode"
    LibCode.Width = CType((Me.grdLicenses.Width - Me.grdLicenses.RowHeaderWidth - scarto) / 100 * 25, Integer)
    LibCode.TextBox.Font = New Font("Courier New", 8)
    LibCode.NullText = ""
    TableStylegrdLicenses.GridColumnStyles.Add(LibCode)

    TableStylegrdLicenses.PreferredRowHeight = 100

    Me.grdLicenses.TableStyles.Clear()
    Me.grdLicenses.TableStyles.Add(TableStylegrdLicenses)

    grdLicenses.Font = New Font("Courier New", 8)

    gridHelper = New GridLayoutHelper(grdLicenses, _
        grdLicenses.TableStyles(0), _
        New Decimal() {0D, 0.1D, 0.03D, 0.09D, 0.09D, 0.07D, 0.05D, 0.07D, 0.13D, 0.1D, 0.25D}, _
        New Integer() {0, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20})

  End Sub


  Public Sub ArchiveLicense(ByVal productName As String, ByVal productversion As String, ByVal userName As String, _
ByVal registrationDate As String, ByVal expiresAfter As String, ByVal licenseType As String, ByVal lockType As String, _
ByVal registeredLevel As String, ByVal installationCode As String, ByVal licenseCode As String)
    Dim cm As New OleDbCommand

    cm.CommandText = "INSERT INTO license ( Progname, progver, RegDate, ExpDate, LicType, LockType, RegLevel, InstCode, UserName, LibCode )" & _
                    " VALUES(@productName,@productversion,@registrationDate,@expiresAfter,@licenseType,@lockType,@registeredLevel,@installationCode,@userName,@licenseCode)"
    cm.Parameters.Clear()
        'walter'obsolete'cm.Parameters.Add("@productName", productName)
        'walter'obsolete'cm.Parameters.Add("@productversion", productversion)
        'walter'obsolete'cm.Parameters.Add("@registrationDate", registrationDate)
        'walter'obsolete'cm.Parameters.Add("@expiresAfter", expiresAfter)
        'walter'obsolete'cm.Parameters.Add("@licenseType", licenseType)
        'walter'obsolete'cm.Parameters.Add("@lockType", lockType)
        'walter'obsolete'cm.Parameters.Add("@registeredLevel", registeredLevel)
        'walter'obsolete'cm.Parameters.Add("@installationCode", installationCode)
        'walter'obsolete'cm.Parameters.Add("@userName", userName)
        'walter'obsolete'cm.Parameters.Add("@licenseCode", licenseCode)
        cm.Parameters.AddWithValue("@productName", productName)
        cm.Parameters.AddWithValue("@productversion", productversion)
        cm.Parameters.AddWithValue("@registrationDate", registrationDate)
        cm.Parameters.AddWithValue("@expiresAfter", expiresAfter)
        cm.Parameters.AddWithValue("@licenseType", licenseType)
        cm.Parameters.AddWithValue("@lockType", lockType)
        cm.Parameters.AddWithValue("@registeredLevel", registeredLevel)
        cm.Parameters.AddWithValue("@installationCode", installationCode)
        cm.Parameters.AddWithValue("@userName", userName)
        cm.Parameters.AddWithValue("@licenseCode", licenseCode)
    Try
      If conn Is Nothing Then InitConnection()
      conn.Open()
      cm.Connection = conn
      cm.ExecuteNonQuery()
    Catch ex As Exception
      MessageBox.Show(ex.Message, ACTIVELOCKSTRING, MessageBoxButtons.OK, MessageBoxIcon.Error)
    Finally
      conn.Close()
    End Try

        Cursor = Cursors.Default
  End Sub

  Public Sub PopulateCboProductName()
    Dim oRow As DataRow
    Dim QueryStr As String
    Dim Da As New OleDbDataAdapter
    Dim cm As New OleDbCommand
    Dim ds As New DataSet

    Me.cboProductName.Items.Clear()
    Dim itemProductAll As New ProductInfoItem("(All)", "")
    Me.cboProductName.Items.Add(itemProductAll)

    Da.SelectCommand = cm

    QueryStr = "SELECT DISTINCT license.Progname,license.Progver" & _
                " FROM(license)" & _
                " ORDER BY license.Progname "
    cm.CommandText = QueryStr
    ds.Clear()

    Try
      If conn Is Nothing Then InitConnection()
      conn.Open()
      cm.Connection = conn
      Da.Fill(ds, "progname")
    Catch ex As Exception
      MessageBox.Show(ex.Message, ACTIVELOCKSTRING, MessageBoxButtons.OK, MessageBoxIcon.Error)
    Finally
      conn.Close()
    End Try

    'Load all items into combo
    For Each oRow In ds.Tables("progname").Rows
      Dim itemProduct As New ProductInfoItem(oRow.Item("Progname").ToString(), oRow.Item("Progver").ToString())
      Me.cboProductName.Items.Add(itemProduct)
    Next

    Me.cboProductName.DisplayMember = "ProductNameVersion"
    Me.cboProductName.ValueMember = "ProductNameVersion"
    Me.cboProductName.SelectedIndex = 0

  End Sub

  Public Sub PopulateCboRegLevel()
    'populate registered levels combo
    Dim oRow As DataRow
    Dim QueryStr As String
    Dim Da As New OleDbDataAdapter
    Dim cm As New OleDbCommand
    Dim ds As New DataSet

    Me.cboRegisteredLevel.Items.Clear()
    Me.cboRegisteredLevel.Items.Add("(All)")

    Da.SelectCommand = cm

    QueryStr = "SELECT DISTINCT license.RegLevel" & _
                " FROM (license)" & _
                " ORDER BY license.RegLevel "
    cm.CommandText = QueryStr
    ds.Clear()

    Try
      If conn Is Nothing Then InitConnection()
      conn.Open()
      cm.Connection = conn
      Da.Fill(ds, "reglevels")
    Catch ex As Exception
      MessageBox.Show(ex.Message, ACTIVELOCKSTRING, MessageBoxButtons.OK, MessageBoxIcon.Error)
    Finally
      conn.Close()
    End Try

    'Load all items into combo
    For Each oRow In ds.Tables("reglevels").Rows
      Me.cboRegisteredLevel.Items.Add(oRow.Item(0).ToString())
    Next

    Me.cboRegisteredLevel.SelectedIndex = 0
  End Sub

  Public Sub PopulateCboLicType()
    'populate license type combo
    With Me.cboLicType
      .Items.Clear()
      .Items.Add("(All)")
      .Items.Add("Time Locked")
      .Items.Add("Periodic")
      .Items.Add("Permanent")
      .SelectedIndex = 0
    End With
  End Sub

  Private Sub LoadImages()
    'load buttons images
        cmdReset.Image = CType(frmMain.resxList("reset"), Image)
        cmdFilter.Image = CType(frmMain.resxList("filter"), Image)
        cmdPrintLicensesList.Image = CType(frmMain.resxList("print"), Image)
  End Sub

  Private Function BuildQueryString(ByVal strDefaultQuery As String) As String
    'build main query string
    Dim strResult As String = strDefaultQuery
    Dim locktype, username, progname, progver, _
        selcboLicType, selcboRegisteredLevel As String

    'resolve filters
    'Locktype
    If Me.txtlocktype.Text <> "" Then
      locktype = Me.txtlocktype.Text
    Else
      locktype = "%"
    End If
    'username
    If Me.txtusername.Text <> "" Then
      username = Me.txtusername.Text
    Else
      username = "%"
    End If
    'product name and version
    If Me.cboProductName.SelectedIndex <> 0 Then
      Dim selProductInfo As ProductInfoItem = CType(Me.cboProductName.SelectedItem, ProductInfoItem)
      progname = selProductInfo.ProductName
      progver = selProductInfo.ProductVersion
    Else
      progname = "%"
      progver = "%"
    End If
    'license type
    If Me.cboLicType.SelectedIndex <> 0 Then
      selcboLicType = CType(Me.cboLicType.SelectedItem, String)
    Else
      selcboLicType = "%"
    End If
    'registered level
    If Me.cboRegisteredLevel.SelectedIndex <> 0 Then
      selcboRegisteredLevel = CType(Me.cboRegisteredLevel.SelectedItem, String)
    Else
      selcboRegisteredLevel = "%"
    End If

    'replace in main template
    strResult = strResult.Replace("@locktype", "'" & locktype & "'")
    strResult = strResult.Replace("@username", "'" & username & "'")
    strResult = strResult.Replace("@progname", "'" & progname & "'")
    strResult = strResult.Replace("@progver", "'" & progver & "'")
    strResult = strResult.Replace("@selcboLicType", "'" & selcboLicType & "'")
    strResult = strResult.Replace("@selcboRegisteredLevel", "'" & selcboRegisteredLevel & "'")

    If Not dtpRegDateFrom.Value Is Nothing Then
      strResult = strResult & " AND datediff('d',RegDate,'" & Convert.ToString(dtpRegDateFrom.Value) & "')<=0"
    End If
    If Not dtpRegDateTo.Value Is Nothing Then
      strResult = strResult & " AND datediff('d',RegDate,'" & Convert.ToString(dtpRegDateTo.Value) & "')>=0"
    End If
    If Not dtpExpDateFrom.Value Is Nothing Then
      strResult = strResult & " AND datediff('d',ExpDate,'" & Convert.ToString(dtpExpDateFrom.Value) & "')<=0"
    End If
    If Not dtpExpDateTo.Value Is Nothing Then
      strResult = strResult & " AND datediff('d',ExpDate,'" & Convert.ToString(dtpExpDateTo.Value) & "')>=0"
    End If

    Return strResult
  End Function

  Private Sub LoadData()
    Dim daLicenses As New OleDbDataAdapter
    Dim dsLicenses As New DataSet
    Dim cmLicenses As New OleDbCommand

    daLicenses.SelectCommand = cmLicenses
    dsLicenses.Clear()
    cmLicenses.CommandText = BuildQueryString(DefaultQueryTemplate)

    Try
      If conn Is Nothing Then InitConnection()
      conn.Open()
      cmLicenses.Connection = conn
      'fill dataset
      daLicenses.Fill(dsLicenses, LicensesTableMapping)
    Catch ex As Exception
      MessageBox.Show(ex.Message, ACTIVELOCKSTRING, MessageBoxButtons.OK, MessageBoxIcon.Error)
    Finally
      conn.Close()
    End Try

    'set datagrid source
    grdLicenses.DataSource = dsLicenses.Tables.Item(LicensesTableMapping)
  End Sub

  Private Sub InitConnection()
    conn = New OleDbConnection(stringaconn)
  End Sub

#End Region


End Class
