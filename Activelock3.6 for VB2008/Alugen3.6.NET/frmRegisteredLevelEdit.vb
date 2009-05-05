' *   ActiveLock
' *   Copyright 1998-2002 Nelson Ferraz
' *   Copyright 2003-2008 The ActiveLock Software Group (ASG)
' *   All material is the property of the contributing authors.
' *
' *   Redistribution and use in source and binary forms, with or without
' *   modification, are permitted provided that the following conditions are
' *   met:
' *
' *     [o] Redistributions of source code must retain the above copyright
' *         notice, this list of conditions and the following disclaimer.
' *
' *     [o] Redistributions in binary form must reproduce the above
' *         copyright notice, this list of conditions and the following
' *         disclaimer in the documentation and/or other materials provided
' *         with the distribution.
' *
' *   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' *   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' *   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
' *   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
' *   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
' *   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
' *   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
' *   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' *   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' *   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' *   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' *
' *
'--------------------------------------------------------------------------------
'    Component  : dlgRegisteredLevel
'    Project    : ALUGEN3
'
'    Description: For Add and Modify a Registered Level
'
'    Modified   : By Kirtaph on 2006-02-16
'--------------------------------------------------------------------------------
'</CSCC>
Option Strict On
Option Explicit On 
Imports System.Windows.Forms

Friend Class frmRegisteredLevelEdit
  Inherits System.Windows.Forms.Form

#Region "Windows Form Designer generated code "
  Public Sub New()
    MyBase.New()
    'This call is required by the Windows Form Designer.
    InitializeComponent()
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
  Public ToolTip1 As System.Windows.Forms.ToolTip
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents cmdOk As System.Windows.Forms.Button
  Friend WithEvents txtItemData As System.Windows.Forms.TextBox
  Friend WithEvents txtDescription As System.Windows.Forms.TextBox
  Friend WithEvents Line2 As System.Windows.Forms.Label
  Friend WithEvents Line1 As System.Windows.Forms.Label
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents Label1 As System.Windows.Forms.Label
  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.
  'Do not modify it using the code editor.
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container
    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.cmdOk = New System.Windows.Forms.Button
    Me.txtItemData = New System.Windows.Forms.TextBox
    Me.txtDescription = New System.Windows.Forms.TextBox
    Me.Line2 = New System.Windows.Forms.Label
    Me.Line1 = New System.Windows.Forms.Label
    Me.Label2 = New System.Windows.Forms.Label
    Me.Label1 = New System.Windows.Forms.Label
    Me.SuspendLayout()
    '
    'cmdCancel
    '
    Me.cmdCancel.BackColor = System.Drawing.SystemColors.Control
    Me.cmdCancel.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdCancel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdCancel.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdCancel.Location = New System.Drawing.Point(308, 66)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdCancel.Size = New System.Drawing.Size(76, 25)
    Me.cmdCancel.TabIndex = 5
    Me.cmdCancel.Text = "Cancel"
    Me.cmdCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.ToolTip1.SetToolTip(Me.cmdCancel, "Cancel")
    '
    'cmdOk
    '
    Me.cmdOk.BackColor = System.Drawing.SystemColors.Control
    Me.cmdOk.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdOk.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdOk.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdOk.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdOk.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdOk.Location = New System.Drawing.Point(224, 66)
    Me.cmdOk.Name = "cmdOk"
    Me.cmdOk.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdOk.Size = New System.Drawing.Size(76, 25)
    Me.cmdOk.TabIndex = 4
    Me.cmdOk.Text = "Ok"
    Me.cmdOk.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.ToolTip1.SetToolTip(Me.cmdOk, "Ok")
    '
    'txtItemData
    '
    Me.txtItemData.AcceptsReturn = True
    Me.txtItemData.AutoSize = False
    Me.txtItemData.BackColor = System.Drawing.SystemColors.Window
    Me.txtItemData.Cursor = System.Windows.Forms.Cursors.IBeam
    Me.txtItemData.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtItemData.ForeColor = System.Drawing.SystemColors.WindowText
    Me.txtItemData.Location = New System.Drawing.Point(309, 21)
    Me.txtItemData.MaxLength = 0
    Me.txtItemData.Name = "txtItemData"
    Me.txtItemData.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.txtItemData.Size = New System.Drawing.Size(76, 22)
    Me.txtItemData.TabIndex = 3
    Me.txtItemData.Text = ""
    Me.ToolTip1.SetToolTip(Me.txtItemData, "Eneter registered level value")
    '
    'txtDescription
    '
    Me.txtDescription.AcceptsReturn = True
    Me.txtDescription.AutoSize = False
    Me.txtDescription.BackColor = System.Drawing.SystemColors.Window
    Me.txtDescription.Cursor = System.Windows.Forms.Cursors.IBeam
    Me.txtDescription.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtDescription.ForeColor = System.Drawing.SystemColors.WindowText
    Me.txtDescription.Location = New System.Drawing.Point(9, 21)
    Me.txtDescription.MaxLength = 0
    Me.txtDescription.Name = "txtDescription"
    Me.txtDescription.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.txtDescription.Size = New System.Drawing.Size(291, 22)
    Me.txtDescription.TabIndex = 1
    Me.txtDescription.Text = ""
    Me.ToolTip1.SetToolTip(Me.txtDescription, "Enter registered level description")
    '
    'Line2
    '
    Me.Line2.BackColor = System.Drawing.SystemColors.ControlLightLight
    Me.Line2.Location = New System.Drawing.Point(9, 58)
    Me.Line2.Name = "Line2"
    Me.Line2.Size = New System.Drawing.Size(375, 1)
    Me.Line2.TabIndex = 6
    '
    'Line1
    '
    Me.Line1.BackColor = System.Drawing.SystemColors.ControlDarkDark
    Me.Line1.Location = New System.Drawing.Point(9, 57)
    Me.Line1.Name = "Line1"
    Me.Line1.Size = New System.Drawing.Size(372, 1)
    Me.Line1.TabIndex = 7
    '
    'Label2
    '
    Me.Label2.BackColor = System.Drawing.SystemColors.Control
    Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
    Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
    Me.Label2.Location = New System.Drawing.Point(309, 6)
    Me.Label2.Name = "Label2"
    Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label2.Size = New System.Drawing.Size(76, 16)
    Me.Label2.TabIndex = 2
    Me.Label2.Text = "Item Data:"
    '
    'Label1
    '
    Me.Label1.BackColor = System.Drawing.SystemColors.Control
    Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
    Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
    Me.Label1.Location = New System.Drawing.Point(9, 6)
    Me.Label1.Name = "Label1"
    Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label1.Size = New System.Drawing.Size(157, 16)
    Me.Label1.TabIndex = 0
    Me.Label1.Text = "Description:"
    '
    'frmRegisteredLevelEdit
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.BackColor = System.Drawing.SystemColors.Control
    Me.ClientSize = New System.Drawing.Size(395, 98)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOk)
    Me.Controls.Add(Me.txtItemData)
    Me.Controls.Add(Me.txtDescription)
    Me.Controls.Add(Me.Line2)
    Me.Controls.Add(Me.Line1)
    Me.Controls.Add(Me.Label2)
    Me.Controls.Add(Me.Label1)
    Me.Cursor = System.Windows.Forms.Cursors.Default
    Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Location = New System.Drawing.Point(3, 28)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.Name = "frmRegisteredLevelEdit"
    Me.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Registered Level Properties"
    Me.ResumeLayout(False)

  End Sub
#End Region



#Region "Private members"
  Private Const NumberOfControl As Short = 2
  Private blnAnnullato As Boolean
  Private blnCanEnabled(NumberOfControl - 1) As Boolean
  Private m_strDescription As String
  Private m_intItemData As Integer
#End Region



#Region "Properties"
  Public Property ItemData() As Integer
    Get
      If Val(txtItemData.Text) > 32767 Then
        MsgBox("ItemData cannot be greater than 32767.", vbExclamation)
        Exit Property
      End If
      ItemData = Convert.ToInt32(txtItemData.Text)
    End Get
    Set(ByVal Value As Integer)
      m_intItemData = Value
    End Set
  End Property

  Public Property Description() As String
    Get
      Description = txtDescription.Text
    End Get
    Set(ByVal Value As String)
      m_strDescription = Value
    End Set
  End Property

  Public ReadOnly Property Annullato() As Boolean
    Get
      Annullato = blnAnnullato
    End Get
  End Property
#End Region



#Region "Methods"
  Private Function CanEnable() As Boolean
    Dim inti As Integer
    CanEnable = True
    For inti = 0 To UBound(blnCanEnabled)
      CanEnable = CanEnable And blnCanEnabled(inti)
    Next
  End Function

  Private Function InitializeEnable() As Object
    Dim inti As Integer
    For inti = 0 To UBound(blnCanEnabled)
      blnCanEnabled(inti) = False
        Next
        Return Nothing
  End Function

  Private Sub LoadImages()
    'load buttons images
        cmdCancel.Image = CType(frmMain.resxList("close"), Image)
        cmdOk.Image = CType(frmMain.resxList("validate"), Image)
  End Sub
#End Region



#Region "Events"
  Private Sub dlgRegisteredLevel_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
    'load buttons images
    LoadImages()

    InitializeEnable()
    txtItemData.Text = CStr(m_intItemData)
    txtDescription.Text = m_strDescription
    cmdOk.Enabled = CanEnable()
  End Sub

  Private Sub dlgRegisteredLevel_Closing(ByVal eventSender As System.Object, ByVal eventArgs As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
    Dim Cancel As Boolean = eventArgs.Cancel
    'blnAnnullato = True
    eventArgs.Cancel = Cancel
  End Sub

  Private Sub cmdCancel_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCancel.Click
    blnAnnullato = True
    Me.Hide()
  End Sub

  Private Sub cmdOk_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdOk.Click
    blnAnnullato = False
    Me.Hide()
  End Sub

  Private Sub txtDescription_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtDescription.TextChanged
    If txtDescription.Text = "" Then
      blnCanEnabled(0) = False
    Else
      blnCanEnabled(0) = True
    End If
    If txtItemData.Text = "" Then
      blnCanEnabled(1) = False
    Else
      blnCanEnabled(1) = True
    End If
    cmdOk.Enabled = CanEnable()
  End Sub

  Private Sub txtItemData_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtItemData.TextChanged
    If txtDescription.Text = "" Then
      blnCanEnabled(0) = False
    Else
      blnCanEnabled(0) = True
    End If
    If txtItemData.Text = "" Then
      blnCanEnabled(1) = False
    Else
      blnCanEnabled(1) = True
    End If
    cmdOk.Enabled = CanEnable()
  End Sub
#End Region


End Class