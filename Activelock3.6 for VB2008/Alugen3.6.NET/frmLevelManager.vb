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
'    Component  : frmLevelManager
'    Project    : ALUGEN3
'
'    Description: For organize your personal Registered Level
'
'    Modified   : By Kirtaph On 2006-02-16
'--------------------------------------------------------------------------------
'</CSCC>
Option Strict On
Option Explicit On 

Friend Class frmLevelManager
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
  Friend WithEvents lstRegisteredLevel As System.Windows.Forms.ListBox
  Friend WithEvents cmdClose As System.Windows.Forms.Button
  Friend WithEvents cmdRemove As System.Windows.Forms.Button
  Friend WithEvents cmdAdd As System.Windows.Forms.Button
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents Line2 As System.Windows.Forms.Label
  Friend WithEvents Line1 As System.Windows.Forms.Label
	'NOTE: The following procedure is required by the Windows Form Designer
	'It can be modified using the Windows Form Designer.
	'Do not modify it using the code editor.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Me.components = New System.ComponentModel.Container
    Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmLevelManager))
    Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
    Me.lstRegisteredLevel = New System.Windows.Forms.ListBox
    Me.cmdClose = New System.Windows.Forms.Button
    Me.cmdRemove = New System.Windows.Forms.Button
    Me.cmdAdd = New System.Windows.Forms.Button
    Me.Label1 = New System.Windows.Forms.Label
    Me.Line2 = New System.Windows.Forms.Label
    Me.Line1 = New System.Windows.Forms.Label
    Me.SuspendLayout()
    '
    'lstRegisteredLevel
    '
    Me.lstRegisteredLevel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.lstRegisteredLevel.BackColor = System.Drawing.SystemColors.Window
    Me.lstRegisteredLevel.Cursor = System.Windows.Forms.Cursors.Default
    Me.lstRegisteredLevel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.lstRegisteredLevel.ForeColor = System.Drawing.SystemColors.WindowText
    Me.lstRegisteredLevel.ItemHeight = 14
    Me.lstRegisteredLevel.Location = New System.Drawing.Point(6, 27)
    Me.lstRegisteredLevel.Name = "lstRegisteredLevel"
    Me.lstRegisteredLevel.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.lstRegisteredLevel.Size = New System.Drawing.Size(290, 270)
    Me.lstRegisteredLevel.TabIndex = 3
    Me.ToolTip1.SetToolTip(Me.lstRegisteredLevel, "Registered levels list. Doubleclick to edit position.")
    '
    'cmdClose
    '
    Me.cmdClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdClose.BackColor = System.Drawing.SystemColors.Control
    Me.cmdClose.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdClose.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdClose.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdClose.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdClose.Location = New System.Drawing.Point(226, 318)
    Me.cmdClose.Name = "cmdClose"
    Me.cmdClose.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdClose.Size = New System.Drawing.Size(70, 23)
    Me.cmdClose.TabIndex = 2
    Me.cmdClose.Text = "Close"
    Me.cmdClose.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.ToolTip1.SetToolTip(Me.cmdClose, "Close this window")
    '
    'cmdRemove
    '
    Me.cmdRemove.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.cmdRemove.BackColor = System.Drawing.SystemColors.Control
    Me.cmdRemove.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdRemove.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdRemove.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdRemove.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdRemove.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdRemove.Location = New System.Drawing.Point(84, 318)
    Me.cmdRemove.Name = "cmdRemove"
    Me.cmdRemove.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdRemove.Size = New System.Drawing.Size(72, 25)
    Me.cmdRemove.TabIndex = 1
    Me.cmdRemove.Text = "Remove"
    Me.cmdRemove.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.ToolTip1.SetToolTip(Me.cmdRemove, "Remove selected registered level")
    '
    'cmdAdd
    '
    Me.cmdAdd.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.cmdAdd.BackColor = System.Drawing.SystemColors.Control
    Me.cmdAdd.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdAdd.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdAdd.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdAdd.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdAdd.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdAdd.Location = New System.Drawing.Point(6, 318)
    Me.cmdAdd.Name = "cmdAdd"
    Me.cmdAdd.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdAdd.Size = New System.Drawing.Size(71, 25)
    Me.cmdAdd.TabIndex = 0
    Me.cmdAdd.Text = "Add"
    Me.cmdAdd.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    Me.ToolTip1.SetToolTip(Me.cmdAdd, "Add new registered level")
    '
    'Label1
    '
    Me.Label1.BackColor = System.Drawing.SystemColors.Control
    Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
    Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
    Me.Label1.Location = New System.Drawing.Point(6, 6)
    Me.Label1.Name = "Label1"
    Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.Label1.Size = New System.Drawing.Size(223, 16)
    Me.Label1.TabIndex = 4
    Me.Label1.Text = "List of Registered Level:"
    '
    'Line2
    '
    Me.Line2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Line2.BackColor = System.Drawing.SystemColors.ControlLightLight
    Me.Line2.Location = New System.Drawing.Point(6, 308)
    Me.Line2.Name = "Line2"
    Me.Line2.Size = New System.Drawing.Size(289, 1)
    Me.Line2.TabIndex = 5
    '
    'Line1
    '
    Me.Line1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.Line1.BackColor = System.Drawing.SystemColors.ControlDarkDark
    Me.Line1.Location = New System.Drawing.Point(6, 306)
    Me.Line1.Name = "Line1"
    Me.Line1.Size = New System.Drawing.Size(289, 1)
    Me.Line1.TabIndex = 6
    '
    'frmLevelManager
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.BackColor = System.Drawing.SystemColors.Control
    Me.ClientSize = New System.Drawing.Size(302, 353)
    Me.Controls.Add(Me.lstRegisteredLevel)
    Me.Controls.Add(Me.cmdClose)
    Me.Controls.Add(Me.cmdRemove)
    Me.Controls.Add(Me.cmdAdd)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.Line2)
    Me.Controls.Add(Me.Line1)
    Me.Cursor = System.Windows.Forms.Cursors.Default
    Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Location = New System.Drawing.Point(3, 28)
    Me.MaximizeBox = False
    Me.MinimizeBox = False
    Me.MinimumSize = New System.Drawing.Size(310, 380)
    Me.Name = "frmLevelManager"
    Me.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.ShowInTaskbar = False
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Level Manager"
    Me.ResumeLayout(False)

  End Sub
#End Region 

  Private Sub cmdAdd_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdAdd.Click
    Dim mdlgRegisteredLevel As New frmRegisteredLevelEdit
    With mdlgRegisteredLevel
      .Description = ""
      .ItemData = 0
      .ShowDialog()
      If Not .Annullato Then
        lstRegisteredLevel.Items.Add(New Mylist(.Description, .ItemData))
      End If
    End With
    mdlgRegisteredLevel.Close()
    mdlgRegisteredLevel = Nothing
  End Sub

  Private Sub cmdClose_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdClose.Click
    SaveComboBox(strRegisteredLevelDBName, lstRegisteredLevel.Items.GetEnumerator, True)
    Me.Hide()
  End Sub

  Private Sub cmdRemove_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRemove.Click
    On Error Resume Next
    lstRegisteredLevel.Items.RemoveAt(lstRegisteredLevel.SelectedIndex)
    lstRegisteredLevel.SelectedIndex = 0
    lstRegisteredLevel.Focus()
  End Sub

  Private Sub frmLevelManager_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
    'load images for buttons..
    LoadImages()

    lstRegisteredLevel.Items.Clear()
    LoadListBox(strRegisteredLevelDBName, lstRegisteredLevel, True)
  End Sub

  Private Sub lstRegisteredLevel_DoubleClick(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles lstRegisteredLevel.DoubleClick
    Dim mList As Mylist

    If lstRegisteredLevel.Items.Count > 0 Then
      Dim mdlgRegisteredLevel As New frmRegisteredLevelEdit
      With mdlgRegisteredLevel
        mList = CType(lstRegisteredLevel.Items(lstRegisteredLevel.SelectedIndex), Mylist)
        .ItemData = mList.ItemData
        .Description = mList.Name
        .ShowDialog()
        If Not .Annullato Then
          'lstRegisteredLevel.Items.Remove(lstRegisteredLevel.SelectedIndex)
          mList.Name = .Description
          mList.ItemData = .ItemData
          lstRegisteredLevel.Items(lstRegisteredLevel.SelectedIndex) = mList
        End If
      End With
      mdlgRegisteredLevel.Close()
      mdlgRegisteredLevel = Nothing
    End If
  End Sub

  Private Sub LoadImages()
    'load buttons images
        cmdAdd.Image = CType(frmMain.resxList("add"), Image)
        cmdRemove.Image = CType(frmMain.resxList("delete"), Image)
        cmdClose.Image = CType(frmMain.resxList("close"), Image)
  End Sub

End Class