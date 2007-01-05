<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LicenseManager
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lvLicenses = New System.Windows.Forms.ListView
        Me.User_Name = New System.Windows.Forms.ColumnHeader
        Me.Machine_Name = New System.Windows.Forms.ColumnHeader
        Me.IP_Address = New System.Windows.Forms.ColumnHeader
        Me.License_Date = New System.Windows.Forms.ColumnHeader
        Me.MAC_Address = New System.Windows.Forms.ColumnHeader
        Me.bntRemove = New System.Windows.Forms.Button
        Me.bntRemoveAll = New System.Windows.Forms.Button
        Me.bntExit = New System.Windows.Forms.Button
        Me.bntRefresh = New System.Windows.Forms.Button
        Me.lblLicenseStatus = New System.Windows.Forms.Label
        Me.txtRequestCode = New System.Windows.Forms.RichTextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.txtLicenseCode = New System.Windows.Forms.RichTextBox
        Me.bntGenerate = New System.Windows.Forms.Button
        Me.bntRegister = New System.Windows.Forms.Button
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'lvLicenses
        '
        Me.lvLicenses.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.MAC_Address, Me.User_Name, Me.Machine_Name, Me.IP_Address, Me.License_Date})
        Me.lvLicenses.FullRowSelect = True
        Me.lvLicenses.GridLines = True
        Me.lvLicenses.Location = New System.Drawing.Point(12, 53)
        Me.lvLicenses.MultiSelect = False
        Me.lvLicenses.Name = "lvLicenses"
        Me.lvLicenses.Size = New System.Drawing.Size(507, 211)
        Me.lvLicenses.TabIndex = 0
        Me.lvLicenses.UseCompatibleStateImageBehavior = False
        Me.lvLicenses.View = System.Windows.Forms.View.Details
        '
        'User_Name
        '
        Me.User_Name.Text = "User Name"
        Me.User_Name.Width = 105
        '
        'Machine_Name
        '
        Me.Machine_Name.Text = "Machine Name"
        Me.Machine_Name.Width = 114
        '
        'IP_Address
        '
        Me.IP_Address.Text = "IP Address"
        Me.IP_Address.Width = 116
        '
        'License_Date
        '
        Me.License_Date.Text = "License Date"
        Me.License_Date.Width = 148
        '
        'MAC_Address
        '
        Me.MAC_Address.Text = "MAC Address"
        Me.MAC_Address.Width = 0
        '
        'bntRemove
        '
        Me.bntRemove.Location = New System.Drawing.Point(534, 134)
        Me.bntRemove.Name = "bntRemove"
        Me.bntRemove.Size = New System.Drawing.Size(75, 23)
        Me.bntRemove.TabIndex = 1
        Me.bntRemove.Text = "Remove"
        Me.bntRemove.UseVisualStyleBackColor = True
        '
        'bntRemoveAll
        '
        Me.bntRemoveAll.Location = New System.Drawing.Point(534, 172)
        Me.bntRemoveAll.Name = "bntRemoveAll"
        Me.bntRemoveAll.Size = New System.Drawing.Size(75, 23)
        Me.bntRemoveAll.TabIndex = 2
        Me.bntRemoveAll.Text = "Remove All"
        Me.bntRemoveAll.UseVisualStyleBackColor = True
        '
        'bntExit
        '
        Me.bntExit.Location = New System.Drawing.Point(534, 12)
        Me.bntExit.Name = "bntExit"
        Me.bntExit.Size = New System.Drawing.Size(75, 23)
        Me.bntExit.TabIndex = 3
        Me.bntExit.Text = "Exit"
        Me.bntExit.UseVisualStyleBackColor = True
        '
        'bntRefresh
        '
        Me.bntRefresh.Location = New System.Drawing.Point(534, 53)
        Me.bntRefresh.Name = "bntRefresh"
        Me.bntRefresh.Size = New System.Drawing.Size(75, 23)
        Me.bntRefresh.TabIndex = 4
        Me.bntRefresh.Text = "Refresh"
        Me.bntRefresh.UseVisualStyleBackColor = True
        '
        'lblLicenseStatus
        '
        Me.lblLicenseStatus.AutoSize = True
        Me.lblLicenseStatus.Location = New System.Drawing.Point(9, 12)
        Me.lblLicenseStatus.Name = "lblLicenseStatus"
        Me.lblLicenseStatus.Size = New System.Drawing.Size(128, 13)
        Me.lblLicenseStatus.TabIndex = 5
        Me.lblLicenseStatus.Text = "Really long label text here"
        '
        'txtRequestCode
        '
        Me.txtRequestCode.Location = New System.Drawing.Point(97, 49)
        Me.txtRequestCode.Name = "txtRequestCode"
        Me.txtRequestCode.ReadOnly = True
        Me.txtRequestCode.Size = New System.Drawing.Size(310, 58)
        Me.txtRequestCode.TabIndex = 6
        Me.txtRequestCode.Text = ""
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(6, 52)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(85, 13)
        Me.Label2.TabIndex = 26
        Me.Label2.Text = "Installation Code"
        '
        'txtUser
        '
        Me.txtUser.Location = New System.Drawing.Point(97, 23)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(310, 20)
        Me.txtUser.TabIndex = 27
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(31, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 28
        Me.Label1.Text = "User Name"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(26, 116)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 13)
        Me.Label3.TabIndex = 29
        Me.Label3.Text = "License Key"
        '
        'txtLicenseCode
        '
        Me.txtLicenseCode.Location = New System.Drawing.Point(97, 113)
        Me.txtLicenseCode.Name = "txtLicenseCode"
        Me.txtLicenseCode.Size = New System.Drawing.Size(310, 58)
        Me.txtLicenseCode.TabIndex = 30
        Me.txtLicenseCode.Text = ""
        '
        'bntGenerate
        '
        Me.bntGenerate.Location = New System.Drawing.Point(413, 52)
        Me.bntGenerate.Name = "bntGenerate"
        Me.bntGenerate.Size = New System.Drawing.Size(75, 23)
        Me.bntGenerate.TabIndex = 31
        Me.bntGenerate.Text = "Generate"
        Me.bntGenerate.UseVisualStyleBackColor = True
        '
        'bntRegister
        '
        Me.bntRegister.Location = New System.Drawing.Point(413, 119)
        Me.bntRegister.Name = "bntRegister"
        Me.bntRegister.Size = New System.Drawing.Size(75, 23)
        Me.bntRegister.TabIndex = 32
        Me.bntRegister.Text = "Register"
        Me.bntRegister.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.txtRequestCode)
        Me.GroupBox1.Controls.Add(Me.bntRegister)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.bntGenerate)
        Me.GroupBox1.Controls.Add(Me.txtUser)
        Me.GroupBox1.Controls.Add(Me.txtLicenseCode)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 280)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(507, 187)
        Me.GroupBox1.TabIndex = 33
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Registration"
        Me.GroupBox1.Visible = False
        '
        'LicenseServer
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(621, 503)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.lblLicenseStatus)
        Me.Controls.Add(Me.bntRefresh)
        Me.Controls.Add(Me.bntExit)
        Me.Controls.Add(Me.bntRemoveAll)
        Me.Controls.Add(Me.bntRemove)
        Me.Controls.Add(Me.lvLicenses)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "LicenseServer"
        Me.Text = "License Manager"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lvLicenses As System.Windows.Forms.ListView
    Friend WithEvents bntRemove As System.Windows.Forms.Button
    Friend WithEvents bntRemoveAll As System.Windows.Forms.Button
    Friend WithEvents bntExit As System.Windows.Forms.Button
    Friend WithEvents bntRefresh As System.Windows.Forms.Button
    Friend WithEvents lblLicenseStatus As System.Windows.Forms.Label
    Friend WithEvents txtRequestCode As System.Windows.Forms.RichTextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtUser As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtLicenseCode As System.Windows.Forms.RichTextBox
    Friend WithEvents bntGenerate As System.Windows.Forms.Button
    Friend WithEvents bntRegister As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Machine_Name As System.Windows.Forms.ColumnHeader
    Friend WithEvents IP_Address As System.Windows.Forms.ColumnHeader
    Friend WithEvents User_Name As System.Windows.Forms.ColumnHeader
    Friend WithEvents License_Date As System.Windows.Forms.ColumnHeader
    Friend WithEvents MAC_Address As System.Windows.Forms.ColumnHeader

End Class
