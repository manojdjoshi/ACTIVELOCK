<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.txtInstCode = New System.Windows.Forms.TextBox
        Me.txtProductName = New System.Windows.Forms.TextBox
        Me.txtProductVersion = New System.Windows.Forms.TextBox
        Me.txtLicenceType = New System.Windows.Forms.TextBox
        Me.txtRegisteredLevel = New System.Windows.Forms.TextBox
        Me.txtStatus = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.ButtonRegister = New System.Windows.Forms.Button
        Me.Label8 = New System.Windows.Forms.Label
        Me.txtLicence = New System.Windows.Forms.TextBox
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtUser
        '
        Me.txtUser.Location = New System.Drawing.Point(112, 16)
        Me.txtUser.Name = "txtUser"
        Me.txtUser.Size = New System.Drawing.Size(400, 20)
        Me.txtUser.TabIndex = 0
        '
        'txtInstCode
        '
        Me.txtInstCode.Location = New System.Drawing.Point(112, 42)
        Me.txtInstCode.Multiline = True
        Me.txtInstCode.Name = "txtInstCode"
        Me.txtInstCode.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtInstCode.Size = New System.Drawing.Size(400, 68)
        Me.txtInstCode.TabIndex = 1
        '
        'txtProductName
        '
        Me.txtProductName.Location = New System.Drawing.Point(112, 194)
        Me.txtProductName.Name = "txtProductName"
        Me.txtProductName.Size = New System.Drawing.Size(400, 20)
        Me.txtProductName.TabIndex = 2
        '
        'txtProductVersion
        '
        Me.txtProductVersion.Location = New System.Drawing.Point(112, 221)
        Me.txtProductVersion.Name = "txtProductVersion"
        Me.txtProductVersion.Size = New System.Drawing.Size(400, 20)
        Me.txtProductVersion.TabIndex = 3
        '
        'txtLicenceType
        '
        Me.txtLicenceType.Location = New System.Drawing.Point(112, 247)
        Me.txtLicenceType.Name = "txtLicenceType"
        Me.txtLicenceType.Size = New System.Drawing.Size(400, 20)
        Me.txtLicenceType.TabIndex = 4
        '
        'txtRegisteredLevel
        '
        Me.txtRegisteredLevel.Location = New System.Drawing.Point(112, 273)
        Me.txtRegisteredLevel.Name = "txtRegisteredLevel"
        Me.txtRegisteredLevel.Size = New System.Drawing.Size(400, 20)
        Me.txtRegisteredLevel.TabIndex = 5
        '
        'txtStatus
        '
        Me.txtStatus.Location = New System.Drawing.Point(112, 299)
        Me.txtStatus.Name = "txtStatus"
        Me.txtStatus.Size = New System.Drawing.Size(400, 20)
        Me.txtStatus.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(5, 23)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "User:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(5, 45)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(78, 13)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "Install Code:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(5, 197)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(91, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Product Name:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(5, 225)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(101, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Product Version:"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(5, 250)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(88, 13)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "Licence Type:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(4, 277)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(107, 13)
        Me.Label6.TabIndex = 12
        Me.Label6.Text = "Registered Level:"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(5, 306)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(47, 13)
        Me.Label7.TabIndex = 13
        Me.Label7.Text = "Status:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.ButtonRegister)
        Me.GroupBox1.Controls.Add(Me.Label8)
        Me.GroupBox1.Controls.Add(Me.txtLicence)
        Me.GroupBox1.Controls.Add(Me.txtInstCode)
        Me.GroupBox1.Controls.Add(Me.Label7)
        Me.GroupBox1.Controls.Add(Me.txtUser)
        Me.GroupBox1.Controls.Add(Me.txtStatus)
        Me.GroupBox1.Controls.Add(Me.Label6)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.txtRegisteredLevel)
        Me.GroupBox1.Controls.Add(Me.txtProductName)
        Me.GroupBox1.Controls.Add(Me.txtLicenceType)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.txtProductVersion)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(526, 333)
        Me.GroupBox1.TabIndex = 14
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Registration [Type your name to start]"
        '
        'ButtonRegister
        '
        Me.ButtonRegister.Location = New System.Drawing.Point(11, 132)
        Me.ButtonRegister.Name = "ButtonRegister"
        Me.ButtonRegister.Size = New System.Drawing.Size(95, 52)
        Me.ButtonRegister.TabIndex = 16
        Me.ButtonRegister.Text = "REGISTER"
        Me.ButtonRegister.UseVisualStyleBackColor = True
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(5, 116)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(79, 13)
        Me.Label8.TabIndex = 15
        Me.Label8.Text = "License key:"
        '
        'txtLicence
        '
        Me.txtLicence.Location = New System.Drawing.Point(112, 116)
        Me.txtLicence.Multiline = True
        Me.txtLicence.Name = "txtLicence"
        Me.txtLicence.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtLicence.Size = New System.Drawing.Size(400, 68)
        Me.txtLicence.TabIndex = 14
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(548, 420)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = " Register Form"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents txtUser As System.Windows.Forms.TextBox
    Friend WithEvents txtInstCode As System.Windows.Forms.TextBox
    Friend WithEvents txtProductName As System.Windows.Forms.TextBox
    Friend WithEvents txtProductVersion As System.Windows.Forms.TextBox
    Friend WithEvents txtLicenceType As System.Windows.Forms.TextBox
    Friend WithEvents txtRegisteredLevel As System.Windows.Forms.TextBox
    Friend WithEvents txtStatus As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents txtLicence As System.Windows.Forms.TextBox
    Friend WithEvents ButtonRegister As System.Windows.Forms.Button
End Class
