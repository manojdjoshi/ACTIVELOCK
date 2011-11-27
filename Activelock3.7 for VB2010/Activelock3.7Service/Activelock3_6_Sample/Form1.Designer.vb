<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
		Me.cmdGetTrial = New System.Windows.Forms.Button
		Me.Button2 = New System.Windows.Forms.Button
		Me.txtInstallCode = New System.Windows.Forms.TextBox
		Me.Label1 = New System.Windows.Forms.Label
		Me.txtLicenceCode = New System.Windows.Forms.TextBox
		Me.Label2 = New System.Windows.Forms.Label
		Me.Label3 = New System.Windows.Forms.Label
		Me.txtUserName = New System.Windows.Forms.TextBox
		Me.SuspendLayout()
		'
		'cmdGetTrial
		'
		Me.cmdGetTrial.Location = New System.Drawing.Point(90, 101)
		Me.cmdGetTrial.Name = "cmdGetTrial"
		Me.cmdGetTrial.Size = New System.Drawing.Size(75, 23)
		Me.cmdGetTrial.TabIndex = 0
		Me.cmdGetTrial.Text = "Get Trial"
		Me.cmdGetTrial.UseVisualStyleBackColor = True
		'
		'Button2
		'
		Me.Button2.Location = New System.Drawing.Point(171, 101)
		Me.Button2.Name = "Button2"
		Me.Button2.Size = New System.Drawing.Size(75, 23)
		Me.Button2.TabIndex = 1
		Me.Button2.Text = "Get License"
		Me.Button2.UseVisualStyleBackColor = True
		'
		'txtInstallCode
		'
		Me.txtInstallCode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
					Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtInstallCode.Location = New System.Drawing.Point(90, 32)
		Me.txtInstallCode.Multiline = True
		Me.txtInstallCode.Name = "txtInstallCode"
		Me.txtInstallCode.ReadOnly = True
		Me.txtInstallCode.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtInstallCode.Size = New System.Drawing.Size(450, 63)
		Me.txtInstallCode.TabIndex = 2
		Me.txtInstallCode.Text = "AAAAB3NzaC1yc2EAAAABJQAAAIB8/B2KWoai2WSGTRPcgmMoczeXpd8nv0Y4r1sJ1wV3vH21q4rTpEYuB" & _
			"iD4HFOpkbNBSRdpBHJGWec7jUi8ISV0pM6i2KznjhCms5CEtYHRybbiYvRXleGzFsAAP817PLN3JYo3W" & _
			"kErT2ofR5RCkfhmx060BT8waPoqnn3AB7sZ0Q=="
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(12, 35)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(62, 13)
		Me.Label1.TabIndex = 3
		Me.Label1.Text = "Install Code"
		'
		'txtLicenceCode
		'
		Me.txtLicenceCode.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
					Or System.Windows.Forms.AnchorStyles.Left) _
					Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtLicenceCode.Location = New System.Drawing.Point(90, 130)
		Me.txtLicenceCode.Multiline = True
		Me.txtLicenceCode.Name = "txtLicenceCode"
		Me.txtLicenceCode.ReadOnly = True
		Me.txtLicenceCode.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
		Me.txtLicenceCode.Size = New System.Drawing.Size(450, 171)
		Me.txtLicenceCode.TabIndex = 8
		'
		'Label2
		'
		Me.Label2.AutoSize = True
		Me.Label2.Location = New System.Drawing.Point(12, 133)
		Me.Label2.Name = "Label2"
		Me.Label2.Size = New System.Drawing.Size(72, 13)
		Me.Label2.TabIndex = 9
		Me.Label2.Text = "License Code"
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Location = New System.Drawing.Point(12, 9)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(60, 13)
		Me.Label3.TabIndex = 10
		Me.Label3.Text = "User Name"
		'
		'txtUserName
		'
		Me.txtUserName.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
					Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.txtUserName.Location = New System.Drawing.Point(90, 6)
		Me.txtUserName.Name = "txtUserName"
		Me.txtUserName.Size = New System.Drawing.Size(450, 20)
		Me.txtUserName.TabIndex = 11
		Me.txtUserName.Text = "SampleUser"
		'
		'Form1
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(551, 313)
		Me.Controls.Add(Me.txtUserName)
		Me.Controls.Add(Me.Label3)
		Me.Controls.Add(Me.Label2)
		Me.Controls.Add(Me.txtLicenceCode)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.txtInstallCode)
		Me.Controls.Add(Me.Button2)
		Me.Controls.Add(Me.cmdGetTrial)
		Me.Name = "Form1"
		Me.Text = "Activelock36 Sample"
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
    Friend WithEvents cmdGetTrial As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents txtInstallCode As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtLicenceCode As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtUserName As System.Windows.Forms.TextBox

End Class
