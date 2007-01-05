<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Client
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
        Me.btnRelease = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'btnRelease
        '
        Me.btnRelease.Location = New System.Drawing.Point(27, 99)
        Me.btnRelease.Name = "btnRelease"
        Me.btnRelease.Size = New System.Drawing.Size(146, 23)
        Me.btnRelease.TabIndex = 0
        Me.btnRelease.Text = "Release License"
        Me.btnRelease.UseVisualStyleBackColor = True
        '
        'Client
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(635, 318)
        Me.Controls.Add(Me.btnRelease)
        Me.Name = "Client"
        Me.Text = "License Client"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnRelease As System.Windows.Forms.Button

End Class
