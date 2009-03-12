Option Strict Off
Option Explicit On 
Imports System.IO
Imports VB = Microsoft.VisualBasic
Friend Class frmMain
	Inherits System.Windows.Forms.Form
#Region "Windows Form Designer generated code "
    Public Sub New()

        MyBase.New()
        If m_vb6FormDefInstance Is Nothing Then
            If m_InitializingDefInstance Then
                m_vb6FormDefInstance = Me
            Else
                Try
                    'For the start-up form, the first instance created is the default instance.
                    If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
                        m_vb6FormDefInstance = Me
                    End If
                Catch
                End Try
            End If
        End If
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
    'Public WithEvents sbStatus As AxComctlLib.AxStatusBar
	Public WithEvents txtLicenseType As System.Windows.Forms.TextBox
	Public WithEvents cmdResetTrial As System.Windows.Forms.Button
	Public WithEvents cmdKillTrial As System.Windows.Forms.Button
	Public WithEvents Picture2 As System.Windows.Forms.PictureBox
	Public WithEvents txtRegisteredLevel As System.Windows.Forms.TextBox
	Public WithEvents txtChecksum As System.Windows.Forms.TextBox
	Public WithEvents txtVersion As System.Windows.Forms.TextBox
	Public WithEvents txtName As System.Windows.Forms.TextBox
	Public WithEvents txtExpiration As System.Windows.Forms.TextBox
	Public WithEvents txtUsedDays As System.Windows.Forms.TextBox
	Public WithEvents txtRegStatus As System.Windows.Forms.TextBox
	Public WithEvents Label9 As System.Windows.Forms.Label
	Public WithEvents Label16 As System.Windows.Forms.Label
	Public WithEvents Label5 As System.Windows.Forms.Label
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents Label2 As System.Windows.Forms.Label
	Public WithEvents Label1 As System.Windows.Forms.Label
	Public WithEvents Label8 As System.Windows.Forms.Label
	Public WithEvents Label7 As System.Windows.Forms.Label
	Public WithEvents Label6 As System.Windows.Forms.Label
	Public WithEvents fraRegStatus As System.Windows.Forms.GroupBox
	Public WithEvents cmdPaste As System.Windows.Forms.Button
	Public WithEvents cmdCopy As System.Windows.Forms.Button
	Public WithEvents cmdKillLicense As System.Windows.Forms.Button
	Public WithEvents txtUser As System.Windows.Forms.TextBox
	Public WithEvents cmdReqGen As System.Windows.Forms.Button
	Public WithEvents txtReqCodeGen As System.Windows.Forms.TextBox
	Public WithEvents cmdRegister As System.Windows.Forms.Button
	Public WithEvents txtLibKeyIn As System.Windows.Forms.TextBox
	Public WithEvents Label13 As System.Windows.Forms.Label
	Public WithEvents Label11 As System.Windows.Forms.Label
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents fraReg As System.Windows.Forms.GroupBox
	Public WithEvents _SSTab1_TabPage0 As System.Windows.Forms.TabPage
	Public WithEvents lblLockStatus As System.Windows.Forms.Label
	Public WithEvents lblLockStatus2 As System.Windows.Forms.Label
	Public WithEvents lblTrialInfo As System.Windows.Forms.Label
	Public WithEvents _optForm_1 As System.Windows.Forms.RadioButton
	Public WithEvents _optForm_0 As System.Windows.Forms.RadioButton
	Public WithEvents cboSpeed As System.Windows.Forms.ComboBox
	Public WithEvents chkPause As System.Windows.Forms.CheckBox
	Public WithEvents chkFlash As System.Windows.Forms.CheckBox
	Public WithEvents chkScroll As System.Windows.Forms.CheckBox
	Public WithEvents lblHost As System.Windows.Forms.Label
	Public WithEvents lblSpeed As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.Panel
    Public WithEvents _SSTab1_TabPage1 As System.Windows.Forms.TabPage
    Public WithEvents SSTab1 As System.Windows.Forms.TabControl
    Public WithEvents optForm As Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray
    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.
    'Do not modify it using the code editor.
    Public WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents txtNetworkedLicense As System.Windows.Forms.TextBox
    Public WithEvents txtMaxCount As System.Windows.Forms.TextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Public WithEvents lblConcurrentUsers As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmdResetTrial = New System.Windows.Forms.Button
        Me.cmdKillTrial = New System.Windows.Forms.Button
        Me.cmdKillLicense = New System.Windows.Forms.Button
        Me.cmdReqGen = New System.Windows.Forms.Button
        Me.cmdRegister = New System.Windows.Forms.Button
        Me.SSTab1 = New System.Windows.Forms.TabControl
        Me._SSTab1_TabPage0 = New System.Windows.Forms.TabPage
        Me.fraRegStatus = New System.Windows.Forms.GroupBox
        Me.lblConcurrentUsers = New System.Windows.Forms.Label
        Me.txtMaxCount = New System.Windows.Forms.TextBox
        Me.txtNetworkedLicense = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.txtLicenseType = New System.Windows.Forms.TextBox
        Me.Picture2 = New System.Windows.Forms.PictureBox
        Me.txtRegisteredLevel = New System.Windows.Forms.TextBox
        Me.txtChecksum = New System.Windows.Forms.TextBox
        Me.txtVersion = New System.Windows.Forms.TextBox
        Me.txtName = New System.Windows.Forms.TextBox
        Me.txtExpiration = New System.Windows.Forms.TextBox
        Me.txtUsedDays = New System.Windows.Forms.TextBox
        Me.txtRegStatus = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.Label16 = New System.Windows.Forms.Label
        Me.Label5 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label8 = New System.Windows.Forms.Label
        Me.Label7 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.fraReg = New System.Windows.Forms.GroupBox
        Me.cmdPaste = New System.Windows.Forms.Button
        Me.cmdCopy = New System.Windows.Forms.Button
        Me.txtUser = New System.Windows.Forms.TextBox
        Me.txtReqCodeGen = New System.Windows.Forms.TextBox
        Me.txtLibKeyIn = New System.Windows.Forms.TextBox
        Me.Label13 = New System.Windows.Forms.Label
        Me.Label11 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me._SSTab1_TabPage1 = New System.Windows.Forms.TabPage
        Me.lblLockStatus = New System.Windows.Forms.Label
        Me.lblLockStatus2 = New System.Windows.Forms.Label
        Me.lblTrialInfo = New System.Windows.Forms.Label
        Me.Frame1 = New System.Windows.Forms.Panel
        Me._optForm_1 = New System.Windows.Forms.RadioButton
        Me._optForm_0 = New System.Windows.Forms.RadioButton
        Me.cboSpeed = New System.Windows.Forms.ComboBox
        Me.chkPause = New System.Windows.Forms.CheckBox
        Me.chkFlash = New System.Windows.Forms.CheckBox
        Me.chkScroll = New System.Windows.Forms.CheckBox
        Me.lblHost = New System.Windows.Forms.Label
        Me.lblSpeed = New System.Windows.Forms.Label
        Me.optForm = New Microsoft.VisualBasic.Compatibility.VB6.RadioButtonArray(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SSTab1.SuspendLayout()
        Me._SSTab1_TabPage0.SuspendLayout()
        Me.fraRegStatus.SuspendLayout()
        CType(Me.Picture2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.fraReg.SuspendLayout()
        Me._SSTab1_TabPage1.SuspendLayout()
        Me.Frame1.SuspendLayout()
        CType(Me.optForm, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'cmdResetTrial
        '
        Me.cmdResetTrial.BackColor = System.Drawing.SystemColors.Control
        Me.cmdResetTrial.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdResetTrial.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdResetTrial.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdResetTrial.Location = New System.Drawing.Point(552, 88)
        Me.cmdResetTrial.Name = "cmdResetTrial"
        Me.cmdResetTrial.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdResetTrial.Size = New System.Drawing.Size(73, 21)
        Me.cmdResetTrial.TabIndex = 42
        Me.cmdResetTrial.Text = "&Reset Trial"
        Me.ToolTip1.SetToolTip(Me.cmdResetTrial, "Reset the Free Trial")
        Me.cmdResetTrial.UseVisualStyleBackColor = False
        '
        'cmdKillTrial
        '
        Me.cmdKillTrial.BackColor = System.Drawing.SystemColors.Control
        Me.cmdKillTrial.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdKillTrial.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdKillTrial.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdKillTrial.Location = New System.Drawing.Point(552, 116)
        Me.cmdKillTrial.Name = "cmdKillTrial"
        Me.cmdKillTrial.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdKillTrial.Size = New System.Drawing.Size(73, 21)
        Me.cmdKillTrial.TabIndex = 41
        Me.cmdKillTrial.Text = "&Kill Trial"
        Me.ToolTip1.SetToolTip(Me.cmdKillTrial, "End the Free Trial")
        Me.cmdKillTrial.UseVisualStyleBackColor = False
        '
        'cmdKillLicense
        '
        Me.cmdKillLicense.BackColor = System.Drawing.SystemColors.Control
        Me.cmdKillLicense.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdKillLicense.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdKillLicense.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdKillLicense.Location = New System.Drawing.Point(548, 226)
        Me.cmdKillLicense.Name = "cmdKillLicense"
        Me.cmdKillLicense.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdKillLicense.Size = New System.Drawing.Size(73, 21)
        Me.cmdKillLicense.TabIndex = 43
        Me.cmdKillLicense.Text = "&Kill License"
        Me.ToolTip1.SetToolTip(Me.cmdKillLicense, "Kill the License")
        Me.cmdKillLicense.UseVisualStyleBackColor = False
        '
        'cmdReqGen
        '
        Me.cmdReqGen.BackColor = System.Drawing.SystemColors.Control
        Me.cmdReqGen.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdReqGen.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdReqGen.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdReqGen.Location = New System.Drawing.Point(548, 36)
        Me.cmdReqGen.Name = "cmdReqGen"
        Me.cmdReqGen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdReqGen.Size = New System.Drawing.Size(73, 21)
        Me.cmdReqGen.TabIndex = 9
        Me.cmdReqGen.Text = "&Generate"
        Me.ToolTip1.SetToolTip(Me.cmdReqGen, "Generate Installation Code")
        Me.cmdReqGen.UseVisualStyleBackColor = False
        '
        'cmdRegister
        '
        Me.cmdRegister.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRegister.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRegister.Enabled = False
        Me.cmdRegister.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdRegister.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRegister.Location = New System.Drawing.Point(548, 198)
        Me.cmdRegister.Name = "cmdRegister"
        Me.cmdRegister.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRegister.Size = New System.Drawing.Size(73, 21)
        Me.cmdRegister.TabIndex = 11
        Me.cmdRegister.Text = "&Register"
        Me.ToolTip1.SetToolTip(Me.cmdRegister, "Register the License")
        Me.cmdRegister.UseVisualStyleBackColor = False
        '
        'SSTab1
        '
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage0)
        Me.SSTab1.Controls.Add(Me._SSTab1_TabPage1)
        Me.SSTab1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SSTab1.ForeColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.SSTab1.ItemSize = New System.Drawing.Size(42, 18)
        Me.SSTab1.Location = New System.Drawing.Point(0, 0)
        Me.SSTab1.Name = "SSTab1"
        Me.SSTab1.SelectedIndex = 0
        Me.SSTab1.Size = New System.Drawing.Size(649, 641)
        Me.SSTab1.TabIndex = 12
        '
        '_SSTab1_TabPage0
        '
        Me._SSTab1_TabPage0.Controls.Add(Me.fraRegStatus)
        Me._SSTab1_TabPage0.Controls.Add(Me.fraReg)
        Me._SSTab1_TabPage0.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage0.Name = "_SSTab1_TabPage0"
        Me._SSTab1_TabPage0.Size = New System.Drawing.Size(641, 615)
        Me._SSTab1_TabPage0.TabIndex = 0
        Me._SSTab1_TabPage0.Text = "Registration"
        '
        'fraRegStatus
        '
        Me.fraRegStatus.BackColor = System.Drawing.SystemColors.Control
        Me.fraRegStatus.Controls.Add(Me.lblConcurrentUsers)
        Me.fraRegStatus.Controls.Add(Me.txtMaxCount)
        Me.fraRegStatus.Controls.Add(Me.txtNetworkedLicense)
        Me.fraRegStatus.Controls.Add(Me.Label10)
        Me.fraRegStatus.Controls.Add(Me.txtLicenseType)
        Me.fraRegStatus.Controls.Add(Me.cmdResetTrial)
        Me.fraRegStatus.Controls.Add(Me.cmdKillTrial)
        Me.fraRegStatus.Controls.Add(Me.Picture2)
        Me.fraRegStatus.Controls.Add(Me.txtRegisteredLevel)
        Me.fraRegStatus.Controls.Add(Me.txtChecksum)
        Me.fraRegStatus.Controls.Add(Me.txtVersion)
        Me.fraRegStatus.Controls.Add(Me.txtName)
        Me.fraRegStatus.Controls.Add(Me.txtExpiration)
        Me.fraRegStatus.Controls.Add(Me.txtUsedDays)
        Me.fraRegStatus.Controls.Add(Me.txtRegStatus)
        Me.fraRegStatus.Controls.Add(Me.Label9)
        Me.fraRegStatus.Controls.Add(Me.Label16)
        Me.fraRegStatus.Controls.Add(Me.Label5)
        Me.fraRegStatus.Controls.Add(Me.Label3)
        Me.fraRegStatus.Controls.Add(Me.Label2)
        Me.fraRegStatus.Controls.Add(Me.Label1)
        Me.fraRegStatus.Controls.Add(Me.Label8)
        Me.fraRegStatus.Controls.Add(Me.Label7)
        Me.fraRegStatus.Controls.Add(Me.Label6)
        Me.fraRegStatus.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraRegStatus.ForeColor = System.Drawing.Color.Blue
        Me.fraRegStatus.Location = New System.Drawing.Point(8, 28)
        Me.fraRegStatus.Name = "fraRegStatus"
        Me.fraRegStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraRegStatus.Size = New System.Drawing.Size(633, 182)
        Me.fraRegStatus.TabIndex = 13
        Me.fraRegStatus.TabStop = False
        Me.fraRegStatus.Text = "Status"
        '
        'lblConcurrentUsers
        '
        Me.lblConcurrentUsers.BackColor = System.Drawing.SystemColors.Control
        Me.lblConcurrentUsers.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblConcurrentUsers.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblConcurrentUsers.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblConcurrentUsers.Location = New System.Drawing.Point(230, 158)
        Me.lblConcurrentUsers.Name = "lblConcurrentUsers"
        Me.lblConcurrentUsers.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblConcurrentUsers.Size = New System.Drawing.Size(126, 17)
        Me.lblConcurrentUsers.TabIndex = 49
        Me.lblConcurrentUsers.Text = "No. of Concurrent Users:"
        '
        'txtMaxCount
        '
        Me.txtMaxCount.AcceptsReturn = True
        Me.txtMaxCount.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtMaxCount.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtMaxCount.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtMaxCount.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtMaxCount.Location = New System.Drawing.Point(358, 156)
        Me.txtMaxCount.MaxLength = 0
        Me.txtMaxCount.Name = "txtMaxCount"
        Me.txtMaxCount.ReadOnly = True
        Me.txtMaxCount.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtMaxCount.Size = New System.Drawing.Size(36, 20)
        Me.txtMaxCount.TabIndex = 48
        '
        'txtNetworkedLicense
        '
        Me.txtNetworkedLicense.AcceptsReturn = True
        Me.txtNetworkedLicense.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtNetworkedLicense.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNetworkedLicense.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNetworkedLicense.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNetworkedLicense.Location = New System.Drawing.Point(104, 156)
        Me.txtNetworkedLicense.MaxLength = 0
        Me.txtNetworkedLicense.Name = "txtNetworkedLicense"
        Me.txtNetworkedLicense.ReadOnly = True
        Me.txtNetworkedLicense.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNetworkedLicense.Size = New System.Drawing.Size(120, 20)
        Me.txtNetworkedLicense.TabIndex = 47
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.SystemColors.Control
        Me.Label10.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label10.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label10.Location = New System.Drawing.Point(8, 158)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(89, 17)
        Me.Label10.TabIndex = 46
        Me.Label10.Text = "License Class:"
        '
        'txtLicenseType
        '
        Me.txtLicenseType.AcceptsReturn = True
        Me.txtLicenseType.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtLicenseType.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLicenseType.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLicenseType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtLicenseType.Location = New System.Drawing.Point(276, 136)
        Me.txtLicenseType.MaxLength = 0
        Me.txtLicenseType.Name = "txtLicenseType"
        Me.txtLicenseType.ReadOnly = True
        Me.txtLicenseType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLicenseType.Size = New System.Drawing.Size(117, 20)
        Me.txtLicenseType.TabIndex = 45
        '
        'Picture2
        '
        Me.Picture2.BackColor = System.Drawing.SystemColors.Window
        Me.Picture2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Picture2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Picture2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.Picture2.Image = CType(resources.GetObject("Picture2.Image"), System.Drawing.Image)
        Me.Picture2.Location = New System.Drawing.Point(558, 14)
        Me.Picture2.Name = "Picture2"
        Me.Picture2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Picture2.Size = New System.Drawing.Size(55, 55)
        Me.Picture2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize
        Me.Picture2.TabIndex = 39
        Me.Picture2.TabStop = False
        '
        'txtRegisteredLevel
        '
        Me.txtRegisteredLevel.AcceptsReturn = True
        Me.txtRegisteredLevel.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtRegisteredLevel.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRegisteredLevel.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRegisteredLevel.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRegisteredLevel.Location = New System.Drawing.Point(104, 116)
        Me.txtRegisteredLevel.MaxLength = 0
        Me.txtRegisteredLevel.Name = "txtRegisteredLevel"
        Me.txtRegisteredLevel.ReadOnly = True
        Me.txtRegisteredLevel.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRegisteredLevel.Size = New System.Drawing.Size(289, 20)
        Me.txtRegisteredLevel.TabIndex = 36
        '
        'txtChecksum
        '
        Me.txtChecksum.AcceptsReturn = True
        Me.txtChecksum.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtChecksum.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtChecksum.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtChecksum.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtChecksum.Location = New System.Drawing.Point(104, 136)
        Me.txtChecksum.MaxLength = 0
        Me.txtChecksum.Name = "txtChecksum"
        Me.txtChecksum.ReadOnly = True
        Me.txtChecksum.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtChecksum.Size = New System.Drawing.Size(81, 20)
        Me.txtChecksum.TabIndex = 6
        '
        'txtVersion
        '
        Me.txtVersion.AcceptsReturn = True
        Me.txtVersion.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtVersion.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtVersion.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtVersion.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtVersion.Location = New System.Drawing.Point(104, 36)
        Me.txtVersion.MaxLength = 0
        Me.txtVersion.Name = "txtVersion"
        Me.txtVersion.ReadOnly = True
        Me.txtVersion.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtVersion.Size = New System.Drawing.Size(289, 20)
        Me.txtVersion.TabIndex = 2
        Me.txtVersion.Text = "1.0"
        '
        'txtName
        '
        Me.txtName.AcceptsReturn = True
        Me.txtName.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtName.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtName.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtName.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtName.Location = New System.Drawing.Point(104, 16)
        Me.txtName.MaxLength = 0
        Me.txtName.Name = "txtName"
        Me.txtName.ReadOnly = True
        Me.txtName.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtName.Size = New System.Drawing.Size(289, 20)
        Me.txtName.TabIndex = 1
        Me.txtName.Text = "TestApp"
        '
        'txtExpiration
        '
        Me.txtExpiration.AcceptsReturn = True
        Me.txtExpiration.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtExpiration.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtExpiration.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtExpiration.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtExpiration.Location = New System.Drawing.Point(104, 96)
        Me.txtExpiration.MaxLength = 0
        Me.txtExpiration.Name = "txtExpiration"
        Me.txtExpiration.ReadOnly = True
        Me.txtExpiration.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtExpiration.Size = New System.Drawing.Size(289, 20)
        Me.txtExpiration.TabIndex = 5
        '
        'txtUsedDays
        '
        Me.txtUsedDays.AcceptsReturn = True
        Me.txtUsedDays.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtUsedDays.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUsedDays.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUsedDays.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUsedDays.Location = New System.Drawing.Point(104, 76)
        Me.txtUsedDays.MaxLength = 0
        Me.txtUsedDays.Name = "txtUsedDays"
        Me.txtUsedDays.ReadOnly = True
        Me.txtUsedDays.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUsedDays.Size = New System.Drawing.Size(289, 20)
        Me.txtUsedDays.TabIndex = 4
        '
        'txtRegStatus
        '
        Me.txtRegStatus.AcceptsReturn = True
        Me.txtRegStatus.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtRegStatus.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtRegStatus.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRegStatus.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtRegStatus.Location = New System.Drawing.Point(104, 56)
        Me.txtRegStatus.MaxLength = 0
        Me.txtRegStatus.Name = "txtRegStatus"
        Me.txtRegStatus.ReadOnly = True
        Me.txtRegStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtRegStatus.Size = New System.Drawing.Size(289, 20)
        Me.txtRegStatus.TabIndex = 3
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.SystemColors.Control
        Me.Label9.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label9.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label9.Location = New System.Drawing.Point(196, 138)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(77, 17)
        Me.Label9.TabIndex = 44
        Me.Label9.Text = "License Type:"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'Label16
        '
        Me.Label16.BackColor = System.Drawing.SystemColors.Control
        Me.Label16.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label16.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.ForeColor = System.Drawing.Color.Blue
        Me.Label16.Location = New System.Drawing.Point(550, 70)
        Me.Label16.Name = "Label16"
        Me.Label16.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label16.Size = New System.Drawing.Size(71, 11)
        Me.Label16.TabIndex = 40
        Me.Label16.Text = "Activelock V3"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'Label5
        '
        Me.Label5.BackColor = System.Drawing.SystemColors.Control
        Me.Label5.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label5.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label5.Location = New System.Drawing.Point(8, 118)
        Me.Label5.Name = "Label5"
        Me.Label5.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label5.Size = New System.Drawing.Size(92, 17)
        Me.Label5.TabIndex = 37
        Me.Label5.Text = "Registered Level:"
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.SystemColors.Control
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label3.Location = New System.Drawing.Point(8, 138)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(89, 17)
        Me.Label3.TabIndex = 35
        Me.Label3.Text = "DLL Checksum:"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label2.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label2.Location = New System.Drawing.Point(8, 38)
        Me.Label2.Name = "Label2"
        Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label2.Size = New System.Drawing.Size(88, 17)
        Me.Label2.TabIndex = 32
        Me.Label2.Text = "App Version:"
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.SystemColors.Control
        Me.Label1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(8, 18)
        Me.Label1.Name = "Label1"
        Me.Label1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label1.Size = New System.Drawing.Size(65, 17)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "App Name:"
        '
        'Label8
        '
        Me.Label8.BackColor = System.Drawing.SystemColors.Control
        Me.Label8.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label8.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label8.Location = New System.Drawing.Point(8, 98)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(65, 17)
        Me.Label8.TabIndex = 16
        Me.Label8.Text = "Expiry Date:"
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.SystemColors.Control
        Me.Label7.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label7.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label7.Location = New System.Drawing.Point(8, 78)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(65, 17)
        Me.Label7.TabIndex = 15
        Me.Label7.Text = "Days Used:"
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.SystemColors.Control
        Me.Label6.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label6.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label6.Location = New System.Drawing.Point(8, 58)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(92, 17)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "License Status:"
        '
        'fraReg
        '
        Me.fraReg.BackColor = System.Drawing.SystemColors.Control
        Me.fraReg.Controls.Add(Me.cmdPaste)
        Me.fraReg.Controls.Add(Me.cmdCopy)
        Me.fraReg.Controls.Add(Me.cmdKillLicense)
        Me.fraReg.Controls.Add(Me.txtUser)
        Me.fraReg.Controls.Add(Me.cmdReqGen)
        Me.fraReg.Controls.Add(Me.txtReqCodeGen)
        Me.fraReg.Controls.Add(Me.cmdRegister)
        Me.fraReg.Controls.Add(Me.txtLibKeyIn)
        Me.fraReg.Controls.Add(Me.Label13)
        Me.fraReg.Controls.Add(Me.Label11)
        Me.fraReg.Controls.Add(Me.Label4)
        Me.fraReg.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.fraReg.ForeColor = System.Drawing.Color.Blue
        Me.fraReg.Location = New System.Drawing.Point(8, 216)
        Me.fraReg.Name = "fraReg"
        Me.fraReg.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.fraReg.Size = New System.Drawing.Size(633, 396)
        Me.fraReg.TabIndex = 17
        Me.fraReg.TabStop = False
        Me.fraReg.Text = "Register"
        '
        'cmdPaste
        '
        Me.cmdPaste.BackColor = System.Drawing.SystemColors.Control
        Me.cmdPaste.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdPaste.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdPaste.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdPaste.Image = CType(resources.GetObject("cmdPaste.Image"), System.Drawing.Image)
        Me.cmdPaste.Location = New System.Drawing.Point(548, 170)
        Me.cmdPaste.Name = "cmdPaste"
        Me.cmdPaste.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdPaste.Size = New System.Drawing.Size(23, 23)
        Me.cmdPaste.TabIndex = 47
        Me.cmdPaste.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdPaste.UseVisualStyleBackColor = False
        '
        'cmdCopy
        '
        Me.cmdCopy.BackColor = System.Drawing.SystemColors.Control
        Me.cmdCopy.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdCopy.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmdCopy.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdCopy.Image = CType(resources.GetObject("cmdCopy.Image"), System.Drawing.Image)
        Me.cmdCopy.Location = New System.Drawing.Point(548, 60)
        Me.cmdCopy.Name = "cmdCopy"
        Me.cmdCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdCopy.Size = New System.Drawing.Size(23, 23)
        Me.cmdCopy.TabIndex = 46
        Me.cmdCopy.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.cmdCopy.UseVisualStyleBackColor = False
        '
        'txtUser
        '
        Me.txtUser.AcceptsReturn = True
        Me.txtUser.BackColor = System.Drawing.SystemColors.Window
        Me.txtUser.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtUser.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtUser.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtUser.Location = New System.Drawing.Point(96, 20)
        Me.txtUser.MaxLength = 0
        Me.txtUser.Name = "txtUser"
        Me.txtUser.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtUser.Size = New System.Drawing.Size(445, 20)
        Me.txtUser.TabIndex = 7
        Me.txtUser.Text = "Evaluation User"
        '
        'txtReqCodeGen
        '
        Me.txtReqCodeGen.AcceptsReturn = True
        Me.txtReqCodeGen.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtReqCodeGen.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtReqCodeGen.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtReqCodeGen.ForeColor = System.Drawing.Color.Green
        Me.txtReqCodeGen.Location = New System.Drawing.Point(96, 40)
        Me.txtReqCodeGen.MaxLength = 0
        Me.txtReqCodeGen.Multiline = True
        Me.txtReqCodeGen.Name = "txtReqCodeGen"
        Me.txtReqCodeGen.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtReqCodeGen.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtReqCodeGen.Size = New System.Drawing.Size(445, 107)
        Me.txtReqCodeGen.TabIndex = 8
        '
        'txtLibKeyIn
        '
        Me.txtLibKeyIn.AcceptsReturn = True
        Me.txtLibKeyIn.BackColor = System.Drawing.SystemColors.ButtonFace
        Me.txtLibKeyIn.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtLibKeyIn.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtLibKeyIn.ForeColor = System.Drawing.Color.Green
        Me.txtLibKeyIn.Location = New System.Drawing.Point(96, 150)
        Me.txtLibKeyIn.MaxLength = 0
        Me.txtLibKeyIn.Multiline = True
        Me.txtLibKeyIn.Name = "txtLibKeyIn"
        Me.txtLibKeyIn.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtLibKeyIn.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtLibKeyIn.Size = New System.Drawing.Size(445, 240)
        Me.txtLibKeyIn.TabIndex = 10
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.SystemColors.Control
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label13.Location = New System.Drawing.Point(8, 20)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(89, 17)
        Me.Label13.TabIndex = 20
        Me.Label13.Text = "User Name:"
        '
        'Label11
        '
        Me.Label11.BackColor = System.Drawing.SystemColors.Control
        Me.Label11.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label11.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label11.Location = New System.Drawing.Point(8, 40)
        Me.Label11.Name = "Label11"
        Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label11.Size = New System.Drawing.Size(89, 32)
        Me.Label11.TabIndex = 19
        Me.Label11.Text = "Installation Code (Site Code):"
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.SystemColors.Control
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label4.Location = New System.Drawing.Point(8, 150)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(89, 50)
        Me.Label4.TabIndex = 18
        Me.Label4.Text = "License Key (Liberation Key):"
        '
        '_SSTab1_TabPage1
        '
        Me._SSTab1_TabPage1.Controls.Add(Me.lblLockStatus)
        Me._SSTab1_TabPage1.Controls.Add(Me.lblLockStatus2)
        Me._SSTab1_TabPage1.Controls.Add(Me.lblTrialInfo)
        Me._SSTab1_TabPage1.Controls.Add(Me.Frame1)
        Me._SSTab1_TabPage1.Location = New System.Drawing.Point(4, 22)
        Me._SSTab1_TabPage1.Name = "_SSTab1_TabPage1"
        Me._SSTab1_TabPage1.Size = New System.Drawing.Size(641, 615)
        Me._SSTab1_TabPage1.TabIndex = 1
        Me._SSTab1_TabPage1.Text = "Sample App"
        '
        'lblLockStatus
        '
        Me.lblLockStatus.BackColor = System.Drawing.SystemColors.Control
        Me.lblLockStatus.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLockStatus.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLockStatus.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLockStatus.Location = New System.Drawing.Point(6, 34)
        Me.lblLockStatus.Name = "lblLockStatus"
        Me.lblLockStatus.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLockStatus.Size = New System.Drawing.Size(214, 25)
        Me.lblLockStatus.TabIndex = 33
        Me.lblLockStatus.Text = "Application Functionalities Are Currently: "
        '
        'lblLockStatus2
        '
        Me.lblLockStatus2.BackColor = System.Drawing.SystemColors.Control
        Me.lblLockStatus2.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblLockStatus2.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLockStatus2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblLockStatus2.Location = New System.Drawing.Point(222, 34)
        Me.lblLockStatus2.Name = "lblLockStatus2"
        Me.lblLockStatus2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblLockStatus2.Size = New System.Drawing.Size(301, 17)
        Me.lblLockStatus2.TabIndex = 34
        Me.lblLockStatus2.Text = "Disabled"
        '
        'lblTrialInfo
        '
        Me.lblTrialInfo.BackColor = System.Drawing.SystemColors.Control
        Me.lblTrialInfo.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTrialInfo.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTrialInfo.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTrialInfo.Location = New System.Drawing.Point(6, 58)
        Me.lblTrialInfo.Name = "lblTrialInfo"
        Me.lblTrialInfo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTrialInfo.Size = New System.Drawing.Size(318, 25)
        Me.lblTrialInfo.TabIndex = 38
        Me.lblTrialInfo.Text = "NOTE: All application functionalities are available in Trial Mode."
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me._optForm_1)
        Me.Frame1.Controls.Add(Me._optForm_0)
        Me.Frame1.Controls.Add(Me.cboSpeed)
        Me.Frame1.Controls.Add(Me.chkPause)
        Me.Frame1.Controls.Add(Me.chkFlash)
        Me.Frame1.Controls.Add(Me.chkScroll)
        Me.Frame1.Controls.Add(Me.lblHost)
        Me.Frame1.Controls.Add(Me.lblSpeed)
        Me.Frame1.Cursor = System.Windows.Forms.Cursors.Default
        Me.Frame1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(10, 96)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(503, 125)
        Me.Frame1.TabIndex = 21
        '
        '_optForm_1
        '
        Me._optForm_1.BackColor = System.Drawing.SystemColors.Control
        Me._optForm_1.Cursor = System.Windows.Forms.Cursors.Default
        Me._optForm_1.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._optForm_1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optForm.SetIndex(Me._optForm_1, CType(1, Short))
        Me._optForm_1.Location = New System.Drawing.Point(202, 83)
        Me._optForm_1.Name = "_optForm_1"
        Me._optForm_1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optForm_1.Size = New System.Drawing.Size(93, 15)
        Me._optForm_1.TabIndex = 27
        Me._optForm_1.TabStop = True
        Me._optForm_1.Text = "Option 2"
        Me._optForm_1.UseVisualStyleBackColor = False
        '
        '_optForm_0
        '
        Me._optForm_0.BackColor = System.Drawing.SystemColors.Control
        Me._optForm_0.Checked = True
        Me._optForm_0.Cursor = System.Windows.Forms.Cursors.Default
        Me._optForm_0.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me._optForm_0.ForeColor = System.Drawing.SystemColors.ControlText
        Me.optForm.SetIndex(Me._optForm_0, CType(0, Short))
        Me._optForm_0.Location = New System.Drawing.Point(202, 66)
        Me._optForm_0.Name = "_optForm_0"
        Me._optForm_0.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me._optForm_0.Size = New System.Drawing.Size(93, 15)
        Me._optForm_0.TabIndex = 26
        Me._optForm_0.TabStop = True
        Me._optForm_0.Text = "Option 1"
        Me._optForm_0.UseVisualStyleBackColor = False
        '
        'cboSpeed
        '
        Me.cboSpeed.BackColor = System.Drawing.SystemColors.Window
        Me.cboSpeed.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSpeed.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSpeed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSpeed.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSpeed.Items.AddRange(New Object() {"Slowest", "Slow", "Normal", "Fast", "Fastest"})
        Me.cboSpeed.Location = New System.Drawing.Point(356, 4)
        Me.cboSpeed.Name = "cboSpeed"
        Me.cboSpeed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSpeed.Size = New System.Drawing.Size(109, 22)
        Me.cboSpeed.TabIndex = 25
        '
        'chkPause
        '
        Me.chkPause.BackColor = System.Drawing.SystemColors.Control
        Me.chkPause.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPause.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPause.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkPause.Location = New System.Drawing.Point(0, 26)
        Me.chkPause.Name = "chkPause"
        Me.chkPause.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPause.Size = New System.Drawing.Size(171, 21)
        Me.chkPause.TabIndex = 24
        Me.chkPause.Text = "Checkbox for Level 3 only"
        Me.chkPause.UseVisualStyleBackColor = False
        '
        'chkFlash
        '
        Me.chkFlash.BackColor = System.Drawing.SystemColors.Control
        Me.chkFlash.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkFlash.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkFlash.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkFlash.Location = New System.Drawing.Point(0, 46)
        Me.chkFlash.Name = "chkFlash"
        Me.chkFlash.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkFlash.Size = New System.Drawing.Size(174, 33)
        Me.chkFlash.TabIndex = 23
        Me.chkFlash.Text = "Checkbox for ALL Levels"
        Me.chkFlash.UseVisualStyleBackColor = False
        '
        'chkScroll
        '
        Me.chkScroll.BackColor = System.Drawing.SystemColors.Control
        Me.chkScroll.Checked = True
        Me.chkScroll.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkScroll.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkScroll.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkScroll.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkScroll.Location = New System.Drawing.Point(0, 6)
        Me.chkScroll.Name = "chkScroll"
        Me.chkScroll.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkScroll.Size = New System.Drawing.Size(178, 15)
        Me.chkScroll.TabIndex = 22
        Me.chkScroll.Text = "Checkbox for ALL Levels"
        Me.chkScroll.UseVisualStyleBackColor = False
        '
        'lblHost
        '
        Me.lblHost.BackColor = System.Drawing.SystemColors.Control
        Me.lblHost.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblHost.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblHost.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblHost.Location = New System.Drawing.Point(204, 46)
        Me.lblHost.Name = "lblHost"
        Me.lblHost.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblHost.Size = New System.Drawing.Size(184, 17)
        Me.lblHost.TabIndex = 29
        Me.lblHost.Text = "Option Buttons for ALL Levels:"
        '
        'lblSpeed
        '
        Me.lblSpeed.BackColor = System.Drawing.SystemColors.Control
        Me.lblSpeed.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSpeed.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSpeed.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblSpeed.Location = New System.Drawing.Point(202, 6)
        Me.lblSpeed.Name = "lblSpeed"
        Me.lblSpeed.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSpeed.Size = New System.Drawing.Size(148, 17)
        Me.lblSpeed.TabIndex = 28
        Me.lblSpeed.Text = "Activated with Level 4 Only"
        '
        'frmMain
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(648, 653)
        Me.Controls.Add(Me.SSTab1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.Name = "frmMain"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "ALTestApp - ActiveLock v3.6 Test Application for VB2008"
        Me.SSTab1.ResumeLayout(False)
        Me._SSTab1_TabPage0.ResumeLayout(False)
        Me.fraRegStatus.ResumeLayout(False)
        Me.fraRegStatus.PerformLayout()
        CType(Me.Picture2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.fraReg.ResumeLayout(False)
        Me.fraReg.PerformLayout()
        Me._SSTab1_TabPage1.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        CType(Me.optForm, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
#End Region
#Region "Upgrade Support "
    Private Shared m_vb6FormDefInstance As frmMain
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmMain
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmMain
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal Value As frmMain)
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region
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
    ' This test app is used to exercise all functionalities of ActiveLock.
    '
    ' If you're running this from inside VB and would like to bypass dll-checksumming,
    ' Add the following compilation flag to your Project Properties (Make tab)
    '   AL_DEBUG = 1
    '
    ' @author th2tran
    ' @version 2.0.0
    ' @date 20030715

    '  ///////////////////////////////////////////////////////////////////////
    '  /                        MODULE TO DO LIST                            /
    '  ///////////////////////////////////////////////////////////////////////
    '
    '   [ ] Re: GetMACAndUserFromRequestCode(), try to move the decoding of the
    '       request code inside ActiveLock.  We need to abstract this, if possible,
    '       such that the client app doesn't need to understand how the request
    '       code was encoded.
    '
    Private MyActiveLock As ActiveLock3_6NET._IActiveLock
    Private WithEvents ActiveLockEventSink As ActiveLock3_6NET.ActiveLockEventNotifier

    ' Trial mode variables
    Dim noTrialThisTime As Boolean 'ialkan - needed for registration while form was loaded via trial
    Dim expireTrialLicense As Boolean
    Dim strKeyStorePath As String
    Dim strAutoRegisterKeyPath As String

    'Application name used
    Const LICENSE_ROOT As String = "TestApp"

    ' Timer count to check the license
    Dim timerCount As Long

    ' The following declarations are used by the IsDLLAvailable function
    ' provided by the Activelock user Pinheiro
    Private Declare Function GetLastError Lib "kernel32" () As Integer
    Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Integer, ByRef lpSource As Integer, ByVal dwMessageId As Integer, ByVal dwLanguageId As Integer, ByVal lpBuffer As String, ByVal nSize As Integer, ByRef Arguments As Integer) As Integer
    Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Integer
    Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Integer) As Integer
    Private Const FORMAT_MESSAGE_FROM_SYSTEM As Short = &H1000S
    Private Const MAX_MESSAGE_LENGTH As Short = 512

    'Windows and System directory API
    Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer

    'Splash Screen API
    Private Declare Function SetWindowWord Lib "user32" (ByVal hWnd As Integer, ByVal nIndex As Integer, ByVal wNewWord As Integer) As Integer
    Const GWW_HWNDPARENT As Short = (-8)
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Integer)

    Private Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As IntPtr, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean


    Public Function LooseSpace(ByRef invoer As String) As String
        'This routine terminates a string if it detects char 0.

        Dim P As Integer

        P = InStr(invoer, Chr(0))
        If P <> 0 Then
            LooseSpace = VB.Left(invoer, P - 1)
            Exit Function
        End If
        LooseSpace = invoer

    End Function
    Private Function WindowsDirectory() As String
        'This function gets the windows directory name
        Dim WinPath As String
        Dim Temp As Object
        WinPath = New String(Chr(0), 145)
        Temp = GetWindowsDirectory(WinPath, 145)
        WindowsDirectory = VB.Left(WinPath, InStr(WinPath, Chr(0)) - 1)
    End Function
    Private Function IsDLLAvailable(ByVal DllFilename As String) As Boolean
        ' Code provided by Activelock user Pinheiro
        Dim hModule As Integer
        hModule = LoadLibrary(DllFilename) 'attempt to load DLL
        If hModule > 32 Then
            FreeLibrary(hModule) 'decrement the DLL usage counter
            IsDLLAvailable = True 'Return true
        Else
            IsDLLAvailable = False 'Return False
        End If
    End Function

    Private Sub cmdCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopy.Click
        Dim aDataObject As New DataObject
        aDataObject.SetData(DataFormats.Text, txtReqCodeGen.Text)
        Clipboard.SetDataObject(aDataObject)
    End Sub

    Private Sub cmdKillLicense_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdKillLicense.Click
        ' Kill the License File
        ' Let Activelock handle this
        MyActiveLock.KillLicense(MyActiveLock.SoftwareName & MyActiveLock.SoftwareVersion, strKeyStorePath)

        MsgBox("Your license has been killed." & vbCrLf & _
            "You need to get a new license for this application if you want to use it.", vbInformation)
        txtUsedDays.Text = ""
        txtExpiration.Text = ""
        txtRegisteredLevel.Text = ""
        txtNetworkedLicense.Text = ""
        txtMaxCount.Text = ""

        frmMain_Load(Me, New System.EventArgs)
        cmdResetTrial.Visible = True
    End Sub
    Private Sub cmdKillTrial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdKillTrial.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        MyActiveLock.KillTrial()
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Free Trial has been Killed." & vbCrLf & "There will be no more Free Trial next time you start this application." & vbCrLf & vbCrLf & "You must register this application for further use.", MsgBoxStyle.Information)
        txtRegStatus.Text = "Free Trial has been Killed"
        txtUsedDays.Text = ""
        txtExpiration.Text = ""
        txtRegisteredLevel.Text = ""
        txtLicenseType.Text = "None"
        txtNetworkedLicense.Text = ""
        txtMaxCount.Text = ""

    End Sub
    Private Sub cmdPaste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPaste.Click
        If Clipboard.GetDataObject.GetDataPresent(DataFormats.Text) Then
            If Clipboard.GetDataObject.GetData(DataFormats.Text) = txtReqCodeGen.Text Then
                MsgBox("You cannot paste the Installation Code into the Liberation Key field.", MsgBoxStyle.Exclamation)
                Exit Sub
            End If
            txtLibKeyIn.Text = Clipboard.GetDataObject.GetData(DataFormats.Text)
        End If
    End Sub
    Private Sub cmdResetTrial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdResetTrial.Click
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        MyActiveLock.ResetTrial()
        MyActiveLock.ResetTrial() ' DO NOT REMOVE, NEED TO CALL TWICE
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        MsgBox("Free Trial has been Reset." & vbCrLf & "Please restart the application for a new Free Trial." & vbCrLf & vbCrLf & "Note: This feature is provided for the developers only to test their products;" & vbCrLf & "DO NOT provide this feature in your application.", MsgBoxStyle.Information)
        txtRegStatus.Text = "Free Trial has been Reset"
        txtUsedDays.Text = ""
        txtExpiration.Text = ""
        txtRegisteredLevel.Text = ""
        txtLicenseType.Text = "None"
        txtNetworkedLicense.Text = ""
        txtMaxCount.Text = ""

    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ActiveLock Initialization
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        Dim autoRegisterKey As String = Nothing
        Dim boolAutoRegisterKeyPath As Boolean
        Dim Msg As String
        On Error GoTo NotRegistered

        ' Form's caption
        Me.Text = "ALTestApp3NET - ActiveLock Test Application for VB2008 - v3.6" '& Application.ProductVersion

        ' Check the existence of necessary files to run this application
        ' This is not necessary if you're not using these controls in your app.
        Call CheckForResources("#Alcrypto3NET.dll", "#ActiveLock3_6Net.dll", "comctl32.ocx", "tabctl32.ocx")

        On Error GoTo NotRegistered

        ' Set the path to the license file (LIC) and ALL file (if it exists)
        Dim Ret As Long
        Dim AppfilePath As String
        AppfilePath = Space$(260)
        Ret = SHGetSpecialFolderPath(0, AppfilePath, 46, False) ' 46 is for ...\All Users\Documents folder.
        If AppfilePath.Trim <> Chr(0) Then
            AppfilePath = VB.Left(AppfilePath, InStr(AppfilePath, Chr(0)) - 1)
        End If

        'The second line is used when unmanaged Activelock3NET.dll is used
        Dim MyAL As New ActiveLock3_6NET.Globals
        'Dim MyAL As New ActiveLock3.Globals

        ' Set a new instance of the Activelock object
        MyActiveLock = MyAL.NewInstance()

        With MyActiveLock

            ' Set the software/application name
            .SoftwareName = LICENSE_ROOT
            txtName.Text = .SoftwareName

            ' Set the software/application version number
            ' Note: Do not use (App.Major & "." & App.Minor & "." & App.Revision)
            ' since the license will fail with version incremented via exe rebuilds
            ' THE FOLLOWING IS A SAMPLE USAGE
            ' WARNING *** WARNING *** DO NOT USE App.Major & "." & App.Minor & "." & App.Revision
            '.SoftwareVersion = "1.3.2"   
            .SoftwareVersion = "3" ' WARNING *** WARNING *** DO NOT USE App.Major & "." & App.Minor & "." & App.Revision
            txtVersion.Text = .SoftwareVersion

            ' Set the software/application password
            ' This should be set to protect yourself against ResetTrial abuse
            ' The password is also used by the short keys
            ' Regular licensing does not use this password, but you should still use a password
            ' WARNING: You can not ignore this property. You *must* set a password.
            .SoftwarePassword = Chr(99) & Chr(111) & Chr(111) & Chr(108)

            ' Set whether the software/application will use a short key or RSA method
            ' alsRSA covers both ALCrypto and RSA native classes approach.
            ' RSA classes in .NET allows you to pick from several cipher strengths
            ' however ALCrypto uses 1024 bit strength key only.
            ' alsShortKeyMD5 is for short key protection only
            ' WARNING: Short key licenses use the lockFingerprint by default
            '.LicenseKeyType = ActiveLock3_6NET.IActiveLock.ALLicenseKeyTypes.alsShortKeyMD5
            .LicenseKeyType = ActiveLock3_6NET.IActiveLock.ALLicenseKeyTypes.alsRSA

            ' Set the Trial Feature properties
            ' If you don't want to use the trial feature in your app, set the TrialType
            ' property to trialNone.

            ' Set the trial type property
            ' this is either trialDays, or trialRuns or trialNone.
            .TrialType = ActiveLock3_6NET.IActiveLock.ALTrialTypes.trialDays

            ' Set the Trial Length property.
            ' This number represents the number of days or the number of runs (whichever is applicable).
            .TrialLength = 15
            If .TrialType <> ActiveLock3_6NET.IActiveLock.ALTrialTypes.trialNone And .TrialLength = 0 Then
                ' In such cases Activelock automatically generates errors -11001100 or -11001101
                ' to indicate that you're using the trial feature but, trial length was not specified
            End If

            ' Comment the following statement to use a certain trial data hiding technique
            ' Use OR to combine one or more trial hiding techniques
            ' or don't use this property to use ALL techniques
            ' WARNING: trialRegistryPerUser is "Per User"; this means each user trial feature 
            ' is controlled that user's own registry hive.
            ' This means initiating a trial with one user does not initiate a trial for another user.
            ' trialHiddenFolder and trialSteganography are for "All Users"
            .TrialHideType = ActiveLock3_6NET.IActiveLock.ALTrialHideTypes.trialHiddenFolder Or ActiveLock3_6NET.IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or ActiveLock3_6NET.IActiveLock.ALTrialHideTypes.trialSteganography

            ' Set the Software code
            ' This is the same thing as VCode
            ' Run Alugen first and create a VCode and GCode 
            ' for the software name and version number you used above
            ' Then copy and use the VCode as the PUB_KEY here.
            ' It's up to you to encrypt it; just makes it more secure
            ' Enc encodes, Dec decodes the public key (VCode)
            ' Change Enc() and Dec(0 the way you want.
            .SoftwareCode = Dec(PUB_KEY)

            ' uncomment the following when unmanaged Activelock3NET.dll is used
            '.LockType = ActiveLock3.ALLockTypes.lockNone

            ' Set the Hardware keys
            ' In order to pick the keys that you want to lock to in Alugen, use lockNone only
            ' Example: lockWindows Or lockComp
            ' You can combine any lockType(s) using OR as above
            ' WARNING: Short key licenses use the lockFingerprint by default
            ' WARNING: This has no effect for short key licenses
            .LockType = ActiveLock3_6NET.IActiveLock.ALLockTypes.lockNone
            '.LockType = ActiveLock3_6NET.IActiveLock.ALLockTypes.lockIP Or _
            'ActiveLock3_6NET.IActiveLock.ALLockTypes.lockComp()

            ' If you want to lock to any keys explicitly, combine them using OR
            ' But you won't be able to uncheck/check any of them while in Alugen (too late at that point).
            '.LockType = _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockBIOS Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockComp Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockHD Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockHDFirmware Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockIP Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockMotherboard Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockWindows Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockMAC Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockExternalIP Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockFingerprint Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockSID Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockCPUID Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockBaseboardID Or _
            ' ActiveLock3_6NET.IActiveLock.ALLockTypes.lockVideoID

            ' Set the .ALL file path if you're using an ALL file.
            ' .ALL is an auto registration file.
            ' You generate .ALL files via Alugen and then send to the users
            ' They put the .ALL file in the directory you specify below
            ' .ALL simply contains the license key
            ' WARNING: .ALL files are deleted after they are used.
            ' It is recommended that you use both SoftwareName and Version in the .ALL filename
            ' since multiple .ALL files might exist in the same directory
            ' If you don't want to use the software name and version number explicitly, use an .ALL
            ' filename that is specific to this application
            If Directory.Exists(AppfilePath & "\" & .SoftwareName & .SoftwareVersion) = False Then
                MkDir(AppfilePath & "\" & .SoftwareName & .SoftwareVersion)
            End If
            strAutoRegisterKeyPath = AppfilePath & "\" & .SoftwareName & .SoftwareVersion & "\" & .SoftwareName & .SoftwareVersion & ".all"
            ' AppPath could be an option for XP, but not so for Vista
            'If Directory.Exists(AppPath & "\" & .SoftwareName & .SoftwareVersion ) = False Then
            '    MkDir(AppPath & "\" & .SoftwareName & .SoftwareVersion)
            'End If
            'strAutoRegisterKeyPath = AppPath() & "\" & .SoftwareName & .SoftwareVersion & "\" & .SoftwareName & .SoftwareVersion & ".all"
            .AutoRegisterKeyPath = strAutoRegisterKeyPath
            If File.Exists(strAutoRegisterKeyPath) Then boolAutoRegisterKeyPath = True

            ' Set if auto registration will be used.
            ' Auto registration uses the ALL file for license registration.
            .AutoRegister = ActiveLock3_6NET.IActiveLock.ALAutoRegisterTypes.alsEnableAutoRegistration

            ' Set the Time Server check for Clock Tampering
            ' This is optional but highly recommended.
            ' Although Activelock makes every effort to check if the system clock was tampered,
            ' checking a time server is the guaranteed way of knowing the correct UTC time/day.
            ' This feature might add some delay to your apps start-up time.
            .CheckTimeServerForClockTampering = ActiveLock3_6NET.IActiveLock.ALTimeServerTypes.alsDontCheckTimeServer       ' use alsCheckTimeServer to enforce time server checks for clock tampering check
            '.CheckTimeServerForClockTampering = ActiveLock3_6NET.IActiveLock.ALTimeServerTypes.alsCheckTimeServer

            ' Set the system files clock tampering check
            ' This feature might add some delay to your apps start-up time.
            .CheckSystemFilesForClockTampering = ActiveLock3_6NET.IActiveLock.ALSystemFilesTypes.alsDontCheckSystemFiles    ' use alsCheckSystemFiles to enforce system files scanning for clock tampering check
            '.CheckSystemFilesForClockTampering = ActiveLock3_6NET.IActiveLock.ALSystemFilesTypes.alsCheckSystemFiles

            ' Set the license file format; this could be encrypted or plain
            ' Even in a plain file format, certain keys and dates are still encrypted.
            .LicenseFileType = ActiveLock3_6NET.IActiveLock.ALLicenseFileTypes.alsLicenseFilePlain

            ' Verify Activelock DLL's authenticity by checking its CRC
            ' This checkes the CRC of the Activelock DLL and compares it with the embedded value
            ' To change the embedded value; find the "VerifyActiveLockNETdll" check in this project,
            ' and change the VerifyActiveLockNETdll() function to make it the same as the current CRC.
            txtChecksum.Text = modMain.VerifyActiveLockNETdll()

            ' Initialize the keystore. We use a File keystore in this case.
            ' The other type alsRegistry is NOT supported.
            .KeyStoreType = ActiveLock3_6NET.IActiveLock.LicStoreType.alsFile
            ' uncomment the following when unmanaged Activelock3NET.dll is used
            'MyActiveLock.KeyStoreType = ActiveLock3.LicStoreType.alsFile

            ' The code below will put the LIC file inside the "...\All Users\Documents" folder
            ' You can hard code this path and put the LIC file anywhere you want
            ' But be careful with limited user accounts in Vista.
            ' It's recommended that you put this file in shared and accessible folders in Vista
            ' It is recommended that you use both SoftwareName and Version in the LIC filename
            ' since multiple LIC files might exist in the same directory
            ' If you don't want to use the software name and version number explicitly, use an LIC
            ' filename that is specific to this application
            If Directory.Exists(AppfilePath & "\" & .SoftwareName & .SoftwareVersion) = False Then
                MkDir(AppfilePath & "\" & .SoftwareName & .SoftwareVersion)
            End If
            strKeyStorePath = AppfilePath & "\" & .SoftwareName & .SoftwareVersion & "\" & .SoftwareName & .SoftwareVersion & ".lic"
            ' AppPath could be an option for XP, but not so for Vista
            'If Directory.Exists(AppPath & "\" & .SoftwareName & .SoftwareVersion ) = False Then
            '    MkDir(AppPath & "\" & .SoftwareName & .SoftwareVersion)
            'End If
            'strKeyStorePath = AppPath() & "\" & .SoftwareName & .SoftwareVersion & "\" & .SoftwareName & .SoftwareVersion & ".lic"
            .KeyStorePath = strKeyStorePath

            ' Obtain the EventNotifier so that we can receive notifications from AL.
            ' Do Not Change This  - unless you know what this is for.
            ActiveLockEventSink = .EventNotifier

            ' Initialize Activelock 
            ' This is for handling the ALCrypto DLL used by Activelock. 
            ' Init() method below checkes the ALCrypto CRC, and registeres the 
            ' application if an ALL file was used.
            ' Important: If you're not going to put Alcrypto3NET.dll under
            ' the system32 directory, you should pass the path of the exe
            ' to the Init() method otherwise this call will fail
            ' Putting Alcrypto3NET.dll under the system32 is a problem with the ASP.NET apps
            ' Since Activelock3NET can be used by ASP.NET apps, setting the first arguments below,
            ' will help you put the ALcrypto DLL into a location you want; mostly the app folder.
            ' This is particularly useful for hosted ASP.NET apps where 
            ' you don't have the server control (no system32 access)
            ' Use the following with ASP.NET applications
            ' MyActiveLock.Init(Application.StartupPath & "\bin")
            ' Use the following with VB.NET applications
            .Init(Application.StartupPath, strKeyStorePath)
            If File.Exists(strKeyStorePath) And boolAutoRegisterKeyPath = True And autoRegisterKey <> "" Then
                ' This means, an ALL file existed and was used to create a LIC file
                ' Init() method successfully registered the ALL file
                ' and returned the license key
                ' You can process that key here to see if there is any abuse, etc.
                ' ie. whether the key was used before, etc.
            End If

        End With

        cboSpeed.SelectedItem = 2

        ' Check registration status
        ' Acquire() method does both trial and regular licensing
        ' If it generates an error, that means there NO trial, NO license
        ' If no error and returns a string, there's a trial but No license. Parse the string to display a trial message.
        ' If no error and no string returned, you've got a valid license.

        ' In case the Acquire method generates an error, so no license and no trial:
        ' If InStr(1, Err.Description, "No valid license") > 0 Or InStr(1, Err.Description, "license invalid") > 0 Then '-2147221503 & -2147221502

        MyActiveLock.Acquire(strMsg, strRemainingTrialDays, strRemainingTrialRuns, strTrialLength, strUsedDays, strExpirationDate, strRegisteredUser, strRegisteredLevel, strLicenseClass, strMaxCount, strLicenseFileType, strLicenseType, strUsedLockType)
        ' strMsg is to get the trial status
        ' All other parameters are Optional and you can actually get all of them
        ' using MyActivelock.Property usage, but keep in mind that 
        ' doing so will check the license every time making this a time consuming 
        ' way of reading those properties
        ' The fastest approach is to use the arguments from Acquire() method.
        If strMsg <> "" Then 'There's a trial
            ' Parse the returned string to display it on a form
            Dim A() As String = strMsg.Split(vbCrLf)
            txtRegStatus.Text = A(0).Trim
            txtUsedDays.Text = A(1).Trim

            ' You can also get the RemainingTrialDays or RemainingTrialRuns directly by:
            'txtRemainingTrialDays.Text = MyActiveLock.RemainingTrialDays OR MyActiveLock.RemainingTrialRuns
            ' At this point RemainingTrialDays and RemainingTrialRuns properties are directly accessible
            ' Even if you don't want to show a trial form at this point, you still know the 
            ' trial status with one of these two properties (whichever is applicable).

            FunctionalitiesEnabled = True

            ' ALTERNATIVE 1
            ' Show a splash form for 3 seconds
            ' The form displays the trial days/run total allowed and remaining
            ' Splash form to display the trial period/run information.
            Dim frmsplash As New frmSplash
            frmsplash.lblInfo.Text = vbCrLf & strMsg
            frmsplash.Visible = True
            frmsplash.Refresh()
            Sleep(3000) 'wait about 3 seconds
            frmsplash.Close()

            ' ALTERNATIVE 2
            ' Show a splash form that shows two choices, register or try
            ' User must chosse one option for the form to close
            'MyActiveLock.RemainingTrialDays
            remainingDays = CInt(strRemainingTrialDays)
            'MyActiveLock.TrialLength
            totalDays = CInt(strTrialLength)
            Dim frmsplash1 As New frmSplash1
            frmsplash1.Visible = False
            frmsplash1.ShowDialog()

            cmdKillTrial.Visible = True
            cmdResetTrial.Visible = True
            txtLicenseType.Text = "Free Trial"
            Me.Refresh()
            Exit Sub
        Else
            cmdKillTrial.Visible = False
            cmdResetTrial.Visible = False
        End If

        ' You can retrieve the LockTypes set inside Alugen
        ' by accessing the UsedLockType property
        'For example, if only lockHDfirmware was used, this will return 256
        'MsgBox(MyActiveLock.UsedLockType)

        ' So far so good... 
        ' If you are here already, that means you have a valid license.
        ' Set the textboxes in your app accordingly.
        txtRegStatus.Text = "Registered"
        'MyActiveLock.UsedDays
        txtUsedDays.Text = strUsedDays
        'MyActiveLock.ExpirationDate
        txtExpiration.Text = strExpirationDate
        If txtExpiration.Text = "" Then txtExpiration.Text = "Permanent" 'App has a permanent license

        Dim arrProdVer() As String
        'MyActiveLock.RegisteredUser
        arrProdVer = Split(strRegisteredUser, "&&&") ' Extract the software name and version
        txtUser.Text = arrProdVer(0)

        'MyActiveLock.RegisteredLevel
        txtRegisteredLevel.Text = strRegisteredLevel

        ' Set Networked Licenses if applicable
        'MyActiveLock.LicenseClass
        If strLicenseClass = "MultiUser" Then
            txtNetworkedLicense.Text = "Networked"
        Else
            txtNetworkedLicense.Text = "Single User"
            txtMaxCount.Visible = False
            lblConcurrentUsers.Visible = False
        End If
        ' This is for number of concurrent users count in a netwrok license
        'MyActiveLock.MaxCount
        txtMaxCount.Text = strMaxCount

        ' Read the license file into a string to determine the license type
        ' You can read the LicenseType from the LIC file directly
        ' However, the LIC file should be in Plain format for this to work.
        ' MyActiveLock.LicenseFileType
        If strLicenseFileType = ActiveLock3_6NET.IActiveLock.ALLicenseFileTypes.alsLicenseFilePlain Then
            Dim strBuff As String
            Dim fNum As Short
            fNum = FreeFile()
            FileOpen(fNum, strKeyStorePath, OpenMode.Input)
            strBuff = InputString(1, LOF(1))
            FileClose(fNum)
            If Instring(strBuff, "LicenseType=3") Then
                txtLicenseType.Text = "Time Limited"
            ElseIf Instring(strBuff, "LicenseType=1") Then
                txtLicenseType.Text = "Periodic"
            ElseIf Instring(strBuff, "LicenseType=2") Then
                txtLicenseType.Text = "Permanent"
            End If
        Else
            If strLicenseType = "3" Then
                txtLicenseType.Text = "Time Limited"
            ElseIf strLicenseType = "1" Then
                txtLicenseType.Text = "Periodic"
            ElseIf strLicenseType = "2" Then
                txtLicenseType.Text = "Permanent"
            End If
        End If

        FunctionalitiesEnabled = True
        txtUser.BackColor = Color.FromKnownColor(KnownColor.ButtonFace)

        ' If your code has reached this point successfully, then you're good.
        ' If not, revisit your code, check and recheck,
        ' watch the video tutorial and if still in doubt,
        ' post a question in the forums
        ' http://www.activelocksoftware.com

        ' The following used with a timer control to
        ' check the license periodically
        ' Check the Timer1_Tick event for further details
        timerCount = 5000
        Timer1.Interval = 1
        Timer1.Enabled = True

        Exit Sub

NotRegistered:
        ' There's no valid trial or license - let the user know.
        ' and cripple your application if necessary
        ' or kill it if that's what you want.
        FunctionalitiesEnabled = False
        If Instring(Err.Description, "no valid license") = False And noTrialThisTime = False Then
            MsgBox(Err.Number & ": " & Err.Description)
        End If
        txtRegStatus.Text = Err.Description
        txtLicenseType.Text = "None"
        txtUser.Text = System.Environment.UserName
        If strMsg <> "" Then
            MsgBox(strMsg, MsgBoxStyle.Information)
        End If
        Exit Sub

DLLnotRegistered:
        End

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

        On Error GoTo checkForResourcesError
        Dim foundIt As Boolean
        Dim y As Object
        Dim i, j As Short
        Dim systemDir, s, pathName As String

        'WhereIsDLL("") 'initialize

        systemDir = WindowsSystemDirectory() 'Get the Windows system directory
        For Each y In MyArray
            foundIt = False
            s = CStr(y)

            If VB.Left(s, 1) = "#" Then
                pathName = VB6.GetPath
                s = Mid(s, 2)
            ElseIf Instring(s, "\") Then
                j = InStrRev(s, "\")
                pathName = VB.Left(s, j - 1)
                s = Mid(s, j + 1)
            Else
                pathName = systemDir
            End If

            If Instring(s, ".") Then
                If File.Exists(pathName & "\" & s) Then foundIt = True
            ElseIf File.Exists(pathName & "\" & s & ".DLL") Then
                foundIt = True
            ElseIf File.Exists(pathName & "\" & s & ".OCX") Then
                foundIt = True
                s = s & ".OCX"
            End If

            If Not foundIt Then
                MsgBox(s & " could not be found in " & pathName & "." & vbCrLf & System.Reflection.Assembly.GetExecutingAssembly.GetName.Name & " cannot run without this library file!" & vbCrLf & vbCrLf & "Exiting!", MsgBoxStyle.Critical, "Missing Resource")
                End
            End If
        Next y

        CheckForResources = True
        Exit Function

checkForResourcesError:
        MsgBox("CheckForResources error", MsgBoxStyle.Critical, "Error")
        End 'an error kills the program
    End Function
    Private Function WindowsSystemDirectory() As String

        Dim cnt As Integer
        Dim s As String
        Dim dl As Integer

        cnt = 254
        s = New String(Chr(0), 254)
        dl = GetSystemDirectory(s, cnt)
        WindowsSystemDirectory = LooseSpace(VB.Left(s, cnt))

    End Function

    'Function WhereIsDLL(ByVal T As String) As String
    '    'Places where programs look for DLLs
    '    '   1 directory containing the EXE
    '    '   2 current directory
    '    '   3 32 bit system directory   possibly \Windows\system32
    '    '   4 16 bit system directory   possibly \Windows\system
    '    '   5 windows directory         possibly \Windows
    '    '   6 path

    '    'The current directory may be changed in the course of the program
    '    'but current directory -- when the program starts -- is what matters
    '    'so a call should be made to this function early on to "lock" the paths.

    '    'Add a call at the beginning of checkForResources
    '    Dim A() As Object = Nothing
    '    Dim s, d As String
    '    Dim EnvString As String
    '    Dim Indx As Short ' Declare variables.
    '    Dim i As Short

    '    On Error Resume Next
    '    'i = UBound(A)
    '    'If i = 0 Then
    '    s = VB6.GetPath & ";" & CurDir() & ";"

    '    d = WindowsSystemDirectory()
    '    s = s & d & ";"

    '    If VB.Right(d, 2) = "32" Then 'I'm guessing at the name of the 16 bit windows directory (assuming it exists)
    '        i = Len(d)
    '        s = s & VB.Left(d, i - 2) & ";"
    '    End If

    '    s = s & WindowsDirectory() & ";"
    '    Indx = 1 ' Initialize index to 1.
    '    Do
    '        EnvString = Environ(Indx) ' Get environment variable.
    '        If StrComp(VB.Left(EnvString, 5), "PATH=", CompareMethod.Text) = 0 Then ' Check PATH entry.
    '            s = s & Mid(EnvString, 6)
    '            Exit Do
    '        End If
    '        Indx = Indx + 1
    '    Loop Until EnvString = ""
    '    A = s.Split(";")
    '    'End If

    '    T = Trim(T)
    '    If T = "" Then Return Nothing
    '    If Not Instring(VB.Right(T, 4), ".") Then T = T & ".DLL" 'default extension
    '    For i = 0 To UBound(A)
    '        If File.Exists(A(i) & "\" & T) Then
    '            WhereIsDLL = A(i)
    '            Exit Function
    '        End If
    '    Next i
    '    Return Nothing
    'End Function
    Function Instring(ByVal x As String, ByVal ParamArray MyArray() As Object) As Boolean
        'Do ANY of a group of sub-strings appear in within the first string?
        'Case doesn't count and we don't care WHERE or WHICH
        Dim y As Object 'member of array that holds all arguments except the first
        For Each y In MyArray
            If InStr(1, x, y, 1) > 0 Then 'the "ones" make the comparison case-insensitive
                Instring = True
                Exit Function
            End If
        Next y
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Key Validation Functionalities
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' ActiveLock raises this event typically when it needs a value to be encrypted.
    ' We can use any kind of encryption we'd like here, as long as it's deterministic.
    ' i.e. there's a one-to-one correspondence between unencrypted value and encrypted value.
    ' NOTE: BlowFish is NOT an example of deterministic encryption so you can't use it here.
    Private Sub ActiveLockEventSink_ValidateValue(ByRef Value As String) Handles ActiveLockEventSink.ValidateValue
        Value = Encrypt(Value)
    End Sub

    Private Function Encrypt(ByRef strdata As String) As String
        Dim i, n As Integer
        Dim sResult As String = Nothing
        n = Len(strdata)
        For i = 1 To n
            sResult = sResult & Asc(Mid(strdata, i, 1)) * 7
        Next i
        Encrypt = sResult
    End Function

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Key Request and Registration Functionalities
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub cmdReqGen_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdReqGen.Click
        Dim instCode As String
        ' Generate Request code to Lock
        If MyActiveLock Is Nothing Then
            noTrialThisTime = True
            frmMain_Load(Me, New System.EventArgs)
        End If
        If txtRegStatus.Text <> "Registered" Then txtRegStatus.Text = ""
        If Not IsNumeric(txtUsedDays.Text) Then txtUsedDays.Text = ""
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        instCode = MyActiveLock.InstallationCode(txtUser.Text)
        If Len(instCode) = 8 Then
            instCode = "You must send all of the following for authorization:" & vbCrLf & _
                "Serial Number: " & instCode & vbCrLf & _
                "Application Name: " & txtName.Text & " - Version " & txtVersion.Text & vbCrLf & _
                "User Name: " & txtUser.Text
        End If
        txtReqCodeGen.Text = instCode
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
    End Sub

    Private Sub cmdRegister_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRegister.Click
        Dim ok As Boolean, LibKey As String
        On Error GoTo errHandler
        ' Register this key
        LibKey = txtLibKeyIn.Text
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor
        If Mid(LibKey, 5, 1) = "-" And Mid(LibKey, 10, 1) = "-" And Mid(LibKey, 15, 1) = "-" And Mid(LibKey, 20, 1) = "-" Then
            MyActiveLock.Register(LibKey, txtUser.Text) 'YOU MUST SPECIFY THE USER NAME WITH SHORT KEYS !!!
        Else    ' ALCRYPTO RSA
            MyActiveLock.Register(LibKey)
        End If
        MsgBox(modMain.Dec("386.457.46D.483.4F1.4FC.4E6.42B.4FC.483.4C5.4BA.160.4F1.507.441.441.457.4F1.4F1.462.507.4A4.16B"), MsgBoxStyle.Information) ' "Registration successful!"
        MyActiveLock = Nothing
        strMsg = String.Empty
        remainingDays = 0
        totalDays = 0
        txtReqCodeGen.Text = String.Empty
        txtLibKeyIn.Text = String.Empty
        frmMain_Load(Me, New System.EventArgs)
        Me.Visible = True
        System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
        Exit Sub

errHandler:
        MsgBox(Err.Number & ": " & Err.Description)
    End Sub
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' The rest of the application's functionalities
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Private Sub cboSpeed_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSpeed.SelectedIndexChanged

    End Sub

    Private WriteOnly Property FunctionalitiesEnabled() As Boolean
        Set(ByVal Value As Boolean)
            chkScroll.Enabled = Value
            chkFlash.Enabled = Value
            lblHost.Enabled = Value
            optForm(0).Enabled = Value
            optForm(1).Enabled = Value
            chkPause.Enabled = Value
            lblSpeed.Enabled = Value
            cboSpeed.Enabled = Value
            If Value Then
                If txtRegisteredLevel.Text <> "" Then
                    lblLockStatus2.Text = "Enabled with " & txtRegisteredLevel.Text
                    chkPause.Enabled = (InStr(1, txtRegisteredLevel.Text, "Level 3") > 0)
                    lblSpeed.Enabled = (InStr(1, txtRegisteredLevel.Text, "Level 4") > 0)
                    cboSpeed.Enabled = (InStr(1, txtRegisteredLevel.Text, "Level 4") > 0)
                Else
                    lblLockStatus2.Text = "Enabled with " & txtUsedDays.Text
                End If
            Else
                lblLockStatus2.Text = "Disabled (Registration Required)"
                chkPause.Enabled = False
                lblSpeed.Enabled = False
                cboSpeed.Enabled = False
            End If
        End Set
    End Property
    Private Sub frmMain_Closed(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Closed
        frmMain.DefInstance = Nothing
        'DO NOT ADD THE "END" STATEMENT INTO THIS SUB
        'Form reloads upon registration
    End Sub
    Private Sub txtLibKeyIn_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtLibKeyIn.TextChanged
        cmdRegister.Enabled = CBool(Trim(txtLibKeyIn.Text) <> "")
    End Sub
    Private Sub txtName_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtName.TextChanged
        'MyActiveLock.SoftwareName = txtName
    End Sub
    Private Sub txtVersion_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtVersion.TextChanged
        'MyActiveLock.SoftwareVersion = txtVersion
    End Sub
    Public Function AppPath() As String
        Return System.Windows.Forms.Application.StartupPath
    End Function

    Private Sub Timer1_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ' You can use this timer control and this event in your
        ' own application to check the Activelock license periodically
        ' If there's anything wrong, this event will end your application
        ' In addition to the code below, make sure to declare timerCount
        ' and add the following in your Form_Load:
        '    timerCount = 5000
        '    Timer1.Interval = 1
        '    Timer1.Enabled = True

        On Error GoTo errorTrap
        timerCount = timerCount - 1
        If timerCount = 0 Then
            ' Check the license if there was a license at form_load
            ' and it was not a trial license
            If txtRegStatus.Text = Chr(82) & Chr(101) & Chr(103) & Chr(105) & Chr(115) & Chr(116) & Chr(101) & Chr(114) & Chr(101) & Chr(100) And strMsg = "" Then ' "Registered" and no trial message
                MyActiveLock.Acquire(strMsg)
                timerCount = 5000
            End If
        End If
        Exit Sub

errorTrap:
        MsgBox("License error. Quitting.", vbCritical)
        End
    End Sub

End Class