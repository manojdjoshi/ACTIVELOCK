
Public Class frmAlugendb_details
  Inherits System.Windows.Forms.Form

  Private printPreviewDialog1 As New PrintPreviewDialog

#Region " Codice generato da Progettazione Windows Form "

  Public Sub New()
    MyBase.New()

    'Chiamata richiesta da Progettazione Windows Form.
    InitializeComponent()

    'Aggiungere le eventuali istruzioni di inizializzazione dopo la chiamata a InitializeComponent()

  End Sub

  'Form esegue l'override del metodo Dispose per pulire l'elenco dei componenti.
  Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
    If disposing Then
      If Not (components Is Nothing) Then
        components.Dispose()
      End If
    End If
    MyBase.Dispose(disposing)
  End Sub

  'Richiesto da Progettazione Windows Form
  Private components As System.ComponentModel.IContainer

  'NOTA: la procedura che segue è richiesta da Progettazione Windows Form.
  'Può essere modificata in Progettazione Windows Form.  
  'Non modificarla nell'editor del codice.
  Friend WithEvents txtprogname As System.Windows.Forms.TextBox
  Friend WithEvents txtprogver As System.Windows.Forms.TextBox
  Friend WithEvents lblprogname As System.Windows.Forms.Label
  Friend WithEvents lblprogver As System.Windows.Forms.Label
  Friend WithEvents lblexpdate As System.Windows.Forms.Label
  Friend WithEvents lblregdate As System.Windows.Forms.Label
  Friend WithEvents txtexpdate As System.Windows.Forms.TextBox
  Friend WithEvents txtregdate As System.Windows.Forms.TextBox
  Friend WithEvents lblLockType As System.Windows.Forms.Label
  Friend WithEvents lblLicType As System.Windows.Forms.Label
  Friend WithEvents txtLicType As System.Windows.Forms.TextBox
  Friend WithEvents txtLockType As System.Windows.Forms.TextBox
  Friend WithEvents lblLibCode As System.Windows.Forms.Label
  Friend WithEvents lblusername As System.Windows.Forms.Label
  Friend WithEvents lblinstcode As System.Windows.Forms.Label
  Friend WithEvents txtinstcode As System.Windows.Forms.TextBox
  Friend WithEvents txtusername As System.Windows.Forms.TextBox
  Friend WithEvents lblreglevel As System.Windows.Forms.Label
  Friend WithEvents txtreglevel As System.Windows.Forms.TextBox
  Friend WithEvents picALBanner As System.Windows.Forms.PictureBox
  Friend WithEvents cmdEmailLicenseKey As System.Windows.Forms.Button
  Friend WithEvents cmdPrintLicenseKey As System.Windows.Forms.Button
  Friend WithEvents cmdCopy As System.Windows.Forms.Button
  Friend WithEvents txtLicenseKey As System.Windows.Forms.TextBox
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmAlugendb_details))
    Me.txtprogname = New System.Windows.Forms.TextBox
    Me.txtprogver = New System.Windows.Forms.TextBox
    Me.lblprogname = New System.Windows.Forms.Label
    Me.lblprogver = New System.Windows.Forms.Label
    Me.lblexpdate = New System.Windows.Forms.Label
    Me.lblregdate = New System.Windows.Forms.Label
    Me.txtexpdate = New System.Windows.Forms.TextBox
    Me.txtregdate = New System.Windows.Forms.TextBox
    Me.lblinstcode = New System.Windows.Forms.Label
    Me.lblreglevel = New System.Windows.Forms.Label
    Me.txtinstcode = New System.Windows.Forms.TextBox
    Me.txtreglevel = New System.Windows.Forms.TextBox
    Me.lblLockType = New System.Windows.Forms.Label
    Me.lblLicType = New System.Windows.Forms.Label
    Me.txtLockType = New System.Windows.Forms.TextBox
    Me.txtLicType = New System.Windows.Forms.TextBox
    Me.lblLibCode = New System.Windows.Forms.Label
    Me.lblusername = New System.Windows.Forms.Label
    Me.txtLicenseKey = New System.Windows.Forms.TextBox
    Me.txtusername = New System.Windows.Forms.TextBox
    Me.picALBanner = New System.Windows.Forms.PictureBox
    Me.cmdEmailLicenseKey = New System.Windows.Forms.Button
    Me.cmdPrintLicenseKey = New System.Windows.Forms.Button
    Me.cmdCopy = New System.Windows.Forms.Button
    Me.SuspendLayout()
    '
    'txtprogname
    '
    Me.txtprogname.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtprogname.Location = New System.Drawing.Point(80, 9)
    Me.txtprogname.Name = "txtprogname"
    Me.txtprogname.Size = New System.Drawing.Size(296, 20)
    Me.txtprogname.TabIndex = 0
    Me.txtprogname.Text = ""
    '
    'txtprogver
    '
    Me.txtprogver.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtprogver.Location = New System.Drawing.Point(80, 33)
    Me.txtprogver.Name = "txtprogver"
    Me.txtprogver.Size = New System.Drawing.Size(296, 20)
    Me.txtprogver.TabIndex = 1
    Me.txtprogver.Text = ""
    '
    'lblprogname
    '
    Me.lblprogname.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.lblprogname.Location = New System.Drawing.Point(8, 8)
    Me.lblprogname.Name = "lblprogname"
    Me.lblprogname.Size = New System.Drawing.Size(64, 23)
    Me.lblprogname.TabIndex = 5
    Me.lblprogname.Text = "Prog name"
    Me.lblprogname.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblprogver
    '
    Me.lblprogver.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.lblprogver.Location = New System.Drawing.Point(8, 32)
    Me.lblprogver.Name = "lblprogver"
    Me.lblprogver.Size = New System.Drawing.Size(64, 23)
    Me.lblprogver.TabIndex = 6
    Me.lblprogver.Text = "Prog Ver."
    Me.lblprogver.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblexpdate
    '
    Me.lblexpdate.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.lblexpdate.Location = New System.Drawing.Point(8, 80)
    Me.lblexpdate.Name = "lblexpdate"
    Me.lblexpdate.Size = New System.Drawing.Size(64, 23)
    Me.lblexpdate.TabIndex = 10
    Me.lblexpdate.Text = "Exp Date"
    Me.lblexpdate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblregdate
    '
    Me.lblregdate.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.lblregdate.Location = New System.Drawing.Point(8, 56)
    Me.lblregdate.Name = "lblregdate"
    Me.lblregdate.Size = New System.Drawing.Size(64, 23)
    Me.lblregdate.TabIndex = 9
    Me.lblregdate.Text = "Reg Date"
    Me.lblregdate.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'txtexpdate
    '
    Me.txtexpdate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtexpdate.Location = New System.Drawing.Point(80, 80)
    Me.txtexpdate.Name = "txtexpdate"
    Me.txtexpdate.Size = New System.Drawing.Size(296, 20)
    Me.txtexpdate.TabIndex = 8
    Me.txtexpdate.Text = ""
    '
    'txtregdate
    '
    Me.txtregdate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtregdate.Location = New System.Drawing.Point(80, 56)
    Me.txtregdate.Name = "txtregdate"
    Me.txtregdate.Size = New System.Drawing.Size(296, 20)
    Me.txtregdate.TabIndex = 7
    Me.txtregdate.Text = ""
    '
    'lblinstcode
    '
    Me.lblinstcode.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.lblinstcode.Location = New System.Drawing.Point(8, 176)
    Me.lblinstcode.Name = "lblinstcode"
    Me.lblinstcode.Size = New System.Drawing.Size(64, 23)
    Me.lblinstcode.TabIndex = 18
    Me.lblinstcode.Text = "Inst Code"
    Me.lblinstcode.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblreglevel
    '
    Me.lblreglevel.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.lblreglevel.Location = New System.Drawing.Point(8, 152)
    Me.lblreglevel.Name = "lblreglevel"
    Me.lblreglevel.Size = New System.Drawing.Size(64, 23)
    Me.lblreglevel.TabIndex = 17
    Me.lblreglevel.Text = "Reg Level"
    Me.lblreglevel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'txtinstcode
    '
    Me.txtinstcode.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtinstcode.Location = New System.Drawing.Point(80, 176)
    Me.txtinstcode.Name = "txtinstcode"
    Me.txtinstcode.Size = New System.Drawing.Size(296, 20)
    Me.txtinstcode.TabIndex = 16
    Me.txtinstcode.Text = ""
    '
    'txtreglevel
    '
    Me.txtreglevel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtreglevel.Location = New System.Drawing.Point(80, 152)
    Me.txtreglevel.Name = "txtreglevel"
    Me.txtreglevel.Size = New System.Drawing.Size(296, 20)
    Me.txtreglevel.TabIndex = 15
    Me.txtreglevel.Text = ""
    '
    'lblLockType
    '
    Me.lblLockType.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.lblLockType.Location = New System.Drawing.Point(8, 128)
    Me.lblLockType.Name = "lblLockType"
    Me.lblLockType.Size = New System.Drawing.Size(64, 23)
    Me.lblLockType.TabIndex = 14
    Me.lblLockType.Text = "Lock Type"
    Me.lblLockType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblLicType
    '
    Me.lblLicType.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.lblLicType.Location = New System.Drawing.Point(8, 104)
    Me.lblLicType.Name = "lblLicType"
    Me.lblLicType.Size = New System.Drawing.Size(64, 23)
    Me.lblLicType.TabIndex = 13
    Me.lblLicType.Text = "Lic Type"
    Me.lblLicType.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'txtLockType
    '
    Me.txtLockType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtLockType.Location = New System.Drawing.Point(80, 128)
    Me.txtLockType.Name = "txtLockType"
    Me.txtLockType.Size = New System.Drawing.Size(296, 20)
    Me.txtLockType.TabIndex = 12
    Me.txtLockType.Text = ""
    '
    'txtLicType
    '
    Me.txtLicType.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtLicType.Location = New System.Drawing.Point(80, 104)
    Me.txtLicType.Name = "txtLicType"
    Me.txtLicType.Size = New System.Drawing.Size(296, 20)
    Me.txtLicType.TabIndex = 11
    Me.txtLicType.Text = ""
    '
    'lblLibCode
    '
    Me.lblLibCode.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.lblLibCode.Location = New System.Drawing.Point(8, 224)
    Me.lblLibCode.Name = "lblLibCode"
    Me.lblLibCode.Size = New System.Drawing.Size(64, 23)
    Me.lblLibCode.TabIndex = 22
    Me.lblLibCode.Text = "Lib Code"
    Me.lblLibCode.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'lblusername
    '
    Me.lblusername.Font = New System.Drawing.Font("Arial", 8.0!)
    Me.lblusername.Location = New System.Drawing.Point(8, 200)
    Me.lblusername.Name = "lblusername"
    Me.lblusername.Size = New System.Drawing.Size(64, 23)
    Me.lblusername.TabIndex = 21
    Me.lblusername.Text = "User name"
    Me.lblusername.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
    '
    'txtLicenseKey
    '
    Me.txtLicenseKey.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtLicenseKey.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.txtLicenseKey.Location = New System.Drawing.Point(80, 224)
    Me.txtLicenseKey.Multiline = True
    Me.txtLicenseKey.Name = "txtLicenseKey"
    Me.txtLicenseKey.Size = New System.Drawing.Size(464, 122)
    Me.txtLicenseKey.TabIndex = 20
    Me.txtLicenseKey.Text = ""
    '
    'txtusername
    '
    Me.txtusername.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.txtusername.Location = New System.Drawing.Point(80, 200)
    Me.txtusername.Name = "txtusername"
    Me.txtusername.Size = New System.Drawing.Size(296, 20)
    Me.txtusername.TabIndex = 19
    Me.txtusername.Text = ""
    '
    'picALBanner
    '
    Me.picALBanner.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
    Me.picALBanner.BackColor = System.Drawing.SystemColors.ActiveCaptionText
    Me.picALBanner.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
    Me.picALBanner.Cursor = System.Windows.Forms.Cursors.Hand
    Me.picALBanner.Location = New System.Drawing.Point(2, 308)
    Me.picALBanner.Name = "picALBanner"
    Me.picALBanner.Size = New System.Drawing.Size(74, 38)
    Me.picALBanner.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
    Me.picALBanner.TabIndex = 65
    Me.picALBanner.TabStop = False
    '
    'cmdEmailLicenseKey
    '
    Me.cmdEmailLicenseKey.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdEmailLicenseKey.BackColor = System.Drawing.SystemColors.Control
    Me.cmdEmailLicenseKey.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdEmailLicenseKey.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdEmailLicenseKey.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdEmailLicenseKey.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdEmailLicenseKey.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdEmailLicenseKey.Location = New System.Drawing.Point(464, 63)
    Me.cmdEmailLicenseKey.Name = "cmdEmailLicenseKey"
    Me.cmdEmailLicenseKey.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdEmailLicenseKey.Size = New System.Drawing.Size(80, 24)
    Me.cmdEmailLicenseKey.TabIndex = 68
    Me.cmdEmailLicenseKey.Text = "Email  key"
    Me.cmdEmailLicenseKey.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    '
    'cmdPrintLicenseKey
    '
    Me.cmdPrintLicenseKey.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdPrintLicenseKey.BackColor = System.Drawing.SystemColors.Control
    Me.cmdPrintLicenseKey.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdPrintLicenseKey.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdPrintLicenseKey.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdPrintLicenseKey.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdPrintLicenseKey.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdPrintLicenseKey.Location = New System.Drawing.Point(464, 35)
    Me.cmdPrintLicenseKey.Name = "cmdPrintLicenseKey"
    Me.cmdPrintLicenseKey.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdPrintLicenseKey.Size = New System.Drawing.Size(80, 24)
    Me.cmdPrintLicenseKey.TabIndex = 67
    Me.cmdPrintLicenseKey.Text = "Print  key"
    Me.cmdPrintLicenseKey.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    '
    'cmdCopy
    '
    Me.cmdCopy.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.cmdCopy.BackColor = System.Drawing.SystemColors.Control
    Me.cmdCopy.Cursor = System.Windows.Forms.Cursors.Default
    Me.cmdCopy.FlatStyle = System.Windows.Forms.FlatStyle.Popup
    Me.cmdCopy.Font = New System.Drawing.Font("Arial", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
    Me.cmdCopy.ForeColor = System.Drawing.SystemColors.ControlText
    Me.cmdCopy.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
    Me.cmdCopy.Location = New System.Drawing.Point(464, 8)
    Me.cmdCopy.Name = "cmdCopy"
    Me.cmdCopy.RightToLeft = System.Windows.Forms.RightToLeft.No
    Me.cmdCopy.Size = New System.Drawing.Size(80, 24)
    Me.cmdCopy.TabIndex = 66
    Me.cmdCopy.Text = "Copy  key"
    Me.cmdCopy.TextAlign = System.Drawing.ContentAlignment.MiddleRight
    '
    'frmAlugendb_details
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.ClientSize = New System.Drawing.Size(552, 357)
    Me.Controls.Add(Me.cmdEmailLicenseKey)
    Me.Controls.Add(Me.cmdPrintLicenseKey)
    Me.Controls.Add(Me.cmdCopy)
    Me.Controls.Add(Me.picALBanner)
    Me.Controls.Add(Me.lblLibCode)
    Me.Controls.Add(Me.lblusername)
    Me.Controls.Add(Me.txtLicenseKey)
    Me.Controls.Add(Me.txtusername)
    Me.Controls.Add(Me.lblinstcode)
    Me.Controls.Add(Me.lblreglevel)
    Me.Controls.Add(Me.txtinstcode)
    Me.Controls.Add(Me.txtreglevel)
    Me.Controls.Add(Me.lblLockType)
    Me.Controls.Add(Me.lblLicType)
    Me.Controls.Add(Me.txtLockType)
    Me.Controls.Add(Me.txtLicType)
    Me.Controls.Add(Me.lblexpdate)
    Me.Controls.Add(Me.lblregdate)
    Me.Controls.Add(Me.txtexpdate)
    Me.Controls.Add(Me.txtregdate)
    Me.Controls.Add(Me.lblprogver)
    Me.Controls.Add(Me.lblprogname)
    Me.Controls.Add(Me.txtprogver)
    Me.Controls.Add(Me.txtprogname)
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmAlugendb_details"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "License key details"
    Me.ResumeLayout(False)

  End Sub

#End Region

  Dim progname, progver, regdate, expdate, lictype, locktype, reglevel, instcode, username, libcode As String

  Public Sub New(ByVal new_progname As String, ByVal new_progver As String, ByVal new_regdate As String, ByVal new_expdate As String, ByVal new_lictype As String, ByVal new_locktype As String, ByVal new_reglevel As String, ByVal new_instcode As String, ByVal new_username As String, ByVal new_libcode As String)

    MyBase.New()
    InitializeComponent()

    progname = new_progname
    progver = new_progver
    regdate = new_regdate
    expdate = new_expdate
    lictype = new_lictype
    locktype = new_locktype
    reglevel = new_reglevel
    instcode = new_instcode
    username = new_username
    libcode = new_libcode

  End Sub

  Private Sub frmAlugendb_details_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    'load buttons pictures
    LoadImages()

    Me.txtprogname.Text = progname
    Me.txtprogver.Text = progver
    Me.txtregdate.Text = regdate
    Me.txtexpdate.Text = expdate
    Me.txtLicType.Text = lictype
    Me.txtLockType.Text = locktype
    Me.txtreglevel.Text = reglevel
    Me.txtinstcode.Text = instcode
    Me.txtusername.Text = username
    Me.txtLicenseKey.Text = libcode
  End Sub

  Private Sub picALBanner_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles picALBanner.Click
    'navigate to www.activelocksoftware.com
    System.Diagnostics.Process.Start(ACTIVELOCKSOFTWAREWEB)
  End Sub

  Private Sub cmdCopy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCopy.Click
    Dim aDataObject As New DataObject
    aDataObject.SetData(DataFormats.Text, txtLicenseKey.Text)
    Clipboard.SetDataObject(aDataObject)
  End Sub

  Private Sub cmdPrintLicenseKey_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPrintLicenseKey.Click
    Dim daPrintDocument As New daReport.DaPrintDocument
    Dim hashParameters As New Hashtable

    'set .xml file for printing
    daPrintDocument.setXML("reports\repLicenseKey.xml")

    'build parameters
    hashParameters.Add("pProductName", txtprogname.Text)
    hashParameters.Add("pProductVersion", txtprogver.Text)
    hashParameters.Add("pRegisteredLevel", txtreglevel.Text)
    hashParameters.Add("pLicenseType", txtLicType.Text)
    hashParameters.Add("pRegisteredDate", txtregdate.Text)
    hashParameters.Add("pExpireDate", txtexpdate.Text)
    hashParameters.Add("pInstallCode", txtinstcode.Text)
    hashParameters.Add("pUserName", txtusername.Text)
    hashParameters.Add("pLicenseKey", txtLicenseKey.Text)

    'setting parameters
    daPrintDocument.SetParameters(hashParameters)
    'daPrintDocument.DocumentName = "License key"

    'print preview
    printPreviewDialog1.Icon = CType(frmMain.resxList("report.ico"), Icon)
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
    Dim strNewLine As String = "%0D%0A"

    strSubject = String.Format("License key for application {0} ({1}), user [{2}]", txtprogname.Text, txtprogver.Text, txtusername.Text)
    strBodyMessage = strNewLine & String.Format("Install code:" & strNewLine & "{0}", txtinstcode.Text)
    strBodyMessage = strBodyMessage & strNewLine & strNewLine & String.Format("License key:" & strNewLine & "{0}", txtLicenseKey.Text.Replace(Chr(13), strNewLine))

    'final constructor
    mailToString = String.Format("mailto:{0}?subject={1}&body={2}", emailAddress, strSubject, strBodyMessage)

    'launch default email client
    System.Diagnostics.Process.Start(mailToString)
  End Sub

  Private Sub LoadImages()
    'load buttons images
    picALBanner.Image = CType(frmMain.resxList("I_Trust_AL_small.gif"), Image)
    cmdCopy.Image = CType(frmMain.resxList("copy.gif"), Image)
    cmdEmailLicenseKey.Image = CType(frmMain.resxList("email.gif"), Image)
    cmdPrintLicenseKey.Image = CType(frmMain.resxList("print.gif"), Image)
  End Sub
End Class
