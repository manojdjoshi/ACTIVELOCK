Option Strict On
Option Explicit On 
Imports ActiveLock3_6NET

Public Class frmProductsStorage
  Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

  Public Sub New()
    MyBase.New()

    'This call is required by the Windows Form Designer.
    InitializeComponent()

    'Add any initialization after the InitializeComponent() call

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

  'NOTE: The following procedure is required by the Windows Form Designer
  'It can be modified using the Windows Form Designer.  
  'Do not modify it using the code editor.
  Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
  Friend WithEvents optStorageIniFile As System.Windows.Forms.RadioButton
  Friend WithEvents optStorageXmlFile As System.Windows.Forms.RadioButton
  Friend WithEvents optStorageMDBFile As System.Windows.Forms.RadioButton
  Friend WithEvents optStorageMSSQL As System.Windows.Forms.RadioButton
  Friend WithEvents Label1 As System.Windows.Forms.Label
  Friend WithEvents txtStorageFile As System.Windows.Forms.TextBox
  Friend WithEvents Label2 As System.Windows.Forms.Label
  Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
  Friend WithEvents cmdBrowseForStorageFile As System.Windows.Forms.Button
  Friend WithEvents cmdOK As System.Windows.Forms.Button
  Friend WithEvents cmdCancel As System.Windows.Forms.Button
  Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
  <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmProductsStorage))
    Me.GroupBox1 = New System.Windows.Forms.GroupBox
    Me.cmdBrowseForStorageFile = New System.Windows.Forms.Button
    Me.TextBox1 = New System.Windows.Forms.TextBox
    Me.Label2 = New System.Windows.Forms.Label
    Me.txtStorageFile = New System.Windows.Forms.TextBox
    Me.Label1 = New System.Windows.Forms.Label
    Me.optStorageMSSQL = New System.Windows.Forms.RadioButton
    Me.optStorageMDBFile = New System.Windows.Forms.RadioButton
    Me.optStorageXmlFile = New System.Windows.Forms.RadioButton
    Me.optStorageIniFile = New System.Windows.Forms.RadioButton
    Me.cmdOK = New System.Windows.Forms.Button
    Me.cmdCancel = New System.Windows.Forms.Button
    Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
    Me.GroupBox1.SuspendLayout()
    Me.SuspendLayout()
    '
    'GroupBox1
    '
    Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
    Me.GroupBox1.Controls.Add(Me.cmdBrowseForStorageFile)
    Me.GroupBox1.Controls.Add(Me.TextBox1)
    Me.GroupBox1.Controls.Add(Me.Label2)
    Me.GroupBox1.Controls.Add(Me.txtStorageFile)
    Me.GroupBox1.Controls.Add(Me.Label1)
    Me.GroupBox1.Controls.Add(Me.optStorageMSSQL)
    Me.GroupBox1.Controls.Add(Me.optStorageMDBFile)
    Me.GroupBox1.Controls.Add(Me.optStorageXmlFile)
    Me.GroupBox1.Controls.Add(Me.optStorageIniFile)
    Me.GroupBox1.Location = New System.Drawing.Point(8, 8)
    Me.GroupBox1.Name = "GroupBox1"
    Me.GroupBox1.Size = New System.Drawing.Size(440, 134)
    Me.GroupBox1.TabIndex = 0
    Me.GroupBox1.TabStop = False
    Me.GroupBox1.Text = "Storage Type"
    '
    'cmdBrowseForStorageFile
    '
    Me.cmdBrowseForStorageFile.Location = New System.Drawing.Point(414, 32)
    Me.cmdBrowseForStorageFile.Name = "cmdBrowseForStorageFile"
    Me.cmdBrowseForStorageFile.Size = New System.Drawing.Size(23, 18)
    Me.cmdBrowseForStorageFile.TabIndex = 8
    Me.cmdBrowseForStorageFile.Text = "..."
    '
    'TextBox1
    '
    Me.TextBox1.Enabled = False
    Me.TextBox1.Location = New System.Drawing.Point(128, 110)
    Me.TextBox1.Name = "TextBox1"
    Me.TextBox1.Size = New System.Drawing.Size(306, 20)
    Me.TextBox1.TabIndex = 7
    Me.TextBox1.Text = ""
    '
    'Label2
    '
    Me.Label2.Enabled = False
    Me.Label2.Location = New System.Drawing.Point(128, 94)
    Me.Label2.Name = "Label2"
    Me.Label2.Size = New System.Drawing.Size(100, 16)
    Me.Label2.TabIndex = 6
    Me.Label2.Text = "Connection String:"
    '
    'txtStorageFile
    '
    Me.txtStorageFile.Location = New System.Drawing.Point(128, 30)
    Me.txtStorageFile.Name = "txtStorageFile"
    Me.txtStorageFile.Size = New System.Drawing.Size(284, 20)
    Me.txtStorageFile.TabIndex = 5
    Me.txtStorageFile.Text = ""
    '
    'Label1
    '
    Me.Label1.Location = New System.Drawing.Point(128, 16)
    Me.Label1.Name = "Label1"
    Me.Label1.TabIndex = 4
    Me.Label1.Text = "Storage file:"
    '
    'optStorageMSSQL
    '
    Me.optStorageMSSQL.Enabled = False
    Me.optStorageMSSQL.Location = New System.Drawing.Point(8, 88)
    Me.optStorageMSSQL.Name = "optStorageMSSQL"
    Me.optStorageMSSQL.Size = New System.Drawing.Size(120, 24)
    Me.optStorageMSSQL.TabIndex = 3
    Me.optStorageMSSQL.Text = "MSSQL database"
    '
    'optStorageMDBFile
    '
    Me.optStorageMDBFile.Location = New System.Drawing.Point(8, 64)
    Me.optStorageMDBFile.Name = "optStorageMDBFile"
    Me.optStorageMDBFile.Size = New System.Drawing.Size(120, 24)
    Me.optStorageMDBFile.TabIndex = 2
    Me.optStorageMDBFile.Text = "MDB file"
    '
    'optStorageXmlFile
    '
    Me.optStorageXmlFile.Location = New System.Drawing.Point(8, 40)
    Me.optStorageXmlFile.Name = "optStorageXmlFile"
    Me.optStorageXmlFile.Size = New System.Drawing.Size(120, 24)
    Me.optStorageXmlFile.TabIndex = 1
    Me.optStorageXmlFile.Text = "XML file"
    '
    'optStorageIniFile
    '
    Me.optStorageIniFile.Checked = True
    Me.optStorageIniFile.Location = New System.Drawing.Point(8, 16)
    Me.optStorageIniFile.Name = "optStorageIniFile"
    Me.optStorageIniFile.Size = New System.Drawing.Size(120, 24)
    Me.optStorageIniFile.TabIndex = 0
    Me.optStorageIniFile.TabStop = True
    Me.optStorageIniFile.Text = "INI file"
    '
    'cmdOK
    '
    Me.cmdOK.Location = New System.Drawing.Point(292, 148)
    Me.cmdOK.Name = "cmdOK"
    Me.cmdOK.TabIndex = 1
    Me.cmdOK.Text = "OK"
    '
    'cmdCancel
    '
    Me.cmdCancel.Location = New System.Drawing.Point(374, 148)
    Me.cmdCancel.Name = "cmdCancel"
    Me.cmdCancel.TabIndex = 2
    Me.cmdCancel.Text = "Cancel"
    '
    'frmProductsStorage
    '
    Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
    Me.ClientSize = New System.Drawing.Size(456, 177)
    Me.Controls.Add(Me.cmdCancel)
    Me.Controls.Add(Me.cmdOK)
    Me.Controls.Add(Me.GroupBox1)
    Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
    Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
    Me.Name = "frmProductsStorage"
    Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
    Me.Text = "Products Storage"
    Me.GroupBox1.ResumeLayout(False)
    Me.ResumeLayout(False)

  End Sub

#End Region

  Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
    Me.Close()
  End Sub

  Private Sub cmdOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdOK.Click
    SaveSettings()
    mainForm.frmMain_Load(sender, e)
    Me.Close()
  End Sub

  Private Sub frmProductsStorage_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    LoadSettings()
  End Sub

  Private Sub LoadSettings()
    On Error GoTo LoadSettings_Error
    txtStorageFile.Text = mainForm.mProductsStoragePath
    Select Case mainForm.mProductsStoreType
      Case IActiveLock.ProductsStoreType.alsINIFile
        optStorageIniFile.Checked = True
      Case IActiveLock.ProductsStoreType.alsXMLFile
        optStorageXmlFile.Checked = True
      Case IActiveLock.ProductsStoreType.alsMDBFile
        optStorageMDBFile.Checked = True
        'Case alsMSSQL
    End Select

    On Error GoTo 0
    Exit Sub

LoadSettings_Error:

    MsgBox("Error " & Err.Number & " (" & Err.Description & ") in procedure LoadSettings of Form frmProductsStorage")
  End Sub

  Private Sub SaveSettings()
    On Error GoTo SaveSettings_Error
    mainForm.mProductsStoragePath = txtStorageFile.Text
    mainForm.mProductsStoreType = CType(IIf(optStorageIniFile.Checked, IActiveLock.ProductsStoreType.alsINIFile, _
      IIf(optStorageXmlFile.Checked, IActiveLock.ProductsStoreType.alsXMLFile, _
      IIf(optStorageMDBFile.Checked, IActiveLock.ProductsStoreType.alsMDBFile, _
      IActiveLock.ProductsStoreType.alsINIFile))) _
      , IActiveLock.ProductsStoreType)

    On Error GoTo 0
    Exit Sub

SaveSettings_Error:

    MsgBox("Error " & Err.Number & " (" & Err.Description & ") in procedure SaveSettings of Form frmProductsStorage")
  End Sub

  Private Sub cmdBrowseForStorageFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowseForStorageFile.Click
    'browse for storage file
    Try
      OpenFileDialog1.ValidateNames = True
      OpenFileDialog1.InitialDirectory = System.IO.Path.GetFullPath(txtStorageFile.Text)
      OpenFileDialog1.RestoreDirectory = True
      OpenFileDialog1.Title = "Select products storage file"

            If OpenFileDialog1.ShowDialog(Me) = Windows.Forms.DialogResult.OK Then
                txtStorageFile.Text = OpenFileDialog1.FileName
            End If
    Catch
      MessageBox.Show("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdBrowseForStorageFile_Click of Form frmProductsStorage", modALUGEN.ACTIVELOCKSTRING)
    End Try
  End Sub
End Class
