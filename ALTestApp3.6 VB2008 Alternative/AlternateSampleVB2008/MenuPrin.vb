Imports ActiveLock3_6NET


Public Class MenuPrin
    'Define vars 
    Private Shared ActiveLock As ActiveLock3_6NET._IActiveLock
    Private Shared ALGlobals As New ActiveLock3_6NET.Globals
    Private Shared WithEvents ActiveLockEventSink As ActiveLock3_6NET.ActiveLockEventNotifier
    Private strKeyStorePath As String

    Private Sub MenuPrin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitiializeActiveLock()
        Call ValidateLicence()
    End Sub

    Private Sub InitiializeActiveLock()

        ' Obtain an instance of ActiveLock object 
        ActiveLock = ALGlobals.NewInstance()
        ActiveLockEventSink = ActiveLock.EventNotifier

        With ActiveLock
            ' Specify License parameters 
            .KeyStoreType = ActiveLock3_6NET.IActiveLock.LicStoreType.alsFile
            strKeyStorePath = System.Windows.Forms.Application.StartupPath & "\License.lic"
            .KeyStorePath = strKeyStorePath
            .AutoRegisterKeyPath = System.Windows.Forms.Application.StartupPath & "\License.all"
            'ActiveLock.LicenseFileType = ActiveLock3_6NET.IActiveLock.ALLicenseFileTypes.alsLicenseFileEncrypted 
            ' 

            ' Specify Software parameters 
            ' Specify Lock Type as Lock-to-HardDrive Firmware and Lock Type as lockMAC. 
            .LockType = ActiveLock3_6NET.IActiveLock.ALLockTypes.lockHDFirmware Or ActiveLock3_6NET.IActiveLock.ALLockTypes.lockMAC
            .SoftwareName = "MyNewSoftware"
            .SoftwareVersion = "1.0"
            .SoftwarePassword = "Cool"
            .SoftwareCode = "AAAAB3NzaC1yc2EAAAABJQAAAIDKMvVASZ0SAT/umCzuyAQo8ZPb4iwaDns+fGeu0HHngvaLeX2zefNQmecM4PJAEH4ZntRY2a/mxtdUNq1Q2aiXnACbg9XcON67fO2ypZh5IGNH/L9v4x4/hHDYKq3kfSIOfcXSJRY6d6iF9DpxiKpGikRNnA186SgAuUE+ycVJrw=="
        End With

    End Sub

    Private Sub ValidateLicence()
        Dim A() As String
        Dim Mensaje As String = Nothing
        Try
            ActiveLock.Init()
            ActiveLock.Acquire(Mensaje)  ' Acquire will raise an error if no valid license exists.
            If Mensaje <> "" Then
                'There's a trial
                A = Split(Mensaje, vbCrLf)
                MsgBox("Aplication is running on Trial MODE")
                'A(0) & " - " & A(1) & " days"
            Else
                StatusBarPanel1.Text = "DataBase File: "
                StatusBarPanel2.Text = "User: "
                StatusBarPanel3.Text = "Profile: "
            End If
        Catch ALInit As Exception
            Dim z As New MainForm
            z.ShowDialog()
        End Try
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        'Exit app
        End
    End Sub
End Class
