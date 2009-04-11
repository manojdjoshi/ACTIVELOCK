Imports ActiveLock3_6NET

Public Class MainForm

    'Define vars 
    Private Shared ActiveLock As ActiveLock3_6NET._IActiveLock
    Private Shared ALGlobals As New ActiveLock3_6NET.Globals
    Private Shared WithEvents ActiveLockEventSink As ActiveLock3_6NET.ActiveLockEventNotifier
    Private strKeyStorePath As String

    Private Sub MainForm_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        End
    End Sub


    Private Sub MainForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call InitiializeActiveLock()
        Call ValidateLicence()
        Call Updatecontrols()

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
            ' Initialize AL
            ' Important: If you're not going to put Alcrypto3NET.dll under
            ' the system32 directory, you should pass the path of the exe
            ' to the Init() method otherwise this call will fail
            ' Putting Alcrypto3NET.dll under the system32 is a problem with ASP.NET apps
            ' since Activelock3NET is shared between ASP.NET and VB.NET apps.
            ActiveLock.Init(Application.StartupPath, strKeyStorePath)
            ActiveLock.Acquire(Mensaje) ' Acquire will raise an error if no valid license exists. 
            If Mensaje <> "" Then
                'There's a trial 
                A = Split(Mensaje, vbCrLf)
                txtStatus.Text = A(0) & " - " & A(1) & " days"
            Else
                txtStatus.Text = "Registered"
            End If
            'txtUser.Text = ActiveLock.RegisteredUser
            'txtLicenceKey.Text = ActiveLock.InstallationCode

            'TextBox1.Text = ActiveLock.InstallationCode(txtUser) 
        Catch ALInit As Exception
            MsgBox("Aplication is Unregistered, please give your name and generete your INSTALL CODE" & Chr(13) & _
                   "And send us to mymail@mycompany.com and we send you your Licence Key to register this aplication.")
            txtStatus.Text = "UnRegistered"
            Exit Try
        End Try
    End Sub

    Private Sub Updatecontrols()
        txtProductName.Text = ActiveLock.SoftwareName
        txtProductVersion.Text = ActiveLock.SoftwareVersion
        If Not ActiveLock Is Nothing Then
            txtLicenceType.Text = ActiveLock.LockType
            If txtStatus.Text = "Registered" Then
                txtRegisteredLevel.Text = ActiveLock.RegisteredLevel
                txtUser.Text = ActiveLock.RegisteredUser
                MsgBox("Welcome to MyNewSoftware, to access the main window, just click on the App button")
            End If
        End If
    End Sub

    Private Sub txtProductVersion_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtProductVersion.TextChanged

    End Sub

    Private Sub txtUser_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUser.TextChanged
        Try
            txtInstCode.Text = ActiveLock.InstallationCode(txtUser.Text) 'ActiveLock.InstallCode(txtUser) 
        Catch ex As Exception
            MsgBox("EXCEPTION : " & ex.Message)
        End Try

    End Sub

    Private Sub ButtonRegister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonRegister.Click
        Try
            ActiveLock.Register(txtLicence.Text, txtUser.Text)
            MsgBox("Application REGISTERED successfully!, " & Chr(13) & "Please restart the aplication.")
            End
        Catch ex As Exception
            MsgBox("EXCEPTION: " & ex.Message)
        End Try
    End Sub

    Private Sub ButtonApp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'Dim y As New MenuPrin
        'y.ShowDialog()
    End Sub
End Class
