Imports ActiveLock3_6NET
Imports System.IO

Public Class Form1

	Public Shared lock As ActiveLock3_6NET._IActiveLock

    Private Sub cmdGetTrial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdGetTrial.Click
		Using proxy As New Activelock36Service.Activelock3_6ServiceSoapClient
			txtLicenceCode.Text = proxy.GetTrial(txtInstallCode.Text, _
				ApplicationIdentifier.ApplicationName, _
				ApplicationIdentifier.ApplicationVersion)
		End Using
	End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Using proxy As New Activelock36Service.Activelock3_6ServiceSoapClient
            txtLicenceCode.Text = proxy.GetLicense(txtInstallCode.Text, _
                    ApplicationIdentifier.ApplicationName, _
                    ApplicationIdentifier.ApplicationVersion, _
                    "All", _
                    Activelock36Service.LicFlags.alfSingle, _
                    1, _
                    Activelock36Service.ALLicType.allicPermanent, _
                    100)
        End Using
    End Sub

	Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
		txtInstallCode.Text = GetInstallCode()
	End Sub

	Function GetInstallCode() As String

		Dim myAl As New ActiveLock3_6NET.Globals
		lock = myAl.NewInstance()
		lock.SoftwareName = ApplicationIdentifier.ApplicationName
		lock.SoftwareVersion = ApplicationIdentifier.ApplicationVersion
		lock.LicenseKeyType = IActiveLock.ALLicenseKeyTypes.alsRSA

		lock.TrialType = IActiveLock.ALTrialTypes.trialDays
		lock.TrialLength = 15
		lock.TrialHideType = IActiveLock.ALTrialHideTypes.trialHiddenFolder Or IActiveLock.ALTrialHideTypes.trialRegistryPerUser Or IActiveLock.ALTrialHideTypes.trialSteganography
		lock.TrialWarning = ActiveLock3_6NET.IActiveLock.ALTrialWarningTypes.trialWarningPersistent

		lock.SoftwareCode = ApplicationIdentifier.Vcode
		lock.LockType = IActiveLock.ALLockTypes.lockHD
		lock.CheckTimeServerForClockTampering = ActiveLock3_6NET.IActiveLock.ALTimeServerTypes.alsDontCheckTimeServer
		lock.CheckSystemFilesForClockTampering = ActiveLock3_6NET.IActiveLock.ALSystemFilesTypes.alsDontCheckSystemFiles

		lock.LicenseFileType = IActiveLock.ALLicenseFileTypes.alsLicenseFilePlain
		lock.KeyStoreType = IActiveLock.LicStoreType.alsFile
		lock.AutoRegisterKeyPath = Path.GetDirectoryName(Application.StartupPath) & "\" & ApplicationIdentifier.ApplicationName & ".all"

		Dim strKeyStorePath = Path.GetDirectoryName(Application.StartupPath) & "\" & ApplicationIdentifier.ApplicationName & ".lic"
		lock.KeyStoreType = IActiveLock.LicStoreType.alsFile
		lock.KeyStorePath = strKeyStorePath
		lock.Init(Application.StartupPath, strKeyStorePath)

		Return lock.InstallationCode(txtUserName.Text, Nothing)
	End Function

	Private Sub txtUserName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtUserName.TextChanged
		txtInstallCode.Text = GetInstallCode()
	End Sub
End Class
