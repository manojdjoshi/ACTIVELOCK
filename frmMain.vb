'You Can call it from commandline with someting like this
'"Activelock Wizard.exe" "AppName=Test App" "AppVersion=1.0.0" "PUB_KEY=frf4378t475gy54vgyw5g5y6cg556g576fg6f"
Public Class frmMain
#Region "Form Routines"

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'LoadDefaultValuesForTesting()


        LoadValuesFromAlugen()
    End Sub

#End Region '"Form Routines"

#Region "Form Button Routines"

    Private Sub rbtnTrialTypeNone_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtnTrialTypeNone.CheckedChanged
        If rbtnTrialTypeNone.Checked = True Then
            txtTrialLength.Enabled = False
            GroupTrialHideTypes.Enabled = False
            lblTrialLength.Enabled = False
        Else
            txtTrialLength.Enabled = True
            GroupTrialHideTypes.Enabled = True
            lblTrialLength.Enabled = True
        End If
    End Sub

    Private Sub btnSubmitDetails_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmitDetails.Click
        Try
            ClearAllVariableValues()
            GetSoftwareName()
            GetSoftwareVersion()
            GetSoftwarePassword()
            GetSoftwareCode()
            GetLicenceKeyType()
            GetLicenceFileType()
            GetAutoRegister()
            GetTimeServerClockTampering()
            GetSystemFilesClockTampering()
            GetTrialType()
            GetTrialHideType()
            GetTrialLength()
            GetLockTypes()
            GetKeyStoreTypes()
            GetDevEnvironment()
            LaunchCreateDevEnvModBuilder()
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub chkLicenseLockTypeNone_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeNone.CheckedChanged
        If chkLicenseLockTypeNone.Checked = True Then
            chkLicenseLockTypeBios.Checked = False
            chkLicenseLockTypeComp.Checked = False
            chkLicenseLockTypeHD.Checked = False
            chkLicenseLockTypeHDFW.Checked = False
            chkLicenseLockTypeIP.Checked = False
            chkLicenseLockTypeMAC.Checked = False
            chkLicenseLockTypeMotherboard.Checked = False
            chkLicenseLockTypeWindows.Checked = False
        Else
            If chkLicenseLockTypeBios.Checked = False And _
                chkLicenseLockTypeComp.Checked = False And _
                chkLicenseLockTypeHD.Checked = False And _
                chkLicenseLockTypeHDFW.Checked = False And _
                chkLicenseLockTypeIP.Checked = False And _
                chkLicenseLockTypeMAC.Checked = False And _
                chkLicenseLockTypeMotherboard.Checked = False And _
                chkLicenseLockTypeWindows.Checked = False Then
                chkLicenseLockTypeNone.Checked = True
            End If
        End If
    End Sub

    Private Sub chkLicenseLockTypeBios_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeBios.CheckedChanged
        If chkLicenseLockTypeBios.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            If chkLicenseLockTypeBios.Checked = False And _
                chkLicenseLockTypeComp.Checked = False And _
                chkLicenseLockTypeHD.Checked = False And _
                chkLicenseLockTypeHDFW.Checked = False And _
                chkLicenseLockTypeIP.Checked = False And _
                chkLicenseLockTypeMAC.Checked = False And _
                chkLicenseLockTypeMotherboard.Checked = False And _
                chkLicenseLockTypeWindows.Checked = False Then
                chkLicenseLockTypeNone.Checked = True
            End If
        End If

    End Sub

    Private Sub chkLicenseLockTypeComp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeComp.CheckedChanged
        If chkLicenseLockTypeComp.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            If chkLicenseLockTypeBios.Checked = False And _
                chkLicenseLockTypeComp.Checked = False And _
                chkLicenseLockTypeHD.Checked = False And _
                chkLicenseLockTypeHDFW.Checked = False And _
                chkLicenseLockTypeIP.Checked = False And _
                chkLicenseLockTypeMAC.Checked = False And _
                chkLicenseLockTypeMotherboard.Checked = False And _
                chkLicenseLockTypeWindows.Checked = False Then
                chkLicenseLockTypeNone.Checked = True
            End If
        End If

    End Sub

    Private Sub chkLicenseLockTypeHD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeHD.CheckedChanged
        If chkLicenseLockTypeHD.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            If chkLicenseLockTypeBios.Checked = False And _
                chkLicenseLockTypeComp.Checked = False And _
                chkLicenseLockTypeHD.Checked = False And _
                chkLicenseLockTypeHDFW.Checked = False And _
                chkLicenseLockTypeIP.Checked = False And _
                chkLicenseLockTypeMAC.Checked = False And _
                chkLicenseLockTypeMotherboard.Checked = False And _
                chkLicenseLockTypeWindows.Checked = False Then
                chkLicenseLockTypeNone.Checked = True
            End If
        End If

    End Sub

    Private Sub chkLicenseLockTypeHDFW_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeHDFW.CheckedChanged
        If chkLicenseLockTypeHDFW.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            If chkLicenseLockTypeBios.Checked = False And _
                chkLicenseLockTypeComp.Checked = False And _
                chkLicenseLockTypeHD.Checked = False And _
                chkLicenseLockTypeHDFW.Checked = False And _
                chkLicenseLockTypeIP.Checked = False And _
                chkLicenseLockTypeMAC.Checked = False And _
                chkLicenseLockTypeMotherboard.Checked = False And _
                chkLicenseLockTypeWindows.Checked = False Then
                chkLicenseLockTypeNone.Checked = True
            End If
        End If

    End Sub

    Private Sub chkLicenseLockTypeIP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeIP.CheckedChanged
        If chkLicenseLockTypeIP.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            If chkLicenseLockTypeBios.Checked = False And _
                chkLicenseLockTypeComp.Checked = False And _
                chkLicenseLockTypeHD.Checked = False And _
                chkLicenseLockTypeHDFW.Checked = False And _
                chkLicenseLockTypeIP.Checked = False And _
                chkLicenseLockTypeMAC.Checked = False And _
                chkLicenseLockTypeMotherboard.Checked = False And _
                chkLicenseLockTypeWindows.Checked = False Then
                chkLicenseLockTypeNone.Checked = True
            End If
        End If

    End Sub

    Private Sub chkLicenseLockTypeMAC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeMAC.CheckedChanged
        If chkLicenseLockTypeMAC.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            If chkLicenseLockTypeBios.Checked = False And _
                chkLicenseLockTypeComp.Checked = False And _
                chkLicenseLockTypeHD.Checked = False And _
                chkLicenseLockTypeHDFW.Checked = False And _
                chkLicenseLockTypeIP.Checked = False And _
                chkLicenseLockTypeMAC.Checked = False And _
                chkLicenseLockTypeMotherboard.Checked = False And _
                chkLicenseLockTypeWindows.Checked = False Then
                chkLicenseLockTypeNone.Checked = True
            End If
        End If

    End Sub

    Private Sub chkLicenseLockTypeMotherboard_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeMotherboard.CheckedChanged
        If chkLicenseLockTypeMotherboard.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            If chkLicenseLockTypeBios.Checked = False And _
                chkLicenseLockTypeComp.Checked = False And _
                chkLicenseLockTypeHD.Checked = False And _
                chkLicenseLockTypeHDFW.Checked = False And _
                chkLicenseLockTypeIP.Checked = False And _
                chkLicenseLockTypeMAC.Checked = False And _
                chkLicenseLockTypeMotherboard.Checked = False And _
                chkLicenseLockTypeWindows.Checked = False Then
                chkLicenseLockTypeNone.Checked = True
            End If
        End If

    End Sub

    Private Sub chkLicenseLockTypeWindows_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeWindows.CheckedChanged
        If chkLicenseLockTypeWindows.Checked = True Then chkLicenseLockTypeNone.Checked = False
    End Sub

#End Region '"Form Button Routines"

#Region "Get Values From Form"

    Private Sub GetSoftwareName()
        'SoftwareName
        If txtSoftwareName.Text.Length = 0 Then
            Err.Raise(1, "Private Sub GetSoftwareName", "Please Enter a Name For The Software")
        Else
            SoftwareName = txtSoftwareName.Text
        End If
    End Sub

    Private Sub GetSoftwareVersion()
        'SoftwareVersion
        If txtSoftwareVersion.Text.Length = 0 Then
            Err.Raise(1, "Private Sub GetSoftwareVersion", "Please Enter a Version For The Software")
        Else
            SoftwareVersion = txtSoftwareVersion.Text
        End If
    End Sub

    Private Sub GetSoftwarePassword()
        'SoftwarePassword
        If txtSoftwarePassword.Text.Length = 0 Then
            Err.Raise(1, "Private Sub GetSoftwarePassword", "Please Enter a Password For The Software")
        Else
            SoftwarePassword = txtSoftwarePassword.Text
        End If
    End Sub

    Private Sub GetSoftwareCode()
        'SoftwareCode
        If txtSoftwareCode.Text.Length = 0 Then
            Err.Raise(1, "Private Sub GetSoftwareCode", "Please Enter a Software Code")
        Else
            SoftwareCode = txtSoftwareCode.Text
        End If
    End Sub

    Private Sub GetLicenceKeyType()
        'LicenceKeyType
        If rbtnLicenceKeyTypeRSA.Checked = True Then LicenceKeyType = LicenceKeyType_t.alsRSA
        If rbtnLicenceKeyTypeShortMD5.Checked = True Then LicenceKeyType = LicenceKeyType_t.alsShortKeyMD5
    End Sub

    Private Sub GetLicenceFileType()
        'LicenceFileType
        If rbtnLicenceFileTypePlain.Checked = True Then LicenceFileType = LicenceFileType_t.alsLicenseFilePlain
        If rbtnLicenceFileTypeEncrypted.Checked = True Then LicenceFileType = LicenceFileType_t.alsLicenseFileEncrypted
    End Sub

    Private Sub GetAutoRegister()
        'AutoRegister
        If rbtnAutoRegisterEnable.Checked = True Then AutoRegister = AutoRegister_t.alsEnableAutoRegistration
        If rbtnAutoRegisterDisable.Checked = True Then AutoRegister = AutoRegister_t.alsDisableAutoRegistration
    End Sub

    Private Sub GetTimeServerClockTampering()
        'TimeServerClockTampering
        If rbtnTimeServerEnable.Checked = True Then TimeServerClockTampering = TimeServerClockTampering_t.alsCheckTimeServer
        If rbtnTimeServerDisable.Checked = True Then TimeServerClockTampering = TimeServerClockTampering_t.alsDontCheckTimeServer
    End Sub

    Private Sub GetSystemFilesClockTampering()
        'SystemFilesClockTampering
        If rbtnSystemFilesEnable.Checked = True Then SystemFilesClockTampering = SystemFilesClockTampering_t.alsCheckSystemFiles
        If rbtnSystemFilesDisable.Checked = True Then SystemFilesClockTampering = SystemFilesClockTampering_t.alsDontCheckSystemFiles
    End Sub

    Private Sub GetTrialType()
        'TrialType
        If rbtnTrialTypeDays.Checked = True Then TrialType = TrialType_t.trialDays
        If rbtnTrialTypeRuns.Checked = True Then TrialType = TrialType_t.trialRuns
        If rbtnTrialTypeNone.Checked = True Then TrialType = TrialType_t.trialNone
    End Sub

    Private Sub GetTrialHideType()
        'TrialHideType
        If TrialType <> TrialType_t.trialNone Then
            If chkTrialHideTypeHiddenFolder.Checked = True Then TrialHideType.Add("trialHiddenFolder")
            If chkTrialHideTypeRegistry.Checked = True Then TrialHideType.Add("trialRegistry")
            If chkTrialHideTypeSteganography.Checked = True Then TrialHideType.Add("trialSteganography")
        End If
    End Sub

    Private Sub GetTrialLength()
        'TrialLength
        If TrialType <> TrialType_t.trialNone Then
            Try
                TrialLength = CType(txtTrialLength.Text, Integer)
                If TrialLength < 0 Then TrialLength = 15
            Catch ex As Exception
                TrialLength = 15
            End Try
        End If
    End Sub

    Private Sub GetLockTypes()
        'LockTypes
        If chkLicenseLockTypeNone.Checked = True Then LockTypes.Add("lockNone")
        If chkLicenseLockTypeBios.Checked = True Then LockTypes.Add("lockBIOS")
        If chkLicenseLockTypeComp.Checked = True Then LockTypes.Add("lockComp")
        If chkLicenseLockTypeHD.Checked = True Then LockTypes.Add("lockHD")
        If chkLicenseLockTypeHDFW.Checked = True Then LockTypes.Add("lockHDFirmware")
        If chkLicenseLockTypeIP.Checked = True Then LockTypes.Add("lockIP")
        If chkLicenseLockTypeMAC.Checked = True Then LockTypes.Add("lockMAC")
        If chkLicenseLockTypeMotherboard.Checked = True Then LockTypes.Add("lockMotherboard")
        If chkLicenseLockTypeWindows.Checked = True Then LockTypes.Add("lockWindows")
        If LockTypes.Count = 0 Then LockTypes.Add("lockNone")
    End Sub

    Private Sub GetKeyStoreTypes()
        'KeyStoreType
        If rbtnKeyStorageTypeFile.Checked = True Then KeyStoreType = KeyStoreType_t.alsFile
        If rbtnKeyStorageTypeRegistry.Checked = True Then KeyStoreType = KeyStoreType_t.alsRegistry
    End Sub

    Private Sub GetDevEnvironment()
        'DevEnvironment
        If cboDevEnviroment.Text.ToUpper = "VB6" Then DevEnvironment = DevEnvironment_t.vb6
        If cboDevEnviroment.Text.ToUpper = "VB2003" Then DevEnvironment = DevEnvironment_t.vb2003
        If cboDevEnviroment.Text.ToUpper = "VB2005" Then DevEnvironment = DevEnvironment_t.vb2005
    End Sub

#End Region '"Get Values From Form"

#Region "General Form Routines"

    Private Sub ClearAllVariableValues()
        SoftwareName = Nothing
        SoftwareVersion = Nothing
        SoftwarePassword = Nothing
        SoftwareCode = Nothing
        LicenceKeyType = Nothing
        LicenceFileType = Nothing
        AutoRegister = Nothing
        TimeServerClockTampering = Nothing
        SystemFilesClockTampering = Nothing
        TrialType = Nothing
        LockTypes.Clear()
        KeyStoreType = Nothing
        TrialHideType.Clear()
        TrialLength = 0
        DevEnvironment = Nothing
    End Sub

    Private Sub LaunchCreateDevEnvModBuilder()
        If DevEnvironment = DevEnvironment_t.vb6 Then CreateVB6Module()
        If DevEnvironment = DevEnvironment_t.vb2003 Then CreateVB2003Module()
        If DevEnvironment = DevEnvironment_t.vb2005 Then CreateVB2005Module()
    End Sub

    Private Sub LoadValuesFromAlugen()
        txtSoftwareName.Text = SoftwareName
        txtSoftwareVersion.Text = SoftwareVersion
        txtSoftwareCode.Text = SoftwareCode
    End Sub

#End Region '"General Form Routines"

#Region "ForDev"

    Private Sub LoadDefaultValuesForTesting()
        'SoftwareName
        txtSoftwareName.Text = "TestApp"
        'SoftwareVersion
        txtSoftwareVersion.Text = "1.0.0"
        'SoftwarePassword
        txtSoftwarePassword.Text = "walterPassword"
        'SoftwareCode
        txtSoftwareCode.Text = "2343r34r4hrj4fhj54f5hjghjgg6j5gh6g3vh5bh5ch5hc35hvbbvbvh5v6h"
        'LicenceKeyType
        rbtnLicenceKeyTypeRSA.Checked = True
        'LicenceFileType
        rbtnLicenceFileTypeEncrypted.Checked = True
        'AutoRegister
        rbtnAutoRegisterEnable.Checked = True
        'TimeServerClockTampering
        rbtnTimeServerEnable.Checked = True
        'SystemFilesClockTampering
        rbtnSystemFilesEnable.Checked = True
        'TrialType
        rbtnTrialTypeDays.Checked = True
        'TrialHideType
        chkTrialHideTypeHiddenFolder.Checked = True
        chkTrialHideTypeRegistry.Checked = True
        chkTrialHideTypeSteganography.Checked = True
        'TrialLength
        txtTrialLength.Text = "10"
        'DevEnvironment
        cboDevEnviroment.Text = "VB2005"
    End Sub

#End Region '"ForDev"

    Private Sub ImageApplicationHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp.Click, PictureBox12.Click, PictureBox1.Click
        MsgBox("Enter Your Application Name And Version", MsgBoxStyle.Information)
    End Sub

    Private Sub PicActiveLock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicActiveLock.Click
        System.Diagnostics.Process.Start("http://www.activelocksoftware.com")
    End Sub

    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click
        HelpProvider.SetHelpString(Me.txtSoftwareName, "Enter the street address in this text box.")
        HelpProvider.SetShowHelp(Me.txtSoftwareName, True)
    End Sub
End Class
