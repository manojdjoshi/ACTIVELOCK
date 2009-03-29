'You can call the program from commandline with someting like this
'"Activelock Wizard.exe" "AppName=Test App" "AppVersion=1.0.0" "PUB_KEY=frf4378t475gy54vgyw5g5y6cg556g576fg6f"
Option Strict Off
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

Public Class frmMain
#Region "Form Routines"

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        LoadValuesFromAlugen()
    End Sub

    Private Sub cboDevEnviroment_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cboDevEnviroment.SelectedIndexChanged

        If Strings.StrComp(cboDevEnviroment.Text, "VB2003", CompareMethod.Binary) = 0 Then
            btnSubmitDetails.Enabled = False
            MsgBox("Sorry VB2003 is not available in this" & vbCrLf & _
                   "version, please let us know if you " & vbCrLf & _
                   "want this included in the next version", MsgBoxStyle.Information)
            'ws' ElseIf Strings.StrComp(cboDevEnviroment.Text, "VB2008", CompareMethod.Binary) = 0 Then
            'ws' This is the same as for VB2005 (using DotNet Framework 2.0)
            'MsgBox("VB2008 is not implemented in this" & vbCrLf & _
            '       "version, it will create a VB2005 " & vbCrLf & _
            '       "file instead that can be imported", MsgBoxStyle.Information)
        Else
            btnSubmitDetails.Enabled = True
        End If
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
            If RadioButton1.Checked = True Then 'added
                GetCRC32Encoded() 'added
            Else
                GetOwnCRC32Encoded() 'added V3.6
            End If
            LaunchCreateDevEnvModBuilder()
        Catch ex As Exception
            MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub

#Region "License Lock Types"

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
            chkLicenseLockTypeVideoID.Checked = False
            chkLicenseLockTypeBaseboardID.Checked = False
            chkLicenseLockTypeCPUID.Checked = False
            chkLicenseLockTypeMemory.Checked = False
            chkLicenseLockTypeFingerprint.Checked = False
            chkLicenseLockTypeExternalID.Checked = False
        End If
    End Sub

    Private Sub chkLicenseLockTypeBios_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeBios.CheckedChanged
        If chkLicenseLockTypeBios.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeComp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeComp.CheckedChanged
        If chkLicenseLockTypeComp.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeHD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeHD.CheckedChanged
        If chkLicenseLockTypeHD.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeHDFW_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeHDFW.CheckedChanged
        If chkLicenseLockTypeHDFW.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeIP_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeIP.CheckedChanged
        If chkLicenseLockTypeIP.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeMAC_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeMAC.CheckedChanged
        If chkLicenseLockTypeMAC.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeMotherboard_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeMotherboard.CheckedChanged
        If chkLicenseLockTypeMotherboard.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeWindows_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeWindows.CheckedChanged
        If chkLicenseLockTypeWindows.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeVideoID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeVideoID.CheckedChanged
        If chkLicenseLockTypeVideoID.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeBaseboardID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeBaseboardID.CheckedChanged
        If chkLicenseLockTypeBaseboardID.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeCPUID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeCPUID.CheckedChanged
        If chkLicenseLockTypeCPUID.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeMemory_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeMemory.CheckedChanged
        If chkLicenseLockTypeMemory.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeFingerprint_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeFingerprint.CheckedChanged
        If chkLicenseLockTypeFingerprint.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub chkLicenseLockTypeExternalID_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkLicenseLockTypeExternalID.CheckedChanged
        If chkLicenseLockTypeExternalID.Checked = True Then
            chkLicenseLockTypeNone.Checked = False
        Else
            CheckifNoneIsSelected()
        End If
    End Sub

    Private Sub CheckifNoneIsSelected()
        If chkLicenseLockTypeBios.Checked = False And _
           chkLicenseLockTypeComp.Checked = False And _
           chkLicenseLockTypeHD.Checked = False And _
           chkLicenseLockTypeHDFW.Checked = False And _
           chkLicenseLockTypeIP.Checked = False And _
           chkLicenseLockTypeMAC.Checked = False And _
           chkLicenseLockTypeMotherboard.Checked = False And _
           chkLicenseLockTypeWindows.Checked = False And _
           chkLicenseLockTypeVideoID.Checked = False And _
           chkLicenseLockTypeBaseboardID.Checked = False And _
           chkLicenseLockTypeCPUID.Checked = False And _
           chkLicenseLockTypeMemory.Checked = False And _
           chkLicenseLockTypeFingerprint.Checked = False And _
           chkLicenseLockTypeExternalID.Checked = False Then
            chkLicenseLockTypeNone.Checked = True
        End If
    End Sub

#End Region '"License Lock Types"


    Private Sub PicActiveLock_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PicActiveLock.Click
        System.Diagnostics.Process.Start("http://www.activelocksoftware.com")
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
        If rbtnTrialTypeDays.Checked = False And rbtnTrialTypeRuns.Checked = False Then TrialType = TrialType_t.trialNone
    End Sub

    Private Sub GetTrialHideType()
        'TrialHideType
        If TrialType <> TrialType_t.trialNone Then
            If chkTrialHideTypeHiddenFolder.Checked = True Or chkTrialHideTypeSteganography.Checked = True Or chkTrialHideTypeIsolatedStorage.Checked = True Then
                If chkTrialHideTypeHiddenFolder.Checked = True Then TrialHideType.Add("trialHiddenFolder")
                If chkTrialHideTypeRegistry.Checked = True Then TrialHideType.Add("trialRegistryPerUser")
                If chkTrialHideTypeSteganography.Checked = True Then TrialHideType.Add("trialSteganography")
                If chkTrialHideTypeIsolatedStorage.Checked = True Then TrialHideType.Add("trialIsolatedStorage")
            Else
                Err.Raise(1, "GetTrialHideType", "Please Select The Trial hiding Types")
            End If
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
        If chkLicenseLockTypeVideoID.Checked = True Then LockTypes.Add("lockvideoID")
        If chkLicenseLockTypeBaseboardID.Checked = True Then LockTypes.Add("lockBaseboardID")
        If chkLicenseLockTypeCPUID.Checked = True Then LockTypes.Add("lockCPUID")
        If chkLicenseLockTypeMemory.Checked = True Then LockTypes.Add("lockMemory")
        If chkLicenseLockTypeFingerprint.Checked = True Then LockTypes.Add("lockFingerprint")
        If chkLicenseLockTypeExternalID.Checked = True Then LockTypes.Add("lockExternalIP")
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
        If cboDevEnviroment.Text.ToUpper = "VB2008" Then DevEnvironment = DevEnvironment_t.vb2008
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
        If DevEnvironment = DevEnvironment_t.vb2008 Then CreateVB2008Module()
    End Sub

    Private Sub LoadValuesFromAlugen()
        txtSoftwareName.Text = SoftwareName
        txtSoftwareVersion.Text = SoftwareVersion
        txtSoftwarePassword.Text = SoftwarePassword
        txtSoftwareCode.Text = SoftwareCode
        cboDevEnviroment.Text = DevEnvironment.ToString.ToUpper
    End Sub

    Private Sub GetCRC32Encoded() 'added
        'get the path from txtActiveLockPath
        'get its crc32
        'encode the crc
        Dim raw As String
        Dim c As New CRC32
        Dim crc As Integer = 0
        Dim fileName As String = txtActivelockPath.Text
        If txtActivelockPath.Text.Length = 0 Then
            Err.Raise(1, "Private Sub GetCRC32Encoded", "Please Select The Activelock DLL first")
        End If
        If File.Exists(fileName) Then
            If cboDevEnviroment.Text = "VB6" Then 'added V1.0.4
                'This Is for VB6 ONLY
                Dim vb6crc As Long = CRCCheckSumTypeLib(fileName)
                CrcDataEnc = vb6crc - 121 'this will be added in the value() routine
            Else
                'This is for VB2005 and VB2008
                Dim f As FileStream = New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
                crc = c.GetCrc32(f)
                ' as the most commonly known format
                raw = String.Format("{0:X8}", crc)
                f.Close()
                CrcDataEnc = Enc(raw)
                System.Diagnostics.Debug.WriteLine("Hash: " & crc)
            End If
        End If
    End Sub

    Private Sub GetOwnCRC32Encoded() 'added V1.0.4
        'Use the dll in System32 folder
        'get its crc32
        'encode the crc
        Dim raw As String
        Dim c As New CRC32
        Dim crc As Integer = 0
        Dim fileName As String = Nothing
        If cboDevEnviroment.Text = "VB6" Then
            'This Is for VB6 ONLY
            fileName = System.Windows.Forms.Application.StartupPath & "\ActiveLock3.6.dll"
            Dim vb6crc As Long = CRCCheckSumTypeLib(fileName)
            CrcDataEnc = vb6crc - 121 'this will be added in the value() routine
            System.Diagnostics.Debug.WriteLine("Hash: " & vb6crc)
        Else
            'This is for VB2005 and VB2008
            fileName = System.Windows.Forms.Application.StartupPath & "\ActiveLock3_6Net.dll"
            Dim f As FileStream = New FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            crc = c.GetCrc32(f)
            ' as the most commonly known format
            raw = String.Format("{0:X8}", crc)
            f.Close()
            CrcDataEnc = Enc(raw)
            System.Diagnostics.Debug.WriteLine("Hash: " & crc)
        End If
    End Sub

#End Region '"General Form Routines"

#Region "Calculate VB6 Checksum" 'added V1.0.4
    Private Declare Function MapFileAndCheckSum Lib "imagehlp" Alias "MapFileAndCheckSumA" (ByVal Filename As String, ByRef HeaderSum As Long, ByRef CheckSum As Long) As Long

    Private Function CRCCheckSumTypeLib(ByVal filename As String) As Long
        Dim HeaderSum As Long
        Dim RealSum As Long
        MapFileAndCheckSum(filename, HeaderSum, RealSum)
        CRCCheckSumTypeLib = RealSum
    End Function

#End Region '"Calculate VB6 Checksum"

#Region "Help Provider"

    Private Sub ImageApplicationHelp1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp1.Click
        MsgBox("Provide a name for your application", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp2.Click
        MsgBox("Provide a version of the software", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp3.Click
        MsgBox("Provide a password for the application", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp4.Click
        MsgBox("Provide the Vcode (Public Key) from Alugen", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp5.Click
        MsgBox("Select which type of licence file must be generated," & vbCrLf & "Please note that short Key does not allow all the locks to be used", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp6.Click
        MsgBox("Select if the license file on the customer PC must be encrypted" & vbCrLf & "Please note even if the file is unencrypted it is not a security risk", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp7.Click
        MsgBox("Select if the application can auto register from a file in the application directory", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp8.Click
        MsgBox("Select where the license info is stored on the customer pc" & vbCrLf & "Please note only type file is curently supported", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp9.Click
        MsgBox("Select if you want the application to check on a Time server(On the internet)for clock tampering" & vbCrLf & "Please note that internet access will be needed for the application tho run", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp10.Click
        MsgBox("Select if you want the application to check system files for clock tampering", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp11.Click
        MsgBox("Select what type of trial the application should have", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp12.Click
        MsgBox("Select where the trial info should be hidden on the pc" & vbCrLf & "Please note that diffent options will afffect the speed of the startup", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp13.Click
        MsgBox("select how many times the application can run in the trial" & vbCrLf & " or how many days the application can run in the trial", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp14.Click
        MsgBox("Select the development environment you want to create the file for", MsgBoxStyle.Information)
    End Sub

    Private Sub ImageApplicationHelp15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp15.Click
        MsgBox("Select what the software must be locked to on the customer PC" & vbCrLf & "Please note that some IP addresses is dynamic and will change, so be carefull", MsgBoxStyle.Information)
    End Sub

    'added

    Private Sub ImageApplicationHelp16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImageApplicationHelp16.Click
        MsgBox("Just select the activelock dll that you are going to use." & vbCrLf & "The CRC will be Encoded by the Wizard for better protection")
    End Sub

#End Region '"Help Provider"

#Region "CRC provider" 'added

    Private Sub GroupBox14_Enter(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles GroupBox14.Enter

    End Sub

    'added

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        txtActivelockPath.Text = (OpenFileDialog1.FileName)
    End Sub

    'added

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        OpenFileDialog1.ShowDialog()
    End Sub

    'added

    Private Sub txtActivelockPath_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtActivelockPath.TextChanged

    End Sub

    'added

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            RadioButton1.Checked = False
            txtActivelockPath.Enabled = False
            btnDLLPath.Enabled = False
        Else
            RadioButton1.Checked = True
            txtActivelockPath.Enabled = True
            btnDLLPath.Enabled = True
        End If
        RadioButton1.Refresh()
    End Sub

    'added

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            RadioButton2.Checked = False
        Else
            RadioButton2.Checked = True
        End If
        RadioButton2.Refresh()
    End Sub

    'added

    Private Function Encrypt(ByRef strdata As String) As String
        Dim i, n As Integer
        Dim sResult As String = Nothing
        n = Len(strdata)
        For i = 1 To n
            sResult = sResult & Asc(Mid(strdata, i, 1)) * 7
        Next i
        Encrypt = sResult
    End Function

#End Region

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDLLPath.Click
        OpenFileDialog1.ShowDialog()
    End Sub

End Class
