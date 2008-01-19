Option Strict Off
Option Explicit On
Friend Class frmMain
Inherits System.Windows.Forms.Form
#Region "Form Controls"

    Private Sub frmMain_Load(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles MyBase.Load
        InitMyProgram()
        Me.Text = ActivelockValues.AppName & " - " & ActivelockValues.AppVersion
    End Sub

#End Region '"Form Controls"

#Region "Form Buttons"

Private Sub cmdCopy_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdCopy.Click
    If Len(txtReqCodeGen.Text) = 0 Then
        txtReqCodeGen.Text = GetTheInstallationCode(txtUsername.Text)
    End If
    My.Computer.Clipboard.Clear()
    My.Computer.Clipboard.SetText(txtReqCodeGen.Text)
End Sub

Private Sub cmdGenerateCode_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdGenerateCode.Click
    txtReqCodeGen.Text = GetTheInstallationCode(txtUsername.Text)
End Sub

Private Sub cmdKillLic_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdKillLic.Click
    KillTheLic()
    InitMyProgram()
End Sub

Private Sub cmdKillTrial_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKillTrial.Click
    KillTheTrial()
    InitMyProgram()
End Sub

Private Sub cmdPaste_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdPaste.Click
    If My.Computer.Clipboard.GetText = txtReqCodeGen.Text Then
        MsgBox("You cannot paste the Installation Code into the Liberation Key field.", MsgBoxStyle.Exclamation)
        Exit Sub
    End If
    txtLibKeyIn.Text = My.Computer.Clipboard.GetText
End Sub

Private Sub cmdResetTrial_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdResetTrial.Click
    ResetTheTrial()
    InitMyProgram()
End Sub

Private Sub cmdRegister_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRegister.Click
    Dim GoodRegister As Boolean
    GoodRegister = RegisterTheApplication(txtLibKeyIn.Text, txtUsername.Text)
    If GoodRegister = True Then
        MsgBox("Registration successful!")
        InitMyProgram()
    Else
        MsgBox("Error Registering Your Application")
    End If
End Sub

#End Region '"Form Buttons"

#Region "My Routines"

Public Sub InitMyProgram()
    Dim CanRun As Boolean
    CanRun = InitActivelock()
    If CanRun = True Then
        btnControl1.Enabled = True
        If Strings.StrComp(ActivelockValues.RegisteredLevel, "Trial", CompareMethod.Binary) <> 0 Then
            btnControl2.Enabled = True
            btnControl3.Enabled = False
        End If
        If Strings.StrComp(ActivelockValues.RegisteredLevel, "Full Version-USA", CompareMethod.Binary) = 0 Then
            btnControl2.Enabled = True
            btnControl3.Enabled = True
        End If
    Else
        btnControl1.Enabled = False
        btnControl2.Enabled = False
        btnControl3.Enabled = False
        'Trial Or License Expired or Tampered
        'Disable everything and ask to register
    End If
    Text1.Text = ActivelockValues.RegStatus
    Text2.Text = ActivelockValues.UsedDaysOrRuns
    Text3.Text = CStr(ActivelockValues.ValidTrial)
    Text4.Text = ActivelockValues.LicenceType
    Text5.Text = ActivelockValues.ExpirationDate
    Text6.Text = ActivelockValues.RegisteredUser
    Text7.Text = ActivelockValues.RegisteredLevel
    Text8.Text = ActivelockValues.LicenseClass
    Text9.Text = CStr(ActivelockValues.MaxCount)
End Sub

#End Region '"My Routines"

End Class