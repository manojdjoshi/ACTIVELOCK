Public Class frmSplash1

    Private Sub frmSplash1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblVersion.Text = "Version " & Application.ProductVersion
        lblProductName.Text = Application.ProductName
        If strMsg.Contains("Trial Period") And totalDays <> 0 Then
            ProgressBar1.Value = CInt(Math.Abs(1 - remainingDays / totalDays) * 100)
            If remainingDays > 0 Then lblInfo.Text = CStr(totalDays - remainingDays) & " day(s) used out of " & CStr(totalDays) & " trial days"
        ElseIf strMsg.Contains("Trial Runs") And totalRuns <> 0 Then
            ProgressBar1.Value = CInt(Math.Abs(1 - remainingRuns / totalRuns) * 100)
            If remainingRuns > 0 Then lblInfo.Text = CStr(totalRuns - remainingRuns) & " run(s) used out of " & CStr(totalRuns) & " trial days"
        End If

    End Sub
    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        Application.Exit()
    End Sub

    Private Sub cmdRegister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRegister.Click
        MsgBox("Show your registration form here after running Activelock with no Trial or after KillTrial.", vbInformation)
        useTrial = False
        Dim frmMain As New frmMain
        frmMain.Visible = True
        Me.Hide()
    End Sub
    Private Sub cmdTry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTry.Click
        MsgBox("Show your main form here after running Activelock with the Trial feature on.", vbInformation)
        Me.Close()
        useTrial = True
        Dim frmMain As New frmMain
        frmMain.Visible = True
        Application.DoEvents()
        frmMain.Hide()
         'Show some other form of your app here
    End Sub
End Class