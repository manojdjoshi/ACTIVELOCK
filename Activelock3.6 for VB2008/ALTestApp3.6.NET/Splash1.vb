Public Class frmSplash1

    Private Sub frmSplash1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        lblVersion.Text = "Version " & Application.ProductVersion
        lblProductName.Text = Application.ProductName
        If strMsg <> "" Then
            ProgressBar1.Value = CInt((1 - remainingDays / totalDays) * 100)
            If remainingDays > 0 Then lblInfo.Text = CStr(totalDays - remainingDays) & " day(s) used out of " & CStr(totalDays) & " trial days"
        End If

    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        End
    End Sub

    Private Sub cmdRefgister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRefgister.Click
        Me.Close()
    End Sub

    Private Sub cmdTry_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdTry.Click
        Me.Close()
    End Sub
End Class