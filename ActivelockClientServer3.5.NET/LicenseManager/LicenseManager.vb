Imports System.Windows.Forms
Imports ActiveLockServer
Imports System.Runtime.Remoting
'Imports System.Management

Public Class LicenseManager

    Private Sub LockManager_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '
        ' Local variables
        '
        Dim winSize As System.Drawing.Size
        Dim objInterface As ALServer
        '
        '
        '
        Try
            objInterface = Activator.GetObject(GetType(ALServer), LicenseData.LockInfo.URL)
        Catch ex As Exception
            MsgBox("License server could not be contacted")
            End
        End Try
        '
        ' If we do not have an acquired license then exit this application
        '
        Try
            If objInterface.Initialized() = False Then
                MsgBox("License Server is not running")
                End
            End If
            '
            ' Display the status of the license
            '
            lblLicenseStatus.Text = objInterface.LIcenseStatus
            '
            ' Determine whether we should show or hide the registration components
            '
            If objInterface.Acquired() = False Then
                GroupBox1.Visible = True
                '
                ' Set the Window Size so you can see the Registration controls
                '
                winSize = Me.Size
                winSize.Height = 519
                Me.Size = winSize
            Else
                GroupBox1.Visible = False
                '
                ' Set the Window Size so you don't see the Registration controls
                '
                winSize = Me.Size
                winSize.Height = 313
                Me.Size = winSize
            End If
        Catch ex As Exception
            MsgBox("License server could not be contacted")
            End
        End Try
        '
        ' Load the ListView data
        '
        LoadListView()
        '
        ' Done
        '
    End Sub

    Private Sub LoadListView()
        '
        ' Declare local variables
        '
        Dim Lock As LicenseData.LockInfo
        Dim lvItem As New ListViewItem
        Dim objInterface As ALServer
        Dim LicenseList As Licenses
        '
        ' Get the LicenseList Now
        '
        Try
            objInterface = Activator.GetObject(GetType(ALServer), LicenseData.LockInfo.URL)
            LicenseList = objInterface.LicenseList
        Catch ex As Exception
            MsgBox("License server could not be contacted")
            End
        End Try
        '
        ' Load the ListView with the lock data in the LicenseList collection
        '
        lvLicenses.Items.Clear()
        For Each Lock In LicenseList.Collection
            lvItem = New ListViewItem
            lvItem.Text = Lock.MacAddress
            lvItem.SubItems.Add(Lock.UserName)
            lvItem.SubItems.Add(Lock.MachineName)
            lvItem.SubItems.Add(Lock.IPAddress)
            lvItem.SubItems.Add(Lock.RequestDate)
            lvLicenses.Items.Add(lvItem)
        Next
        lvLicenses.View = View.Details
        '
        ' Done
        '
    End Sub

    Private Sub bntRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntRefresh.Click
        LoadListView()
    End Sub

    Private Sub bntExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntExit.Click
        Me.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub bntRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntRemove.Click
        '
        ' Declare local variables
        '
        Dim MacAddress As String
        Dim lvItem As ListViewItem
        Dim objInterface As ALServer
        Dim LicenseList As Licenses
        '
        ' Get the LicenseList Now
        '
        Try
            objInterface = Activator.GetObject(GetType(ALServer), LicenseData.LockInfo.URL)
            LicenseList = objInterface.LicenseList
        Catch ex As Exception
            MsgBox("License server could not be contacted")
            End
        End Try
        '
        ' Get the MacAddress of the licenses to remove, remove them from the collection
        '
        For Each lvItem In lvLicenses.SelectedItems
            MacAddress = lvItem.Text
            LicenseList.Remove(MacAddress)
        Next
        '
        ' Write the New License List back to the Server now
        '
        Try
            objInterface = Activator.GetObject(GetType(ALServer), LicenseData.LockInfo.URL)
            objInterface.LicenseList = LicenseList
        Catch ex As Exception
            MsgBox("License server could not be contacted")
            End
        End Try
        '
        ' Reload the ListView now that we have removed licenses
        '
        LoadListView()
        '
        ' Done
        '
    End Sub

    Private Sub bntRemoveAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntRemoveAll.Click
        '
        ' Local Variables
        '
        Dim objInterface As ALServer
        '
        ' Clear the LicenseList now
        '
        Try
            objInterface = Activator.GetObject(GetType(ALServer), LicenseData.LockInfo.URL)
            objInterface.LicenseList.Clear()
        Catch ex As Exception
            MsgBox("License server could not be contacted")
            End
        End Try
        LoadListView()
    End Sub

    Private Sub bntGenerate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntGenerate.Click
        '
        ' Local Variables
        '
        Dim objInterface As ALServer
        '
        ' If we have a user name, generate the InstallationCode now
        '
        If Not txtUser.Text Is Nothing Then
            If txtUser.Text <> "" Then
                Try
                    objInterface = Activator.GetObject(GetType(ALServer), LicenseData.LockInfo.URL)
                    txtRequestCode.Text = objInterface.InstallationCode(txtUser.Text)
                    Exit Sub
                Catch ex As Exception
                    MsgBox("License server could not be contacted")
                    End
                End Try
            End If
        End If
        '
        ' Provide message to the user that a user name is required
        '
        MsgBox("A user name must be supplied to generate a registration code")
        '
        ' Done
        '
    End Sub

    Private Sub bntRegister_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bntRegister.Click
        '
        ' Declare local variables
        '
        Dim strCode As String = ""
        Dim strMsg As String = ""
        Dim winSize As System.Drawing.Size
        Dim objInterface As ALServer
        '
        ' Register and aquire the lock
        '
        Try
            '
            ' Remove the line feeds from the license code now (I get them from ALUGEN 3.5) and register the license now
            '
            strCode = Replace(txtLicenseCode.Text, Chr(10), "")
            objInterface = Activator.GetObject(GetType(ALServer), LicenseData.LockInfo.URL)
            If objInterface.Register(strCode) = True Then
                '
                ' Hide the registration controls now
                '
                GroupBox1.Visible = False
                '
                ' Set the Window Size so you don't see the Registration controls
                '
                winSize = Me.Size
                winSize.Height = 313
                Me.Size = winSize
                '
                ' Display the status of the license
                '
                lblLicenseStatus.Text = ALServices.LicesneStatus
            Else
                MsgBox("Registration Failed, code may be incorrect")
            End If
        Catch ex1 As Exception
            '
            ' Registration failed
            '
            MsgBox("Registration Failed, code may be incorrect")
        End Try
        '
        ' Done
        '
    End Sub

End Class
