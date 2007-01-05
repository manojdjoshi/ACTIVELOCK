Imports System.Threading
Imports System.Runtime.Remoting
Imports System.Management

Public Class Client
    '
    ' Global variables
    '
    Dim _RequestData As New LicenseData.LockInfo
    Dim _Licensed As Boolean = False
    Dim TThread As Thread
    Dim TThreadStart As New ThreadStart(AddressOf Me.LicenseTimer)
    '
    ' Private Methods
    '
    Private Sub Client_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '
        ' Populate the License request data that we will use request a license from the server
        '
        PopulateRequestData()
        '
        ' If we don't get a License exit the client application
        '
        If AcquireLock() = False Then
            End
        End If
        '
        ' Done
        '
    End Sub

    Private Sub PopulateRequestData()
        '
        ' Local Variables
        '
        Dim NACClass As New System.Management.ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim NACCollection As System.Management.ManagementObjectCollection = NACClass.GetInstances()
        Dim NACObject As System.Management.ManagementObject
        '
        ' Get the MAC Address and IP Address of the Client now
        '
        For Each NACObject In NACCollection
            If NACObject.Item("IPEnabled") = True Then
                _RequestData.MacAddress = NACObject.Item("MacAddress").ToString()
                _RequestData.IPAddress = NACObject.Item("IPAddress")(0).ToString()
                Exit For
            End If
        Next
        '
        ' Get the User Name, Machine Name, and Request Date now
        '
        _RequestData.MachineName = System.Windows.Forms.SystemInformation.ComputerName
        _RequestData.UserName = System.Windows.Forms.SystemInformation.UserName
        _RequestData.RequestDate = Now
        '
        ' Done
        '
    End Sub

    Private Function AcquireLock() As Boolean
        '
        ' Declare local variables
        '
        Dim strValidationCode As String
        Dim boolSuccess As String
        Dim objInterface As LicenseData.iLicense
        '
        ' Instantiate the LicenseServer object on the server now and validate it's authenticity
        '
        Try
            objInterface = Activator.GetObject(GetType(LicenseData.iLicense), LicenseData.LockInfo.URL)
            strValidationCode = objInterface.Validate(_RequestData)
            If strValidationCode <> "Cool" Then
                MsgBox("License server invalid")
                Return False
            End If
        Catch ex As Exception
            MsgBox("License server could not be contacted")
            Return False
        End Try
        '
        ' Get a license from the License Server
        '
        Try
            objInterface = Activator.GetObject(GetType(LicenseData.iLicense), LicenseData.LockInfo.URL)
            boolSuccess = objInterface.Acquire(_RequestData)
            If boolSuccess Then
                _Licensed = True
            Else
                MsgBox("License request has been rejected by the server")
                Return False
            End If
        Catch ex As Exception
            MsgBox("License server could not be contacted")
            Return False
        End Try
        '
        ' Start the License Timer to periodically check to see if the license is still valid
        '
        If _Licensed = True Then
            Try
                TThread = New Thread(TThreadStart)
                TThread.IsBackground = True
                TThread.Name = "LicenseTimer"
                TThread.Start()
            Catch ex As Exception
                MsgBox("Could not start License Timer Thread")
                Return False
            End Try
        End If
        '
        ' Return Successful
        '
        Return True
        '
        ' Done
        '
    End Function

    Private Sub LicenseTimer()
        '
        ' Declare local variables
        '
        Dim objInterface As LicenseData.iLicense
        '
        ' Get a license from the License Server
        '
        Try
            objInterface = Activator.GetObject(GetType(LicenseData.iLicense), LicenseData.LockInfo.URL)
            '
            ' Check the status of the license to make sure it's still valid
            '
            While objInterface.Verify(_RequestData)
                Thread.Sleep(5000)    'Sleep 5 seconds and check license status again
            End While
            '
            ' License is not valid, exit the client application
            '
            MsgBox("License is no longer valid")
            End
        Catch ex As Exception
            MsgBox("License server could not be contacted")
            End
        End Try
        '
        ' Done
        '
    End Sub

    Private Sub btnRelease_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRelease.Click
        '
        ' Declare local variables
        '
        Dim objInterface As LicenseData.iLicense
        '
        ' Release the license now
        '
        Try
            objInterface = Activator.GetObject(GetType(LicenseData.iLicense), LicenseData.LockInfo.URL)
            objInterface.Release(_RequestData)
        Catch ex As Exception
            MsgBox("License server could not be contacted")
        End Try
        End
    End Sub
    '
    ' Done
    '
End Class
