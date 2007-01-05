Imports System.IO
Imports Microsoft.Win32
Imports LicenseData

Public Class ALServer
    Inherits MarshalByRefObject
    Implements iLicense
    '
    ' Public Interface Methods for Client Applications
    '
    Public Function Acquire(ByVal RequestData As LockInfo) As Boolean Implements iLicense.Acquire
        '
        ' This function acquires a license from the server
        '
        Try
            '
            ' If the LicenseManager has not acquired a valid ActiveLock license, reject the request
            '
            If ALServices.Acquired = False Then
                ALServices.WriteError("LicenseManager is not intialized. License request has been rejected", ALServices.iError)
                Return False
            End If
            '
            ' If there are more licenses than allowed reject this request, even if it's a duplicate
            '
            If OverLimit() Then
                Return False
            End If
            '
            ' If the license requested is already on the list accept the new request
            '
            If Not (ALServices.LicenseList.LicenseInfo(RequestData.MacAddress) Is Nothing) Then
                ALServices.LicenseList.LicenseInfo(RequestData.MacAddress) = RequestData    'Update the License information now
                ALServices.WriteError("License Request Granted", ALServices.iInfo)
                Return True
            End If
            '
            ' If the current count of active licenses is less than the ActiveLock MaxCount, add the lock and accept the license
            '
            If ALServices.LicenseList.Count < ALServices.LicenseLimit Then
                ALServices.LicenseList.Add(RequestData, RequestData.MacAddress)
                Return True
            End If
            '
            ' Return False (reject the license request)
            '
            ALServices.WriteError("Maximum concurrent license limit has been reached.  License Request Rejected", ALServices.iWarn)
            Return False
        Catch ex As Exception
            ALServices.WriteError("Exception was thrown requesting a License.  Exception Message: " & ex.Message, ALServices.iError)
            Return False
        End Try
        '
        ' Done
        '
    End Function

    Public Sub Release(ByVal RequestData As LockInfo) Implements iLicense.Release
        '
        ' Remove the License from the list now
        '
        ALServices.WriteError("License Released", ALServices.iInfo)
        ALServices.LicenseList.Remove(RequestData.MacAddress)
        '
        ' Done
        '
    End Sub

    Public Function Validate(ByVal RequestData As LockInfo) As String Implements iLicense.Validate
        '
        ' Return the ValidationCode so that the client can authenticate this is a valid Server
        '
        Return ALServices.ValidationCode
        '
        ' Done
        '
    End Function

    Public Function Verify(ByVal RequestData As LockInfo) As Boolean Implements iLicense.Verify
        '
        ' If the Lock is on the list, return true (still active) else return false (removed by license administrator)
        '
        Try
            If Not (ALServices.LicenseList.LicenseInfo(RequestData.MacAddress) Is Nothing) Then
                If OverLimit() Then
                    Return False
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            ALServices.WriteError("License is no longer valid", ALServices.iWarn)
            Return False
        End Try
        '
        ' Done
        '
    End Function
    '
    ' Public Methods and Properties for License Manager Application
    ' These are not part of the iLicense Interface so the consumer of these has to be located with the ActiveLockServer.exe
    '
    Public ReadOnly Property Initialized() As Boolean
        Get
            Return ALServices.Initialized
        End Get
    End Property

    Public ReadOnly Property Acquired() As Boolean
        Get
            Return ALServices.Acquired
        End Get
    End Property

    Public Property LicenseList() As Licenses
        Get
            Return ALServices.LicenseList
        End Get
        Set(ByVal value As Licenses)
            ALServices.LicenseList = value
        End Set
    End Property

    Public ReadOnly Property LIcenseStatus()
        Get
            Return ALServices.LicesneStatus
        End Get
    End Property

    Public Function Initialize() As Boolean
        ALServices.Initialize()
        Return ALServices.Initialized
    End Function

    Public Function InstallationCode(ByVal UserName As String) As String
        Return ALServices.InstallationCode(UserName)
    End Function

    Public Function Register(ByVal strCode As String) As Boolean
        ALServices.Register(strCode)
        ALServices.Initialize()
        Return ALServices.Acquired
    End Function
    '
    ' Private Methods
    '
    Private Function OverLimit() As Boolean
        '
        ' If there are more licenses than allowed reject this request, even if it's a duplicate
        '
        If ALServices.LicenseList.Count > ALServices.LicenseLimit Then
            ALServices.WriteError("More licenses granted than allowed.", ALServices.iError)
            Return True
        End If
        Return False
        '
        ' Done
        '
    End Function
    '
    ' Done
    '
End Class
