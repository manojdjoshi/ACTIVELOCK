Imports System.IO
Imports Microsoft.Win32
Imports LicenseData

Public Class Server
    Inherits MarshalByRefObject
    Implements iLicense
    '
    ' Public Interface Methods
    '
    Public Function Acquire(ByVal RequestData As LockInfo) As Boolean Implements iLicense.Acquire
        '
        ' This function acquires a license from the server
        '
        Try
            '
            ' If the LicenseManager has not acquired a valid ActiveLock license, reject the request
            '
            If LicenseServer.Acquired = False Then
                LicenseServer.WriteError("LicenseManager is not intialized. License request has been rejected", LicenseServer.iError)
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
            If Not (LicenseServer.LicenseList.LicenseInfo(RequestData.MacAddress) Is Nothing) Then
                LicenseServer.LicenseList.LicenseInfo(RequestData.MacAddress) = RequestData    'Update the License information now
                LicenseServer.WriteError("License Request Granted", LicenseServer.iInfo)
                Return True
            End If
            '
            ' If the current count of active licenses is less than the ActiveLock MaxCount, add the lock and accept the license
            '
            If LicenseServer.LicenseList.Count < LicenseServer.LicenseLimit Then
                LicenseServer.LicenseList.Add(RequestData, RequestData.MacAddress)
                Return True
            End If
            '
            ' Return False (reject the license request)
            '
            LicenseServer.WriteError("Maximum concurrent license limit has been reached.  License Request Rejected", LicenseServer.iWarn)
            Return False
        Catch ex As Exception
            LicenseServer.WriteError("Exception was thrown requesting a License.  Exception Message: " & ex.Message, LicenseServer.iError)
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
        LicenseServer.WriteError("License Released", LicenseServer.iInfo)
        LicenseServer.LicenseList.Remove(RequestData.MacAddress)
        '
        ' Done
        '
    End Sub

    Public Function Validate(ByVal RequestData As LockInfo) As String Implements iLicense.Validate
        '
        ' Return the ValidationCode so that the client can authenticate this is a valid Server
        '
        Return LicenseServer.ValidationCode
        '
        ' Done
        '
    End Function

    Public Function Verify(ByVal RequestData As LockInfo) As Boolean Implements iLicense.Verify
        '
        ' If the Lock is on the list, return true (still active) else return false (removed by license administrator)
        '
        Try
            If Not (LicenseServer.LicenseList.LicenseInfo(RequestData.MacAddress) Is Nothing) Then
                If OverLimit() Then
                    Return False
                End If
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            LicenseServer.WriteError("License is no longer valid", LicenseServer.iWarn)
            Return False
        End Try
        '
        ' Done
        '
    End Function

    Private Function OverLimit() As Boolean
        '
        ' If there are more licenses than allowed reject this request, even if it's a duplicate
        '
        If LicenseServer.LicenseList.Count > LicenseServer.LicenseLimit Then
            LicenseServer.WriteError("More licenses granted than allowed.", LicenseServer.iError)
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
