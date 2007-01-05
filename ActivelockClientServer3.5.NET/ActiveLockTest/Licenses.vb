Imports System.Collections.Specialized
Imports LicenseData

Public Class Licenses
    Inherits NameObjectCollectionBase
    '
    ' Public Properties
    '
    Public Property LicenseInfo(ByVal MacAddress As String) As LockInfo
        Get
            Return Me.BaseGet(MacAddress)
        End Get
        Set(ByVal value As LockInfo)
            Me.BaseSet(MacAddress, value)
        End Set
    End Property

    Public ReadOnly Property Collection() As Array
        Get
            Return Me.BaseGetAllValues
        End Get
    End Property
    '
    ' Public Methods
    '
    Public Sub New()
        '
        ' Create a new instance of this object class, and initialize class variables
        '
        MyBase.New()
        '
        ' Done
        '
    End Sub

    Public Sub Add(ByVal LockData As LockInfo, ByVal MacAddress As String)
        '
        ' Adds a Lock object to the Locks collection
        '
        Try
            Me.BaseAdd(MacAddress, LockData)
        Catch ex As Exception
            Throw ex
        End Try
        '
        ' Done
        '
    End Sub

    Public Sub Remove(ByVal MacAddress As String)
        '
        ' Removes a Lock object to the Locks collection
        '
        Try
            Me.BaseRemove(MacAddress)
        Catch ex As Exception
            Throw ex
        End Try
        '
        ' Done
        '
    End Sub

    Public Sub Clear()
        '
        ' Clears the Locks collection of all Lock objects
        '
        Try
            Me.BaseClear()
        Catch ex As Exception
            Throw ex
        End Try
        '
        ' Done
        '
    End Sub
    '
    ' Done
    '
End Class
