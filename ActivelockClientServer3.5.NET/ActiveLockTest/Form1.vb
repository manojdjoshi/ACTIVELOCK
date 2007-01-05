Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Runtime.Remoting.Channels.Http
Imports LicenseData

Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '
        ' Declare local variables
        '
        Dim tcpChannel As TcpChannel
        Dim httpChannel As HttpChannel
        'Dim objLogging As Logging
        '
        ' Initialize ActiveLock and Acquire the license now
        '
        LicenseServer.Initialize()
        '
        ' Register the channel now
        '
        If LCase(LockInfo.Channel) = "http" Then
            httpChannel = New HttpChannel(LockInfo.Port)
            ChannelServices.RegisterChannel(httpChannel, False)
        Else
            tcpChannel = New TcpChannel(LockInfo.Port)
            ChannelServices.RegisterChannel(tcpChannel, False)
        End If
        '
        ' Register the objects that we will be serving on this channel
        '
        RemotingConfiguration.RegisterWellKnownServiceType(GetType(Server), LockInfo.URI, WellKnownObjectMode.SingleCall)
        '
        ' Done
        '
    End Sub

    Private Sub btnManage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnManage.Click
        '
        '
        '
        Dim Manager As New LicenseManager
        '
        '
        '
        Manager.ShowDialog()
    End Sub
End Class
