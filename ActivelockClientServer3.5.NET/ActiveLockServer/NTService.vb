Imports System
Imports System.ServiceProcess
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Remoting.Channels.Tcp
Imports System.Runtime.Remoting.Channels.Http
Imports LicenseData
Imports ActiveLockServer.ALServices
Imports ActiveLock3_5NET
Imports System.IO
Imports Microsoft.Win32
Imports System.Threading

Public Class NTService

    Protected Overrides Sub OnStart(ByVal args() As String)
        '
        ' Declare local variables
        '
        Dim tcpChannel As TcpChannel
        Dim httpChannel As HttpChannel
        '
        ' Initialize ActiveLock and Acquire the license now
        '
        ALServices.Initialize()
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
        RemotingConfiguration.RegisterWellKnownServiceType(GetType(ALServer), LockInfo.URI, WellKnownObjectMode.SingleCall)
        '
        ' Done
        '
    End Sub

    Protected Overrides Sub OnStop()

    End Sub

End Class
