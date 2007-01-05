Imports ActiveLock3_5NET
Imports System.IO
Imports Microsoft.Win32
Imports LicenseData
Imports System.Threading

Public Class ALServices
    '
    ' Class Constants
    '
    Friend Const ValidationCode As String = "Cool"
    Public Const iInfo As Integer = 0
    Public Const iWarn As Integer = 1
    Public Const iError As Integer = 2
    '
    ' Class Variables
    '
    Private Shared _LicenseList As New Licenses
    Friend Shared _Initialized As Boolean = False
    Friend Shared _Acquired As Boolean = False
    Private Shared _LicenseLimit As Integer = 0
    Private Shared _LicenseStatus As String = "No License"
    Private Shared ActiveLock As ActiveLock3_5NET._IActiveLock
    Private Shared ALGlobals As New ActiveLock3_5NET.Globals_Renamed
    Private Shared WithEvents ActiveLockEventSink As ActiveLock3_5NET.ActiveLockEventNotifier
    '
    ' Public Properties
    '
    Public Shared Property LicenseList() As Licenses
        Get
            Return _LicenseList
        End Get
        Set(ByVal value As Licenses)
            _LicenseList = value
        End Set
    End Property

    Public Shared ReadOnly Property LicenseLimit()
        Get
            Return _LicenseLimit
        End Get
    End Property

    Public Shared ReadOnly Property LicesneStatus() As String
        Get
            Return _LicenseStatus
        End Get
    End Property

    Public Shared ReadOnly Property Acquired() As Boolean
        Get
            Return _Acquired
        End Get
    End Property

    Public Shared ReadOnly Property Initialized() As Boolean
        Get
            Return _Initialized
        End Get
    End Property

    Public Shared ReadOnly Property InstallationCode(Optional ByVal User As String = "") As String
        Get
            If Initialized Then
                Return ActiveLock.InstallationCode(User)
            Else
                Return ""
            End If
        End Get
    End Property
    '
    ' Public Methods
    '
    Public Shared Sub WriteError(ByVal Message As String, ByVal Level As Integer)
        '
        ' Local variables
        '
        Dim WriteLog As New Logging(Message, Level)
        Dim WriteThread As Thread
        Dim WriteThreadStart As New ThreadStart(AddressOf WriteLog.Log)
        '
        ' Spawn a new thread to write to the Event viewer now
        '
        WriteThread = New Thread(WriteThreadStart)
        WriteThread.IsBackground = True
        WriteThread.Name = "WriteLog"
        WriteThread.Start()
        '
        ' Done
        '
    End Sub

    Public Shared Sub Register(ByVal RegistrationCode As String)
        If Initialized Then
            ActiveLock.Register(RegistrationCode)
        End If
    End Sub

    Public Shared Sub Initialize()
        '
        ' Initialize the LicenseList Object to track all Active Licenses
        '
        _LicenseList.Clear()
        '
        ' Initialize ActiveLock and Acquire the license now
        '
        If _Acquired = False Then
            _Acquired = AcquireLock()
        End If
        '
        ' Done
        '
    End Sub
    '
    ' Private Methods
    '
    Private Shared Function AcquireLock() As Boolean
        '
        ' Declare local variables
        '
        Dim strMsg As String = ""
        '
        ' Get a new instance of ActiveLock and set the EventNotifier so that we can receive notifications
        '
        ActiveLock = ALGlobals.NewInstance()
        ActiveLockEventSink = ActiveLock.EventNotifier
        '
        ' Specify License parameters
        '
        ActiveLock.KeyStoreType = ActiveLock3_5NET.IActiveLock.LicStoreType.alsFile
        ActiveLock.KeyStorePath = System.Windows.Forms.Application.StartupPath & "\License.lic"
        ActiveLock.AutoRegisterKeyPath = System.Windows.Forms.Application.StartupPath & "\License.all"
        ActiveLock.LicenseFileType = ActiveLock3_5NET.IActiveLock.ALLicenseFileTypes.alsLicenseFileEncrypted
        '
        ' Specify Software parameters
        '
        ActiveLock.LockType = ActiveLock3_5NET.IActiveLock.ALLockTypes.lockWindows Or ActiveLock3_5NET.IActiveLock.ALLockTypes.lockMAC
        ActiveLock.SoftwareName = "LicenseServer"
        ActiveLock.SoftwareVersion = "1.0"
        ActiveLock.SoftwarePassword = "Cool"
        ActiveLock.SoftwareCode = "AAAAB3NzaC1yc2EAAAABJQAAAICj3T75VHLRAe/Bx9mU1gqwSKPwx86eB1Ah1TTopUvz7vRs27l3j37dtrz9Vr3eDPRBE96byIyEGgZFoW3SjGvX0+YI5IdPD8OaCBeE66eKghM9KvpCPC0fqTh+eyAC7mkU0+AXnl7JbpidTtMhPwgzmQotY07wmgK0oWWCDOcgLw=="
        '
        ' Set Trial information
        '
        ActiveLock.TrialType = IActiveLock.ALTrialTypes.trialNone
        '
        ' Initialize ActiveLock now
        '
        Try
            ActiveLock.Init()
            _Initialized = True
        Catch ALInit As Exception
            ALServices.WriteError("Exception thrown intializing lock.  Exception message: " & ALInit.Message, ALServices.iError)
            Return False
        End Try
        '
        ' Acquire a lock.  If it's a trail handle it now
        '
        Try
            ActiveLock.Acquire(strMsg)
            If strMsg = "" Then
                _LicenseStatus = "Licensed for " & ActiveLock.MaxCount & " concurrent users"
                _LicenseLimit = ActiveLock.MaxCount
            End If
            Return True
        Catch ALAcquire As Exception
            ALServices.WriteError("Exception thrown acquiring lock.  Exception message: " & ALAcquire.Message, ALServices.iError)
            Return False
        End Try
        '
        ' If we didn't get a valid license for some reason, return false now
        '
        Return False
        '
        ' Done
        '    
    End Function
    '
    ' Done
    '
End Class
