<SerializableAttribute()> Public Class LockInfo

    '
    ' Declare Class Variables
    '
    Dim _UserName As String
    Dim _MachineName As String
    Dim _IPAddress As String
    Dim _MacAddress As String
    Dim _RequestDate As Date
    '
    ' Declare Public Constants used for setting up Remoting
    '
    Public Const Channel As String = "tcp"
    Public Const Port As Integer = 2020
    Public Const URI As String = "ActiveLock.rem"
    Public Const URL As String = "tcp://127.0.0.1:2020/ActiveLock.rem"
    '
    ' Declare Public Properties
    '
    Public Property UserName() As String
        Get
            Return _UserName
        End Get
        Set(ByVal value As String)
            _UserName = value
        End Set
    End Property

    Public Property MachineName() As String
        Get
            Return _MachineName
        End Get
        Set(ByVal value As String)
            _MachineName = value
        End Set
    End Property

    Public Property IPAddress() As String
        Get
            Return _IPAddress
        End Get
        Set(ByVal value As String)
            _IPAddress = value
        End Set
    End Property

    Public Property MacAddress() As String
        Get
            Return _MacAddress
        End Get
        Set(ByVal value As String)
            _MacAddress = value
        End Set
    End Property

    Public Property RequestDate()
        Get
            Return _RequestDate
        End Get
        Set(ByVal value)
            _RequestDate = value
        End Set
    End Property
    '
    '
    '
    Public Sub New()
        MyBase.New()
        'Do nothing but is required for deserialization when remoting a collection of these objects
    End Sub
    '
    ' Done
    '
End Class
