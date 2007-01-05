Imports System.Diagnostics

Public Class Logging
    '
    ' Declare local constants
    '
    Public Const strSource As String = "ShopTrack"
    Public Shared ReadOnly iInfo As Integer = 0
    Public Shared ReadOnly iWarn As Integer = 1
    Public Shared ReadOnly iError As Integer = 2
    '
    ' Declare class variables
    '
    Dim strMessage As String = ""
    Dim intLevel As Integer = 0
    '
    ' Public Properties
    '
    Public Property Message() As String
        Get
            Message = strMessage
        End Get
        Set(ByVal Value As String)
            strMessage = Value
        End Set
    End Property

    Public Property Level() As Integer
        Get
            Level = intLevel
        End Get
        Set(ByVal Value As Integer)
            intLevel = Value
            If intLevel < 0 Or intLevel > 2 Then
                intLevel = 2
            End If
        End Set
    End Property
    '
    ' Public Methods
    '
    Public Sub New(ByVal Message As String, ByVal Level As Integer)
        '
        ' Create a new instance of this object
        '
        MyBase.New()
        '
        ' Set the class variables now
        '
        strMessage = Message
        intLevel = Level
        '
        ' Done
        '
    End Sub

    Public Sub Log()
        '
        ' Call this method with a new thread to write the log, so we can have multiple threads working at once
        '
        Select Case intLevel
            Case iInfo
                WriteInformation()
            Case iWarn
                WriteWarning()
            Case iError
                WriteError()
            Case Else
                WriteError()
        End Select
        '
        ' Done
        '
    End Sub
    '
    ' Private Methods
    '
    Private Sub WriteInformation()

        EventLog.WriteEntry(strSource, strMessage, EventLogEntryType.Information)

    End Sub

    Private Sub WriteWarning()

        EventLog.WriteEntry(strSource, strMessage, EventLogEntryType.Warning)

    End Sub

    Private Sub WriteError()

        EventLog.WriteEntry(strSource, strMessage, EventLogEntryType.Error)

    End Sub
    '
    ' Done
    '
End Class
