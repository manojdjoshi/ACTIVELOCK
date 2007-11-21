Module modGlobal
#Region "enumerations"
    Public Enum LicenceKeyType_t
        alsRSA = 1
        alsShortKeyMD5 = 2
    End Enum

    Public Enum TrialType_t
        trialNone = 0
        trialDays = 1
        trialRuns = 2
    End Enum

    Public Enum KeyStoreType_t
        alsFile = 0
        alsRegistry = 1
    End Enum
    Public Enum AutoRegister_t
        alsDisableAutoRegistration = 0
        alsEnableAutoRegistration = 1
    End Enum
    Public Enum LicenceFileType_t
        alsLicenseFilePlain = 0
        alsLicenseFileEncrypted = 1
    End Enum
    Public Enum TimeServerClockTampering_t
        alsDontCheckTimeServer = 0
        alsCheckTimeServer = 1
    End Enum

    Public Enum SystemFilesClockTampering_t
        alsDontCheckSystemFiles = 0
        alsCheckSystemFiles = 1
    End Enum

    Public Enum DevEnvironment_t
        vb6 = 0
        vb2003 = 1
        vb2005 = 2
    End Enum
#End Region '"enumerations"

#Region "LocalVariables"
    Public SoftwareName As String = Nothing
    Public SoftwareVersion As String = Nothing
    Public SoftwarePassword As String = Nothing
    Public SoftwareCode As String = Nothing
    Public LicenceKeyType As LicenceKeyType_t
    Public LicenceFileType As LicenceFileType_t
    Public LockTypes As New ArrayList
    Public AutoRegister As AutoRegister_t
    Public TimeServerClockTampering As TimeServerClockTampering_t
    Public SystemFilesClockTampering As SystemFilesClockTampering_t
    Public TrialType As TrialType_t
    Public TrialHideType As New ArrayList
    Public TrialLength As Integer = 0
    Public KeyStoreType As KeyStoreType_t
    Public DevEnvironment As DevEnvironment_t
#End Region '"LocalVariables"
    Sub main(ByVal Args() As String)
        Dim i As Integer = 0
        Dim j As Integer = 0
        For Each value As String In Args

            If value.Contains("AppName=") Then
                i = value.Length
                j = value.IndexOf("=") + 1
                SoftwareName = Strings.Right(value, i - j)
            End If
            If value.Contains("AppVersion=") Then
                i = value.Length
                j = value.IndexOf("=") + 1
                SoftwareVersion = Strings.Right(value, i - j)
            End If
            If value.Contains("PUB_KEY=") Then
                i = value.Length
                j = value.IndexOf("=") + 1
                SoftwareCode = Strings.Right(value, i - j)
            End If

        Next
        frmMain.ShowDialog()
    End Sub
    Public Function EncodedPassword(ByVal Password As String) As String
        Dim encPassword As String = Nothing
        For Each c As Char In Password
            encPassword = encPassword & "Chr(" & Asc(c) & ") & "
        Next
        encPassword = Strings.Left(encPassword, (encPassword.Length - 3))
        Return encPassword
    End Function
    Public Function GetHideTypeString(ByVal HideType As ArrayList) As String
        Dim strHideType As String = Nothing
        For Each item As String In HideType
            strHideType = strHideType & "ActiveLock3_5NET.IActiveLock.ALTrialHideTypes." & item.ToString & " Or "
        Next
        strHideType = Strings.Left(strHideType, (strHideType.Length - 4))
        Return strHideType
    End Function

    Public Function GetLockTypesString(ByVal LockType As ArrayList) As String
        Dim stringLockType As String = Nothing
        For Each item As String In LockType
            stringLockType = stringLockType & "IActiveLock.ALLockTypes." & item.ToString & " Or "
        Next
        stringLockType = Strings.Left(stringLockType, (stringLockType.Length - 4))
        Return stringLockType
    End Function

    Public Function Enc(ByRef strdata As String) As String
        Dim i, n As Integer
        Dim sResult As String = Nothing
        n = Len(strdata)
        Dim l As Integer
        For i = 1 To n
            l = Asc(Mid(strdata, i, 1)) * 11
            If sResult = "" Then
                sResult = Hex(l)
            Else
                sResult = sResult & "." & Hex(l)
            End If
        Next i
        Enc = sResult
    End Function

End Module
'"AppName=test_app" "AppVersion=1.0.0" "PUB_KEY=123456789"
