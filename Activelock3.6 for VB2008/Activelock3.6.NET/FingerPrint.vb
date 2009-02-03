Imports System
Public Class FingerPrint

    Private m_CpuID As String = ""
    Private m_BiosID As String = ""
    Private m_DiskID As String = ""
    Private m_BaseID As String = ""
    Private m_VideoID As String = ""
    Private m_MacID As String = ""

    Private m_UseCpuID As Boolean = True
    Private m_UseBiosID As Boolean = True
    Private m_UseDiskID As Boolean = True
    Private m_UseBaseID As Boolean = True
    Private m_UseVideoID As Boolean = True
    Private m_UseMacID As Boolean = True

    Private m_ReturnLength As Long = 8
    Private m_TotalLength As Long = 8

    Public Event StartingWith(ByVal Text As String)
    Public Event DoneWith(ByVal Text As String)

    Public ReadOnly Property TotalLength() As Long
        Get
            Return m_TotalLength
        End Get
    End Property

    Public Property ReturnLength() As Long
        Get
            Return m_ReturnLength
        End Get
        Set(ByVal value As Long)
            If m_ReturnLength < 0 Then m_ReturnLength = 0
            m_ReturnLength = value
        End Set
    End Property

    Public Property UseCpuID() As Boolean
        Get
            Return m_UseCpuID
        End Get
        Set(ByVal value As Boolean)
            m_UseCpuID = value
        End Set
    End Property

    Public Property UseBiosID() As Boolean
        Get
            Return m_UseBiosID
        End Get
        Set(ByVal value As Boolean)
            m_UseBiosID = value
        End Set
    End Property

    Public Property UseDiskID() As Boolean
        Get
            Return m_UseDiskID
        End Get
        Set(ByVal value As Boolean)
            m_UseDiskID = value
        End Set
    End Property

    Public Property UseBaseID() As Boolean
        Get
            Return m_UseBaseID
        End Get
        Set(ByVal value As Boolean)
            m_UseBaseID = value
        End Set
    End Property

    Public Property UseVideoID() As Boolean
        Get
            Return m_UseVideoID
        End Get
        Set(ByVal value As Boolean)
            m_UseVideoID = value
        End Set
    End Property

    Public Property UseMacID() As Boolean
        Get
            Return m_UseMacID
        End Get
        Set(ByVal value As Boolean)
            m_UseMacID = value
        End Set
    End Property

    Public ReadOnly Property Value() As String
        Get
            RaiseEvent StartingWith("All")

            m_CpuID = ""
            If m_UseCpuID Then m_CpuID = CpuID()
            m_BiosID = ""
            If m_UseBiosID Then m_BiosID = BiosID()
            m_DiskID = ""
            If m_UseDiskID Then m_DiskID = DiskID()
            m_BaseID = ""
            If m_UseBaseID Then m_BaseID = BaseID()
            m_VideoID = ""
            If m_UseVideoID Then m_VideoID = VideoID()
            m_MacID = ""
            If m_UseMacID Then m_MacID = MacID()

            Return Pack(m_CpuID & m_BiosID & m_DiskID & m_BaseID & m_VideoID & m_MacID)

            RaiseEvent DoneWith("All")
        End Get
    End Property

    Private Function Identifier(ByVal wmiClass As String, ByVal wmiProperty As String, ByVal wmiMustBeTrue As String) As String
        'Return a hardware identifier

        Dim Result As String = ""
        Dim mc As New System.Management.ManagementClass(wmiClass)
        Dim moc As System.Management.ManagementObjectCollection = mc.GetInstances
        Dim mo As System.Management.ManagementObject

        For Each mo In moc
            If mo(wmiMustBeTrue).ToString = "True" Then
                'Only get the first one
                If Result = "" Then
                    Try
                        Result = mo(wmiProperty).ToString
                        Exit For
                    Catch ex As Exception
                        'Ignore error
                    End Try
                End If
            End If
        Next mo

        Return Result
    End Function

    Private Function Identifier(ByVal wmiClass As String, ByVal wmiProperty As String) As String
        'Return a hardware identifier

        Dim Result As String = ""
        Dim mc As New System.Management.ManagementClass(wmiClass)
        Dim moc As System.Management.ManagementObjectCollection = mc.GetInstances
        Dim mo As System.Management.ManagementObject

        For Each mo In moc
            'Only get the first one
            If Result = "" Then
                Try
                    Result = mo(wmiProperty).ToString
                    Exit For
                Catch ex As Exception
                    'Ignore error
                End Try
            End If
        Next mo

        Return Result
    End Function

    Private Function CpuID() As String
        RaiseEvent StartingWith("CpuID")

        'Uses first CPU identifier available in order of preference
        'Don't get all identifiers as very time consuming
        ' Do not get the following because it's mostly unavailable
        'Dim RetVal As String = Identifier("Win32_Processor", "UniqueId")

        Dim RetVal As String = String.Empty
        If RetVal = "" Then   'If no UniqueId, use ProcessorID
            RetVal = Identifier("Win32_Processor", "ProcessorId")

            If RetVal = "" Then   'If no ProcessorID, use Name
                RetVal = Identifier("Win32_Processor", "Name")

                If RetVal = "" Then   'If no Name, use Manufacturer
                    RetVal = Identifier("Win32_Processor", "Manufacturer")
                End If

                'Add clock speed for extra security
                RetVal += Identifier("Win32_Processor", "MaxClockSpeed")
            End If
        End If

        Return RetVal

        RaiseEvent DoneWith("CpuID")
    End Function

    Private Function BiosID() As String
        RaiseEvent StartingWith("BiosID")

        'BIOS Identifier

        Return Identifier("Win32_BIOS", "Manufacturer") _
          & Identifier("Win32_BIOS", "SMBIOSBIOSVersion") _
          & Identifier("Win32_BIOS", "SerialNumber") _
          & Identifier("Win32_BIOS", "ReleaseDate") _
          & Identifier("Win32_BIOS", "Version")
        '          & Identifier("Win32_BIOS", "IdentificationCode") _

        RaiseEvent DoneWith("BiosID")
    End Function

    Private Function DiskID() As String
        RaiseEvent StartingWith("DiskID")

        'Main physical hard drive ID

        Return Identifier("Win32_DiskDrive", "Manufacturer") _
          & Identifier("Win32_DiskDrive", "Signature") _
          & Identifier("Win32_DiskDrive", "TotalHeads")
        'Identifier("Win32_DiskDrive", "Model") _

        RaiseEvent DoneWith("CpuID")
    End Function

    Private Function BaseID() As String
        RaiseEvent StartingWith("BaseID")

        'Motherboard ID

        Return Identifier("Win32_BaseBoard", "Model") _
          & Identifier("Win32_BaseBoard", "Manufacturer") _
          & Identifier("Win32_BaseBoard", "Name") _
          & Identifier("Win32_BaseBoard", "SerialNumber")

        RaiseEvent DoneWith("BaseID")
    End Function

    Private Function VideoID() As String
        RaiseEvent StartingWith("VideoID")

        'Primary video controller ID

        Return Identifier("Win32_VideoController", "DriverVersion") _
          & Identifier("Win32_VideoController", "Name")

        RaiseEvent DoneWith("VideoID")
    End Function

    Private Function MacID() As String
        RaiseEvent StartingWith("MacID")

        'First enabled network card ID

        Return Identifier("Win32_NetworkAdapterConfiguration", "MACAddress", "IPEnabled")

        RaiseEvent StartingWith("MacID")
    End Function

    Private Function Pack(ByVal Text As String) As String
        RaiseEvent StartingWith("Packing")

        'Packs the string to m_ReturnLength digits
        'If m_ReturnLength=-1 : Return complete string

        Dim RetVal As String
        Dim X As Long
        Dim Y As Long
        Dim N As Char

        For Each N In Text
            Y += 1
            X += (Asc(N) * Y)
        Next N

        If m_ReturnLength > 0 Then
            RetVal = X.ToString.PadRight(CType(m_ReturnLength, Integer), Chr(48))
        Else
            RetVal = X.ToString
        End If

        If m_ReturnLength = 0 Then
            m_TotalLength = RetVal.Length
            Return RetVal
        Else
            m_TotalLength = RetVal.Length
            Return RetVal.Substring(0, CType(m_ReturnLength, Integer))
        End If

        RaiseEvent DoneWith("Packing")
    End Function
End Class




