Imports System
Imports System.Runtime.InteropServices
Imports System.Globalization
Imports System.Diagnostics
Imports System.Text

#Region "Copyright"
' This project is available from SVN on SourceForge.net under the main project, Activelock !
'
' ProjectPage: http://sourceforge.net/projects/activelock
' WebSite: http://www.activeLockSoftware.com
' DeveloperForums: http://forums.activelocksoftware.com
' ProjectManager: Ismail Alkan - http://activelocksoftware.com/simplemachinesforum/index.php?action=profile;u=1
' ProjectLicense: BSD Open License - http://www.opensource.org/licenses/bsd-license.php
' ProjectPurpose: Copy Protection, Software Locking, Anti Piracy
'
' //////////////////////////////////////////////////////////////////////////////////////////
' *   ActiveLock
' *   Copyright 1998-2002 Nelson Ferraz
' *   Copyright 2003-2009 The Activelock - Ismail Alkan (ASG)
' *   All material is the property of the contributing authors.
' *
' *   Redistribution and use in source and binary forms, with or without
' *   modification, are permitted provided that the following conditions are
' *   met:
' *
' *     [o] Redistributions of source code must retain the above copyright
' *         notice, this list of conditions and the following disclaimer.
' *
' *     [o] Redistributions in binary form must reproduce the above
' *         copyright notice, this list of conditions and the following
' *         disclaimer in the documentation and/or other materials provided
' *         with the distribution.
' *
' *   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' *   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' *   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
' *   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
' *   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
' *   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
' *   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
' *   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' *   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' *   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' *   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' *
#End Region

'Ported from C# and modified as needed 11/11/06 By Dwain Snickles 
'Left Out Windows 64 Bit.
Namespace CWindows

    Public Enum OSPlatformId
        Win32s = 0
        Win32Windows = 1
        Win32NT = 2
        WinCE = 3
    End Enum 'OSPlatformId

    Public Enum OSBuildNumber
        None = 0
        Win2000SP4 = 2195
        WinXPSP2 = 2600
        Win2003SP1 = 3790
    End Enum 'OSBuildNumber

    <Flags()> _
    Public Enum OSSuites
        None = 0
        SmallBusiness = &H1
        Enterprise = &H2
        BackOffice = &H4
        Communications = &H8
        Terminal = &H10
        SmallBusinessRestricted = &H20
        EmbeddedNT = &H40
        Datacenter = &H80
        SingleUserTS = &H100
        Personal = &H200
        Blade = &H400
        EmbeddedRestricted = &H800
    End Enum 'OSSuites

    Public Enum OSProductType
        Invalid = 0
        Workstation = 1
        DomainController = 2
        Server = 3
    End Enum 'OSProductType

    Public Enum OSVersion
        Win32s
        Win95
        Win98
        WinME
        WinNT351
        WinNT4
        Win2000
        WinXP
        Win2003
        WinCE
        WinVista
        Unknown
    End Enum 'OSVersion

    '-----------------------------------------------------------------------------
    ' OSVersionInfo
    Public Class OSVersionInfo
        'Implements ICloneable 'ToDo: Add Implements Clauses for implementation methods of these interface(s)

        '-----------------------------------------------------------------------------
        ' Constants
        Private Class MajorVersionConst
            Public Const Win32s As Integer = 0 ' TODO: check
            Public Const Win95 As Integer = 4
            Public Const Win98 As Integer = 4
            Public Const WinME As Integer = 4
            Public Const WinNT351 As Integer = 3
            Public Const WinNT4 As Integer = 4
            Public Const WinNT5 As Integer = 5
            Public Const WinNT6 As Integer = 6
            Public Const Win2000 As Integer = WinNT5
            Public Const WinXP As Integer = WinNT5
            Public Const Win2003 As Integer = WinNT5
            Public Const WinVista As Integer = WinNT6

            Private Sub New()
            End Sub 'New
        End Class 'MajorVersionConst 

        Private Class MinorVersionConst
            Public Const Win32s As Integer = 0 ' TODO: check
            Public Const Win95 As Integer = 0
            Public Const Win98 As Integer = 10
            Public Const WinME As Integer = 90
            Public Const WinNT351 As Integer = 51
            Public Const WinNT4 As Integer = 0
            Public Const Win2000 As Integer = 0
            Public Const WinXP As Integer = 1
            Public Const Win2003 As Integer = 2
            Public Const WinVista As Integer = 0

            Private Sub New()
            End Sub 'New
        End Class 'MinorVersionConst 
        '-----------------------------------------------------------------------------
        ' Static Fields
        Private Shared _Win32s As New OSVersionInfo(CWindows.OSPlatformId.Win32s, MajorVersionConst.Win32s, MinorVersionConst.Win32s, True)
        Private Shared _Win95 As New OSVersionInfo(CWindows.OSPlatformId.Win32Windows, MajorVersionConst.Win95, MinorVersionConst.Win95, True)
        Private Shared _Win98 As New OSVersionInfo(CWindows.OSPlatformId.Win32Windows, MajorVersionConst.Win98, MinorVersionConst.Win98, True)
        Private Shared _WinME As New OSVersionInfo(CWindows.OSPlatformId.Win32Windows, MajorVersionConst.WinME, MinorVersionConst.WinME, True)
        Private Shared _WinNT351 As New OSVersionInfo(CWindows.OSPlatformId.Win32NT, MajorVersionConst.WinNT351, MinorVersionConst.WinNT351, True)
        Private Shared _WinNT4 As New OSVersionInfo(CWindows.OSPlatformId.Win32NT, MajorVersionConst.WinNT4, MinorVersionConst.WinNT4, True)
        Private Shared _Win2000 As New OSVersionInfo(CWindows.OSPlatformId.Win32NT, MajorVersionConst.Win2000, MinorVersionConst.Win2000, True)
        Private Shared _WinXP As New OSVersionInfo(CWindows.OSPlatformId.Win32NT, MajorVersionConst.WinXP, MinorVersionConst.WinXP, True)
        Private Shared _Win2003 As New OSVersionInfo(CWindows.OSPlatformId.Win32NT, MajorVersionConst.Win2003, MinorVersionConst.Win2003, True)
        Private Shared _WinVista As New OSVersionInfo(CWindows.OSPlatformId.Win32NT, MajorVersionConst.WinVista, MinorVersionConst.WinVista, True)
        Private Shared _WinCE As New OSVersionInfo(CWindows.OSPlatformId.WinCE, True) ' TODO: WinCE version
        '-----------------------------------------------------------------------------
        ' Static Properties
        Private Shared ReadOnly Property Win32s() As OSVersionInfo
            Get
                Return _Win32s
            End Get
        End Property

        Private Shared ReadOnly Property Win95() As OSVersionInfo
            Get
                Return _Win95
            End Get
        End Property

        Private Shared ReadOnly Property Win98() As OSVersionInfo
            Get
                Return _Win98
            End Get
        End Property

        Private Shared ReadOnly Property WinME() As OSVersionInfo
            Get
                Return _WinME
            End Get
        End Property

        Private Shared ReadOnly Property WinNT351() As OSVersionInfo
            Get
                Return _WinNT351
            End Get
        End Property

        Private Shared ReadOnly Property WinNT4() As OSVersionInfo
            Get
                Return _WinNT4
            End Get
        End Property

        Private Shared ReadOnly Property Win2000() As OSVersionInfo
            Get
                Return _Win2000
            End Get
        End Property

        Private Shared ReadOnly Property WinXP() As OSVersionInfo
            Get
                Return _WinXP
            End Get
        End Property

        Private Shared ReadOnly Property Win2003() As OSVersionInfo
            Get
                Return _Win2003
            End Get
        End Property

        Private Shared ReadOnly Property WinVista() As OSVersionInfo
            Get
                Return _WinVista
            End Get
        End Property

        Private Shared ReadOnly Property WinCE() As OSVersionInfo
            Get
                Return _WinCE
            End Get
        End Property
        '-----------------------------------------------------------------------------
        ' Static methods
        Public Shared Function GetOSVersionInfo(ByVal v As OSVersion) As OSVersionInfo
            Select Case v
                Case CWindows.OSVersion.Win32s
                    Return Win32s
                Case CWindows.OSVersion.Win95
                    Return Win95
                Case CWindows.OSVersion.Win98
                    Return Win98
                Case CWindows.OSVersion.WinME
                    Return WinME
                Case CWindows.OSVersion.WinNT351
                    Return WinNT351
                Case CWindows.OSVersion.WinNT4
                    Return WinNT4
                Case CWindows.OSVersion.Win2000
                    Return Win2000
                Case CWindows.OSVersion.WinXP
                    Return WinXP
                Case CWindows.OSVersion.Win2003
                    Return Win2003
                Case CWindows.OSVersion.WinVista
                    Return WinVista
                Case CWindows.OSVersion.WinCE
                    Return WinCE
                Case Else
                    Return Nothing
            End Select
        End Function 'GetOSVersionInfo

        '-----------------------------------------------------------------------------
        ' Fields
        ' normal fields
        Private _OSPlatformId As OSPlatformId

        Private _MajorVersion As Integer = -1
        Private _MinorVersion As Integer = -1
        Private _BuildNumber As Integer = -1
        '		private int _PlatformId;
        Private _CSDVersion As String = String.Empty

        ' extended fields
        Private _OSSuiteFlags As OSSuites
        Private _OSProductType As OSProductType

        Private _ServicePackMajor As Int16 = -1
        Private _ServicePackMinor As Int16 = -1
        '		private UInt16 _SuiteMask;
        '		private byte _ProductType;
        Private _Reserved As Byte

        ' state fields
        Private _Locked As Boolean = False
        Private _ExtendedPropertiesAreSet As Boolean = False
        '-----------------------------------------------------------------------------
        ' Normal Properties

        Public Property OSPlatformId() As CWindows.OSPlatformId
            Get
                Return _OSPlatformId
            End Get
            Set(ByVal Value As CWindows.OSPlatformId)
                CheckLock("OSPlatformId")

                _OSPlatformId = Value
            End Set
        End Property

        Public Property OSMajorVersion() As Integer
            Get
                Return _MajorVersion
            End Get
            Set(ByVal Value As Integer)
                CheckLock("MajorVersion")

                _MajorVersion = Value
            End Set
        End Property

        Public Property OSMinorVersion() As Integer
            Get
                Return _MinorVersion
            End Get
            Set(ByVal Value As Integer)
                CheckLock("MinorVersion")

                _MinorVersion = Value
            End Set
        End Property

        Public Property BuildNumber() As Integer
            Get
                Return _BuildNumber
            End Get
            Set(ByVal Value As Integer)
                CheckLock("BuildNumber")

                _BuildNumber = Value
            End Set
        End Property

        Public Property OSCSDVersion() As String
            Get
                Return _CSDVersion
            End Get
            Set(ByVal Value As String)
                CheckLock("CSDVersion")

                _CSDVersion = Value
            End Set
        End Property
        '-----------------------------------------------------------------------------
        ' Extended Properties

        Public Property OSSuiteFlags() As CWindows.OSSuites
            Get
                CheckExtendedProperty("OSSuiteFlags")

                Return _OSSuiteFlags
            End Get

            Set(ByVal Value As CWindows.OSSuites)
                CheckLock("OSSuiteFlags")

                _OSSuiteFlags = Value
            End Set
        End Property

        Public Property OSProductType() As CWindows.OSProductType
            Get
                CheckExtendedProperty("OSProductType")

                Return _OSProductType
            End Get

            Set(ByVal Value As CWindows.OSProductType)
                CheckLock("OSProductType")

                _OSProductType = Value
            End Set
        End Property

        Public Property OSServicePackMajor() As Int16
            Get
                CheckExtendedProperty("ServicePackMajor")

                Return _ServicePackMajor
            End Get

            Set(ByVal Value As Int16)
                CheckLock("ServicePackMajor")

                _ServicePackMajor = Value
            End Set
        End Property

        Public Property OSServicePackMinor() As Int16
            Get
                CheckExtendedProperty("ServicePackMinor")

                Return _ServicePackMinor
            End Get

            Set(ByVal Value As Int16)
                CheckLock("ServicePackMinor")

                _ServicePackMinor = Value
            End Set
        End Property

        Public Property OSReserved() As Byte
            Get
                CheckExtendedProperty("Reserved")

                Return _Reserved
            End Get

            Set(ByVal Value As Byte)
                CheckLock("Reserved")

                _Reserved = Value
            End Set
        End Property

        '-----------------------------------------------------------------------------
        ' Get Properties

        Public ReadOnly Property Platform() As Integer
            Get
                Return Fix(_OSPlatformId)
            End Get
        End Property

        Public ReadOnly Property SuiteMask() As Integer
            Get
                CheckExtendedProperty("SuiteMask")

                Return Fix(_OSSuiteFlags)
            End Get
        End Property

        Public ReadOnly Property ProductType() As Byte
            Get
                CheckExtendedProperty("ProductType")

                Return CByte(_OSProductType)
            End Get
        End Property

        '-----------------------------------------------------------------------------
        ' Calculated Properties

        Public ReadOnly Property Version() As System.Version
            Get
                If OSMajorVersion < 0 OrElse OSMinorVersion < 0 Then
                    Return New Version
                End If
                If BuildNumber < 0 Then
                    Return New Version(OSMajorVersion, OSMinorVersion)
                End If
                Return New Version(OSMajorVersion, OSMinorVersion, BuildNumber)
            End Get
        End Property

        Public ReadOnly Property VersionString() As String
            Get
                Return Version.ToString()
            End Get
        End Property

        Public ReadOnly Property OSPlatformIdString() As String
            Get
                Select Case OSPlatformId
                    Case CWindows.OSPlatformId.Win32s
                        Return "Windows 32s"
                    Case CWindows.OSPlatformId.Win32Windows
                        Return "Windows 32"
                    Case CWindows.OSPlatformId.Win32NT
                        Return "Windows NT"
                    Case CWindows.OSPlatformId.WinCE
                        Return "Windows CE"

                    Case Else
                        'Throw New InvalidOperationException("Invalid OSPlatformId: " + OSPlatformId)
                        Return Nothing
                End Select
            End Get
        End Property

        Friend Shared Function OSSuiteFlag(ByVal flags As CWindows.OSSuites, ByVal test As CWindows.OSSuites) As Boolean
            Return (flags And test) > 0
        End Function 'OSSuiteFlag

        Public ReadOnly Property OSSuiteString() As String
            Get
                Dim s As String = String.Empty

                Dim flags As CWindows.OSSuites = OSSuiteFlags

                If OSSuiteFlag(flags, CWindows.OSSuites.SmallBusiness) Then
                    OSSuiteStringAdd(s, "Small Business")
                End If
                If OSSuiteFlag(flags, CWindows.OSSuites.Enterprise) Then
                    Select Case OSVersion
                        Case CWindows.OSVersion.WinNT4
                            OSSuiteStringAdd(s, "Enterprise")
                        Case CWindows.OSVersion.Win2000
                            OSSuiteStringAdd(s, "Advanced")
                        Case CWindows.OSVersion.Win2003
                            OSSuiteStringAdd(s, "Enterprise")
                    End Select
                End If
                If OSSuiteFlag(flags, CWindows.OSSuites.BackOffice) Then
                    OSSuiteStringAdd(s, "BackOffice")
                End If
                If OSSuiteFlag(flags, CWindows.OSSuites.Communications) Then
                    OSSuiteStringAdd(s, "Communications")
                End If
                If OSSuiteFlag(flags, CWindows.OSSuites.Terminal) Then
                    OSSuiteStringAdd(s, "Terminal Services")
                End If
                If OSSuiteFlag(flags, CWindows.OSSuites.SmallBusinessRestricted) Then
                    OSSuiteStringAdd(s, "Small Business Restricted")
                End If
                If OSSuiteFlag(flags, CWindows.OSSuites.EmbeddedNT) Then
                    OSSuiteStringAdd(s, "Embedded")
                End If
                If OSSuiteFlag(flags, CWindows.OSSuites.Datacenter) Then
                    OSSuiteStringAdd(s, "Datacenter")
                End If
                '				if ( OSSuiteFlag( flags, CWindows.OSSuites.SingleUserTS ) )
                '					OSSuiteStringAdd( ref s, "Single User Terminal Services" );
                If OSSuiteFlag(flags, CWindows.OSSuites.Personal) Then
                    OSSuiteStringAdd(s, "Home Edition")
                End If
                If OSSuiteFlag(flags, CWindows.OSSuites.Blade) Then
                    OSSuiteStringAdd(s, "Web Edition")
                End If
                If OSSuiteFlag(flags, CWindows.OSSuites.EmbeddedRestricted) Then
                    OSSuiteStringAdd(s, "Embedded Restricted")
                End If
                Return s
            End Get
        End Property

        Private Shared Sub OSSuiteStringAdd(ByRef s As String, ByVal suite As String)
            If s.Length > 0 Then
                s += ", "
            End If
            s += suite
        End Sub 'OSSuiteStringAdd

        Public ReadOnly Property OSProductTypeString() As String
            Get
                Select Case OSProductType
                    Case CWindows.OSProductType.Workstation

                        Select Case OSVersion
                            Case CWindows.OSVersion.Win32s
                                Return String.Empty
                            Case CWindows.OSVersion.Win95
                                Return String.Empty
                            Case CWindows.OSVersion.Win98
                                Return String.Empty
                            Case CWindows.OSVersion.WinME
                                Return String.Empty
                            Case CWindows.OSVersion.WinNT351
                                Return String.Empty
                            Case CWindows.OSVersion.WinNT4
                                Return "Workstation"
                            Case CWindows.OSVersion.Win2000
                                Return "Professional"
                            Case CWindows.OSVersion.WinXP
                                If OSSuiteFlag(OSSuiteFlags, CWindows.OSSuites.Personal) Then
                                    Return "Home Edition"
                                Else
                                    Return "Professional"
                                End If
                            Case CWindows.OSVersion.Win2003
                                Return String.Empty
                            Case CWindows.OSVersion.WinVista
                                Return String.Empty
                            Case CWindows.OSVersion.WinCE
                                Return String.Empty
                            Case Else
                                Return "Unkown Windows ProductType"
                        End Select

                    Case CWindows.OSProductType.DomainController
                        Dim s As String = OSSuiteString

                        If s.Length > 0 Then
                            s += " "
                        End If
                        Return s + "Domain Controller"

                    Case CWindows.OSProductType.Server
                        Dim s As String = OSSuiteString

                        If s.Length > 0 Then
                            s += " "
                        End If
                        Return s + "Server"

                    Case Else
                        Return "Unkown Windows ProductType"
                End Select
            End Get
        End Property

        Public ReadOnly Property OSVersion() As CWindows.OSVersion
            Get
                Select Case OSPlatformId
                    Case CWindows.OSPlatformId.Win32s
                        Return CWindows.OSVersion.Win32s

                    Case CWindows.OSPlatformId.Win32Windows

                        Select Case OSMinorVersion
                            Case MinorVersionConst.Win95
                                Return CWindows.OSVersion.Win95
                            Case MinorVersionConst.Win98
                                Return CWindows.OSVersion.Win98
                            Case MinorVersionConst.WinME
                                Return CWindows.OSVersion.WinME
                            Case Else
                                Return CWindows.OSVersion.Unknown
                        End Select

                    Case CWindows.OSPlatformId.Win32NT

                        Select Case OSMajorVersion
                            Case MajorVersionConst.WinNT351
                                Return CWindows.OSVersion.WinNT351
                            Case MajorVersionConst.WinNT4
                                Return CWindows.OSVersion.WinNT4

                            Case MajorVersionConst.WinNT5

                                Select Case OSMinorVersion
                                    Case MinorVersionConst.Win2000
                                        Return CWindows.OSVersion.Win2000
                                    Case MinorVersionConst.WinXP
                                        Return CWindows.OSVersion.WinXP
                                    Case MinorVersionConst.Win2003
                                        Return CWindows.OSVersion.Win2003
                                    Case Else
                                        Return CWindows.OSVersion.Unknown
                                End Select

                            Case MajorVersionConst.WinNT6

                                Select Case OSMinorVersion
                                    Case MinorVersionConst.WinVista
                                        Return CWindows.OSVersion.WinVista
                                End Select
                            Case Else
                                Return CWindows.OSVersion.Unknown
                        End Select

                    Case CWindows.OSPlatformId.WinCE
                        Return CWindows.OSVersion.WinCE

                    Case Else
                        Return CWindows.OSVersion.Unknown
                End Select
            End Get
        End Property

        Public ReadOnly Property OSVersionString() As String
            Get
                Select Case OSVersion
                    Case CWindows.OSVersion.Win32s
                        Return "Windows 32s"
                    Case CWindows.OSVersion.Win95
                        Return "Windows 95"
                    Case CWindows.OSVersion.Win98
                        Return "Windows 98"
                    Case CWindows.OSVersion.WinME
                        Return "Windows ME"
                    Case CWindows.OSVersion.WinNT351
                        Return "Windows NT 3.51"
                    Case CWindows.OSVersion.WinNT4
                        Return "Windows NT 4"
                    Case CWindows.OSVersion.Win2000
                        Return "Windows 2000"
                    Case CWindows.OSVersion.WinXP
                        Return "Windows XP"
                    Case CWindows.OSVersion.Win2003
                        Return "Windows 2003"
                    Case CWindows.OSVersion.WinVista
                        Return "Windows Vista"
                    Case CWindows.OSVersion.WinCE
                        Return "Windows CE"

                    Case Else
                        'Throw New InvalidOperationException("Invalid OSVersion: " + OSVersion)
                        Return Nothing
                End Select
            End Get
        End Property

        '-----------------------------------------------------------------------------
        ' State Properties

        Public Property ExtendedPropertiesAreSet() As Boolean
            Get
                Return _ExtendedPropertiesAreSet
            End Get
            Set(ByVal Value As Boolean)
                _ExtendedPropertiesAreSet = Value
            End Set
        End Property

        Public ReadOnly Property IsLocked() As Boolean
            Get
                Return _Locked
            End Get
        End Property

        Public Sub Lock()
            _Locked = True
        End Sub 'Lock

        '-----------------------------------------------------------------------------
        ' Property helpers
        Private Sub CheckExtendedProperty(ByVal [property] As String)
            If _ExtendedPropertiesAreSet Then
                Return
            End If
            'Throw New InvalidOperationException("'" + [property] + "' is not set")
        End Sub 'CheckExtendedProperty

        Private Sub CheckLock(ByVal [property] As String)
            If Not _Locked Then
                Return
            End If
            Throw New InvalidOperationException("Cannot set '" + [property] + "' on locked instance")
        End Sub 'CheckLock

        '-----------------------------------------------------------------------------
        ' Constructors
        Public Sub New()
        End Sub 'New

        Private Sub New(ByVal osPlatformId As OSPlatformId)
            _OSPlatformId = osPlatformId
        End Sub 'New

        Private Sub New(ByVal osPlatformId As OSPlatformId, ByVal locked As Boolean)
            _OSPlatformId = osPlatformId

            _Locked = locked
        End Sub 'New

        Private Sub New(ByVal osPlatformId As OSPlatformId, ByVal majorVersion As Integer, ByVal minorVersion As Integer)
            _OSPlatformId = osPlatformId
            _MajorVersion = majorVersion
            _MinorVersion = minorVersion
        End Sub 'New

        Private Sub New(ByVal osPlatformId As OSPlatformId, ByVal majorVersion As Integer, ByVal minorVersion As Integer, ByVal locked As Boolean)
            _OSPlatformId = osPlatformId
            _MajorVersion = majorVersion
            _MinorVersion = minorVersion

            _Locked = locked
        End Sub 'New

        Public Overrides Function GetHashCode() As Integer
            Return MyBase.GetHashCode()
        End Function 'GetHashCode

        Public Overrides Function ToString() As String
            Dim s As String = OSVersionString

            If ExtendedPropertiesAreSet Then
                s += " " + OSProductTypeString
            End If
            If OSCSDVersion.Length > 0 Then
                s += " " + OSCSDVersion
            End If
            s += " v" + VersionString

            Return s
        End Function 'ToString

    End Class 'OSVersionInfo 
    '-----------------------------------------------------------------------------
    ' OSVersionInfo
    '-----------------------------------------------------------------------------
    ' OperatingSystemVersion
    Public Class OperatingSystemVersion
        Inherits CWindows.OSVersionInfo

        <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
                             Private Overloads Shared Function GetVersionEx(<[In](), Out()> ByVal osVersionInfo As OSVERSIONINFO) As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
                  Private Overloads Shared Function GetVersionEx(<[In](), Out()> ByVal osVersionInfoEx As OSVERSIONINFOEX) As Boolean
        End Function

        Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Integer) As Integer

        Public Declare Auto Function GetVolumeInformation Lib "kernel32" (ByVal lpRootPathName As String, _
            ByVal lpVolumeNameBuffer As StringBuilder, ByVal nVolumeNameSize As Integer, _
            ByRef lpVolumeSerialNumber As Integer, ByRef lpMaximumComponentLength As Integer, _
            ByRef lpFileSystemFlags As Integer, ByVal lpFileSystemNameBuffer As StringBuilder, _
            ByRef nFileSystemNameSize As Integer) As Boolean

        'XXXXXX Windows Plus Version XXXXX

        Public Function IsWinNT4Plus() As Boolean
            Dim Host As New OperatingSystemVersion
            IsWinNT4Plus = (Host.OSPlatformId = OSPlatformId.Win32NT) And _
                          (Host.Version.Major >= 4)
        End Function

        Public Function IsWin2000Plus() As Boolean
            Dim Host As New OperatingSystemVersion
            IsWin2000Plus = (Host.OSPlatformId = OSPlatformId.Win32NT) And _
                            (Host.Version.Major = 5 And Host.Version.Minor >= 0)
        End Function

        Public Function IsWinXPPlus() As Boolean
            Dim Host As New OperatingSystemVersion
            IsWinXPPlus = (Host.OSPlatformId = OSPlatformId.Win32NT) And _
                              (Host.Version.Major >= 5 And Host.Version.Minor >= 1)
        End Function

        Public Function IsWinVistaPlus() As Boolean
            Dim Host As New OperatingSystemVersion
            IsWinVistaPlus = (Host.OSPlatformId = OSPlatformId.Win32NT) And _
                         (Host.Version.Major >= 6)
        End Function

        'XXXXX Server Indformation XXXXX

        Public Function IsBackOfficeServer() As Boolean
            If IsWinNT4Plus() Then
                IsBackOfficeServer = CType((SuiteMask And VerSuiteMask.VER_SUITE_BACKOFFICE), Boolean)
            End If
        End Function

        Public Function IsBladeServer() As Boolean
            Dim Host As New OperatingSystemVersion
            If IsWin2003Server() Then
                IsBladeServer = CType((SuiteMask And VerSuiteMask.VER_SUITE_BLADE), Boolean)
            End If
        End Function

        Public Function IsDomainController() As Boolean
            'Returns True if the server is a domain
            'controller (Win 2000 or later), including
            'under active directory
            'OSVERSIONINFOEX supported on NT4 or
            'later only, so a test is required
            'before using
            If IsWin2000Server() Then
                IsDomainController = (OSProductType = VerProductType.VER_NT_SERVER) And (OSProductType = VerProductType.VER_NT_DOMAIN_CONTROLLER)
            End If
        End Function

        Public Function IsEnterpriseServer() As Boolean
            'Returns True if Windows NT 4.0 Enterprise Edition,
            'Windows 2000 Advanced Server, or Windows Server 2003
            'Enterprise Edition is installed.
            'OSVERSIONINFOEX supported on NT4 or
            'later only, so a test is required
            'before using
            If IsWinNT4Plus() Then
                IsEnterpriseServer = (OSProductType = VerProductType.VER_NT_SERVER) And CType((SuiteMask And VerSuiteMask.VER_SUITE_ENTERPRISE), Boolean)
            End If
        End Function

        Public Function IsSmallBusinessServer() As Boolean
            'Returns True if Microsoft Small Business Server is installed
            'OSVERSIONINFOEX supported on NT4 or
            'later only, so a test is required
            'before using
            If IsWinNT4Plus() Then
                IsSmallBusinessServer = CType((SuiteMask And VerSuiteMask.VER_SUITE_SMALLBUSINESS), Boolean)
            End If
        End Function

        Public Function IsSmallBusinessRestrictedServer() As Boolean
            'Returns True if Microsoft Small Business Server
            'is installed with the restrictive client license
            'in force
            'OSVERSIONINFOEX supported on NT4 or
            'later only, so a test is required
            'before using
            If IsWinNT4Plus() Then
                IsSmallBusinessRestrictedServer = CType((SuiteMask And VerSuiteMask.VER_SUITE_SMALLBUSINESS_RESTRICTED), Boolean)
            End If
        End Function

        Public Function IsTerminalServer() As Boolean
            'Returns True if Terminal Services is installed
            'OSVERSIONINFOEX supported on NT4 or
            'later only, so a test is required
            'before using
            If IsWinNT4Plus() Then
                IsTerminalServer = CType((SuiteMask And VerSuiteMask.VER_SUITE_TERMINAL), Boolean)
            End If
        End Function

        'XXXXXX Windows Operatong Systems XXXXXX

        Public Function IsWin95() As Boolean
            IsWin95 = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_WINDOWS) And _
                          (Version.Major = 4 And Version.Minor = 0) And _
                          (BuildNumber = 950)
        End Function

        Public Function IsWin95OSR2() As Boolean
            IsWin95OSR2 = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_WINDOWS) And _
                               (Version.Major = 4 And Version.Minor = 0) And _
                               (BuildNumber = 1111)
        End Function

        Public Function IsWin98() As Boolean
            IsWin98 = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_WINDOWS) And _
                      (Version.Major = 4 And Version.Minor = 10) And _
                      (BuildNumber >= 1998)
        End Function

        Public Function IsWinME() As Boolean
            IsWinME = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_WINDOWS) And _
                          (Version.Major = 4 And Version.Minor = 90) And _
                          (BuildNumber >= 3000)
        End Function

        Public Function IsWinNT3() As Boolean
            IsWinNT3 = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_NT) And _
                       (Version.Major = 3 And Version.Minor = 0)
        End Function

        Public Function IsWinNT31() As Boolean
            IsWinNT31 = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_NT) And _
                       (Version.Major = 3 And Version.Minor = 1)
        End Function

        Public Function IsWinNT35() As Boolean
            IsWinNT35 = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_NT) And _
                       (Version.Major = 3 And Version.Minor = 5)
        End Function

        Public Function IsWinNT351() As Boolean
            IsWinNT351 = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_NT) And _
                       (Version.Major = 3 And Version.Minor = 51)
        End Function

        Public Function IsWinNT4() As Boolean
            IsWinNT4 = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_NT) And _
                       (Version.Major = 4 And Version.Minor = 0) And _
                       (BuildNumber >= 1381)
        End Function

        Public Function IsWinNT4Workstation() As Boolean
            'returns True if running Windows NT4 Workstation
            If IsWinNT4() Then
                IsWinNT4Workstation = CType((OSProductType And VerProductType.VER_NT_WORKSTATION), Boolean)
            End If
        End Function

        Public Function IsWinNT4Server() As Boolean
            'Dim Sys As OperatingSystemVersion
            If IsWinNT4() Then
                IsWinNT4Server = CType((OSProductType And VerProductType.VER_NT_SERVER), Boolean)
            End If
        End Function

        Public Function IsWin2000() As Boolean
            IsWin2000 = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_NT) And _
                        (Version.Major = 5 And Version.Minor = 0) And _
                        (BuildNumber >= 2195)
        End Function

        Public Function IsWin2000AdvancedServer() As Boolean
            If IsWin2000Plus() Then
                IsWin2000AdvancedServer = ((OSProductType = VerProductType.VER_NT_SERVER) Or _
                                           (OSProductType = VerProductType.VER_NT_DOMAIN_CONTROLLER)) And _
                                           CType((SuiteMask And VerSuiteMask.VER_SUITE_ENTERPRISE), Boolean)
            End If
        End Function

        Public Function IsWin2000Server() As Boolean
            If IsWin2000() Then
                IsWin2000Server = (OSProductType = VerProductType.VER_NT_SERVER)
            End If
        End Function

        Public Function IsWin2000Workstation() As Boolean
            If IsWin2000() Then
                IsWin2000Workstation = CType((OSProductType And VerProductType.VER_NT_WORKSTATION), Boolean)
            End If
        End Function

        Public Function IsWin2003Server() As Boolean
            'returns True if running Windows 2003 (.NET) Server
            IsWin2003Server = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_NT) And _
                              (Version.Major = 5 And Version.Minor = 2) And _
                              (BuildNumber = 3790)
        End Function

        Public Function IsWin2003ServerR2() As Boolean
            'returns True if running
            'Windows 2003 (.NET) Server Release 2
            If IsWin2003Server() Then
                IsWin2003ServerR2 = CType(GetSystemMetrics(VerSuiteMask.VER_SM_SERVERR2), Boolean)
            End If
        End Function

        Public Function IsWinXP() As Boolean
            IsWinXP = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_NT) And _
                (Version.Major = 5 And Version.Minor = 1) And _
                (BuildNumber >= 2600)
        End Function

        Public Function IsWinXPSP2() As Boolean
            'returns True if running Windows XP SP2 (Service Pack 2)
            If IsWinXP() Then
                IsWinXPSP2 = InStr(OSCSDVersion, "Service Pack 2") > 0
            End If
        End Function

        Public Function IsWinXPMediaCenter() As Boolean
            'returns True if running Windows XP Media Centre
            If IsWinXP() Then
                IsWinXPMediaCenter = CType(GetSystemMetrics(VerSuiteMask.VER_SM_MEDIACENTER), Boolean)
            End If
        End Function

        Public Function IsWinXPHomeEdition() As Boolean
            'returns True if running Windows XP Home Edition
            If IsWinXP() Then
                IsWinXPHomeEdition = ((SuiteMask And VerSuiteMask.VER_SUITE_PERSONAL) = VerSuiteMask.VER_SUITE_PERSONAL)
            End If
        End Function

        Public Function IsWinXPProEdition() As Boolean
            'returns True if running WinXP Pro
            If IsWinXP() Then
                IsWinXPProEdition = Not ((SuiteMask And VerSuiteMask.VER_SUITE_PERSONAL) = VerSuiteMask.VER_SUITE_PERSONAL)
            End If
        End Function

        'Returns False positive?
        Public Function IsWinXPStarter() As Boolean
            'returns True if running Windows XP Starter
            If IsWinXP() Then
                IsWinXPStarter = CType(GetSystemMetrics(VerSuiteMask.VER_SM_STARTER), Boolean)
            End If
        End Function

        'Returns False positive?
        Public Function IsWinXPTabletPc() As Boolean
            ''returns True if running Windows XP Tablet Pc
            If IsWinXP() Then
                IsWinXPTabletPc = CType(GetSystemMetrics(VerSuiteMask.VER_SM_TABLETPC), Boolean)
            End If
        End Function

        Public Function IsWinXPEmbedded() As Boolean
            'Returns True if OS is Windows XP Embedded
            'OSVERSIONINFOEX supported on NT4 or
            'later only, so a test is required
            'before using
            If IsWinXP() Then
                IsWinXPEmbedded = CType((SuiteMask And VerSuiteMask.VER_SUITE_EMBEDDEDNT), Boolean)
            End If
        End Function

        'Left out is needed will need to be added
        Public Function IsWinXP64() As Boolean

        End Function

        Public Function IsWinVista() As Boolean
            'returns True if running Windows Vista
            IsWinVista = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_NT) And _
                         (Version.Major = 6)
        End Function

        Public Function IsWinLonghornServer() As Boolean
            'returns True if running Windows Longhorn Server
            IsWinLonghornServer = (OSPlatformId = VerSuiteMask.VER_PLATFORM_WIN32_NT) And _
                                  (Version.Major = 6 And Version.Minor = 0) And _
                                  (OSProductType <> VerProductType.VER_NT_WORKSTATION)
        End Function

        '-----------------------------------------------------------------------------
        ' Interop Types
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
        Private Class OSVERSIONINFO
            Public OSVersionInfoSize As Integer
            Public MajorVersion As Integer
            Public MinorVersion As Integer
            Public BuildNumber As Integer
            Public PlatformId As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=&H80)> _
            Public CSDVersion As String

            Public Sub New()
                OSVersionInfoSize = Marshal.SizeOf(Me)
            End Sub 'New

            Private Sub StopTheCompilerComplaining()
                MajorVersion = 0
                MinorVersion = 0
                BuildNumber = 0
                PlatformId = 0
                CSDVersion = String.Empty
            End Sub 'StopTheCompilerComplaining
        End Class 'OSVERSIONINFO

        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Auto)> _
        Private Class OSVERSIONINFOEX
            Public OSVersionInfoSize As Integer
            Public MajorVersion As Integer
            Public MinorVersion As Integer
            Public BuildNumber As Integer
            Public PlatformId As Integer
            <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=&H80)> _
            Public CSDVersion As String
            Public ServicePackMajor As Int16
            Public ServicePackMinor As Int16
            Public SuiteMask As Int16 'UInt16
            Public ProductType As Byte
            Public Reserved As Byte

            Public Sub New()
                OSVersionInfoSize = Marshal.SizeOf(Me)
            End Sub 'New

            Private Sub StopTheCompilerComplaining()
                MajorVersion = 0
                MinorVersion = 0
                BuildNumber = 0
                PlatformId = 0
                CSDVersion = String.Empty
                ServicePackMajor = 0
                ServicePackMinor = 0
                'SuiteMask = 0
                ProductType = 0
                Reserved = 0
            End Sub 'StopTheCompilerComplaining
        End Class 'OSVERSIONINFOEX
        '-----------------------------------------------------------------------------
        ' Interop constants
        Private Class VerPlatformId
            Public Const Win32s As Int32 = 0
            Public Const Win32Windows As Int32 = 1
            Public Const Win32NT As Int32 = 2
            Public Const WinCE As Int32 = 3

            Private Sub New()
            End Sub 'New
        End Class 'VerPlatformId 

        Private Class VerSuiteMask
            '--------------------------------------------------------------^--- Specified cast is not valid.
            Public Const VER_PLATFORM_WIN32s As Integer = 0
            Public Const VER_PLATFORM_WIN32_WINDOWS As Integer = 1
            Public Const VER_PLATFORM_WIN32_NT As Integer = 2
            Public Const VER_SM_MEDIACENTER As Long = 87
            Public Const VER_SM_SERVERR2 As Long = 89
            Public Const VER_SM_TABLETPC As Long = 86
            Public Const VER_SM_STARTER As Long = 88
            Public Const VER_WORKSTATION_NT As Integer = &H40000000
            Public Const VER_SUITE_SMALLBUSINESS As Integer = &H1
            Public Const VER_SUITE_ENTERPRISE As Integer = &H2
            Public Const VER_SERVER_NT As Integer = &H3
            Public Const VER_SUITE_BACKOFFICE As Integer = &H4
            Public Const VER_SUITE_COMMUNICATIONS As Integer = &H8
            Public Const VER_SUITE_TERMINAL As Integer = &H10
            Public Const VER_SUITE_SMALLBUSINESS_RESTRICTED As Integer = &H20
            Public Const VER_SUITE_EMBEDDEDNT As Integer = &H40
            Public Const VER_SUITE_DATACENTER As Integer = &H80
            Public Const VER_SUITE_SINGLEUSERTS As Integer = &H100
            Public Const VER_SUITE_PERSONAL As Integer = &H200
            Public Const VER_SUITE_BLADE As Integer = &H400
            Public Const VER_SUITE_EMBEDDED_RESTRICTED As Integer = &H800
            Public Const PROCESSOR_ARCHITECTURE_IA64 As Long = 6

            Private Sub New()
            End Sub 'New
        End Class 'VerSuiteMask 

        Private Class VerProductType
            Public Const VER_NT_WORKSTATION As Byte = &H1
            Public Const VER_NT_DOMAIN_CONTROLLER As Byte = &H2
            Public Const VER_NT_SERVER As Byte = &H3

            Private Sub New()
            End Sub 'New
        End Class 'VerProductType 
        '-----------------------------------------------------------------------------
        ' Interop methods
        Private Class NativeMethods
            Private Sub New()
            End Sub 'New

            <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
                      Public Overloads Shared Function GetVersionEx(<[In](), Out()> ByVal osVersionInfo As OSVERSIONINFO) As Boolean
            End Function

            <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
                      Public Overloads Shared Function GetVersionEx(<[In](), Out()> ByVal osVersionInfoEx As OSVERSIONINFOEX) As Boolean
            End Function

        End Class 'NativeMethods

        '-----------------------------------------------------------------------------
        ' Constructors
        Friend Sub New()
            Dim osVersionInfo As New OSVERSIONINFO

            If Not UseOSVersionInfoEx(osVersionInfo) Then
                InitOsVersionInfo(osVersionInfo)
            Else
                InitOsVersionInfoEx()
            End If
        End Sub 'New

        ' check for NT4 SP6 or later
        Private Shared Function UseOSVersionInfoEx(ByVal info As OSVERSIONINFO) As Boolean
            Dim b As Boolean = NativeMethods.GetVersionEx(info)

            'If Not b Then
            '    Dim [error] As Integer = Marshal.GetLastWin32Error()

            '    Throw New InvalidOperationException("Failed to get OSVersionInfo. Error = 0x" + [error].ToString("8X", CultureInfo.CurrentCulture))
            'End If

            If info.MajorVersion < 4 Then
                Return False
            End If
            If info.MajorVersion > 4 Then
                Return True
            End If
            If info.MinorVersion < 0 Then
                Return False
            End If
            If info.MinorVersion > 0 Then
                Return True
            End If
            ' TODO: CSDVersion for NT4 SP6
            If info.CSDVersion = "Service Pack 6" Then
                Return True
            End If
            Return False
        End Function 'UseOSVersionInfoEx

        Private Sub InitOsVersionInfo(ByVal info As OSVERSIONINFO)
            OSPlatformId = GetOSPlatformId(info.PlatformId)

            OSMajorVersion = info.MajorVersion
            OSMinorVersion = info.MinorVersion
            BuildNumber = info.BuildNumber
            '			PlatformId     = info.PlatformId   ;
            OSCSDVersion = info.CSDVersion
        End Sub 'InitOsVersionInfo

        Private Sub InitOsVersionInfoEx()
            Dim info As New OSVERSIONINFOEX

            Dim b As Boolean = NativeMethods.GetVersionEx(info)

            'If Not b Then
            '   Dim [error] As Integer = Marshal.GetLastWin32Error()

            '   Throw New InvalidOperationException("Failed to get OSVersionInfoEx. Error = 0x" + [error].ToString("8X", CultureInfo.CurrentCulture))
            'End If

            OSPlatformId = GetOSPlatformId(info.PlatformId)

            OSMajorVersion = info.MajorVersion
            OSMinorVersion = info.MinorVersion
            BuildNumber = info.BuildNumber
            '			PlatformId         = info.PlatformId       ;
            OSCSDVersion = info.CSDVersion

            OSSuiteFlags = GetOSSuiteFlags(info.SuiteMask)
            OSProductType = GetOSProductType(info.ProductType)

            OSServicePackMajor = info.ServicePackMajor
            OSServicePackMinor = info.ServicePackMinor
            '			SuiteMask          = info.SuiteMask        ;
            '			ProductType        = info.ProductType      ;
            OSReserved = info.Reserved

            ExtendedPropertiesAreSet = True
        End Sub 'InitOsVersionInfoEx

        Private Shared Function GetOSPlatformId(ByVal platformId As Integer) As CWindows.OSPlatformId
            Select Case platformId
                Case VerPlatformId.Win32s
                    Return CWindows.OSPlatformId.Win32s
                Case VerPlatformId.Win32Windows
                    Return CWindows.OSPlatformId.Win32Windows
                Case VerPlatformId.Win32NT
                    Return CWindows.OSPlatformId.Win32NT
                Case VerPlatformId.WinCE
                    Return CWindows.OSPlatformId.WinCE

                    'Case Else
                    'Throw New InvalidOperationException("Invalid PlatformId: " + platformId)
            End Select
        End Function 'GetOSPlatformId

        Private Shared Function GetOSSuiteFlags(ByVal suiteMask As Int16) As CWindows.OSSuites
            Return CType(suiteMask, CWindows.OSSuites)
        End Function 'GetOSSuiteFlags

        Private Shared Function GetOSProductType(ByVal productType As Byte) As CWindows.OSProductType
            Select Case productType
                Case VerProductType.VER_NT_WORKSTATION
                    Return CWindows.OSProductType.Workstation
                Case VerProductType.VER_NT_DOMAIN_CONTROLLER
                    Return CWindows.OSProductType.DomainController
                Case VerProductType.VER_NT_SERVER
                    Return CWindows.OSProductType.Server
                    'Case Else
                    'Throw New InvalidOperationException("Invalid ProductType: " + productType)
            End Select
        End Function 'GetOSProductType

        Public Function WinDir() As String
            Dim WinSysPath As String = System.Environment.GetFolderPath(Environment.SpecialFolder.System)
            WinDir = WinSysPath.Substring(0, WinSysPath.LastIndexOf("\"))
        End Function

        'Purpose to return the windows system directory
        Public Function WinSysDir() As String
            WinSysDir = System.Environment.SystemDirectory
        End Function

        'Purpose to return the drive type eg NTFS, Fat32 etc
        Public Function SystemType() As String
            Const StringBufferLength As Integer = 255
            Dim lsRootPathName As String = IO.Directory.GetDirectoryRoot(Application.StartupPath)
            Dim lsFileSystemNameBuffer As New StringBuilder(StringBufferLength)
            GetVolumeInformation(lsRootPathName, Nothing, Nothing, Nothing, Nothing, Nothing, lsFileSystemNameBuffer, Nothing)
            Return lsFileSystemNameBuffer.ToString
        End Function

        'Purpose to send a true value if drive type is NTFS
        Public Function IsNTFSDrive() As Boolean
            If SystemType.ToString = "NTFS" Then Return True
        End Function

    End Class 'OperatingSystemVersion 
End Namespace 'Common '-----------------------------------------------------------------------------
' OperatingSystemVersion
'-----------------------------------------------------------------------------