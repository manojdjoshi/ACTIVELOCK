Option Strict Off
Option Explicit On 
Imports System
Imports System.net
Imports System.Management
Imports Microsoft.win32
Imports System.Runtime.InteropServices

Module modHardware
	'*   ActiveLock
	'*   Copyright 1998-2002 Nelson Ferraz
	'*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
	'*   All material is the property of the contributing authors.
	'*
	'*   Redistribution and use in source and binary forms, with or without
	'*   modification, are permitted provided that the following conditions are
	'*   met:
	'*
	'*     [o] Redistributions of source code must retain the above copyright
	'*         notice, this list of conditions and the following disclaimer.
	'*
	'*     [o] Redistributions in binary form must reproduce the above
	'*         copyright notice, this list of conditions and the following
	'*         disclaimer in the documentation and/or other materials provided
	'*         with the distribution.
	'*
	'*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
	'*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
	'*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
	'*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
	'*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
	'*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
	'*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
	'*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
	'*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
	'*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
	'*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
	'*
	'*
	'===============================================================================
	' Name: modHardware
	' Purpose: Gets all the hardware signatures of the current machine
	' Date Created:
	' Functions:
	' Properties:
	' Methods:
	' Started: 08.15.2005
    ' Modified: 03.25.2006
	'===============================================================================
	
    '****** SMART DECLARATIONS ******
    Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef LpVersionInformation As OSVERSIONINFO) As Integer
    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByVal lpSecurityAttributes As Integer, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As Integer) As Integer
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
    Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Integer, ByVal dwIoControlCode As Integer, ByRef lpInBuffer As SENDCMDINPARAMS, ByVal nInBufferSize As Integer, ByRef lpOutBuffer As Integer, ByVal nOutBufferSize As Integer, ByRef lpBytesReturned As Integer, ByVal lpOverlapped As Integer) As Integer
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Integer, ByRef Source As Integer, ByVal Length As Integer)
    Public Declare Unicode Function CreateFile2 Lib "kernel32" Alias "CreateFileW" (ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByVal lpSecurityAttributes As IntPtr, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As IntPtr) As IntPtr
    Public Declare Unicode Function CloseHandle2 Lib "kernel32" Alias "CloseHandle" (ByVal hObject As IntPtr) As Boolean
    Public Declare Ansi Function DeviceIoControl2 Lib "kernel32" Alias "DeviceIoControl" (ByVal hDevice As IntPtr, ByVal dwIoControlCode As Integer, ByVal lpInBuffer As IntPtr, ByVal nInBufferSize As Integer, ByVal lpOutBuffer As IntPtr, ByVal nOutBufferSize As Integer, ByRef lpBytesReturned As Integer, ByVal lpOverlapped As IntPtr) As Boolean

    Dim colAttrNames As Collection

    '---------------------------------------------------------------------
    ' The following structure defines the structure of a Drive Attribute
    '---------------------------------------------------------------------
    Public Const NUM_ATTRIBUTE_STRUCTS As Short = 30

    Public Structure DRIVEATTRIBUTE
        Dim bAttrID As Byte ' Identifies which attribute
        Dim wStatusFlags As Short 'Integer ' see bit definitions below
        Dim bAttrValue As Byte ' Current normalized value
        Dim bWorstValue As Byte ' How bad has it ever been?
        <VBFixedArray(5)> Dim bRawValue() As Byte ' Un-normalized value
        Dim bReserved As Byte ' ...
        Public Sub Initialize()
            ReDim bRawValue(5)
        End Sub
    End Structure
    '---------------------------------------------------------------------
    ' The following structure defines the structure of a Warranty Threshold
    ' Obsoleted in ATA4!
    '---------------------------------------------------------------------
    'Public Structure ATTRTHRESHOLD
    '    Dim bAttrID As Byte ' Identifies which attribute
    '    Dim bWarrantyThreshold As Byte ' Triggering value
    '    <VBFixedArray(9)> Dim bReserved() As Byte
    '    Public Sub Initialize()
    '        ReDim bReserved(9)
    '    End Sub
    'End Structure

    '---------------------------------------------------------------------
    ' Valid Attribute IDs
    '---------------------------------------------------------------------
    Public Enum ATTRIBUTE_ID
        ATTR_INVALID = 0
        ATTR_READ_ERROR_RATE = 1
        ATTR_THROUGHPUT_PERF = 2
        ATTR_SPIN_UP_TIME = 3
        ATTR_START_STOP_COUNT = 4
        ATTR_REALLOC_SECTOR_COUNT = 5
        ATTR_READ_CHANNEL_MARGIN = 6
        ATTR_SEEK_ERROR_RATE = 7
        ATTR_SEEK_TIME_PERF = 8
        ATTR_POWER_ON_HRS_COUNT = 9
        ATTR_SPIN_RETRY_COUNT = 10
        ATTR_CALIBRATION_RETRY_COUNT = 11
        ATTR_POWER_CYCLE_COUNT = 12
        ATTR_SOFT_READ_ERROR_RATE = 13
        ATTR_G_SENSE_ERROR_RATE = 191
        ATTR_POWER_OFF_RETRACT_CYCLE = 192
        ATTR_LOAD_UNLOAD_CYCLE_COUNT = 193
        ATTR_TEMPERATURE = 194
        ATTR_REALLOCATION_EVENTS_COUNT = 196
        ATTR_CURRENT_PENDING_SECTOR_COUNT = 197
        ATTR_UNCORRECTABLE_SECTOR_COUNT = 198
        ATTR_ULTRADMA_CRC_ERROR_RATE = 199
        ATTR_WRITE_ERROR_RATE = 200
        ATTR_DISK_SHIFT = 220
        ATTR_G_SENSE_ERROR_RATEII = 221
        ATTR_LOADED_HOURS = 222
        ATTR_LOAD_UNLOAD_RETRY_COUNT = 223
        ATTR_LOAD_FRICTION = 224
        ATTR_LOAD_UNLOAD_CYCLE_COUNTII = 225
        ATTR_LOAD_IN_TIME = 226
        ATTR_TORQUE_AMPLIFICATION_COUNT = 227
        ATTR_POWER_OFF_RETRACT_COUNT = 228
        ATTR_GMR_HEAD_AMPLITUDE = 230
        ATTR_TEMPERATUREII = 231
        ATTR_READ_ERROR_RETRY_RATE = 250
    End Enum

    '***** SMART DECLARATIONS *****

    'HDD firmware serial number
    Private Const GENERIC_READ As Integer = &H80000000
    Private Const GENERIC_WRITE As Integer = &H40000000
    Private Const FILE_SHARE_READ As Short = &H1S
    Private Const FILE_SHARE_WRITE As Short = &H2S
    Private Const OPEN_EXISTING As Short = 3
    Private Const CREATE_NEW As Short = 1
    Private Const INVALID_HANDLE_VALUE As Short = -1
    Private Const VER_PLATFORM_WIN32_NT As Short = 2
    Private Const IDENTIFY_BUFFER_SIZE As Short = 512
    Public Const READ_THRESHOLD_BUFFER_SIZE As Short = 512
    Private Const OUTPUT_DATA_SIZE As Integer = IDENTIFY_BUFFER_SIZE + 16

    Public Declare Function apiGetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, ByRef lpVolumeSerialNumber As Integer, ByRef lpMaximumComponentLength As Integer, ByRef lpFileSystemFlags As Integer, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Integer) As Integer

    'GETVERSIONOUTPARAMS contains the data returned
    'from the Get Driver Version function
    Private Structure GETVERSIONOUTPARAMS
        Dim bVersion As Byte 'Binary driver version.
        Dim bRevision As Byte 'Binary driver revision
        Dim bReserved As Byte 'Not used
        Dim bIDEDeviceMap As Byte 'Bit map of IDE devices
        Dim fCapabilities As Integer 'Bit mask of driver capabilities
        <VBFixedArray(3)> Dim dwReserved() As Integer 'For future use
        Public Sub Initialize()
            ReDim dwReserved(3)
        End Sub
    End Structure

    'IDE registers
    Private Structure IDEREGS
        Dim bFeaturesReg As Byte 'Used for specifying SMART "commands"
        Dim bSectorCountReg As Byte 'IDE sector count register
        Dim bSectorNumberReg As Byte 'IDE sector number register
        Dim bCylLowReg As Byte 'IDE low order cylinder value
        Dim bCylHighReg As Byte 'IDE high order cylinder value
        Dim bDriveHeadReg As Byte 'IDE drive/head register
        Dim bCommandReg As Byte 'Actual IDE command
        Dim bReserved As Byte 'reserved for future use - must be zero
    End Structure

    'SENDCMDINPARAMS contains the input parameters for the
    'Send Command to Drive function
    Private Structure SENDCMDINPARAMS
        Dim cBufferSize As Integer 'Buffer size in bytes
        Dim irDriveRegs As IDEREGS 'Structure with drive register values.
        Dim bDriveNumber As Byte 'Physical drive number to send command to (0,1,2,3).
        <VBFixedArray(2)> Dim bReserved() As Byte 'Bytes reserved
        <VBFixedArray(3)> Dim dwReserved() As Integer 'DWORDS reserved
        Dim bBuffer() As Byte 'Input buffer.
        Public Sub Initialize()
            ReDim bReserved(2)
            ReDim dwReserved(3)
        End Sub
    End Structure

    'Valid values for the bCommandReg member of IDEREGS.
    Private Const IDE_ATAPI_ID As Short = &HA1S ' Returns ID sector for ATAPI.
    Private Const IDE_ID_FUNCTION As Short = &HECS 'Returns ID sector for ATA.
    Private Const IDE_EXECUTE_SMART_FUNCTION As Short = &HB0S 'Performs SMART cmd.
    'Requires valid bFeaturesReg,
    'bCylLowReg, and bCylHighReg

    'Cylinder register values required when issuing SMART command
    Private Const SMART_CYL_LOW As Short = &H4FS
    Private Const SMART_CYL_HI As Short = &HC2S

    'Status returned from driver
    Private Structure DRIVERSTATUS
        Dim bDriverError As Byte 'Error code from driver, or 0 if no error
        Dim bIDEStatus As Byte 'Contents of IDE Error register
        'Only valid when bDriverError is SMART_IDE_ERROR
        <VBFixedArray(1)> Dim bReserved() As Byte
        <VBFixedArray(1)> Dim dwReserved() As Integer
        Public Sub Initialize()
            ReDim bReserved(1)
            ReDim dwReserved(1)
        End Sub
    End Structure

    Private Structure IDSECTOR
        Dim wGenConfig As Short
        Dim wNumCyls As Short
        Dim wReserved As Short
        Dim wNumHeads As Short
        Dim wBytesPerTrack As Short
        Dim wBytesPerSector As Short
        Dim wSectorsPerTrack As Short
        <VBFixedArray(2)> Dim wVendorUnique() As Short
        <VBFixedArray(19)> Dim sSerialNumber() As Byte
        Dim wBufferType As Short
        Dim wBufferSize As Short
        Dim wECCSize As Short
        <VBFixedArray(7)> Dim sFirmwareRev() As Byte
        <VBFixedArray(39)> Dim sModelNumber() As Byte
        Dim wMoreVendorUnique As Short
        Dim wDoubleWordIO As Short
        Dim wCapabilities As Short
        Dim wReserved1 As Short
        Dim wPIOTiming As Short
        Dim wDMATiming As Short
        Dim wBS As Short
        Dim wNumCurrentCyls As Short
        Dim wNumCurrentHeads As Short
        Dim wNumCurrentSectorsPerTrack As Short
        Dim ulCurrentSectorCapacity As Integer
        Dim wMultSectorStuff As Short
        Dim ulTotalAddressableSectors As Integer
        Dim wSingleWordDMA As Short
        Dim wMultiWordDMA As Short
        <VBFixedArray(127)> Dim bReserved() As Byte
        Public Sub Initialize()
            ReDim wVendorUnique(2)
            ReDim sSerialNumber(19)
            ReDim sFirmwareRev(7)
            ReDim sModelNumber(39)
            ReDim bReserved(127)
        End Sub
    End Structure

    'Structure returned by SMART IOCTL commands
    Private Structure SENDCMDOUTPARAMS
        Dim cBufferSize As Integer 'Size of Buffer in bytes
        Dim DRIVERSTATUS As DRIVERSTATUS 'Driver status structure
        Dim bBuffer() As Byte 'Buffer of arbitrary length for data read from drive
        Public Sub Initialize()
            DRIVERSTATUS.Initialize()
        End Sub
    End Structure

    'Vendor specific feature register defines
    'for SMART "sub commands"
    Private Const SMART_READ_ATTRIBUTE_VALUES As Short = &HD0S
    Private Const SMART_READ_ATTRIBUTE_THRESHOLDS As Short = &HD1S
    Private Const SMART_ENABLE_DISABLE_ATTRIBUTE_AUTOSAVE As Short = &HD2S
    Private Const SMART_SAVE_ATTRIBUTE_VALUES As Short = &HD3S
    Private Const SMART_EXECUTE_OFFLINE_IMMEDIATE As Short = &HD4S
    ' Vendor specific commands:
    Private Const SMART_ENABLE_SMART_OPERATIONS As Short = &HD8S
    Private Const SMART_DISABLE_SMART_OPERATIONS As Short = &HD9S
    Private Const SMART_RETURN_SMART_STATUS As Short = &HDAS

    'Status Flags Values
    Public Enum STATUS_FLAGS
        PRE_FAILURE_WARRANTY = &H1S
        ON_LINE_COLLECTION = &H2S
        PERFORMANCE_ATTRIBUTE = &H4S
        ERROR_RATE_ATTRIBUTE = &H8S
        EVENT_COUNT_ATTRIBUTE = &H10S
        SELF_PRESERVING_ATTRIBUTE = &H20S
    End Enum

    'IOCTL commands
    Private Const DFP_GET_VERSION As Integer = &H74080
    Private Const DFP_SEND_DRIVE_COMMAND As Integer = &H7C084
    Private Const DFP_RECEIVE_DRIVE_DATA As Integer = &H7C088

    Public Structure ATTR_DATA
        Dim AttrID As Byte
        Dim AttrName As String
        Dim AttrValue As Byte
        Dim ThresholdValue As Byte
        Dim WorstValue As Byte
        Dim StatusFlags As STATUS_FLAGS
    End Structure

    Private Structure OSVERSIONINFO
        Dim dwOSVersionInfoSize As Integer
        Dim dwMajorVersion As Integer
        Dim dwMinorVersion As Integer
        Dim dwBuildNumber As Integer
        Dim dwPlatformId As Integer
        <VBFixedString(128), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=128)> Public szCSDVersion As String
    End Structure

    Public Structure DRIVE_INFO
        Dim bDriveType As Byte
        Dim SerialNumber As String
        Dim Model As String
        Dim FirmWare As String
        Dim Cilinders As Integer
        Dim Heads As Integer
        Dim SecPerTrack As Integer
        Dim BytesPerSector As Integer
        Dim BytesperTrack As Integer
        Dim NumAttributes As Byte
        Dim Attributes() As ATTR_DATA
    End Structure
    Dim di As DRIVE_INFO

    Public Enum IDE_DRIVE_NUMBER
        PRIMARY_MASTER
        PRIMARY_SLAVE
        SECONDARY_MASTER
        SECONDARY_SLAVE
        TERTIARY_MASTER
        TERTIARY_SLAVE
        QUARTIARY_MASTER
        QUARTIARY_SLAVE
    End Enum
    Private Structure BufferType
        <VBFixedArray(559)> Dim myBuffer() As Byte
        Public Sub Initialize()
            ReDim myBuffer(559)
        End Sub
    End Structure
    Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Integer, ByVal numBytes As Integer)
    Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Integer, ByVal dwIoControlCode As Integer, ByRef lpInBuffer As Object, ByVal nInBufferSize As Integer, ByRef lpOutBuffer As Object, ByVal nOutBufferSize As Integer, ByRef lpBytesReturned As Integer, ByRef lpOverlapped As Integer) As Integer

    ' The following UDT and the DLL function is for getting
    ' the serial number from a C++ DLL in case the VB6 APIs fail
    ' Currently, VB code cannot handle the serial numbers
    ' coming from computers with non-admin rights; in those
    ' cases the C++ DLL function "getHardDriveFirmware" should
    ' work properly.
    ' Neither of the two methods work for the SATA and SCSI drives
    ' ialkan - 8312005
    Private Structure MyUDT2
        <VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=30)> Public myStr As String
        Dim mL As Integer
    End Structure
    Private Declare Function getHardDriveFirmware Lib "ALCrypto3.dll" (ByRef myU As MyUDT2) As Integer

    'MAC Address
    Public Const NCBASTAT As Integer = &H33S
    Public Const NCBNAMSZ As Integer = 16
    Public Const HEAP_ZERO_MEMORY As Integer = &H8S
    Public Const HEAP_GENERATE_EXCEPTIONS As Integer = &H4S
    Public Const NCBRESET As Integer = &H32S

    Public Structure NET_CONTROL_BLOCK 'NCB
        Dim ncb_command As Byte
        Dim ncb_retcode As Byte
        Dim ncb_lsn As Byte
        Dim ncb_num As Byte
        Dim ncb_buffer As Integer
        Dim ncb_length As Short
        <VBFixedString(NCBNAMSZ), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=NCBNAMSZ)> Public ncb_callname As String
        <VBFixedString(NCBNAMSZ), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=NCBNAMSZ)> Public ncb_name As String
        Dim ncb_rto As Byte
        Dim ncb_sto As Byte
        Dim ncb_post As Integer
        Dim ncb_lana_num As Byte
        Dim ncb_cmd_cplt As Byte
        <VBFixedArray(9)> Dim ncb_reserve() As Byte ' Reserved, must be 0
        Dim ncb_event As Integer
        Public Sub Initialize()
            ReDim ncb_reserve(9)
        End Sub
    End Structure

    Public Structure ADAPTER_STATUS
        <VBFixedArray(5)> Dim adapter_address() As Byte
        Dim rev_major As Byte
        Dim reserved0 As Byte
        Dim adapter_type As Byte
        Dim rev_minor As Byte
        Dim duration As Short
        Dim frmr_recv As Short
        Dim frmr_xmit As Short
        Dim iframe_recv_err As Short
        Dim xmit_aborts As Short
        Dim xmit_success As Integer
        Dim recv_success As Integer
        Dim iframe_xmit_err As Short
        Dim recv_buff_unavail As Short
        Dim t1_timeouts As Short
        Dim ti_timeouts As Short
        Dim Reserved1 As Integer
        Dim free_ncbs As Short
        Dim max_cfg_ncbs As Short
        Dim max_ncbs As Short
        Dim xmit_buf_unavail As Short
        Dim max_dgram_size As Short
        Dim pending_sess As Short
        Dim max_cfg_sess As Short
        Dim max_sess As Short
        Dim max_sess_pkt_size As Short
        Dim name_count As Short
        Public Sub Initialize()
            ReDim adapter_address(5)
        End Sub
    End Structure

    Public Structure NAME_BUFFER
        <VBFixedString(NCBNAMSZ), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=NCBNAMSZ)> Public name As String
        Dim name_num As Short
        Dim name_flags As Short
    End Structure

    Public Structure ASTAT
        Dim adapt As ADAPTER_STATUS
        <VBFixedArray(30)> Dim NameBuff() As NAME_BUFFER
        Public Sub Initialize()
            adapt.Initialize()
            ReDim NameBuff(30)
        End Sub
    End Structure

    'Structure NET_CONTROL_BLOCK may require marshalling attributes to be passed as an argument in this Declare statement
    Public Declare Function Netbios Lib "netapi32.dll" (ByRef pncb As NET_CONTROL_BLOCK) As Byte
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef hpvDest As Object, ByVal hpvSource As Integer, ByVal cbCopy As Integer)
    Public Declare Function GetProcessHeap Lib "kernel32" () As Integer
    Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByVal dwBytes As Integer) As Integer
    Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Integer, ByVal dwFlags As Integer, ByVal lpMem As Integer) As Integer
    '===============================================================================
    ' Name: Function GetComputerName
    ' Input: None
    ' Output:
    '   String - Computer name
    ' Purpose: Gets the computer name on the network
    '===============================================================================
    Public Function GetComputerName() As String
        On Error GoTo GetComputerNameError
        GetComputerName = System.Environment.ExpandEnvironmentVariables("%ComputerName%")
GetComputerNameError:
        If GetComputerName = "" Then
            GetComputerName = "Not Available"
        End If

    End Function

    '===============================================================================
    ' Name: Function HDSerial
    ' Input:
    '   ByRef path As String - Drive letter
    ' Output:
    '   String - The serial number for the drive alock is on, formatted as "xxxx-xxxx"
    ' Purpose: Function to return the serial number for a hard drive
    ' Currently works on local drives, mapped drives, and shared drives.
    ' Remarks: TODO: Decide what to to about shared folders and RAID arrays
    '===============================================================================
    Public Function HDSerial(ByRef path As String) As String

        Dim lngDummy2, lngReturn, lngDummy1, lngSerial As Integer
        Dim strDummy2, strDummy1, strSerial As String

        Dim strDriveLetter As String
        Dim lngFirstSlash As Integer

        strDriveLetter = path

        '(Just in case... It's better to be safe than sorry.)
        strDriveLetter = Replace_Renamed(strDriveLetter, "/", "\")

        'Check the drive type
        If Not Left(strDriveLetter, 1) = "\" Then
            'Good... The path is a local drive
            strDriveLetter = Left(strDriveLetter, 3)
        Else
            'It's a network drive
            'This will return 0000-0000 on shared folders or RAID arrays
            'Shared drives work fine
            lngFirstSlash = InStr(3, strDriveLetter, "\")
            strDriveLetter = Left(strDriveLetter, lngFirstSlash)
        End If

        'Set up dimmies
        strDummy1 = Space(260)
        strDummy2 = Space(260)

        'Call the API function
        lngReturn = apiGetVolumeInformation(strDriveLetter, strDummy1, Len(strDummy1), lngSerial, lngDummy1, lngDummy2, strDummy2, Len(strDummy2))

        'Format the serial
        strSerial = Trim(Hex(lngSerial))
        strSerial = New String("0", 8 - Len(strSerial)) & strSerial
        strSerial = Left(strSerial, 4) & "-" & Right(strSerial, 4)
        HDSerial = strSerial

        ' Alternative Code - Short Version
        '    Dim volbuf$, sysname$, sysflags&, componentlength&
        '    Dim serialnum As Long
        '    GetVolumeInformation "C:\", volbuf$, 255, serialnum, componentlength, sysflags, sysname$, 255
        '    HDSerial = CStr(serialnum)

    End Function

    '===============================================================================
    ' Name: Function GetHDSerial
    ' Input: None
    ' Output:
    '   String - The serial number for the drive alock is on, formatted as "xxxx-xxxx"
    ' Purpose: Function to return the serial number for a hard drive
    '   Currently works on local drives, mapped drives, and shared drives.
    '   Checks windir if it cant get a serial, then c:, then returns 0000-0000
    ' Remarks: I think that this is 99.999999897456284893% effective.
    '===============================================================================
    Public Function GetHDSerial() As String
        Dim strSerial As String
        strSerial = HDSerial(Microsoft.VisualBasic.Compatibility.VB6.GetPath)

        If strSerial = "0000-0000" Then
            'Calculate WINDIR drive if couldn't retrieve app.path serial
            strSerial = HDSerial(WinDir())
        End If

        If strSerial = "0000-0000" Then
            'If it still can't get a serial, revert to c:.
            'If no c: is present (or c: is RAID), 0000-0000 is returned
            strSerial = HDSerial("C:\")
        End If

        GetHDSerial = strSerial

GetHDSeriAlerror:
        If GetHDSerial = "" Then
            GetHDSerial = "Not Available"
        End If
    End Function

    '===============================================================================
    ' Name: Function GetHDSerialFirmware
    ' Input: None
    ' Output:
    '   String - HDD Firmware Serial Number
    ' Purpose: Function to return the HDD Firmware Serial Number (Actual Physical Serial Number)
    ' Remarks: None
    '===============================================================================
    Function GetHDSerialFirmware() As String
        Dim jj As Short
        On Error GoTo GetHDSerialFirmwareError

        Dim drvNumber As Integer

        ' We just need the Primary Master Drive ID - ialkan 8312005
        Dim mU As MyUDT2 = Nothing
        Dim a As String

        GetHDSerialFirmware = Trim(GetDriveInfo(IDE_DRIVE_NUMBER.PRIMARY_MASTER))
        If GetHDSerialFirmware <> "" Then
            Exit Function
        End If

        ' ialkan 2-12-06
        ' Pure VB6 version of the code found in several online resources
        ' described in GetHDSerialFirmwareVB6 function
        ' This eliminates the dependency of the HDD firmware serial number
        ' function from ALCrypto3.dll
        For jj = 0 To 15 ' Controller index
            GetHDSerialFirmware = Trim(GetHDSerialFirmwareVBNET(jj, True)) ' Check the Master drive
            If GetHDSerialFirmware <> "" Then
                Exit Function
            End If
            GetHDSerialFirmware = Trim(GetHDSerialFirmwareVBNET(jj, False)) ' Now check the Slave Drive
            If GetHDSerialFirmware <> "" Then
                Exit Function
            End If
        Next

        ' Still nothing... Use ALCrypto DLL
        Call getHardDriveFirmware(mU)
        GetHDSerialFirmware = Trim(StripControlChars(mU.myStr, False))
        If GetHDSerialFirmware <> "" Then Exit Function

        Exit Function

        ' Well, this is not so good, because we still don't have
        ' a serial number in our hands...
        ' Cannot return an empty string...
GetHDSerialFirmwareError:
        If GetHDSerialFirmware = "" Then
            GetHDSerialFirmware = "Not Available"
        End If

    End Function

    '===============================================================================
    ' Name: Function StripControlChars
    ' Input:
    '   ByVal source As String - String to be stripped off the control characters
    '   ByVal KeepCRLF As Boolean - If the second argument is True or omitted, CR-LF pairs are preserved
    ' Output:
    '   String - String stripped off the control characters
    ' Purpose: Strips all control characters (ASCII code < 32)
    ' Remarks: None
    '===============================================================================
    Function StripControlChars(ByRef Source As String, Optional ByRef KeepCRLF As Boolean = True) As String
        Dim Index As Integer
        Dim bytes() As Byte

        ' the fastest way to process this string
        ' is copy it into an array of Bytes
        bytes = System.Text.UnicodeEncoding.Unicode.GetBytes(Source)
        For Index = 0 To UBound(bytes) Step 2
            ' if this is a control character
            If bytes(Index) < 32 And bytes(Index + 1) = 0 Then
                If Not KeepCRLF Or (bytes(Index) <> 13 And bytes(Index) <> 10) Then
                    ' the user asked to trim CRLF or this
                    ' character isn't a CR or a LF, so clear it
                    bytes(Index) = 0
                End If
            End If
        Next

        ' return this string, after filtering out all null chars
        StripControlChars = Replace(System.Text.UnicodeEncoding.Unicode.GetString(bytes), vbNullChar, "")
    End Function

    '***************************************************************************
    ' Open SMART to allow DeviceIoControl communications. Return SMART handle
    '***************************************************************************
    Private Function OpenSmart(ByRef drv_num As IDE_DRIVE_NUMBER) As Integer
        If IsWindowsNT() Then
            OpenSmart = CreateFile("\\.\PhysicalDrive" & CStr(drv_num), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
        Else
            OpenSmart = CreateFile("\\.\SMARTVSD", 0, 0, 0, CREATE_NEW, 0, 0)
        End If
    End Function


    '****************************************************************************
    ' ReadAttributesCmd
    ' FUNCTION: Send a SMART_READ_ATTRIBUTE_VALUES command to the drive
    ' bDriveNum = 0-3
    '***************************************************************************}
    Private Function ReadAttributesCmd(ByVal hDrive As Integer, ByRef DriveNum As IDE_DRIVE_NUMBER) As Boolean
        Dim READ_ATTRIBUTE_BUFFER_SIZE As Object = Nothing
        Dim cbBytesReturned As Integer
        Dim SCIP As SENDCMDINPARAMS = Nothing
        Dim drv_attr As DRIVEATTRIBUTE
        Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
        Dim sMsg As String
        Dim i As Integer
        With SCIP
            ' Set up data structures for Read Attributes SMART Command.
            .cBufferSize = READ_ATTRIBUTE_BUFFER_SIZE
            .bDriveNumber = DriveNum
            With .irDriveRegs
                .bFeaturesReg = SMART_READ_ATTRIBUTE_VALUES
                .bSectorCountReg = 1
                .bSectorNumberReg = 1
                .bCylLowReg = SMART_CYL_LOW
                .bCylHighReg = SMART_CYL_HI
                '  Compute the drive number.
                .bDriveHeadReg = &HA0S
                If Not IsWindowsNT() Then .bDriveHeadReg = .bDriveHeadReg Or CShort(DriveNum And 1) * 16
                .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
            End With
        End With
        ReadAttributesCmd = DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, cbBytesReturned, 0)
        On Error Resume Next
        For i = 0 To NUM_ATTRIBUTE_STRUCTS - 1
            If bArrOut(18 + i * 12) > 0 Then
                di.Attributes(di.NumAttributes).AttrID = bArrOut(18 + i * 12)
                di.Attributes(di.NumAttributes).AttrName = "Unknown value (" & bArrOut(18 + i * 12) & ")"
                di.Attributes(di.NumAttributes).AttrName = colAttrNames.Item(CStr(bArrOut(18 + i * 12)))
                di.NumAttributes = di.NumAttributes + 1
                ReDim Preserve di.Attributes(di.NumAttributes)
                CopyMemory(di.Attributes(di.NumAttributes).StatusFlags, bArrOut(19 + i * 12), 2)
                di.Attributes(di.NumAttributes).AttrValue = bArrOut(21 + i * 12)
                di.Attributes(di.NumAttributes).WorstValue = bArrOut(22 + i * 12)
            End If
        Next i
    End Function

    Private Function IsWindowsNT() As Boolean
        'Dim verinfo As OSVERSIONINFO
        'verinfo.dwOSVersionInfoSize = Len(verinfo)
        'If (GetVersionEx(verinfo)) = 0 Then Exit Function
        'If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
        Dim MyHost As New CWindows.OperatingSystemVersion
        IsWindowsNT = MyHost.IsWinNT4Plus
    End Function

    Private Function IsBitSet(ByRef iBitString As Byte, ByVal lBitNo As Short) As Boolean
        If lBitNo = 7 Then
            IsBitSet = iBitString < 0
        Else
            IsBitSet = iBitString And (2 ^ lBitNo)
        End If
    End Function

    Private Function SwapStringBytes(ByVal sIn As String) As String
        Dim sTemp As String
        Dim i As Short
        sTemp = Space(Len(sIn))
        For i = 1 To Len(sIn) - 1 Step 2
            Mid(sTemp, i, 1) = Mid(sIn, i + 1, 1)
            Mid(sTemp, i + 1, 1) = Mid(sIn, i, 1)
        Next i
        SwapStringBytes = sTemp
    End Function

    Public Sub FillAttrNameCollection()
        colAttrNames = New Collection
        With colAttrNames
            .Add("ATTR_INVALID", "0")
            .Add("READ_ERROR_RATE", "1")
            .Add("THROUGHPUT_PERF", "2")
            .Add("SPIN_UP_TIME", "3")
            .Add("START_STOP_COUNT", "4")
            .Add("REALLOC_SECTOR_COUNT", "5")
            .Add("READ_CHANNEL_MARGIN", "6")
            .Add("SEEK_ERROR_RATE", "7")
            .Add("SEEK_TIME_PERF", "8")
            .Add("POWER_ON_HRS_COUNT", "9")
            .Add("SPIN_RETRY_COUNT", "10")
            .Add("CALIBRATION_RETRY_COUNT", "11")
            .Add("POWER_CYCLE_COUNT", "12")
            .Add("SOFT_READ_ERROR_RATE", "13")
            .Add("G_SENSE_ERROR_RATE", "191")
            .Add("POWER_OFF_RETRACT_CYCLE", "192")
            .Add("LOAD_UNLOAD_CYCLE_COUNT", "193")
            .Add("TEMPERATURE", "194")
            .Add("REALLOCATION_EVENTS_COUNT", "196")
            .Add("CURRENT_PENDING_SECTOR_COUNT", "197")
            .Add("UNCORRECTABLE_SECTOR_COUNT", "198")
            .Add("ULTRADMA_CRC_ERROR_RATE", "199")
            .Add("WRITE_ERROR_RATE", "200")
            .Add("DISK_SHIFT", "220")
            .Add("G_SENSE_ERROR_RATEII", "221")
            .Add("LOADED_HOURS", "222")
            .Add("LOAD_UNLOAD_RETRY_COUNT", "223")
            .Add("LOAD_FRICTION", "224")
            .Add("LOAD_UNLOAD_CYCLE_COUNTII", "225")
            .Add("LOAD_IN_TIME", "226")
            .Add("TORQUE_AMPLIFICATION_COUNT", "227")
            .Add("POWER_OFF_RETRACT_COUNT", "228")
            .Add("GMR_HEAD_AMPLITUDE", "230")
            .Add("TEMPERATUREII", "231")
            .Add("READ_ERROR_RETRY_RATE", "250")
        End With
    End Sub


    '===============================================================================
    ' Name: Function SmartGetVersion
    ' Input:
    '   ByVal hDrive As Long - SMART drive handle
    ' Output:
    '   Boolean - True if successful
    ' Purpose: Given the SMART drive handle, gets the version
    ' Remarks: None
    '===============================================================================
    Private Function SmartGetVersion(ByVal hDrive As Integer) As Boolean

        Dim cbBytesReturned As Integer
        'Arrays in structure GVOP may need to be initialized before they can be used
        Dim GVOP As GETVERSIONOUTPARAMS = Nothing

        SmartGetVersion = DeviceIoControl(hDrive, DFP_GET_VERSION, 0, 0, GVOP, Len(GVOP), cbBytesReturned, 0)

    End Function
    '===============================================================================
    ' Name: Function SwapBytes
    ' Input:
    '   ByRef b As Byte - Input byte array
    ' Output:
    '   Byte - Swapped byte array
    ' Purpose: Swaps byte arrays
    ' Remarks: None
    '===============================================================================
    Private Function SwapBytes(ByRef b() As Byte) As Byte()

        'Note: VB4-32 and VB5 do not support the
        'return of arrays from a function. For
        'developers using these VB versions there
        'are two workarounds to this restriction:
        '
        '1) Change the return data type ( As Byte() )
        '   to As Variant (no brackets). No change
        '   to the calling code is required.
        '
        '2) Change the function to a sub, remove
        '   the last line of code (SwapBytes = b()),
        '   and take advantage of the fact the
        '   original byte array is being passed
        '   to the function ByRef, therefore any
        '   changes made to the passed data are
        '   actually being made to the original data.
        '   With this workaround the calling code
        '   also requires modification:
        '
        '      di.Model = StrConv(SwapBytes(IDSEC.sModelNumber), vbUnicode)
        '
        '   ... to ...
        '
        '      Call SwapBytes(IDSEC.sModelNumber)
        '      di.Model = StrConv(IDSEC.sModelNumber, vbUnicode)

        Dim bTemp As Byte
        Dim cnt As Integer

        For cnt = LBound(b) To UBound(b) Step 2
            bTemp = b(cnt)
            b(cnt) = b(cnt + 1)
            b(cnt + 1) = bTemp
        Next cnt

        SwapBytes = VB6.CopyArray(b)

    End Function
    '===============================================================================
    ' Name: Function GetMACAddress
    ' Input: None
    ' Output:
    '   String - MAC address of the computer NIC
    ' Purpose: Retrieves the MAC Address for the network controller installed, returning a formatted string
    ' Remarks: None
    '===============================================================================
    Public Function GetMACAddress() As String
        GetMACAddress = String.Empty
        'On Error Resume Next
        'Dim tmp As String
        'Dim pASTAT As Integer
        'Dim NCB As NET_CONTROL_BLOCK
        'Dim AST As ASTAT

        ''The IBM NetBIOS 3.0 specifications defines four basic
        ''NetBIOS environments under the NCBRESET command. Win32
        ''follows the OS/2 Dynamic Link Routine (DLR) environment.
        ''This means that the first NCB issued by an application
        ''must be a NCBRESET, with the exception of NCBENUM.
        ''The Windows NT implementation differs from the IBM
        ''NetBIOS 3.0 specifications in the NCB_CALLNAME field.
        'NCB.ncb_command = NCBRESET
        'Call Netbios(NCB)

        ''To get the Media Access Control (MAC) address for an
        ''ethernet adapter programmatically, use the Netbios()
        ''NCBASTAT command and provide a "*" as the name in the
        ''NCB.ncb_CallName field (in a 16-chr string).
        'NCB.ncb_callname = "*               "
        'NCB.ncb_command = NCBASTAT

        ''For machines with multiple network adapters you need to
        ''enumerate the LANA numbers and perform the NCBASTAT
        ''command on each. Even when you have a single network
        ''adapter, it is a good idea to enumerate valid LANA numbers
        ''first and perform the NCBASTAT on one of the valid LANA
        ''numbers. It is considered bad programming to hardcode the
        ''LANA number to 0 (see the comments section below).
        'NCB.ncb_lana_num = 0
        'NCB.ncb_length = Len(AST)

        'pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, NCB.ncb_length)

        'If pASTAT = 0 Then
        '    System.Diagnostics.Debug.WriteLine("memory allocation failed!")
        '    Exit Function
        'End If

        'NCB.ncb_buffer = pASTAT
        'Call Netbios(NCB)

        'CopyMemory(AST, NCB.ncb_buffer, Len(AST))

        'tmp = Right("00" & Hex(AST.adapt.adapter_address(0)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(1)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(2)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(3)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(4)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(5)), 2)

        'HeapFree(GetProcessHeap(), 0, pASTAT)

        ''GetMACAddress = Replace(tmp, " ", "")
        'GetMACAddress = tmp

        Dim netInfo As New clsNetworkStats
        Dim netStruct As clsNetworkStats.IFROW_HELPER
        netStruct = netInfo.GetAdapter
        GetMACAddress = netStruct.PhysAddr.ToString().Replace("-", " ")
        If GetMACAddress <> "00 00 00 00 00 00" Then Exit Function

        ' Here we are assuming that the user is NOT running .NET in a Win98/Me machine...
        Dim mc As ManagementClass
        Dim mo As ManagementObject
        mc = New ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim moc As ManagementObjectCollection = mc.GetInstances()
        For Each mo In moc
            If mo.Item("IPEnabled").ToString() = "True" Then
                Return mo.Item("MacAddress").ToString().Replace(":", " ")
            End If
        Next

        'another user provided the code below that seems to work well
        'if the adapter card is "Ethernet 802.3", then the code below will work
        Dim objset, obj As Object
        objset = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_NetworkAdapter")
        For Each obj In objset
            On Error Resume Next
            If Not IsDBNull(obj.MACAddress) Then
                If obj.AdapterType = "Ethernet 802.3" Then
                    If InStr(obj.PNPDeviceID, "PCI\") <> 0 Then
                        GetMACAddress = Replace_Renamed(obj.MACAddress, ":", " ")
                        Exit Function
                    End If
                End If
            End If
        Next obj

GetMACAddressError:
        If GetMACAddress = "" Then
            GetMACAddress = "Not Available"
        End If

    End Function
    '===============================================================================
    ' Name: Function GetWindowsSerial
    ' Input: None
    ' Output:
    '   String - Windows serial number
    ' Purpose: Gets the Windows Serial Number
    ' Remarks: .NET way of doing things added
    '===============================================================================
    Public Function GetWindowsSerial() As String
        Dim myReg As RegistryKey = Registry.LocalMachine
        Dim MyRegKey As RegistryKey
        MyRegKey = myReg.OpenSubKey("Software\Microsoft\Windows NT\CurrentVersion")
        GetWindowsSerial = MyRegKey.GetValue("ProductID")
        MyRegKey.Close()
    End Function

    '===============================================================================
    ' Name: Function GetBiosVersion
    ' Input: None
    ' Output:
    '   String - BIOS serial number
    ' Purpose: Gets the BIOS Serial Number
    ' Remarks: Uses the WMI
    '===============================================================================
    Public Function GetBiosVersion() As String
        Dim BiosSet As Object
        Dim obj As Object

        GetBiosVersion = String.Empty

        On Error GoTo GetBiosVersionerror
        BiosSet = GetObject("WinMgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_BIOS")
        For Each obj In BiosSet
            GetBiosVersion = obj.Version
            If GetBiosVersion <> "" Then Exit Function
        Next obj
GetBiosVersionerror:
        If GetBiosVersion = "" Then
            GetBiosVersion = "Not Available"
        End If
    End Function

    '===============================================================================
    ' Name: Function GetMotherboardSerial
    ' Input: None
    ' Output:
    '   String - Motherboard serial number
    ' Purpose: Gets the Motherboard Serial Number
    ' Remarks: Uses the WMI
    '===============================================================================
    Public Function GetMotherboardSerial() As String
        Dim MotherboardSet As Object
        Dim obj As Object

        GetMotherboardSerial = String.Empty
        On Error GoTo GetMotherboardSeriAlerror

        MotherboardSet = GetObject("WinMgmts:{impersonationLevel=impersonate}").InstancesOf("CIM_Chassis")
        Dim bytes() As Byte
        For Each obj In MotherboardSet
            GetMotherboardSerial = obj.SerialNumber
            If GetMotherboardSerial <> "" Then
                ' Strip any periods
                bytes = System.Text.UnicodeEncoding.Unicode.GetBytes(GetMotherboardSerial)
                GetMotherboardSerial = Replace(System.Text.UnicodeEncoding.Unicode.GetString(bytes), ".", "")
                Exit Function
            End If
        Next obj

GetMotherboardSeriAlerror:
        If GetMotherboardSerial = "" Then
            GetMotherboardSerial = "Not Available"
        End If
    End Function
    '===============================================================================
    ' Name: Function GetIPaddress
    ' Input: None
    ' Output:
    '   String - IP address
    ' Purpose: Gets the IP address
    ' Remarks:
    '===============================================================================
    Public Function GetIPaddress() As String
        On Error GoTo GetIPaddressError
        Dim ipEntry As IPHostEntry = Dns.GetHostEntry(Environment.MachineName)
        Dim IpAddr As IPAddress() = ipEntry.AddressList
        Dim i As Integer
        'A hostmachine can have more than one IP assigned 
        GetIPaddress = IpAddr(0).ToString()

GetIPaddressError:
        If GetIPaddress = "" Then
            GetIPaddress = "Not Available"
        End If
    End Function
    Private Function GetHDSerialFirmwareVBNET(ByVal controller As Integer, Optional ByVal masterDrive As Boolean = True) As Object

        ' Created with the help of the following articles and clues from ALCrypto3.dll
        ' SOURCE 1: http://discuss.develop.com/archives/wa.exe?A2=ind0309a&L=advanced-dotnet&D=0&T=0&P=3760
        ' SOURCE 2: http://www.visual-basic.it/scarica.asp?ID=611
        ' SOURCE 3: ALCrypto3.dll and DISKID32 program

        ' This code DOES NOT require admin rights in the user's machine
        ' This code DOES NOT require WMI
        ' This code DOES NOT require SMART VXD drivers for Win95/98/Me

        Dim myStr As String = String.Empty
        Dim str1, reversedStr, str2 As String
        Dim jj As Short
        Dim dummy As Integer = 0
        Dim hdh As IntPtr, newHandle As Boolean

        GetHDSerialFirmwareVBNET = ""
        hdh = CreateFile2("\\.\Scsi" & controller.ToString() & ":", GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero)

        Dim bin(559) As Byte
        Dim bout As IntPtr = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(560)
        If (hdh.ToInt32 <> -1) Then
            bin(0) = 28
            bin(4) = 83
            bin(5) = 67
            bin(6) = 83
            bin(7) = 73
            bin(8) = 68
            bin(9) = 73
            bin(10) = 83
            bin(11) = 75
            bin(12) = 16
            bin(13) = 39
            bin(16) = 1
            bin(17) = 5
            bin(18) = 27
            bin(24) = 20 '17?
            bin(25) = 2
            bin(38) = 236 '&HEC

            If masterDrive = True Then
                bin(40) = 0 'master drive
            Else
                bin(40) = 1 'slave drive
            End If

            System.Runtime.InteropServices.Marshal.Copy(bin, 0, bout, 560)
            newHandle = DeviceIoControl2(hdh, 315400, bout, 63, bout, 560, dummy, IntPtr.Zero)

            If newHandle Then
                System.Runtime.InteropServices.Marshal.Copy(bout, bin, 0, 560)
                ' HDD Firmware Serial Number is between 64 to 83 - 19 digits as we had from ALCrypto before
                ' HDD Model Number is between 98 and 137
                ' HDD Controller Revision is between 90 and 97
                For jj = 64 To 83
                    myStr = myStr & Convert.ToString(Convert.ToChar(bin(jj)))
                Next

                ' Seems like some swapping is needed at this point
                reversedStr = ""
                For jj = 0 To Len(myStr) / 2
                    str1 = Mid(myStr, jj * 2 + 1, 1)
                    str2 = Mid(myStr, jj * 2 + 2, 1)
                    reversedStr = reversedStr & str2 & str1
                Next
                GetHDSerialFirmwareVBNET = StripControlChars(Trim(reversedStr), False)
            End If

        End If

        System.Runtime.InteropServices.Marshal.FreeHGlobal(bout)
        CloseHandle2(hdh)

    End Function

    Private Function GetSerialNumberFromWMI(ByVal wmi_selection As String) As String
        GetSerialNumberFromWMI = ""

    End Function
End Module