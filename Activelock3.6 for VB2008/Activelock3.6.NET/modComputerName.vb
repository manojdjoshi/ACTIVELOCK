Option Strict Off
Option Explicit On 
Imports System
Imports System.net
Imports System.Text
Imports System.Management
Imports Microsoft.win32
Imports System.Runtime.InteropServices
Imports System.Security
'Imports System.DirectoryServices ' This isn't VS2005 compatible

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
' *   Copyright 2003-2009 The ActiveLock Software Group (ASG)
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

''' <summary>
''' Gets all the hardware signatures of the current machine
''' </summary>
''' <remarks></remarks>
Module modHardware

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

    ' ****** SMART DECLARATIONS ******
    ''' <summary>
    ''' Retrieves information about the current operating system.
    ''' </summary>
    ''' <param name="LpVersionInformation">[in, out] An OSVERSIONINFO or OSVERSIONINFOEX structure that receives the operating system information.
    ''' <para>Before calling the GetVersionEx function, set the dwOSVersionInfoSize member of the structure as appropriate to indicate which data structure is being passed to this function.</para>
    ''' </param>
    ''' <returns>If the function succeeds, the return value is a nonzero value.
    ''' <para>If the function fails, the return value is zero. To get extended error information, call GetLastError. The function fails if you specify an invalid value for the dwOSVersionInfoSize member of the OSVERSIONINFO or OSVERSIONINFOEX structure.</para>
    ''' </returns>
    ''' <remarks>see http://msdn.microsoft.com/en-us/library/ms724451(VS.85).aspx for more information</remarks>
    Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef LpVersionInformation As OSVERSIONINFO) As Integer
    ''' <summary>
    ''' Creates or opens a file or I/O device. The most commonly used I/O devices are as follows: file, file stream, directory, physical disk, volume, console buffer, tape drive, communications resource, mailslot, and pipe. The function returns a handle that can be used to access the file or device for various types of I/O depending on the file or device and the flags and attributes specified.
    ''' <para>To perform this operation as a transacted operation, which results in a handle that can be used for transacted I/O, use the CreateFileTransacted function</para>
    ''' </summary>
    ''' <param name="lpFileName"></param>
    ''' <param name="dwDesiredAccess"></param>
    ''' <param name="dwShareMode"></param>
    ''' <param name="lpSecurityAttributes"></param>
    ''' <param name="dwCreationDisposition"></param>
    ''' <param name="dwFlagsAndAttributes"></param>
    ''' <param name="hTemplateFile"></param>
    ''' <returns></returns>
    ''' <remarks>see http://msdn.microsoft.com/en-us/library/aa363858(VS.85).aspx for more information</remarks>
    Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByVal lpSecurityAttributes As Integer, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As Integer) As Integer
    ''' <summary>
    ''' Closes an open object handle.
    ''' </summary>
    ''' <param name="hObject">[in] A valid handle to an open object.</param>
    ''' <returns>If the function succeeds, the return value is nonzero.
    ''' <para>If the function fails, the return value is zero. To get extended error information, call GetLastError.</para>
    ''' <para>If the application is running under a debugger, the function will throw an exception if it receives either a handle value that is not valid or a pseudo-handle value. This can happen if you close a handle twice, or if you call CloseHandle on a handle returned by the FindFirstFile function instead of calling the FindClose function.</para>
    ''' </returns>
    ''' <remarks>see http://msdn.microsoft.com/en-us/library/ms724211(VS.85).aspx for more information</remarks>
    Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Integer) As Integer
    ''' <summary>
    ''' Sends a control code directly to a specified device driver, causing the corresponding device to perform the corresponding operation.
    ''' </summary>
    ''' <param name="hDevice">[in] A handle to the device on which the operation is to be performed. The device is typically a volume, directory, file, or stream. To retrieve a device handle, use the CreateFile function. For more information, see Remarks.</param>
    ''' <param name="dwIoControlCode">[in] The control code for the operation. This value identifies the specific operation to be performed and the type of device on which to perform it. 
    ''' <para>For a list of the control codes, see Remarks. The documentation for each control code provides usage details for the lpInBuffer, nInBufferSize, lpOutBuffer, and nOutBufferSize parameters.</para>
    ''' </param>
    ''' <param name="lpInBuffer">[in, optional] A pointer to the input buffer that contains the data required to perform the operation. The format of this data depends on the value of the dwIoControlCode parameter. 
    ''' <para>This parameter can be NULL if dwIoControlCode specifies an operation that does not require input data.</para>
    ''' </param>
    ''' <param name="nInBufferSize">[in] The size of the input buffer, in bytes.</param>
    ''' <param name="lpOutBuffer">[out, optional] A pointer to the output buffer that is to receive the data returned by the operation. The format of this data depends on the value of the dwIoControlCode parameter. 
    ''' <para>This parameter can be NULL if dwIoControlCode specifies an operation that does not return data.</para>
    ''' </param>
    ''' <param name="nOutBufferSize">[in] The size of the output buffer, in bytes.</param>
    ''' <param name="lpBytesReturned">see http://msdn.microsoft.com/en-us/library/aa363216(VS.85).aspx for more information</param>
    ''' <param name="lpOverlapped">see http://msdn.microsoft.com/en-us/library/aa363216(VS.85).aspx for more information</param>
    ''' <returns>If the operation completes successfully, the return value is nonzero.
    ''' <para>If the operation fails or is pending, the return value is zero. To get extended error information, call GetLastError.</para>
    ''' </returns>
    ''' <remarks>see http://msdn.microsoft.com/en-us/library/aa363216(VS.85).aspx for more information</remarks>
    Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Integer, ByVal dwIoControlCode As Integer, ByRef lpInBuffer As SENDCMDINPARAMS, ByVal nInBufferSize As Integer, ByRef lpOutBuffer As Integer, ByVal nOutBufferSize As Integer, ByRef lpBytesReturned As Integer, ByVal lpOverlapped As Integer) As Integer
    ''' <summary>
    ''' Copies a block of memory from one location to another.
    ''' </summary>
    ''' <param name="Destination">[in] A pointer to the starting address of the copied block's destination.</param>
    ''' <param name="Source">[in] A pointer to the starting address of the block of memory to copy.</param>
    ''' <param name="Length">[in] The size of the block of memory to copy, in bytes.</param>
    ''' <remarks>see http://msdn.microsoft.com/en-us/library/aa366535%28VS.85%29.aspx for more information</remarks>
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Integer, ByRef Source As Integer, ByVal Length As Integer)
    ''' <summary>
    ''' Creates or opens a file or I/O device.
    ''' <para>see http://msdn.microsoft.com/en-us/library/aa363858(VS.85).aspx for more information</para>
    ''' </summary>
    ''' <param name="lpFileName"></param>
    ''' <param name="dwDesiredAccess"></param>
    ''' <param name="dwShareMode"></param>
    ''' <param name="lpSecurityAttributes"></param>
    ''' <param name="dwCreationDisposition"></param>
    ''' <param name="dwFlagsAndAttributes"></param>
    ''' <param name="hTemplateFile"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
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

    ' ***** SMART DECLARATIONS *****

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
    ''' <summary>
    ''' contains the data returned from the Get Driver Version function
    ''' </summary>
    ''' <remarks></remarks>
    Private Structure GETVERSIONOUTPARAMS
        ''' <summary>
        ''' Binary driver version.
        ''' </summary>
        ''' <remarks></remarks>
        Dim bVersion As Byte 'Binary driver version.
        ''' <summary>
        ''' Binary driver revision
        ''' </summary>
        ''' <remarks></remarks>
        Dim bRevision As Byte 'Binary driver revision
        ''' <summary>
        ''' Not used
        ''' </summary>
        ''' <remarks></remarks>
        Dim bReserved As Byte 'Not used
        ''' <summary>
        ''' Bit map of IDE devices
        ''' </summary>
        ''' <remarks></remarks>
        Dim bIDEDeviceMap As Byte 'Bit map of IDE devices
        ''' <summary>
        ''' Bit mask of driver capabilities
        ''' </summary>
        ''' <remarks></remarks>
        Dim fCapabilities As Integer 'Bit mask of driver capabilities
        ''' <summary>
        ''' For future use
        ''' </summary>
        ''' <remarks></remarks>
        <VBFixedArray(3)> Dim dwReserved() As Integer 'For future use
        Public Sub Initialize()
            ReDim dwReserved(3)
        End Sub
    End Structure

    'IDE registers
    ''' <summary>
    ''' IDE registers
    ''' </summary>
    ''' <remarks></remarks>
    Private Structure IDEREGS
        ''' <summary>
        ''' Used for specifying SMART "commands"
        ''' </summary>
        ''' <remarks></remarks>
        Dim bFeaturesReg As Byte 'Used for specifying SMART "commands"
        ''' <summary>
        ''' IDE sector count register
        ''' </summary>
        ''' <remarks></remarks>
        Dim bSectorCountReg As Byte 'IDE sector count register
        ''' <summary>
        ''' IDE sector number register
        ''' </summary>
        ''' <remarks></remarks>
        Dim bSectorNumberReg As Byte 'IDE sector number register
        ''' <summary>
        ''' IDE low order cylinder value
        ''' </summary>
        ''' <remarks></remarks>
        Dim bCylLowReg As Byte 'IDE low order cylinder value
        ''' <summary>
        ''' IDE high order cylinder value
        ''' </summary>
        ''' <remarks></remarks>
        Dim bCylHighReg As Byte 'IDE high order cylinder value
        ''' <summary>
        ''' IDE drive/head register
        ''' </summary>
        ''' <remarks></remarks>
        Dim bDriveHeadReg As Byte 'IDE drive/head register
        ''' <summary>
        ''' Actual IDE command
        ''' </summary>
        ''' <remarks></remarks>
        Dim bCommandReg As Byte 'Actual IDE command
        ''' <summary>
        ''' reserved for future use - must be zero
        ''' </summary>
        ''' <remarks></remarks>
        Dim bReserved As Byte 'reserved for future use - must be zero
    End Structure

    'SENDCMDINPARAMS contains the input parameters for the
    'Send Command to Drive function
    ''' <summary>
    ''' contains the input parameters for the Send Command to Drive function
    ''' </summary>
    ''' <remarks></remarks>
    Private Structure SENDCMDINPARAMS
        ''' <summary>
        ''' Buffer size in bytes
        ''' </summary>
        ''' <remarks></remarks>
        Dim cBufferSize As Integer 'Buffer size in bytes
        ''' <summary>
        ''' Structure with drive register values.
        ''' </summary>
        ''' <remarks></remarks>
        Dim irDriveRegs As IDEREGS 'Structure with drive register values.
        ''' <summary>
        ''' Physical drive number to send command to (0,1,2,3).
        ''' </summary>
        ''' <remarks></remarks>
        Dim bDriveNumber As Byte 'Physical drive number to send command to (0,1,2,3).
        ''' <summary>
        ''' Bytes reserved
        ''' </summary>
        ''' <remarks></remarks>
        <VBFixedArray(2)> Dim bReserved() As Byte 'Bytes reserved
        ''' <summary>
        ''' DWORDS reserved
        ''' </summary>
        ''' <remarks></remarks>
        <VBFixedArray(3)> Dim dwReserved() As Integer 'DWORDS reserved
        ''' <summary>
        ''' Input buffer.
        ''' </summary>
        ''' <remarks></remarks>
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
    ''' <summary>
    ''' Status returned from driver
    ''' </summary>
    ''' <remarks></remarks>
    Private Structure DRIVERSTATUS
        ''' <summary>
        ''' Error code from driver, or 0 if no error
        ''' </summary>
        ''' <remarks></remarks>
        Dim bDriverError As Byte 'Error code from driver, or 0 if no error
        ''' <summary>
        ''' Contents of IDE Error register
        ''' </summary>
        ''' <remarks></remarks>
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
    ''' <summary>
    ''' Status Flags Values
    ''' </summary>
    ''' <remarks></remarks>
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
    ''' <summary>
    ''' Fills a block of memory with zeros.
    ''' <para>To avoid any undesired effects of optimizing compilers, use the SecureZeroMemory function.</para>
    ''' </summary>
    ''' <param name="dest">[in] A pointer to the starting address of the block of memory to fill with zeros.</param>
    ''' <param name="numBytes">[in] The size of the block of memory to fill with zeros, in bytes.</param>
    ''' <remarks>see http://msdn.microsoft.com/en-us/library/aa366920(VS.85).aspx for more information</remarks>
    Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef dest As Integer, ByVal numBytes As Integer)
    ''' <summary>
    ''' Sends a control code directly to a specified device driver, causing the corresponding device to perform the corresponding operation.
    ''' </summary>
    ''' <param name="hDevice"></param>
    ''' <param name="dwIoControlCode"></param>
    ''' <param name="lpInBuffer"></param>
    ''' <param name="nInBufferSize"></param>
    ''' <param name="lpOutBuffer"></param>
    ''' <param name="nOutBufferSize"></param>
    ''' <param name="lpBytesReturned"></param>
    ''' <param name="lpOverlapped"></param>
    ''' <returns></returns>
    ''' <remarks>see http://msdn.microsoft.com/en-us/library/aa363216(VS.85).aspx for more information</remarks>
    Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Integer, ByVal dwIoControlCode As Integer, ByRef lpInBuffer As Object, ByVal nInBufferSize As Integer, ByRef lpOutBuffer As Object, ByVal nOutBufferSize As Integer, ByRef lpBytesReturned As Integer, ByRef lpOverlapped As Integer) As Integer

    ' ALCrypto Removal
    '' The following UDT and the DLL function is for getting
    '' the serial number from a C++ DLL in case the VB6 APIs fail
    '' Currently, VB code cannot handle the serial numbers
    '' coming from computers with non-admin rights; in those
    '' cases the C++ DLL function "getHardDriveFirmware" should
    '' work properly.
    '' Neither of the two methods work for the SATA and SCSI drives
    '' ialkan - 8312005
    'Private Structure MyUDT2
    '    <VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=30)> Public myStr As String
    '    Dim mL As Integer
    'End Structure
    'ALCrypto Removal
    'Private Declare Function getHardDriveFirmware Lib "ALCrypto3NET.dll" (ByRef myU As MyUDT2) As Integer

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

    Private WithEvents fp As FingerPrint = New FingerPrint
    '===============================================================================
    ' Name: Function GetComputerName
    ' Input: None
    ' Output:
    '   String - Computer name
    ' Purpose: Gets the computer name on the network
    '===============================================================================
    ''' <summary>
    ''' Gets the computer name on the network
    ''' </summary>
    ''' <returns>Computer name</returns>
    ''' <remarks></remarks>
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
    ''' <summary>
    ''' Function to return the serial number for a hard drive Currently works on local drives, mapped drives, and shared drives.
    ''' </summary>
    ''' <param name="path">String - Drive letter</param>
    ''' <returns>The serial number for the drive alock is on, formatted as "xxxx-xxxx"</returns>
    ''' <remarks>TODO: Decide what to to about shared folders and RAID arrays</remarks>
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
    ''' <summary>
    ''' Function to return the serial number for a hard drive. Currently works on local drives, mapped drives, and shared drives. Checks windir if it cant get a serial, then c:, then returns 0000-0000
    ''' </summary>
    ''' <returns>The serial number for the drive alock is on, formatted as "xxxx-xxxx"</returns>
    ''' <remarks>I think that this is 99.999999897456284893% effective.</remarks>
    Public Function GetHDSerial() As String
        Dim strSerial As String

        strSerial = HDSerial(My.Application.Info.DirectoryPath)

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
    ''' <summary>
    ''' Function to return the HDD Firmware Serial Number (Actual Physical Serial Number)
    ''' </summary>
    ''' <returns>HDD Firmware Serial Number</returns>
    ''' <remarks></remarks>
    Function GetHDSerialFirmware() As String
        Dim jj As Short
        Dim drvNumber As Integer
        ' We just need the Primary Master Drive ID - ialkan 8312005
        'ALCrypto Removal
        'Dim mU As MyUDT2 = Nothing
        Dim a As String
        On Error GoTo GetHDSerialFirmwareError
        GetHDSerialFirmware = ""
        a = ""

        ' ************** METHOD 1 **************
        ' ialkan 2-12-06
        ' Pure VB6 version of the code found in several online resources
        ' described in GetHDSerialFirmwareVB6 function
        ' This eliminates the dependency of the HDD firmware serial number
        ' function from ALCrypto3NET.dll
        For jj = 0 To 15 ' Controller index
            a = GetHDSerialFirmwareVBNET(jj, True) ' Check the Master drive
            If a <> "" Then GetHDSerialFirmware = a.Trim
            If GetHDSerialFirmware <> "" Then
                Exit Function
            End If
            a = GetHDSerialFirmwareVBNET(jj, False) ' Now check the Slave Drive
            If a <> "" Then GetHDSerialFirmware = a.Trim
            If GetHDSerialFirmware <> "" Then
                Exit Function
            End If
        Next

        ' ALCrypto Removal
        '' ************** METHOD 2 **************
        '' Still nothing... Use ALCrypto DLL
        'Call getHardDriveFirmware(mU)
        'If mU.myStr <> "" Then a = StripControlChars(mU.myStr, False)
        'If a <> "" Then GetHDSerialFirmware = a.Trim
        'If GetHDSerialFirmware <> "" Then
        '    Exit Function
        'End If

        ' ************** METHOD 3 **************
        a = GetDriveInfo(IDE_DRIVE_NUMBER.PRIMARY_MASTER)
        If a <> "" Then GetHDSerialFirmware = a.Trim
        If GetHDSerialFirmware <> "" Then
            Exit Function
        End If

        ' ************** METHOD 4 **************
        a = GetHDSerialFirmwareWMI()
        If a <> "" Then GetHDSerialFirmware = a.Trim
        If GetHDSerialFirmware <> "" Then
            Exit Function
        End If

        Exit Function

        ' Well, this is not so good, because we still don't have
        ' a serial number in our hands...
        ' Cannot return an empty string...
GetHDSerialFirmwareError:
        If GetHDSerialFirmware = "" Then
            'GetHDSerialFirmware = "Not Available"
            ' Per suggestion by Jeroen, we must have something decent returned from this
            GetHDSerialFirmware = "NA" & GetHDSerial() & GetMotherboardSerial()   '"Not Available"
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
    ''' <summary>
    ''' Strips all control characters (ASCII code &lt; 32)
    ''' </summary>
    ''' <param name="Source">String to be stripped off the control characters</param>
    ''' <param name="KeepCRLF">If the second argument is True or omitted, CR-LF pairs are preserved</param>
    ''' <returns>String stripped off the control characters</returns>
    ''' <remarks></remarks>
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

    ' ***************************************************************************
    ' Open SMART to allow DeviceIoControl communications. Return SMART handle
    ' ***************************************************************************
    Private Function OpenSmart(ByRef drv_num As IDE_DRIVE_NUMBER) As Integer
        If IsWindowsNT() Then
            OpenSmart = CreateFile("\\.\PhysicalDrive" & CStr(drv_num), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0)
        Else
            OpenSmart = CreateFile("\\.\SMARTVSD", 0, 0, 0, CREATE_NEW, 0, 0)
        End If
    End Function


    ' ****************************************************************************
    ' ReadAttributesCmd
    ' FUNCTION: Send a SMART_READ_ATTRIBUTE_VALUES command to the drive
    ' bDriveNum = 0-3
    ' ***************************************************************************}
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
    ''' <summary>
    ''' Given the SMART drive handle, gets the version
    ''' </summary>
    ''' <param name="hDrive">SMART drive handle</param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
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
    ''' <summary>
    ''' Swaps byte arrays
    ''' </summary>
    ''' <param name="b">Input byte array</param>
    ''' <returns>Swapped byte array</returns>
    ''' <remarks>see code for more information</remarks>
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

        SwapBytes = b   ' VB6.CopyArray(b)

    End Function
    '===============================================================================
    ' Name: Function GetMACAddress
    ' Input: None
    ' Output:
    '   String - MAC address of the computer NIC
    ' Purpose: Retrieves the MAC Address for the network controller installed, returning a formatted string
    ' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' Retrieves the MAC Address for the network controller installed, returning a formatted string
    ''' </summary>
    ''' <returns>MAC address of the computer NIC</returns>
    ''' <remarks></remarks>
    Public Function GetMACAddress() As String

        '' ******* METHOD 1 *******
        '' This was causing problems and therefore was commented out

        ''On Error Resume Next
        ''Dim tmp As String
        ''Dim pASTAT As Integer
        ''Dim NCB As NET_CONTROL_BLOCK
        ''Dim AST As ASTAT

        ' ''The IBM NetBIOS 3.0 specifications defines four basic
        ' ''NetBIOS environments under the NCBRESET command. Win32
        ' ''follows the OS/2 Dynamic Link Routine (DLR) environment.
        ' ''This means that the first NCB issued by an application
        ' ''must be a NCBRESET, with the exception of NCBENUM.
        ' ''The Windows NT implementation differs from the IBM
        ' ''NetBIOS 3.0 specifications in the NCB_CALLNAME field.
        ''NCB.ncb_command = NCBRESET
        ''Call Netbios(NCB)

        ' ''To get the Media Access Control (MAC) address for an
        ' ''ethernet adapter programmatically, use the Netbios()
        ' ''NCBASTAT command and provide a "*" as the name in the
        ' ''NCB.ncb_CallName field (in a 16-chr string).
        ''NCB.ncb_callname = "*               "
        ''NCB.ncb_command = NCBASTAT

        ' ''For machines with multiple network adapters you need to
        ' ''enumerate the LANA numbers and perform the NCBASTAT
        ' ''command on each. Even when you have a single network
        ' ''adapter, it is a good idea to enumerate valid LANA numbers
        ' ''first and perform the NCBASTAT on one of the valid LANA
        ' ''numbers. It is considered bad programming to hardcode the
        ' ''LANA number to 0 (see the comments section below).
        ''NCB.ncb_lana_num = 0
        ''NCB.ncb_length = Len(AST)

        ''pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, NCB.ncb_length)

        ''If pASTAT = 0 Then
        ''    System.Diagnostics.Debug.WriteLine("memory allocation failed!")
        ''    Exit Function
        ''End If

        ''NCB.ncb_buffer = pASTAT
        ''Call Netbios(NCB)

        ''CopyMemory(AST, NCB.ncb_buffer, Len(AST))

        ''tmp = Right("00" & Hex(AST.adapt.adapter_address(0)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(1)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(2)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(3)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(4)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(5)), 2)

        ''HeapFree(GetProcessHeap(), 0, pASTAT)

        ' ''GetMACAddress = Replace(tmp, " ", "")
        ''GetMACAddress = tmp
        'Dim foundOne As Boolean = False

        '' ******* METHOD 2 *******
        '' Replacement for the abandoned method
        '' Here we are assuming that the user is NOT running .NET in a Win98/Me machine...
        'Dim mc As New System.Management.ManagementClass("Win32_NetworkAdapterConfiguration")
        'Dim moc As System.Management.ManagementObjectCollection = mc.GetInstances
        'Dim mo As System.Management.ManagementObject
        'For Each mo In moc
        '    If mo("IPEnabled").ToString = "True" Then
        '        'Get all the MAC addresses separated by a +++
        '        If GetMACAddress = "" Then
        '            GetMACAddress = mo("MACAddress").ToString.Replace(":", "-")
        '        Else
        '            GetMACAddress = GetMACAddress & "-" & mo("MACAddress").ToString.Replace(":", "-")
        '        End If
        '        If mo("MACAddress").ToString.Replace(":", "-") <> "00-00-00-00-00-00" Then foundOne = True
        '    End If
        'Next mo

        'If foundOne Then Exit Function

        '' ******* METHOD 3 *******
        'Dim netInfo As New clsNetworkStats
        'Dim netStruct As clsNetworkStats.IFROW_HELPER
        'netStruct = netInfo.GetAdapter
        'GetMACAddress = netStruct.PhysAddr.ToString()
        'If GetMACAddress <> "00-00-00-00-00-00" Then Exit Function

        ' '' ******* METHOD 3 *******
        ' '' Here we are assuming that the user is NOT running .NET in a Win98/Me machine...
        ''Dim mc As ManagementClass
        ''Dim mo As ManagementObject
        ''mc = New ManagementClass("Win32_NetworkAdapterConfiguration")
        ''Dim moc As ManagementObjectCollection = mc.GetInstances()
        ''For Each mo In moc
        ''    If mo.Item("IPEnabled").ToString() = "True" Then
        ''        GetMACAddress = mo.Item("MacAddress").ToString().Replace(":", "-")
        ''        If GetMACAddress <> "00-00-00-00-00-00" Then Exit Function
        ''    End If
        ''Next

        '' ******* METHOD 4 *******
        'Dim objMOS As ManagementObjectSearcher
        'Dim objMOC As Management.ManagementObjectCollection
        'Dim objMO As Management.ManagementObject
        'objMOS = New ManagementObjectSearcher("Select * From Win32_NetworkAdapter")
        'objMOC = objMOS.Get
        'For Each objMO In objMOC
        '    GetMACAddress = objMO("MACAddress").ToString
        '    If GetMACAddress <> "00-00-00-00-00-00" Then Exit Function
        'Next
        'objMOS.Dispose()
        'objMOS = Nothing
        'objMO = Nothing

        '' ******* METHOD 5 *******
        'Dim nic As System.Net.NetworkInformation.NetworkInterface = Nothing
        'For Each nic In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces
        '    GetMACAddress = nic.GetPhysicalAddress().ToString
        '    If GetMACAddress <> "" And GetMACAddress <> "00-00-00-00-00-00" Then
        '        nic = Nothing
        '        Exit Function
        '    End If
        'Next

        '' ******* METHOD 6 *******
        ''another user provided the code below that seems to work well
        ''if the adapter card is "Ethernet 802.3", then the code below will work
        'Dim objset, obj As Object
        'objset = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_NetworkAdapter")
        'For Each obj In objset
        '    On Error Resume Next
        '    If Not IsDBNull(obj.MACAddress) Then
        '        If obj.AdapterType = "Ethernet 802.3" Then
        '            If InStr(obj.PNPDeviceID, "PCI\") <> 0 Then
        '                GetMACAddress = Replace_Renamed(obj.MACAddress, ":", "-")
        '                Exit Function
        '            End If
        '        End If
        '    End If
        'Next obj

        GetMACAddress = String.Empty
        Dim foundOne As Boolean = False

        ' ******* METHOD 1 *******
        ' Replacement for the abandoned method
        ' Here we are assuming that the user is NOT running .NET in a Win98/Me machine...
        Dim mc As New System.Management.ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim moc As System.Management.ManagementObjectCollection = mc.GetInstances
        Dim mo As System.Management.ManagementObject
        Dim strO As String = ""
        For Each mo In moc
            If mo("IPEnabled").ToString = "True" Then
                strO = mo("MACAddress").ToString.Replace(":", "-")
                'Get all the MAC addresses separated by a +++
                If GetMACAddress = "" Then
                    GetMACAddress = strO
                Else
                    GetMACAddress = GetMACAddress & "___" & strO
                End If
                If strO <> "00-00-00-00-00-00" And strO <> "000000000000" And strO <> "" Then foundOne = True
            End If
        Next mo
        moc.Dispose()
        moc = Nothing
        mo = Nothing

        If foundOne Then Exit Function

        ' ******* METHOD 2 *******
        Dim nic As System.Net.NetworkInformation.NetworkInterface = Nothing
        Dim strM As String = ""
        For Each nic In System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces
            strM = nic.GetPhysicalAddress().ToString
            'Get all the MAC addresses separated by a +++
            If GetMACAddress = "" Then
                GetMACAddress = strM
            Else
                If strM <> "" Then GetMACAddress = GetMACAddress & "___" & strM
            End If
            If strM <> "00-00-00-00-00-00" And strM <> "000000000000" And strM <> "" Then foundOne = True
        Next
        nic = Nothing

        If foundOne Then Exit Function

        ' ******* METHOD 3 *******
        'another user provided the code below that seems to work well
        'if the adapter card is "Ethernet 802.3", then the code below will work
        Dim objset, obj As Object
        Dim strT As String = ""
        objset = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_NetworkAdapter")
        For Each obj In objset
            On Error Resume Next
            If Not IsDBNull(obj.MACAddress) Then
                If obj.AdapterType = "Ethernet 802.3" Then
                    If InStr(obj.PNPDeviceID, "PCI\") <> 0 Then
                        strT = Replace(obj.MACAddress, ":", "-")
                        If GetMACAddress = "" Then
                            GetMACAddress = strT
                        Else
                            If strT <> "" Then GetMACAddress = GetMACAddress & "___" & strT
                        End If
                        If strT <> "00-00-00-00-00-00" And strT <> "000000000000" And strT <> "" Then foundOne = True
                    End If
                End If
            End If
        Next obj
        objset = Nothing
        obj = Nothing

GetMACAddressError:
        If GetMACAddress = "" Then
            GetMACAddress = "Not Available"
        End If

    End Function
    Public Function WirelessIsFoundAndConnected() As Boolean
        Dim adapters() As NetworkInformation.NetworkInterface = NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()
        Dim wirelessfound As Boolean = False
        For Each adapter As NetworkInformation.NetworkInterface In adapters
            If adapter.NetworkInterfaceType = NetworkInformation.NetworkInterfaceType.Wireless80211 Then
                'If adapter.Description.ToLower.Contains("wireless") Then  ' this also works
                If adapter.GetIPProperties().UnicastAddresses.Count > 0 Then
                    WirelessIsFoundAndConnected = True
                End If
                'End If

                'Debug.WriteLine("Name " & adapter.Name)
                'Debug.WriteLine("Status:" & adapter.OperationalStatus.ToString)
                'Debug.WriteLine("Speed:" & adapter.Speed.ToString())
                'Debug.WriteLine(adapter.GetIPProperties.GetIPv4Properties)
                'MessageBox.Show(adapter.GetIPProperties.GetIPv4Properties.IsDhcpEnabled.ToString)
                'If adapter.GetIPProperties.GetIPv4Properties.IsDhcpEnabled Then
                '    Debug.WriteLine("Dynamic IP")
                'Else
                '    Debug.WriteLine("Static IP")
                'End If
                'WirelessIsFoundAndConnected = True

            End If
        Next
    End Function
    '===============================================================================
    ' Name: Function GetWindowsSerial
    ' Input: None
    ' Output:
    '   String - Windows serial number
    ' Purpose: Gets the Windows Serial Number
    ' Remarks: .NET way of doing things added
    '===============================================================================
    ''' <summary>
    ''' Gets the Windows Serial Number
    ''' </summary>
    ''' <returns>Windows serial number</returns>
    ''' <remarks>.NET way of doing things added</remarks>
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
    ''' <summary>
    ''' Gets the BIOS Serial Number
    ''' </summary>
    ''' <returns>BIOS serial number</returns>
    ''' <remarks>Uses the WMI</remarks>
    Public Function GetBiosVersion() As String
        Dim BiosSet As Object
        Dim obj As Object

        GetBiosVersion = String.Empty

        On Error GoTo GetBiosVersionerror
        BiosSet = GetObject("WinMgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_BIOS")
        For Each obj In BiosSet
            GetBiosVersion = obj.Version
            GetBiosVersion = GetBiosVersion.Replace(" ", "")
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
    ''' <summary>
    ''' Gets the Motherboard Serial Number
    ''' </summary>
    ''' <returns>Motherboard serial number</returns>
    ''' <remarks>Uses the WMI</remarks>
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
                GetMotherboardSerial = GetMotherboardSerial.Trim
                If GetMotherboardSerial = "" Then
                    GetMotherboardSerial = "Not Available"
                End If
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
    ''' <summary>
    ''' Gets the IP address
    ''' </summary>
    ''' <returns>IP address</returns>
    ''' <remarks></remarks>
    Public Function GetIPaddress() As String
        On Error GoTo GetIPaddressError
        GetIPaddress = String.Empty

        If IsWebConnected() = False Then
            GetIPaddress = "-1"
            Exit Function
        End If

        ' This is the old method 
        ' It worked but not necessarily all the time
        ' There could be many IP addresses and one has to check if they are empty
        'Dim ipEntry As IPHostEntry = Dns.GetHostEntry(Environment.MachineName)
        'Dim IpAddr As IPAddress() = ipEntry.AddressList
        'GetIPaddress = IpAddr(0).ToString()

        'A hostmachine can have more than one IP assigned 
        Dim NIC_IPs() As IPAddress = Dns.GetHostEntry(Dns.GetHostName()).AddressList
        For Each IPAdr As IPAddress In NIC_IPs
            GetIPaddress = IPAdr.ToString
            If GetIPaddress <> "0.0.0.0" And GetIPaddress <> "127.0.0.1" And GetIPaddress.Contains(":") = False Then Exit Function
        Next

GetIPaddressError:
        If GetIPaddress = "" Then
            GetIPaddress = "Not Available"
        End If
    End Function
    Private Function GetHDSerialFirmwareVBNET(ByVal controller As Integer, Optional ByVal masterDrive As Boolean = True) As Object

        ' Created with the help of the following articles and clues from ALCrypto3NET.dll
        ' SOURCE 1: http://discuss.develop.com/archives/wa.exe?A2=ind0309a&L=advanced-dotnet&D=0&T=0&P=3760
        ' SOURCE 2: http://www.visual-basic.it/scarica.asp?ID=611
        ' SOURCE 3: ALCrypto3NET.dll and DISKID32 program

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
    Public Function GetHDSerialFirmwareWMI() As String
        GetHDSerialFirmwareWMI = ""

        Dim managementScope As New ManagementScope("\root\cimv2")
        managementScope.Options.Impersonation = System.Management.ImpersonationLevel.Impersonate

        Dim searcher As New ManagementObjectSearcher(managementScope, New ObjectQuery("SELECT * FROM Win32_DiskDrive WHERE InterfaceType=""IDE"" or InterfaceType=""SCSI"""))
        For Each disk As ManagementObject In searcher.[Get]()
            If disk("PNPDeviceID") IsNot Nothing Then
                Dim pnpDeviceID As String = disk("PNPDeviceID").ToString()

                Dim split As String() = pnpDeviceID.Split(New String() {"\"}, StringSplitOptions.None)
                If split.Length = 3 Then
                    If Not split(2).Contains("&") Then
                        If split(2).Contains("_") Then split(2) = split(2).Substring(0, split(2).IndexOf("_"))
                        Dim bytes As Byte() = GetHexStringBytes(split(2))
                        If bytes.Length > 0 Then
                            GetHDSerialFirmwareWMI = ReverseSerialNumber(System.Text.Encoding.UTF8.GetString(bytes)).Trim()
                        End If
                    Else
                        ' Custom checks go into here
                        Dim parts() As String
                        parts = pnpDeviceID.Split("\".ToCharArray)
                        ' The serial number should be the next to the last element
                        GetHDSerialFirmwareWMI = parts(parts.Length - 1)
                        GetHDSerialFirmwareWMI = Replace(GetHDSerialFirmwareWMI, "&", "")
                    End If
                End If
            End If
        Next
    End Function

    Private Function ReverseSerialNumber(ByVal serialNumber As String) As String
        serialNumber = serialNumber.Trim()
        Dim sb As New StringBuilder()
        For i As Integer = 0 To serialNumber.Length - 1 Step 2
            sb.Append(serialNumber(i + 1).ToString() + serialNumber(i).ToString())
        Next
        serialNumber = sb.ToString()
        sb = Nothing
        Return serialNumber
    End Function

    Private Function GetHexStringBytes(ByVal hex As String) As Byte()
        Try
            If hex.Contains([String].Empty) Then
                hex = hex.Replace(" ", [String].Empty)
            End If
            If hex.Length Mod 2 = 1 Then
                hex = "0" & hex
            End If
            Dim size As Integer = CInt(CDbl(hex.Length) / CDbl(2))
            Dim bytes As Byte() = New Byte(size - 1) {}
            For i As Integer = 0 To size - 1
                bytes(i) = Convert.ToByte(hex.Substring(i * 2, 2), 16)
            Next
            Return bytes
        Catch
            Return New Byte() {}
        End Try
    End Function

    Public Function GetCPUID() As String
        GetCPUID = String.Empty

        Dim managementScope As New ManagementScope("\root\cimv2")
        managementScope.Options.Impersonation = System.Management.ImpersonationLevel.Impersonate

        Dim searcher As New ManagementObjectSearcher(managementScope, New ObjectQuery("SELECT * FROM Win32_Processor"))
        For Each disk As ManagementObject In searcher.[Get]()
            If disk("ProcessorID") IsNot Nothing Then
                GetCPUID = disk("ProcessorID").ToString()
            Else
                GetCPUID = "Not Available"
            End If
        Next
    End Function
    Public Function GetBaseBoardID() As String
        GetBaseBoardID = String.Empty

        Dim managementScope As New ManagementScope("\root\cimv2")
        managementScope.Options.Impersonation = System.Management.ImpersonationLevel.Impersonate

        Dim searcher As New ManagementObjectSearcher(managementScope, New ObjectQuery("SELECT * FROM Win32_Baseboard"))
        For Each disk As ManagementObject In searcher.[Get]()
            If disk("SerialNumber") IsNot Nothing Then
                GetBaseBoardID = disk("SerialNumber").ToString()
                GetBaseBoardID = GetBaseBoardID.Replace(".", "")
            Else
                GetBaseBoardID = "Not Available"
            End If
        Next
    End Function

    Public Function GetVideoID() As String
        GetVideoID = String.Empty

        Dim managementScope As New ManagementScope("\root\cimv2")
        managementScope.Options.Impersonation = System.Management.ImpersonationLevel.Impersonate

        Dim searcher As New ManagementObjectSearcher(managementScope, New ObjectQuery("SELECT * FROM Win32_VideoController"))
        For Each disk As ManagementObject In searcher.[Get]()
            If disk("DriverVersion") IsNot Nothing Then
                GetVideoID = disk("DriverVersion").ToString()
            Else
                GetVideoID = "Not Available"
            End If
        Next
    End Function

    Public Function GetMemoryID() As String
        GetMemoryID = String.Empty

        Dim managementScope As New ManagementScope("\root\cimv2")
        managementScope.Options.Impersonation = System.Management.ImpersonationLevel.Impersonate

        Dim searcher As New ManagementObjectSearcher(managementScope, New ObjectQuery("SELECT * FROM Win32_MemoryDevice"))
        For Each disk As ManagementObject In searcher.[Get]()
            If disk("SystemName") IsNot Nothing Then
                Dim MemoryID As String = disk("EndingAddress").ToString() & "-" & My.Computer.Info.TotalPhysicalMemory.ToString
                Return MemoryID
            Else
                GetMemoryID = "Not Available"
            End If
        Next
    End Function

    'Private deDomainRoot As DirectoryEntry ' This isn't VS2005 compatible - commented out for now
    Private strDomainPath As String
    Private Declare Auto Function ConvertSidToStringSid Lib "advapi32.dll" (ByVal bSID As IntPtr, <System.Runtime.InteropServices.In(), System.Runtime.InteropServices.Out(), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPTStr)> ByRef SIDString As String) As Integer

    ' The following function was added here to
    ' get the computer SID but turned out to be problematic
    ' This isn't VS2005 compatible and the function requires admin rights which is no good
    'Private Function ConnectToAD() As Boolean
    '    Try
    '        deDomainRoot = New DirectoryEntry("LDAP://rootDSE")
    '        strDomainPath = "LDAP://" + deDomainRoot.Properties("DefaultNamingContext")(0).ToString()
    '        deDomainRoot = New DirectoryEntry(strDomainPath)
    '        Return True
    '    Catch ex As Exception
    '        Return False
    '    End Try
    'End Function
    Private query As ManagementObjectSearcher
    Private queryCollection As ManagementObjectCollection

    Public Function GetSID() As String
        GetSID = String.Empty

        Dim co As New ConnectionOptions()
        co.Username = System.Environment.UserDomainName
        Dim msc As New ManagementScope("\root\cimv2", co)
        Dim queryString As String = "SELECT * FROM Win32_UserAccount where name='" & co.Username & "'"
        Dim q As New SelectQuery(queryString)
        query = New ManagementObjectSearcher(msc, q)
        queryCollection = query.[Get]()
        Dim res As String = [String].Empty
        For Each mo As ManagementObject In queryCollection
            ' there should be only one here! 
            res += mo("SID").ToString()
        Next
        GetSID = res

        'If (ConnectToAD()) Then
        '    Dim dirSearcher As New DirectorySearcher
        '    Dim singleQueryResult As SearchResult
        '    Dim strSID As String = ""
        '    Dim intSuccess As Integer
        '    Dim userName As String = "administrator"
        '    Try
        '        dirSearcher.SearchScope = SearchScope.Subtree
        '        dirSearcher.SearchRoot = deDomainRoot
        '        dirSearcher.Filter = "(&(sAMAccountName=" & userName & "))"
        '        singleQueryResult = dirSearcher.FindOne()
        '        Dim sidBytes As Byte() = CType(singleQueryResult.Properties("objectSid")(0), Byte())
        '        Dim sidPtr As IntPtr = System.Runtime.InteropServices.Marshal.AllocHGlobal(sidBytes.Length)
        '        System.Runtime.InteropServices.Marshal.Copy(sidBytes, 0, sidPtr, sidBytes.Length)
        '        intSuccess = ConvertSidToStringSid(sidPtr, strSID)
        '        GetSID = strSID.Trim()
        '    Catch ex As Exception
        '        Return "Not Available"
        '    End Try
        'End If
        If GetSID() = "" Then
            GetSID = "Not Available"
        End If

    End Function
    Public Function GetExternalIP() As String
        Dim IP_URL As String = "http://checkip.dyndns.org"
        Dim strHTML, strIP As String
        Try
            Dim objWebReq As System.Net.WebRequest = System.Net.WebRequest.Create(IP_URL)
            Dim objWebResp As System.Net.WebResponse = objWebReq.GetResponse()
            Dim strmResp As System.IO.Stream = objWebResp.GetResponseStream()
            Dim srResp As System.IO.StreamReader = New System.IO.StreamReader(strmResp, System.Text.Encoding.UTF8)
            strHTML = srResp.ReadToEnd()
            Dim regexIP As System.Text.RegularExpressions.Regex
            regexIP = New System.Text.RegularExpressions.Regex("\b\d{1,3}.\d{1,3}.\d{1,3}.\d{1,3}\b")
            strIP = regexIP.Match(strHTML).Value
            If strIP.Contains("HTML") Or strIP.Contains("DOC") Or strIP.Contains("Transitional") Then
                strIP = "Not Available"
            End If
            Return strIP
        Catch ex As Exception
            GetExternalIP = "Not Available"
        End Try
    End Function
    Public Function GetFingerprint() As String
        fp.UseCpuID = True
        fp.UseBiosID = True
        fp.UseBaseID = True
        fp.UseDiskID = True
        ' Do not use the Video ID and MAC Address in Fingerprint
        ' since these are not very reliable and might change
        ' from admin to limited user accounts
        fp.UseVideoID = False
        fp.UseMacID = False
        fp.ReturnLength = 8
        GetFingerprint = fp.Value
    End Function
    Public Function CheckMACaddress(ByVal usedMACaddress As String) As Boolean
        CheckMACaddress = False
        Dim mc As ManagementClass
        Dim mo As ManagementObject
        Dim nicMACaddress As String
        mc = New ManagementClass("Win32_NetworkAdapterConfiguration")
        Dim moc As ManagementObjectCollection = mc.GetInstances()
        For Each mo In moc
            'If mo.Item("IPEnabled").ToString() = "True" Then
            nicMACaddress = mo.Item("MacAddress").ToString().Replace(":", "-")
            If nicMACaddress = usedMACaddress Then
                Return True
            End If
            'End If
        Next

    End Function
End Module