Attribute VB_Name = "modHardware"
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
' Modified: 08.15.2005
'===============================================================================

Public Declare Function ComputerName Lib "kernel32" Alias _
        "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Declare Function apiGetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
    (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

'HDD firmware serial number
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2
Private Const OPEN_EXISTING = 3
Private Const CREATE_NEW = 1
Private Const INVALID_HANDLE_VALUE = -1
Private Const VER_PLATFORM_WIN32_NT = 2
Private Const IDENTIFY_BUFFER_SIZE = 512
Public Const READ_THRESHOLD_BUFFER_SIZE = 512
Private Const OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16

Dim colAttrNames As Collection

'---------------------------------------------------------------------
' The following structure defines the structure of a Drive Attribute
'---------------------------------------------------------------------
Public Const NUM_ATTRIBUTE_STRUCTS = 30

Public Type DRIVEATTRIBUTE
       bAttrID As Byte         ' Identifies which attribute
       wStatusFlags As Integer 'Integer ' see bit definitions below
       bAttrValue As Byte      ' Current normalized value
       bWorstValue As Byte     ' How bad has it ever been?
       bRawValue(5) As Byte    ' Un-normalized value
       bReserved As Byte       ' ...
End Type

'---------------------------------------------------------------------
' The following structure defines the structure of a Warranty Threshold
' Obsoleted in ATA4!
'---------------------------------------------------------------------
Public Type ATTRTHRESHOLD
       bAttrID As Byte            ' Identifies which attribute
       bWarrantyThreshold As Byte ' Triggering value
       bReserved(9) As Byte       ' ...
End Type

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

'GETVERSIONOUTPARAMS contains the data returned
'from the Get Driver Version function
Private Type GETVERSIONOUTPARAMS
   bVersion       As Byte 'Binary driver version.
   bRevision      As Byte 'Binary driver revision
   bReserved      As Byte 'Not used
   bIDEDeviceMap  As Byte 'Bit map of IDE devices
   fCapabilities  As Long 'Bit mask of driver capabilities
   dwReserved(3)  As Long 'For future use
End Type

'IDE registers
Private Type IDEREGS
   bFeaturesReg     As Byte 'Used for specifying SMART "commands"
   bSectorCountReg  As Byte 'IDE sector count register
   bSectorNumberReg As Byte 'IDE sector number register
   bCylLowReg       As Byte 'IDE low order cylinder value
   bCylHighReg      As Byte 'IDE high order cylinder value
   bDriveHeadReg    As Byte 'IDE drive/head register
   bCommandReg      As Byte 'Actual IDE command
   bReserved        As Byte 'reserved for future use - must be zero
End Type

'SENDCMDINPARAMS contains the input parameters for the
'Send Command to Drive function
Private Type SENDCMDINPARAMS
   cBufferSize     As Long     'Buffer size in bytes
   irDriveRegs     As IDEREGS  'Structure with drive register values.
   bDriveNumber    As Byte     'Physical drive number to send command to (0,1,2,3).
   bReserved(2)    As Byte     'Bytes reserved
   dwReserved(3)   As Long     'DWORDS reserved
   bBuffer()      As Byte      'Input buffer.
End Type

'Valid values for the bCommandReg member of IDEREGS.
Private Const IDE_ATAPI_ID = &HA1               ' Returns ID sector for ATAPI.
Private Const IDE_ID_FUNCTION = &HEC            'Returns ID sector for ATA.
Private Const IDE_EXECUTE_SMART_FUNCTION = &HB0 'Performs SMART cmd.
                                                'Requires valid bFeaturesReg,
                                                'bCylLowReg, and bCylHighReg

'Cylinder register values required when issuing SMART command
Private Const SMART_CYL_LOW = &H4F
Private Const SMART_CYL_HI = &HC2

'Status returned from driver
Private Type DRIVERSTATUS
   bDriverError  As Byte          'Error code from driver, or 0 if no error
   bIDEStatus    As Byte          'Contents of IDE Error register
                                  'Only valid when bDriverError is SMART_IDE_ERROR
   bReserved(1)  As Byte
   dwReserved(1) As Long
 End Type

Private Type IDSECTOR
   wGenConfig                 As Integer
   wNumCyls                   As Integer
   wReserved                  As Integer
   wNumHeads                  As Integer
   wBytesPerTrack             As Integer
   wBytesPerSector            As Integer
   wSectorsPerTrack           As Integer
   wVendorUnique(2)           As Integer
   sSerialNumber(19)          As Byte
   wBufferType                As Integer
   wBufferSize                As Integer
   wECCSize                   As Integer
   sFirmwareRev(7)            As Byte
   sModelNumber(39)           As Byte
   wMoreVendorUnique          As Integer
   wDoubleWordIO              As Integer
   wCapabilities              As Integer
   wReserved1                 As Integer
   wPIOTiming                 As Integer
   wDMATiming                 As Integer
   wBS                        As Integer
   wNumCurrentCyls            As Integer
   wNumCurrentHeads           As Integer
   wNumCurrentSectorsPerTrack As Integer
   ulCurrentSectorCapacity    As Long
   wMultSectorStuff           As Integer
   ulTotalAddressableSectors  As Long
   wSingleWordDMA             As Integer
   wMultiWordDMA              As Integer
   bReserved(127)             As Byte
End Type

'Structure returned by SMART IOCTL commands
Private Type SENDCMDOUTPARAMS
  cBufferSize   As Long         'Size of Buffer in bytes
  DRIVERSTATUS  As DRIVERSTATUS 'Driver status structure
  bBuffer()    As Byte          'Buffer of arbitrary length for data read from drive
End Type

'Vendor specific feature register defines
'for SMART "sub commands"
Private Const SMART_READ_ATTRIBUTE_VALUES = &HD0
Private Const SMART_READ_ATTRIBUTE_THRESHOLDS = &HD1
Private Const SMART_ENABLE_DISABLE_ATTRIBUTE_AUTOSAVE = &HD2
Private Const SMART_SAVE_ATTRIBUTE_VALUES = &HD3
Private Const SMART_EXECUTE_OFFLINE_IMMEDIATE = &HD4
' Vendor specific commands:
Private Const SMART_ENABLE_SMART_OPERATIONS = &HD8
Private Const SMART_DISABLE_SMART_OPERATIONS = &HD9
Private Const SMART_RETURN_SMART_STATUS = &HDA

'Status Flags Values
Public Enum STATUS_FLAGS
   PRE_FAILURE_WARRANTY = &H1
   ON_LINE_COLLECTION = &H2
   PERFORMANCE_ATTRIBUTE = &H4
   ERROR_RATE_ATTRIBUTE = &H8
   EVENT_COUNT_ATTRIBUTE = &H10
   SELF_PRESERVING_ATTRIBUTE = &H20
End Enum

'IOCTL commands
Private Const DFP_GET_VERSION = &H74080
Private Const DFP_SEND_DRIVE_COMMAND = &H7C084
Private Const DFP_RECEIVE_DRIVE_DATA = &H7C088

Private Type ATTR_DATA
   AttrID As Byte
   AttrName As String
   AttrValue As Byte
   ThresholdValue As Byte
   WorstValue As Byte
   StatusFlags As STATUS_FLAGS
End Type

Public Type DRIVE_INFO
   bDriveType As Byte
   SerialNumber As String
   Model As String
   FirmWare As String
   Cilinders As Long
   Heads As Long
   SecPerTrack As Long
   BytesPerSector As Long
   BytesperTrack As Long
   NumAttributes As Byte
   Attributes() As ATTR_DATA
End Type
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

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function DeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, lpInBuffer As Any, ByVal nInBufferSize As Long, lpOutBuffer As Any, ByVal nOutBufferSize As Long, lpBytesReturned As Long, ByVal lpOverlapped As Long) As Long

Private Type BufferType
     myBuffer(559) As Byte
End Type

' The following UDT and the DLL function is for getting
' the serial number from a C++ DLL in case the VB6 APIs fail
' Currently, VB code cannot handle the serial numbers
' coming from computers with non-admin rights; in those
' cases the C++ DLL function "getHardDriveFirmware" should
' work properly.
' Neither of the two methods work for the SATA and SCSI drives
' ialkan - 8312005
Private Type MyUDT2
    myStr As String * 30
    mL As Long
End Type
Private Declare Function getHardDriveFirmware Lib "ALCrypto3.dll" (myU As MyUDT2) As Long

'MAC Address
Public Const NCBASTAT As Long = &H33
Public Const NCBNAMSZ As Long = 16
Public Const HEAP_ZERO_MEMORY As Long = &H8
Public Const HEAP_GENERATE_EXCEPTIONS As Long = &H4
Public Const NCBRESET As Long = &H32

Public Type NET_CONTROL_BLOCK  'NCB
   ncb_command    As Byte
   ncb_retcode    As Byte
   ncb_lsn        As Byte
   ncb_num        As Byte
   ncb_buffer     As Long
   ncb_length     As Integer
   ncb_callname   As String * NCBNAMSZ
   ncb_name       As String * NCBNAMSZ
   ncb_rto        As Byte
   ncb_sto        As Byte
   ncb_post       As Long
   ncb_lana_num   As Byte
   ncb_cmd_cplt   As Byte
   ncb_reserve(9) As Byte ' Reserved, must be 0
   ncb_event      As Long
End Type

Public Type ADAPTER_STATUS
   adapter_address(5) As Byte
   rev_major         As Byte
   reserved0         As Byte
   adapter_type      As Byte
   rev_minor         As Byte
   duration          As Integer
   frmr_recv         As Integer
   frmr_xmit         As Integer
   iframe_recv_err   As Integer
   xmit_aborts       As Integer
   xmit_success      As Long
   recv_success      As Long
   iframe_xmit_err   As Integer
   recv_buff_unavail As Integer
   t1_timeouts       As Integer
   ti_timeouts       As Integer
   Reserved1         As Long
   free_ncbs         As Integer
   max_cfg_ncbs      As Integer
   max_ncbs          As Integer
   xmit_buf_unavail  As Integer
   max_dgram_size    As Integer
   pending_sess      As Integer
   max_cfg_sess      As Integer
   max_sess          As Integer
   max_sess_pkt_size As Integer
   name_count        As Integer
End Type
   
Public Type NAME_BUFFER
   name        As String * NCBNAMSZ
   name_num    As Integer
   name_flags  As Integer
End Type

Public Type ASTAT
   adapt          As ADAPTER_STATUS
   NameBuff(30)   As NAME_BUFFER
End Type

Public Declare Function Netbios Lib "netapi32.dll" (pncb As NET_CONTROL_BLOCK) As Byte
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function GetProcessHeap Lib "kernel32" () As Long
Public Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal lpMem As Long) As Long
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)

' Following used to get the IP address
Public Const MIN_SOCKETS_REQD As Long = 1
Public Const WS_VERSION_REQD As Long = &H101
Public Const WS_VERSION_MAJOR As Long = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR As Long = WS_VERSION_REQD And &HFF&
Public Const SOCKET_ERROR As Long = -1
Public Const WSADESCRIPTION_LEN = 257
Public Const WSASYS_STATUS_LEN = 129
Public Const MAX_WSADescription = 256
Public Const MAX_WSASYSStatus = 128
Public Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Integer
    wMaxUDPDG As Integer
    dwVendorInfo As Long
End Type
Type WSADataInfo
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * WSADESCRIPTION_LEN
    szSystemStatus As String * WSASYS_STATUS_LEN
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As String
End Type
Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLen As Integer
    hAddrList As Long
End Type
Declare Function WSAStartupInfo Lib "WSOCK32" Alias "WSAStartup" (ByVal wVersionRequested As Integer, lpWSADATA As WSADataInfo) As Long
Declare Function WSACleanup Lib "WSOCK32" () As Long
Declare Function WSAGetLastError Lib "WSOCK32" () As Long
Declare Function WSAStartup Lib "WSOCK32" (ByVal wVersionRequired As Long, lpWSADATA As WSAData) As Long
Declare Function gethostname Lib "WSOCK32" (ByVal szHost As String, ByVal dwHostLen As Long) As Long
Declare Function gethostbyname Lib "WSOCK32" (ByVal szHost As String) As Long
Declare Sub CopyMemoryIP Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, ByVal hpvSource As Long, ByVal cbCopy As Long)

' Declarations needed for GetAdaptersInfo & GetIfTable to get MAC address
Private Const MIB_IF_TYPE_OTHER                   As Long = 1
Private Const MIB_IF_TYPE_ETHERNET                As Long = 6
Private Const MIB_IF_TYPE_TOKENRING               As Long = 9
Private Const MIB_IF_TYPE_FDDI                    As Long = 15
Private Const MIB_IF_TYPE_PPP                     As Long = 23
Private Const MIB_IF_TYPE_LOOPBACK                As Long = 24
Private Const MIB_IF_TYPE_SLIP                    As Long = 28

Private Const MIB_IF_ADMIN_STATUS_UP              As Long = 1
Private Const MIB_IF_ADMIN_STATUS_DOWN            As Long = 2
Private Const MIB_IF_ADMIN_STATUS_TESTING         As Long = 3

Private Const MIB_IF_OPER_STATUS_NON_OPERATIONAL  As Long = 0
Private Const MIB_IF_OPER_STATUS_UNREACHABLE      As Long = 1
Private Const MIB_IF_OPER_STATUS_DISCONNECTED     As Long = 2
Private Const MIB_IF_OPER_STATUS_CONNECTING       As Long = 3
Private Const MIB_IF_OPER_STATUS_CONNECTED        As Long = 4
Private Const MIB_IF_OPER_STATUS_OPERATIONAL      As Long = 5

Private Const MAX_ADAPTER_DESCRIPTION_LENGTH      As Long = 128
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH_p    As Long = MAX_ADAPTER_DESCRIPTION_LENGTH + 4
Private Const MAX_ADAPTER_NAME_LENGTH             As Long = 256
Private Const MAX_ADAPTER_NAME_LENGTH_p           As Long = MAX_ADAPTER_NAME_LENGTH + 4
Private Const MAX_ADAPTER_ADDRESS_LENGTH          As Long = 8
Private Const DEFAULT_MINIMUM_ENTITIES            As Long = 32
Private Const MAX_HOSTNAME_LEN                    As Long = 128
Private Const MAX_DOMAIN_NAME_LEN                 As Long = 128
Private Const MAX_SCOPE_ID_LEN                    As Long = 256

Private Const MAXLEN_IFDESCR                      As Long = 256
Private Const MAX_INTERFACE_NAME_LEN              As Long = MAXLEN_IFDESCR * 2
Private Const MAXLEN_PHYSADDR                     As Long = 8

' Information structure returned by GetIfEntry/GetIfTable
Private Type MIB_IFROW
    wszName(0 To MAX_INTERFACE_NAME_LEN - 1) As Byte    ' MSDN Docs say pointer, but it is WCHAR array
    dwIndex             As Long
    dwType              As Long
    dwMtu               As Long
    dwSpeed             As Long
    dwPhysAddrLen       As Long
    bPhysAddr(MAXLEN_PHYSADDR - 1) As Byte
    dwAdminStatus       As Long
    dwOperStatus        As Long
    dwLastChange        As Long
    dwInOctets          As Long
    dwInUcastPkts       As Long
    dwInNUcastPkts      As Long
    dwInDiscards        As Long
    dwInErrors          As Long
    dwInUnknownProtos   As Long
    dwOutOctets         As Long
    dwOutUcastPkts      As Long
    dwOutNUcastPkts     As Long
    dwOutDiscards       As Long
    dwOutErrors         As Long
    dwOutQLen           As Long
    dwDescrLen          As Long
    bDescr As String * MAXLEN_IFDESCR
End Type

Public Declare Function GetAdaptersInfo Lib "iphlpapi.dll" (ByRef pAdapterInfo As Any, ByRef pOutBufLen As Long) As Long
Public Declare Function GetNumberOfInterfaces Lib "iphlpapi.dll" (ByRef pdwNumIf As Long) As Long
Public Declare Function GetIfEntry Lib "iphlpapi.dll" (ByRef pIfRow As Any) As Long
Private Declare Function GetIfTable Lib "iphlpapi.dll" _
        (ByRef pIfTable As Any, ByRef pdwSize As Long, ByVal bOrder As Long) As Long

' Getting External IP Address
Dim VbString, IP As String
Dim StrEnd, a As Integer
  
Public Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Public Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hOpen As Long, ByVal sUrl As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Public Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Public Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
  
Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3
  
Public Const scUserAgent = "VB OpenUrl"
Public Const INTERNET_FLAG_RELOAD = &H80000000

Function GetExternalIP(URL As String) As String
  
Dim hOpen As Long
Dim hOpenUrl As Long
Dim bDoLoop As Boolean
Dim bRet As Boolean
Dim sReadBuffer As String * 2048
Dim lNumberOfBytesRead As Long
Dim sBuffer As String

hOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
hOpenUrl = InternetOpenUrl(hOpen, URL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)

bDoLoop = True
While bDoLoop
    sReadBuffer = vbNullString
    bRet = InternetReadFile(hOpenUrl, sReadBuffer, Len(sReadBuffer), lNumberOfBytesRead)
    sBuffer = sBuffer & Left$(sReadBuffer, lNumberOfBytesRead)
    If Not CBool(lNumberOfBytesRead) Then bDoLoop = False
Wend
 
VbString = sBuffer
 
VbString = Mid(VbString, InStr(VbString, "IP Address:") + 12, 20)
StrEnd = InStr(VbString, "<br>") - 2

For a = 1 To StrEnd
    IP = IP + Mid(VbString, a, 1)
Next
 
GetExternalIP = IP
 
If hOpenUrl <> 0 Then InternetCloseHandle (hOpenUrl)
If hOpen <> 0 Then InternetCloseHandle (hOpen)
  
End Function

'===============================================================================
' Name: Function GetComputerName
' Input: None
' Output:
'   String - Computer name
' Purpose: Gets the computer name on the network
' Remarks: Stolen from ActiveLock 1.89. Author Unknown
'===============================================================================
Public Function GetComputerName() As String
Dim strString As String
'Create a buffer
strString = String(255, Chr$(0))
'Get the computer name
ComputerName strString, 255
'Remove the unnecessary Chr$(0)
strString = Left$(strString, InStr(1, strString, Chr$(0)) - 1)
GetComputerName = strString

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
Public Function HDSerial(path As String) As String
    
    Dim lngReturn As Long, lngDummy1 As Long, lngDummy2 As Long, lngSerial As Long
    Dim strDummy1 As String, strDummy2 As String, strSerial As String
    
    Dim strDriveLetter As String, lngFirstSlash As Long
    
    strDriveLetter = path
    
    '(Just in case... It's better to be safe than sorry.)
    strDriveLetter = Replace(strDriveLetter, "/", "\")
    
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
    strSerial = String(8 - Len(strSerial), "0") & strSerial
    strSerial = Left(strSerial, 4) & "-" & Right(strSerial, 4)
    HDSerial = strSerial
    
    ' Alternative Code - Short Version
'    Dim volbuf$, sysname$, sysflags&, componentlength&
'    Dim serialnum As Long
'    GetVolumeInformation "C:\", volbuf$, 255, serialnum, componentlength, sysflags, sysname$, 255
'    HDSerial = CStr(serialnum)

End Function

'===============================================================================
' Name: Function GetHDSerialFirmware
' Input: None
' Output:
'   String - HDD Firmware Serial Number
' Purpose: Function to return the HDD Firmware Serial Number (Actual Physical Serial Number)
' Remarks: None
'===============================================================================
Function GetHDSerialFirmware(ii As Integer) As String
Dim jj As Integer
On Error GoTo GetHDSerialFirmwareError

' ******* METHOD 1 - SMART *******
If ii = 0 Then
    Dim di As DRIVE_INFO
    Dim drvNumber As Long
    di = GetDriveInfo(PRIMARY_MASTER)
    GetHDSerialFirmware = Trim(di.SerialNumber)
    Exit Function
End If

' ******* METHOD 2 - SCSI - PURE VB6 *******
' ialkan 2-12-06
' Pure VB6 version of the code found in several online resources
' described in GetHDSerialFirmwareVB6 function
' This eliminates the dependency of the HDD firmware serial number
' function from ALCrypto3.dll
If ii = 2 Then
    For jj = 0 To 15 ' Controller index
        GetHDSerialFirmware = Trim(GetHDSerialFirmwareVB6(jj, True))    ' Check the Master drive
        If GetHDSerialFirmware <> "" Then Exit Function
        GetHDSerialFirmware = Trim(GetHDSerialFirmwareVB6(jj, False))   ' Now check the Slave Drive
        If GetHDSerialFirmware <> "" Then Exit Function
    Next
    Exit Function
End If

' ******* METHOD 3 - ALCRYPTO *******
' Use ALCrypto DLL
If ii = 1 Then
    Dim mU As MyUDT2
    Call getHardDriveFirmware(mU)
    GetHDSerialFirmware = Trim(StripControlChars(mU.myStr, False))
    Exit Function
End If

' ******* METHOD 4 - WMI *******
' Works under Vista and with UAC
' This code was supplied by Daniel Gochin because the Vista code works without admin rights and under UAC !!!
' Modified a bit by Ismail Alkan - Nov'2008
'If ii = 3 Then
'    GetHDSerialFirmware = GetSerialNumberFromWMI("Win32_PhysicalMedia")
'    Exit Function
'End If
If ii = 4 Then
    GetHDSerialFirmware = GetSerialNumberFromWMI("Win32_DiskDrive")
    Exit Function
End If

' Well, this is not so good, because we still don't have
' a serial number in our hands...
' Cannot return an empty string...
GetHDSerialFirmwareError:
If GetHDSerialFirmware = "" Then
    GetHDSerialFirmware = "Not Available"
End If

End Function
Private Function IsNull2(vValue As Variant, vReturnValue As Variant) As Variant
    If IsNull(vValue) = True Then
        IsNull2 = vReturnValue
    Else
        IsNull2 = vValue
    End If
End Function
Private Function GetSerialNumberFromWMI(wmi_selection As String) As String
    Dim o As Integer
    Dim sHDNoHex, reversedStr As String
    Dim sHDNoHexToChar As String
    Dim str1, str2 As String
    Dim jj As Integer
    Dim svc As Object
    Dim objEnum As WbemScripting.SWbemObjectSet
    Dim obj As WbemScripting.SWbemObject
    Set svc = GetObject("winmgmts:root\cimv2")
    Set objEnum = svc.ExecQuery("select * from " & wmi_selection)
    
    ' Check SerialNumber property
'    For Each obj In objEnum
'        Dim i As WbemScripting.SWbemProperty
'
'        For Each i In obj.Properties_
'            If IsNull2(i.name, "") = "SerialNumber" Then
'                'Debug.Print i.Value
'                sHDNoHex = IsNull2(i.Value, "")
'            End If
'        Next i
'    Next obj
'    If Len(sHDNoHex) > 0 Then
'        sHDNoHexToChar = ""
'        For o = 1 To Len(sHDNoHex) Step 2
'            sHDNoHexToChar = sHDNoHexToChar & Chr(CDec(("&H" & Trim(Mid(sHDNoHex, o, 2)))))
'        Next
'        reversedStr = ""
'        For jj = 0 To Len(sHDNoHexToChar) / 2
'            str1 = Mid(sHDNoHexToChar, jj * 2 + 1, 1)
'            str2 = Mid(sHDNoHexToChar, jj * 2 + 2, 1)
'            reversedStr = reversedStr & str2 & str1
'        Next
'        GetSerialNumberFromWMI = StripControlChars(Trim(reversedStr), False)
'    Else
'        GetSerialNumberFromWMI = sHDNoHex
'    End If

    Dim mPos  As Integer
    Dim k As Integer
    Dim mChar As String
    Dim mChars As String
    Dim mSerial As String
    ' Check PNPDevideID property
    If GetSerialNumberFromWMI = "" Then
        For Each obj In objEnum
            Dim ii As WbemScripting.SWbemProperty
    
            For Each ii In obj.Properties_
                If IsNull2(ii.name, "") = "PNPDeviceID" Then
                    'Debug.Print ii.Value
                    If Left(ii.Value, 3) = "IDE" Or Left(ii.Value, 4) = "SCSI" Then
                        sHDNoHex = IsNull2(ii.Value, "")
                        Exit For
                    End If
                End If
            Next ii
        Next obj
        
        Dim myString() As String
        myString = Split(sHDNoHex, "\", , vbTextCompare)
        
        If InStr(1, myString(2), "&") = 0 Then
            If InStr(1, myString(2), "_") Then myString(2) = Mid(myString(2), InStr(1, myString(2), "_"))
            sHDNoHexToChar = ""
            sHDNoHex = myString(2)
            For o = 1 To Len(sHDNoHex) Step 2
                sHDNoHexToChar = sHDNoHexToChar & Chr(CDec(("&H" & Trim(Mid(sHDNoHex, o, 2)))))
            Next
            reversedStr = ""
            For jj = 0 To Len(sHDNoHexToChar) / 2
                str1 = Mid(sHDNoHexToChar, jj * 2 + 1, 1)
                str2 = Mid(sHDNoHexToChar, jj * 2 + 2, 1)
                reversedStr = reversedStr & str2 & str1
            Next
            GetSerialNumberFromWMI = StripControlChars(Trim(reversedStr), False)
        Else
        
            mPos = InStrRev(sHDNoHex, "\")
            If mPos > 0 Then
                sHDNoHex = Mid(sHDNoHex, mPos + 1)
    
                ' Strip & characters
                Dim Index As Long
                Dim Bytes() As Byte
                
                Bytes() = sHDNoHex
                For Index = 0 To UBound(Bytes) Step 2
                    If Bytes(Index) = 38 Then
                        Bytes(Index) = 0
                    End If
                Next
                sHDNoHex = VBA.Replace(Bytes(), vbNullChar, "")
                mSerial = sHDNoHex
            
            End If
        End If
        
        GetSerialNumberFromWMI = StripControlChars(Trim(mSerial), False)
    End If
End Function

'****************************************************************************
' CheckSMARTEnable - Check if SMART enable
' FUNCTION: Send a SMART_ENABLE_SMART_OPERATIONS command to the drive
' bDriveNum = 0-3
'***************************************************************************}
Private Function CheckSMARTEnable(ByVal hDrive As Long, DriveNum As IDE_DRIVE_NUMBER) As Boolean
   'Set up data structures for Enable SMART Command.
   Dim SCIP As SENDCMDINPARAMS
   Dim SCOP As SENDCMDOUTPARAMS
   Dim lpcbBytesReturned As Long
   With SCIP
       .cBufferSize = 0
       With .irDriveRegs
            .bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI
   'Compute the drive number.
            .bDriveHeadReg = &HA0 ' Or (DriveNum And 1) * 16
            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
        End With
        .bDriveNumber = DriveNum
   End With
   CheckSMARTEnable = DeviceIoControl(hDrive, DFP_SEND_DRIVE_COMMAND, SCIP, Len(SCIP) - 4, SCOP, Len(SCOP) - 4, lpcbBytesReturned, ByVal 0&)
End Function

Public Sub FillAttrNameCollection()
   Set colAttrNames = New Collection
   With colAttrNames
       .Add "ATTR_INVALID", "0"
       .Add "READ_ERROR_RATE", "1"
       .Add "THROUGHPUT_PERF", "2"
       .Add "SPIN_UP_TIME", "3"
       .Add "START_STOP_COUNT", "4"
       .Add "REALLOC_SECTOR_COUNT", "5"
       .Add "READ_CHANNEL_MARGIN", "6"
       .Add "SEEK_ERROR_RATE", "7"
       .Add "SEEK_TIME_PERF", "8"
       .Add "POWER_ON_HRS_COUNT", "9"
       .Add "SPIN_RETRY_COUNT", "10"
       .Add "CALIBRATION_RETRY_COUNT", "11"
       .Add "POWER_CYCLE_COUNT", "12"
       .Add "SOFT_READ_ERROR_RATE", "13"
       .Add "G_SENSE_ERROR_RATE", "191"
       .Add "POWER_OFF_RETRACT_CYCLE", "192"
       .Add "LOAD_UNLOAD_CYCLE_COUNT", "193"
       .Add "TEMPERATURE", "194"
       .Add "REALLOCATION_EVENTS_COUNT", "196"
       .Add "CURRENT_PENDING_SECTOR_COUNT", "197"
       .Add "UNCORRECTABLE_SECTOR_COUNT", "198"
       .Add "ULTRADMA_CRC_ERROR_RATE", "199"
       .Add "WRITE_ERROR_RATE", "200"
       .Add "DISK_SHIFT", "220"
       .Add "G_SENSE_ERROR_RATEII", "221"
       .Add "LOADED_HOURS", "222"
       .Add "LOAD_UNLOAD_RETRY_COUNT", "223"
       .Add "LOAD_FRICTION", "224"
       .Add "LOAD_UNLOAD_CYCLE_COUNTII", "225"
       .Add "LOAD_IN_TIME", "226"
       .Add "TORQUE_AMPLIFICATION_COUNT", "227"
       .Add "POWER_OFF_RETRACT_COUNT", "228"
       .Add "GMR_HEAD_AMPLITUDE", "230"
       .Add "TEMPERATUREII", "231"
       .Add "READ_ERROR_RETRY_RATE", "250"
   End With
End Sub

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
Function StripControlChars(Source As String, Optional KeepCRLF As Boolean = True) As String
Dim Index As Long
Dim Bytes() As Byte

' the fastest way to process this string
' is copy it into an array of Bytes
Bytes() = Source
For Index = 0 To UBound(Bytes) Step 2
    ' if this is a control character
    If Bytes(Index) < 32 And Bytes(Index + 1) = 0 Then
        If Not KeepCRLF Or (Bytes(Index) <> 13 And Bytes(Index) <> 10) Then
            ' the user asked to trim CRLF or this
            ' character isn't a CR or a LF, so clear it
            Bytes(Index) = 0
        End If
    End If
Next

' return this string, after filtering out all null chars
StripControlChars = VBA.Replace(Bytes(), vbNullChar, "")
End Function

'===============================================================================
' Name: Function GetDriveInfo
' Input:
'   ByRef drvNumber As IDE_DRIVE_NUMBER - IDE Drive type object
' Output:
'   DRIVE_INFO - Drive properties object
' Purpose: Returns the drive properties object structure
' Remarks: None
'===============================================================================
Public Function GetDriveInfo(DriveNum As IDE_DRIVE_NUMBER) As DRIVE_INFO
    Dim hDrive As Long
    Dim VerParam As GETVERSIONOUTPARAMS
    Dim cb As Long
    di.bDriveType = 0
    di.NumAttributes = 0
    ReDim di.Attributes(0)
    hDrive = OpenSmart(DriveNum)
    If hDrive = INVALID_HANDLE_VALUE Then Exit Function
    If Not GetSmartVersion(hDrive, VerParam) Then Exit Function
    'If Not IsBitSet(VerParam.bIDEDeviceMap, DriveNum) Then Exit Function
    di.bDriveType = 1 + Abs(IsBitSet(VerParam.bIDEDeviceMap, DriveNum + 4))
    If Not CheckSMARTEnable(hDrive, DriveNum) Then Exit Function
    FillAttrNameCollection
    Call IdentifyDrive(hDrive, IDE_ID_FUNCTION, DriveNum)
    Call ReadAttributesCmd(hDrive, DriveNum)
    Call ReadThresholdsCmd(hDrive, DriveNum)
    GetDriveInfo = di
    CloseHandle hDrive
    Set colAttrNames = Nothing
End Function

'===============================================================================
' Name: Function IdentifyDrive
' Input:
'   ByVal hDrive As Long - SMART drive handle
'   ByVal IDCmd As Byte - Command that can either be IDE identify or ATAPI identify
'   ByRef drvNumber As IDE_DRIVE_NUMBER - IDE Drive type object
'   ByRef di As DRIVE_INFO - Drive properties object
' Output:
'   Boolean - True if successful
' Purpose: Sends an IDENTIFY command to the drive
' Remarks: drvNumber = 0-3, IDCmd = IDE_ID_FUNCTION or IDE_ATAPI_ID
'===============================================================================
Private Function IdentifyDrive(ByVal hDrive As Long, ByVal IDCmd As Byte, ByVal DriveNum As IDE_DRIVE_NUMBER) As Boolean
    Dim SCIP As SENDCMDINPARAMS
    Dim IDSEC As IDSECTOR
    Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
    Dim sMsg As String
    Dim lpcbBytesReturned As Long
    Dim barrfound(100) As Long
    Dim i As Long
    Dim lng As Long
'   Set up data structures for IDENTIFY command.
    With SCIP
        .cBufferSize = IDENTIFY_BUFFER_SIZE
        .bDriveNumber = CByte(DriveNum)
        With .irDriveRegs
             .bFeaturesReg = 0
             .bSectorCountReg = 1
             .bSectorNumberReg = 1
             .bCylLowReg = 0
             .bCylHighReg = 0
' Compute the drive number.
             .bDriveHeadReg = &HA0
             If Not IsWinNT4Plus Then .bDriveHeadReg = .bDriveHeadReg Or (DriveNum And 1) * 16
' The command can either be IDE identify or ATAPI identify.
             .bCommandReg = CByte(IDCmd)
        End With
    End With
    If DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, lpcbBytesReturned, ByVal 0&) Then
       IdentifyDrive = True
       CopyMemory IDSEC, bArrOut(16), Len(IDSEC)
       di.Model = SwapStringBytes(StrConv(IDSEC.sModelNumber, vbUnicode))
       di.FirmWare = SwapStringBytes(StrConv(IDSEC.sFirmwareRev, vbUnicode))
       di.SerialNumber = SwapStringBytes(StrConv(IDSEC.sSerialNumber, vbUnicode))
       di.Cilinders = IDSEC.wNumCyls
       di.Heads = IDSEC.wNumHeads
       di.SecPerTrack = IDSEC.wSectorsPerTrack
    End If
End Function

Private Function IsBitSet(iBitString As Byte, ByVal lBitNo As Integer) As Boolean
    If lBitNo = 7 Then
        IsBitSet = iBitString < 0
    Else
        IsBitSet = iBitString And (2 ^ lBitNo)
    End If
End Function

'===============================================================================
' Name: Function SmartCheckEnabled
' Input:
'   ByVal hDrive As Long - SMART drive handle
'   ByRef drvNumber As IDE_DRIVE_NUMBER - IDE Drive type structure
' Output:
'   Boolean - Returns True if SMART enabled
' Purpose: Sends a SMART_ENABLE_SMART_OPERATIONS command to the drive
' Remarks: bDriveNum = 0-3
'===============================================================================
Private Function SmartCheckEnabled(ByVal hDrive As Long, _
                                   drvNumber As IDE_DRIVE_NUMBER) As Boolean
   
   Dim SCIP As SENDCMDINPARAMS
   Dim SCOP As SENDCMDOUTPARAMS
   Dim cbBytesReturned As Long
   
   With SCIP
   
      .cBufferSize = 0
      
      With .irDriveRegs
           .bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
           .bSectorCountReg = 1
           .bSectorNumberReg = 1
           .bCylLowReg = SMART_CYL_LOW
           .bCylHighReg = SMART_CYL_HI

           .bDriveHeadReg = &HA0
            If Not IsWinNT4Plus Then
               .bDriveHeadReg = .bDriveHeadReg Or ((drvNumber And 1) * 16)
            End If
           .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
           
       End With
       
       .bDriveNumber = drvNumber
       
   End With
   
   SmartCheckEnabled = DeviceIoControl(hDrive, _
                                      DFP_SEND_DRIVE_COMMAND, _
                                      SCIP, _
                                      Len(SCIP) - 4, _
                                      SCOP, _
                                      Len(SCOP) - 4, _
                                      cbBytesReturned, _
                                      ByVal 0&)
End Function


'===============================================================================
' Name: Function SmartGetVersion
' Input:
'   ByVal hDrive As Long - SMART drive handle
' Output:
'   Boolean - True if successful
' Purpose: Given the SMART drive handle, gets the version
' Remarks: None
'===============================================================================
Private Function SmartGetVersion(ByVal hDrive As Long) As Boolean
   
   Dim cbBytesReturned As Long
   Dim GVOP As GETVERSIONOUTPARAMS
   
   SmartGetVersion = DeviceIoControl(hDrive, _
                                     DFP_GET_VERSION, _
                                     ByVal 0&, 0, _
                                     GVOP, _
                                     Len(GVOP), _
                                     cbBytesReturned, _
                                     ByVal 0&)
   
End Function


'===============================================================================
' Name: Function OpenSmart
' Input:
'   ByRef drvNumber As IDE_DRIVE_NUMBER - IDE Drive type structure
' Output:
'   Long - SMART handle
' Purpose: Open SMART to allow DeviceIoControl communications and return SMART handle
' Remarks: None
'===============================================================================
Private Function OpenSmart(drv_num As IDE_DRIVE_NUMBER) As Long
   If IsWinNT4Plus Then
      OpenSmart = CreateFile("\\.\PhysicalDrive" & CStr(drv_num), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)
   Else
      OpenSmart = CreateFile("\\.\SMARTVSD", 0, 0, ByVal 0&, CREATE_NEW, 0, 0)
   End If
End Function
'****************************************************************************
' ReadAttributesCmd
' FUNCTION: Send a SMART_READ_ATTRIBUTE_VALUES command to the drive
' bDriveNum = 0-3
'***************************************************************************}
Private Function ReadAttributesCmd(ByVal hDrive As Long, DriveNum As IDE_DRIVE_NUMBER) As Boolean
   Dim cbBytesReturned As Long
   Dim SCIP As SENDCMDINPARAMS
   Dim drv_attr As DRIVEATTRIBUTE
   Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
   Dim sMsg As String
   Dim i As Long
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
            .bDriveHeadReg = &HA0
            If Not IsWinNT4Plus Then .bDriveHeadReg = .bDriveHeadReg Or (DriveNum And 1) * 16
            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
       End With
  End With
  ReadAttributesCmd = DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, cbBytesReturned, ByVal 0&)
  On Error Resume Next
  For i = 0 To NUM_ATTRIBUTE_STRUCTS - 1
      If bArrOut(18 + i * 12) > 0 Then
         di.Attributes(di.NumAttributes).AttrID = bArrOut(18 + i * 12)
         di.Attributes(di.NumAttributes).AttrName = "Unknown value (" & bArrOut(18 + i * 12) & ")"
         di.Attributes(di.NumAttributes).AttrName = colAttrNames(CStr(bArrOut(18 + i * 12)))
         di.NumAttributes = di.NumAttributes + 1
         ReDim Preserve di.Attributes(di.NumAttributes)
         CopyMemory di.Attributes(di.NumAttributes).StatusFlags, bArrOut(19 + i * 12), 2
         di.Attributes(di.NumAttributes).AttrValue = bArrOut(21 + i * 12)
         di.Attributes(di.NumAttributes).WorstValue = bArrOut(22 + i * 12)
      End If
  Next i
End Function
Private Function ReadThresholdsCmd(ByVal hDrive As Long, DriveNum As IDE_DRIVE_NUMBER) As Boolean
   Dim cbBytesReturned As Long
   Dim SCIP As SENDCMDINPARAMS
   Dim IDSEC As IDSECTOR
   Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
   Dim sMsg As String
   Dim thr_attr As ATTRTHRESHOLD
   Dim i As Long, j As Long
   With SCIP
 ' Set up data structures for Read Attributes SMART Command.
       .cBufferSize = READ_THRESHOLD_BUFFER_SIZE
       .bDriveNumber = DriveNum
       With .irDriveRegs
            .bFeaturesReg = SMART_READ_ATTRIBUTE_THRESHOLDS
            .bSectorCountReg = 1
            .bSectorNumberReg = 1
            .bCylLowReg = SMART_CYL_LOW
            .bCylHighReg = SMART_CYL_HI
'  Compute the drive number.
            .bDriveHeadReg = &HA0
            If Not IsWinNT4Plus Then .bDriveHeadReg = .bDriveHeadReg Or (DriveNum And 1) * 16
            .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
       End With
  End With
  ReadThresholdsCmd = DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Len(SCIP) - 4, bArrOut(0), OUTPUT_DATA_SIZE, cbBytesReturned, ByVal 0&)
  For i = 0 To NUM_ATTRIBUTE_STRUCTS - 1
      CopyMemory thr_attr, bArrOut(18 + i * Len(thr_attr)), Len(thr_attr)
      If thr_attr.bAttrID > 0 Then
         For j = 0 To UBound(di.Attributes)
             If thr_attr.bAttrID = di.Attributes(j).AttrID Then
                di.Attributes(j).ThresholdValue = thr_attr.bWarrantyThreshold
                Exit For
             End If
         Next j
      End If
  Next i
End Function

'===============================================================================
' Name: Function SwapStringBytes
' Input:
'   ByVal sIn As String - Input string
' Output:
'   String - Swapped string
' Purpose: Swaps byte arrays
' Remarks: None
'===============================================================================
Private Function SwapStringBytes(ByVal sIn As String) As String
   Dim sTemp As String
   Dim i As Integer
   sTemp = Space(Len(sIn))
   For i = 1 To Len(sIn) - 1 Step 2
       Mid(sTemp, i, 1) = Mid(sIn, i + 1, 1)
       Mid(sTemp, i + 1, 1) = Mid(sIn, i, 1)
   Next i
   SwapStringBytes = sTemp
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

' v3.5 introduces the use of the GetIfTable method which
' works very fast
If IsWinNT4Plus = True Then
    GetMACAddress = GetMACs_IfTable
    If GetMACAddress <> "" Then Exit Function
End If

'On Error Resume Next
'Dim tmp As String
'Dim pASTAT As Long
'Dim NCB As NET_CONTROL_BLOCK
'Dim AST As ASTAT
'
''The IBM NetBIOS 3.0 specifications defines four basic
''NetBIOS environments under the NCBRESET command. Win32
''follows the OS/2 Dynamic Link Routine (DLR) environment.
''This means that the first NCB issued by an application
''must be a NCBRESET, with the exception of NCBENUM.
''The Windows NT implementation differs from the IBM
''NetBIOS 3.0 specifications in the NCB_CALLNAME field.
'NCB.ncb_command = NCBRESET
'Call Netbios(NCB)
'
''To get the Media Access Control (MAC) address for an
''ethernet adapter programmatically, use the Netbios()
''NCBASTAT command and provide a "*" as the name in the
''NCB.ncb_CallName field (in a 16-chr string).
'NCB.ncb_callname = "*               "
'NCB.ncb_command = NCBASTAT
'
''For machines with multiple network adapters you need to
''enumerate the LANA numbers and perform the NCBASTAT
''command on each. Even when you have a single network
''adapter, it is a good idea to enumerate valid LANA numbers
''first and perform the NCBASTAT on one of the valid LANA
''numbers. It is considered bad programming to hardcode the
''LANA number to 0 (see the comments section below).
'NCB.ncb_lana_num = 0
'NCB.ncb_length = Len(AST)
'
'pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS _
'        Or HEAP_ZERO_MEMORY, NCB.ncb_length)
'
'If pASTAT = 0 Then
'    Debug.Print "memory allocation failed!"
'    Exit Function
'End If
'
'NCB.ncb_buffer = pASTAT
'Call Netbios(NCB)
'
'CopyMemory AST, NCB.ncb_buffer, Len(AST)
'
'tmp = Right$("00" & Hex(AST.adapt.adapter_address(0)), 2) & " " & _
'        Right$("00" & Hex(AST.adapt.adapter_address(1)), 2) & " " & _
'        Right$("00" & Hex(AST.adapt.adapter_address(2)), 2) & " " & _
'        Right$("00" & Hex(AST.adapt.adapter_address(3)), 2) & " " & _
'        Right$("00" & Hex(AST.adapt.adapter_address(4)), 2) & " " & _
'        Right$("00" & Hex(AST.adapt.adapter_address(5)), 2)
'
'HeapFree GetProcessHeap(), 0, pASTAT
'GetMACAddress = tmp

'another user provided the code below that seems to work well
'if the adapter card is "Ethernet 802.3", then the code below will work
Dim objset, obj
Set objset = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_NetworkAdapter")
For Each obj In objset
    On Error Resume Next
    If Not IsNull(obj.MACAddress) Then
        If obj.AdapterType = "Ethernet 802.3" Then
            If InStr(obj.PNPDeviceID, "PCI\") <> 0 Then
                GetMACAddress = Replace(obj.MACAddress, ":", " ")
                Exit Function
            End If
        End If
    End If
Next

GetMACAddressError:
If GetMACAddress = "" Then
    GetMACAddress = "Not Available"
End If

End Function

'-----------------------------------------------------------------------------------
' Get the system's MAC address(es) via GetIfTable API function (IPHLPAPI.DLL)
'
' Note: GetIfTable returns information also about the virtual loopback adapter
'-----------------------------------------------------------------------------------
Public Function GetMACs_IfTable() As String

    Dim NumAdapts As Long, nRowSize As Long, i%, retStr As String
    Dim IfInfo As MIB_IFROW, IPinfoBuf() As Byte, bufLen As Long, sts As Long

    ' Get # of interfaces defined (sometimes 1 more than GetIfTable)
    sts = GetNumberOfInterfaces(NumAdapts)

    ' Get size of buffer to allocate
    sts = GetIfTable(ByVal 0&, bufLen, 1)
    If (bufLen = 0) Then Exit Function

    ' reserve byte buffer & get it filled with adapter information
    ReDim IPinfoBuf(0 To bufLen - 1) As Byte
    sts = GetIfTable(IPinfoBuf(0), bufLen, 1)
    If (sts <> 0) Then Exit Function

    NumAdapts = IPinfoBuf(0)
    nRowSize = Len(IfInfo)
    'retStr = NumAdapts & " Interface(s):" & vbCrLf

    For i = 1 To NumAdapts
        ' copy one IfRow chunk of byte data into an MIB_IFROW structure
        Call CopyMemory(IfInfo, IPinfoBuf(4 + (i - 1) * nRowSize), nRowSize)

        ' Take adapter address if correct type
        With IfInfo
            'retStr = retStr & vbCrLf & "[" & i & "] " & Left$(.bDescr, .dwDescrLen - 1) & vbCrLf
            If (.dwType = MIB_IF_TYPE_ETHERNET And .dwOperStatus = MIB_IF_OPER_STATUS_OPERATIONAL) Then
                retStr = Replace(MAC2String(.bPhysAddr), "-", " ") 'retStr & vbTab & MAC2String(.bPhysAddr) & vbCrLf
            End If
        End With
    Next i

    GetMACs_IfTable = retStr

End Function

' Convert a byte array containing a MAC address to a hex string
Private Function MAC2String(AdrArray() As Byte) As String
    Dim aStr As String, hexStr As String, i%
    
    For i = 0 To 5
        If (i > UBound(AdrArray)) Then
            hexStr = "00"
        Else
            hexStr = Hex$(AdrArray(i))
        End If
        
        If (Len(hexStr) < 2) Then hexStr = "0" & hexStr
        aStr = aStr & hexStr
        If (i < 5) Then aStr = aStr & "-"
    Next i
    
    MAC2String = aStr
    
End Function

'===============================================================================
' Name: Function GeetBiosVersion
' Input: None
' Output:
'   String - BIOS serial number
' Purpose: Gets the BIOS Serial Number
' Remarks: Uses the WMI
'===============================================================================
Public Function GeetBiosVersion() As String
Dim BiosSet As Object
Dim obj As Object
On Error GoTo GeetBiosVersionError
Set BiosSet = GetObject("WinMgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_BIOS")
For Each obj In BiosSet
    GeetBiosVersion = obj.Version
    If GeetBiosVersion <> "" Then Exit Function
Next
GeetBiosVersionError:
If GeetBiosVersion = "" Then
    GeetBiosVersion = "Not Available"
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
On Error GoTo GetMotherboardSerialError

Set MotherboardSet = GetObject("WinMgmts:{impersonationLevel=impersonate}").InstancesOf("CIM_Chassis")
For Each obj In MotherboardSet
    GetMotherboardSerial = obj.SerialNumber
    If GetMotherboardSerial <> "" Then
        ' Strip any periods
        Dim Bytes() As Byte
        Bytes() = GetMotherboardSerial
        GetMotherboardSerial = VBA.Replace(Bytes(), ".", "")
        Exit Function
    End If
Next

GetMotherboardSerialError:
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

Dim sHostName As String * 256
Dim lpHost As Long
Dim HOST As HOSTENT
Dim dwIPAddr As Long
Dim tmpIPAddr() As Byte
Dim i As Integer
Dim sIPAddr As String
If Not SocketsInitialize() Then
    GetIPaddress = ""
    GoTo GetIPaddressError
End If
If gethostname(sHostName, 256) = SOCKET_ERROR Then
    GetIPaddress = ""
    'MsgBox "Windows Sockets error " & str$(WSAGetLastError()) & " has occurred. Unable to successfully get Host Name."
    SocketsCleanup
    GoTo GetIPaddressError
End If
sHostName = Trim$(sHostName)
lpHost = gethostbyname(sHostName)
If lpHost = 0 Then
    GetIPaddress = ""
    'MsgBox "Windows Sockets are not responding. " & "Unable to successfully get Host Name."
    SocketsCleanup
    GoTo GetIPaddressError
End If
CopyMemoryIP HOST, lpHost, Len(HOST)
CopyMemoryIP dwIPAddr, HOST.hAddrList, 4
ReDim tmpIPAddr(1 To HOST.hLen)
CopyMemoryIP tmpIPAddr(1), dwIPAddr, HOST.hLen
For i = 1 To HOST.hLen
    sIPAddr = sIPAddr & tmpIPAddr(i) & "."
Next
GetIPaddress = Mid$(sIPAddr, 1, Len(sIPAddr) - 1)
SocketsCleanup

' for external IP address, use the following
'GetIPaddress = GetExternalIP("http://checkip.dyndns.org/")

GetIPaddressError:
If GetIPaddress = "" Then
    GetIPaddress = "Not Available"
End If
End Function

Public Function HiByte(ByVal wParam As Integer)
    HiByte = wParam \ &H100 And &HFF&
End Function
Public Function LoByte(ByVal wParam As Integer)
    LoByte = wParam And &HFF&
End Function
Public Sub SocketsCleanup()
    If WSACleanup() <> ERROR_SUCCESS Then
        'MsgBox "Socket error occurred in Cleanup."
    End If
End Sub
Public Function SocketsInitialize() As Boolean
    Dim WSAD As WSAData
    Dim sLoByte As String
    Dim sHiByte As String
    If WSAStartup(WS_VERSION_REQD, WSAD) <> ERROR_SUCCESS Then
        'MsgBox "The 32-bit Windows Socket is not responding."
        SocketsInitialize = False
        Exit Function
    End If
    If WSAD.wMaxSockets < MIN_SOCKETS_REQD Then
        'MsgBox "This application requires a minimum of " & CStr(MIN_SOCKETS_REQD) & " supported sockets."
        SocketsInitialize = False
        Exit Function
    End If
    If LoByte(WSAD.wVersion) < WS_VERSION_MAJOR Or (LoByte(WSAD.wVersion) = WS_VERSION_MAJOR And HiByte(WSAD.wVersion) < WS_VERSION_MINOR) Then
        sHiByte = CStr(HiByte(WSAD.wVersion))
        sLoByte = CStr(LoByte(WSAD.wVersion))
        'MsgBox "Sockets version " & sLoByte & "." & sHiByte & " is not supported by 32-bit Windows Sockets."
        SocketsInitialize = False
        Exit Function
    End If
    'must be OK, so lets do it
    SocketsInitialize = True
End Function

Private Function GetHDSerialFirmwareVB6(ByVal controller As Integer, Optional ByVal masterDrive As Boolean = True)

' Created with the help of the following articles and clues from ALCrypto3.dll
' SOURCE 1: http://discuss.develop.com/archives/wa.exe?A2=ind0309a&L=advanced-dotnet&D=0&T=0&P=3760
' SOURCE 2: http://www.visual-basic.it/scarica.asp?ID=611
' SOURCE 3: ALCrypto3.dll and DISKID32 program

' This code DOES NOT require admin rights in the user's machine
' This code DOES NOT require WMI
' This code DOES NOT require SMART VXD drivers for Win95/98/Me

Dim myStr As String
Dim reversedStr As String, str1 As String, str2 As String
Dim kk As Integer, jj As Integer
Dim hdh As Long, dummy As Long, newHandle As Long

GetHDSerialFirmwareVB6 = ""
hdh = CreateFile("\\.\Scsi" & controller & ":", GENERIC_READ Or GENERIC_WRITE, _
    FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0&, OPEN_EXISTING, 0, 0)

If (hdh <> -1) Then

    Dim bin As BufferType, bout As BufferType
    ZeroMemory bin, Len(bin)
    ZeroMemory bout, Len(bout)
    
    bin.myBuffer(0) = 28
    bin.myBuffer(4) = 83
    bin.myBuffer(5) = 67
    bin.myBuffer(6) = 83
    bin.myBuffer(7) = 73
    bin.myBuffer(8) = 68
    bin.myBuffer(9) = 73
    bin.myBuffer(10) = 83
    bin.myBuffer(11) = 75
    bin.myBuffer(12) = 16
    bin.myBuffer(13) = 39
    bin.myBuffer(16) = 1
    bin.myBuffer(17) = 5
    bin.myBuffer(18) = 27
    bin.myBuffer(24) = 20  '17?
    bin.myBuffer(25) = 2
    bin.myBuffer(38) = 236 '&HEC
    
    If masterDrive = True Then
        bin.myBuffer(40) = 0 'master drive
    Else
        bin.myBuffer(40) = 1 'slave drive
    End If
    
    newHandle = DeviceIoControl(hdh, 315400, bin, Len(bin), bout, Len(bout), dummy, 0)
    If (newHandle) Then
        ' HDD Firmware Serial Number is between 64 to 83 - 19 digits as we had from ALCrypto before
        ' HDD Model Number is between 98 and 137
        ' HDD Controller Revision is between 90 and 97
        For jj = 64 To 83
            myStr = myStr & Chr(bout.myBuffer(jj))
        Next
        
        ' Seems like some swapping is needed at this point
        reversedStr = ""
        For jj = 0 To Len(myStr) / 2
            str1 = Mid(myStr, jj * 2 + 1, 1)
            str2 = Mid(myStr, jj * 2 + 2, 1)
            reversedStr = reversedStr & str2 & str1
        Next
        GetHDSerialFirmwareVB6 = StripControlChars(Trim(reversedStr), False)
    End If

End If

CloseHandle hdh
    
End Function
Private Function GetSmartVersion(ByVal hDrive As Long, VersionParams As GETVERSIONOUTPARAMS) As Boolean
   Dim cbBytesReturned As Long
   GetSmartVersion = DeviceIoControl(hDrive, DFP_GET_VERSION, ByVal 0&, 0, VersionParams, Len(VersionParams), cbBytesReturned, ByVal 0&)
End Function

