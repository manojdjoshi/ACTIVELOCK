using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
 // ERROR: Not supported in C#: OptionDeclaration
using System.Net;
using System.Text;
using System.Management;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Security;
using System.DirectoryServices;

static class modHardware
{
	[DllImport("kernel32", EntryPoint = "GetVersionExA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int GetVersionEx(ref OSVERSIONINFO LpVersionInformation);
	[DllImport("kernel32", EntryPoint = "CreateFileA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int CreateFile(string lpFileName, int dwDesiredAccess, int dwShareMode, int lpSecurityAttributes, int dwCreationDisposition, int dwFlagsAndAttributes, int hTemplateFile);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int CloseHandle(int hObject);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int DeviceIoControl(int hDevice, int dwIoControlCode, ref SENDCMDINPARAMS lpInBuffer, int nInBufferSize, ref int lpOutBuffer, int nOutBufferSize, ref int lpBytesReturned, int lpOverlapped);
	[DllImport("kernel32", EntryPoint = "RtlMoveMemory", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern void CopyMemory(ref int Destination, ref int Source, int Length);
	[DllImport("kernel32", EntryPoint = "CreateFileW", CharSet = CharSet.Unicode, SetLastError = true, ExactSpelling = true)]
	public static extern IntPtr CreateFile2(string lpFileName, int dwDesiredAccess, int dwShareMode, IntPtr lpSecurityAttributes, int dwCreationDisposition, int dwFlagsAndAttributes, IntPtr hTemplateFile);
	[DllImport("kernel32", EntryPoint = "CloseHandle", CharSet = CharSet.Unicode, SetLastError = true, ExactSpelling = true)]
	public static extern bool CloseHandle2(IntPtr hObject);
	[DllImport("kernel32", EntryPoint = "DeviceIoControl", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern bool DeviceIoControl2(IntPtr hDevice, int dwIoControlCode, IntPtr lpInBuffer, int nInBufferSize, IntPtr lpOutBuffer, int nOutBufferSize, ref int lpBytesReturned, IntPtr lpOverlapped);
	//*   ActiveLock
	//*   Copyright 1998-2002 Nelson Ferraz
	//*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
	//*   All material is the property of the contributing authors.
	//*
	//*   Redistribution and use in source and binary forms, with or without
	//*   modification, are permitted provided that the following conditions are
	//*   met:
	//*
	//*     [o] Redistributions of source code must retain the above copyright
	//*         notice, this list of conditions and the following disclaimer.
	//*
	//*     [o] Redistributions in binary form must reproduce the above
	//*         copyright notice, this list of conditions and the following
	//*         disclaimer in the documentation and/or other materials provided
	//*         with the distribution.
	//*
	//*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
	//*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
	//*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
	//*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
	//*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
	//*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
	//*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
	//*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
	//*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
	//*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
	//*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
	//*
	//*
	//===============================================================================
	// Name: modHardware
	// Purpose: Gets all the hardware signatures of the current machine
	// Date Created:
	// Functions:
	// Properties:
	// Methods:
	// Started: 08.15.2005
	// Modified: 03.25.2006
	//===============================================================================

	//****** SMART DECLARATIONS ******


	static Collection colAttrNames;
	//---------------------------------------------------------------------
	// The following structure defines the structure of a Drive Attribute
	//---------------------------------------------------------------------

	public const short NUM_ATTRIBUTE_STRUCTS = 30;
	public struct DRIVEATTRIBUTE
	{
			// Identifies which attribute
		public byte bAttrID;
			//Integer ' see bit definitions below
		public short wStatusFlags;
			// Current normalized value
		public byte bAttrValue;
			// How bad has it ever been?
		public byte bWorstValue;
		[VBFixedArray(5)]
			// Un-normalized value
		public byte[] bRawValue;
			// ...
		public byte bReserved;
		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

		}
	}
	//---------------------------------------------------------------------
	// The following structure defines the structure of a Warranty Threshold
	// Obsoleted in ATA4!
	//---------------------------------------------------------------------
	//Public Structure ATTRTHRESHOLD
	//    Dim bAttrID As Byte ' Identifies which attribute
	//    Dim bWarrantyThreshold As Byte ' Triggering value
	//    <VBFixedArray(9)> Dim bReserved() As Byte
	//    Public Sub Initialize()
	//        ReDim bReserved(9)
	//    End Sub
	//End Structure

	//---------------------------------------------------------------------
	// Valid Attribute IDs
	//---------------------------------------------------------------------
	public enum ATTRIBUTE_ID
	{
		ATTR_INVALID = 0,
		ATTR_READ_ERROR_RATE = 1,
		ATTR_THROUGHPUT_PERF = 2,
		ATTR_SPIN_UP_TIME = 3,
		ATTR_START_STOP_COUNT = 4,
		ATTR_REALLOC_SECTOR_COUNT = 5,
		ATTR_READ_CHANNEL_MARGIN = 6,
		ATTR_SEEK_ERROR_RATE = 7,
		ATTR_SEEK_TIME_PERF = 8,
		ATTR_POWER_ON_HRS_COUNT = 9,
		ATTR_SPIN_RETRY_COUNT = 10,
		ATTR_CALIBRATION_RETRY_COUNT = 11,
		ATTR_POWER_CYCLE_COUNT = 12,
		ATTR_SOFT_READ_ERROR_RATE = 13,
		ATTR_G_SENSE_ERROR_RATE = 191,
		ATTR_POWER_OFF_RETRACT_CYCLE = 192,
		ATTR_LOAD_UNLOAD_CYCLE_COUNT = 193,
		ATTR_TEMPERATURE = 194,
		ATTR_REALLOCATION_EVENTS_COUNT = 196,
		ATTR_CURRENT_PENDING_SECTOR_COUNT = 197,
		ATTR_UNCORRECTABLE_SECTOR_COUNT = 198,
		ATTR_ULTRADMA_CRC_ERROR_RATE = 199,
		ATTR_WRITE_ERROR_RATE = 200,
		ATTR_DISK_SHIFT = 220,
		ATTR_G_SENSE_ERROR_RATEII = 221,
		ATTR_LOADED_HOURS = 222,
		ATTR_LOAD_UNLOAD_RETRY_COUNT = 223,
		ATTR_LOAD_FRICTION = 224,
		ATTR_LOAD_UNLOAD_CYCLE_COUNTII = 225,
		ATTR_LOAD_IN_TIME = 226,
		ATTR_TORQUE_AMPLIFICATION_COUNT = 227,
		ATTR_POWER_OFF_RETRACT_COUNT = 228,
		ATTR_GMR_HEAD_AMPLITUDE = 230,
		ATTR_TEMPERATUREII = 231,
		ATTR_READ_ERROR_RETRY_RATE = 250
	}

	//***** SMART DECLARATIONS *****

	//HDD firmware serial number
	private const int GENERIC_READ = 0x80000000;
	private const int GENERIC_WRITE = 0x40000000;
	private const short FILE_SHARE_READ = 0x1;
	private const short FILE_SHARE_WRITE = 0x2;
	private const short OPEN_EXISTING = 3;
	private const short CREATE_NEW = 1;
	private const short INVALID_HANDLE_VALUE = -1;
	private const short VER_PLATFORM_WIN32_NT = 2;
	private const short IDENTIFY_BUFFER_SIZE = 512;
	public const short READ_THRESHOLD_BUFFER_SIZE = 512;

	private const int OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16;
	[DllImport("kernel32", EntryPoint = "GetVolumeInformationA", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int apiGetVolumeInformation(string lpRootPathName, string lpVolumeNameBuffer, int nVolumeNameSize, ref int lpVolumeSerialNumber, ref int lpMaximumComponentLength, ref int lpFileSystemFlags, string lpFileSystemNameBuffer, int nFileSystemNameSize);

	//GETVERSIONOUTPARAMS contains the data returned
	//from the Get Driver Version function
	private struct GETVERSIONOUTPARAMS
	{
			//Binary driver version.
		public byte bVersion;
			//Binary driver revision
		public byte bRevision;
			//Not used
		public byte bReserved;
			//Bit map of IDE devices
		public byte bIDEDeviceMap;
			//Bit mask of driver capabilities
		public int fCapabilities;
		[VBFixedArray(3)]
			//For future use
		public int[] dwReserved;
		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

		}
	}

	//IDE registers
	private struct IDEREGS
	{
			//Used for specifying SMART "commands"
		public byte bFeaturesReg;
			//IDE sector count register
		public byte bSectorCountReg;
			//IDE sector number register
		public byte bSectorNumberReg;
			//IDE low order cylinder value
		public byte bCylLowReg;
			//IDE high order cylinder value
		public byte bCylHighReg;
			//IDE drive/head register
		public byte bDriveHeadReg;
			//Actual IDE command
		public byte bCommandReg;
			//reserved for future use - must be zero
		public byte bReserved;
	}

	//SENDCMDINPARAMS contains the input parameters for the
	//Send Command to Drive function
	private struct SENDCMDINPARAMS
	{
			//Buffer size in bytes
		public int cBufferSize;
			//Structure with drive register values.
		public IDEREGS irDriveRegs;
			//Physical drive number to send command to (0,1,2,3).
		public byte bDriveNumber;
		[VBFixedArray(2)]
			//Bytes reserved
		public byte[] bReserved;
		[VBFixedArray(3)]
			//DWORDS reserved
		public int[] dwReserved;
			//Input buffer.
		public byte[] bBuffer;
		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

			 // ERROR: Not supported in C#: ReDimStatement

		}
	}

	//Valid values for the bCommandReg member of IDEREGS.
		// Returns ID sector for ATAPI.
	private const short IDE_ATAPI_ID = 0xa1;
		//Returns ID sector for ATA.
	private const short IDE_ID_FUNCTION = 0xec;
		//Performs SMART cmd.
	private const short IDE_EXECUTE_SMART_FUNCTION = 0xb0;
	//Requires valid bFeaturesReg,
	//bCylLowReg, and bCylHighReg

	//Cylinder register values required when issuing SMART command
	private const short SMART_CYL_LOW = 0x4f;

	private const short SMART_CYL_HI = 0xc2;
	//Status returned from driver
	private struct DRIVERSTATUS
	{
			//Error code from driver, or 0 if no error
		public byte bDriverError;
			//Contents of IDE Error register
		public byte bIDEStatus;
		//Only valid when bDriverError is SMART_IDE_ERROR
		[VBFixedArray(1)]
		public byte[] bReserved;
		[VBFixedArray(1)]
		public int[] dwReserved;
		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

			 // ERROR: Not supported in C#: ReDimStatement

		}
	}

	private struct IDSECTOR
	{
		public short wGenConfig;
		public short wNumCyls;
		public short wReserved;
		public short wNumHeads;
		public short wBytesPerTrack;
		public short wBytesPerSector;
		public short wSectorsPerTrack;
		[VBFixedArray(2)]
		public short[] wVendorUnique;
		[VBFixedArray(19)]
		public byte[] sSerialNumber;
		public short wBufferType;
		public short wBufferSize;
		public short wECCSize;
		[VBFixedArray(7)]
		public byte[] sFirmwareRev;
		[VBFixedArray(39)]
		public byte[] sModelNumber;
		public short wMoreVendorUnique;
		public short wDoubleWordIO;
		public short wCapabilities;
		public short wReserved1;
		public short wPIOTiming;
		public short wDMATiming;
		public short wBS;
		public short wNumCurrentCyls;
		public short wNumCurrentHeads;
		public short wNumCurrentSectorsPerTrack;
		public int ulCurrentSectorCapacity;
		public short wMultSectorStuff;
		public int ulTotalAddressableSectors;
		public short wSingleWordDMA;
		public short wMultiWordDMA;
		[VBFixedArray(127)]
		public byte[] bReserved;
		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

			 // ERROR: Not supported in C#: ReDimStatement

			 // ERROR: Not supported in C#: ReDimStatement

			 // ERROR: Not supported in C#: ReDimStatement

			 // ERROR: Not supported in C#: ReDimStatement

		}
	}

	//Structure returned by SMART IOCTL commands
	private struct SENDCMDOUTPARAMS
	{
			//Size of Buffer in bytes
		public int cBufferSize;
			//Driver status structure
		public DRIVERSTATUS DRIVERSTATUS;
			//Buffer of arbitrary length for data read from drive
		public byte[] bBuffer;
		public void Initialize()
		{
			DRIVERSTATUS.Initialize();
		}
	}

	//Vendor specific feature register defines
	//for SMART "sub commands"
	private const short SMART_READ_ATTRIBUTE_VALUES = 0xd0;
	private const short SMART_READ_ATTRIBUTE_THRESHOLDS = 0xd1;
	private const short SMART_ENABLE_DISABLE_ATTRIBUTE_AUTOSAVE = 0xd2;
	private const short SMART_SAVE_ATTRIBUTE_VALUES = 0xd3;
	private const short SMART_EXECUTE_OFFLINE_IMMEDIATE = 0xd4;
	// Vendor specific commands:
	private const short SMART_ENABLE_SMART_OPERATIONS = 0xd8;
	private const short SMART_DISABLE_SMART_OPERATIONS = 0xd9;

	private const short SMART_RETURN_SMART_STATUS = 0xda;
	//Status Flags Values
	public enum STATUS_FLAGS
	{
		PRE_FAILURE_WARRANTY = 0x1,
		ON_LINE_COLLECTION = 0x2,
		PERFORMANCE_ATTRIBUTE = 0x4,
		ERROR_RATE_ATTRIBUTE = 0x8,
		EVENT_COUNT_ATTRIBUTE = 0x10,
		SELF_PRESERVING_ATTRIBUTE = 0x20
	}

	//IOCTL commands
	private const int DFP_GET_VERSION = 0x74080;
	private const int DFP_SEND_DRIVE_COMMAND = 0x7c084;

	private const int DFP_RECEIVE_DRIVE_DATA = 0x7c088;
	public struct ATTR_DATA
	{
		public byte AttrID;
		public string AttrName;
		public byte AttrValue;
		public byte ThresholdValue;
		public byte WorstValue;
		public STATUS_FLAGS StatusFlags;
	}

	private struct OSVERSIONINFO
	{
		public int dwOSVersionInfoSize;
		public int dwMajorVersion;
		public int dwMinorVersion;
		public int dwBuildNumber;
		public int dwPlatformId;
		[VBFixedString(128), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = 128)]
		public string szCSDVersion;
	}

	public struct DRIVE_INFO
	{
		public byte bDriveType;
		public string SerialNumber;
		public string Model;
		public string FirmWare;
		public int Cilinders;
		public int Heads;
		public int SecPerTrack;
		public int BytesPerSector;
		public int BytesperTrack;
		public byte NumAttributes;
		public ATTR_DATA[] Attributes;
	}

	static DRIVE_INFO di;
	public enum IDE_DRIVE_NUMBER
	{
		PRIMARY_MASTER,
		PRIMARY_SLAVE,
		SECONDARY_MASTER,
		SECONDARY_SLAVE,
		TERTIARY_MASTER,
		TERTIARY_SLAVE,
		QUARTIARY_MASTER,
		QUARTIARY_SLAVE
	}
	private struct BufferType
	{
		[VBFixedArray(559)]
		public byte[] myBuffer;
		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

		}
	}
	[DllImport("kernel32", EntryPoint = "RtlZeroMemory", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern void ZeroMemory(ref int dest, int numBytes);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int DeviceIoControl(int hDevice, int dwIoControlCode, ref object lpInBuffer, int nInBufferSize, ref object lpOutBuffer, int nOutBufferSize, ref int lpBytesReturned, ref int lpOverlapped);

	// The following UDT and the DLL function is for getting
	// the serial number from a C++ DLL in case the VB6 APIs fail
	// Currently, VB code cannot handle the serial numbers
	// coming from computers with non-admin rights; in those
	// cases the C++ DLL function "getHardDriveFirmware" should
	// work properly.
	// Neither of the two methods work for the SATA and SCSI drives
	// ialkan - 8312005
	private struct MyUDT2
	{
		[VBFixedString(30), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = 30)]
		public string myStr;
		public int mL;
	}
	[DllImport("ALCrypto3.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	private static extern int getHardDriveFirmware(ref MyUDT2 myU);

	//MAC Address
	public const int NCBASTAT = 0x33;
	public const int NCBNAMSZ = 16;
	public const int HEAP_ZERO_MEMORY = 0x8;
	public const int HEAP_GENERATE_EXCEPTIONS = 0x4;

	public const int NCBRESET = 0x32;
	public struct NET_CONTROL_BLOCK
	{
		//NCB
		public byte ncb_command;
		public byte ncb_retcode;
		public byte ncb_lsn;
		public byte ncb_num;
		public int ncb_buffer;
		public short ncb_length;
		[VBFixedString(modHardware.NCBNAMSZ), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = modHardware.NCBNAMSZ)]
		public string ncb_callname;
		[VBFixedString(modHardware.NCBNAMSZ), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = modHardware.NCBNAMSZ)]
		public string ncb_name;
		public byte ncb_rto;
		public byte ncb_sto;
		public int ncb_post;
		public byte ncb_lana_num;
		public byte ncb_cmd_cplt;
		[VBFixedArray(9)]
			// Reserved, must be 0
		public byte[] ncb_reserve;
		public int ncb_event;
		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

		}
	}

	public struct ADAPTER_STATUS
	{
		[VBFixedArray(5)]
		public byte[] adapter_address;
		public byte rev_major;
		public byte reserved0;
		public byte adapter_type;
		public byte rev_minor;
		public short duration;
		public short frmr_recv;
		public short frmr_xmit;
		public short iframe_recv_err;
		public short xmit_aborts;
		public int xmit_success;
		public int recv_success;
		public short iframe_xmit_err;
		public short recv_buff_unavail;
		public short t1_timeouts;
		public short ti_timeouts;
		public int Reserved1;
		public short free_ncbs;
		public short max_cfg_ncbs;
		public short max_ncbs;
		public short xmit_buf_unavail;
		public short max_dgram_size;
		public short pending_sess;
		public short max_cfg_sess;
		public short max_sess;
		public short max_sess_pkt_size;
		public short name_count;
		public void Initialize()
		{
			 // ERROR: Not supported in C#: ReDimStatement

		}
	}

	public struct NAME_BUFFER
	{
		[VBFixedString(modHardware.NCBNAMSZ), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = modHardware.NCBNAMSZ)]
		public string name;
		public short name_num;
		public short name_flags;
	}

	public struct ASTAT
	{
		public ADAPTER_STATUS adapt;
		[VBFixedArray(30)]
		public NAME_BUFFER[] NameBuff;
		public void Initialize()
		{
			adapt.Initialize();
			 // ERROR: Not supported in C#: ReDimStatement

		}
	}
	[DllImport("netapi32.dll", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern byte Netbios(ref NET_CONTROL_BLOCK pncb);
	[DllImport("kernel32", EntryPoint = "RtlMoveMemory", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern void CopyMemory(ref object hpvDest, int hpvSource, int cbCopy);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int GetProcessHeap();
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int HeapAlloc(int hHeap, int dwFlags, int dwBytes);
	[DllImport("kernel32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]
	public static extern int HeapFree(int hHeap, int dwFlags, int lpMem);

	//Structure NET_CONTROL_BLOCK may require marshalling attributes to be passed as an argument in this Declare statement

	private static FingerPrint fp = new FingerPrint();
	//===============================================================================
	// Name: Function GetComputerName
	// Input: None
	// Output:
	//   String - Computer name
	// Purpose: Gets the computer name on the network
	//===============================================================================
	public static string GetComputerName()
	{
		string functionReturnValue = null;
		 // ERROR: Not supported in C#: OnErrorStatement

		functionReturnValue = System.Environment.ExpandEnvironmentVariables("%ComputerName%");
		GetComputerNameError:
		if (string.IsNullOrEmpty(functionReturnValue)) {
			functionReturnValue = "Not Available";
		}
		return functionReturnValue;

	}

	//===============================================================================
	// Name: Function HDSerial
	// Input:
	//   ByRef path As String - Drive letter
	// Output:
	//   String - The serial number for the drive alock is on, formatted as "xxxx-xxxx"
	// Purpose: Function to return the serial number for a hard drive
	// Currently works on local drives, mapped drives, and shared drives.
	// Remarks: TODO: Decide what to to about shared folders and RAID arrays
	//===============================================================================

	public static string HDSerial(ref string path)
	{
		int lngDummy2 = 0;
		int lngReturn = 0;
		int lngDummy1 = 0;
		int lngSerial = 0;
		string strDummy2 = null;
		string strDummy1 = null;
		string strSerial = null;

		string strDriveLetter = null;
		int lngFirstSlash = 0;

		strDriveLetter = path;

		//(Just in case... It's better to be safe than sorry.)
		strDriveLetter = modTrial.Replace_Renamed(strDriveLetter, "/", "\\");

		//Check the drive type
		if (!(Strings.Left(strDriveLetter, 1) == "\\")) {
			//Good... The path is a local drive
			strDriveLetter = Strings.Left(strDriveLetter, 3);
		}
		else {
			//It's a network drive
			//This will return 0000-0000 on shared folders or RAID arrays
			//Shared drives work fine
			lngFirstSlash = Strings.InStr(3, strDriveLetter, "\\");
			strDriveLetter = Strings.Left(strDriveLetter, lngFirstSlash);
		}

		//Set up dimmies
		strDummy1 = Strings.Space(260);
		strDummy2 = Strings.Space(260);

		//Call the API function
		lngReturn = apiGetVolumeInformation(strDriveLetter, strDummy1, Strings.Len(strDummy1), ref lngSerial, ref lngDummy1, ref lngDummy2, strDummy2, Strings.Len(strDummy2));

		//Format the serial
		strSerial = Strings.Trim(Conversion.Hex(lngSerial));
		strSerial = new string("0", 8 - Strings.Len(strSerial)) + strSerial;
		strSerial = Strings.Left(strSerial, 4) + "-" + Strings.Right(strSerial, 4);
		return strSerial;

		// Alternative Code - Short Version
		//    Dim volbuf$, sysname$, sysflags&, componentlength&
		//    Dim serialnum As Long
		//    GetVolumeInformation "C:\", volbuf$, 255, serialnum, componentlength, sysflags, sysname$, 255
		//    HDSerial = CStr(serialnum)

	}

	//===============================================================================
	// Name: Function GetHDSerial
	// Input: None
	// Output:
	//   String - The serial number for the drive alock is on, formatted as "xxxx-xxxx"
	// Purpose: Function to return the serial number for a hard drive
	//   Currently works on local drives, mapped drives, and shared drives.
	//   Checks windir if it cant get a serial, then c:, then returns 0000-0000
	// Remarks: I think that this is 99.999999897456284893% effective.
	//===============================================================================
	public static string GetHDSerial()
	{
		string functionReturnValue = null;
		string strSerial = null;

		strSerial = HDSerial(ref ActiveLock3_6NET.My.MyProject.Application.Info.DirectoryPath);

		if (strSerial == "0000-0000") {
			//Calculate WINDIR drive if couldn't retrieve app.path serial
			strSerial = HDSerial(ref modActiveLock.WinDir());
		}

		if (strSerial == "0000-0000") {
			//If it still can't get a serial, revert to c:.
			//If no c: is present (or c: is RAID), 0000-0000 is returned
			strSerial = HDSerial(ref "C:\\");
		}

		functionReturnValue = strSerial;
		GetHDSeriAlerror:

		if (string.IsNullOrEmpty(functionReturnValue)) {
			functionReturnValue = "Not Available";
		}
		return functionReturnValue;
	}

	//===============================================================================
	// Name: Function GetHDSerialFirmware
	// Input: None
	// Output:
	//   String - HDD Firmware Serial Number
	// Purpose: Function to return the HDD Firmware Serial Number (Actual Physical Serial Number)
	// Remarks: None
	//===============================================================================
	public static string GetHDSerialFirmware()
	{
		string functionReturnValue = null;
		short jj = 0;
		int drvNumber = 0;
		// We just need the Primary Master Drive ID - ialkan 8312005
		MyUDT2 mU = null;
		string a = null;
		 // ERROR: Not supported in C#: OnErrorStatement

		functionReturnValue = "";
		a = "";

		//************** METHOD 1 **************
		// ialkan 2-12-06
		// Pure VB6 version of the code found in several online resources
		// described in GetHDSerialFirmwareVB6 function
		// This eliminates the dependency of the HDD firmware serial number
		// function from ALCrypto3.dll
		// Controller index
		for (jj = 0; jj <= 15; jj++) {
			a = GetHDSerialFirmwareVBNET(jj, true);
			// Check the Master drive
			if (!string.IsNullOrEmpty(a)) functionReturnValue = a.Trim(); 
			if (!string.IsNullOrEmpty(functionReturnValue)) {
				return;
			}
			a = GetHDSerialFirmwareVBNET(jj, false);
			// Now check the Slave Drive
			if (!string.IsNullOrEmpty(a)) functionReturnValue = a.Trim(); 
			if (!string.IsNullOrEmpty(functionReturnValue)) {
				return;
			}
		}

		//************** METHOD 2 **************
		// Still nothing... Use ALCrypto DLL
		getHardDriveFirmware(ref mU);
		if (!string.IsNullOrEmpty(mU.myStr)) a = StripControlChars(ref mU.myStr, ref false); 
		if (!string.IsNullOrEmpty(a)) functionReturnValue = a.Trim(); 
		if (!string.IsNullOrEmpty(functionReturnValue)) {
			return;
		}

		//************** METHOD 3 **************
		a = mSmartCall.GetDriveInfo(IDE_DRIVE_NUMBER.PRIMARY_MASTER);
		if (!string.IsNullOrEmpty(a)) functionReturnValue = a.Trim(); 
		if (!string.IsNullOrEmpty(functionReturnValue)) {
			return;
		}

		//************** METHOD 4 **************
		a = GetHDSerialFirmwareWMI();
		if (!string.IsNullOrEmpty(a)) functionReturnValue = a.Trim(); 
		if (!string.IsNullOrEmpty(functionReturnValue)) {
			return;
		}

		return;
		GetHDSerialFirmwareError:

		// Well, this is not so good, because we still don't have
		// a serial number in our hands...
		// Cannot return an empty string...
		Interaction.MsgBox(Err().Description);
		if (string.IsNullOrEmpty(functionReturnValue)) {
			//GetHDSerialFirmware = "Not Available"
			// Per suggestion by Jeroen, we must have something decent returned from this
			functionReturnValue = "NA" + GetHDSerial() + GetMotherboardSerial();
			//"Not Available"
		}
		return functionReturnValue;

	}

	//===============================================================================
	// Name: Function StripControlChars
	// Input:
	//   ByVal source As String - String to be stripped off the control characters
	//   ByVal KeepCRLF As Boolean - If the second argument is True or omitted, CR-LF pairs are preserved
	// Output:
	//   String - String stripped off the control characters
	// Purpose: Strips all control characters (ASCII code < 32)
	// Remarks: None
	//===============================================================================
	public static string StripControlChars(ref string Source, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(true)] ref  // ERROR: Optional parameters aren't supported in C#
bool KeepCRLF)
	{
		int Index = 0;
		byte[] bytes = null;

		// the fastest way to process this string
		// is copy it into an array of Bytes
		bytes = System.Text.UnicodeEncoding.Unicode.GetBytes(Source);
		for (Index = 0; Index <= Information.UBound(bytes); Index += 2) {
			// if this is a control character
			if (bytes[Index] < 32 & bytes[Index + 1] == 0) {
				if (!KeepCRLF | (bytes[Index] != 13 & bytes[Index] != 10)) {
					// the user asked to trim CRLF or this
					// character isn't a CR or a LF, so clear it
					bytes[Index] = 0;
				}
			}
		}

		// return this string, after filtering out all null chars
		return Strings.Replace(System.Text.UnicodeEncoding.Unicode.GetString(bytes), Constants.vbNullChar, "");
	}

	//***************************************************************************
	// Open SMART to allow DeviceIoControl communications. Return SMART handle
	//***************************************************************************
	private static int OpenSmart(ref IDE_DRIVE_NUMBER drv_num)
	{
		int functionReturnValue = 0;
		if (IsWindowsNT()) {
			functionReturnValue = CreateFile("\\\\.\\PhysicalDrive" + (string)drv_num, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, 0, OPEN_EXISTING, 0, 0);
		}
		else {
			functionReturnValue = CreateFile("\\\\.\\SMARTVSD", 0, 0, 0, CREATE_NEW, 0, 0);
		}
		return functionReturnValue;
	}


	//****************************************************************************
	// ReadAttributesCmd
	// FUNCTION: Send a SMART_READ_ATTRIBUTE_VALUES command to the drive
	// bDriveNum = 0-3
	//***************************************************************************}
	private static bool ReadAttributesCmd(int hDrive, ref IDE_DRIVE_NUMBER DriveNum)
	{
		bool functionReturnValue = false;
		object READ_ATTRIBUTE_BUFFER_SIZE = null;
		int cbBytesReturned = 0;
		SENDCMDINPARAMS SCIP = null;
		DRIVEATTRIBUTE drv_attr = default(DRIVEATTRIBUTE);
		byte[] bArrOut = new byte[OUTPUT_DATA_SIZE];
		string sMsg = null;
		int i = 0;
		{
			// Set up data structures for Read Attributes SMART Command.
			SCIP.cBufferSize = READ_ATTRIBUTE_BUFFER_SIZE;
			SCIP.bDriveNumber = DriveNum;
			{
				SCIP.irDriveRegs.bFeaturesReg = SMART_READ_ATTRIBUTE_VALUES;
				SCIP.irDriveRegs.bSectorCountReg = 1;
				SCIP.irDriveRegs.bSectorNumberReg = 1;
				SCIP.irDriveRegs.bCylLowReg = SMART_CYL_LOW;
				SCIP.irDriveRegs.bCylHighReg = SMART_CYL_HI;
				//  Compute the drive number.
				SCIP.irDriveRegs.bDriveHeadReg = 0xa0;
				if (!IsWindowsNT()) SCIP.irDriveRegs.bDriveHeadReg = SCIP.irDriveRegs.bDriveHeadReg | (short)DriveNum & 1 * 16; 
				SCIP.irDriveRegs.bCommandReg = IDE_EXECUTE_SMART_FUNCTION;
			}
		}
		functionReturnValue = DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, ref SCIP, Strings.Len(SCIP) - 4, ref bArrOut[0], OUTPUT_DATA_SIZE, ref cbBytesReturned, 0);
		 // ERROR: Not supported in C#: OnErrorStatement

		for (i = 0; i <= NUM_ATTRIBUTE_STRUCTS - 1; i++) {
			if (bArrOut[18 + i * 12] > 0) {
				di.Attributes(di.NumAttributes).AttrID = bArrOut[18 + i * 12];
				di.Attributes(di.NumAttributes).AttrName = "Unknown value (" + bArrOut[18 + i * 12] + ")";
				di.Attributes(di.NumAttributes).AttrName = colAttrNames[(string)bArrOut[18 + i * 12]];
				di.NumAttributes = di.NumAttributes + 1;
				Array.Resize(ref di.Attributes, di.NumAttributes + 1);
				CopyMemory(ref di.Attributes(di.NumAttributes).StatusFlags, bArrOut[19 + i * 12], 2);
				di.Attributes(di.NumAttributes).AttrValue = bArrOut[21 + i * 12];
				di.Attributes(di.NumAttributes).WorstValue = bArrOut[22 + i * 12];
			}
		}
		return functionReturnValue;
	}

	private static bool IsWindowsNT()
	{
		//Dim verinfo As OSVERSIONINFO
		//verinfo.dwOSVersionInfoSize = Len(verinfo)
		//If (GetVersionEx(verinfo)) = 0 Then Exit Function
		//If verinfo.dwPlatformId = 2 Then IsWindowsNT = True
		CWindows.OperatingSystemVersion MyHost = new CWindows.OperatingSystemVersion();
		return MyHost.IsWinNT4Plus();
	}

	private static bool IsBitSet(ref byte iBitString, short lBitNo)
	{
		bool functionReturnValue = false;
		if (lBitNo == 7) {
			functionReturnValue = iBitString < 0;
		}
		else {
			functionReturnValue = iBitString & (Math.Pow(2, lBitNo));
		}
		return functionReturnValue;
	}

	private static string SwapStringBytes(string sIn)
	{
		string sTemp = null;
		short i = 0;
		sTemp = Strings.Space(Strings.Len(sIn));
		for (i = 1; i <= Strings.Len(sIn) - 1; i += 2) {
			Strings.Mid(sTemp, i, 1) = Strings.Mid(sIn, i + 1, 1);
			Strings.Mid(sTemp, i + 1, 1) = Strings.Mid(sIn, i, 1);
		}
		return sTemp;
	}

	public static void FillAttrNameCollection()
	{
		colAttrNames = new Collection();
		{
			colAttrNames.Add("ATTR_INVALID", "0");
			colAttrNames.Add("READ_ERROR_RATE", "1");
			colAttrNames.Add("THROUGHPUT_PERF", "2");
			colAttrNames.Add("SPIN_UP_TIME", "3");
			colAttrNames.Add("START_STOP_COUNT", "4");
			colAttrNames.Add("REALLOC_SECTOR_COUNT", "5");
			colAttrNames.Add("READ_CHANNEL_MARGIN", "6");
			colAttrNames.Add("SEEK_ERROR_RATE", "7");
			colAttrNames.Add("SEEK_TIME_PERF", "8");
			colAttrNames.Add("POWER_ON_HRS_COUNT", "9");
			colAttrNames.Add("SPIN_RETRY_COUNT", "10");
			colAttrNames.Add("CALIBRATION_RETRY_COUNT", "11");
			colAttrNames.Add("POWER_CYCLE_COUNT", "12");
			colAttrNames.Add("SOFT_READ_ERROR_RATE", "13");
			colAttrNames.Add("G_SENSE_ERROR_RATE", "191");
			colAttrNames.Add("POWER_OFF_RETRACT_CYCLE", "192");
			colAttrNames.Add("LOAD_UNLOAD_CYCLE_COUNT", "193");
			colAttrNames.Add("TEMPERATURE", "194");
			colAttrNames.Add("REALLOCATION_EVENTS_COUNT", "196");
			colAttrNames.Add("CURRENT_PENDING_SECTOR_COUNT", "197");
			colAttrNames.Add("UNCORRECTABLE_SECTOR_COUNT", "198");
			colAttrNames.Add("ULTRADMA_CRC_ERROR_RATE", "199");
			colAttrNames.Add("WRITE_ERROR_RATE", "200");
			colAttrNames.Add("DISK_SHIFT", "220");
			colAttrNames.Add("G_SENSE_ERROR_RATEII", "221");
			colAttrNames.Add("LOADED_HOURS", "222");
			colAttrNames.Add("LOAD_UNLOAD_RETRY_COUNT", "223");
			colAttrNames.Add("LOAD_FRICTION", "224");
			colAttrNames.Add("LOAD_UNLOAD_CYCLE_COUNTII", "225");
			colAttrNames.Add("LOAD_IN_TIME", "226");
			colAttrNames.Add("TORQUE_AMPLIFICATION_COUNT", "227");
			colAttrNames.Add("POWER_OFF_RETRACT_COUNT", "228");
			colAttrNames.Add("GMR_HEAD_AMPLITUDE", "230");
			colAttrNames.Add("TEMPERATUREII", "231");
			colAttrNames.Add("READ_ERROR_RETRY_RATE", "250");
		}
	}


	//===============================================================================
	// Name: Function SmartGetVersion
	// Input:
	//   ByVal hDrive As Long - SMART drive handle
	// Output:
	//   Boolean - True if successful
	// Purpose: Given the SMART drive handle, gets the version
	// Remarks: None
	//===============================================================================

	private static bool SmartGetVersion(int hDrive)
	{
		int cbBytesReturned = 0;
		//Arrays in structure GVOP may need to be initialized before they can be used
		GETVERSIONOUTPARAMS GVOP = null;

		return DeviceIoControl(hDrive, DFP_GET_VERSION, ref 0, 0, ref GVOP, Strings.Len(GVOP), ref cbBytesReturned, 0);

	}
	//===============================================================================
	// Name: Function SwapBytes
	// Input:
	//   ByRef b As Byte - Input byte array
	// Output:
	//   Byte - Swapped byte array
	// Purpose: Swaps byte arrays
	// Remarks: None
	//===============================================================================

	private static byte[] SwapBytes(ref byte[] b)
	{
		//Note: VB4-32 and VB5 do not support the
		//return of arrays from a function. For
		//developers using these VB versions there
		//are two workarounds to this restriction:
		//
		//1) Change the return data type ( As Byte() )
		//   to As Variant (no brackets). No change
		//   to the calling code is required.
		//
		//2) Change the function to a sub, remove
		//   the last line of code (SwapBytes = b()),
		//   and take advantage of the fact the
		//   original byte array is being passed
		//   to the function ByRef, therefore any
		//   changes made to the passed data are
		//   actually being made to the original data.
		//   With this workaround the calling code
		//   also requires modification:
		//
		//      di.Model = StrConv(SwapBytes(IDSEC.sModelNumber), vbUnicode)
		//
		//   ... to ...
		//
		//      Call SwapBytes(IDSEC.sModelNumber)
		//      di.Model = StrConv(IDSEC.sModelNumber, vbUnicode)

		byte bTemp = 0;
		int cnt = 0;

		for (cnt = Information.LBound(b); cnt <= Information.UBound(b); cnt += 2) {
			bTemp = b[cnt];
			b[cnt] = b[cnt + 1];
			b[cnt + 1] = bTemp;
		}

		return b;
		// VB6.CopyArray(b)

	}
	//===============================================================================
	// Name: Function GetMACAddress
	// Input: None
	// Output:
	//   String - MAC address of the computer NIC
	// Purpose: Retrieves the MAC Address for the network controller installed, returning a formatted string
	// Remarks: None
	//===============================================================================
	public static string GetMACAddress()
	{
		string functionReturnValue = null;
		functionReturnValue = string.Empty;

		// ******* METHOD 1 *******
		// This was causing problems and therefore was commented out

		//On Error Resume Next
		//Dim tmp As String
		//Dim pASTAT As Integer
		//Dim NCB As NET_CONTROL_BLOCK
		//Dim AST As ASTAT

		//'The IBM NetBIOS 3.0 specifications defines four basic
		//'NetBIOS environments under the NCBRESET command. Win32
		//'follows the OS/2 Dynamic Link Routine (DLR) environment.
		//'This means that the first NCB issued by an application
		//'must be a NCBRESET, with the exception of NCBENUM.
		//'The Windows NT implementation differs from the IBM
		//'NetBIOS 3.0 specifications in the NCB_CALLNAME field.
		//NCB.ncb_command = NCBRESET
		//Call Netbios(NCB)

		//'To get the Media Access Control (MAC) address for an
		//'ethernet adapter programmatically, use the Netbios()
		//'NCBASTAT command and provide a "*" as the name in the
		//'NCB.ncb_CallName field (in a 16-chr string).
		//NCB.ncb_callname = "*               "
		//NCB.ncb_command = NCBASTAT

		//'For machines with multiple network adapters you need to
		//'enumerate the LANA numbers and perform the NCBASTAT
		//'command on each. Even when you have a single network
		//'adapter, it is a good idea to enumerate valid LANA numbers
		//'first and perform the NCBASTAT on one of the valid LANA
		//'numbers. It is considered bad programming to hardcode the
		//'LANA number to 0 (see the comments section below).
		//NCB.ncb_lana_num = 0
		//NCB.ncb_length = Len(AST)

		//pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or HEAP_ZERO_MEMORY, NCB.ncb_length)

		//If pASTAT = 0 Then
		//    System.Diagnostics.Debug.WriteLine("memory allocation failed!")
		//    Exit Function
		//End If

		//NCB.ncb_buffer = pASTAT
		//Call Netbios(NCB)

		//CopyMemory(AST, NCB.ncb_buffer, Len(AST))

		//tmp = Right("00" & Hex(AST.adapt.adapter_address(0)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(1)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(2)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(3)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(4)), 2) & " " & Right("00" & Hex(AST.adapt.adapter_address(5)), 2)

		//HeapFree(GetProcessHeap(), 0, pASTAT)

		//'GetMACAddress = Replace(tmp, " ", "")
		//GetMACAddress = tmp

		// ******* METHOD 2 *******
		clsNetworkStats netInfo = new clsNetworkStats();
		clsNetworkStats.IFROW_HELPER netStruct = default(clsNetworkStats.IFROW_HELPER);
		netStruct = netInfo.GetAdapter();
		functionReturnValue = netStruct.PhysAddr.ToString();
		if (functionReturnValue != "00-00-00-00-00-00") return;
 

		// ******* METHOD 3 *******
		// Here we are assuming that the user is NOT running .NET in a Win98/Me machine...
		ManagementClass mc = null;
		ManagementObject mo = null;
		mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
		ManagementObjectCollection moc = mc.GetInstances();
		foreach (ManagementObject mo_loopVariable in moc) {
			mo = mo_loopVariable;
			if (mo["IPEnabled"].ToString() == "True") {
				functionReturnValue = mo["MacAddress"].ToString().Replace(":", "-");
				if (functionReturnValue != "00-00-00-00-00-00") return;
 
			}
		}

		// ******* METHOD 4 *******
		ManagementObjectSearcher objMOS = null;
		Management.ManagementObjectCollection objMOC = null;
		Management.ManagementObject objMO = null;
		objMOS = new ManagementObjectSearcher("Select * From Win32_NetworkAdapter");
		objMOC = objMOS.Get();
		foreach (ManagementObject objMO_loopVariable in objMOC) {
			objMO = objMO_loopVariable;
			functionReturnValue = objMO["MACAddress"].ToString();
			if (functionReturnValue != "00-00-00-00-00-00") return;
 
		}
		objMOS.Dispose();
		objMOS = null;
		objMO = null;

		// ******* METHOD 5 *******
		System.Net.NetworkInformation.NetworkInterface nic = null;
		foreach (System.Net.NetworkInformation.NetworkInterface nic_loopVariable in System.Net.NetworkInformation.NetworkInterface.GetAllNetworkInterfaces()) {
			nic = nic_loopVariable;
			functionReturnValue = nic.GetPhysicalAddress().ToString();
			if (!string.IsNullOrEmpty(functionReturnValue) & functionReturnValue != "00-00-00-00-00-00") {
				nic = null;
				return;
			}
		}

		// ******* METHOD 6 *******
		//another user provided the code below that seems to work well
		//if the adapter card is "Ethernet 802.3", then the code below will work
		object objset = null;
		object obj = null;
		objset = Interaction.GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_NetworkAdapter");
		foreach (object obj_loopVariable in objset) {
			obj = obj_loopVariable;
			 // ERROR: Not supported in C#: OnErrorStatement

			if (!Information.IsDBNull(obj.MACAddress)) {
				if (obj.AdapterType == "Ethernet 802.3") {
					if (Strings.InStr(obj.PNPDeviceID, "PCI\\") != 0) {
						functionReturnValue = modTrial.Replace_Renamed(obj.MACAddress, ":", "-");
						return;
					}
				}
			}
		}
		GetMACAddressError:

		if (string.IsNullOrEmpty(functionReturnValue)) {
			functionReturnValue = "Not Available";
		}
		return functionReturnValue;

	}
	public static bool WirelessIsFoundAndConnected()
	{
		bool functionReturnValue = false;
		NetworkInformation.NetworkInterface[] adapters = NetworkInformation.NetworkInterface.GetAllNetworkInterfaces();
		bool wirelessfound = false;
		foreach (NetworkInformation.NetworkInterface adapter in adapters) {
			if (adapter.NetworkInterfaceType == NetworkInformation.NetworkInterfaceType.Wireless80211) {
				//If adapter.Description.ToLower.Contains("wireless") Then  ' this also works
				if (adapter.GetIPProperties().UnicastAddresses.Count > 0) {
					functionReturnValue = true;
				}
				//End If

				//Debug.WriteLine("Name " & adapter.Name)
				//Debug.WriteLine("Status:" & adapter.OperationalStatus.ToString)
				//Debug.WriteLine("Speed:" & adapter.Speed.ToString())
				//Debug.WriteLine(adapter.GetIPProperties.GetIPv4Properties)
				//MessageBox.Show(adapter.GetIPProperties.GetIPv4Properties.IsDhcpEnabled.ToString)
				//If adapter.GetIPProperties.GetIPv4Properties.IsDhcpEnabled Then
				//    Debug.WriteLine("Dynamic IP")
				//Else
				//    Debug.WriteLine("Static IP")
				//End If
				//WirelessIsFoundAndConnected = True

			}
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function GetWindowsSerial
	// Input: None
	// Output:
	//   String - Windows serial number
	// Purpose: Gets the Windows Serial Number
	// Remarks: .NET way of doing things added
	//===============================================================================
	public static string GetWindowsSerial()
	{
		string functionReturnValue = null;
		RegistryKey myReg = Registry.LocalMachine;
		RegistryKey MyRegKey = null;
		MyRegKey = myReg.OpenSubKey("Software\\Microsoft\\Windows NT\\CurrentVersion");
		functionReturnValue = MyRegKey.GetValue("ProductID");
		MyRegKey.Close();
		return functionReturnValue;
	}

	//===============================================================================
	// Name: Function GetBiosVersion
	// Input: None
	// Output:
	//   String - BIOS serial number
	// Purpose: Gets the BIOS Serial Number
	// Remarks: Uses the WMI
	//===============================================================================
	public static string GetBiosVersion()
	{
		string functionReturnValue = null;
		object BiosSet = null;
		object obj = null;

		functionReturnValue = string.Empty;

		 // ERROR: Not supported in C#: OnErrorStatement

		BiosSet = Interaction.GetObject("WinMgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_BIOS");
		foreach (object obj_loopVariable in BiosSet) {
			obj = obj_loopVariable;
			functionReturnValue = obj.Version;
			functionReturnValue = functionReturnValue.Replace(" ", "");
			if (!string.IsNullOrEmpty(functionReturnValue)) return;
 
		}
		GetBiosVersionerror:
		if (string.IsNullOrEmpty(functionReturnValue)) {
			functionReturnValue = "Not Available";
		}
		return functionReturnValue;
	}

	//===============================================================================
	// Name: Function GetMotherboardSerial
	// Input: None
	// Output:
	//   String - Motherboard serial number
	// Purpose: Gets the Motherboard Serial Number
	// Remarks: Uses the WMI
	//===============================================================================
	public static string GetMotherboardSerial()
	{
		string functionReturnValue = null;
		object MotherboardSet = null;
		object obj = null;

		functionReturnValue = string.Empty;
		 // ERROR: Not supported in C#: OnErrorStatement


		MotherboardSet = Interaction.GetObject("WinMgmts:{impersonationLevel=impersonate}").InstancesOf("CIM_Chassis");
		byte[] bytes = null;
		foreach (object obj_loopVariable in MotherboardSet) {
			obj = obj_loopVariable;
			functionReturnValue = obj.SerialNumber;
			if (!string.IsNullOrEmpty(functionReturnValue)) {
				// Strip any periods
				bytes = System.Text.UnicodeEncoding.Unicode.GetBytes(GetMotherboardSerial());
				functionReturnValue = Strings.Replace(System.Text.UnicodeEncoding.Unicode.GetString(bytes), ".", "");
				return;
			}
		}
		GetMotherboardSeriAlerror:

		if (string.IsNullOrEmpty(functionReturnValue)) {
			functionReturnValue = "Not Available";
		}
		return functionReturnValue;
	}
	//===============================================================================
	// Name: Function GetIPaddress
	// Input: None
	// Output:
	//   String - IP address
	// Purpose: Gets the IP address
	// Remarks:
	//===============================================================================
	public static string GetIPaddress()
	{
		string functionReturnValue = null;
		 // ERROR: Not supported in C#: OnErrorStatement

		functionReturnValue = string.Empty;

		if (modActiveLock.IsWebConnected() == false) {
			functionReturnValue = "-1";
			return;
		}

		// This is the old method 
		// It worked but not necessarily all the time
		// There could be many IP addresses and one has to check if they are empty
		//Dim ipEntry As IPHostEntry = Dns.GetHostEntry(Environment.MachineName)
		//Dim IpAddr As IPAddress() = ipEntry.AddressList
		//GetIPaddress = IpAddr(0).ToString()

		//A hostmachine can have more than one IP assigned 
		IPAddress[] NIC_IPs = Dns.GetHostEntry(Dns.GetHostName()).AddressList;
		foreach (IPAddress IPAdr in NIC_IPs) {
			functionReturnValue = IPAdr.ToString();
			if (functionReturnValue != "0.0.0.0" & functionReturnValue != "127.0.0.1" & functionReturnValue.Contains(":") == false) return;
 
		}
		GetIPaddressError:

		if (string.IsNullOrEmpty(functionReturnValue)) {
			functionReturnValue = "Not Available";
		}
		return functionReturnValue;
	}

	private static object GetHDSerialFirmwareVBNET(int controller, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute(true)]  // ERROR: Optional parameters aren't supported in C#
bool masterDrive)
	{
		object functionReturnValue = null;
		// Created with the help of the following articles and clues from ALCrypto3.dll
		// SOURCE 1: http://discuss.develop.com/archives/wa.exe?A2=ind0309a&L=advanced-dotnet&D=0&T=0&P=3760
		// SOURCE 2: http://www.visual-basic.it/scarica.asp?ID=611
		// SOURCE 3: ALCrypto3.dll and DISKID32 program

		// This code DOES NOT require admin rights in the user's machine
		// This code DOES NOT require WMI
		// This code DOES NOT require SMART VXD drivers for Win95/98/Me

		string myStr = string.Empty;
		string str1 = null;
		string reversedStr = null;
		string str2 = null;
		short jj = 0;
		int dummy = 0;
		IntPtr hdh = default(IntPtr);
		bool newHandle = false;

		functionReturnValue = "";
		hdh = CreateFile2("\\\\.\\Scsi" + controller.ToString() + ":", GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);

		byte[] bin = new byte[560];
		IntPtr bout = System.Runtime.InteropServices.Marshal.AllocCoTaskMem(560);
		if ((hdh.ToInt32() != -1)) {
			bin[0] = 28;
			bin[4] = 83;
			bin[5] = 67;
			bin[6] = 83;
			bin[7] = 73;
			bin[8] = 68;
			bin[9] = 73;
			bin[10] = 83;
			bin[11] = 75;
			bin[12] = 16;
			bin[13] = 39;
			bin[16] = 1;
			bin[17] = 5;
			bin[18] = 27;
			bin[24] = 20;
			//17?
			bin[25] = 2;
			bin[38] = 236;
			//&HEC

			if (masterDrive == true) {
				bin[40] = 0;
				//master drive
			}
			else {
				bin[40] = 1;
				//slave drive
			}

			System.Runtime.InteropServices.Marshal.Copy(bin, 0, bout, 560);
			newHandle = DeviceIoControl2(hdh, 315400, bout, 63, bout, 560, ref dummy, IntPtr.Zero);

			if (newHandle) {
				System.Runtime.InteropServices.Marshal.Copy(bout, bin, 0, 560);
				// HDD Firmware Serial Number is between 64 to 83 - 19 digits as we had from ALCrypto before
				// HDD Model Number is between 98 and 137
				// HDD Controller Revision is between 90 and 97
				for (jj = 64; jj <= 83; jj++) {
					myStr = myStr + Convert.ToString(Convert.ToChar(bin[jj]));
				}

				// Seems like some swapping is needed at this point
				reversedStr = "";
				for (jj = 0; jj <= Strings.Len(myStr) / 2; jj++) {
					str1 = Strings.Mid(myStr, jj * 2 + 1, 1);
					str2 = Strings.Mid(myStr, jj * 2 + 2, 1);
					reversedStr = reversedStr + str2 + str1;
				}
				functionReturnValue = StripControlChars(ref Strings.Trim(reversedStr), ref false);
			}

		}

		System.Runtime.InteropServices.Marshal.FreeHGlobal(bout);
		CloseHandle2(hdh);
		return functionReturnValue;

	}
	public static string GetHDSerialFirmwareWMI()
	{
		string functionReturnValue = null;
		functionReturnValue = "";

		ManagementScope managementScope = new ManagementScope("\\root\\cimv2");
		managementScope.Options.Impersonation = System.Management.ImpersonationLevel.Impersonate;

		ManagementObjectSearcher searcher = new ManagementObjectSearcher(managementScope, new ObjectQuery("SELECT * FROM Win32_DiskDrive WHERE InterfaceType=\"IDE\" or InterfaceType=\"SCSI\""));
		foreach (ManagementObject disk in searcher.Get()) {
			if (disk["PNPDeviceID"] != null) {
				string pnpDeviceID = disk["PNPDeviceID"].ToString();

				string[] split = pnpDeviceID.Split(new string[] { "\\" }, StringSplitOptions.None);
				if (split.Length == 3) {
					if (!Strings.Split(2).Contains("&")) {
						if (Strings.Split(2).Contains("_")) Strings.Split(2) = Strings.Split(2).Substring(0, Strings.Split(2).IndexOf("_")); 
						byte[] bytes = GetHexStringBytes(Strings.Split(2));
						if (bytes.Length > 0) {
							functionReturnValue = ReverseSerialNumber(System.Text.Encoding.UTF8.GetString(bytes)).Trim();
						}
					}
					else {
						// Custom checks go into here
						string[] parts = null;
						parts = pnpDeviceID.Split("\\".ToCharArray());
						// The serial number should be the next to the last element
						functionReturnValue = parts[parts.Length - 1];
						functionReturnValue = Strings.Replace(GetHDSerialFirmwareWMI(), "&", "");
					}
				}
			}
		}
		return functionReturnValue;
	}

	private static string ReverseSerialNumber(string serialNumber)
	{
		serialNumber = serialNumber.Trim();
		StringBuilder sb = new StringBuilder();
		for (int i = 0; i <= serialNumber.Length - 1; i += 2) {
			sb.Append(serialNumber[i + 1].ToString() + serialNumber[i].ToString());
		}
		serialNumber = sb.ToString();
		sb = null;
		return serialNumber;
	}

	private static byte[] GetHexStringBytes(string hex)
	{
		try {
			if (hex.Contains(String.Empty)) {
				hex = hex.Replace(" ", String.Empty);
			}
			if (hex.Length % 2 == 1) {
				hex = "0" + hex;
			}
			int size = (int)(double)hex.Length / (double)2;
			byte[] bytes = new byte[size];
			for (int i = 0; i <= size - 1; i++) {
				bytes[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
			}
			return bytes;
		}
		catch {
			return new byte[];
		}
	}

	public static string GetCPUID()
	{
		string functionReturnValue = null;
		functionReturnValue = string.Empty;

		ManagementScope managementScope = new ManagementScope("\\root\\cimv2");
		managementScope.Options.Impersonation = System.Management.ImpersonationLevel.Impersonate;

		ManagementObjectSearcher searcher = new ManagementObjectSearcher(managementScope, new ObjectQuery("SELECT * FROM Win32_Processor"));
		foreach (ManagementObject disk in searcher.Get()) {
			if (disk["ProcessorID"] != null) {
				functionReturnValue = disk["ProcessorID"].ToString();
			}
			else {
				functionReturnValue = "Not Available";
			}
		}
		return functionReturnValue;
	}
	public static string GetBaseBoardID()
	{
		string functionReturnValue = null;
		functionReturnValue = string.Empty;

		ManagementScope managementScope = new ManagementScope("\\root\\cimv2");
		managementScope.Options.Impersonation = System.Management.ImpersonationLevel.Impersonate;

		ManagementObjectSearcher searcher = new ManagementObjectSearcher(managementScope, new ObjectQuery("SELECT * FROM Win32_Baseboard"));
		foreach (ManagementObject disk in searcher.Get()) {
			if (disk["SerialNumber"] != null) {
				functionReturnValue = disk["SerialNumber"].ToString();
				functionReturnValue = functionReturnValue.Replace(".", "");
			}
			else {
				functionReturnValue = "Not Available";
			}
		}
		return functionReturnValue;
	}

	public static string GetVideoID()
	{
		string functionReturnValue = null;
		functionReturnValue = string.Empty;

		ManagementScope managementScope = new ManagementScope("\\root\\cimv2");
		managementScope.Options.Impersonation = System.Management.ImpersonationLevel.Impersonate;

		ManagementObjectSearcher searcher = new ManagementObjectSearcher(managementScope, new ObjectQuery("SELECT * FROM Win32_VideoController"));
		foreach (ManagementObject disk in searcher.Get()) {
			if (disk["DriverVersion"] != null) {
				functionReturnValue = disk["DriverVersion"].ToString();
			}
			else {
				functionReturnValue = "Not Available";
			}
		}
		return functionReturnValue;
	}

	public static string GetMemoryID()
	{
		string functionReturnValue = null;
		functionReturnValue = string.Empty;

		ManagementScope managementScope = new ManagementScope("\\root\\cimv2");
		managementScope.Options.Impersonation = System.Management.ImpersonationLevel.Impersonate;

		ManagementObjectSearcher searcher = new ManagementObjectSearcher(managementScope, new ObjectQuery("SELECT * FROM Win32_MemoryDevice"));
		foreach (ManagementObject disk in searcher.Get()) {
			if (disk["SystemName"] != null) {
				string MemoryID = disk["EndingAddress"].ToString() + "-" + ActiveLock3_6NET.My.MyProject.Computer.Info.TotalPhysicalMemory.ToString();
				return MemoryID;
			}
			else {
				functionReturnValue = "Not Available";
			}
		}
		return functionReturnValue;
	}

	private static DirectoryEntry deDomainRoot;
	private static string strDomainPath;
	[DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
	private static extern int ConvertSidToStringSid(IntPtr bSID, 	[System.Runtime.InteropServices.In(), System.Runtime.InteropServices.Out(), System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.LPTStr)]
ref string SIDString);

	private static bool ConnectToAD()
	{
		try {
			deDomainRoot = new DirectoryEntry("LDAP://rootDSE");
			strDomainPath = "LDAP://" + deDomainRoot.Properties("DefaultNamingContext")(0).ToString();
			deDomainRoot = new DirectoryEntry(strDomainPath);
			return true;
		}
		catch (Exception ex) {
			return false;
		}
	}
	private static ManagementObjectSearcher query;

	private static ManagementObjectCollection queryCollection;
	public static string GetSID()
	{
		string functionReturnValue = null;
		functionReturnValue = string.Empty;

		ConnectionOptions co = new ConnectionOptions();
		co.Username = System.Environment.UserDomainName;
		ManagementScope msc = new ManagementScope("\\root\\cimv2", co);
		string queryString = "SELECT * FROM Win32_UserAccount where name='" + co.Username + "'";
		SelectQuery q = new SelectQuery(queryString);
		query = new ManagementObjectSearcher(msc, q);
		queryCollection = query.Get();
		string res = String.Empty;
		foreach (ManagementObject mo in queryCollection) {
			// there should be only one here! 
			res += mo["SID"].ToString();
		}
		functionReturnValue = res;

		//If (ConnectToAD()) Then
		//    Dim dirSearcher As New DirectorySearcher
		//    Dim singleQueryResult As SearchResult
		//    Dim strSID As String = ""
		//    Dim intSuccess As Integer
		//    Dim userName As String = "administrator"
		//    Try
		//        dirSearcher.SearchScope = SearchScope.Subtree
		//        dirSearcher.SearchRoot = deDomainRoot
		//        dirSearcher.Filter = "(&(sAMAccountName=" & userName & "))"
		//        singleQueryResult = dirSearcher.FindOne()
		//        Dim sidBytes As Byte() = CType(singleQueryResult.Properties("objectSid")(0), Byte())
		//        Dim sidPtr As IntPtr = System.Runtime.InteropServices.Marshal.AllocHGlobal(sidBytes.Length)
		//        System.Runtime.InteropServices.Marshal.Copy(sidBytes, 0, sidPtr, sidBytes.Length)
		//        intSuccess = ConvertSidToStringSid(sidPtr, strSID)
		//        GetSID = strSID.Trim()
		//    Catch ex As Exception
		//        Return "Not Available"
		//    End Try
		//End If
		if (string.IsNullOrEmpty(GetSID())) {
			functionReturnValue = "Not Available";
		}
		return functionReturnValue;

	}
	public static string GetExternalIP()
	{
		string functionReturnValue = null;
		string IP_URL = "http://checkip.dyndns.org";
		string strHTML = null;
		string strIP = null;
		try {
			System.Net.WebRequest objWebReq = System.Net.WebRequest.Create(IP_URL);
			System.Net.WebResponse objWebResp = objWebReq.GetResponse();
			System.IO.Stream strmResp = objWebResp.GetResponseStream();
			System.IO.StreamReader srResp = new System.IO.StreamReader(strmResp, System.Text.Encoding.UTF8);
			strHTML = srResp.ReadToEnd();
			System.Text.RegularExpressions.Regex regexIP = null;
			regexIP = new System.Text.RegularExpressions.Regex("\\b\\d{1,3}.\\d{1,3}.\\d{1,3}.\\d{1,3}\\b");
			strIP = regexIP.Match(strHTML).Value;
			return strIP;
		}
		catch (Exception ex) {
			functionReturnValue = "Not Available";
		}
		return functionReturnValue;
	}
	public static string GetFingerprint()
	{
		fp.UseCpuID = true;
		fp.UseBiosID = true;
		fp.UseBaseID = true;
		fp.UseDiskID = true;
		fp.UseVideoID = true;
		fp.UseMacID = true;
		fp.ReturnLength = 8;
		return fp.Value;
	}
	public static bool CheckMACaddress(string usedMACaddress)
	{
		bool functionReturnValue = false;
		functionReturnValue = false;
		ManagementClass mc = null;
		ManagementObject mo = null;
		string nicMACaddress = null;
		mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
		ManagementObjectCollection moc = mc.GetInstances();
		foreach (ManagementObject mo_loopVariable in moc) {
			mo = mo_loopVariable;
			//If mo.Item("IPEnabled").ToString() = "True" Then
			nicMACaddress = mo["MacAddress"].ToString().Replace(":", "-");
			if (nicMACaddress == usedMACaddress) {
				return true;
			}
			//End If
		}
		return functionReturnValue;

	}
}
