using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

using System.Runtime.InteropServices;
using System.Text;
static class mSmartCall
{
	private const int IDENTIFY_BUFFER_SIZE = 512;

	private const int OUTPUT_DATA_SIZE = IDENTIFY_BUFFER_SIZE + 16;
	//IOCTL commands
	private const int DFP_SEND_DRIVE_COMMAND = 0x7c084;

	private const int DFP_RECEIVE_DRIVE_DATA = 0x7c088;
	//---------------------------------------------------------------------
	// IDE registers
	//---------------------------------------------------------------------
	[StructLayout(LayoutKind.Sequential)]
	private struct IDEREGS
	{
			// // Used for specifying SMART "commands".
		public byte bFeaturesReg;
			// // IDE sector count register
		public byte bSectorCountReg;
			// // IDE sector number register
		public byte bSectorNumberReg;
			// // IDE low order cylinder value
		public byte bCylLowReg;
			// // IDE high order cylinder value
		public byte bCylHighReg;
			// // IDE drive/head register
		public byte bDriveHeadReg;
			// // Actual IDE command.
		public byte bCommandReg;
			// // reserved for future use.  Must be zero.
		public byte bReserved;
	}

	//---------------------------------------------------------------------
	// SENDCMDINPARAMS contains the input parameters for the
	// Send Command to Drive function.
	//---------------------------------------------------------------------
	[StructLayout(LayoutKind.Sequential)]
	private struct SENDCMDINPARAMS
	{
			// Buffer size in bytes
		public int cBufferSize;
			// Structure with drive register values.
		public IDEREGS irDriveRegs;
			// Physical drive number to send command to (0,1,2,3).
		public byte bDriveNumber;
		[MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
			// Bytes reserved
		public byte[] bReserved;
		[MarshalAs(UnmanagedType.ByValArray, SizeConst = 4)]
			// DWORDS reserved
		public int[] dwReserved;
			// Input buffer.
		public IntPtr bBuffer;
	}

	// Valid values for the bCommandReg member of IDEREGS.
		// Returns ID sector for ATA.
	private const byte IDE_ID_FUNCTION = 0xec;
		// Performs SMART cmd. Requires valid bFeaturesReg, bCylLowReg, and bCylHighReg
	private const byte IDE_EXECUTE_SMART_FUNCTION = 0xb0;

	// Cylinder register values required when issuing SMART command
	private const byte SMART_CYL_LOW = 0x4f;

	private const byte SMART_CYL_HI = 0xc2;
	//---------------------------------------------------------------------
	// Status returned from driver
	//---------------------------------------------------------------------
	[StructLayout(LayoutKind.Sequential)]
	private struct DRIVERSTATUS
	{
			// Error code from driver, or 0 if no error.
		public byte bDriverError;
			// Contents of IDE Error register. Only valid when bDriverError is SMART_IDE_ERROR.
		public byte bIDEStatus;
		[MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
		public byte[] bReserved;
		[MarshalAs(UnmanagedType.ByValArray, SizeConst = 2)]
		public int[] dwReserved;
	}

	//---------------------------------------------------------------------
	// The following struct defines the interesting part of the IDENTIFY
	// buffer:
	//---------------------------------------------------------------------
	[StructLayout(LayoutKind.Sequential)]
	private struct IDSECTOR
	{
		public short wGenConfig;
		public short wNumCyls;
		public short wReserved;
		public short wNumHeads;
		public short wBytesPerTrack;
		public short wBytesPerSector;
		public short wSectorsPerTrack;
		[MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
		public short[] wVendorUnique;
		[MarshalAs(UnmanagedType.ByValArray, SizeConst = 20)]
		public byte[] sSerialNumber;
	}

	//---------------------------------------------------------------------
	// Structure returned by SMART IOCTL for several commands
	[StructLayout(LayoutKind.Sequential)]
	private struct SENDCMDOUTPARAMS
	{
			// Size of bBuffer in bytes (IDENTIFY_BUFFER_SIZE in our case)
		public int cBufferSize;
			// Driver status structure.
		public DRIVERSTATUS DRIVERSTATUS;
			// Buffer of arbitrary length in which to store the data read from the drive.
		public IntPtr bBuffer;
	}

	// Vendor specific commands:

	private const byte SMART_ENABLE_SMART_OPERATIONS = 0xd8;
	public enum IDE_DRIVE_NUMBER : byte
	{
		PRIMARY_MASTER,
		PRIMARY_SLAVE,
		SECONDARY_MASTER,
		SECONDARY_SLAVE
	}

	[DllImport("kernel32", CharSet = CharSet.Auto, SetLastError = true)]
	private static int CreateFile(string lpFileName, int dwDesiredAccess, int dwShareMode, int ByVallpSecurityAttributes, int dwCreationDisposition, int dwFlagsAndAttributes, int hTemplateFile)
	{
		//
	}

	[DllImport("kernel32", SetLastError = true)]
	private static int CloseHandle(int hObject)
	{
		//
	}

	[DllImport("kernel32", SetLastError = true)]
	private static int DeviceIoControl(int hDevice, int dwIoControlCode, ref SENDCMDINPARAMS lpInBuffer, int nInBufferSize, ref SENDCMDOUTPARAMS lpOutBuffer, int nOutBufferSize, ref int lpBytesReturned, int ByVallpOverlapped)
	{
		//
	}

	[DllImport("kernel32", SetLastError = true)]
	private static int DeviceIoControl(int hDevice, int ByValdwIoControlCode, ref SENDCMDINPARAMS lpInBuffer, int nInBufferSize, byte[] lpOutBuffer, int nOutBufferSize, ref int lpBytesReturned, int lpOverlapped)
	{
		//
	}

	private const int GENERIC_READ = 0x80000000;

	private const int GENERIC_WRITE = 0x40000000;
	private const int FILE_SHARE_READ = 0x1;
	private const int FILE_SHARE_WRITE = 0x2;
	private const int OPEN_EXISTING = 3;
	private const int FILE_ATTRIBUTE_SYSTEM = 0x4;

	private const int CREATE_NEW = 1;

	private const int INVALID_HANDLE_VALUE = -1;

	//***************************************************************************
	// Open SMART to allow DeviceIoControl communications. Return SMART handle

	//***************************************************************************

	private static int OpenSmart(IDE_DRIVE_NUMBER drv_num)
	{
		return CreateFile(string.Format("\\\\.\\PhysicalDrive{0}", Convert.ToInt32(drv_num)), GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ | FILE_SHARE_WRITE, 0L, OPEN_EXISTING, 0, 0);

	}


	//****************************************************************************
	// CheckSMARTEnable - Check if SMART enable
	// FUNCTION: Send a SMART_ENABLE_SMART_OPERATIONS command to the drive
	// bDriveNum = 0-3

	//***************************************************************************
	private static bool CheckSMARTEnable(int hDrive, IDE_DRIVE_NUMBER DriveNum)
	{
		//Set up data structures for Enable SMART Command.
		SENDCMDINPARAMS SCIP = null;
		SENDCMDOUTPARAMS SCOP = null;
		int lpcbBytesReturned = 0;
		{
			SCIP.cBufferSize = 0;
			{
				SCIP.irDriveRegs.bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS;
				SCIP.irDriveRegs.bCylLowReg = SMART_CYL_LOW;
				SCIP.irDriveRegs.bCylHighReg = SMART_CYL_HI;
				//Compute the drive number.
				SCIP.irDriveRegs.bDriveHeadReg = 0xa0;
				// Or (DriveNum And 1) * 16
				SCIP.irDriveRegs.bCommandReg = IDE_EXECUTE_SMART_FUNCTION;
			}
			SCIP.bDriveNumber = DriveNum;
		}
		return Convert.ToBoolean(DeviceIoControl(hDrive, DFP_SEND_DRIVE_COMMAND, ref SCIP, Marshal.SizeOf(SCIP) - 4, ref SCOP, Marshal.SizeOf(SCOP) - 4, ref lpcbBytesReturned, 0L));
	}


	//***************************************************************************
	// DoIdentify
	// Function: Send an IDENTIFY command to the drive
	// DriveNum = 0-3
	// IDCmd = IDE_ID_FUNCTION or IDE_ATAPI_ID
	//***************************************************************************
	private static string IdentifyDrive(int hDrive, byte IDCmd)
	{
		SENDCMDINPARAMS SCIP = null;
		byte[] bArrOut = new byte[OUTPUT_DATA_SIZE];
		byte[] bSerial = new byte[20];
		int lpcbBytesReturned = 0;
		//   Set up data structures for IDENTIFY command.

		// Compute the drive number.
		// The command can either be IDE identify or ATAPI identify.
		SCIP.irDriveRegs.bCommandReg = (byte)IDCmd;

		if (DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, ref SCIP, Marshal.SizeOf(SCIP) - 4, bArrOut, OUTPUT_DATA_SIZE, ref lpcbBytesReturned, 0L) != 0) {
			System.Buffer.BlockCopy(bArrOut, 36, bSerial, 0, 20);
			bSerial = SwapBytes(bSerial);
			return Encoding.ASCII.GetString(bSerial);
		}
		return null;
	}

	private static byte[] SwapBytes(byte[] bIn)
	{
		byte[] bTemp = new byte[bIn.GetUpperBound(0) + 1];
		int i = 0;

		for (i = 0; i <= bIn.Length - 1; i += 2) {
			bTemp[i] = bIn[i + 1];
			bTemp[i + 1] = bIn[i];
		}
		return bTemp;
	}

	//***************************************************************************
	// ReadAttributesCmd
	// FUNCTION: Send a SMART_READ_ATTRIBUTE_VALUES command to the drive
	// bDriveNum = 0-3
	//***************************************************************************

	public static string GetDriveInfo(IDE_DRIVE_NUMBER DriveNum)
	{
		string functionReturnValue = null;
		int hDrive = 0;

		hDrive = OpenSmart(DriveNum);
		if (hDrive == INVALID_HANDLE_VALUE) return null; 

		if (CheckSMARTEnable(hDrive, DriveNum)) {
			functionReturnValue = IdentifyDrive(hDrive, IDE_ID_FUNCTION);
		}
		else {
			functionReturnValue = null;
		}

		CloseHandle(hDrive);
		return functionReturnValue;
	}

	// test method to make sure all is working....
	//Sub Main()
	//    Console.Write(GetDriveInfo(IDE_DRIVE_NUMBER.PRIMARY_MASTER))
	//    Console.ReadLine()
	//End Sub

}
