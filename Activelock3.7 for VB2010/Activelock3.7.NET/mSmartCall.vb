Option Explicit On 
Option Strict On
Imports System.Runtime.InteropServices
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

Module mSmartCall
    Private Const IDENTIFY_BUFFER_SIZE As Integer = 512
    Private Const OUTPUT_DATA_SIZE As Integer = IDENTIFY_BUFFER_SIZE + 16

    'IOCTL commands
    Private Const DFP_SEND_DRIVE_COMMAND As Integer = &H7C084
    Private Const DFP_RECEIVE_DRIVE_DATA As Integer = &H7C088

    '---------------------------------------------------------------------
    ' IDE registers
    '---------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure IDEREGS
        Public bFeaturesReg As Byte     ' // Used for specifying SMART "commands".
        Public bSectorCountReg As Byte  ' // IDE sector count register
        Public bSectorNumberReg As Byte    ' // IDE sector number register
        Public bCylLowReg As Byte    ' // IDE low order cylinder value
        Public bCylHighReg As Byte   ' // IDE high order cylinder value
        Public bDriveHeadReg As Byte    ' // IDE drive/head register
        Public bCommandReg As Byte   ' // Actual IDE command.
        Public bReserved As Byte     ' // reserved for future use.  Must be zero.
    End Structure

    '---------------------------------------------------------------------
    ' SENDCMDINPARAMS contains the input parameters for the
    ' Send Command to Drive function.
    '---------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure SENDCMDINPARAMS
        Public cBufferSize As Integer   ' Buffer size in bytes
        Public irDriveRegs As IDEREGS   ' Structure with drive register values.
        Public bDriveNumber As Byte     ' Physical drive number to send command to (0,1,2,3).
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=3)> _
        Public bReserved() As Byte     ' Bytes reserved
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=4)> _
        Public dwReserved() As Integer    ' DWORDS reserved
        Public bBuffer As IntPtr     ' Input buffer.
    End Structure

    ' Valid values for the bCommandReg member of IDEREGS.
    Private Const IDE_ID_FUNCTION As Byte = &HEC  ' Returns ID sector for ATA.
    Private Const IDE_EXECUTE_SMART_FUNCTION As Byte = &HB0   ' Performs SMART cmd. Requires valid bFeaturesReg, bCylLowReg, and bCylHighReg

    ' Cylinder register values required when issuing SMART command
    Private Const SMART_CYL_LOW As Byte = &H4F
    Private Const SMART_CYL_HI As Byte = &HC2

    '---------------------------------------------------------------------
    ' Status returned from driver
    '---------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure DRIVERSTATUS
        Public bDriverError As Byte     ' Error code from driver, or 0 if no error.
        Public bIDEStatus As Byte    ' Contents of IDE Error register. Only valid when bDriverError is SMART_IDE_ERROR.
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)> _
        Public bReserved() As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=2)> _
        Public dwReserved() As Integer
    End Structure

    '---------------------------------------------------------------------
    ' The following struct defines the interesting part of the IDENTIFY
    ' buffer:
    '---------------------------------------------------------------------
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure IDSECTOR
        Public wGenConfig As Short
        Public wNumCyls As Short
        Public wReserved As Short
        Public wNumHeads As Short
        Public wBytesPerTrack As Short
        Public wBytesPerSector As Short
        Public wSectorsPerTrack As Short
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=3)> _
        Public wVendorUnique() As Short
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=20)> _
        Public sSerialNumber() As Byte
    End Structure

    '---------------------------------------------------------------------
    ' Structure returned by SMART IOCTL for several commands
    <StructLayout(LayoutKind.Sequential)> _
    Private Structure SENDCMDOUTPARAMS
        Public cBufferSize As Integer     ' Size of bBuffer in bytes (IDENTIFY_BUFFER_SIZE in our case)
        Public DRIVERSTATUS As DRIVERSTATUS  ' Driver status structure.
        Public bBuffer As IntPtr    ' Buffer of arbitrary length in which to store the data read from the drive.
    End Structure

    ' Vendor specific commands:
    Private Const SMART_ENABLE_SMART_OPERATIONS As Byte = &HD8

    Public Enum IDE_DRIVE_NUMBER As Byte
        PRIMARY_MASTER
        PRIMARY_SLAVE
        SECONDARY_MASTER
        SECONDARY_SLAVE
    End Enum

    <DllImport("kernel32", CharSet:=CharSet.Auto, SetLastError:=True)> _
    Private Function CreateFile(ByVal lpFileName As String, ByVal dwDesiredAccess As Integer, ByVal dwShareMode As Integer, ByVal ByVallpSecurityAttributes As Integer, ByVal dwCreationDisposition As Integer, ByVal dwFlagsAndAttributes As Integer, ByVal hTemplateFile As Integer) As Integer
        '
    End Function

    <DllImport("kernel32", SetLastError:=True)> _
    Private Function CloseHandle(ByVal hObject As Integer) As Integer
        '
    End Function

    <DllImport("kernel32", SetLastError:=True)> _
    Private Function DeviceIoControl(ByVal hDevice As Integer, ByVal dwIoControlCode As Integer, ByRef lpInBuffer As SENDCMDINPARAMS, ByVal nInBufferSize As Integer, ByRef lpOutBuffer As SENDCMDOUTPARAMS, ByVal nOutBufferSize As Integer, ByRef lpBytesReturned As Integer, ByVal ByVallpOverlapped As Integer) As Integer
        '
    End Function

    <DllImport("kernel32", SetLastError:=True)> _
    Private Function DeviceIoControl(ByVal hDevice As Integer, ByVal ByValdwIoControlCode As Integer, ByRef lpInBuffer As SENDCMDINPARAMS, ByVal nInBufferSize As Integer, ByVal lpOutBuffer() As Byte, ByVal nOutBufferSize As Integer, ByRef lpBytesReturned As Integer, ByVal lpOverlapped As Integer) As Integer
        '
    End Function

    Private Const GENERIC_READ As Integer = &H80000000
    Private Const GENERIC_WRITE As Integer = &H40000000

    Private Const FILE_SHARE_READ As Integer = &H1
    Private Const FILE_SHARE_WRITE As Integer = &H2
    Private Const OPEN_EXISTING As Integer = 3
    Private Const FILE_ATTRIBUTE_SYSTEM As Integer = &H4
    Private Const CREATE_NEW As Integer = 1

    Private Const INVALID_HANDLE_VALUE As Integer = -1


    ' ***************************************************************************
    ' Open SMART to allow DeviceIoControl communications. Return SMART handle

    ' ***************************************************************************
    Private Function OpenSmart(ByVal drv_num As IDE_DRIVE_NUMBER) As Integer

        OpenSmart = CreateFile(String.Format("\\.\PhysicalDrive{0}", Convert.ToInt32(drv_num)), GENERIC_READ Or GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0, 0)

    End Function


    ' ****************************************************************************
    ' CheckSMARTEnable - Check if SMART enable
    ' FUNCTION: Send a SMART_ENABLE_SMART_OPERATIONS command to the drive
    ' bDriveNum = 0-3

    ' ***************************************************************************
    Private Function CheckSMARTEnable(ByVal hDrive As Integer, ByVal DriveNum As IDE_DRIVE_NUMBER) As Boolean
        'Set up data structures for Enable SMART Command.
        Dim SCIP As SENDCMDINPARAMS = Nothing
        Dim SCOP As SENDCMDOUTPARAMS = Nothing
        Dim lpcbBytesReturned As Integer
        With SCIP
            .cBufferSize = 0
            With .irDriveRegs
                .bFeaturesReg = SMART_ENABLE_SMART_OPERATIONS
                .bCylLowReg = SMART_CYL_LOW
                .bCylHighReg = SMART_CYL_HI
                'Compute the drive number.
                .bDriveHeadReg = &HA0    ' Or (DriveNum And 1) * 16
                .bCommandReg = IDE_EXECUTE_SMART_FUNCTION
            End With
            .bDriveNumber = DriveNum
        End With
        Return Convert.ToBoolean(DeviceIoControl(hDrive, DFP_SEND_DRIVE_COMMAND, SCIP, Marshal.SizeOf(SCIP) - 4, SCOP, Marshal.SizeOf(SCOP) - 4, lpcbBytesReturned, 0&))
    End Function


    ' ***************************************************************************
    ' DoIdentify
    ' Function: Send an IDENTIFY command to the drive
    ' DriveNum = 0-3
    ' IDCmd = IDE_ID_FUNCTION or IDE_ATAPI_ID
    ' ***************************************************************************
    Private Function IdentifyDrive(ByVal hDrive As Integer, ByVal IDCmd As Byte) As String
        Dim SCIP As SENDCMDINPARAMS = Nothing
        Dim bArrOut(OUTPUT_DATA_SIZE - 1) As Byte
        Dim bSerial(19) As Byte
        Dim lpcbBytesReturned As Integer
        '   Set up data structures for IDENTIFY command.

        ' Compute the drive number.
        ' The command can either be IDE identify or ATAPI identify.
        SCIP.irDriveRegs.bCommandReg = CByte(IDCmd)

        If DeviceIoControl(hDrive, DFP_RECEIVE_DRIVE_DATA, SCIP, Marshal.SizeOf(SCIP) - 4, bArrOut, OUTPUT_DATA_SIZE, lpcbBytesReturned, 0&) <> 0 Then
            System.Buffer.BlockCopy(bArrOut, 36, bSerial, 0, 20)
            bSerial = SwapBytes(bSerial)
            Return Encoding.ASCII.GetString(bSerial)
        End If
        Return Nothing
    End Function

    Private Function SwapBytes(ByVal bIn() As Byte) As Byte()
        Dim bTemp(bIn.GetUpperBound(0)) As Byte
        Dim i As Integer

        For i = 0 To bIn.Length - 1 Step 2
            bTemp(i) = bIn(i + 1)
            bTemp(i + 1) = bIn(i)
        Next i
        SwapBytes = bTemp
    End Function

    ' ***************************************************************************
    ' ReadAttributesCmd
    ' FUNCTION: Send a SMART_READ_ATTRIBUTE_VALUES command to the drive
    ' bDriveNum = 0-3
    ' ***************************************************************************

    Public Function GetDriveInfo(ByVal DriveNum As IDE_DRIVE_NUMBER) As String
        Dim hDrive As Integer

        hDrive = OpenSmart(DriveNum)
        If hDrive = INVALID_HANDLE_VALUE Then Return Nothing

        If CheckSMARTEnable(hDrive, DriveNum) Then
            GetDriveInfo = IdentifyDrive(hDrive, IDE_ID_FUNCTION)
        Else
            GetDriveInfo = Nothing
        End If

        CloseHandle(hDrive)
    End Function

    ' test method to make sure all is working....
    'Sub Main()
    '    Console.Write(GetDriveInfo(IDE_DRIVE_NUMBER.PRIMARY_MASTER))
    '    Console.ReadLine()
    'End Sub

End Module
