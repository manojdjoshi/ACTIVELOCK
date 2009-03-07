using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
// TODO: ADSFile.vb - Add comment as to what this does/should do! Not Documented!
/// <summary>
/// Not Documented!
/// </summary>
/// <remarks></remarks>
static class ADSFile
{

	private static string _stream = string.Empty;

	private static string _FileName;
	#region "Win32 Constants"
		//ToDo: Unsigned Integers not supported
	private const System.Int32 GENERIC_WRITE = 0x40000000;
	private const System.Int32 GENERIC_READ = 0x80000000;
	//ToDo: Unsigned Integers not supported
	//
	//ToDo: Error processing original source shown below
	//      private  const    uint     GENERIC_WRITE                 = 0x40000000;
	//      private  const    uint     GENERIC_READ                  = 0x80000000;
	//------------------------------------------------------------------^--- Specified cast is not valid.

		//ToDo: Unsigned Integers not supported
	private const System.Int32 FILE_SHARE_READ = 0x1;
		//ToDo: Unsigned Integers not supported
	private const System.Int32 FILE_SHARE_WRITE = 0x2;
		//ToDo: Unsigned Integers not supported
	private const System.Int32 CREATE_NEW = 1;
		//ToDo: Unsigned Integers not supported
	private const System.Int32 CREATE_ALWAYS = 2;
		//ToDo: Unsigned Integers not supported
	private const System.Int32 OPEN_EXISTING = 3;
		//ToDo: Unsigned Integers not supported
	private const System.Int32 OPEN_ALWAYS = 4;
	#endregion

	#region "Win32 API Defines"
	/// <summary>
	/// Retrieves the size of the specified file, in bytes.
	/// </summary>
	/// <param name="handle">A handle to the file.</param>
	/// <param name="size">A pointer to the variable where the high-order doubleword of the file size is returned. This parameter can be NULL if the application does not require the high-order doubleword.</param>
	/// <returns>If the function succeeds, the return value is the low-order doubleword of the file size, and, if lpFileSizeHigh is non-NULL, the function puts the high-order doubleword of the file size into the variable pointed to by that parameter.</returns>
	/// <remarks>Note that if the return value is INVALID_FILE_SIZE (0xffffffff), an application must call GetLastError to determine whether the function has succeeded or failed. The reason the function may appear to fail when it has not is that lpFileSizeHigh could be non-NULL or the file size could be 0xffffffff. In this case, GetLastError will return NO_ERROR (0) upon success. Because of this behavior, it is recommended that you use GetFileSizeEx instead.</remarks>
	[System.Runtime.InteropServices.DllImport("kernel32", SetLastError = true)]
	public static System.Int32 GetFileSize(System.Int32 handle, IntPtr size)
	{
		//ToDo: Unsigned Integers not supported
		//ToDo: Unsigned Integers not supported
	}

	/// <summary>
	/// <para>Reads data from the specified file or input/output (I/O) device. Reads occur at the position specified by the file pointer if supported by the device.</para>
	/// <para>This function is designed for both synchronous and asynchronous operations. For a similar function designed solely for asynchronous operation, see <a href="http://msdn.microsoft.com/en-us/library/aa365468(VS.85).aspx">ReadFileEx</a>.</para>
	/// </summary>
	/// <param name="handle">A handle to the device (for example, a file, file stream, physical disk, volume, console buffer, tape drive, socket, communications resource, mailslot, or pipe).</param>
	/// <param name="buffer">A pointer to the buffer that receives the data read from a file or device.</param>
	/// <param name="byteToRead">The maximum number of bytes to be read.</param>
	/// <param name="bytesRead">A pointer to the variable that receives the number of bytes read when using a synchronous hFile parameter. ReadFile sets this value to zero before doing any work or error checking. Use NULL for this parameter if this is an asynchronous operation to avoid potentially erroneous results.</param>
	/// <param name="lpOverlapped">A pointer to an OVERLAPPED structure is required if the hFile parameter was opened with FILE_FLAG_OVERLAPPED, otherwise it can be NULL.</param>
	/// <returns>If the function succeeds, the return value is nonzero (TRUE).</returns>
	/// <remarks>see http://msdn.microsoft.com/en-us/library/aa365467(VS.85).aspx </remarks>
	[System.Runtime.InteropServices.DllImport("kernel32", SetLastError = true)]
	public static System.Int32 ReadFile(System.Int32 handle, byte[] buffer, System.Int32 byteToRead, ref System.Int32 bytesRead, IntPtr lpOverlapped)
	{
		//ToDo: Unsigned Integers not supported
		//ToDo: Unsigned Integers not supported
		//ToDo: Unsigned Integers not supported
		//ToDo: Unsigned Integers not supported
	}

	/// <summary>
	/// <para>Creates or opens a file or I/O device. The most commonly used I/O devices are as follows: file, file stream, directory, physical disk, volume, console buffer, tape drive, communications resource, mailslot, and pipe. The function returns a handle that can be used to access the file or device for various types of I/O depending on the file or device and the flags and attributes specified.</para>
	/// <para>To perform this operation as a transacted operation, which results in a handle that can be used for transacted I/O, use the <a href="http://msdn.microsoft.com/en-us/library/aa363859(VS.85).aspx">CreateFileTransacted</a> function.</para>
	/// </summary>
	/// <param name="filename">The name of the file or device to be created or opened. </param>
	/// <param name="desiredAccess">The requested access to the file or device, which can be summarized as read, write, both or neither (zero).</param>
	/// <param name="shareMode">The requested sharing mode of the file or device, which can be read, write, both, delete, all of these, or none (refer to the following table). Access requests to attributes or extended attributes are not affected by this flag.</param>
	/// <param name="attributes">A pointer to a SECURITY_ATTRIBUTES structure that contains two separate but related data members: an optional security descriptor, and a Boolean value that determines whether the returned handle can be inherited by child processes.</param>
	/// <param name="creationDisposition">An action to take on a file or device that exists or does not exist.</param>
	/// <param name="flagsAndAttributes">The file or device attributes and flags, FILE_ATTRIBUTE_NORMAL being the most common default value for files.</param>
	/// <param name="templateFile">A valid handle to a template file with the GENERIC_READ access right. The template file supplies file attributes and extended attributes for the file that is being created.</param>
	/// <returns>
	/// <para>If the function succeeds, the return value is an open handle to the specified file, device, named pipe, or mail slot.</para>
	/// <para>If the function fails, the return value is INVALID_HANDLE_VALUE. To get extended error information, call GetLastError.</para>
	/// </returns>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/aa363858.aspx for full documentation!</remarks>
	[System.Runtime.InteropServices.DllImport("kernel32", SetLastError = true)]
	public static System.Int32 CreateFile(string filename, System.Int32 desiredAccess, System.Int32 shareMode, IntPtr attributes, System.Int32 creationDisposition, System.Int32 flagsAndAttributes, IntPtr templateFile)
	{
		//ToDo: Unsigned Integers not supported
		//ToDo: Unsigned Integers not supported
		//ToDo: Unsigned Integers not supported
		//ToDo: Unsigned Integers not supported
		//ToDo: Unsigned Integers not supported
	}

	/// <summary>
	/// Writes data to the specified file or input/output (I/O) device. Writes occur at the position specified by the file pointer, if the handle refers to a seeking device.
	/// </summary>
	/// <param name="hFile">A handle to the file or I/O device (for example, a file, file stream, physical disk, volume, console buffer, tape drive, socket, communications resource, mailslot, or pipe).</param>
	/// <param name="lpBuffer">A pointer to the buffer containing the data to be written to the file or device.</param>
	/// <param name="nNumberOfBytesToWrite">The number of bytes to be written to the file or device.</param>
	/// <param name="lpNumberOfBytesWritten">A pointer to the variable that receives the number of bytes written when using a synchronous hFile parameter. WriteFile sets this value to zero before doing any work or error checking. Use NULL for this parameter if this is an asynchronous operation to avoid potentially erroneous results.</param>
	/// <param name="lpOverlapped">A pointer to an <a href="http://msdn.microsoft.com/en-us/library/ms684342(VS.85).aspx">OVERLAPPED</a> structure is required if the hFile parameter was opened with FILE_FLAG_OVERLAPPED, otherwise this parameter can be NULL.</param>
	/// <returns>If the function succeeds, the return value is nonzero (TRUE).</returns>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/aa365747(VS.85).aspx for full documentation!</remarks>
	[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
	public static bool WriteFile(System.Int32 hFile, byte[] lpBuffer, System.Int32 nNumberOfBytesToWrite, ref System.Int32 lpNumberOfBytesWritten, IntPtr lpOverlapped)
	{
		//ToDo: Unsigned Integers not supported
		//ToDo: Unsigned Integers not supported
		//ToDo: Unsigned Integers not supported
	}

	/// <summary>
	/// Closes an open object handle.
	/// </summary>
	/// <param name="hFile">A valid handle to an open object.</param>
	/// <returns>If the function succeeds, the return value is nonzero.</returns>
	/// <remarks>See http://msdn.microsoft.com/en-us/library/ms724211(VS.85).aspx for full Documentation!</remarks>
	[System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError = true)]
	//ToDo: Unsigned Integers not supported
	public static bool CloseHandle(System.Int32 hFile)
	{
	}

	#endregion

	#region "ctors"

	/// <summary>
	/// Private constructor.  No instances of the class can be created.
	/// </summary>
	public ADSFile()
	{
	}
	//New
	#endregion

	#region "Public Static Methods"

	/// <summary>
	/// Method called when an alternate data stream must be read from.
	/// </summary>
	/// <param name="file">The fully qualified name of the file from which
	/// the ADS data will be read.</param>
	/// <param name="stream">The name of the stream within the "normal" file
	/// from which to read.</param>
	/// <returns>The contents of the file as a string.  It will always return
	/// at least a zero-length string, even if the file does not exist.
	/// </returns>
	public static string Read(string file, string stream)
	{
		System.Int32 fHandle = CreateFile(file + ":" + stream, GENERIC_READ, FILE_SHARE_READ, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero);
		//ToDo: Unsigned Integers not supported ' Filename
		// Desired access
		// Share more
		// Attributes
		// Creation attributes
		// Flags and attributes
		// Template file
		// if the handle returned is uint.MaxValue, the stream doesn't exist.
		//ToDo: Unsigned Integers not supported
		if (fHandle != System.Int32.MaxValue) {
			// A handle to the stream within the file was created successfully.
			System.Int32 size = GetFileSize(fHandle, IntPtr.Zero);
			//ToDo: Unsigned Integers not supported
			byte[] buffer = new byte[size + 1];
			System.Int32 reader = System.Int32.MinValue;
			//ToDo: Unsigned Integers not supported
			//ToDo: Unsigned Integers not supported

			System.Int32 result = ReadFile(fHandle, buffer, size, ref reader, IntPtr.Zero);
			//ToDo: Unsigned Integers not supported ' Handle
			// Data buffer
			// Bytes to read
			// Bytes actually read
			// Overlapped
			CloseHandle(fHandle);

			// Convert the bytes read into an ASCII string and return it to the caller.
			return System.Text.Encoding.ASCII.GetString(buffer);
		}
		else {
			ReadError(file, stream);
			return null;
		}
	}
	//Read

	/// <summary>
	/// The static method to call when data must be written to a stream.
	/// </summary>
	/// <param name="data">The string data to embed in the stream in the file</param>
	/// <param name="file">The fully qualified name of the file with the
	/// stream into which the data will be written.</param>
	/// <param name="stream">The name of the stream within the normal file to
	/// write the data.</param>
	/// <returns>An unsigned integer of how many bytes were actually written.</returns>
	//ToDo: Unsigned Integers not supported
	public static System.Int32 Write(string data, string file, string stream)
	{
		// Convert the string data to be written to an array of ascii characters.
		byte[] barData = System.Text.Encoding.ASCII.GetBytes(data);
		System.Int32 nReturn = 0;
		//ToDo: Unsigned Integers not supported

		System.Int32 fHandle = CreateFile(file + ":" + stream, GENERIC_WRITE, FILE_SHARE_WRITE, IntPtr.Zero, CREATE_ALWAYS, 0, IntPtr.Zero);
		//ToDo: Unsigned Integers not supported ' File name
		// Desired access
		// Share mode
		// Attributes
		// Creation disposition
		// Flags and attributes
		// Template file
		bool bOK = WriteFile(fHandle, barData, (System.Int32)barData.Length, ref nReturn, IntPtr.Zero);
		//ToDo: Unsigned Integers not supported ' Handle
		// Data buffer
		// Buffer size
		// Bytes written
		// Overlapped
		CloseHandle(fHandle);

		// Throw an exception if the data wasn't written successfully.
		if (!bOK) {
			throw new System.ComponentModel.Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error());
		}
		return nReturn;
	}
	//Write

	/// <summary>
	/// Not Documented!
	/// </summary>
	/// <param name="FileName"></param>
	/// <param name="stream"></param>
	/// <returns></returns>
	/// <remarks></remarks>
	public static string ReadError(string FileName, string stream)
	{
		_stream = stream;
		_FileName = FileName;
		return "Stream \"" + _stream + "\" not found in \"" + _FileName + "\"";
	}

	#endregion
}

//Module StreamNotFoundException
//#Region "Private Members"
//    'Private _stream As String = String.Empty
//    'Private _FileName As String
//#End Region

//#Region "ctors"

//    '/ <summary>
//    '/ Constructor called with the name of the file and stream which was
//    '/ unsuccessfully opened.
//    '/ </summary>
//    '/ <param name="file">Fully qualified name of the file in which the stream
//    '/ was supposed to reside.</param>
//    '/ <param name="stream">Stream within the file to open.</param>
//    'Public Sub New(ByVal file As String, ByVal stream As String)
//    '    MyBase.New(String.Empty, file)
//    '    
//    'End Sub 'New
//    'Public Sub ReadError(ByVal FileName As String, ByVal stream As String)
//    '    _stream = stream
//    '    _FileName = FileName
//    'End Sub


//#End Region

//#Region "Public Properties"
//    '/ <summary>
//    '/ Read-only property to allow the user to query the exception to determine
//    '/ the name of the stream that couldn't be found.
//    '/ </summary>

//    'Public ReadOnly Property Stream() As String
//    '    Get
//    '        Return _stream
//    '    End Get
//    'End Property
//#End Region

//#Region "Overridden Properties"
//    '/ <summary>
//    '/ Overridden Message property to return a concise string describing the
//    '/ exception.
//    '/ </summary>

//    Public ReadOnly Property Message() As String
//        Get
//            Return "Stream """ + _stream + """ not found in """ + _FileName + """"
//        End Get
//    End Property
//#End Region

//End Module
