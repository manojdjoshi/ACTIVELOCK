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
' TODO: ADSFile.vb - Add comment as to what this does/should do! Not Documented!
''' <summary>
''' Not Documented!
''' </summary>
''' <remarks></remarks>
Module ADSFile

    Private _stream As String = String.Empty
    Private _FileName As String

#Region "Win32 Constants"
    Private Const GENERIC_WRITE As System.Int32 = &H40000000 'ToDo: Unsigned Integers not supported
    Private Const GENERIC_READ As System.Int32 = &H80000000
    'ToDo: Unsigned Integers not supported
    '
    'ToDo: Error processing original source shown below
    '      private  const    uint     GENERIC_WRITE                 = 0x40000000;
    '      private  const    uint     GENERIC_READ                  = 0x80000000;
    '------------------------------------------------------------------^--- Specified cast is not valid.

    Private Const FILE_SHARE_READ As System.Int32 = &H1 'ToDo: Unsigned Integers not supported
    Private Const FILE_SHARE_WRITE As System.Int32 = &H2 'ToDo: Unsigned Integers not supported
    Private Const CREATE_NEW As System.Int32 = 1 'ToDo: Unsigned Integers not supported
    Private Const CREATE_ALWAYS As System.Int32 = 2 'ToDo: Unsigned Integers not supported
    Private Const OPEN_EXISTING As System.Int32 = 3 'ToDo: Unsigned Integers not supported
    Private Const OPEN_ALWAYS As System.Int32 = 4 'ToDo: Unsigned Integers not supported
#End Region

#Region "Win32 API Defines"
    ''' <summary>
    ''' Retrieves the size of the specified file, in bytes.
    ''' </summary>
    ''' <param name="handle">A handle to the file.</param>
    ''' <param name="size">A pointer to the variable where the high-order doubleword of the file size is returned. This parameter can be NULL if the application does not require the high-order doubleword.</param>
    ''' <returns>If the function succeeds, the return value is the low-order doubleword of the file size, and, if lpFileSizeHigh is non-NULL, the function puts the high-order doubleword of the file size into the variable pointed to by that parameter.</returns>
    ''' <remarks>Note that if the return value is INVALID_FILE_SIZE (0xffffffff), an application must call GetLastError to determine whether the function has succeeded or failed. The reason the function may appear to fail when it has not is that lpFileSizeHigh could be non-NULL or the file size could be 0xffffffff. In this case, GetLastError will return NO_ERROR (0) upon success. Because of this behavior, it is recommended that you use GetFileSizeEx instead.</remarks>
    <System.Runtime.InteropServices.DllImport("kernel32", SetLastError:=True)> _
     Function GetFileSize(ByVal handle As System.Int32, ByVal size As IntPtr) As System.Int32
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
    End Function

    ''' <summary>
    ''' <para>Reads data from the specified file or input/output (I/O) device. Reads occur at the position specified by the file pointer if supported by the device.</para>
    ''' <para>This function is designed for both synchronous and asynchronous operations. For a similar function designed solely for asynchronous operation, see <a href="http://msdn.microsoft.com/en-us/library/aa365468(VS.85).aspx">ReadFileEx</a>.</para>
    ''' </summary>
    ''' <param name="handle">A handle to the device (for example, a file, file stream, physical disk, volume, console buffer, tape drive, socket, communications resource, mailslot, or pipe).</param>
    ''' <param name="buffer">A pointer to the buffer that receives the data read from a file or device.</param>
    ''' <param name="byteToRead">The maximum number of bytes to be read.</param>
    ''' <param name="bytesRead">A pointer to the variable that receives the number of bytes read when using a synchronous hFile parameter. ReadFile sets this value to zero before doing any work or error checking. Use NULL for this parameter if this is an asynchronous operation to avoid potentially erroneous results.</param>
    ''' <param name="lpOverlapped">A pointer to an OVERLAPPED structure is required if the hFile parameter was opened with FILE_FLAG_OVERLAPPED, otherwise it can be NULL.</param>
    ''' <returns>If the function succeeds, the return value is nonzero (TRUE).</returns>
    ''' <remarks>see http://msdn.microsoft.com/en-us/library/aa365467(VS.85).aspx </remarks>
    <System.Runtime.InteropServices.DllImport("kernel32", SetLastError:=True)> _
     Function ReadFile(ByVal handle As System.Int32, ByVal buffer() As Byte, ByVal byteToRead As System.Int32, ByRef bytesRead As System.Int32, ByVal lpOverlapped As IntPtr) As System.Int32
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
    End Function

    ''' <summary>
    ''' <para>Creates or opens a file or I/O device. The most commonly used I/O devices are as follows: file, file stream, directory, physical disk, volume, console buffer, tape drive, communications resource, mailslot, and pipe. The function returns a handle that can be used to access the file or device for various types of I/O depending on the file or device and the flags and attributes specified.</para>
    ''' <para>To perform this operation as a transacted operation, which results in a handle that can be used for transacted I/O, use the <a href="http://msdn.microsoft.com/en-us/library/aa363859(VS.85).aspx">CreateFileTransacted</a> function.</para>
    ''' </summary>
    ''' <param name="filename">The name of the file or device to be created or opened. </param>
    ''' <param name="desiredAccess">The requested access to the file or device, which can be summarized as read, write, both or neither (zero).</param>
    ''' <param name="shareMode">The requested sharing mode of the file or device, which can be read, write, both, delete, all of these, or none (refer to the following table). Access requests to attributes or extended attributes are not affected by this flag.</param>
    ''' <param name="attributes">A pointer to a SECURITY_ATTRIBUTES structure that contains two separate but related data members: an optional security descriptor, and a Boolean value that determines whether the returned handle can be inherited by child processes.</param>
    ''' <param name="creationDisposition">An action to take on a file or device that exists or does not exist.</param>
    ''' <param name="flagsAndAttributes">The file or device attributes and flags, FILE_ATTRIBUTE_NORMAL being the most common default value for files.</param>
    ''' <param name="templateFile">A valid handle to a template file with the GENERIC_READ access right. The template file supplies file attributes and extended attributes for the file that is being created.</param>
    ''' <returns>
    ''' <para>If the function succeeds, the return value is an open handle to the specified file, device, named pipe, or mail slot.</para>
    ''' <para>If the function fails, the return value is INVALID_HANDLE_VALUE. To get extended error information, call GetLastError.</para>
    ''' </returns>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/aa363858.aspx for full documentation!</remarks>
    <System.Runtime.InteropServices.DllImport("kernel32", SetLastError:=True)> _
     Function CreateFile(ByVal filename As String, ByVal desiredAccess As System.Int32, ByVal shareMode As System.Int32, ByVal attributes As IntPtr, ByVal creationDisposition As System.Int32, ByVal flagsAndAttributes As System.Int32, ByVal templateFile As IntPtr) As System.Int32
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
    End Function

    ''' <summary>
    ''' Writes data to the specified file or input/output (I/O) device. Writes occur at the position specified by the file pointer, if the handle refers to a seeking device.
    ''' </summary>
    ''' <param name="hFile">A handle to the file or I/O device (for example, a file, file stream, physical disk, volume, console buffer, tape drive, socket, communications resource, mailslot, or pipe).</param>
    ''' <param name="lpBuffer">A pointer to the buffer containing the data to be written to the file or device.</param>
    ''' <param name="nNumberOfBytesToWrite">The number of bytes to be written to the file or device.</param>
    ''' <param name="lpNumberOfBytesWritten">A pointer to the variable that receives the number of bytes written when using a synchronous hFile parameter. WriteFile sets this value to zero before doing any work or error checking. Use NULL for this parameter if this is an asynchronous operation to avoid potentially erroneous results.</param>
    ''' <param name="lpOverlapped">A pointer to an <a href="http://msdn.microsoft.com/en-us/library/ms684342(VS.85).aspx">OVERLAPPED</a> structure is required if the hFile parameter was opened with FILE_FLAG_OVERLAPPED, otherwise this parameter can be NULL.</param>
    ''' <returns>If the function succeeds, the return value is nonzero (TRUE).</returns>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/aa365747(VS.85).aspx for full documentation!</remarks>
    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)> _
    Function WriteFile(ByVal hFile As System.Int32, ByVal lpBuffer() As Byte, ByVal nNumberOfBytesToWrite As System.Int32, ByRef lpNumberOfBytesWritten As System.Int32, ByVal lpOverlapped As IntPtr) As Boolean
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
    End Function

    ''' <summary>
    ''' Closes an open object handle.
    ''' </summary>
    ''' <param name="hFile">A valid handle to an open object.</param>
    ''' <returns>If the function succeeds, the return value is nonzero.</returns>
    ''' <remarks>See http://msdn.microsoft.com/en-us/library/ms724211(VS.85).aspx for full Documentation!</remarks>
    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)> _
     Function CloseHandle(ByVal hFile As System.Int32) As Boolean 'ToDo: Unsigned Integers not supported
    End Function

#End Region

#Region "ctors"

    '/ <summary>
    '/ Private constructor.  No instances of the class can be created.
    '/ </summary>
    Sub New()
    End Sub 'New
#End Region

#Region "Public Static Methods"

    ''' <summary>
    ''' Method called when an alternate data stream must be read from.
    ''' </summary>
    ''' <param name="file">The fully qualified name of the file from which
    ''' the ADS data will be read.</param>
    ''' <param name="stream">The name of the stream within the "normal" file
    ''' from which to read.</param>
    ''' <returns>The contents of the file as a string.  It will always return
    ''' at least a zero-length string, even if the file does not exist.
    ''' </returns>
    Public Function Read(ByVal file As String, ByVal stream As String) As String
        Dim fHandle As System.Int32 = CreateFile(file + ":" + stream, GENERIC_READ, FILE_SHARE_READ, IntPtr.Zero, OPEN_EXISTING, 0, IntPtr.Zero) 'ToDo: Unsigned Integers not supported ' Filename
        ' Desired access
        ' Share more
        ' Attributes
        ' Creation attributes
        ' Flags and attributes
        ' Template file
        ' if the handle returned is uint.MaxValue, the stream doesn't exist.
        If fHandle <> System.Int32.MaxValue Then 'ToDo: Unsigned Integers not supported
            ' A handle to the stream within the file was created successfully.
            Dim size As System.Int32 = GetFileSize(fHandle, IntPtr.Zero) 'ToDo: Unsigned Integers not supported
            Dim buffer(size) As Byte
            Dim reader As System.Int32 = System.Int32.MinValue
            'ToDo: Unsigned Integers not supported
            'ToDo: Unsigned Integers not supported

            Dim result As System.Int32 = ReadFile(fHandle, buffer, size, reader, IntPtr.Zero) 'ToDo: Unsigned Integers not supported ' Handle
            ' Data buffer
            ' Bytes to read
            ' Bytes actually read
            ' Overlapped
            CloseHandle(fHandle)

            ' Convert the bytes read into an ASCII string and return it to the caller.
            Return System.Text.Encoding.ASCII.GetString(buffer)
        Else
            ReadError(file, stream)
            Return Nothing
        End If
    End Function 'Read

    ''' <summary>
    ''' The static method to call when data must be written to a stream.
    ''' </summary>
    ''' <param name="data">The string data to embed in the stream in the file</param>
    ''' <param name="file">The fully qualified name of the file with the
    ''' stream into which the data will be written.</param>
    ''' <param name="stream">The name of the stream within the normal file to
    ''' write the data.</param>
    ''' <returns>An unsigned integer of how many bytes were actually written.</returns>
    Public Function Write(ByVal data As String, ByVal file As String, ByVal stream As String) As System.Int32  'ToDo: Unsigned Integers not supported
        ' Convert the string data to be written to an array of ascii characters.
        Dim barData As Byte() = System.Text.Encoding.ASCII.GetBytes(data)
        Dim nReturn As System.Int32 = 0 'ToDo: Unsigned Integers not supported

        Dim fHandle As System.Int32 = CreateFile(file + ":" + stream, GENERIC_WRITE, FILE_SHARE_WRITE, IntPtr.Zero, CREATE_ALWAYS, 0, IntPtr.Zero) 'ToDo: Unsigned Integers not supported ' File name
        ' Desired access
        ' Share mode
        ' Attributes
        ' Creation disposition
        ' Flags and attributes
        ' Template file
        Dim bOK As Boolean = WriteFile(fHandle, barData, CType(barData.Length, System.Int32), nReturn, IntPtr.Zero) 'ToDo: Unsigned Integers not supported ' Handle
        ' Data buffer
        ' Buffer size
        ' Bytes written
        ' Overlapped
        CloseHandle(fHandle)

        ' Throw an exception if the data wasn't written successfully.
        If Not bOK Then
            Throw New System.ComponentModel.Win32Exception(System.Runtime.InteropServices.Marshal.GetLastWin32Error())
        End If
        Return nReturn
    End Function 'Write

    ''' <summary>
    ''' Not Documented!
    ''' </summary>
    ''' <param name="FileName"></param>
    ''' <param name="stream"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ReadError(ByVal FileName As String, ByVal stream As String) As String
        _stream = stream
        _FileName = FileName
        Return "Stream """ + _stream + """ not found in """ + _FileName + """"
    End Function

#End Region
End Module

'Module StreamNotFoundException
'#Region "Private Members"
'    'Private _stream As String = String.Empty
'    'Private _FileName As String
'#End Region

'#Region "ctors"

'    '/ <summary>
'    '/ Constructor called with the name of the file and stream which was
'    '/ unsuccessfully opened.
'    '/ </summary>
'    '/ <param name="file">Fully qualified name of the file in which the stream
'    '/ was supposed to reside.</param>
'    '/ <param name="stream">Stream within the file to open.</param>
'    'Public Sub New(ByVal file As String, ByVal stream As String)
'    '    MyBase.New(String.Empty, file)
'    '    
'    'End Sub 'New
'    'Public Sub ReadError(ByVal FileName As String, ByVal stream As String)
'    '    _stream = stream
'    '    _FileName = FileName
'    'End Sub


'#End Region

'#Region "Public Properties"
'    '/ <summary>
'    '/ Read-only property to allow the user to query the exception to determine
'    '/ the name of the stream that couldn't be found.
'    '/ </summary>

'    'Public ReadOnly Property Stream() As String
'    '    Get
'    '        Return _stream
'    '    End Get
'    'End Property
'#End Region

'#Region "Overridden Properties"
'    '/ <summary>
'    '/ Overridden Message property to return a concise string describing the
'    '/ exception.
'    '/ </summary>

'    Public ReadOnly Property Message() As String
'        Get
'            Return "Stream """ + _stream + """ not found in """ + _FileName + """"
'        End Get
'    End Property
'#End Region

'End Module