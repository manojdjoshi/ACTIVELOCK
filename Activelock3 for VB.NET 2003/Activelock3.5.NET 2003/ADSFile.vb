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

    <System.Runtime.InteropServices.DllImport("kernel32", SetLastError:=True)> _
     Function GetFileSize(ByVal handle As System.Int32, ByVal size As IntPtr) As System.Int32
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
    End Function

    <System.Runtime.InteropServices.DllImport("kernel32", SetLastError:=True)> _
     Function ReadFile(ByVal handle As System.Int32, ByVal buffer() As Byte, ByVal byteToRead As System.Int32, ByRef bytesRead As System.Int32, ByVal lpOverlapped As IntPtr) As System.Int32
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
    End Function

    <System.Runtime.InteropServices.DllImport("kernel32", SetLastError:=True)> _
     Function CreateFile(ByVal filename As String, ByVal desiredAccess As System.Int32, ByVal shareMode As System.Int32, ByVal attributes As IntPtr, ByVal creationDisposition As System.Int32, ByVal flagsAndAttributes As System.Int32, ByVal templateFile As IntPtr) As System.Int32
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
    End Function

    <System.Runtime.InteropServices.DllImport("kernel32.dll", SetLastError:=True)> _
    Function WriteFile(ByVal hFile As System.Int32, ByVal lpBuffer() As Byte, ByVal nNumberOfBytesToWrite As System.Int32, ByRef lpNumberOfBytesWritten As System.Int32, ByVal lpOverlapped As IntPtr) As Boolean
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
        'ToDo: Unsigned Integers not supported
    End Function

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

    '/ <summary>
    '/ Method called when an alternate data stream must be read from.
    '/ </summary>
    '/ <param name="file">The fully qualified name of the file from which
    '/ the ADS data will be read.</param>
    '/ <param name="stream">The name of the stream within the "normal" file
    '/ from which to read.</param>
    '/ <returns>The contents of the file as a string.  It will always return
    '/ at least a zero-length string, even if the file does not exist.</returns>
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
        End If
    End Function 'Read

    '/ <summary>
    '/ The static method to call when data must be written to a stream.
    '/ </summary>
    '/ <param name="data">The string data to embed in the stream in the file</param>
    '/ <param name="file">The fully qualified name of the file with the
    '/ stream into which the data will be written.</param>
    '/ <param name="stream">The name of the stream within the normal file to
    '/ write the data.</param>
    '/ <returns>An unsigned integer of how many bytes were actually written.</returns>
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