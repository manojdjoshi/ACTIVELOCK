Imports System.IO
Imports System.Diagnostics
Imports System.Text
Imports System.Security.Cryptography
Imports System.Runtime.InteropServices

Module modProjectSpecific

    Public fDisableNotifications As Boolean
    Public Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
    Public Function Replace_Renamed(ByRef sIn As String, ByRef sFind As String, ByRef sReplace As String, Optional ByRef nStart As Integer = 1, Optional ByRef nCount As Integer = -1, Optional ByRef bCompare As CompareMethod = CompareMethod.Binary) As String

        Dim nC As Integer
        Dim nPos As Short
        Dim sOut As String
        sOut = sIn
        nPos = InStr(nStart, sOut, sFind, bCompare)
        If nPos = 0 Then GoTo EndFn
        Do
            nC = nC + 1
            sOut = Left(sOut, nPos - 1) & sReplace & Mid(sOut, nPos + Len(sFind))
            If nCount <> -1 And nC >= nCount Then Exit Do
            nPos = InStr(nStart, sOut, sFind, bCompare)
        Loop While nPos > 0
EndFn:
        Replace_Renamed = sOut
    End Function
    Structure sData_LocalType
        Dim sDataString As String
    End Structure
    Public Function ReadFile(ByVal sPath As String, ByRef sData As String) As Integer

        Dim c As New CRC32
        Dim crc As Integer = 0

        ' CRC32 Hash:
        Dim f As FileStream = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        crc = c.GetCrc32(f)
        f.Close()

        ' File size:
        'f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        'txtSize.Text = String.Format("{0}", f.Length)
        'f.Close()
        'txtCrc32.Text = String.Format("{0:X8}", crc)
        'txtTime.Text = String.Format("{0}", h.ElapsedTime)

        ' Run MD5 Hash
        f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        Dim md5 As MD5CryptoServiceProvider = New MD5CryptoServiceProvider
        md5.ComputeHash(f)
        f.Close()

        Dim hash As Byte() = md5.Hash
        Dim buff As StringBuilder = New StringBuilder
        Dim hashByte As Byte
        For Each hashByte In hash
            buff.Append(String.Format("{0:X1}", hashByte))
        Next
        sData = buff.ToString() 'MD5 String

        ' Run SHA-1 Hash
        'f = New FileStream(sPath, FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
        'Dim sha1 As SHA1CryptoServiceProvider = New SHA1CryptoServiceProvider
        'sha1.ComputeHash(f)
        'f.Close()
        'hash = SHA1.Hash
        'buff = New StringBuilder
        'For Each hashByte In hash
        '    buff.Append(String.Format("{0:X1}", hashByte))
        'Next
        'txtSHA1.Text = buff.ToString()

        ReadFile = Len(sData)
        Exit Function
Hell:
        Err.Raise(Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext)
    End Function
    ' Gets the system directory
    '
    Public Function WinSysDir() As String
        Const FIX_LENGTH As Short = 4096
        Dim length As Short
        Dim Buffer As New Microsoft.VisualBasic.Compatibility.VB6.FixedLengthString(FIX_LENGTH)

        length = GetSystemDirectory(Buffer.Value, FIX_LENGTH - 1)
        WinSysDir = Left(Buffer.Value, length)
    End Function

End Module
