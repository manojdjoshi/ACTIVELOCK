Option Strict Off
Option Explicit On 
Imports System.IO
Module modMain
	'*   ActiveLock
	'*   Copyright 1998-2002 Nelson Ferraz
	'*   Copyright 2003 The ActiveLock Software Group (ASG)
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
	
	''
	' This module handles contains common utility routines that can be shared
	' between ActiveLock and the client application.
	'
    Private Declare Function MapFileAndCheckSum Lib "imagehlp" Alias "MapFileAndCheckSumA" (ByVal FileName As String, ByRef HeaderSum As Integer, ByRef CheckSum As Integer) As Integer
	Public Declare Function GetSystemDirectory Lib "kernel32.dll"  Alias "GetSystemDirectoryA"(ByVal lpBuffer As String, ByVal nSize As Integer) As Integer
	
	' Application Encryption keys:
	' !!!WARNING!!!
	' It is alright to use these same keys for testing your application.  But it is highly recommended
	' that you generate your own set of keys to use before deploying your app.

    'VCode for TestApp, 1.0, ALCrypto
    'Public Const PUB_KEY As String = "2CB.2CB.2CB.2CB.2D6.231.35A.53E.42B.2E1.21B.533.441.226.2F7.2CB.2CB.2CB.2CB.2D6.32E.37B.2CB.2CB.2CB.323.2D6.268.205.2D6.226.339.3BD.4C5.42B.483.226.3BD.391.30D.39C.386.370.441.46D.4AF.34F.4C5.441.53E.457.3C8.4D0.44C.268.4BA.512.210.3D3.23C.4E6.21B.4F1.32E.21B.51D.3B2.231.512.318.226.21B.4DB.23C.4E6.39C.4D0.2F7.3D3.507.2D6.483.2EC.23C.318.302.365.4D0.499.436.35A.2D6.391.386.44C.4D0.2D6.318.32E.30D.3BD.457.441.25D.48E.3A7.483.268.323.391.3B2.210.4D0.34F.252.483.226.339.53E.4BA.48E.478.2E1.4AF.4F1.247.2E1.2F7.4FC.3D3.318.386.533.436.436.483.3D3.512.386.3C8.4A4.457.30D.53E.302.4F1.2CB.2CB.370.268.21B.25D.370.344.35A.231.32E.3D3.4C5.231.3BD.499.2F7.4E6.39C.226.4C5.462.386.247.386.2E1.499.462.478.4AF.528.210.252.210.2D6.39C.268.51D.42B.370.4C5.4DB.4BA.4BA.231.2CB.2D6.25D.4F1.3DE.210.37B.29F.29F"

    'VCode for TestApp, version 3.0, NET RSA 1024-bit
    Public Const PUB_KEY As String = "386.391.2CB.21B.210.226.23C.294.386.391.2CB.339.457.533.3B2.42B.4A4.507.457.2AA.294.34F.4C5.44C.507.4A4.507.4F1.2AA.533.2F7.35A.21B.23C.323.39C.462.48E.46D.4E6.512.318.2D6.441.441.4FC.4E6.3D3.3B2.2CB.370.344.205.462.21B.512.344.483.318.2E1.4C5.436.231.210.1D9.51D.23C.247.478.210.441.4F1.436.3BD.2CB.34F.205.441.436.507.273.34F.3D3.4F1.53E.3DE.39C.210.318.1D9.483.226.231.4F1.391.339.339.46D.4AF.4D0.30D.441.4E6.4DB.226.3B2.21B.4C5.3D3.3BD.365.2D6.344.391.35A.365.42B.499.39C.4BA.512.3D3.3B2.3D3.533.34F.35A.302.205.2F7.323.533.1D9.32E.273.344.2F7.3A7.42B.2D6.4F1.48E.2F7.478.226.226.48E.2EC.46D.39C.365.4C5.370.478.370.3A7.252.1D9.533.3A7.478.499.268.507.273.39C.46D.3BD.457.2EC.391.4D0.512.35A.23C.4C5.436.231.252.210.318.4C5.231.365.3D3.44C.512.2CB.53E.247.3DE.478.370.344.323.205.37B.42B.4AF.499.29F.294.205.34F.4C5.44C.507.4A4.507.4F1.2AA.294.2F7.528.4D0.4C5.4BA.457.4BA.4FC.2AA.2CB.37B.2CB.2D6.294.205.2F7.528.4D0.4C5.4BA.457.4BA.4FC.2AA.294.205.386.391.2CB.339.457.533.3B2.42B.4A4.507.457.2AA"
    'VCode for TestApp, version 3.5, NET RSA 1024-bit
    'Public Const PUB_KEY As String = "386.391.2CB.21B.210.226.23C.294.386.391.2CB.339.457.533.3B2.42B.4A4.507.457.2AA.294.34F.4C5.44C.507.4A4.507.4F1.2AA.4FC.2D6.247.39C.4FC.4D0.2E1.3B2.273.512.1D9.4FC.34F.533.2EC.512.499.2D6.44C.499.4FC.4F1.2F7.386.323.344.21B.318.483.2D6.344.302.23C.4E6.533.44C.2CB.205.478.34F.4C5.226.231.4FC.483.4C5.344.226.231.30D.483.507.231.46D.370.35A.302.4DB.226.323.4DB.4C5.51D.3D3.499.441.4E6.42B.344.4AF.4C5.4FC.365.462.528.3C8.252.3C8.37B.3C8.268.42B.252.210.1D9.51D.3B2.4E6.48E.462.386.210.210.4AF.32E.436.25D.3BD.23C.4AF.205.507.457.507.528.3C8.4A4.4DB.4DB.4BA.34F.210.4C5.268.226.268.273.4F1.268.37B.273.210.3B2.252.51D.4D0.4C5.2D6.370.252.210.436.386.210.339.441.231.252.318.528.34F.499.370.436.48E.39C.365.2E1.3A7.21B.344.4A4.483.273.533.252.2E1.533.370.205.507.4A4.2CB.2D6.344.457.4FC.4DB.344.4AF.34F.29F.294.205.34F.4C5.44C.507.4A4.507.4F1.2AA.294.2F7.528.4D0.4C5.4BA.457.4BA.4FC.2AA.2CB.37B.2CB.2D6.294.205.2F7.528.4D0.4C5.4BA.457.4BA.4FC.2AA.294.205.386.391.2CB.339.457.533.3B2.42B.4A4.507.457.2AA"
    
    'Test_App
    'Public Const PUB_KEY$ = "2CB.2CB.2CB.2CB.2D6.231.35A.53E.42B.2E1.21B.533.441.226.2F7.2CB.2CB.2CB.2CB.2D6.32E.37B.2CB.2CB.2CB.323.2E1.48E.4A4.3DE.252.210.39C.210.37B.51D.344.441.386.252.2EC.1D9.51D.323.23C.25D.3A7.2E1.210.35A.1D9.339.39C.2E1.25D.3A7.339.44C.507.32E.231.441.2F7.499.3DE.2E1.318.46D.231.4A4.2EC.365.533.2CB.3D3.37B.3C8.4AF.53E.4A4.51D.25D.23C.4A4.4F1.48E.32E.42B.457.3A7.318.365.4F1.21B.2F7.3B2.339.436.318.34F.4AF.512.4A4.302.23C.507.365.42B.302.323.1D9.4FC.252.457.35A.2D6.4DB.3BD.3D3.4C5.268.4AF.210.441.3A7.528.386.318.3BD.386.231.210.4C5.23C.457.4C5.226.3DE.2D6.273.3DE.302.533.48E.48E.370.4A4.478.2F7.4E6.25D.231.3A7.4D0.4AF.51D.226.226.478.205.25D.48E.32E.533.35A.323.21B.462.4AF.4FC.46D.2D6.34F.37B.3B2.4AF.323.3B2.457.365.528.1D9.51D.23C.46D.25D.441.339.318.4A4.4AF.386.457.35A.4F1.507.21B.37B.29F.29F"

    ''
    ' Verifies the checksum of the typelib containing the specified object.
    ' Returns the checksum.
    '
    Public Function VerifyActiveLockNETdll() As String
        ' CRC32 Hash...
        ' I have modified this routine to read the crc32
        ' of the Activelock3NET.dll directly
        ' since the assembly is not a COM object anymore
        ' the method below is very suitable for .NET and more appropriate
        Dim c As New CRC32
        Dim crc As Integer = 0
        Dim fileName As String = System.Windows.Forms.Application.StartupPath & "\ActiveLock3_5Net.dll"
        If File.Exists(fileName) Then
            Dim f As FileStream = New FileStream(System.Windows.Forms.Application.StartupPath & "\ActiveLock3_5Net.dll", FileMode.Open, FileAccess.Read, FileShare.Read, 8192)
            crc = c.GetCrc32(f)
            ' as the most commonly known format
            VerifyActiveLockNETdll = String.Format("{0:X8}", crc)
            f.Close()
            System.Diagnostics.Debug.WriteLine("Hash: " & crc)
            If VerifyActiveLockNETdll <> Dec("210.247.23C.2CB.210.2E1.226.25D") Then
                ' Encrypted version of "activelock3NET.dll has been corrupted. If you were running a real application, it should terminate at this point."
                MsgBox(Dec("42B.441.4FC.483.512.457.4A4.4C5.441.499.231.35A.2F7.39C.1FA.44C.4A4.4A4.160.478.42B.4F1.160.436.457.457.4BA.160.441.4C5.4E6.4E6.507.4D0.4FC.457.44C.1FA"))
                End
            End If
        Else
            MsgBox(Dec("42B.441.4FC.483.512.457.4A4.4C5.441.499.231.35A.2F7.39C.1FA.44C.4A4.4A4.160.4BA.4C5.4FC.160.462.4C5.507.4BA.44C"))
            End
        End If
    End Function
    ' Simple encrypt of a string
    Public Function Enc(ByRef strdata As String) As String
        Dim i, n As Integer
        Dim sResult As String = Nothing
        n = Len(strdata)
        Dim l As Integer
        For i = 1 To n
            l = Asc(Mid(strdata, i, 1)) * 11
            If sResult = "" Then
                sResult = Hex(l)
            Else
                sResult = sResult & "." & Hex(l)
            End If
        Next i
        Enc = sResult
    End Function

    Public Function Dec(ByRef strdata As String) As String
        Dim arr() As String = Nothing
        arr = Split(strdata, ".")
        Dim sRes As String = Nothing
        Dim i As Integer
        For i = LBound(arr) To UBound(arr)
            sRes = sRes & Chr(CInt("&h" & arr(i)) / 11)
        Next
        Dec = sRes
    End Function

    ' Callback function for rsa_generate()
    Public Sub ProgressUpdate(ByVal param As Integer, ByVal action As Integer, ByVal phase As Integer, ByVal iprogress As Integer)
        'frmMain.DefInstance.UpdateStatus("Progress Update received " & param & ", action: " & action & ", iprogress: " & iprogress)
    End Sub
    '===============================================================================
    ' Name: Function WinSysDir
    ' Input: None
    ' Output:
    '   String - Windows system directory path
    ' Purpose: Gets the Windows system directory
    ' Remarks: None
    '===============================================================================
    Public Function WinSysDir() As String
        Const FIX_LENGTH As Short = 4096
        Dim Length As Short
        Dim Buffer As New VB6.FixedLengthString(FIX_LENGTH)

        Length = GetSystemDirectory(Buffer.Value, FIX_LENGTH - 1)
        WinSysDir = Left(Buffer.Value, Length)
    End Function
End Module