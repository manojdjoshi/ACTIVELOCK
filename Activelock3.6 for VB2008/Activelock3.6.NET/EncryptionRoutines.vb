
Imports System.Text
Imports System.Text.UnicodeEncoding
Imports System.Security.Cryptography
Imports System.IO

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

Friend NotInheritable Class EncryptionRoutines

#Region " Private Instance Members "
    Private bKey() As Byte
    Private bIV() As Byte
    Private bInitialised As Boolean = False
    Private rijM As RijndaelManaged = Nothing

    Private headerString As String = "CRYPTOR"
    Private headerBytes(7) As Byte

    Private bCancel As Boolean = False
#End Region
#Region " Public Events And Enums "
    Public Event Progress(ByVal prog As Integer)
    Public Event Finished(ByVal retType As ReturnType)

    Public Enum ReturnType As Integer
        Well = 0
        Badly = 1
        IncorrectPassword = 2
    End Enum
#End Region

    Public Function GenerateHash(ByVal strSource As String) As String
        Return System.Convert.ToBase64String(New SHA384Managed().ComputeHash(New UnicodeEncoding().GetBytes(strSource)))
    End Function

    Public Sub Initialise(ByVal sPWH As String)
        'initialise rijM
        rijM = New RijndaelManaged
        'derive the key and IV using the 
        'PasswordDeriveBytes class
        Dim pdb As Rfc2898DeriveBytes = _
                   New Rfc2898DeriveBytes(sPWH, New MD5CryptoServiceProvider().ComputeHash(ConvertStringToBytes(sPWH)))
        'extract the key and IV
        bKey = pdb.GetBytes(32)
        bIV = pdb.GetBytes(16)

        'initialise headerBytes
        headerBytes = ConvertStringToBytes(headerString)

        With rijM
            .Key = bKey '256 bit key
            .IV = bIV '128 bit IV
            .BlockSize = 128 '128 bit BlockSize
            .Padding = PaddingMode.PKCS7
        End With

        bInitialised = True
    End Sub

    Public Sub CancelTransform()
        If Not bInitialised Then Return
        bCancel = True
    End Sub

    Public Function TransformFile(ByVal sInFile As String, ByVal sOutFile As String, Optional ByVal encrypt As Boolean = True) As Boolean
        'make sure that all the initialisation has been completed:
        If Not bInitialised Then RaiseEvent Finished(ReturnType.Badly) : Return False
        If Not IO.File.Exists(sInFile) Then RaiseEvent Finished(ReturnType.Badly) : Return False

        Dim fsIn As FileStream = Nothing
        Dim fsOut As FileStream = Nothing
        Dim encStream As CryptoStream = Nothing
        Dim retVal As ReturnType = ReturnType.Badly
        Try
            'create the input and output streams:
            fsIn = New FileStream(sInFile, FileMode.Open, FileAccess.Read)
            fsOut = New FileStream(sOutFile, FileMode.Create, FileAccess.Write)

            'some helper variables
            Dim bBuffer(4096) As Byte '4KB buffer
            Dim lBytesRead As Long = 0
            Dim lFileSize As Long = fsIn.Length
            Dim lBytesToWrite As Integer

            If encrypt Then
                encStream = New CryptoStream(fsOut, rijM.CreateEncryptor(bKey, bIV), CryptoStreamMode.Write)
                'write the header to the output file for use when decrypting it
                encStream.Write(headerBytes, 0, headerBytes.Length)
                'this is the main encryption routine. it loops over the input data in blocks of 4KB,
                'and writes the encrypted data to disk
                Do
                    If bCancel Then Exit Try
                    lBytesToWrite = fsIn.Read(bBuffer, 0, 4096)
                    If lBytesToWrite = 0 Then Exit Do
                    encStream.Write(bBuffer, 0, lBytesToWrite)
                    lBytesRead += lBytesToWrite
                    RaiseEvent Progress(CInt((lBytesRead / lFileSize) * 100))
                Loop
                RaiseEvent Progress(100)
                retVal = ReturnType.Well
            Else
                encStream = New CryptoStream(fsIn, rijM.CreateDecryptor(bKey, bIV), CryptoStreamMode.Read)

                'read in the header
                Dim test(headerBytes.Length) As Byte
                encStream.Read(test, 0, headerBytes.Length)

                'check to see if the file header reads correctly.
                'if it doesn't, then close the stream & jump out
                If ConvertBytesToString(test) <> headerString Then
                    encStream.Clear()
                    encStream = Nothing
                    retVal = ReturnType.IncorrectPassword
                    Exit Try
                End If

                'this is the main decryption routine. it loops over the input data in blocks of 4KB,
                'and writes the decrypted data to disk
                Do
                    If bCancel Then
                        'if the cancel flag is set,
                        'then jump out
                        encStream.Clear()
                        encStream = Nothing
                        Exit Try
                    End If
                    lBytesToWrite = encStream.Read(bBuffer, 0, 4096)
                    If lBytesToWrite = 0 Then Exit Do
                    fsOut.Write(bBuffer, 0, lBytesToWrite)
                    lBytesRead += lBytesToWrite
                    RaiseEvent Progress(CInt((lBytesRead / lFileSize) * 100))
                Loop
                RaiseEvent Progress(100)
                retVal = ReturnType.Well
            End If
        Catch ex As Exception
            Console.WriteLine("*****************ERROR*****************")
            Console.WriteLine(ex.ToString)
            Console.WriteLine("****************/ERROR*****************")
        Finally
            'close all I/O streams (encStream first)
            If Not encStream Is Nothing Then
                encStream.Close()
            End If
            If Not fsOut Is Nothing Then
                fsOut.Close()
            End If
            If Not fsIn Is Nothing Then
                fsIn.Close()
            End If
        End Try
        'only delete the file if the password was bad, and
        'therefore its only an empty file
        If retVal = ReturnType.IncorrectPassword Then
            IO.File.Delete(sOutFile)
        End If
        'raise the Finished event, and then reset bCancel
        RaiseEvent Finished(retVal)
        bCancel = False
    End Function

    Public Function ConvertStringToBytes(ByVal sString As String) As Byte()
        Return New UnicodeEncoding().GetBytes(sString)
    End Function

    Public Function ConvertBytesToString(ByVal bytes() As Byte) As String
        Return New UnicodeEncoding().GetString(bytes)
    End Function
End Class
