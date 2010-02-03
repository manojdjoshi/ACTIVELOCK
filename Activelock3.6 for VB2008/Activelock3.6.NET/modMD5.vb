Option Strict Off
Option Explicit On 
Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

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

''' <summary>
''' MD5 Hashing Module
''' </summary>
''' <remarks></remarks>
Module modMD5

    '===============================================================================
    ' Name: modMD5
    ' Purpose: MD5 Hashing Module
    ' Date Created: 
    ' Date Last Modified:
    ' Functions:
    ' Properties:
    ' Methods:
    ' Started: 06.16.1998
    ' Modified: 03.25.2006
    '===============================================================================
    '===============================================================================
    ' Name: Function Hash
    ' Input:
    '   ByRef strMessage As String - String to be hashed
    ' Output:
    '   String - Hashed string
    ' Purpose: MD5 Hash function
    ' Remarks: None
    '===============================================================================
    ''' <summary>
    ''' MD5 Hash function
    ''' </summary>
    ''' <param name="strMessage">String to be hashed</param>
    ''' <returns>Hashed string</returns>
    ''' <remarks></remarks>
	Public Function Hash(ByRef strMessage As String) As String
        'Create an encoding object to ensure the encoding standard for the source text
        Dim Ue As New UnicodeEncoding
        'Retrieve a byte array based on the source text
        Dim ByteSourceText() As Byte = Ue.GetBytes(strMessage)
        'Instantiate an MD5 Provider object
        Dim Md5 As New MD5CryptoServiceProvider
        'Compute the hash value from the source
        Dim ByteHash() As Byte = Md5.ComputeHash(ByteSourceText)
        'And convert it to String format for return
        Return Convert.ToBase64String(ByteHash)
    End Function
    '===============================================================================
    ' Name: Function ComputeHash
    ' Input:
    '   ByRef strMessage As String - String to be hashed
    ' Output:
    '   String - Hashed string
    ' Purpose: MD5 Hash function
    ' Remarks: This function is primarily used by the Short Key Function 
    ' to hash strings; it matches the hash generated in the VB6 version
    ' therefore both .NET and VB6 versions generate the same serial number and key
    '===============================================================================
    ''' <summary>
    ''' MD5 Hash function
    ''' </summary>
    ''' <param name="strMessage">String to be hashed</param>
    ''' <returns>Hashed string</returns>
    ''' <remarks>This function is primarily used by the Short Key Function to hash strings; it matches the hash generated in the VB6 version therefore both .NET and VB6 versions generate the same serial number and key</remarks>
    Public Function ComputeHash(ByRef strMessage As String) As String
        Dim hashedDataBytes As Byte()
        Dim encoder As New UTF7Encoding
        Dim md5Hasher As New MD5CryptoServiceProvider
        hashedDataBytes = md5Hasher.ComputeHash(encoder.GetBytes(strMessage))
        Dim strHash As String = BitConverter.ToString(hashedDataBytes)
        Return strHash.Replace("-", "").ToLower
    End Function

End Module