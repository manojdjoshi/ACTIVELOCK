'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
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
Option Strict On
Option Explicit On 

Imports System.IO
Imports System.Text
Imports System.Security.Cryptography


Namespace ASPNETAlugen3


Public Module modMD5
    '  ///////////////////////////////////////////////////////////////////////
    '  / Filename:  modMD5.bas                                               /
    '  / Version:   3.0.0.0                                                  /
    '  / Purpose:   MD5 Hashing Module                                       /
    '  / (C) 1998 Joseph Smugeresky                                          /
    '  /                                                                     /
    '  / Date Created:         June 16, 1998 - JS                            /
    '  / Date Last Modified:   July 07, 2003 - MEC                           /
    '  ///////////////////////////////////////////////////////////////////////
    '
    '
    '* ///////////////////////////////////////////////////////////////////////
    '  /                        MODULE TO DO LIST                            /
    '  ///////////////////////////////////////////////////////////////////////
    '
    '   [ ] Nothing to do :)
    '
    '  ///////////////////////////////////////////////////////////////////////
    '  /                        MODULE CHANGE LOG                            /
    '  ///////////////////////////////////////////////////////////////////////
    '
    '   07.07.03 - mcrute  - Updated the header comments for this file.
    '   07.31.03 - th2tran - Changed Integer defs to Long to handle large data.
    '
    '
    '  ///////////////////////////////////////////////////////////////////////
    '  /                MODULE CODE BEGINS BELOW THIS LINE                   /
    '  ///////////////////////////////////////////////////////////////////////

    ' Hash function
    ' @param strMessage     String to be hashed
    ' @return               Hashed string
    '
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
End Module
End Namespace
