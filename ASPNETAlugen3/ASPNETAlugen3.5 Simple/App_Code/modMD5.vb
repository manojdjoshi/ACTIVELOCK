Option Strict Off
Option Explicit On 
Imports System.io
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
