Option Strict Off
Option Explicit On

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

Module modBase64

    '===============================================================================
    ' Name: modBase64
    ' Purpose: This module contains Base-64 encoding and decoding routines.
    ' Functions:
    ' Properties:
    ' Methods:
    ' Started: 04.21.2005
    ' Modified: 03.25.2006
    '===============================================================================
    ' @author activelock-admins
    ' @version 3.3.0
    ' @date 03.25.2006

	Private Const base64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

	'===============================================================================
	' Name: Function Base64_Encode
	' Input:
	'   ByRef DecryptedText As String - The decrypted string
	' Output:
	'   String - Base64 encoded string
	' Purpose: Return the Base64 encoded string
	' Remarks: None
	'===============================================================================
	Public Function Base64_Encode(ByRef DecryptedText As String) As String

        Dim byt As Byte() = System.Text.Encoding.UTF8.GetBytes(DecryptedText)
        ' convert the byte array to a Base64 string
        Base64_Encode = Convert.ToBase64String(byt)

        'Dim c1, c2, c3 As Integer
        'Dim w1 As Short
        'Dim w2 As Short
        'Dim w3 As Short
        'Dim w4 As Short
        'Dim N As Short
        'Dim retry As String = String.Empty

        'For N = 1 To Len(DecryptedText) Step 3
        '    c1 = Asc(Mid(DecryptedText, N, 1))
        '    c2 = Asc(Mid(DecryptedText, N + 1, 1) & Chr(0))
        '    c3 = Asc(Mid(DecryptedText, N + 2, 1) & Chr(0))
        '    w1 = Int(c1 / 4)
        '    w2 = CShort(c1 And 3) * 16 + Int(c2 / 16)
        '    If Len(DecryptedText) >= N + 1 Then w3 = CShort(c2 And 15) * 4 + Int(c3 / 64) Else w3 = -1
        '    If Len(DecryptedText) >= N + 2 Then w4 = c3 And 63 Else w4 = -1
        '    retry = retry & mimeencode(w1) & mimeencode(w2) & mimeencode(w3) & mimeencode(w4)
        'Next
        'Base64_Encode = retry

    End Function

    '===============================================================================
    ' Name: Function Base64_Decode
    ' Input:
    '   ByRef a As String - The string to be decoded
    ' Output:
    '   String - Base64 decoded string
    ' Purpose: Return the Base64 decoded string
    ' Remarks: None
    '===============================================================================
    Public Function Base64_Decode(ByRef a As String) As String

        Dim b As Byte() = Convert.FromBase64String(a)
        Base64_Decode = System.Text.Encoding.UTF8.GetString(b)

        'Dim w1 As Short
        'Dim w2 As Short
        'Dim w3 As Short
        'Dim w4 As Short
        'Dim N As Short
        'Dim retry As String = String.Empty

        'For N = 1 To Len(a) Step 4
        '    w1 = mimedecode(Mid(a, N, 1))
        '    w2 = mimedecode(Mid(a, N + 1, 1))
        '    w3 = mimedecode(Mid(a, N + 2, 1))
        '    w4 = mimedecode(Mid(a, N + 3, 1))
        '    If w2 >= 0 Then retry = retry & Chr((w1 * 4 + Int(w2 / 16)) And 255)
        '    If w3 >= 0 Then retry = retry & Chr((w2 * 16 + Int(w3 / 4)) And 255)
        '    If w4 >= 0 Then retry = retry & Chr((w3 * 64 + w4) And 255)
        'Next
        'Base64_Decode = retry

    End Function

    '===============================================================================
    ' Name: Function mimeencode
    ' Input:
    '   ByRef w As Integer - Input integer
    ' Output:
    '   String
    ' Purpose: Used by the Base64_encode function
    ' Remarks: None
    '===============================================================================
    Private Function mimeencode(ByRef w As Short) As String
        If w >= 0 Then mimeencode = Mid(base64, w + 1, 1) Else mimeencode = ""
    End Function

    '===============================================================================
    ' Name: Function mimedecode
    ' Input:
    '   ByRef a As String - Input string
    ' Output:
    '   Integer
    ' Purpose: Used by the Base64_decode function
    ' Remarks: None
    '===============================================================================
    Private Function mimedecode(ByRef a As String) As Short
        If Len(a) = 0 Then mimedecode = -1 : Exit Function
        mimedecode = InStr(base64, a) - 1
    End Function
End Module