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
'*
'===============================================================================
' Name: modBase64
' Purpose: This module contains Base-64 encoding and decoding routines.
' Functions:
' Properties:
' Methods:
' Started: 04.21.2005
' Modified: 03.25.2006
'===============================================================================
Option Strict On
Option Explicit On 

Module modBase64

  Private Const base64 As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

  Public Function Base64_Encode(ByRef DecryptedText As String) As String
    '===============================================================================
    ' Name: Function Base64_Encode
    ' Input:
    '   ByRef DecryptedText As String - The decrypted string
    ' Output:
    '   String - Base64 encoded string
    ' Purpose: Return the Base64 encoded string
    ' Remarks: None
    '===============================================================================
    Dim c1, c2, c3 As Integer
    Dim w1 As Double
    Dim w2 As Double
    Dim w3 As Double
    Dim w4 As Double
    Dim N As Integer
    Dim retry As String = String.Empty
    For N = 1 To DecryptedText.Length Step 3
      c1 = Convert.ToInt16(DecryptedText.Substring(N - 1, 1).Chars(0))
      c2 = Convert.ToInt16(DecryptedText.Substring(N, 1).Chars(0))
      c3 = Convert.ToInt16(DecryptedText.Substring(N + 1, 1).Chars(0))

      w1 = Int(c1 / 4)
      w2 = CShort(c1 And 3) * 16 + Int(c2 / 16)
      If DecryptedText.Length >= N + 1 Then w3 = CShort(c2 And 15) * 4 + Int(c3 / 64) Else w3 = -1
      If DecryptedText.Length >= N + 2 Then w4 = c3 And 63 Else w4 = -1
      retry = retry & mimeencode(w1) & mimeencode(w2) & mimeencode(w3) & mimeencode(w4)
    Next
    Base64_Encode = retry
  End Function

  Public Function Base64_Decode(ByRef a As String) As String
    '===============================================================================
    ' Name: Function Base64_Decode
    ' Input:
    '   ByRef a As String - The string to be decoded
    ' Output:
    '   String - Base64 decoded string
    ' Purpose: Return the Base64 decoded string
    ' Remarks: None
    '===============================================================================
    Dim w1 As Short
    Dim w2 As Short
    Dim w3 As Short
    Dim w4 As Short
    Dim N As Integer
    Dim retry As String

    For N = 1 To a.Length Step 4
      w1 = mimedecode(a.Substring(N - 1, 1))
      w2 = mimedecode(a.Substring(N, 1))
      w3 = mimedecode(a.Substring(N + 1, 1))
      w4 = mimedecode(a.Substring(N + 2, 1))
      If w2 >= 0 Then retry = retry & Chr(Convert.ToInt32((w1 * 4 + Int(w2 / 16))) And 255)
      If w3 >= 0 Then retry = retry & Chr(Convert.ToInt32(w2 * 16 + Int(w3 / 4)) And 255)
      If w4 >= 0 Then retry = retry & Chr((w3 * 64 + w4) And 255)
    Next
    Base64_Decode = retry
  End Function

  Private Function mimeencode(ByRef w As Double) As String
    '===============================================================================
    ' Name: Function mimeencode
    ' Input:
    '   ByRef w As Integer - Input integer
    ' Output:
    '   String
    ' Purpose: Used by the Base64_encode function
    ' Remarks: None
    '===============================================================================
    If w >= 0 Then
      mimeencode = base64.Substring(Convert.ToInt32(w), 1) 'Mid(base64, w + 1, 1)
    Else
      mimeencode = ""
    End If
  End Function

  Private Function mimedecode(ByRef a As String) As Short
    '===============================================================================
    ' Name: Function mimedecode
    ' Input:
    '   ByRef a As String - Input string
    ' Output:
    '   Integer
    ' Purpose: Used by the Base64_decode function
    ' Remarks: None
    '===============================================================================
    If a.Length = 0 Then
      mimedecode = -1
      Exit Function
    End If
    mimedecode = Convert.ToInt16(base64.IndexOf(a) - 1) 'InStr(base64, a) - 1
  End Function
End Module