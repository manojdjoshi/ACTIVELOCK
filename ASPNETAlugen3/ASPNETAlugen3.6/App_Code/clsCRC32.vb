'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2008 The ActiveLock Software Group (ASG)
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
' This is v2 of the VB CRC32 algorithm provided by Paul
' (wpsjr1@succeed.net) - much quicker than the nasty
' original version I posted.  Excellent work!
'coments by claudiu
'and other comments
'some comments -ismail

Option Strict On
Option Explicit On 

Namespace ASPNETAlugen3

Public Class CRC32
  Private crc32Table() As Integer
  Private Const BUFFER_SIZE As Integer = 1024

  Public Function GetCrc32(ByRef stream As System.IO.Stream) As Integer

    Dim crc32Result As Integer
    crc32Result = &HFFFFFFFF

    Dim buffer(BUFFER_SIZE) As Byte
    Dim readSize As Integer = BUFFER_SIZE

    Dim count As Integer = stream.Read(buffer, 0, readSize)
    Dim i As Integer
    Dim iLookup As Integer
    Do While (count > 0)
      For i = 0 To count - 1
        iLookup = (crc32Result And &HFF) Xor buffer(i)
        crc32Result = ((crc32Result And &HFFFFFF00) \ &H100) And &HFFFFFF   ' nasty shr 8 with vb :/
        crc32Result = crc32Result Xor crc32Table(iLookup)
      Next i
      count = stream.Read(buffer, 0, readSize)
    Loop

    GetCrc32 = Not (crc32Result)

  End Function

  Public Sub New()
    ' This is the official polynomial used by CRC32 in PKZip.
    ' Often the polynomial is shown reversed (04C11DB7).
    Dim dwPolynomial As Integer = &HEDB88320
    Dim i As Integer, j As Integer

    ReDim crc32Table(256)
    Dim dwCrc As Integer

    For i = 0 To 255
      dwCrc = i
      For j = 8 To 1 Step -1
        If (dwCrc And 1) = 0 Then
          dwCrc = Convert.ToInt32(((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF)
          dwCrc = dwCrc Xor dwPolynomial
        Else
          dwCrc = Convert.ToInt32(((dwCrc And &HFFFFFFFE) \ 2&) And &H7FFFFFFF)
        End If
      Next j
      crc32Table(i) = dwCrc
    Next i
  End Sub

End Class

End Namespace
