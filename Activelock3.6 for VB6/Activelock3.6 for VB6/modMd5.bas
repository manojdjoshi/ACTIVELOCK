Attribute VB_Name = "modMD5"
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
'===============================================================================
' Name: modMD5
' Purpose: MD5 Hashing Module
' (C) 1998 Joseph Smugeresky
' Date Created:         June 16, 1998 - JS
' Date Last Modified:   July 07, 2003 - MEC
' Functions:
' Properties:
' Methods:
' Started: 06.16.1998
' Modified: 08.15.2005
'===============================================================================

'  ///////////////////////////////////////////////////////////////////////
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
'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////

Option Explicit

Private lngTrack As Long
Private arrLongConversion() As Long
Private arrSplit64() As Byte

Private Const OFFSET_4 = 4294967296#
Private Const MAXINT_4 = 2147483647

Private Const S11 = 7
Private Const S12 = 12
Private Const S13 = 17
Private Const S14 = 22
Private Const S21 = 5
Private Const S22 = 9
Private Const S23 = 14
Private Const S24 = 20
Private Const S31 = 4
Private Const S32 = 11
Private Const S33 = 16
Private Const S34 = 23
Private Const S41 = 6
Private Const S42 = 10
Private Const S43 = 15
Private Const S44 = 21

Dim w1 As String, w2 As String, w3 As String

'===============================================================================
' Name: Function MD5Round
' Input:
'   ByRef strRound As String
'   ByRef a As Long
'   ByRef b As Long
'   ByRef c As Long
'   ByRef d As Long
'   ByRef X As Long
'   ByRef s As Long
'   ByRef ac As Long
' Output: Long
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Function MD5Round(strRound As String, a As Long, b As Long, C As Long, d As Long, X As Long, S As Long, ac As Long) As Long

    Select Case strRound
        Case Is = "FF"
            a = MD5LongAdd4(a, (b And C) Or (Not (b) And d), X, ac)
            a = MD5Rotate(a, S)
            a = MD5LongAdd(a, b)
        Case Is = "GG"
            a = MD5LongAdd4(a, (b And d) Or (C And Not (d)), X, ac)
            a = MD5Rotate(a, S)
            a = MD5LongAdd(a, b)
        Case Is = "HH"
            a = MD5LongAdd4(a, b Xor C Xor d, X, ac)
            a = MD5Rotate(a, S)
            a = MD5LongAdd(a, b)
        Case Is = "II"
            a = MD5LongAdd4(a, C Xor (b Or Not (d)), X, ac)
            a = MD5Rotate(a, S)
            a = MD5LongAdd(a, b)
    End Select

End Function


'===============================================================================
' Name: Function MD5Rotate
' Input:
'   ByRef lngValue As Long
'   ByRef lngBits As Long
' Output: Long
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Function MD5Rotate(lngValue As Long, lngBits As Long) As Long

    Dim lngSign As Long
    Dim lngI As Long

    lngBits = (lngBits Mod 32)

    If lngBits = 0 Then MD5Rotate = lngValue: Exit Function

    For lngI = 1 To lngBits
        lngSign = lngValue And &HC0000000
        lngValue = (lngValue And &H3FFFFFFF) * 2
        lngValue = lngValue Or ((lngSign < 0) And 1) Or (CBool(lngSign And &H40000000) And &H80000000)
    Next

    MD5Rotate = lngValue

End Function


'===============================================================================
' Name: Function TRID
' Input: None
' Output: String
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Function TRID() As String

    Dim sngNum As Single
    Dim strResult As String

    sngNum = Rnd(2147483648#)
    strResult = CStr(sngNum)

    strResult = Replace(strResult, "0.", "")
    strResult = Replace(strResult, ".", "")
    strResult = Replace(strResult, "E-", "")

    TRID = strResult

End Function


'===============================================================================
' Name: Function MD564Split
' Input:
'   ByRef lngLength As Long
'   ByRef bytBuffer As Byte
' Output: String
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Function MD564Split(lngLength As Long, bytBuffer() As Byte) As String

    Dim lngBytesTotal As Long, lngBytesToAdd As Long
    Dim intLoop As Long, intLoop2 As Long, lngTrace As Long
    Dim intInnerLoop As Long, intLoop3 As Long

    lngBytesTotal = lngTrack Mod 64
    lngBytesToAdd = 64 - lngBytesTotal
    lngTrack = (lngTrack + lngLength)

    If lngLength >= lngBytesToAdd Then
        For intLoop = 0 To lngBytesToAdd - 1
            arrSplit64(lngBytesTotal + intLoop) = bytBuffer(intLoop)
        Next intLoop

        MD5Conversion arrSplit64

        lngTrace = (lngLength) Mod 64

        For intLoop2 = lngBytesToAdd To lngLength - intLoop - lngTrace Step 64
            For intInnerLoop = 0 To 63
                arrSplit64(intInnerLoop) = bytBuffer(intLoop2 + intInnerLoop)
            Next intInnerLoop

            MD5Conversion arrSplit64

        Next intLoop2

        lngBytesTotal = 0
    Else

      intLoop2 = 0

    End If

    For intLoop3 = 0 To lngLength - intLoop2 - 1

        arrSplit64(lngBytesTotal + intLoop3) = bytBuffer(intLoop2 + intLoop3)

    Next intLoop3

End Function


'===============================================================================
' Name: Function MD5StringArray
' Input:
'   ByRef strInput As String
' Output: None
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Function MD5StringArray(strInput As String) As Byte()

    Dim intLoop As Long
    Dim bytBuffer() As Byte
    ReDim bytBuffer(Len(strInput))

    For intLoop = 0 To Len(strInput) - 1
        bytBuffer(intLoop) = Asc(Mid(strInput, intLoop + 1, 1))
    Next intLoop

    MD5StringArray = bytBuffer

End Function


'===============================================================================
' Name: Sub MD5Conversion
' Input:
'   ByRef bytBuffer As Byte
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Private Sub MD5Conversion(bytBuffer() As Byte)

    Dim X(16) As Long, a As Long
    Dim b As Long, C As Long
    Dim d As Long

    a = arrLongConversion(1)
    b = arrLongConversion(2)
    C = arrLongConversion(3)
    d = arrLongConversion(4)

    MD5Decode 64, X, bytBuffer

    MD5Round "FF", a, b, C, d, X(0), S11, -680876936
    MD5Round "FF", d, a, b, C, X(1), S12, -389564586
    MD5Round "FF", C, d, a, b, X(2), S13, 606105819
    MD5Round "FF", b, C, d, a, X(3), S14, -1044525330
    MD5Round "FF", a, b, C, d, X(4), S11, -176418897
    MD5Round "FF", d, a, b, C, X(5), S12, 1200080426
    MD5Round "FF", C, d, a, b, X(6), S13, -1473231341
    MD5Round "FF", b, C, d, a, X(7), S14, -45705983
    MD5Round "FF", a, b, C, d, X(8), S11, 1770035416
    MD5Round "FF", d, a, b, C, X(9), S12, -1958414417
    MD5Round "FF", C, d, a, b, X(10), S13, -42063
    MD5Round "FF", b, C, d, a, X(11), S14, -1990404162
    MD5Round "FF", a, b, C, d, X(12), S11, 1804603682
    MD5Round "FF", d, a, b, C, X(13), S12, -40341101
    MD5Round "FF", C, d, a, b, X(14), S13, -1502002290
    MD5Round "FF", b, C, d, a, X(15), S14, 1236535329

    MD5Round "GG", a, b, C, d, X(1), S21, -165796510
    MD5Round "GG", d, a, b, C, X(6), S22, -1069501632
    MD5Round "GG", C, d, a, b, X(11), S23, 643717713
    MD5Round "GG", b, C, d, a, X(0), S24, -373897302
    MD5Round "GG", a, b, C, d, X(5), S21, -701558691
    MD5Round "GG", d, a, b, C, X(10), S22, 38016083
    MD5Round "GG", C, d, a, b, X(15), S23, -660478335
    MD5Round "GG", b, C, d, a, X(4), S24, -405537848
    MD5Round "GG", a, b, C, d, X(9), S21, 568446438
    MD5Round "GG", d, a, b, C, X(14), S22, -1019803690
    MD5Round "GG", C, d, a, b, X(3), S23, -187363961
    MD5Round "GG", b, C, d, a, X(8), S24, 1163531501
    MD5Round "GG", a, b, C, d, X(13), S21, -1444681467
    MD5Round "GG", d, a, b, C, X(2), S22, -51403784
    MD5Round "GG", C, d, a, b, X(7), S23, 1735328473
    MD5Round "GG", b, C, d, a, X(12), S24, -1926607734

    MD5Round "HH", a, b, C, d, X(5), S31, -378558
    MD5Round "HH", d, a, b, C, X(8), S32, -2022574463
    MD5Round "HH", C, d, a, b, X(11), S33, 1839030562
    MD5Round "HH", b, C, d, a, X(14), S34, -35309556
    MD5Round "HH", a, b, C, d, X(1), S31, -1530992060
    MD5Round "HH", d, a, b, C, X(4), S32, 1272893353
    MD5Round "HH", C, d, a, b, X(7), S33, -155497632
    MD5Round "HH", b, C, d, a, X(10), S34, -1094730640
    MD5Round "HH", a, b, C, d, X(13), S31, 681279174
    MD5Round "HH", d, a, b, C, X(0), S32, -358537222
    MD5Round "HH", C, d, a, b, X(3), S33, -722521979
    MD5Round "HH", b, C, d, a, X(6), S34, 76029189
    MD5Round "HH", a, b, C, d, X(9), S31, -640364487
    MD5Round "HH", d, a, b, C, X(12), S32, -421815835
    MD5Round "HH", C, d, a, b, X(15), S33, 530742520
    MD5Round "HH", b, C, d, a, X(2), S34, -995338651

    MD5Round "II", a, b, C, d, X(0), S41, -198630844
    MD5Round "II", d, a, b, C, X(7), S42, 1126891415
    MD5Round "II", C, d, a, b, X(14), S43, -1416354905
    MD5Round "II", b, C, d, a, X(5), S44, -57434055
    MD5Round "II", a, b, C, d, X(12), S41, 1700485571
    MD5Round "II", d, a, b, C, X(3), S42, -1894986606
    MD5Round "II", C, d, a, b, X(10), S43, -1051523
    MD5Round "II", b, C, d, a, X(1), S44, -2054922799
    MD5Round "II", a, b, C, d, X(8), S41, 1873313359
    MD5Round "II", d, a, b, C, X(15), S42, -30611744
    MD5Round "II", C, d, a, b, X(6), S43, -1560198380
    MD5Round "II", b, C, d, a, X(13), S44, 1309151649
    MD5Round "II", a, b, C, d, X(4), S41, -145523070
    MD5Round "II", d, a, b, C, X(11), S42, -1120210379
    MD5Round "II", C, d, a, b, X(2), S43, 718787259
    MD5Round "II", b, C, d, a, X(9), S44, -343485551

    arrLongConversion(1) = MD5LongAdd(arrLongConversion(1), a)
    arrLongConversion(2) = MD5LongAdd(arrLongConversion(2), b)
    arrLongConversion(3) = MD5LongAdd(arrLongConversion(3), C)
    arrLongConversion(4) = MD5LongAdd(arrLongConversion(4), d)

End Sub


'===============================================================================
' Name: Function MD5LongAdd
' Input:
'   ByRef lngVal1 As Long
'   ByRef lngVal2 As Long
' Output: Long
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Function MD5LongAdd(lngVal1 As Long, lngVal2 As Long) As Long

    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&

    MD5LongAdd = MD5LongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))

End Function


'===============================================================================
' Name: Function MD5LongAdd4
' Input:
'   ByRef lngVal1 As Long
'   ByRef lngVal2 As Long
'   ByRef lngVal3 As Long
'   ByRef lngVal4 As Long
' Output: Long
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Function MD5LongAdd4(lngVal1 As Long, lngVal2 As Long, lngVal3 As Long, lngVal4 As Long) As Long

    Dim lngHighWord As Long
    Dim lngLowWord As Long
    Dim lngOverflow As Long

    lngLowWord = (lngVal1 And &HFFFF&) + (lngVal2 And &HFFFF&) + (lngVal3 And &HFFFF&) + (lngVal4 And &HFFFF&)
    lngOverflow = lngLowWord \ 65536
    lngHighWord = (((lngVal1 And &HFFFF0000) \ 65536) + ((lngVal2 And &HFFFF0000) \ 65536) + ((lngVal3 And &HFFFF0000) \ 65536) + ((lngVal4 And &HFFFF0000) \ 65536) + lngOverflow) And &HFFFF&
    MD5LongAdd4 = MD5LongConversion((lngHighWord * 65536#) + (lngLowWord And &HFFFF&))

End Function


'===============================================================================
' Name: Sub MD5Decode
' Input:
'   ByRef intLength As Integer
'   ByRef lngOutBuffer As Long
'   ByRef bytInBuffer As Byte
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Private Sub MD5Decode(intLength As Integer, lngOutBuffer() As Long, bytInBuffer() As Byte)

    Dim intDblIndex As Integer
    Dim intByteIndex As Integer
    Dim dblSum As Double

    intDblIndex = 0

    For intByteIndex = 0 To intLength - 1 Step 4

        dblSum = bytInBuffer(intByteIndex) + bytInBuffer(intByteIndex + 1) * 256# + bytInBuffer(intByteIndex + 2) * 65536# + bytInBuffer(intByteIndex + 3) * 16777216#
        lngOutBuffer(intDblIndex) = MD5LongConversion(dblSum)
        intDblIndex = (intDblIndex + 1)

    Next intByteIndex

End Sub


'===============================================================================
' Name: Function MD5LongConversion
' Input:
'   ByRef dblValue As Double
' Output: Long
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Function MD5LongConversion(dblValue As Double) As Long

    If dblValue < 0 Or dblValue >= OFFSET_4 Then Error 6

    If dblValue <= MAXINT_4 Then
        MD5LongConversion = dblValue
    Else
        MD5LongConversion = dblValue - OFFSET_4
    End If

End Function


'===============================================================================
' Name: Sub MD5Finish
' Input: None
' Output: None
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Sub MD5Finish()

    Dim dblBits As Double
    Dim arrPadding(72) As Byte
    Dim lngBytesBuffered As Long

    arrPadding(0) = &H80

    dblBits = lngTrack * 8

    lngBytesBuffered = lngTrack Mod 64

    If lngBytesBuffered <= 56 Then
        MD564Split (56 - lngBytesBuffered), arrPadding
    Else
        MD564Split (120 - lngTrack), arrPadding
    End If


    arrPadding(0) = MD5LongConversion(dblBits) And &HFF&
    arrPadding(1) = MD5LongConversion(dblBits) \ 256 And &HFF&
    arrPadding(2) = MD5LongConversion(dblBits) \ 65536 And &HFF&
    arrPadding(3) = MD5LongConversion(dblBits) \ 16777216 And &HFF&
    arrPadding(4) = 0
    arrPadding(5) = 0
    arrPadding(6) = 0
    arrPadding(7) = 0

    MD564Split 8, arrPadding

End Sub


'===============================================================================
' Name: Function MD5StringChange
' Input:
'   ByRef lngnum As Long
' Output: String
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Function MD5StringChange(lngnum As Long) As String

        Dim bytA As Byte
        Dim bytB As Byte
        Dim bytC As Byte
        Dim bytD As Byte

        bytA = lngnum And &HFF&
        If bytA < 16 Then
            MD5StringChange = "0" & Hex(bytA)
        Else
            MD5StringChange = Hex(bytA)
        End If

        bytB = (lngnum And &HFF00&) \ 256
        If bytB < 16 Then
            MD5StringChange = MD5StringChange & "0" & Hex(bytB)
        Else
            MD5StringChange = MD5StringChange & Hex(bytB)
        End If

        bytC = (lngnum And &HFF0000) \ 65536
        If bytC < 16 Then
            MD5StringChange = MD5StringChange & "0" & Hex(bytC)
        Else
            MD5StringChange = MD5StringChange & Hex(bytC)
        End If

        If lngnum < 0 Then
            bytD = ((lngnum And &H7F000000) \ 16777216) Or &H80&
        Else
            bytD = (lngnum And &HFF000000) \ 16777216
        End If

        If bytD < 16 Then
            MD5StringChange = MD5StringChange & "0" & Hex(bytD)
        Else
            MD5StringChange = MD5StringChange & Hex(bytD)
        End If

End Function


'===============================================================================
' Name: Function MD5Value
' Input: None
' Output: String
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Private Function MD5Value() As String

    MD5Value = LCase(MD5StringChange(arrLongConversion(1)) & MD5StringChange(arrLongConversion(2)) & MD5StringChange(arrLongConversion(3)) & MD5StringChange(arrLongConversion(4)))

End Function


'===============================================================================
' Name: Function Hash
' Input:
'   ByRef strMessage As String - String to be hashed
' Output:
'   String - Hashed string
' Purpose: MD5 Hash function
' Remarks: None
'===============================================================================
Public Function Hash(strMessage As String) As String
    ' Re-init mod variables
    lngTrack = 0
    ReDim arrLongConversion(4)
    ReDim arrSplit64(63)
    Dim bytBuffer() As Byte
    bytBuffer = MD5StringArray(strMessage)

    MD5Start
    MD564Split Len(strMessage), bytBuffer
    MD5Finish

    Hash = MD5Value

End Function


'===============================================================================
' Name: Sub MD5Start
' Input: None
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Private Sub MD5Start()

    lngTrack = 0
    arrLongConversion(1) = MD5LongConversion(1732584193#)
    arrLongConversion(2) = MD5LongConversion(4023233417#)
    arrLongConversion(3) = MD5LongConversion(2562383102#)
    arrLongConversion(4) = MD5LongConversion(271733878#)

End Sub

'===============================================================================
' Name: Function MD5AA1F
' Input:
'   ByVal tempstr As String
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: Variant
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Function MD5AA1F(ByVal tempstr As String, ByVal w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    MD5AA1F = BigAA1Mod32Add(BigAA1RotLeft(BigAA1Mod32Add(BigAA1Mod32Add(w, tempstr), BigAA1Mod32Add(in_, qdata)), rots), X)
End Function
'===============================================================================
' Name: Sub MD5AA1F1
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AA1F1(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA1XOR(z, BigAA1AND(X, BigAA1XOR(y, z)))
    w = MD5AA1F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AA1F2
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AA1F2(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA1XOR(y, BigAA1AND(z, BigAA1XOR(X, y)))
    w = MD5AA1F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AA1F3
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AA1F3(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA1XOR(X, BigAA1XOR(y, z))
    w = MD5AA1F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AA1F4
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AA1F4(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA1XOR(y, BigAA1OR(X, BigAA1NOT(z)))
    w = MD5AA1F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Function MD5AA1Hash
' Input:
'   ByVal hashthis As String - String to be hashed
' Output: None
' Purpose: MD5 Hash function
' Remarks: None
'===============================================================================
Function MD5AA1Hash(ByVal hashthis As String) As String
    ReDim Buf(0 To 3) As String
    ReDim in_(0 To 15) As String
    Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
    Dim a As String, b As String, C As String, d As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If
    hashthis = hashthis + String$(tempnum2, Chr$(0))
    For loopit = 1 To 8
        hashthis = hashthis + Chr$(tempnum Mod 256)
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit
    
    ' Set magic numbers
    Buf(0) = "67452301"
    Buf(1) = "efcdab89"
    Buf(2) = "98badcfe"
    Buf(3) = "10325476"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = Buf(0)
        b = Buf(1)
        C = Buf(2)
        d = Buf(3)
    
        ' Get the 512 bits
        For loopit = 0 To 15
            in_(loopit) = ""
            For loopinner = 1 To 4
                in_(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_(loopit)
                If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" + in_(loopit)
            Next loopinner
        Next loopit
        
        ' Round 1
        MD5AA1F1 a, b, C, d, in_(0), "d76aa478", 7
        MD5AA1F1 d, a, b, C, in_(1), "e8c7b756", 12
        MD5AA1F1 C, d, a, b, in_(2), "242070db", 17
        MD5AA1F1 b, C, d, a, in_(3), "c1bdceee", 22
        MD5AA1F1 a, b, C, d, in_(4), "f57c0faf", 7
        MD5AA1F1 d, a, b, C, in_(5), "4787c62a", 12
        MD5AA1F1 C, d, a, b, in_(6), "a8304613", 17
        MD5AA1F1 b, C, d, a, in_(7), "fd469501", 22
        MD5AA1F1 a, b, C, d, in_(8), "698098d8", 7
        MD5AA1F1 d, a, b, C, in_(9), "8b44f7af", 12
        MD5AA1F1 C, d, a, b, in_(10), "ffff5bb1", 17
        MD5AA1F1 b, C, d, a, in_(11), "895cd7be", 22
        MD5AA1F1 a, b, C, d, in_(12), "6b901122", 7
        MD5AA1F1 d, a, b, C, in_(13), "fd987193", 12
        MD5AA1F1 C, d, a, b, in_(14), "a679438e", 17
        MD5AA1F1 b, C, d, a, in_(15), "49b40821", 22
        
        ' Round 2
        MD5AA1F2 a, b, C, d, in_(1), "f61e2562", 5
        MD5AA1F2 d, a, b, C, in_(6), "c040b340", 9
        MD5AA1F2 C, d, a, b, in_(11), "265e5a51", 14
        MD5AA1F2 b, C, d, a, in_(0), "e9b6c7aa", 20
        MD5AA1F2 a, b, C, d, in_(5), "d62f105d", 5
        MD5AA1F2 d, a, b, C, in_(10), "02441453", 9
        MD5AA1F2 C, d, a, b, in_(15), "d8a1e681", 14
        MD5AA1F2 b, C, d, a, in_(4), "e7d3fbc8", 20
        MD5AA1F2 a, b, C, d, in_(9), "21e1cde6", 5
        MD5AA1F2 d, a, b, C, in_(14), "c33707d6", 9
        MD5AA1F2 C, d, a, b, in_(3), "f4d50d87", 14
        MD5AA1F2 b, C, d, a, in_(8), "455a14ed", 20
        MD5AA1F2 a, b, C, d, in_(13), "a9e3e905", 5
        MD5AA1F2 d, a, b, C, in_(2), "fcefa3f8", 9
        MD5AA1F2 C, d, a, b, in_(7), "676f02d9", 14
        MD5AA1F2 b, C, d, a, in_(12), "8d2a4c8a", 20
        
        ' Round 3
        MD5AA1F3 a, b, C, d, in_(5), "fffa3942", 4
        MD5AA1F3 d, a, b, C, in_(8), "8771f681", 11
        MD5AA1F3 C, d, a, b, in_(11), "6d9d6122", 16
        MD5AA1F3 b, C, d, a, in_(14), "fde5380c", 23
        MD5AA1F3 a, b, C, d, in_(1), "a4beea44", 4
        MD5AA1F3 d, a, b, C, in_(4), "4bdecfa9", 11
        MD5AA1F3 C, d, a, b, in_(7), "f6bb4b60", 16
        MD5AA1F3 b, C, d, a, in_(10), "bebfbc70", 23
        MD5AA1F3 a, b, C, d, in_(13), "289b7ec6", 4
        MD5AA1F3 d, a, b, C, in_(0), "eaa127fa", 11
        MD5AA1F3 C, d, a, b, in_(3), "d4ef3085", 16
        MD5AA1F3 b, C, d, a, in_(6), "04881d05", 23
        MD5AA1F3 a, b, C, d, in_(9), "d9d4d039", 4
        MD5AA1F3 d, a, b, C, in_(12), "e6db99e5", 11
        MD5AA1F3 C, d, a, b, in_(15), "1fa27cf8", 16
        MD5AA1F3 b, C, d, a, in_(2), "c4ac5665", 23
        
        ' Round 4
        MD5AA1F4 a, b, C, d, in_(0), "f4292244", 6
        MD5AA1F4 d, a, b, C, in_(7), "432aff97", 10
        MD5AA1F4 C, d, a, b, in_(14), "ab9423a7", 15
        MD5AA1F4 b, C, d, a, in_(5), "fc93a039", 21
        MD5AA1F4 a, b, C, d, in_(12), "655b59c3", 6
        MD5AA1F4 d, a, b, C, in_(3), "8f0ccc92", 10
        MD5AA1F4 C, d, a, b, in_(10), "ffeff47d", 15
        MD5AA1F4 b, C, d, a, in_(1), "85845dd1", 21
        MD5AA1F4 a, b, C, d, in_(8), "6fa87e4f", 6
        MD5AA1F4 d, a, b, C, in_(15), "fe2ce6e0", 10
        MD5AA1F4 C, d, a, b, in_(6), "a3014314", 15
        MD5AA1F4 b, C, d, a, in_(13), "4e0811a1", 21
        MD5AA1F4 a, b, C, d, in_(4), "f7537e82", 6
        MD5AA1F4 d, a, b, C, in_(11), "bd3af235", 10
        MD5AA1F4 C, d, a, b, in_(2), "2ad7d2bb", 15
        MD5AA1F4 b, C, d, a, in_(9), "eb86d391", 21
    
        Buf(0) = BigAA1Add(Buf(0), a)
        Buf(1) = BigAA1Add(Buf(1), b)
        Buf(2) = BigAA1Add(Buf(2), C)
        Buf(3) = BigAA1Add(Buf(3), d)
    Next loopouter
    
    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 3
        For loopinner = 3 To 0 Step -1
            hashthis = hashthis + Mid$(Buf(loopit), 1 + 2 * loopinner, 2)
        Next loopinner
    Next loopit
    
    ' And return it
    MD5AA1Hash = hashthis

End Function

'===============================================================================
' Name: Function MD5AA2F
' Input:
'   ByVal tempstr As String
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: Variant
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Function MD5AA2F(ByVal tempstr As String, ByVal w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    MD5AA2F = BigAA2Mod32Add(BigAA2RotLeft(BigAA2Mod32Add(BigAA2Mod32Add(w, tempstr), BigAA2Mod32Add(in_, qdata)), rots), X)

End Function

'===============================================================================
' Name: Sub MD5AA2F1
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AA2F1(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA2XOR(z, BigAA2AND(X, BigAA2XOR(y, z)))
    w = MD5AA2F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AA2F2
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AA2F2(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA2XOR(y, BigAA2AND(z, BigAA2XOR(X, y)))
    w = MD5AA2F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AA2F3
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AA2F3(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA2XOR(X, BigAA2XOR(y, z))
    w = MD5AA2F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AA2F4
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AA2F4(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    
    tempstr = BigAA2XOR(y, BigAA2OR(X, BigAA2NOT(z)))
    w = MD5AA2F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Function MD5AA2Hash
' Input:
'   ByVal hashthis As String - String to be hashed
' Output: None
' Purpose: MD5 Hash function
' Remarks: None
'===============================================================================
Function MD5AA2Hash(ByVal hashthis As String) As String
    ReDim Buf(0 To 3) As String
    ReDim in_(0 To 15) As String
    Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
    Dim a As String, b As String, C As String, d As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If
    hashthis = hashthis + String$(tempnum2, Chr$(0))
    For loopit = 1 To 8
        hashthis = hashthis + Chr$(tempnum Mod 256)
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit
    
    ' Set magic numbers
    Buf(0) = "67452301"
    Buf(1) = "efcdab89"
    Buf(2) = "98badcfe"
    Buf(3) = "10325476"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = Buf(0)
        b = Buf(1)
        C = Buf(2)
        d = Buf(3)
    
        ' Get the 512 bits
        For loopit = 0 To 15
            in_(loopit) = ""
            For loopinner = 1 To 4
                in_(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_(loopit)
                If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" + in_(loopit)
            Next loopinner
        Next loopit
        
        ' Round 1
        MD5AA2F1 a, b, C, d, in_(0), "d76aa478", 7
        MD5AA2F1 d, a, b, C, in_(1), "e8c7b756", 12
        MD5AA2F1 C, d, a, b, in_(2), "242070db", 17
        MD5AA2F1 b, C, d, a, in_(3), "c1bdceee", 22
        MD5AA2F1 a, b, C, d, in_(4), "f57c0faf", 7
        MD5AA2F1 d, a, b, C, in_(5), "4787c62a", 12
        MD5AA2F1 C, d, a, b, in_(6), "a8304613", 17
        MD5AA2F1 b, C, d, a, in_(7), "fd469501", 22
        MD5AA2F1 a, b, C, d, in_(8), "698098d8", 7
        MD5AA2F1 d, a, b, C, in_(9), "8b44f7af", 12
        MD5AA2F1 C, d, a, b, in_(10), "ffff5bb1", 17
        MD5AA2F1 b, C, d, a, in_(11), "895cd7be", 22
        MD5AA2F1 a, b, C, d, in_(12), "6b901122", 7
        MD5AA2F1 d, a, b, C, in_(13), "fd987193", 12
        MD5AA2F1 C, d, a, b, in_(14), "a679438e", 17
        MD5AA2F1 b, C, d, a, in_(15), "49b40821", 22
        
        ' Round 2
        MD5AA2F2 a, b, C, d, in_(1), "f61e2562", 5
        MD5AA2F2 d, a, b, C, in_(6), "c040b340", 9
        MD5AA2F2 C, d, a, b, in_(11), "265e5a51", 14
        MD5AA2F2 b, C, d, a, in_(0), "e9b6c7aa", 20
        MD5AA2F2 a, b, C, d, in_(5), "d62f105d", 5
        MD5AA2F2 d, a, b, C, in_(10), "02441453", 9
        MD5AA2F2 C, d, a, b, in_(15), "d8a1e681", 14
        MD5AA2F2 b, C, d, a, in_(4), "e7d3fbc8", 20
        MD5AA2F2 a, b, C, d, in_(9), "21e1cde6", 5
        MD5AA2F2 d, a, b, C, in_(14), "c33707d6", 9
        MD5AA2F2 C, d, a, b, in_(3), "f4d50d87", 14
        MD5AA2F2 b, C, d, a, in_(8), "455a14ed", 20
        MD5AA2F2 a, b, C, d, in_(13), "a9e3e905", 5
        MD5AA2F2 d, a, b, C, in_(2), "fcefa3f8", 9
        MD5AA2F2 C, d, a, b, in_(7), "676f02d9", 14
        MD5AA2F2 b, C, d, a, in_(12), "8d2a4c8a", 20
        
        ' Round 3
        MD5AA2F3 a, b, C, d, in_(5), "fffa3942", 4
        MD5AA2F3 d, a, b, C, in_(8), "8771f681", 11
        MD5AA2F3 C, d, a, b, in_(11), "6d9d6122", 16
        MD5AA2F3 b, C, d, a, in_(14), "fde5380c", 23
        MD5AA2F3 a, b, C, d, in_(1), "a4beea44", 4
        MD5AA2F3 d, a, b, C, in_(4), "4bdecfa9", 11
        MD5AA2F3 C, d, a, b, in_(7), "f6bb4b60", 16
        MD5AA2F3 b, C, d, a, in_(10), "bebfbc70", 23
        MD5AA2F3 a, b, C, d, in_(13), "289b7ec6", 4
        MD5AA2F3 d, a, b, C, in_(0), "eaa127fa", 11
        MD5AA2F3 C, d, a, b, in_(3), "d4ef3085", 16
        MD5AA2F3 b, C, d, a, in_(6), "04881d05", 23
        MD5AA2F3 a, b, C, d, in_(9), "d9d4d039", 4
        MD5AA2F3 d, a, b, C, in_(12), "e6db99e5", 11
        MD5AA2F3 C, d, a, b, in_(15), "1fa27cf8", 16
        MD5AA2F3 b, C, d, a, in_(2), "c4ac5665", 23
        
        ' Round 4
        MD5AA2F4 a, b, C, d, in_(0), "f4292244", 6
        MD5AA2F4 d, a, b, C, in_(7), "432aff97", 10
        MD5AA2F4 C, d, a, b, in_(14), "ab9423a7", 15
        MD5AA2F4 b, C, d, a, in_(5), "fc93a039", 21
        MD5AA2F4 a, b, C, d, in_(12), "655b59c3", 6
        MD5AA2F4 d, a, b, C, in_(3), "8f0ccc92", 10
        MD5AA2F4 C, d, a, b, in_(10), "ffeff47d", 15
        MD5AA2F4 b, C, d, a, in_(1), "85845dd1", 21
        MD5AA2F4 a, b, C, d, in_(8), "6fa87e4f", 6
        MD5AA2F4 d, a, b, C, in_(15), "fe2ce6e0", 10
        MD5AA2F4 C, d, a, b, in_(6), "a3014314", 15
        MD5AA2F4 b, C, d, a, in_(13), "4e0811a1", 21
        MD5AA2F4 a, b, C, d, in_(4), "f7537e82", 6
        MD5AA2F4 d, a, b, C, in_(11), "bd3af235", 10
        MD5AA2F4 C, d, a, b, in_(2), "2ad7d2bb", 15
        MD5AA2F4 b, C, d, a, in_(9), "eb86d391", 21
    
        Buf(0) = BigAA2Add(Buf(0), a)
        Buf(1) = BigAA2Add(Buf(1), b)
        Buf(2) = BigAA2Add(Buf(2), C)
        Buf(3) = BigAA2Add(Buf(3), d)
    Next loopouter
    
    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 3
        For loopinner = 3 To 0 Step -1
            hashthis = hashthis + Mid$(Buf(loopit), 1 + 2 * loopinner, 2)
        Next loopinner
    Next loopit
    
    ' And return it
    MD5AA2Hash = hashthis

End Function

'===============================================================================
' Name: Function MD5AB1F
' Input:
'   ByVal tempstr As String
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: Variant
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Function MD5AB1F(ByVal tempstr As String, ByVal w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim valueans As String
    Dim loopit As Integer, tempnum As Integer

    w = String$(8 - Len(w), "0") + w
    tempstr = String$(8 - Len(tempstr), "0") + tempstr
    in_ = String$(8 - Len(in_), "0") + in_
    qdata = String$(8 - Len(qdata), "0") + qdata

    For loopit = 8 To 1 Step -1
        tempnum = tempnum + Val("&H" + Mid$(w, loopit, 1)) + Val("&H" + Mid$(tempstr, loopit, 1)) + Val("&H" + Mid$(in_, loopit, 1)) + Val("&H" + Mid$(qdata, loopit, 1))
        valueans = Hex$(tempnum Mod 16) + valueans
        tempnum = Int(tempnum / 16)
    Next loopit

    MD5AB1F = BigAA1Mod32Add(BigAA1RotLeft(valueans, rots), X)
                   
End Function

'===============================================================================
' Name: Sub MD5AB1F1
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AB1F1(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim loopit As Integer
    Dim xn As Integer, yn As Integer, zn As Integer
    
    X = String$(8 - Len(X), "0") + X
    y = String$(8 - Len(y), "0") + y
    z = String$(8 - Len(z), "0") + z

    For loopit = 1 To 8
        xn = Val("&H" + Mid$(X, loopit, 1))
        yn = Val("&H" + Mid$(y, loopit, 1))
        zn = Val("&H" + Mid$(z, loopit, 1))
        tempstr = tempstr + Hex$(zn Xor (xn And (yn Xor zn)))
    Next loopit

    w = MD5AB1F(tempstr, w, X, y, z, in_, qdata, rots)
    
End Sub

'===============================================================================
' Name: Sub MD5AB1F2
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AB1F2(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim loopit As Integer
    Dim xn As Integer, yn As Integer, zn As Integer
    
    X = String$(8 - Len(X), "0") + X
    y = String$(8 - Len(y), "0") + y
    z = String$(8 - Len(z), "0") + z

    For loopit = 1 To 8
        xn = Val("&H" + Mid$(X, loopit, 1))
        yn = Val("&H" + Mid$(y, loopit, 1))
        zn = Val("&H" + Mid$(z, loopit, 1))
        tempstr = tempstr + Hex$(yn Xor (zn And (xn Xor yn)))
    Next loopit

    w = MD5AB1F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AB1F3
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AB1F3(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim loopit As Integer
    Dim xn As Integer, yn As Integer, zn As Integer
    
    X = String$(8 - Len(X), "0") + X
    y = String$(8 - Len(y), "0") + y
    z = String$(8 - Len(z), "0") + z

    For loopit = 1 To 8
        xn = Val("&H" + Mid$(X, loopit, 1))
        yn = Val("&H" + Mid$(y, loopit, 1))
        zn = Val("&H" + Mid$(z, loopit, 1))
        tempstr = tempstr + Hex$(zn Xor xn Xor yn)
    Next loopit

    w = MD5AB1F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AB1F4
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AB1F4(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim loopit As Integer
    Dim xn As Integer, yn As Integer, zn As Integer
    
    X = String$(8 - Len(X), "0") + X
    y = String$(8 - Len(y), "0") + y
    z = String$(8 - Len(z), "0") + z

    For loopit = 1 To 8
        xn = Val("&H" + Mid$(X, loopit, 1))
        yn = Val("&H" + Mid$(y, loopit, 1))
        zn = Val("&H" + Mid$(z, loopit, 1))
        tempstr = tempstr + Hex$(yn Xor (xn Or (15 Xor zn)))
    Next loopit

    w = MD5AB1F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Function MD5AB1Hash
' Input:
'   ByVal hashthis As String - String to be hashed
' Output: None
' Purpose: MD5 Hash function
' Remarks: None
'===============================================================================
Function MD5AB1Hash(ByVal hashthis As String) As String
    ReDim Buf(0 To 3) As String
    ReDim in_(0 To 15) As String
    Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
    Dim a As String, b As String, C As String, d As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If
    hashthis = hashthis + String$(tempnum2, Chr$(0))
    For loopit = 1 To 8
        hashthis = hashthis + Chr$(tempnum Mod 256)
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit
    
    ' Set magic numbers
    Buf(0) = "67452301"
    Buf(1) = "efcdab89"
    Buf(2) = "98badcfe"
    Buf(3) = "10325476"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = Buf(0)
        b = Buf(1)
        C = Buf(2)
        d = Buf(3)
    
        ' Get the 512 bits
        For loopit = 0 To 15
            in_(loopit) = ""
            For loopinner = 1 To 4
                in_(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_(loopit)
                If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" + in_(loopit)
            Next loopinner
        Next loopit
        
        ' Round 1
        MD5AB1F1 a, b, C, d, in_(0), "d76aa478", 7
        MD5AB1F1 d, a, b, C, in_(1), "e8c7b756", 12
        MD5AB1F1 C, d, a, b, in_(2), "242070db", 17
        MD5AB1F1 b, C, d, a, in_(3), "c1bdceee", 22
        MD5AB1F1 a, b, C, d, in_(4), "f57c0faf", 7
        MD5AB1F1 d, a, b, C, in_(5), "4787c62a", 12
        MD5AB1F1 C, d, a, b, in_(6), "a8304613", 17
        MD5AB1F1 b, C, d, a, in_(7), "fd469501", 22
        MD5AB1F1 a, b, C, d, in_(8), "698098d8", 7
        MD5AB1F1 d, a, b, C, in_(9), "8b44f7af", 12
        MD5AB1F1 C, d, a, b, in_(10), "ffff5bb1", 17
        MD5AB1F1 b, C, d, a, in_(11), "895cd7be", 22
        MD5AB1F1 a, b, C, d, in_(12), "6b901122", 7
        MD5AB1F1 d, a, b, C, in_(13), "fd987193", 12
        MD5AB1F1 C, d, a, b, in_(14), "a679438e", 17
        MD5AB1F1 b, C, d, a, in_(15), "49b40821", 22
        
        ' Round 2
        MD5AB1F2 a, b, C, d, in_(1), "f61e2562", 5
        MD5AB1F2 d, a, b, C, in_(6), "c040b340", 9
        MD5AB1F2 C, d, a, b, in_(11), "265e5a51", 14
        MD5AB1F2 b, C, d, a, in_(0), "e9b6c7aa", 20
        MD5AB1F2 a, b, C, d, in_(5), "d62f105d", 5
        MD5AB1F2 d, a, b, C, in_(10), "02441453", 9
        MD5AB1F2 C, d, a, b, in_(15), "d8a1e681", 14
        MD5AB1F2 b, C, d, a, in_(4), "e7d3fbc8", 20
        MD5AB1F2 a, b, C, d, in_(9), "21e1cde6", 5
        MD5AB1F2 d, a, b, C, in_(14), "c33707d6", 9
        MD5AB1F2 C, d, a, b, in_(3), "f4d50d87", 14
        MD5AB1F2 b, C, d, a, in_(8), "455a14ed", 20
        MD5AB1F2 a, b, C, d, in_(13), "a9e3e905", 5
        MD5AB1F2 d, a, b, C, in_(2), "fcefa3f8", 9
        MD5AB1F2 C, d, a, b, in_(7), "676f02d9", 14
        MD5AB1F2 b, C, d, a, in_(12), "8d2a4c8a", 20
        
        ' Round 3
        MD5AB1F3 a, b, C, d, in_(5), "fffa3942", 4
        MD5AB1F3 d, a, b, C, in_(8), "8771f681", 11
        MD5AB1F3 C, d, a, b, in_(11), "6d9d6122", 16
        MD5AB1F3 b, C, d, a, in_(14), "fde5380c", 23
        MD5AB1F3 a, b, C, d, in_(1), "a4beea44", 4
        MD5AB1F3 d, a, b, C, in_(4), "4bdecfa9", 11
        MD5AB1F3 C, d, a, b, in_(7), "f6bb4b60", 16
        MD5AB1F3 b, C, d, a, in_(10), "bebfbc70", 23
        MD5AB1F3 a, b, C, d, in_(13), "289b7ec6", 4
        MD5AB1F3 d, a, b, C, in_(0), "eaa127fa", 11
        MD5AB1F3 C, d, a, b, in_(3), "d4ef3085", 16
        MD5AB1F3 b, C, d, a, in_(6), "04881d05", 23
        MD5AB1F3 a, b, C, d, in_(9), "d9d4d039", 4
        MD5AB1F3 d, a, b, C, in_(12), "e6db99e5", 11
        MD5AB1F3 C, d, a, b, in_(15), "1fa27cf8", 16
        MD5AB1F3 b, C, d, a, in_(2), "c4ac5665", 23
        
        ' Round 4
        MD5AB1F4 a, b, C, d, in_(0), "f4292244", 6
        MD5AB1F4 d, a, b, C, in_(7), "432aff97", 10
        MD5AB1F4 C, d, a, b, in_(14), "ab9423a7", 15
        MD5AB1F4 b, C, d, a, in_(5), "fc93a039", 21
        MD5AB1F4 a, b, C, d, in_(12), "655b59c3", 6
        MD5AB1F4 d, a, b, C, in_(3), "8f0ccc92", 10
        MD5AB1F4 C, d, a, b, in_(10), "ffeff47d", 15
        MD5AB1F4 b, C, d, a, in_(1), "85845dd1", 21
        MD5AB1F4 a, b, C, d, in_(8), "6fa87e4f", 6
        MD5AB1F4 d, a, b, C, in_(15), "fe2ce6e0", 10
        MD5AB1F4 C, d, a, b, in_(6), "a3014314", 15
        MD5AB1F4 b, C, d, a, in_(13), "4e0811a1", 21
        MD5AB1F4 a, b, C, d, in_(4), "f7537e82", 6
        MD5AB1F4 d, a, b, C, in_(11), "bd3af235", 10
        MD5AB1F4 C, d, a, b, in_(2), "2ad7d2bb", 15
        MD5AB1F4 b, C, d, a, in_(9), "eb86d391", 21
    
        Buf(0) = BigAA1Add(Buf(0), a)
        Buf(1) = BigAA1Add(Buf(1), b)
        Buf(2) = BigAA1Add(Buf(2), C)
        Buf(3) = BigAA1Add(Buf(3), d)
    Next loopouter
    
    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 3
        For loopinner = 3 To 0 Step -1
            hashthis = hashthis + Mid$(Buf(loopit), 1 + 2 * loopinner, 2)
        Next loopinner
    Next loopit
    
    ' And return it
    MD5AB1Hash = hashthis

End Function

'===============================================================================
' Name: Function MD5AB2F
' Input:
'   ByVal tempstr As String
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: Variant
' Purpose: [INTERNAL] MD5 function
' Remarks: None
'===============================================================================
Function MD5AB2F(ByVal tempstr As String, ByVal w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim valueans As String, tempvalstr  As String
    Dim loopit As Integer, tempnum As Long
    Dim temps1 As String, temps2 As String, temps3 As String, temps4 As String

    w = String$(8 - Len(w), "0") + w
    tempstr = String$(8 - Len(tempstr), "0") + tempstr
    in_ = String$(8 - Len(in_), "0") + in_
    qdata = String$(8 - Len(qdata), "0") + qdata

    temps1 = Right$(w, 5)
    temps2 = Right$(tempstr, 5)
    temps3 = Right$(in_, 5)
    temps4 = Right$(qdata, 5)
    tempnum = Val("&H" + temps1 + "&") + Val("&H" + temps2 + "&") + Val("&H" + temps3 + "&") + Val("&H" + temps4 + "&")
    tempvalstr = Hex$(tempnum Mod 1048576)
    valueans = String$(5 - Len(tempvalstr), "0") + tempvalstr + valueans
    tempnum = Int(tempnum / 1048576)
    temps1 = Left$(w, 3)
    temps2 = Left$(tempstr, 3)
    temps3 = Left$(in_, 3)
    temps4 = Left$(qdata, 3)
    tempnum = tempnum + Val("&H" + temps1 + "&") + Val("&H" + temps2 + "&") + Val("&H" + temps3 + "&") + Val("&H" + temps4 + "&")
    valueans = Hex$(tempnum) + valueans

    MD5AB2F = BigAA2Mod32Add(BigAA2RotLeft(valueans, rots), X)
                   

End Function

'===============================================================================
' Name: Sub MD5AB2F1
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AB2F1(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim xn As Long, yn As Long, zn As Long
    
    X = String$(8 - Len(X), "0") + X
    y = String$(8 - Len(y), "0") + y
    z = String$(8 - Len(z), "0") + z

    xn = Val("&H" + Right$(X, 5) + "&")
    yn = Val("&H" + Right$(y, 5) + "&")
    zn = Val("&H" + Right$(z, 5) + "&")
    tempstr = Hex$(zn Xor (xn And (yn Xor zn)))
    tempstr = String$(5 - Len(tempstr), "0") + tempstr
    xn = Val("&H" + Left$(X, 3))
    yn = Val("&H" + Left$(y, 3))
    zn = Val("&H" + Left$(z, 3))
    tempstr = Hex$(zn Xor (xn And (yn Xor zn))) + tempstr
    
    w = MD5AB2F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AB2F2
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AB2F2(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim xn As Long, yn As Long, zn As Long
    
    X = String$(8 - Len(X), "0") + X
    y = String$(8 - Len(y), "0") + y
    z = String$(8 - Len(z), "0") + z

    xn = Val("&H" + Right$(X, 5) + "&")
    yn = Val("&H" + Right$(y, 5) + "&")
    zn = Val("&H" + Right$(z, 5) + "&")
    tempstr = Hex$(yn Xor (zn And (xn Xor yn)))
    tempstr = String$(5 - Len(tempstr), "0") + tempstr
    xn = Val("&H" + Left$(X, 3))
    yn = Val("&H" + Left$(y, 3))
    zn = Val("&H" + Left$(z, 3))
    tempstr = Hex$(yn Xor (zn And (xn Xor yn))) + tempstr
    
    w = MD5AB2F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AB2F3
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AB2F3(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim xn As Long, yn As Long, zn As Long
    
    X = String$(8 - Len(X), "0") + X
    y = String$(8 - Len(y), "0") + y
    z = String$(8 - Len(z), "0") + z

    xn = Val("&H" + Right$(X, 5) + "&")
    yn = Val("&H" + Right$(y, 5) + "&")
    zn = Val("&H" + Right$(z, 5) + "&")
    tempstr = Hex$(xn Xor (yn Xor zn))
    tempstr = String$(5 - Len(tempstr), "0") + tempstr
    xn = Val("&H" + Left$(X, 3))
    yn = Val("&H" + Left$(y, 3))
    zn = Val("&H" + Left$(z, 3))
    tempstr = Hex$(xn Xor (yn Xor zn)) + tempstr
    
    w = MD5AB2F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Sub MD5AB2F4
' Input:
'   ByVal w As String
'   ByVal X As String
'   ByVal Y As String
'   ByVal z As String
'   ByVal in_ As String
'   ByVal qdata As String
'   ByVal rots As Integer
' Output: None
' Purpose: [INTERNAL] MD5 sub
' Remarks: None
'===============================================================================
Sub MD5AB2F4(w As String, ByVal X As String, ByVal y As String, ByVal z As String, ByVal in_ As String, ByVal qdata As String, ByVal rots As Integer)
    Dim tempstr As String
    Dim xn As Long, yn As Long, zn As Long
    
    X = String$(8 - Len(X), "0") + X
    y = String$(8 - Len(y), "0") + y
    z = String$(8 - Len(z), "0") + z

    xn = Val("&H" + Right$(X, 5) + "&")
    yn = Val("&H" + Right$(y, 5) + "&")
    zn = Val("&H" + Right$(z, 5) + "&")
    tempstr = Hex$(yn Xor (xn Or (1048575 Xor zn)))
    tempstr = String$(5 - Len(tempstr), "0") + tempstr
    xn = Val("&H" + Left$(X, 3))
    yn = Val("&H" + Left$(y, 3))
    zn = Val("&H" + Left$(z, 3))
    tempstr = Hex$(yn Xor (xn Or (4095 Xor zn))) + tempstr
    
    w = MD5AB2F(tempstr, w, X, y, z, in_, qdata, rots)

End Sub

'===============================================================================
' Name: Function MD5AB2Hash
' Input:
'   ByVal hashthis As String - String to be hashed
' Output: None
' Purpose: MD5 Hash function
' Remarks: None
'===============================================================================
Function MD5AB2Hash(ByVal hashthis As String) As String
    ReDim Buf(0 To 3) As String
    ReDim in_(0 To 15) As String
    Dim tempnum As Integer, tempnum2 As Integer, loopit As Integer, loopouter As Integer, loopinner As Integer
    Dim a As String, b As String, C As String, d As String

    ' Add padding
    tempnum = 8 * Len(hashthis)
    hashthis = hashthis + Chr$(128) 'Add binary 10000000
    tempnum2 = 56 - Len(hashthis) Mod 64
    If tempnum2 < 0 Then
        tempnum2 = 64 + tempnum2
    End If
    hashthis = hashthis + String$(tempnum2, Chr$(0))
    For loopit = 1 To 8
        hashthis = hashthis + Chr$(tempnum Mod 256)
        tempnum = tempnum - tempnum Mod 256
        tempnum = tempnum / 256
    Next loopit
    
    ' Set magic numbers
    Buf(0) = "67452301"
    Buf(1) = "efcdab89"
    Buf(2) = "98badcfe"
    Buf(3) = "10325476"
    
    ' For each 512 bit section
    For loopouter = 0 To Len(hashthis) / 64 - 1
        a = Buf(0)
        b = Buf(1)
        C = Buf(2)
        d = Buf(3)
    
        ' Get the 512 bits
        For loopit = 0 To 15
            in_(loopit) = ""
            For loopinner = 1 To 4
                in_(loopit) = Hex$(Asc(Mid$(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) + in_(loopit)
                If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" + in_(loopit)
            Next loopinner
        Next loopit
        
        ' Round 1
        MD5AB2F1 a, b, C, d, in_(0), "d76aa478", 7
        MD5AB2F1 d, a, b, C, in_(1), "e8c7b756", 12
        MD5AB2F1 C, d, a, b, in_(2), "242070db", 17
        MD5AB2F1 b, C, d, a, in_(3), "c1bdceee", 22
        MD5AB2F1 a, b, C, d, in_(4), "f57c0faf", 7
        MD5AB2F1 d, a, b, C, in_(5), "4787c62a", 12
        MD5AB2F1 C, d, a, b, in_(6), "a8304613", 17
        MD5AB2F1 b, C, d, a, in_(7), "fd469501", 22
        MD5AB2F1 a, b, C, d, in_(8), "698098d8", 7
        MD5AB2F1 d, a, b, C, in_(9), "8b44f7af", 12
        MD5AB2F1 C, d, a, b, in_(10), "ffff5bb1", 17
        MD5AB2F1 b, C, d, a, in_(11), "895cd7be", 22
        MD5AB2F1 a, b, C, d, in_(12), "6b901122", 7
        MD5AB2F1 d, a, b, C, in_(13), "fd987193", 12
        MD5AB2F1 C, d, a, b, in_(14), "a679438e", 17
        MD5AB2F1 b, C, d, a, in_(15), "49b40821", 22
        
        ' Round 2
        MD5AB2F2 a, b, C, d, in_(1), "f61e2562", 5
        MD5AB2F2 d, a, b, C, in_(6), "c040b340", 9
        MD5AB2F2 C, d, a, b, in_(11), "265e5a51", 14
        MD5AB2F2 b, C, d, a, in_(0), "e9b6c7aa", 20
        MD5AB2F2 a, b, C, d, in_(5), "d62f105d", 5
        MD5AB2F2 d, a, b, C, in_(10), "02441453", 9
        MD5AB2F2 C, d, a, b, in_(15), "d8a1e681", 14
        MD5AB2F2 b, C, d, a, in_(4), "e7d3fbc8", 20
        MD5AB2F2 a, b, C, d, in_(9), "21e1cde6", 5
        MD5AB2F2 d, a, b, C, in_(14), "c33707d6", 9
        MD5AB2F2 C, d, a, b, in_(3), "f4d50d87", 14
        MD5AB2F2 b, C, d, a, in_(8), "455a14ed", 20
        MD5AB2F2 a, b, C, d, in_(13), "a9e3e905", 5
        MD5AB2F2 d, a, b, C, in_(2), "fcefa3f8", 9
        MD5AB2F2 C, d, a, b, in_(7), "676f02d9", 14
        MD5AB2F2 b, C, d, a, in_(12), "8d2a4c8a", 20
        
        ' Round 3
        MD5AB2F3 a, b, C, d, in_(5), "fffa3942", 4
        MD5AB2F3 d, a, b, C, in_(8), "8771f681", 11
        MD5AB2F3 C, d, a, b, in_(11), "6d9d6122", 16
        MD5AB2F3 b, C, d, a, in_(14), "fde5380c", 23
        MD5AB2F3 a, b, C, d, in_(1), "a4beea44", 4
        MD5AB2F3 d, a, b, C, in_(4), "4bdecfa9", 11
        MD5AB2F3 C, d, a, b, in_(7), "f6bb4b60", 16
        MD5AB2F3 b, C, d, a, in_(10), "bebfbc70", 23
        MD5AB2F3 a, b, C, d, in_(13), "289b7ec6", 4
        MD5AB2F3 d, a, b, C, in_(0), "eaa127fa", 11
        MD5AB2F3 C, d, a, b, in_(3), "d4ef3085", 16
        MD5AB2F3 b, C, d, a, in_(6), "04881d05", 23
        MD5AB2F3 a, b, C, d, in_(9), "d9d4d039", 4
        MD5AB2F3 d, a, b, C, in_(12), "e6db99e5", 11
        MD5AB2F3 C, d, a, b, in_(15), "1fa27cf8", 16
        MD5AB2F3 b, C, d, a, in_(2), "c4ac5665", 23
        
        ' Round 4
        MD5AB2F4 a, b, C, d, in_(0), "f4292244", 6
        MD5AB2F4 d, a, b, C, in_(7), "432aff97", 10
        MD5AB2F4 C, d, a, b, in_(14), "ab9423a7", 15
        MD5AB2F4 b, C, d, a, in_(5), "fc93a039", 21
        MD5AB2F4 a, b, C, d, in_(12), "655b59c3", 6
        MD5AB2F4 d, a, b, C, in_(3), "8f0ccc92", 10
        MD5AB2F4 C, d, a, b, in_(10), "ffeff47d", 15
        MD5AB2F4 b, C, d, a, in_(1), "85845dd1", 21
        MD5AB2F4 a, b, C, d, in_(8), "6fa87e4f", 6
        MD5AB2F4 d, a, b, C, in_(15), "fe2ce6e0", 10
        MD5AB2F4 C, d, a, b, in_(6), "a3014314", 15
        MD5AB2F4 b, C, d, a, in_(13), "4e0811a1", 21
        MD5AB2F4 a, b, C, d, in_(4), "f7537e82", 6
        MD5AB2F4 d, a, b, C, in_(11), "bd3af235", 10
        MD5AB2F4 C, d, a, b, in_(2), "2ad7d2bb", 15
        MD5AB2F4 b, C, d, a, in_(9), "eb86d391", 21
    
        Buf(0) = BigAA2Add(Buf(0), a)
        Buf(1) = BigAA2Add(Buf(1), b)
        Buf(2) = BigAA2Add(Buf(2), C)
        Buf(3) = BigAA2Add(Buf(3), d)
    Next loopouter
    
    ' Extract MD5Hash
    hashthis = ""
    For loopit = 0 To 3
        For loopinner = 3 To 0 Step -1
            hashthis = hashthis + Mid$(Buf(loopit), 1 + 2 * loopinner, 2)
        Next loopinner
    Next loopit
    
    ' And return it
    MD5AB2Hash = hashthis

End Function
