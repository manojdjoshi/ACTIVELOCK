VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCryptoAPIBase64"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'******************************************************************************
'*
'* Class clsCryptoAPIBase64
'*
'* This class implements the common BASE64 encoding and decoding. It also facilitates
'* formatting of encoded strings and cleaning of encoded strings from characters not
'* in the BASE64 coding alphabet. Opposite to some RFC-recommendations, the decoding
'* routine does not ignore invalid characters in the string to be decoded, but
'* raises an error instead. Use cleaning if you expect invalid characters in the string
'* due to previous formatting. Due to its special 'bit shuffling' algorithm the
'* implementation is very fast (especially when decoding).
'*
'* Note: if the encoded string is send "over the wire" and as a result from transmission
'*       distortions its contents changed, the only correct means of ensuring data
'*       integrity is to include some kind of checksum into the transmission. The absence
'*       of invalid characters in the string does by no means guarantee correct transmission.
'*       On the other side, if there are invalid characters in the string, which have not
'*       been put into the string by previous formatting, this is a strong hint on
'*       transmission distortions.
'*
'*
'* (c) 2004 Ulrich Korndörfer proSource software development
'*          www.prosource.de
'*          German site with VB articles (in english) and code (comments in english)
'*
'* Precautions: None. May be compiled to native code (which is strongly recommended),
'*              with all extended options selected.
'*
'* External dependencies: - module MCommon
'*                        - some APIs (see below)
'*
'* Version history
'*
'* Version 1.0.1 from 2004.09.24 Uses CompyMem4 now
'* Version 1.0   from 2004.04.01
'*
'*  Disclaimer:
'*  All code in this class is for demonstration purposes only.
'*  It may be used in a production environment, as it is thoroughly tested,
'*  but do not hold us responsible for any consequences resulting of using it.
'*
'******************************************************************************


Private Const mcClassName As String = "clsCryptoAPIBase64"


'******************************************************************************
'* API declarations
'******************************************************************************

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" _
                   (ByRef Destination As Any, _
                    ByRef Source As Any, _
                    ByVal Length As Long)
Private Declare Sub CopyMem4 Lib "msvbvm60.dll" Alias "GetMem4" _
                   (ByRef FromAddr As Any, _
                    ByRef ToAddr As Any)


'******************************************************************************
'* Private consts, member vars and types
'******************************************************************************

Private Const mclsCryptoAPIBase64_CHARSET As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
Private Const mclsCryptoAPIBase64_PADDINGBYTE As Byte = 61 'Chr$(61) = "="
Private Const mcSHIFT6 As Long = 64
Private Const mcSHIFT12 As Long = mcSHIFT6 * mcSHIFT6
Private Const mcSHIFT18 As Long = mcSHIFT12 * mcSHIFT6

'Static arrays wrapped in an UDT are used for performance reasons
'(adressing of array elements is faster).

Private Type TEncodeTable
  Enc(0 To 63) As Byte
End Type

Private Type TDecodeTable
  Dec(0 To 255) As Long
End Type

Private Type TLongByte
  B(0 To 3) As Byte
End Type

Private mEncodeTable As TEncodeTable
Private mDecodeTable As TDecodeTable


'******************************************************************************
'* Public methods
'******************************************************************************

'Encode encodes the byte array Source() to a string using the BASE64 encoding.
'Input param Source() is not altered

'Checks done on the input array:

'- Source must be an initialized, non zombie, one dimensional, zero based array.
'  Source may be a static or dynamic array

'Encode: 8 Bit triplet S(i+1),S(i+1),S(i+2) -> 6 Bit quartet O(j+0),O(j+1),O(j+2),O(j+3)

'   S(i+0)        S(i+1)        S(i+2)
'   |    |        |    |        |    |
'   |    +---+    |    +---+    |    |
' H6|      L2|  H4|      L4|  H2|  L6|
'   |        |    |        |    |    |
'   O(j+0)   O(j+1)        O(j+2)    O(j+3)

'   S(i+0)        S(i+1)        0
'   |    |        |    |        |
'   |    +---+    |    +---+    |
' H6|      L2|  H4|      L4|  02|
'   |        |    |        |    |
'   O(j+0)   O(j+1)        O(j+2)    0

'   S(i+0)        0             0
'   |    |        |
'   |    +---+    |
' H6|      L2|  04|
'   |        |    |
'   O(j+0)   O(j+1)        0         0

Public Function Encode(ByRef Source() As Byte) As String
Const cMETHODNAME As String = "Encode"
Dim Src() As Byte, SrcTop As Long, Out() As Byte, OutTop As Long, Pad As Long
Dim AccuL As Long, AccuB As TLongByte
Dim i As Long, j As Long

'Check array

gSimpleCheckArray Source, 0, 0, mcClassName, cMETHODNAME, "Invalid array"

'Prepare helper arrays

Src = Source: i = UBound(Source)
SrcTop = (i \ 3&) * 3& + 2& 'results to n*3 + 2, where n=0,1,2,...
Pad = SrcTop - i: If Pad > 0& Then ReDim Preserve Src(0& To SrcTop)
OutTop = (i \ 3&) * 4& + 3& 'results to n*4 + 3, where n is from above
ReDim Out(0& To OutTop)

'Encode Src() to Out()

'AccuL = | B(3) | B(2)     | B(1)     | B(0)     |
'AccuL = | 0    | Src(i+0) | Src(i+1) | Src(i+2) |

For i = 0& To SrcTop Step 3&
  With AccuB
    .B(0) = Src(i + 2&)
    .B(1) = Src(i + 1&)
    .B(2) = Src(i)
  End With
  CopyMem4 AccuB, AccuL
  With mEncodeTable
         Out(j) = .Enc((AccuL \ mcSHIFT18))         'H6[ Src(i+0) ]
    Out(j + 1&) = .Enc((AccuL \ mcSHIFT12) And 63&) 'L2[ Src(i+0) ] | H4[ Src(i+1) ]
    Out(j + 2&) = .Enc((AccuL \ mcSHIFT6) And 63&)  'L4[ Src(i+1) ] | H2[ Src(i+2) ]
    Out(j + 3&) = .Enc(AccuL And 63&)               'L6[ Src(i+2) ]
  End With
  j = j + 4&
Next i

'Do padding

If Pad = 2& Then Out(OutTop - 1&) = mclsCryptoAPIBase64_PADDINGBYTE
If Pad > 0& Then Out(OutTop) = mclsCryptoAPIBase64_PADDINGBYTE

'Return encoded

Encode = StrConv(Out, vbUnicode)

End Function

'Decode decodes a BASE64 encoded string to an one dimensional, zero based byte array.
'Input param Source is not altered.

'Checks done on the to be decoded input string:

'- length must be a multiple of 4, not zero
'- last two chars must be a legal combination of padding chars
'- illegal characters in the rest of the string. Illegal is any character, which is not in
'  the BASE64 code table. This excludes the padding chararacter too!

'Decode: 6 Bit quartet S(i+0),S(i+1),S(i+2),S(i+3) -> 8 Bit triplet O(j+0),O(j+1),O(j+2)

'   S(i+0)   S(i+1)        S(i+2)    S(i+3)
'   |        |    |        |    |    |
' L6|      H2|  L4|      H4|  L2|  L6|
'   |    +---+    |    +---+    |    |
'   |    |        |    |        |    |
'   O(j+0)        O(j+1)        O(j+2)

'   S(i+0)   S(i+1)        S(i+2)    0
'   |        |    |        |    |
' L6|      H2|  L4|      H4|  02|
'   |    +---+    |    +---+    |
'   |    |        |    |        |
'   O(j+0)        O(j+1)        0

'   S(i+0)   S(i+1)        0         0
'   |        |    |
' L6|      H2|  04|
'   |    +---+    |
'   |    |        |
'   O(j+0)        0             0

Public Function Decode(ByRef Source As String) As Byte()
Const cMETHODNAME As String = "Decode"
Dim Src() As Byte, SrcTop As Long, Out() As Byte, Pad As Long
Dim AccuL As Long, AccuB As TLongByte
Dim i As Long, j As Long

'Check length

i = Len(Source)
If i = 0 Or (i Mod 4) <> 0 Then ErrRaise cMETHODNAME, "Invalid length"

'Prepare helper arrays

Src = StrConv(Source, vbFromUnicode)
SrcTop = i - 1
ReDim Out(0& To ((i) \ 4&) * 3& - 1&)

'Check padding

If Src(SrcTop) = mclsCryptoAPIBase64_PADDINGBYTE Then
  Pad = 1&: Src(SrcTop) = mEncodeTable.Enc(0)
End If
If Src(SrcTop - 1&) = mclsCryptoAPIBase64_PADDINGBYTE Then
  If Pad = 0 Then
    ErrRaise cMETHODNAME, "Invalid padding"
  Else
    Pad = 2&: Src(SrcTop - 1&) = mEncodeTable.Enc(0)
  End If
End If

'Decode Src() to Out(). Check on invalid characters is included in the decoding algorithm.

'D() = Dec(Src()), having its low *6* bits carrying information, others are 0
'H4[D()] reads: the 4 high bits of D()'s low 6 bits
'AccuL = | 0    | H6[D(i+0)] H2[D(i+1)] | L4[D(i+1)] H4[D(i+2)] | L2[D(i+2)] H6[D(i+3)] |
'AccuL = | B(3) | B(2)                  | B(1)                  | B(0)                  |

For i = 0& To SrcTop Step 4&
  With mDecodeTable
    AccuL = (.Dec(Src(i)) * mcSHIFT18) Or _
            (.Dec(Src(i + 1&)) * mcSHIFT12) Or _
            (.Dec(Src(i + 2&)) * mcSHIFT6) Or _
             .Dec(Src(i + 3&))
  End With
  CopyMem4 AccuL, AccuB
  With AccuB
    If .B(3) = 0 Then
      Out(j + 2&) = .B(0)
      Out(j + 1&) = .B(1)
           Out(j) = .B(2)
    Else
      ErrRaise cMETHODNAME, "Invalid character found"
    End If
  End With
  j = j + 3&
Next i

'Adjust decoded array length according to padding

If Pad > 0 Then ReDim Preserve Out(0 To UBound(Out) - Pad)

Decode = Out

End Function

'Use SplitEncoded to split an encoded string into lines of length LineLen separated by vbCrLfs.
'The last line always has a vbCrLf appended.
'Input param Source is not altered

'Checks:

'- input string must not be empty
'- line length must be greater 0

Public Function SplitEncoded(ByRef Source As String, _
                             Optional ByVal LineLen As Long = 64&) As String
Const cMETHODNAME As String = "SplitEncoded"
Dim Lines() As String, LineCount As Long, SrcLen As Long, i As Long

SrcLen = Len(Source)
If SrcLen = 0& Then ErrRaise cMETHODNAME, "Empty source"
If LineLen <= 0& Then ErrRaise cMETHODNAME, "Invalid line length"

LineCount = (SrcLen - 1&) \ LineLen + 1&
ReDim Lines(0 To LineCount)

For i = 0& To LineCount - 1&
  Lines(i) = Mid$(Source, i * LineLen + 1&, LineLen)
Next i

SplitEncoded = Join$(Lines, vbCrLf)

End Function

'Use CleanEncoded to clean an encoded string before decoding from all invalid characters.
'(like white space). CleanEncoded stops reading from the input string either
'at the end of the input string or after the first occurrence of one or two
'consecutive padding bytes
'Input param Source is not altered

'Checks:

'- input string must not be empty
'- after cleaning the remaining string must not be empty and
'  must have a length being a multiple of 4

Public Function CleanEncoded(ByRef Source As String) As String
Const cMETHODNAME As String = "CleanEncoded"
Dim Src() As Byte, SrcTop As Long, Count As Long
Dim i As Long

Src = StrConv(Source, vbFromUnicode)
SrcTop = UBound(Src)
If SrcTop = -1& Then ErrRaise cMETHODNAME, "Empty source"

With mDecodeTable
  For i = 0 To SrcTop
    Select Case .Dec(Src(i))
      Case -1&
      Case -2&
        Src(Count) = Src(i): Count = Count + 1&
        If i < SrcTop Then
          If Src(i + 1&) = mclsCryptoAPIBase64_PADDINGBYTE Then
            Src(Count) = mclsCryptoAPIBase64_PADDINGBYTE: Count = Count + 1&
          End If
          Exit For
        End If
      Case Else
        Src(Count) = Src(i): Count = Count + 1&
    End Select
  Next i
End With

If Count = 0& Then ErrRaise cMETHODNAME, "All chars invalid"
If (Count Mod 4&) <> 0& Then ErrRaise cMETHODNAME, "Invalid length of cleaned"
ReDim Preserve Src(0& To Count - 1&)

CleanEncoded = StrConv(Src, vbUnicode)
End Function


'******************************************************************************
'* Private helpers
'******************************************************************************

Private Sub ErrRaise(ByVal MethodName As String, Optional ByVal Description As String)
Err.Raise 5, mcClassName & "." & MethodName, Description
End Sub


'******************************************************************************
'* Class de/construction
'******************************************************************************

Private Sub Class_Initialize()
Dim i As Long, Enc() As Byte

Enc = StrConv(mclsCryptoAPIBase64_CHARSET, vbFromUnicode)
RtlMoveMemory mEncodeTable.Enc(0), Enc(0), 64
With mDecodeTable
  For i = 0 To 255: .Dec(i) = -1: Next
  For i = 0 To 63: .Dec(Enc(i)) = i: Next
  .Dec(mclsCryptoAPIBase64_PADDINGBYTE) = -2
End With

End Sub
