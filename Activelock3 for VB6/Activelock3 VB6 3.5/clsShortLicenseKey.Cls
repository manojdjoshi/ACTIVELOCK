VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShortLicenseKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===============================================================================
'
'   LicenseKey Class
'
'   Use to provide license key generation and validation. This class exposes an
'   abstract interface that can be used to implement licensing for all of your
'   commerical and shareware applications.  Keys can be cloaked with a bit
'   swapping technique, and with a private key.  Keys can also be tied to a
'   licensee.
'
'   Use the global conditional compile constants (IncludeCreate, IncludeValidate,
'   and IncludeCheck) to define which members are compiled into your project.
'   For instance, set IncludeCreate = 0 to exclude it from the client app.
'
'   Extra effort is made in the ValidateKey method to so that the entire key is
'   not held in memory at any time. Keep that in mind if you alter the source.
'
'   This implementation breaks a key into the following parts:
'
'       1111-2222-3333-4444-5555
'
'       1111 = Product code
'       2222 = Expiration date (days since 1-1-1970)
'       3333 = Caller definable word (16 bit value)
'       4444 = CRC for key validation
'       5555 = CRC for input validation
'
'   IMPORTANT: Key generators (no matter how good they are) will NOT thwart a
'   cracker! Alter the source code to meet your proprietary needs.
'
'===============================================================================
'
'   Author:             Monte Hansen [monte@killervb.com]
'   Dependencies:       None
'   Invitation:         There is an open invitation to comment on this code,
'                       report bugs or request revisions or enhancements.
'
'===============================================================================
'
'   ==  Copyright © 1999-2001 by Monte Hansen, All Rights Reserved Worldwide  ==
'
'   Monte Hansen  (The Author) grants a royalty-free right to use,  modify,  and
'   distribute this code  (The Code)  in compiled form,  provided that you agree
'   that The Author has no warranty,  obligations  or  liability  for  The Code.
'   You may distribute The Code among peers but may not sell it,  or  distribute
'   it on any electronic or physical media such  as  floppy  diskettes,  compact
'   disks, bulletin boards, web sites, and the like, without first obtaining The
'   Author's consent.
'
'   When distributing The Code among peers,  it is respectfully  requested  that
'   it be distributed as is,  but at no time shall it be distributed without the
'   copyright notice hereinabove.
'
'===============================================================================
Option Explicit

'===============================================================================
'   Constants
'===============================================================================
Private Const Module        As String = "clsShortLicenseKey"

'===============================================================================
'   Enums
'===============================================================================
Private Enum Segments       ' segments to the license key
    iProdCode = 0
    iExpire = 1
    iUserData = 2
    iCRC = 3
    iCRC2 = 4
End Enum

Private Enum MapFileChecksumErrors
    CHECKSUM_SUCCESS = 0
    CHECKSUM_OPEN_FAILURE = 1       ' Could not open the file.
    CHECKSUM_MAP_FAILURE = 2        ' Could not map the file.
    CHECKSUM_MAPVIEW_FAILURE = 3    ' Could not map a view of the file.
    CHECKSUM_UNICODE_FAILURE = 4    ' Could not convert the file name to Unicode.
End Enum

'===============================================================================
'   Types
'===============================================================================
' This structure is used to store a
' reference to two bits that will be
' swapped. Each bit can be from a
' different segment in the key.
' iCRC2 cannot be swapped since it
' is a checksum of the first 4
' segments of the key.
Private Type TBits
    iWord1                  As Byte
    iBit1                   As Byte
    iWord2                  As Byte
    iBit2                   As Byte
End Type

'===============================================================================
'   Private Members
'===============================================================================
Private m_Bits()            As TBits
Private m_nSwaps            As Long

'===============================================================================
'   Declares
'===============================================================================
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSource As Long, ByVal nBytes As Long)
Private Declare Function MapFileAndCheckSumA Lib "IMAGEHLP.DLL" (ByVal FileName As String, HeaderSum As Long, CheckSum As Long) As Long

' Note: add a project conditional compile argument "IncludeCreate"
' if the CreateKey is to be compiled into the application.
'#If IncludeCreate = 1 Then
Friend Function CreateKey(ByVal PrivateKey As String, _
                          ByVal Licensee As String, _
                          ByVal ProductCode As Long, _
                          ByVal ExpireDate As Date, _
                          ByVal UserData As Integer) As String
'===============================================================================
'   CreateKey - Creates a new serial number.
'
'   PrivateKey      This key should be kept private and would probably be
'                   unique to a specific product. It can be any string you
'                   define and is needed for key validation.
'   Licensee        Name of party to whom this license is issued.
'   ProductCode     A unique number assigned to this product. This is intended
'                   to identify a serial number at a glance. If, however, you
'                   expect to issue multiple serial numbers to the same
'                   Licensee then this field could be used to further qualify.
'   ExpireDate      Use this field for time-based serial numbers. This allows
'                   serial number to be issued that expire in two weeks or at
'                   the end of the year.
'   UserData        This field is caller defined. One might use it to identify
'                   the # of licensed CPU's, or perhaps the # of servers
'                   licensed, etc, or both (using a LoByte/HiByte packed field).
'
'   RETURNS         A serial number in the form of "233C-3912-00FF-BE49"
'===============================================================================

    Dim KeySegs(4)  As String
    Dim i           As Long

    ' convert each segment value to a hex string
    KeySegs(iProdCode) = HexWORD(ProductCode)
    KeySegs(iExpire) = HexWORD(DateDiff("d", _
        DateValue(#1/1/1970#), DateValue(CDate(Format(ExpireDate, "mm/dd/yyyy")))))
    KeySegs(iUserData) = HexWORD(UserData)

    ' Compute CRC against pertinent info.
    KeySegs(iCRC) = HexWORD(CRC(UCase$( _
        Licensee & _
        KeySegs(iProdCode) & _
        KeySegs(iExpire) & _
        KeySegs(iUserData) & _
        PrivateKey)))

    ' Perform bit swaps
    For i = 0 To m_nSwaps - 1
        SwapBit m_Bits(i), KeySegs
    Next i
    
    ' Calculate the CRC used to perform
    ' simple user input validation.
    KeySegs(iCRC2) = HexWORD(CRC(UCase$( _
        Licensee & _
        KeySegs(iProdCode) & _
        KeySegs(iExpire) & _
        KeySegs(iUserData) & _
        KeySegs(iCRC))))
    
    ' Return the key to the caller
    CreateKey = UCase$(KeySegs(iProdCode) & "-" & _
        KeySegs(iExpire) & "-" & _
        KeySegs(iUserData) & "-" & _
        KeySegs(iCRC) & "-" & _
        KeySegs(iCRC2))

End Function
'#End If

'#If IncludeCheck = 1 Then
Friend Function TestKey(ByVal LicenseKey As String, ByVal Licensee As String) As Boolean
'===============================================================================
'   TestKey - Performs a simple CRC test to ensure the key was entered
'   "correctly". Does NOT validate that the key is VALID. This function allows
'   the caller to "test" the key input by the user, without having to execute
'   the key validation code, making it more work for a cracker to generate a
'   for your application.
'===============================================================================

    Dim KeySegs         As Variant
    Dim nCRC            As Long
    
    On Error GoTo ExitLabel
    
    ' TODO: don't even call this function if SoftIce was detected!
    
    If Not SplitKey(LicenseKey, KeySegs) Then GoTo ExitLabel
    
    ' NOTE: Licensee can be omitted from the last checksum
    ' if there is no need to bind a serial number to a
    ' customer name.
    nCRC = CRC(UCase$(Licensee & _
        KeySegs(iProdCode) & _
        KeySegs(iExpire) & _
        KeySegs(iUserData) & _
        KeySegs(iCRC)))

    ' Compare check digits
    TestKey = (nCRC = SegmentValue(KeySegs(iCRC2)))
    
ExitLabel:
    
End Function
'#End If

'#If IncludeValidate = 1 Then
Friend Function ValidateKey(ByVal LicenseKey As String, _
                            ByVal PrivateKey As String, _
                            ByVal Licensee As String, _
                            ByVal ProductCode As Long, _
                            Optional ExpireDate As Date, _
                            Optional UserData As Integer) As Boolean
'===============================================================================
'   ValidateKey - Evaluates the supplied license key and tests that it is
'   valid. We do this by recomputing the checksum and comparing it to the one
'   embedded in the serial number.
'
'   LicenseKey      The serial number to validate.
'                   serial number.
'   PrivateKey      A magic string that is application specific. This should
'                   be the same as was used to create the serial number.
'   Licensee        Name of party to whom this license is issued. This should
'                   be the same as was used to create the serial number.
'   ProductCode     A unique number assigned to this product. This should
'                   be the same as was used to create the serial number.
'   ExpireDate      Use this field for time-based serial numbers. This should
'                   be the same as was used to create the serial number.
'   UserData        This field is caller defined. This should be the same as
'                   was used to create the serial number.
'
'   RETURNS         True if the serial number checks out, False otherwise.
'
'===============================================================================

'*******************************************************************************
'   IMPORTANT       This function is where the most care must be given.
'                   You should assume that a cracker has seen this code and can
'                   recognize it from ASM listings, and should be changed.
'                   - Avoid string compares whenever possible.
'                   - Pepper lots of JUNK code.
'                   - Do things in different order (except CRC checks).
'                   - Do not do things in this routine that are being monitored
'                     (registry calls, file-system access, phone home w/TCP).
'                   - Remove the UCase$ statements (just pass serial in ucase).
'*******************************************************************************

    Dim KeySegs         As Variant
    Dim nCrc1           As Long
    Dim nCrc2           As Long
    Dim nCrc3           As Long
    Dim nCrc4           As Long
    Dim i               As Long
    
    On Error GoTo ExitLabel
    
    ' TODO: don't even call this function if SoftIce was detected!
    
    ' ----------------------------------------------------------
    ' This section of code could raise red flags
    ' ----------------------------------------------------------
    
    If Not SplitKey(LicenseKey, KeySegs) Then GoTo ExitLabel

    ' ----------------------------------------------------------

    ' TODO: UCase string before it get's here
    
    ' Get CRC used for input validation
    nCrc1 = CRC(UCase$(Licensee) & _
        KeySegs(iProdCode) & _
        KeySegs(iExpire) & _
        KeySegs(iUserData) & _
        KeySegs(iCRC))
    
    ' Compare check digits
    If (nCrc1 <> SegmentValue(KeySegs(iCRC2))) Then
        GoTo ExitLabel
    End If

    ' Perform bit swaps (in reverse).
    For i = m_nSwaps - 1 To 0 Step -1
        SwapBit m_Bits(i), KeySegs
    Next i
    
    ' Calculate checksum on the license KeySegs.
    ' The LAST thing we want to do is to push a valid
    ' serial number on to the stack. This is the first
    ' thing a cracker will look for. Instead we will
    ' calculate a running checksum on each segment and
    ' compare the checksum to the checksum embedded in
    ' the key.
    
    ' The supplied product code should be
    ' the same as the product code embedded
    ' in the key.
    If ProductCode = SegmentValue(KeySegs(iProdCode)) Then
        ' One more check on the check digits before we
        ' blow away the value stored in nCrc1.
        If (SegmentValue(KeySegs(iCRC2)) = nCrc1) Then
            nCrc1 = CRC(UCase$(Licensee))
        End If
    End If
    
    nCrc2 = CRC(UCase$(KeySegs(iProdCode)), nCrc1)
    nCrc3 = CRC(UCase$(KeySegs(iExpire)), nCrc2)
    nCrc3 = CRC(UCase$(KeySegs(iUserData)), nCrc3)
    nCrc4 = CRC(UCase$(PrivateKey), nCrc3)

    ' Return success and fill outputs IF the license KeySegs is valid
    If nCrc4 = SegmentValue(KeySegs(iCRC)) Then
    
        ' Fill the outputs with expire date and user data.
        ExpireDate = DateAdd("d", SegmentValue(KeySegs(iExpire)), #1/1/1970#)
        UserData = SegmentValue(KeySegs(iUserData))
   
        ' IMPORTANT: This is an easy patch point
        ' if you use real-time key validation.
        ValidateKey = True
    
    End If

ExitLabel:

End Function
'#End If

'#If IncludeValidate = 1 Or IncludeCreate = 1 Then
Friend Sub AddSwapBits(ByVal Word1 As Long, ByVal Bit1 As Long, _
                       ByVal Word2 As Long, ByVal Bit2 As Long)
'===============================================================================
'   AddSwapBits - This is used to swap various bits in the serial number. It's
'   sole purpose is to alter the output serial number.
'
'   This process is "played" forwards during the key creation, and in reverse
'   when validating. This mangling process should be identical for key creation
'   and validation. Add as many combinations as you like.
'
'   Word1/Word2     The words to bit swap. There are 4 words in the serial #.
'                   This parameter is zero-based.
'   Bit1/Bit2       The bits to swap. There are 16 bits to each word.
'                   This parameter is zero-based.
'
'   Example: This scenario causes word 3, bit 8 to be swapped with word 1, bit 3
'
'       KeyGen.AddSwapBits 1, 3, 3, 8
'
'   NOTE:   It is recommended that there be at least 6 combinations in case
'   the bits being swapped are the same (2 swap bits for words 2, 3 & 4).
'===============================================================================

    ' TODO: don't even call this function if SoftIce was detected!

    ' Size array to fit
    If m_nSwaps = 0 Then
        ReDim m_Bits(m_nSwaps)
    Else
        ReDim Preserve m_Bits(m_nSwaps)
    End If
    m_nSwaps = m_nSwaps + 1

    ' This implementation hardcodes keys that are 8 bytes/4 words
    If Word1 < 0 Or Word1 > 3 _
    Or Word2 < 0 Or Word2 > 3 Then
        Err.Raise 5, Module, "Word specification is not within 0-3."
    End If

    ' There are only 16 bits to a word.
    If Bit1 < 0 Or Bit1 > 15 _
    Or Bit2 < 0 Or Bit2 > 15 Then
        Err.Raise 5, Module, "Bit specification is not within 0-15."
    End If

    ' Save the bits to be swapped
    With m_Bits(m_nSwaps - 1)
        .iWord1 = Word1
        .iBit1 = Bit1
        .iWord2 = Word2
        .iBit2 = Bit2
    End With

End Sub
'#End If

'#If IncludeValidate = 1 Or IncludeCheck = 1 Then
Private Function SplitKey(LicenseKey As String, KeySegs As Variant) As Boolean
'===============================================================================
'   SplitKey - Shared code to massage the input serial number, and slice it into
'   the required number of segments.
'===============================================================================

    ' ----------------------------------------------------------
    ' This section of code could raise red flags
    ' ----------------------------------------------------------
    
    ' Sanity check
    If InStr(LicenseKey, "-") = 0 Then GoTo ExitLabel

    ' As a courtesy to the user, we convert the
    ' letter "O" to the number "0". Users hate
    ' serialz that do not have interchangable 0/o's!
    LicenseKey = VBA.Replace(LicenseKey, "o", "0", _
        Compare:=vbTextCompare)
    
    ' Splice the KeySegs into 4 segments,
    ' exit if wrong # of segments.
    KeySegs = VBA.Split(UCase(LicenseKey), "-", 5)
    If UBound(KeySegs) <> 4 Then GoTo ExitLabel

    ' ----------------------------------------------------------

    SplitKey = True
    
ExitLabel:

End Function
'#End If


Private Function SegmentValue(ByVal HexString As String) As Long
'===============================================================================
'   Converts a hex string representation into a 4 byte decimal value.
'===============================================================================

    Dim Buffer(3)   As Byte
    Dim i           As Long
    Dim j           As Long

    ' Exit if each byte not represented by a 2 character string
    If Len(HexString) Mod 2 <> 0 Then Exit Function

    ' Exit if it's larger than a 4 byte value
    If Len(HexString) > 8 Then Exit Function

    ' NOTE: we populate the byte array in little-endian format
    For i = Len(HexString) To 1 Step -2
        Buffer(j) = CByte("&H" & Mid$(HexString, i - 1, 2))
        j = j + 1
    Next i

    ' Return the value
    CopyMemory ByVal VarPtr(SegmentValue), ByVal VarPtr(Buffer(0)), 4

End Function


'#If IncludeCreate = 1 Or IncludeValidate = 1 Then
Private Sub SwapBit(BitList As TBits, KeySegs)
'===============================================================================
'   SwapBit - Swaps any two bits. The bits can differ as long as they are in
'   the range of 0 and 15.
'===============================================================================
    With BitList
        ' Essentially, we swap Bit1 with Bit2. We use a bitwise
        ' OR operator or a bitwise AND operator depending
        ' upon if the subject bit is present. We don't use
        ' local variables to avoid synchronizing, especially
        ' since we may be doing a bit swap on the same word.
        If (SegmentValue(KeySegs(.iWord1)) _
        And (2 ^ .iBit2)) = (2 ^ .iBit2) Then
            KeySegs(.iWord1) = HexWORD(SegmentValue _
                (KeySegs(.iWord1)) Or (2 ^ .iBit2), "")
        Else
            KeySegs(.iWord1) = HexWORD(SegmentValue _
                (KeySegs(.iWord1)) And Not (2 ^ .iBit2), "")
        End If
        
        If (SegmentValue(KeySegs(.iWord2)) _
        And (2 ^ .iBit1)) = (2 ^ .iBit1) Then
            KeySegs(.iWord2) = HexWORD(SegmentValue _
                (KeySegs(.iWord2)) Or (2 ^ .iBit1), "")
        Else
            KeySegs(.iWord2) = HexWORD(SegmentValue _
                (KeySegs(.iWord2)) And Not (2 ^ .iBit1), "")
        End If

    End With

End Sub
'#End If

' Generic helper function
Private Function CRC(Buffer() As Byte, Optional InputCrc As Long) As Long
'===============================================================================
'   Crc - Returns a 16-bit CRC value for a data block.
'
'   Refer to CRC-CCITT compute-on-the-fly implementatations for more info.
'===============================================================================

    Dim Bit             As Long
    Dim i               As Long
    Dim j               As Long

    On Error GoTo ErrHandler

    ' Derive from a prior CRC value if supplied.
    CRC = InputCrc
    
    ' Loop thru entire buffer computing the CRC
    For i = LBound(Buffer) To UBound(Buffer)

        ' Loop thru each of the 8 bits
        For j = 0 To 7
        
            Bit = ((CRC And &H8000&) = &H8000&) _
                Xor ((Buffer(i) And (2 ^ j)) = 2 ^ j)

            CRC = ((CRC And &H7FFF&) * 2)
            
            If Bit <> 0 Then
                CRC = CRC Xor &H1021&
            End If
      
        Next j

    Next i

    Exit Function
    Resume
ErrHandler:
    Debug.Assert 0

End Function

Private Function HexWORD(ByVal WORD As Long, _
                         Optional ByVal Prefix As String) As String
'===============================================================================
'   HexDWORD - Returns a hex string representation of a WORD.
'
'   WORD            The 2 byte value to convert to a hex string.
'   Prefix          A value such as "0x" or "&H".
'
'   NOTE:  It's up to the caller to ensure the subject value is a 16-bit number.
'===============================================================================

    Dim bytes(1 To 2)   As Byte
    Dim i               As Long
    
    CopyMemory ByVal VarPtr(bytes(1)), ByVal VarPtr(WORD), 2

    HexWORD = Prefix
    For i = UBound(bytes) To LBound(bytes) Step -1
        If Len(Hex$(bytes(i))) = 1 Then
            HexWORD = HexWORD & "0" & LCase$(Hex$(bytes(i)))
        Else
            HexWORD = HexWORD & LCase$(Hex$(bytes(i)))
        End If
    Next i
    
End Function

Friend Function ExeIsPatched(ByVal FilePath As String) As Boolean
'===============================================================================
'   ExeIsPatched - Tests if the supplied file has been altered by computing a
'   checksum for the file and comparing it against the checksum in the
'   executable image.
'
'   FileName - Full path to file to check. Caller is responsible for ensuring
'   that the path exists, and that it is an executable.
'===============================================================================

    Dim FileCRC     As Long
    Dim HdrCRC      As Long
    Dim ErrorCode   As Long

    ' NOTE: Many crackers today are smart enough to
    ' update the PE image CRC value. But we check
    ' anyhow, just in case. Otherwise, it could be
    ' embarrassing if the EXE was patched without
    ' updating the PE header.

    ErrorCode = MapFileAndCheckSumA(FilePath, HdrCRC, FileCRC)
    If ErrorCode = CHECKSUM_SUCCESS Then

        If HdrCRC <> 0 _
        And HdrCRC <> FileCRC Then

            ' CRC of file is different than the CRC
            ' embedded in the PE image. Try not to
            ' let the cracker know that you are on
            ' to him. And don't start deleting from
            ' their harddrive!
            ExeIsPatched = True

        End If

    End If

End Function
