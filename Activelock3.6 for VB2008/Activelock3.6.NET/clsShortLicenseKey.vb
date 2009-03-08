Option Strict Off
Option Explicit On

''' <summary>
''' <para>Use to provide license key generation and validation. This class exposes an
''' abstract interface that can be used to implement licensing for all of your
''' commerical and shareware applications.  Keys can be cloaked with a bit
''' swapping technique, and with a private key.  Keys can also be tied to a
''' licensee.</para>
''' </summary>
''' <remarks>See code, within, for more info!</remarks>
Friend Class clsShortLicenseKey

#Region "More Information & Copyright"
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
#End Region

    '===============================================================================
    '   Constants
    '===============================================================================

    ''' <summary>?Not Documented!</summary>
    ''' <remarks>UPGRADE_NOTE: Module was upgraded to Module_Renamed. Click for more: 'ms-help://MS.VSCC.2003/commoner/redir/redirect.htm?keyword="vbup1061"'</remarks>
    Private Const Module_Renamed As String = "clsShortLicenseKey"

#Region "Enums"

    ''' <summary>
    ''' segments to the license key
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum Segments ' segments to the license key
        iProdCode = 0
        iExpire = 1
        iUserData = 2
        iCRC = 3
        iCRC2 = 4
    End Enum

    ''' <summary>
    ''' Undocumented!
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum MapFileChecksumErrors
        CHECKSUM_SUCCESS = 0
        ''' <summary>
        ''' Could not open the file.
        ''' </summary>
        ''' <remarks></remarks>
        CHECKSUM_OPEN_FAILURE = 1
        ''' <summary>
        ''' Could not map the file.
        ''' </summary>
        ''' <remarks></remarks>
        CHECKSUM_MAP_FAILURE = 2
        ''' <summary>
        ''' Could not map a view of the file.
        ''' </summary>
        ''' <remarks></remarks>
        CHECKSUM_MAPVIEW_FAILURE = 3
        ''' <summary>
        ''' Could not convert the file name to Unicode.
        ''' </summary>
        ''' <remarks></remarks>
        CHECKSUM_UNICODE_FAILURE = 4
    End Enum

#End Region

    '===============================================================================
    '   Types
    '===============================================================================
    ''' <summary>This structure is used to store a reference to two bits that will be
    '''  swapped. Each bit can be from a different segment in the key.  iCRC2 cannot
    '''  be swapped since it is a checksum of the first 4 segments of the key.</summary>
    ''' <remarks></remarks>
    Private Structure TBits
        Dim iWord1 As Byte
        Dim iBit1 As Byte
        Dim iWord2 As Byte
        Dim iBit2 As Byte
    End Structure

    '===============================================================================
    '   Private Members
    '===============================================================================
    ''' <summary>?Not Documented!</summary>
    ''' <remarks></remarks>
    Private m_Bits() As TBits
    ''' <summary>?Not Documented!</summary>
    ''' <remarks></remarks>
    Private m_nSwaps As Integer

    '===============================================================================
    '   Declares
    '===============================================================================
    ''' <summary>
    ''' Copies a block of memory from one location to another.
    ''' </summary>
    ''' <param name="lpDest">Integer - A pointer to the starting address of the copied block's destination.</param>
    ''' <param name="lpSource">Integer - A pointer to the starting address of the block of memory to copy.</param>
    ''' <param name="nBytes">Integer - The size of the block of memory to copy, in bytes.</param>
    ''' <remarks>See ms-help://MS.W7SDK.1033/MS.W7SDKCOM.1033/memory/base/copymemory.htm for full documentation.</remarks>
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Integer, ByVal lpSource As Integer, ByVal nBytes As Integer)
    ''' <summary>
    ''' Computes the checksum of the specified file
    ''' </summary>
    ''' <param name="FileName">[in]String - The file name of the file for which the checksum is to be computed.</param>
    ''' <param name="HeaderSum">[out]Integer - A pointer to a variable that receives the original checksum from the image file, or zero if there is an error.</param>
    ''' <param name="CheckSum">[out]Integer - A pointer to a variable that receives the computed checksum.</param>
    ''' <returns>If the function succeeds, the return value is CHECKSUM_SUCCESS (0). See ms-help://MS.W7SDK.1033/MS.W7SDKCOM.1033/debug/base/mapfileandchecksum.htm for failure results</returns>
    ''' <remarks>See ms-help://MS.W7SDK.1033/MS.W7SDKCOM.1033/debug/base/mapfileandchecksum.htm</remarks>
    Private Declare Function MapFileAndCheckSumA Lib "IMAGEHLP.DLL" (ByVal FileName As String, ByRef HeaderSum As Integer, ByRef CheckSum As Integer) As Integer

    ' Note: add a project conditional compile argument "IncludeCreate"
    ' if the CreateShortKey is to be compiled into the application.
    '#If IncludeCreate = 1 Then
    ''' <summary>Creates a new serial number.</summary>
    ''' <param name="SerialNumber">String - The serial number is generated from the app name, version, and password, along with the HDD firmware serial number, which makes it unique for the machine running the app.</param>
    ''' <param name="Licensee">String - Name of party to whom this license is issued.</param>
    ''' <param name="ProductCode">Integer - A unique number assigned to this product. This is created from the app private key and is a 4 digit integer.</param>
    ''' <param name="ExpireDate">Date - Use this field for time-based serial numbers. This allows serial number to be issued that expire in two weeks or at the end of the year.</param>
    ''' <param name="UserData">Short - This field is caller defined. Currently we are using the MaxUser and LicType (using a LoByte/HiByte packed field).</param>
    ''' <param name="RegisteredLevel">Integer - This is the Registered Level from Alugen. Long only.</param>
    ''' <returns>String - A License Key in the form of "233C-3912-00FF-BE49"</returns>
    ''' <remarks></remarks>
    Friend Function CreateShortKey(ByVal SerialNumber As String, ByVal Licensee As String, ByVal ProductCode As Integer, ByVal ExpireDate As Date, ByVal UserData As Short, ByVal RegisteredLevel As Integer) As String

        Dim KeySegs(4) As String
        Dim i As Integer

        ' convert each segment value to a hex string
        KeySegs(Segments.iProdCode) = HexWORD(ProductCode)
        KeySegs(Segments.iExpire) = HexWORD(DateValue(CStr(CDate(ExpireDate.ToString("yyyy/MM/dd")))).Subtract(DateValue(CStr(#1/1/1970#))).Days)
        KeySegs(Segments.iUserData) = HexWORD(UserData)

        ' Compute CRC against pertinent info.
        KeySegs(Segments.iCRC) = HexWORD(CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(UCase(Licensee & KeySegs(Segments.iProdCode) & KeySegs(Segments.iExpire) & KeySegs(Segments.iUserData) & SerialNumber))))

        ' Perform bit swaps
        For i = 0 To m_nSwaps - 1
            SwapBit(m_Bits(i), KeySegs)
        Next i

        ' Calculate the CRC used to perform
        ' simple user input validation.
        KeySegs(Segments.iCRC2) = HexWORD(CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(UCase(Licensee & KeySegs(Segments.iProdCode) & KeySegs(Segments.iExpire) & KeySegs(Segments.iUserData) & KeySegs(Segments.iCRC)))))

        ' Return the key to the caller
        CreateShortKey = UCase(KeySegs(Segments.iProdCode) & "-" & KeySegs(Segments.iExpire) & "-" & KeySegs(Segments.iUserData) & "-" & KeySegs(Segments.iCRC) & "-" & KeySegs(Segments.iCRC2)) & "-" & StrReverse(HexWORD(RegisteredLevel))

    End Function
    '#End If

    '#If IncludeCheck = 1 Then
    ''' <summary>Performs a simple CRC test to ensure the key was entered "correctly". Does NOT validate
    ''' that the key is VALID. This function allows the caller to "test" the key input by the user,
    ''' without having to execute the key validation code, making it more work for a cracker to generate
    ''' a key for your application.
    ''' </summary>
    ''' <param name="LicenseKey">String - ?Not Documented!</param>
    ''' <param name="Licensee">String - ?Not Documented!</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Friend Function TestKey(ByVal LicenseKey As String, ByVal Licensee As String) As Boolean

        Dim KeySegs As Object = Nothing
        Dim nCRC As Integer

        On Error GoTo ExitLabel

        ' TODO: don't even call this function if SoftIce was detected!

        If Not SplitKey(LicenseKey, KeySegs) Then GoTo ExitLabel

        ' NOTE: Licensee can be omitted from the last checksum
        ' if there is no need to bind a serial number to a
        ' customer name.
        nCRC = CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(UCase(Licensee & KeySegs(Segments.iProdCode) & KeySegs(Segments.iExpire) & KeySegs(Segments.iUserData) & KeySegs(Segments.iCRC))))

        ' Compare check digits
        TestKey = (nCRC = SegmentValue(KeySegs(Segments.iCRC2)))

ExitLabel:

    End Function
    '#End If

    '#If IncludeValidate = 1 Then
    ''' <summary>
    ''' Evaluates the supplied license key and tests that it is valid. We do this by recomputing the checksum and comparing it to the one embedded in the serial number.
    ''' </summary>
    ''' <param name="LicenseKey">The license number to validate. Liberation Key.</param>
    ''' <param name="SerialNumber">A magic string that is application specific. This should be the same as was originally created by the application.</param>
    ''' <param name="Licensee">Name of party to whom this license is issued. This should be the same as was used to create the serial number.</param>
    ''' <param name="ProductCode">A unique 4 digit number assigned to this product. This should be the same as was used to create the license key.</param>
    ''' <param name="ExpireDate">Use this field for time-based serial numbers. This should be the same as was used to create the license key.</param>
    ''' <param name="UserData">This field is caller defined. This should be the same as was used to create the license key.</param>
    ''' <param name="RegisteredLevel"></param>
    ''' <returns>True if the license key checks out, False otherwise.</returns>
    ''' <remarks>See code for important notes!</remarks>
    Friend Function ValidateShortKey(ByVal LicenseKey As String, ByVal SerialNumber As String, ByVal Licensee As String, ByVal ProductCode As Integer, Optional ByRef ExpireDate As Date = #1/1/1970#, Optional ByRef UserData As Short = 0, Optional ByRef RegisteredLevel As Integer = 0) As Boolean
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

        Dim KeySegs As Object = Nothing
        Dim nCrc1 As Integer
        Dim nCrc2 As Integer
        Dim nCrc3 As Integer
        Dim nCrc4 As Integer
        Dim i As Integer

        On Error GoTo ExitLabel

        ' TODO: don't even call this function if SoftIce was detected!
        ' ----------------------------------------------------------
        ' This section of code could raise red flags
        ' ----------------------------------------------------------
        RegisteredLevel = SegmentValue(StrReverse(Mid(LicenseKey, 26, 4))) - 200
        LicenseKey = Mid(LicenseKey, 1, 24)
        If Not SplitKey(LicenseKey, KeySegs) Then GoTo ExitLabel

        ' ----------------------------------------------------------
        ' TODO: UCase string before it get's here

        ' Get CRC used for input validation
        nCrc1 = CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(UCase(Licensee) & KeySegs(Segments.iProdCode) & KeySegs(Segments.iExpire) & KeySegs(Segments.iUserData) & KeySegs(Segments.iCRC)))

        ' Compare check digits
        If (nCrc1 <> SegmentValue(KeySegs(Segments.iCRC2))) Then
            GoTo ExitLabel
        End If

        ' Perform bit swaps (in reverse).
        For i = m_nSwaps - 1 To 0 Step -1
            SwapBit(m_Bits(i), KeySegs)
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
        If ProductCode = SegmentValue(KeySegs(Segments.iProdCode)) Then
            ' One more check on the check digits before we
            ' blow away the value stored in nCrc1.
            If (SegmentValue(KeySegs(Segments.iCRC2)) = nCrc1) Then
                nCrc1 = CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(UCase(Licensee)))
            End If
        End If

        nCrc2 = CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(UCase(KeySegs(Segments.iProdCode))), nCrc1)
        nCrc3 = CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(UCase(KeySegs(Segments.iExpire))), nCrc2)
        nCrc3 = CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(UCase(KeySegs(Segments.iUserData))), nCrc3)
        nCrc4 = CRC(System.Text.UnicodeEncoding.Unicode.GetBytes(UCase(SerialNumber)), nCrc3)

        ' Return success and fill outputs IF the license KeySegs is valid
        If nCrc4 = SegmentValue(KeySegs(Segments.iCRC)) Then

            ' Fill the outputs with expire date and user data.
            ExpireDate = #1/1/1970#
            ExpireDate = ExpireDate.AddDays(SegmentValue(KeySegs(Segments.iExpire)))
            UserData = SegmentValue(KeySegs(Segments.iUserData))

            ' IMPORTANT: This is an easy patch point
            ' if you use real-time key validation.
            ValidateShortKey = True

        End If

ExitLabel:

    End Function
    '#End If

    '#If IncludeValidate = 1 Or IncludeCreate = 1 Then
    ''' <summary>
    ''' <para>This is used to swap various bits in the serial number. It's sole purpose is to alter the
    ''' output serial number.</para>
    ''' <para>This process is "played" forwards during the key creation, and in reverse when validating.
    ''' This mangling process should be identical for key creation and validation. Add as many
    ''' combinations as you like.</para>
    ''' </summary>
    ''' <param name="Word1">The words to bit swap. There are 4 words in the serial #. This parameter is zero-based.</param>
    ''' <param name="Bit1">The bits to swap. There are 16 bits to each word. This parameter is zero-based.</param>
    ''' <param name="Word2">The words to bit swap. There are 4 words in the serial #. This parameter is zero-based.</param>
    ''' <param name="Bit2">The bits to swap. There are 16 bits to each word. This parameter is zero-based.</param>
    ''' <remarks>
    ''' <para>Example: This scenario causes word 3, bit 8 to be swapped with word 1, bit 3!</para>
    ''' <para><code>KeyGen.AddSwapBits 1, 3, 3, 8</code></para>
    ''' <para>NOTE: It is recommended that there be at least 6 combinations in case the bits being swapped are the same (2 swap bits for words 2, 3 &amp; 4).</para>
    ''' </remarks>
    Friend Sub AddSwapBits(ByVal Word1 As Integer, ByVal Bit1 As Integer, ByVal Word2 As Integer, ByVal Bit2 As Integer)
        ' TODO: don't even call this function if SoftIce was detected!

        ' Size array to fit
        If m_nSwaps = 0 Then
            ReDim m_Bits(m_nSwaps)
        Else
            ReDim Preserve m_Bits(m_nSwaps)
        End If
        m_nSwaps = m_nSwaps + 1

        ' This implementation hardcodes keys that are 8 bytes/4 words
        If Word1 < 0 Or Word1 > 3 Or Word2 < 0 Or Word2 > 3 Then
            Set_locale(regionalSymbol)
            Err.Raise(5, Module_Renamed, "Word specification is not within 0-3.")
        End If

        ' There are only 16 bits to a word.
        If Bit1 < 0 Or Bit1 > 15 Or Bit2 < 0 Or Bit2 > 15 Then
            Set_locale(regionalSymbol)
            Err.Raise(5, Module_Renamed, "Bit specification is not within 0-15.")
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
    ''' <summary>
    ''' Shared code to massage the input serial number, and slice it into the required number of segments.
    ''' </summary>
    ''' <param name="LicenseKey">String - ?Undocumented!</param>
    ''' <param name="KeySegs">Object - ?Undocumented!</param>
    ''' <returns>Boolean - ?Undocumented!</returns>
    ''' <remarks></remarks>
    Private Function SplitKey(ByRef LicenseKey As String, ByRef KeySegs As Object) As Boolean
        ' ----------------------------------------------------------
        ' This section of code could raise red flags
        ' ----------------------------------------------------------

        ' Sanity check
        If InStr(LicenseKey, "-") = 0 Then GoTo ExitLabel

        ' As a courtesy to the user, we convert the
        ' letter "O" to the number "0". Users hate
        ' serialz that do not have interchangable 0/o's!
        LicenseKey = Replace(LicenseKey, "o", "0", , , CompareMethod.Text)

        ' Splice the KeySegs into 4 segments,
        ' exit if wrong # of segments.
        KeySegs = Split(UCase(LicenseKey), "-", 5)
        If UBound(KeySegs) <> 4 Then GoTo ExitLabel

        ' ----------------------------------------------------------

        SplitKey = True

ExitLabel:

    End Function
    '#End If

    ''' <summary>Converts a hex string representation into a 4 byte decimal value.</summary>
    ''' <param name="HexString">String - The Hex string to convert.</param>
    ''' <returns>Integer - Converted string</returns>
    ''' <remarks></remarks>
    Private Function SegmentValue(ByVal HexString As String) As Integer
        '===============================================================================
        '   Converts a hex string representation into a 4 byte decimal value.
        '===============================================================================

        'Dim Buffer(3) As Byte
        'Dim i As Integer
        'Dim j As Integer

        '' Exit if each byte not represented by a 2 character string
        'If Len(HexString) Mod 2 <> 0 Then Exit Function

        '' Exit if it's larger than a 4 byte value
        'If Len(HexString) > 8 Then Exit Function

        '' NOTE: we populate the byte array in little-endian format
        'For i = Len(HexString) To 1 Step -2
        '    Buffer(j) = CByte("&H" & Mid(HexString, i - 1, 2))
        '    j = j + 1
        'Next i

        '' Return the value
        'CopyMemory(VarPtr(SegmentValue), VarPtr(Buffer(0)), 4)
        SegmentValue = Int32.Parse(HexString, System.Globalization.NumberStyles.HexNumber)

    End Function


    '#If IncludeCreate = 1 Or IncludeValidate = 1 Then
    ''' <summary>
    ''' Swaps any two bits. The bits can differ as long as they are in the range of 0 and 15.
    ''' </summary>
    ''' <param name="BitList">TBits - ?Undocumented!</param>
    ''' <param name="KeySegs">Object - ?Undocumented!</param>
    ''' <remarks></remarks>
    Private Sub SwapBit(ByRef BitList As TBits, ByRef KeySegs As Object)
        With BitList
            ' Essentially, we swap Bit1 with Bit2. We use a bitwise
            ' OR operator or a bitwise AND operator depending
            ' upon if the subject bit is present. We don't use
            ' local variables to avoid synchronizing, especially
            ' since we may be doing a bit swap on the same word.
            If (SegmentValue(KeySegs(.iWord1)) And (2 ^ .iBit2)) = (2 ^ .iBit2) Then
                KeySegs(.iWord1) = HexWORD(SegmentValue(KeySegs(.iWord1)) Or (2 ^ .iBit2), "")
            Else
                KeySegs(.iWord1) = HexWORD(SegmentValue(KeySegs(.iWord1)) And Not (2 ^ .iBit2), "")
            End If

            If (SegmentValue(KeySegs(.iWord2)) And (2 ^ .iBit1)) = (2 ^ .iBit1) Then
                KeySegs(.iWord2) = HexWORD(SegmentValue(KeySegs(.iWord2)) Or (2 ^ .iBit1), "")
            Else
                KeySegs(.iWord2) = HexWORD(SegmentValue(KeySegs(.iWord2)) And Not (2 ^ .iBit1), "")
            End If

        End With

    End Sub
    '#End If

    ' Generic helper function
    ''' <summary>
    ''' Returns a 16-bit CRC value for a data block.
    ''' </summary>
    ''' <param name="Buffer">Byte - </param>
    ''' <param name="InputCrc">Integer - Optional - Defaults to 0</param>
    ''' <returns>Integer - </returns>
    ''' <remarks>Refer to CRC-CCITT compute-on-the-fly implementatations for more info.</remarks>
    Private Function CRC(ByRef Buffer() As Byte, Optional ByRef InputCrc As Integer = 0) As Integer

        Dim Bit As Integer
        Dim i As Integer
        Dim j As Integer

        On Error GoTo ErrHandler

        ' Derive from a prior CRC value if supplied.
        CRC = InputCrc

        ' Loop thru entire buffer computing the CRC
        For i = LBound(Buffer) To UBound(Buffer)

            ' Loop thru each of the 8 bits
            For j = 0 To 7

                Bit = ((CRC And &H8000) = &H8000) Xor ((Buffer(i) And (2 ^ j)) = 2 ^ j)

                CRC = (CShort(CRC And &H7FFF) * 2)

                If Bit <> 0 Then
                    CRC = CRC Xor &H1021
                End If

            Next j

        Next i

        Exit Function
        Resume
ErrHandler:
        System.Diagnostics.Debug.Assert(0, "")

    End Function

    ''' <summary>
    ''' Returns a hex string representation of a WORD.
    ''' </summary>
    ''' <param name="WORD">Integer - The 2 byte value to convert to a hex string.</param>
    ''' <param name="Prefix">A value such as "0x" or "&amp;H".</param>
    ''' <returns>String - The hex string representation of a Word</returns>
    ''' <remarks>NOTE:  It's up to the caller to ensure the subject value is a 16-bit number.</remarks>
    Private Function HexWORD(ByVal WORD As Integer, Optional ByVal Prefix As String = "") As String

        'Dim bytes(1) As Byte
        'Dim i As Integer, str As String

        'CopyMemory(VarPtr(bytes(0)), VarPtr(WORD), 2)

        HexWORD = UCase(System.Convert.ToString(WORD, 16))
        If Len(HexWORD) = 3 Then
            HexWORD = "0" & HexWORD
        ElseIf Len(HexWORD) = 2 Then
            HexWORD = "00" & HexWORD
        ElseIf Len(HexWORD) = 1 Then
            HexWORD = "000" & HexWORD
        End If

        'str = LCase(System.Convert.ToString(WORD, 16))
        'Dim encoding As New System.Text.ASCIIEncoding
        'bytes = encoding.GetBytes(str)

        'HexWORD = Prefix
        'For i = UBound(bytes) To LBound(bytes, ) Step -1
        '    If Len(Hex(bytes(i))) = 1 Then
        '        HexWORD = HexWORD & "0" & LCase(Hex(bytes(i)))
        '    Else
        '        HexWORD = HexWORD & LCase(Hex(bytes(i)))
        '    End If
        'Next i

    End Function

    ''' <summary>
    ''' Tests if the supplied file has been altered by computing a checksum for the file and comparing
    ''' it against the checksum in the executable image.
    ''' </summary>
    ''' <param name="FilePath">Full path to file to check. Caller is responsible for ensuring that the path exists, and that it is an executable.</param>
    ''' <returns>Boolean - </returns>
    ''' <remarks>See notes within for more information.</remarks>
    Friend Function ExeIsPatched(ByVal FilePath As String) As Boolean

        Dim FileCRC As Integer
        Dim HdrCRC As Integer
        Dim ErrorCode As Integer

        ' NOTE: Many crackers today are smart enough to
        ' update the PE image CRC value. But we check
        ' anyhow, just in case. Otherwise, it could be
        ' embarrassing if the EXE was patched without
        ' updating the PE header.

        ErrorCode = MapFileAndCheckSumA(FilePath, HdrCRC, FileCRC)
        If ErrorCode = MapFileChecksumErrors.CHECKSUM_SUCCESS Then

            If HdrCRC <> 0 And HdrCRC <> FileCRC Then

                ' CRC of file is different than the CRC
                ' embedded in the PE image. Try not to
                ' let the cracker know that you are on
                ' to him. And don't start deleting from
                ' their harddrive!
                ExeIsPatched = True

            End If

        End If

    End Function
End Class