VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*******************************************************************************
' This code is based in part on the Registration Class by Phil Fresle
' The KEY part of the Code is Generated using a Supplied Name which is
' parsed thru the Soundex Function. This enables the Developer to register
' the application to a User / Company Name.
'
' DESCRIPTION:
' This class demonstrates how you can create semi-random license keys for
' your applications and test their validity.
'
' The 12 character random key gives around 1 billion different combinations but
' a user would only have to try an average of Modulus keys to randomly find
' a valid value. This works for me as I am pretty sure to get a unique key
' whilst at least discouraging a random attack. By storing the keys in a
' database along with who they have been issued to I can track key use.
'
' If this class was put into a DLL you would not want to expose the Generate
' method on the default interface or anyone could use it.
'
' This is 'free' software with the following restrictions:
'
' You may not redistribute this code as a 'sample' or 'demo'. However, you are free
' to use the source code in your own code, but you may not claim that you created
' the sample code. It is expressly forbidden to sell or profit from this source code
' other than by the knowledge gained or the enhanced value added by your own code.
'
' Use of this software is also done so at your own risk. The code is supplied as
' is without warranty or guarantee of any kind.
'
'*******************************************************************************
Option Explicit

Private oCrypto                 As CryptoX
Private SLEEP_INTERVAL          As Long 'Intercept Brute Force
Private KEYFMT                  As Long 'Used for Formatting Numbers

Private Const BASE_DATE         As Long = 36891 '31 December 2000
Private Const HIGH_DATE         As Long = 99999 '13 October 2173
Private Const PART_LENGTH       As Long = 4
Private Const PART1_LENGTH      As Long = 5
Private Const KEY_LENGTH        As Long = PART1_LENGTH + (3 * PART_LENGTH)

Private Const POS_01_MULTIPLIER As Long = 17
Private Const POS_02_MULTIPLIER As Long = 12
Private Const POS_03_MULTIPLIER As Long = 10
Private Const POS_04_MULTIPLIER As Long = 8
Private Const POS_05_MULTIPLIER As Long = 4
Private Const POS_06_MULTIPLIER As Long = 15
Private Const POS_07_MULTIPLIER As Long = 7
Private Const POS_08_MULTIPLIER As Long = 5
Private Const POS_09_MULTIPLIER As Long = 3
Private Const POS_10_MULTIPLIER As Long = 13
Private Const POS_11_MULTIPLIER As Long = 11
Private Const POS_12_MULTIPLIER As Long = 6
Private Const POS_13_MULTIPLIER As Long = 2
Private Const POS_14_MULTIPLIER As Long = 16
Private Const POS_15_MULTIPLIER As Long = 14
Private Const POS_16_MULTIPLIER As Long = 9
Private Const POS_17_MULTIPLIER As Long = 1

Private mModulus As Long
Private arrPrimes() As Long
Public Enum CDStatus
    CDOkay = 0
    CDTampered = 1
    CDExpired = 2
End Enum

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Property Let Modulus(vData As Long)
    If IsPrime(vData) Then
        mModulus = vData
    Else
        mModulus = 231
    End If
End Property
Public Property Get Modulus() As Long
    Modulus = mModulus
End Property

'*******************************************************************************
' DESCRIPTION:
' Public interface to create a key
'   Licencee and Expiry Date is used in the FIRST KEY PART
'*******************************************************************************
Public Function MakeKeyCode(mKeyValue As String, mExpireDate As Date, EncryptKey As String) As String
    Dim lKeyValue As String
    Dim sPart1    As String
    Dim lSndX     As String

    oCrypto.Keyword = EncryptKey
    lKeyValue = oCrypto.Encrypt(mKeyValue)
    lSndX = Right(SoundsLike(lKeyValue, PART1_LENGTH + 1), PART1_LENGTH)
    sPart1 = ZeroPad(GetDateCode(CLng(lSndX), mExpireDate), PART1_LENGTH)
    MakeKeyCode = GenerateKey(sPart1)
End Function

'*******************************************************************************
' DESCRIPTION:
' Public interface to check a key
' Invalid Codes produce incremental SLEEPS to
'   discourage brute force attempts.
'*******************************************************************************
Public Function Status(ByVal mKeyValue As String, sKey As String, mExpireDate As Date, EncryptKey As String) As CDStatus
    'Brian Gillham: Added auto-increment to discourage BRUTE FORCE
    Dim lKeyValue As String
    Dim lSndX   As String
    Dim sPart1  As String

    oCrypto.Keyword = EncryptKey
    lKeyValue = oCrypto.Encrypt(mKeyValue)
    sKey = Replace(sKey, "-", "")
    sKey = Replace(sKey, " ", "")
    sKey = Replace(sKey, ".", "")
    sKey = Replace(sKey, "+", "")
    Status = CDTampered 'Assume Tampering
    lSndX = Right(SoundsLike(lKeyValue, PART1_LENGTH + 1), PART1_LENGTH)
    sPart1 = ZeroPad(GetDateCode(CLng(lSndX), mExpireDate), PART1_LENGTH)
    If Left(sKey, PART1_LENGTH) = sPart1 Then
        If KeyOK(sKey) Then
            If mExpireDate <= Now Then
                Status = CDExpired
            Else
                Status = CDOkay
            End If
        End If
    End If
    If Status = CDTampered Then
        If SLEEP_INTERVAL < 32000 Then
            'Keep doubling until 65000 ms
            SLEEP_INTERVAL = SLEEP_INTERVAL * SLEEP_INTERVAL
        End If
    Else
        SLEEP_INTERVAL = 2  'Milliseconds
    End If
    Sleep SLEEP_INTERVAL
End Function

Private Function GetDateCode(ByVal SndX As Long, ByVal ChkDate As Date) As Long
    Dim lDate    As Long
    Dim lCode    As Long
    
    lDate = CLng(ChkDate)   ' Convert date to Number
    If lDate > HIGH_DATE Then lDate = HIGH_DATE ' Check if date is valid
    lCode = SndX + lDate
    'Check which is greater
    If lCode > HIGH_DATE Then
        GetDateCode = lCode - HIGH_DATE
    Else
        GetDateCode = lCode
    End If
End Function
'*******************************************************************************
' PartOK (FUNCTION)
'
' PARAMETERS:
' (In) - sPart - String - Part to check
'
' RETURN VALUE:
' Boolean - True if OK
'
' DESCRIPTION:
' Makes sure part of the key is not easily guessed
'*******************************************************************************
Private Function PartOK(ByVal sPart As String) As Boolean
    Dim sCheck As String
    sCheck = "1234567890"
    PartOK = False
    sPart = Trim(sPart)
    
    If (Len(sPart) = PART_LENGTH) Or (IsNumeric(sPart)) Then
        Select Case sPart
            Case String(PART_LENGTH, "1"), String(PART_LENGTH, "2"), String(PART_LENGTH, "3")
            Case String(PART_LENGTH, "4"), String(PART_LENGTH, "5"), String(PART_LENGTH, "6")
            Case String(PART_LENGTH, "7"), String(PART_LENGTH, "8"), String(PART_LENGTH, "9")
            Case String(PART_LENGTH, "0"), Left(sCheck, PART_LENGTH), Left(StrReverse(Left(sCheck, PART_LENGTH)), PART_LENGTH)
            Case Else:  PartOK = True
        End Select
    End If
End Function

Private Function GenerateKey(ByVal sPart1 As String) As String
    Dim sKey        As String
    Dim sPart2      As String
    Dim sPart3      As String
    Dim sPart4      As String
    Dim lCount      As Long
    Dim lModulus    As Long
    Dim lNextNumber As Long

    On Error GoTo LocalError
    If Len(sPart1) <> PART1_LENGTH Then
        Err.Raise vbObjectError, , "Part 1 must be " & PART1_LENGTH & " characters"
    ElseIf Not IsNumeric(sPart1) Then
        Err.Raise vbObjectError, , "Part 1 must be numeric"
    Else
        ' Could be a decimal or negative number, so make sure
        For lCount = 1 To Len(sPart1)
            If Not IsNumeric(Mid(sPart1, lCount, 1)) Then
                Err.Raise vbObjectError, , "Part 1 must be a whole positive numeric number"
            End If
        Next
    End If
    
    ' We want a random key
    Randomize
    
    ' Loop till we get a valid key
    Do
        ' Save time by pre-validating each part
        Do
            sPart2 = ""
            ' Put together four random numbers
            For lCount = 1 To PART_LENGTH
                sPart2 = sPart2 & CStr(Int(10 * Rnd))
            Next
        Loop Until PartOK(sPart2)
    
        Do
            sPart3 = ""
            ' Put together four random numbers
            For lCount = 1 To PART_LENGTH
                sPart3 = sPart3 & CStr(Int(10 * Rnd))
            Next
        Loop Until PartOK(sPart3)
    
        ' PMF - 07-Sep-2000 15:29
        ' This is where changes were made for version 1.1 so that keys are
        ' more quickly created than simply gererating over and over until a valid
        ' key is found
        
        ' Start with one random number
        sPart4 = CStr(Int(10 * Rnd))
        
        ' This part of the code makes sure that the last three numbers when added
        ' to the existing numbers will ensure that the full key is Modulus
        lModulus = Modulus - (GetCheckSum(sPart1 & sPart2 & sPart3 & sPart4) Mod Modulus)
        lNextNumber = lModulus \ POS_15_MULTIPLIER
        If lNextNumber > 9 Then
            lNextNumber = 9
        End If
        sPart4 = sPart4 & CStr(lNextNumber)
        lModulus = lModulus - (lNextNumber * POS_15_MULTIPLIER)
        
        lNextNumber = lModulus \ POS_16_MULTIPLIER
        If lNextNumber > 9 Then
            lNextNumber = 9
        End If
        sPart4 = sPart4 & CStr(lNextNumber)
        lModulus = lModulus - (lNextNumber * POS_16_MULTIPLIER)
        
        lNextNumber = lModulus \ POS_17_MULTIPLIER
        sPart4 = sPart4 & CStr(lNextNumber)
        
        ' Put together the full key for checking
        sKey = sPart1 & sPart2 & sPart3 & sPart4
    Loop Until KeyOK(sKey)
    GenerateKey = sKey
Exit Function
LocalError:
    GenerateKey = "99999999999999999"
End Function

'*******************************************************************************
' KeyOK (FUNCTION)
'
' PARAMETERS:
' (In) - sKey - String - Key to check
'
' RETURN VALUE:
' Boolean - True if OK
'
' DESCRIPTION:
' Checks if the entire key is OK by a customised sum of the key, the way this
' is calculated should be unique to you, so I suggest you change the algorithm.
'*******************************************************************************
Private Function KeyOK(ByVal sKey As String) As Boolean
    Dim lChecksum   As Long
    Dim sPart2      As String
    Dim sPart3      As String
    Dim sPart4      As String
    Dim lCount      As Long

    KeyOK = False
    sKey = Trim(sKey)

    If Len(sKey) <> KEY_LENGTH Then
        KeyOK = False
        Exit Function
    ElseIf Not IsNumeric(sKey) Then
        KeyOK = False
        Exit Function
    Else
        For lCount = 1 To Len(sKey)
            If Not IsNumeric(Mid(sKey, lCount, 1)) Then
                KeyOK = False
                Exit Function
            End If
        Next
    End If

    ' Make sure key parts are valid
    sPart2 = Mid(sKey, PART1_LENGTH + 1, PART_LENGTH)
    sPart3 = Mid(sKey, PART1_LENGTH + PART_LENGTH + 1, PART_LENGTH)
    sPart4 = Mid(sKey, PART1_LENGTH + (2 * PART_LENGTH) + 1, PART_LENGTH)

    If Not PartOK(sPart4) Then
        KeyOK = False
        Exit Function
    ElseIf Not PartOK(sPart3) Then
        KeyOK = False
        Exit Function
    ElseIf Not PartOK(sPart2) Then
        KeyOK = False
        Exit Function
    End If

    ' Compute checksum
    lChecksum = GetCheckSum(sKey)
    ' If the checksum is modulus then it is valid
    If (lChecksum Mod Modulus) = 0 Then
        KeyOK = True
    Else
        KeyOK = False
    End If
End Function

'*******************************************************************************
' GetCheckSum (FUNCTION)
'
' PARAMETERS:
' (In/Out) - sKey - String - Key to compute checksum with.
'
' RETURN VALUE:
' Lone - Value of checksum.
'
' DESCRIPTION:
' This function calculates the custom checksum. It may be called with a full
' key from the validation routine, or a partial key from the key generation
' routine.
'
' NOTE: This function was added for version 1.1
'*******************************************************************************
Private Function GetCheckSum(sKey As String) As Long
    Dim lChecksum As Long
    
    lChecksum = 0
    
    lChecksum = lChecksum + (POS_01_MULTIPLIER * CLng(Mid(sKey, 1, 1)))
    lChecksum = lChecksum + (POS_02_MULTIPLIER * CLng(Mid(sKey, 2, 1)))
    lChecksum = lChecksum + (POS_03_MULTIPLIER * CLng(Mid(sKey, 3, 1)))
    lChecksum = lChecksum + (POS_04_MULTIPLIER * CLng(Mid(sKey, 4, 1)))
    lChecksum = lChecksum + (POS_05_MULTIPLIER * CLng(Mid(sKey, 5, 1)))
    lChecksum = lChecksum + (POS_06_MULTIPLIER * CLng(Mid(sKey, 6, 1)))
    lChecksum = lChecksum + (POS_07_MULTIPLIER * CLng(Mid(sKey, 7, 1)))
    lChecksum = lChecksum + (POS_08_MULTIPLIER * CLng(Mid(sKey, 8, 1)))
    lChecksum = lChecksum + (POS_09_MULTIPLIER * CLng(Mid(sKey, 9, 1)))
    lChecksum = lChecksum + (POS_10_MULTIPLIER * CLng(Mid(sKey, 10, 1)))
    lChecksum = lChecksum + (POS_11_MULTIPLIER * CLng(Mid(sKey, 11, 1)))
    lChecksum = lChecksum + (POS_12_MULTIPLIER * CLng(Mid(sKey, 12, 1)))
    lChecksum = lChecksum + (POS_13_MULTIPLIER * CLng(Mid(sKey, 13, 1)))
    lChecksum = lChecksum + (POS_14_MULTIPLIER * CLng(Mid(sKey, 14, 1)))
    
    If Len(sKey) = KEY_LENGTH Then
        lChecksum = lChecksum + (POS_15_MULTIPLIER * CLng(Mid(sKey, 15, 1)))
        lChecksum = lChecksum + (POS_16_MULTIPLIER * CLng(Mid(sKey, 16, 1)))
        lChecksum = lChecksum + (POS_17_MULTIPLIER * CLng(Mid(sKey, 17, 1)))
    End If
    
    GetCheckSum = lChecksum
End Function

Private Function SoundsLike(ByVal pWord As String, Optional pAccuracy As Byte = 6) As String
    On Error GoTo LocalError
    '  char importance "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim x As Integer
    Dim CChar As String

    If pAccuracy > 10 Then
        pAccuracy = 10              ' maximum accuracy allowed
    ElseIf pAccuracy < 4 Then
        pAccuracy = 4               ' minimum accuracy allowed
    End If
    pAccuracy = pAccuracy - 1       ' account for the first character

    pWord = UCase(pWord)
    
    ' strip all invalid characters
    For x = 1 To Len(pWord)
        If Asc(Mid(pWord, x, 1)) < 65 Or Asc(Mid(pWord, x, 1)) > 90 Then
           Mid(pWord, x, 1) = "@"   ' assign a catchable value
        End If
    Next x
    pWord = Trim(pWord)

    SoundsLike = pWord
    ' assign values to the string
    SoundsLike = Replace(SoundsLike, "A", "0")
    SoundsLike = Replace(SoundsLike, "E", "0")
    SoundsLike = Replace(SoundsLike, "I", "0")
    SoundsLike = Replace(SoundsLike, "O", "0")
    SoundsLike = Replace(SoundsLike, "U", "0")
    SoundsLike = Replace(SoundsLike, "Y", "0")
    SoundsLike = Replace(SoundsLike, "H", "0")
    SoundsLike = Replace(SoundsLike, "W", "0")
    SoundsLike = Replace(SoundsLike, "B", "1")
    SoundsLike = Replace(SoundsLike, "P", "1")
    SoundsLike = Replace(SoundsLike, "F", "1")
    SoundsLike = Replace(SoundsLike, "V", "1")
    SoundsLike = Replace(SoundsLike, "C", "2")
    SoundsLike = Replace(SoundsLike, "S", "2")
    SoundsLike = Replace(SoundsLike, "G", "2")
    SoundsLike = Replace(SoundsLike, "J", "2")
    SoundsLike = Replace(SoundsLike, "K", "2")
    SoundsLike = Replace(SoundsLike, "Q", "2")
    SoundsLike = Replace(SoundsLike, "X", "2")
    SoundsLike = Replace(SoundsLike, "Z", "2")
    SoundsLike = Replace(SoundsLike, "D", "3")
    SoundsLike = Replace(SoundsLike, "T", "3")
    SoundsLike = Replace(SoundsLike, "L", "4")
    SoundsLike = Replace(SoundsLike, "M", "5")
    SoundsLike = Replace(SoundsLike, "N", "5")
    SoundsLike = Replace(SoundsLike, "R", "6")

    CChar = Left(SoundsLike, 1)
    For x = 2 To Len(SoundsLike)
        If Mid(SoundsLike, x, 1) = CChar Then
            Mid(SoundsLike, x, 1) = "@"
        Else
            CChar = Mid(SoundsLike, x, 1)
        End If
    Next x
    SoundsLike = Replace(SoundsLike, "@", "")

    SoundsLike = Mid(SoundsLike, 2)
    SoundsLike = Replace(SoundsLike, "0", "")

    SoundsLike = SoundsLike & String(pAccuracy, "0")
    SoundsLike = Left(pWord, 1) & Left(SoundsLike, pAccuracy)
Exit Function
LocalError:
End Function

Private Function ZeroPad(pNumber As Long, pLength As Long) As String
    If IsNumeric(pNumber) And IsNumeric(pLength) Then
        Dim Padding
        Padding = String(pLength, "0")
        ZeroPad = Padding & CStr(pNumber)
        ZeroPad = Right(ZeroPad, pLength)
    Else
        ZeroPad = pNumber
    End If
End Function

Public Function IsPrime(TestPrime As Long) As Boolean
    Dim TestNum As Long
    Dim TestLimit As Long
    
    '   Eliminate even numbers
    If TestPrime Mod 2 = 0 Then Exit Function
    
    '   Loop through ODD numbers starting with 3
    TestNum = 3
    TestLimit = TestPrime
    Do While TestLimit > TestNum
        If TestPrime Mod TestNum = 0 Then Exit Function
        '   There's logic to this.  Think about it.
        TestLimit = TestPrime \ TestNum
        '   Remember, we only bother to check odd numbers
        TestNum = TestNum + 2
    Loop
    '   If we made it through the loop, the number is a prime.
    IsPrime = True
End Function

Private Sub Class_Initialize()
    Set oCrypto = New CryptoX
    SLEEP_INTERVAL = 2
    KEYFMT = CLng("1" & String(PART1_LENGTH, "0"))
    Modulus = 231
End Sub

Private Sub Class_Terminate()
    Set oCrypto = Nothing
End Sub
