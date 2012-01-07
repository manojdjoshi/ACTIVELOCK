VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CryptoX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Type of encryption to use
Public Enum EncryptionMethod
    BlockEncryption = 1
    StreamEncryption = 2
End Enum
    
' Error numbers (public enum so user's of the DLL can see/use them)
Public Enum CryptoErrors
    ErrorAquiringContext = vbObjectError + 1056
    ErrorCreatingHash = vbObjectError + 1057
    ErrorCreatingHashData = vbObjectError + 1058
    ErrorDerivingKey = vbObjectError + 1059
    ErrorEncryptingData = vbObjectError + 1060
    ErrorDecryptingData = vbObjectError + 1061
    ErrorInvalidHexString = vbObjectError + 1062
    ErrorMissingParameter = vbObjectError + 1063
    ErrorBadEncryptionType = vbObjectError + 1064
End Enum

' Error messages
Private Const ERROR_AQUIRING_CONTEXT    As String = "Could not acquire context"
Private Const ERROR_CREATING_HASH       As String = "Could not create hash"
Private Const ERROR_CREATING_HASH_DATA  As String = "Could not create hash data"
Private Const ERROR_DERIVING_KEY        As String = "Could not derive key"
Private Const ERROR_ENCRYPTING_DATA     As String = "Could not encrypt data"
Private Const ERROR_DECRYPTING_DATA     As String = "Could not decrypt data"
Private Const ERROR_INVALID_HEX_STRING  As String = "Not a valid hex string"
Private Const ERROR_MISSING_PARAMETER   As String = "Both a string and a key are required"
Private Const ERROR_BAD_ENCRYPTION_TYPE As String = "Invalid encryption type specified"

' The provider
Private Const CRYPTO_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"

' For the context
Private Const CRYPT_VERIFYCONTEXT   As Long = &HF0000000
Private Const PROV_RSA_FULL         As Long = 1

' Hashing algorithms for the session key
Private Const ALG_CLASS_HASH    As Long = 32768             ' (4 << 13)
Private Const ALG_TYPE_ANY      As Long = 0
Private Const ALG_SID_MD5       As Long = 3
Private Const CALG_MD5          As Long = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5

' For the session key
Private Const CRYPT_NO_SALT As Long = &H10

' Encryption algorithms
Private Const ALG_CLASS_DATA_ENCRYPT    As Long = 24576     ' (3 << 13)
Private Const ALG_TYPE_BLOCK            As Long = 1536      ' (3 << 9)
Private Const ALG_TYPE_STREAM           As Long = 2048      ' (4 << 9)
Private Const ALG_SID_RC2               As Long = 2
Private Const ALG_SID_RC4               As Long = 1
Private Const CALG_RC2                  As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2
Private Const CALG_RC4                  As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4

' Module level variables
Private m_lProvider As Long

' We will be using this to slow down decryption to discourage hacking
' Declarations for the CryptoAPI functions
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal algID As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal algID As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

'Property Values
Private mErrorMsg  As String             'Trap the Error
Private mKeyword   As String             'Encryption Key
Private mEncrypted As String             'Encryption Key
Private mAsHex     As Boolean            'Return Encryption as / from Hex
Private mMethod    As EncryptionMethod   'Encryption Method

'##############################
'# Publicy Exposed Properties #
'##############################
Public Property Let ErrorMsg(vData As String)
    mErrorMsg = vData
End Property
Public Property Get ErrorMsg() As String
    ErrorMsg = mErrorMsg
End Property

Public Property Let Keyword(vData As String)
    mKeyword = vData
End Property
Public Property Get Keyword() As String
    Keyword = mKeyword
End Property

Public Property Let Encrypted(vData As String)
    mEncrypted = vData
End Property
Public Property Get Encrypted() As String
    Encrypted = mEncrypted
End Property

Public Property Let Method(vData As EncryptionMethod)
    mMethod = vData
End Property
Public Property Get Method() As EncryptionMethod
    Method = mMethod
End Property

Public Property Let AsHex(vData As Boolean)
    mAsHex = vData
End Property
Public Property Get AsHex() As Boolean
    AsHex = mAsHex
End Property

'#############################
'# Publicy Exposed Functions #
'#############################
Public Function Decrypt(Encrypted As String) As String
    ErrorMsg = ""
    If Method <> BlockEncryption And Method <> StreamEncryption Then
        ErrorMsg = ErrorBadEncryptionType & ": Crypto.Decrypt - " & ERROR_BAD_ENCRYPTION_TYPE
        Exit Function
    End If
    If Encrypted = "" Or Keyword = "" Then
        ErrorMsg = ErrorMissingParameter & ": Crypto.Decrypt - " & ERROR_MISSING_PARAMETER
        Exit Function
    End If
    If AsHex Then
        Encrypted = ConvertStringFromHex(Encrypted)
    End If
    Decrypt = EncryptDecrypt(Encrypted, Keyword, False, Method)
End Function

Public Function Encrypt(RawData As String) As String
    Dim sEncryptedText As String
    ErrorMsg = ""
    Encrypted = ""
    If Method <> BlockEncryption And Method <> StreamEncryption Then
        ErrorMsg = ErrorBadEncryptionType & ": Crypto.Encrypt - " & ERROR_BAD_ENCRYPTION_TYPE
        Exit Function
    End If
    If RawData = "" Or Keyword = "" Then
        ErrorMsg = ErrorMissingParameter & ": Crypto.Encrypt - " & ERROR_MISSING_PARAMETER
        Exit Function
    End If
    sEncryptedText = EncryptDecrypt(RawData, Keyword, True, Method)
    If AsHex Then
        Encrypt = ConvertStringToHex(sEncryptedText)
    Else
        Encrypt = sEncryptedText
    End If
    Encrypted = Encrypt
End Function

'############################
'# Private Hidden Functions #
'############################
Private Function EncryptDecrypt(ByVal sText As String, ByVal sKeyRoot As String, ByVal bEncrypt As Boolean, ByVal Method As EncryptionMethod) As String
    Dim lResult             As Long
    Dim lHash               As Long
    Dim lKey                As Long
    Dim lHashPassword       As Long
    Dim lFlags              As Long
    Dim lData               As Long
    Dim lClear              As Long
    Dim sErr                As String
    Dim lErr                As Long
    Dim lEncryptionType     As Long
    
    On Error GoTo ERROR_HANDLER
    
    ' Create a handle to a hash object  using the MD5 algorithm
    lResult = CryptCreateHash(m_lProvider, CALG_MD5, 0, 0, lHash)
    If lResult = 0 Then
        ErrorMsg = ErrorCreatingHash & ":Crypto.EncryptDecrypt - " & ERROR_CREATING_HASH
        Exit Function
    End If
    
    ' Add some data to the hash object for use in generating our key
    ' sKeyRoot is, in effect, our key and should be fairly complex and
    ' not easily guessed
    lHashPassword = Len(sKeyRoot)
    lResult = CryptHashData(lHash, sKeyRoot, lHashPassword, 0)
    If lResult = 0 Then
        ErrorMsg = ErrorCreatingHashData & "Crypto.EncryptDecrypt - " & ERROR_CREATING_HASH_DATA
        Exit Function
    End If
    
    ' Select appropriate encryption method
    If Method = BlockEncryption Then
        lEncryptionType = CALG_RC2
    Else
        lEncryptionType = CALG_RC4
    End If
    
    ' Generate a session key for use with the cypher
    lFlags = CRYPT_NO_SALT
    lResult = CryptDeriveKey(m_lProvider, lEncryptionType, lHash, lFlags, lKey)
    If lResult = 0 Then
        ErrorMsg = ErrorDerivingKey & ": Crypto.EncryptDecrypt" & ERROR_DERIVING_KEY
        Exit Function
    End If
    
    ' Encrypt or decrypt data as required
    lData = Len(sText)
    If bEncrypt Then
        lClear = lData
        ' Call with a null first to see how long the string needs to be
        lResult = CryptEncrypt(lKey, 0, 1, 0, vbNullString, lData, lClear)
        sText = sText & String(lData - lClear, " ")
        ' Encrypt some text
        lResult = CryptEncrypt(lKey, 0, 1, 0, sText, lClear, lData)
        If lResult = 0 Then
            ErrorMsg = ErrorEncryptingData & ": Crypto.EncryptDecrypt - " & ERROR_ENCRYPTING_DATA
            Exit Function
        End If
    Else
        lResult = CryptDecrypt(lKey, 0, 1, 0, sText, lData)
        sText = Left(sText, lData)
        If lResult = 0 Then
            ErrorMsg = ErrorDecryptingData & ": Crypto.EncryptDecrypt - " & ERROR_DECRYPTING_DATA
            Exit Function
        End If
    End If
    
    EncryptDecrypt = sText
TIDY_UP:
    On Error Resume Next
        
    If lHash <> 0 Then
        lResult = CryptDestroyHash(lHash)
    End If
    If lKey <> 0 Then
        lResult = CryptDestroyKey(lKey)
    End If
    Err.Clear
    
    If lErr <> 0 Then
        ErrorMsg = lErr & ": Crypto Error! - " & sErr
        Exit Function
    End If
Exit Function
ERROR_HANDLER:
    lErr = Err.Number
    sErr = Err.Description
    Resume TIDY_UP
End Function

Private Function ConvertStringToHex(ByVal sText As String) As String
    Dim lCount  As Long
    Dim sHex    As String
    Dim sResult As String

    For lCount = 1 To Len(sText)
        sHex = Hex(Asc(Mid(sText, lCount, 1)))
        If Len(sHex) = 1 Then
            sHex = "0" & sHex
        End If
        sResult = sResult & sHex
    Next
    ConvertStringToHex = sResult
End Function

Private Function ConvertStringFromHex(ByVal sText As String) As String
    Dim lCount  As Long
    Dim sChar   As String
    Dim sResult As String
    Dim lLength As Long

    lLength = Len(sText)
    If lLength Mod 2 <> 0 Then
        ErrorMsg = ErrorInvalidHexString & ": Crypo.ConvertStringFromHex - " & ERROR_INVALID_HEX_STRING
        Exit Function
    End If
    For lCount = 1 To lLength
        sChar = Mid(sText, lCount, 1)
        If sChar < "0" Or sChar > "9" Then
            If sChar < "A" Or sChar > "F" Then
                ErrorMsg = ErrorInvalidHexString & ": Crypto.ConvertStringFromHex - " & ERROR_INVALID_HEX_STRING
                Exit Function
            End If
        End If
    Next
    For lCount = 1 To lLength Step 2
        sResult = sResult & Chr("&H" & Mid(sText, lCount, 2))
    Next
    
    ConvertStringFromHex = sResult
End Function

'*******************************************************************************
' Class_Initialize (SUB)
' PARAMETERS:
' None
' DESCRIPTION:
' When class is initialized get a handle to the context
'*******************************************************************************
Private Sub Class_Initialize()
    Dim lResult As Long
    lResult = CryptAcquireContext(m_lProvider, vbNullString, CRYPTO_PROVIDER, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
    If lResult = 0 Then
        ErrorMsg = ErrorAquiringContext & ": Crypto.Class_Initialize - " & ERROR_AQUIRING_CONTEXT
        Exit Sub
    Else
        Method = BlockEncryption
        AsHex = True
    End If
End Sub

'*******************************************************************************
' Class_Terminate (SUB)
' PARAMETERS:
' None
' DESCRIPTION:
' When class terminates, tidy up the context
'*******************************************************************************
Private Sub Class_Terminate()
    Dim lResult As Long
    If m_lProvider <> 0 Then
        lResult = CryptReleaseContext(m_lProvider, 0)
    End If
End Sub
