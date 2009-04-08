VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCryptoAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Wrapper class for Microsoft CryptoAPI (Base CSP).
'I release this code (CryptAPI.cls) to the public domain for free use.
'All other code is kept under the laws of copyright.
'This code was written by Shell Pierce. <drshell@mindspring.com>

'Reference to Doug Gaede, who on December 11, 2000, released his examples
'of using the Windows CryptoAPI...
'I have modified it according to conditions defined and described in
'Richard Bondi's book "Cryptography for Visual Basic"

Private Const INF_PLTFRM_WIN32s = 0
Private Const INF_PLTFRM_WIN32_WINDOWS = 1
Private Const INF_PLTFRM_WIN32_NT = 2

'Provider & Type
Private Const BASE_PROVIDER            As String = "Microsoft Base Cryptographic Provider v1.0"
Private Const ENHANSED_PROVIDER        As String = "Microsoft Enhanced Cryptographic Provider v1.0"
Private Const PROV_RSA_FULL            As Long = 1

'dwFlags definitions for CryptAcquireContext

Private Const CRYPT_VERIFYCONTEXT = &HF0000000 ' The application has no access to the private keys, and the pszContainer parameter must be set to NULL.  This option is intended for applications that do not use private keys.  When CryptAcquireContext is called, many CSPs require input from the owning user before granting access to the private keys in the key container. For example, the private keys can be encrypted, requiring a password from the user before they can be used. However, if the CRYPT_VERIFYCONTEXT flag is specified, access to the private keys is not required and the user interface can be bypassed.
Private Const CRYPT_EXISTINGKEYSET     As Long = &H0
Private Const CRYPT_NEWKEYSET = &H8            ' A new key container is created with the name specified by pszContainer. If pszContainer is NULL, a key container with the default name is created.
Private Const CRYPT_MACHINE_KEYSET = &H20      ' By default, keys and key containers are stored as user keys. For Base Providers, this means that user key containers are stored in the user's profile. The CRYPT_MACHINE_KEYSET flag can be combined with all of the other flags to indicate that the key container of interest is a machine key container and the CSP treats it as such. For Base Providers, this means that the keys are stored locally on the computer that created the key container. If a key container is to be a machine container, the CRYPT_MACHINE_KEYSET flag must be used with all calls to CryptAcquireContext that reference the machine container. The CRYPT_MACHINE_KEYSET flag is useful when the user is accessing from a service or user account that did not log on interactively When key containers are created, most CSPs do not automatically create any public/private key pairs. These keys must be created as a separate step with the CryptGenKey function.
Private Const CRYPT_DELETEKEYSET = &H10        ' The key container specified by pszContainer is deleted. If pszContainer is NULL, the key container with the default name is deleted. All key pairs in the key container are also destroyed.  When this flag is set, the value returned in phProv is undefined, and thus, the CryptReleaseContext function need not be called afterwards.
Private Const CRYPT_SILENT = &H40              ' The application requests that the CSP not display any user interface (UI) for this context. If the CSP must display the UI to operate, the call fails and the NTE_SILENT_CONTEXT error code is set as the last error. In addition, if calls are made to CryptGenKey with the CRYPT_USER_PROTECTED flag with a context that has been acquired with the CRYPT_SILENT flag, the calls fail and the CSP sets NTE_SILENT_CONTEXT.  CRYPT_SILENT is intended for use with applications for which the UI cannot be displayed by the CSP.  This flag is supported with Microsoft® Windows® 2000 or later. It is not supported in Windows 98 or Microsoft® Internet Explorer version 5.0.

'dwFlags definitions for CryptGenKey
Private Const CRYPT_EXPORTABLE         As Long = &H1
Private Const CRYPT_USER_PROTECTED     As Long = &H2
Private Const CRYPT_CREATE_SALT        As Long = &H4
Private Const CRYPT_UPDATE_KEY         As Long = &H8
Private Const CRYPT_NO_SALT            As Long = &H10
Private Const CRYPT_PREGEN             As Long = &H40
Private Const CRYPT_RECIPIENT          As Long = &H10
Private Const CRYPT_INITIATOR          As Long = &H40
Private Const CRYPT_ONLINE             As Long = &H80
Private Const CRYPT_SF                 As Long = &H100
Private Const CRYPT_CREATE_IV          As Long = &H200
Private Const CRYPT_KEK                As Long = &H400
Private Const CRYPT_DATA_KEY           As Long = &H800

Private Const AT_KEYEXCHANGE           As Long = 1
Private Const AT_SIGNATURE             As Long = 2
Private Const HP_HASHVAL               As Long = 2
Private Const SIMPLEBLOB               As Long = 1
Private Const PUBLICKEYBLOB            As Long = 6
Private Const PRIVATEKEYBLOB           As Long = 7
Private Const ALG_CLASS_DATA_ENCRYPT   As Long = 24576
Private Const ALG_CLASS_HASH           As Long = 32768
Private Const ALG_TYPE_ANY             As Long = 0
Private Const ALG_TYPE_STREAM          As Long = 2048
Private Const ALG_SID_RC4              As Long = 1
Private Const ALG_SID_MD5              As Long = 3
Private Const ALG_SID_SHA              As Long = 4
Private Const CALG_MD5 As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
Private Const CALG_RC4 As Long = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
Private Const CALG_SHA As Long = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_SHA)

Private Const NTE_BAD_UID                   As Long = &H80090001
Private Const NTE_BAD_HASH                  As Long = &H80090002
Private Const NTE_BAD_KEY                   As Long = &H80090003
Private Const NTE_BAD_LEN                   As Long = &H80090004
Private Const NTE_BAD_DATA                  As Long = &H80090005
Private Const NTE_BAD_SIGNATURE             As Long = &H80090006
Private Const NTE_BAD_VER                   As Long = &H80090007
Private Const NTE_BAD_ALGID                 As Long = &H80090008
Private Const NTE_BAD_FLAGS                 As Long = &H80090009
Private Const NTE_BAD_TYPE                  As Long = &H8009000A
Private Const NTE_BAD_KEY_STATE             As Long = &H8009000B
Private Const NTE_BAD_HASH_STATE            As Long = &H8009000C
Private Const NTE_NO_KEY                    As Long = &H8009000D
Private Const NTE_NO_MEMORY                 As Long = &H8009000E
Private Const NTE_EXISTS                    As Long = &H8009000F
Private Const NTE_PERM                      As Long = &H80090010
Private Const NTE_NOT_FOUND                 As Long = &H80090011
Private Const NTE_DOUBLE_ENCRYPT            As Long = &H80090012
Private Const NTE_BAD_PROVIDER              As Long = &H80090013
Private Const NTE_BAD_PROV_TYPE             As Long = &H80090014
Private Const NTE_BAD_PUBLIC_KEY            As Long = &H80090015
Private Const NTE_BAD_KEYSET                As Long = &H80090016
Private Const NTE_PROV_TYPE_NOT_DEF         As Long = &H80090017
Private Const NTE_PROV_TYPE_ENTRY_BAD       As Long = &H80090018
Private Const NTE_KEYSET_NOT_DEF            As Long = &H80090019
Private Const NTE_KEYSET_ENTRY_BAD          As Long = &H8009001A
Private Const NTE_PROV_TYPE_NO_MATCH        As Long = &H8009001B
Private Const NTE_SIGNATURE_FILE_BAD        As Long = &H8009001C
Private Const NTE_PROVIDER_DLL_FAIL         As Long = &H8009001D
Private Const NTE_PROV_DLL_NOT_FOUND        As Long = &H8009001E
Private Const NTE_BAD_KEYSET_PARAM          As Long = &H8009001F
Private Const NTE_FAIL                      As Long = &H80090020
Private Const NTE_SYS_ERR                   As Long = &H80090021
Private Const NTE_SILENT_CONTEXT            As Long = &H80090022
Private Const NTE_TOKEN_KEYSET_STORAGE_FULL As Long = &H80090023
Private Const NTE_TEMPORARY_PROFILE         As Long = &H80090024
Private Const NTE_FIXEDPARAMETER            As Long = &H80090025

Private Declare Function CryptAcquireContext _
    Lib "advapi32.dll" _
            Alias "CryptAcquireContextA" ( _
            phProv As Long, _
            pszContainer As Byte, _
            pszProvider As Byte, _
            ByVal dwProvType As Long, _
            ByVal dwFlags As Long) As Long

Private Declare Function CryptCreateHash _
    Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal Algid As Long, _
        ByVal hSessionKey As Long, _
        ByVal dwFlags As Long, _
        phHash As Long _
            ) As Long

Private Declare Function CryptDecrypt _
    Lib "advapi32.dll" ( _
        ByVal hSessionKey As Long, _
        ByVal hHash As Long, _
        ByVal Final As Long, _
        ByVal dwFlags As Long, _
        pbData As Byte, _
        pdwDataLen As Long _
            ) As Long

Private Declare Function CryptDeriveKey _
    Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal Algid As Long, _
        ByVal hBaseData As Long, _
        ByVal dwFlags As Long, _
        hSessionKey As Long _
            ) As Long

Private Declare Function CryptDestroyHash _
    Lib "advapi32.dll" ( _
        ByVal hHash As Long _
            ) As Long

Private Declare Function CryptDestroyKey _
    Lib "advapi32.dll" ( _
        ByVal hSessionKey As Long _
            ) As Long

Private Declare Function CryptEncrypt _
    Lib "advapi32.dll" ( _
        ByVal hSessionKey As Long, _
        ByVal hHash As Long, _
        ByVal Final As Long, _
        ByVal dwFlags As Long, _
        pbData As Byte, _
        pdwDataLen As Long, _
        ByVal dwBufLen As Long _
            ) As Long

Private Declare Function CryptEnumProviders _
    Lib "advapi32.dll" _
        Alias "CryptEnumProvidersA" ( _
            ByVal dwIndex As Long, _
            pdwReserved As Long, _
            ByVal dwFlags As Long, _
            pdwProvType As Long, _
            pszProvName As Byte, _
            pcbProvName As Long)

Private Declare Function CryptEnumProviderTypes _
    Lib "advapi32.dll" _
        Alias "CryptEnumProviderTypesA" ( _
            ByVal dwIndex As Long, _
            pdwReserved As Long, _
            ByVal dwFlags As Long, _
            pdwProvType As Long, _
            pszTypeName As Byte, _
            pcbTypeName As Long)

Private Declare Function CryptExportKey _
    Lib "advapi32.dll" ( _
        ByVal hSessionKey As Long, _
        ByVal hExpKey As Long, _
        ByVal dwBlobType As Long, _
        ByVal dwFlags As Long, _
        pbData As Byte, _
        pdwDataLen As Long _
            ) As Long

Private Declare Function CryptGenKey _
    Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal Algid As Long, _
        ByVal dwFlags As Long, _
        hSessionKey As Long _
            ) As Long

Private Declare Function CryptGenRandom _
    Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal dwLen As Long, _
        pbBuffer As Byte _
            ) As Long

Private Declare Function CryptGetDefaultProvider _
    Lib "advapi32.dll" _
        Alias "CryptGetDefaultProviderA" ( _
            ByVal dwProvType As Long, _
            pdwReserved As Long, _
            ByVal dwFlags As Long, _
            pszProvName As Byte, _
            pcbProvName As Long) As Long

Private Declare Function CryptGetHashParam _
    Lib "advapi32.dll" ( _
        ByVal hHash As Long, _
        ByVal dwParam As Long, _
        pbData As Byte, _
        pdwDataLen As Long, _
        ByVal dwFlags As Long _
            ) As Long

Private Declare Function CryptGetKeyParam _
    Lib "advapi32.dll" ( _
        ByVal hSessionKey As Long, _
        ByVal dwParam As Long, _
        pbData As Byte, _
        pdwDataLen As Long, _
        ByVal dwFlags As Long _
            ) As Long

Private Declare Function CryptGetProvParam _
    Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal dwParam As Long, _
        pbData As Byte, _
        pdwDataLen As Long, _
        ByVal dwFlags As Long) As Long

Private Declare Function CryptGetUserKey _
    Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal dwKeySpec As Long, _
        phUserKey As Long _
            ) As Long

Private Declare Function CryptImportKey _
    Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        pbData As Byte, _
        ByVal dwDataLen As Long, _
        ByVal hPubKey As Long, _
        ByVal dwFlags As Long, _
        phKey As Long _
            ) As Long

Private Declare Function CryptHashData _
    Lib "advapi32.dll" ( _
        ByVal hHash As Long, _
        pbData As Byte, _
        ByVal dwDataLen As Long, _
        ByVal dwFlags As Long _
            ) As Long

Private Declare Function CryptReleaseContext _
    Lib "advapi32.dll" ( _
        ByVal hProv As Long, _
        ByVal dwFlags As Long _
            ) As Long

Private Declare Function CryptSetHashParam _
    Lib "advapi32.dll" ( _
        ByVal hHash As Long, _
        ByVal dwParam As Long, _
        pbData As Byte, _
        ByVal dwFlags As Long _
            ) As Long

Private Declare Function CryptSetKeyParam _
    Lib "advapi32.dll" ( _
        ByVal hSessionKey As Long, _
        ByVal dwParam As Long, _
        pbData As Byte, _
        ByVal dwFlags As Long _
            ) As Long

Private Declare Function CryptSignHash _
    Lib "advapi32.dll" _
        Alias "CryptSignHashA" ( _
            ByVal hHash As Long, _
            ByVal dwKeySpec As Long, _
            ByVal sDescription As String, _
            ByVal dwFlags As Long, _
            pbSignature As Byte, _
            pdwSigLen As Long _
                ) As Long

Private Declare Function CryptVerifySignature Lib "advapi32.dll" Alias "CryptVerifySignatureA" (ByVal hHash As Long, pbSignature As Byte, ByVal dwSigLen As Long, ByVal hPubKey As Long, ByVal sDescription As String, ByVal dwFlags As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Dim bReturnStat As Boolean          'status after encrypt/decrypt
Dim lngType As Long                 'type of key in use (Exchange or Signature)
Dim hCryptProv As Long              'the handle to the CSP's Context (The Key Name)
Dim lngParams As Long               'the parameters to use for a generated key
Dim hKeyPair As Long                'the handle to the current key pair
Dim hSessionKey As Long             'the handle to the current session key
Dim hHash As Long                   'the handle to the current hash
Dim sLastErr As String              'last error code, converted to string desc

Dim abSerPro() As Byte              'byte array of value for Service Provider
Dim abKeyCon() As Byte              'byte array of value for Key Container/Context for CSP
Dim abPublicBlob() As Byte          'byte array (blob) of public key
Dim lngPubBlobLength As Long        'length of public key blob
Dim bolPubBlobType As Boolean       'type of public key blob, T = Exchange, F = Signature
Dim abPrivateBlob() As Byte         'byte array (blob) of private key
Dim lngPvtBlobLength As Long        'length of private key blob
Dim abSessionBlob() As Byte         'session key blob
Dim lngSessionBlobLength As Long    'session key blob length
Dim abSig() As Byte                 'signature blob
Dim lngSigLength As Long            'signature blob length
Dim iOSPlatform As Integer          'OS type, to determine Unicode functions
Private bolUnicode As Boolean       'Unicode flag, established on class init
Private asProviders() As String     'string array of Crypto API Providers names

'The WideCharToMultiByte function maps a wide-character string to a new character string.
'The function is faster when both lpDefaultChar and lpUsedDefaultChar are NULL.

'CodePage
Private Const CP_ACP = 0 'ANSI
Private Const CP_MACCP = 2 'Mac
Private Const CP_OEMCP = 1 'OEM
Private Const CP_UTF7 = 65000
Private Const CP_UTF8 = 65001

'dwFlags
Private Const WC_NO_BEST_FIT_CHARS = &H400
Private Const WC_COMPOSITECHECK = &H200
Private Const WC_DISCARDNS = &H10
Private Const WC_SEPCHARS = &H20 'Default
Private Const WC_DEFAULTCHAR = &H40

Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, _
                                                    ByVal dwFlags As Long, _
                                                    ByVal lpWideCharStr As Long, _
                                                    ByVal cchWideChar As Long, _
                                                    ByVal lpMultiByteStr As Long, _
                                                    ByVal cbMultiByte As Long, _
                                                    ByVal lpDefaultChar As Long, _
                                                    ByVal lpUsedDefaultChar As Long) As Long

Public Property Get DblByteStrings() As Boolean

    DblByteStrings = bolUnicode
End Property

Private Sub ServiceProviderName(Optional bolSrvPvdrBase As Boolean = True)

    If bolSrvPvdrBase Then
        ReDim abSerPro(Len(BASE_PROVIDER))
        If bolUnicode Then
            abSerPro = StrConv(BASE_PROVIDER, vbFromUnicode, 1033)
        Else
            abSerPro = BASE_PROVIDER
        End If
    Else
        ReDim abSerPro(Len(ENHANSED_PROVIDER))
        If bolUnicode Then
            abSerPro = StrConv(ENHANSED_PROVIDER, vbFromUnicode, 1033)
        Else
            abSerPro = ENHANSED_PROVIDER
        End If
    End If
End Sub

Public Property Let ValuePublicKey(strValue As String)  'Import a Public Key
' Currently, this function is not used.

Dim abData() As Byte
Dim sHeader As String
Dim lngW As Long
Dim lngX As Long
Dim lngY As Long

    bReturnStat = False
    If Len(strValue) < 3 Then Exit Property
    ' strip the string length
    lngX = InStr(2, strValue, "|")
    If lngX = 0 Or lngX > 4 Then Exit Property
    lngW = CLng(Mid(strValue, 2, lngX - 2))
    If lngW + lngX > Len(strValue) Then Exit Property
    sHeader = Mid(strValue, lngX + 1, lngW)
    
    lngX = InStr(2, sHeader, Chr(58))
    If lngX = 0 Or lngX > 4 Then Exit Property
    lngPubBlobLength = CLng(Mid(sHeader, 1, lngX - 1))
    If lngPubBlobLength + lngX > Len(sHeader) Then Exit Property
    sHeader = Mid(sHeader, lngX + 1, lngPubBlobLength)
    
    ' setvars
    If lngPubBlobLength < 1 Then Exit Property
    lngY = 0
    ReDim abPublicBlob(lngPubBlobLength - 1)
    ReDim abData(lngPubBlobLength - 1)
    If bolUnicode Then
        abData = StrConv(sHeader, vbFromUnicode, 1033)
    Else
        abData = strValue
    End If
    ReDim Preserve abData(lngPubBlobLength - 1)
    abPublicBlob = abData
    bReturnStat = True

End Property

Public Property Get ValuePublicKey() As String  'Export a Public Key
' Currently, this function is not used.

Dim abData() As Byte
Dim sHeader As String
Dim lngW As Long
Dim lngX As Long


If lngPubBlobLength > 0 Then
    If UBound(abPublicBlob) = lngPubBlobLength - 1 Then
        sHeader = CStr(lngPubBlobLength) & Chr(58)
        lngW = Len(sHeader)
        ReDim abData(lngW + lngPubBlobLength)
        For lngX = 0 To lngW - 1
            abData(lngX) = Asc(Mid(sHeader, lngX + 1, 1))
        Next lngX
        For lngX = 0 To lngPubBlobLength - 1
            abData(lngW + lngX) = abPublicBlob(lngX)
        Next lngX
        If bolUnicode Then
            sHeader = StrConv(abData, vbUnicode, 1033)
        Else
            sHeader = abData
        End If
        sHeader = Mid(sHeader, 1, lngW + lngPubBlobLength)
        ValuePublicKey = IIf(bolPubBlobType, "E", "S") & CStr(Len(sHeader)) & "|" & sHeader
    End If
End If
End Property

Public Property Let HexPublicKey(strValue As String)    'Import a Hex Public Key
Dim lngX As Long, strWorking As String

    bReturnStat = False
    If Len(strValue) < 3 Then Exit Property
    If Mid$(strValue, 1, 1) = "E" Then
        bolPubBlobType = True
    ElseIf Mid$(strValue, 1, 1) = "S" Then
        bolPubBlobType = False
    Else
        Exit Property
    End If
    strWorking = Mid$(strValue, 2)
    If Len(strWorking) Mod 2 <> 0 Then Exit Property
    ReDim abPublicBlob((Len(strWorking) \ 2) - 1)
    For lngX = 0 To UBound(abPublicBlob)
        abPublicBlob(lngX) = Chr(CInt("&h" & Mid$(strWorking, (lngX * 2) + 1, 2)))
    Next lngX
    lngPubBlobLength = lngX
    bReturnStat = True
End Property

Public Property Get HexPublicKey() As String    'Export a Public Key in Hex
Dim lngX As Long

    If lngPubBlobLength > 0 Then
        If UBound(abPublicBlob) = lngPubBlobLength - 1 Then
            HexPublicKey = IIf(bolPubBlobType, "E", "S")
            For lngX = 0 To lngPubBlobLength - 1
                HexPublicKey = HexPublicKey & Hex(Asc(abPublicBlob(lngX)))
            Next lngX
        End If
    End If
End Property

Public Property Let ValuePrivateKey(strKey As String)   'Import a Private Key
' Currently, this function is not used.

Dim abData() As Byte
Dim sHeader As String

bReturnStat = False
If Len(strKey) < 3 Then Exit Property
If Mid$(strKey, 1, 1) = "E" Then
    bolPubBlobType = True
ElseIf Mid$(strKey, 1, 1) = "S" Then
    bolPubBlobType = False
Else
    Exit Property
End If

' Get the blob length
lngPvtBlobLength = CLng(Mid$(strKey, 2, InStr(1, strKey, Chr(58), vbTextCompare) - 2))

If lngPvtBlobLength < 1 Then Exit Property

' Convert to array of bytes, converting from Unicode if necessary
If bolUnicode Then
    abData = StrConv(Mid$(strKey, InStr(1, strKey, Chr(58), vbTextCompare) + 1), vbFromUnicode, 1033)
Else
    abData = Mid$(strKey, InStr(1, strKey, Chr(58), vbTextCompare) + 1)
End If

abPrivateBlob = abData
bReturnStat = False

End Property

Public Property Get ValuePrivateKey() As String     'Export a Private Key.
' Currently, this function is not used.

Dim sHeader As String

If lngPvtBlobLength > 0 Then
    If UBound(abPrivateBlob) = lngPvtBlobLength - 1 Then
        'pack the length & key into the returned string
        sHeader = IIf(bolPubBlobType, "E", "S") & CStr(lngPvtBlobLength) & Chr(58) & CStr(abPrivateBlob)
        If bolUnicode Then _
            sHeader = IIf(bolPubBlobType, "E", "S") & CStr(lngPvtBlobLength) & Chr(58) & StrConv(abPrivateBlob, vbUnicode, 1033)
        ValuePrivateKey = sHeader
    End If
End If

End Property

Private Sub Class_Initialize()
Dim sValue As String
Dim aUniTest() As Byte

    sValue = "CRYPT     "
    aUniTest = StrConv(sValue, vbFromUnicode, 1033)
    If aUniTest(0) = 67 And aUniTest(1) = 82 And _
        aUniTest(2) = 89 And aUniTest(3) = 80 And aUniTest(4) = 84 Then _
            bolUnicode = True
    ServiceProviderName False
End Sub

Private Sub Class_Terminate()

    If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
    If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
    If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0
End Sub

Public Function SessionStart(sKeyCon As String, Optional bFresh As Boolean = False) As Boolean
Dim lngReturnValue As Long

    ReDim abPublicBlob(0)
    SessionStart = False
    If bolUnicode Then
        abKeyCon = StrConv(sKeyCon, vbFromUnicode, 1033) & vbNullChar
        abSerPro = StrConv(BASE_PROVIDER, vbFromUnicode, 1033) & vbNullChar
        'abSerPro = StrConv(ENHANSED_PROVIDER, vbFromUnicode, 1033) & vbNullChar
    Else
        ReDim abKeyCon(Len(sKeyCon))
        ReDim abSerPro(Len(BASE_PROVIDER))
        'ReDim abSerPro(Len(ENHANSED_PROVIDER))
        abKeyCon = sKeyCon & vbNullChar
        abSerPro = BASE_PROVIDER & vbNullChar
        'abSerPro = ENHANSED_PROVIDER & vbNullChar
    End If
    If bFresh Then
        'kill the old key container
        lngReturnValue = CryptAcquireContext( _
            hCryptProv, abKeyCon(0), abSerPro(0), PROV_RSA_FULL, CRYPT_DELETEKEYSET Or CRYPT_MACHINE_KEYSET _
           )
        If lngReturnValue = 0 Then _
            Err.Clear 'The container didn't exist...so what's new?
        'create a handle to the new key container name...
        lngReturnValue = CryptAcquireContext( _
            hCryptProv, abKeyCon(0), abSerPro(0), PROV_RSA_FULL, CRYPT_NEWKEYSET Or CRYPT_MACHINE_KEYSET Or CRYPT_SILENT _
           )
        'and if the new key container fails...raise the error
        If lngReturnValue = 0 Then
            Err.Raise Err.LastDllError, Err.Source, _
            "DLL error code shown above.  " & vbCrLf & _
            "Error during CryptAcquireContext for key container: " & sKeyCon
        End If
        'otherwise things are fine
        SessionStart = True
    Else
        'get the named key container
        lngReturnValue = CryptAcquireContext( _
            hCryptProv, abKeyCon(0), abSerPro(0), PROV_RSA_FULL, CRYPT_EXISTINGKEYSET Or CRYPT_MACHINE_KEYSET _
           )
        'here's a juncture that requires your attention.
        'You have a choice of stopping, by raising the above error,
        '    because the container does not exist,
        '    or, you can go on, which means,
        '    in order for this function to not fail,
        '    you must create a new key container.
        'This code is written to continue if the first AquireContext fails
        '    ...creating a new key container
        If lngReturnValue = 0 Then
            Err.Clear
            lngReturnValue = CryptAcquireContext( _
                hCryptProv, abKeyCon(0), abSerPro(0), PROV_RSA_FULL, CRYPT_NEWKEYSET Or CRYPT_MACHINE_KEYSET Or CRYPT_SILENT _
               )
            'and if the new key container fails...raise the error
            If lngReturnValue = 0 Then
                Err.Raise Err.LastDllError, Err.Source, _
                "DLL error code shown above.  " & vbCrLf & _
                "Error during CryptAcquireContext for key container: " & sKeyCon
            End If
        End If
        'otherwise things are fine
        SessionStart = True
    End If
End Function

Public Function VerifyContainer(sKeyCon As String, Optional SessionEndOnExit As Boolean = True) As Boolean
Dim lngReturnValue As Long

    ReDim abPublicBlob(0)
    If bolUnicode Then
        abKeyCon = StrConv(sKeyCon, vbFromUnicode, 1033) & vbNullChar
        abSerPro = StrConv(BASE_PROVIDER, vbFromUnicode, 1033) & vbNullChar
        'abSerPro = StrConv(ENHANSED_PROVIDER, vbFromUnicode, 1033) & vbNullChar
    Else
        ReDim abKeyCon(Len(sKeyCon))
        ReDim abSerPro(Len(BASE_PROVIDER))
        'ReDim abSerPro(Len(ENHANSED_PROVIDER))
        abKeyCon = sKeyCon & vbNullChar
        abSerPro = BASE_PROVIDER & vbNullChar
        'abSerPro = ENHANSED_PROVIDER & vbNullChar
    End If
    'get the named key container
    lngReturnValue = CryptAcquireContext( _
        hCryptProv, abKeyCon(0), abSerPro(0), PROV_RSA_FULL, CRYPT_EXISTINGKEYSET Or CRYPT_MACHINE_KEYSET _
       )
    If lngReturnValue = 0 Then
        Err.Clear
        Exit Function
    End If
    'otherwise the container exists
    VerifyContainer = True
    If SessionEndOnExit Then Call SessionEnd
End Function

Public Sub SessionEnd()

    If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
    If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
    If hCryptProv <> 0 Then CryptReleaseContext hCryptProv, 0
End Sub

Public Function Generate(Optional UseExchangeKeyPair As Boolean = True, Optional m_lKeyPairModulus As Long = 1024) As Boolean

Dim lngReturnValue As Long, dwFlags As Long
dwFlags = 2 ^ 24 * (m_lKeyPairModulus \ 256) + 2 ^ 16 * (m_lKeyPairModulus Mod 256)

''''''''''''''''''''''''''''''''''''''''''
' Low bytes of dwFlags are various flags:
''''''''''''''''''''''''''''''''''''''''''
lngParams = CRYPT_EXPORTABLE Or CRYPT_NO_SALT Or dwFlags
'lngParams = &H2000000 Or CRYPT_EXPORTABLE
lngType = AT_KEYEXCHANGE
If Not UseExchangeKeyPair Then lngType = AT_SIGNATURE
If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
lngReturnValue = CryptGenKey(hCryptProv, lngType, lngParams, hKeyPair)
If lngReturnValue <> 0 Then
    Generate = True
Else
    Err.Clear
End If

End Function

Public Function DestroyContainer(sKeyCon As String) As Boolean
Dim lngReturnValue As Long

    ReDim abPublicBlob(0)
    'kill the named key container
    lngReturnValue = CryptAcquireContext( _
        hCryptProv, abKeyCon(0), abSerPro(0), PROV_RSA_FULL, CRYPT_DELETEKEYSET Or CRYPT_MACHINE_KEYSET _
       )
    If lngReturnValue = 0 Then 'The container didn't exist
        Err.Clear
        Exit Function
    End If
    DestroyContainer = True
End Function

Public Sub ImportPublicKey(Optional ByVal UseExchangeKeyPair As Boolean = True)
Dim lngReturnValue As Long
Dim lngLength As Long
Dim abKeyBlob() As Byte

    lngType = AT_KEYEXCHANGE
    If Not UseExchangeKeyPair Then lngType = AT_SIGNATURE
    
    ' Added by Ismail
    Dim b64 As clsCryptoAPIBase64
    Set b64 = New clsCryptoAPIBase64
    abPublicBlob = b64.Decode("BgIAAAAkAABSU0ExAAQAAAEAAQCBcwKp9p1rkQhZyxTeREh9EM273wBqpODS+KLkeu/xn/Q0+w8uhBQfZq8f/sdRfL+S5LIBItOv0okG42mKcaNk0mRoSoJkUPMrRp43j8nAKVCmRrD7pZ1Do3uHM4SjydLY0omeK8vOCyZ2WldYy0IxwgQjNqMHLuG1rCg1DR4+5g==")
    lngPubBlobLength = Len(ByteArrayToString(abPublicBlob))
    ' Added by Ismail
    
    If Not (lngPubBlobLength > 0 And UBound(abPublicBlob) = lngPubBlobLength - 1) Then
        Err.Raise vbObjectError + 1, , "The Public blob must hold a valid key"
    End If
    ReDim abKeyBlob(lngPubBlobLength)
    lngLength = lngPubBlobLength
    abKeyBlob = abPublicBlob
    
    If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
    lngReturnValue = CryptImportKey( _
        hCryptProv, abKeyBlob(0), lngLength, 0, lngType, hKeyPair)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. " & _
        "Could not import the Public key"
    If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
End Sub

Public Sub ImportPrivateKey(sPassPhrase As String, Optional ByVal UseExchangeKeyPair As Boolean = True)
Dim lngReturnValue As Long
Dim lngLength As Long
Dim lngD As Long
Dim abKeyBlob() As Byte
Dim abPassPhrase() As Byte
Dim abText() As Byte
Dim lLen As Long

    lngLength = 0
    lngType = AT_SIGNATURE Or CRYPT_EXPORTABLE Or CRYPT_NO_SALT
    If UseExchangeKeyPair Then _
        lngType = AT_KEYEXCHANGE Or CRYPT_EXPORTABLE Or CRYPT_NO_SALT
    
    ' added by Ismail
    Dim b64 As clsCryptoAPIBase64
    Set b64 = New clsCryptoAPIBase64
    abPrivateBlob = b64.Decode("BwIAAAAkAAD4eEFtuTwbjhCgyJLhAzMsX5uF75r6vP5ISv3eKF8ak8NmCp9YxB9LiMyuxvZkl6BSohXROUeRW4CvQzPNO7rnV6wrRFOCn0xX2RVKt3kh4Ers6WImC66l6QWjpotVzVIu6/3pa1zbw7DyixpCuHUhm2Vvt1wAs7EY+tEj4Sh4PYrM3GRJxbp6/7ff5Yi0hWdVntx64HI4K+YABG7ZAj3PEE2AXA4AB7Ghx7Fuqxxm3BNS5xy6XiLJw6FU0lLIe/K4vmqXR1nsQ5vLaHBDCVs3qf2IN9J2N/TgfrFbLRRensgLiGmI1e3POkpDq7IGP4SgNA2XqVPPj0eOqKWSdVJy6v11AuDQuBvischmtE63cg0L/S+qUUr77NLrMb8kH3Bf1as7Y5mCBXqF9W3sFBKG1bHkXMhWlYzEJSxldHA+9velfxsgtJAlfHRL+ItxRzUhUHqTLgDMX19xSXM4E9xNLbAI5mUJluzd7avkDlpOxdPTkP0MeaHC33b4AoQUN/AdTkcF+Cs/Y651kwMQOdY/DaD+rBCDdx3JNRcNzoQL+JFEqii1pDyXiW2Rqv3cNv6JRn7JrZMo1Hn5kRYOHDA/E4cIJdE1xMHu3pGLiuYX7Uu6U9yz/u9H6JASaJnZ8y0K4Atm1igVtl+wxD1tOvoeHcJ2JUTdmS12ae4f40c2uwhv/oD8Ecu6Oc1OuKIN+dlRHuYwvUfzBrcU7BqD63n215P5c5REEbw3ETItv7BEaHHpZ9RaLLjX73XCoXKTb5MV8cgtqLE7MxEqbB4=")
    lngPvtBlobLength = Len(ByteArrayToString(abPrivateBlob))
    ' added by Ismail
    
    If Not (lngPvtBlobLength > 0 And UBound(abPrivateBlob) = lngPvtBlobLength - 1) Or _
            sPassPhrase = vbNullString Then
        Err.Raise vbObjectError + 1, , "The Private blob must hold a valid key and Passphrase parameter must not be null"
    End If
    
    'first release old handles
    If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
    If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
    If hHash <> 0 Then CryptDestroyHash hHash
    
    abPassPhrase = sPassPhrase
    If bolUnicode Then abPassPhrase = StrConv(sPassPhrase, vbFromUnicode, 1033)
    
    'create the Hash Object
    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    If lngReturnValue = 0 Then Exit Sub
    
    'Hash the passphrase
    lngD = UBound(abPassPhrase) + 1
    lngReturnValue = CryptHashData(hHash, abPassPhrase(0), lngD, 0)
    If lngReturnValue = 0 Then Exit Sub
    'Hash the passphrase again
    lngD = UBound(abPassPhrase) + 1
    lngReturnValue = CryptHashData(hHash, abPassPhrase(0), lngD, 0)
    If lngReturnValue = 0 Then Exit Sub
    
    'create Session key from double Hash
    lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, 0, hSessionKey)
    If lngReturnValue = 0 Then Exit Sub
    lngReturnValue = CryptDestroyHash(hHash)
    If lngReturnValue = 0 Then Exit Sub
    
    lngLength = lngPvtBlobLength
    abKeyBlob = abPrivateBlob
    lngReturnValue = CryptImportKey( _
        hCryptProv, abKeyBlob(0), lngLength, hSessionKey, lngType, hKeyPair)
    If lngReturnValue = 0 Then Exit Sub
    If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
End Sub

Public Sub ExportPublicKey(Optional ByVal UseExchangeKeyPair As Boolean = True)
Dim lngReturnValue As Long
Dim lngLength As Long

    lngLength = 0
    If UseExchangeKeyPair Then lngType = AT_KEYEXCHANGE _
                          Else lngType = AT_SIGNATURE
    
    bolPubBlobType = UseExchangeKeyPair
    
    'first release old key pair handle
    If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
    
    'get a handle to the key pair
    lngReturnValue = CryptGetUserKey(hCryptProv, lngType, hKeyPair)
    If lngReturnValue = 0 Then
        Debug.Print Err.LastDllError, Err.Description
        Exit Sub
    End If
    
    'get the size of the buffer needed for the Public BLOB
    lngReturnValue = CryptExportKey(hKeyPair, 0, PUBLICKEYBLOB, 0, vbNull, lngLength)
    ReDim abPublicBlob(lngLength - 1)
    'get the Public Key Blob
    lngReturnValue = CryptExportKey(hKeyPair, 0, PUBLICKEYBLOB, 0, abPublicBlob(0), lngLength)
    If lngReturnValue = 0 Then
        Debug.Print Err.LastDllError, Err.Description
        Exit Sub
    End If
    ReDim Preserve abPublicBlob(lngLength - 1)
    lngPubBlobLength = lngLength
    
    ' Added by Ismail
    Dim b64 As clsCryptoAPIBase64
    Set b64 = New clsCryptoAPIBase64
    publicPublicKeyBlob = b64.Encode(abPublicBlob)
    ' "BgIAAAAkAABSU0ExAAQAAAEAAQCBcwKp9p1rkQhZyxTeREh9EM273wBqpODS+KLkeu/xn/Q0+w8uhBQfZq8f/sdRfL+S5LIBItOv0okG42mKcaNk0mRoSoJkUPMrRp43j8nAKVCmRrD7pZ1Do3uHM4SjydLY0omeK8vOCyZ2WldYy0IxwgQjNqMHLuG1rCg1DR4+5g=="
    ' Added by Ismail

End Sub


Public Function StringToByteArray(strInput As String, _
                                Optional bReturnAsUnicode As Boolean = True, _
                                Optional bAddNullTerminator As Boolean = False) As Byte()
    
    Dim lRet As Long
    Dim bytBuffer() As Byte
    Dim lLenB As Long
    
    If bReturnAsUnicode Then
        'Number of bytes
        lLenB = LenB(strInput)
        'Resize buffer, do we want terminating null?
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        'Copy characters from string to byte array
        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
    Else
        'METHOD ONE
'        'Get rid of embedded nulls
'        strRet = StrConv(strInput, vbFromUnicode, 1033)
'        lLenB = LenB(strRet)
'        If bAddNullTerminator Then
'            ReDim bytBuffer(lLenB)
'        Else
'            ReDim bytBuffer(lLenB - 1)
'        End If
'        CopyMemory bytBuffer(0), ByVal StrPtr(strInput), lLenB
        
        'METHOD TWO
        'Num of characters
        lLenB = Len(strInput)
        If bAddNullTerminator Then
            ReDim bytBuffer(lLenB)
        Else
            ReDim bytBuffer(lLenB - 1)
        End If
        lRet = WideCharToMultiByte(CP_ACP, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(bytBuffer(0)), lLenB, 0&, 0&)
    End If
    
    StringToByteArray = bytBuffer
    
End Function
Public Function ByteArrayToString(Bytes() As Byte) As String
    Dim iUnicode As Long, i As Long, j As Long
    
    On Error Resume Next
    i = UBound(Bytes)
    
    If (i < 1) Then
        'ANSI, just convert to unicode and return
        ByteArrayToString = StrConv(Bytes, vbUnicode, 1033)
        Exit Function
    End If
    i = i + 1
    
    'Examine the first two bytes
    CopyMemory iUnicode, Bytes(0), 2
    
    If iUnicode = Bytes(0) Then 'Unicode
        'Account for terminating null
        If (i Mod 2) Then i = i - 1
        'Set up a buffer to receieve the string
        ByteArrayToString = String$(i / 2, 0)
        'Copy to string
        CopyMemory ByVal StrPtr(ByteArrayToString), Bytes(0), i
    Else 'ANSI
        ByteArrayToString = StrConv(Bytes, vbUnicode, 1033)
    End If
                    
End Function
Public Sub ExportPrivateKey(sPassPhrase As String, Optional ByVal UseExchangeKeyPair As Boolean = True)
Dim abPassPhrase() As Byte
Dim abText() As Byte
Dim lLen As Long
Dim lngD As Long
Dim lngReturnValue As Long
Dim lngLength As Long
    
    lngLength = 0
    If UseExchangeKeyPair Then lngType = AT_KEYEXCHANGE _
                          Else lngType = AT_SIGNATURE
    
    'first release old handles
    If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
    If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
    If hHash <> 0 Then CryptDestroyHash hHash
    
    'get a handle to the key pair
    lngReturnValue = CryptGetUserKey(hCryptProv, lngType, hKeyPair)
    If lngReturnValue = 0 Then
        Debug.Print Err.LastDllError, Err.Description
        Exit Sub
    End If

    abPassPhrase = sPassPhrase
    If bolUnicode Then abPassPhrase = StrConv(sPassPhrase, vbFromUnicode, 1033)
    
    'create the Hash Object
    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    If lngReturnValue = 0 Then
        Debug.Print Err.LastDllError, Err.Description
        Exit Sub
    End If
    
    'Hash the passphrase
    lngD = UBound(abPassPhrase) + 1
    lngReturnValue = CryptHashData(hHash, abPassPhrase(0), lngD, 0)
    If lngReturnValue = 0 Then Exit Sub
    'Hash the passphrase again
    lngD = UBound(abPassPhrase) + 1
    lngReturnValue = CryptHashData(hHash, abPassPhrase(0), lngD, 0)
    If lngReturnValue = 0 Then
        Debug.Print Err.LastDllError, Err.Description
        Exit Sub
    End If
    
    'create Session key from double Hash
    lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, 0, hSessionKey)
    If lngReturnValue = 0 Then
        Debug.Print Err.LastDllError, Err.Description
        Exit Sub
    End If
    lngReturnValue = CryptDestroyHash(hHash)
    If lngReturnValue = 0 Then
        Debug.Print Err.LastDllError, Err.Description
        Exit Sub
    End If
    
    'get the size of the buffer needed for the private BLOB
    lngReturnValue = CryptExportKey(hKeyPair, hSessionKey, PRIVATEKEYBLOB, 0, vbNull, lngLength)
    ReDim abPrivateBlob(lngLength - 1)
    
    'get the Private Blob
    lngReturnValue = CryptExportKey(hKeyPair, hSessionKey, PRIVATEKEYBLOB, 0, abPrivateBlob(0), lngLength)
    If lngReturnValue = 0 Then
        Debug.Print Err.LastDllError, Err.Description
        Exit Sub
    End If
    ReDim Preserve abPrivateBlob(lngLength - 1)
    lngPvtBlobLength = lngLength
    
    ' Added by Ismail
    Dim b64 As clsCryptoAPIBase64
    Set b64 = New clsCryptoAPIBase64
    publicPrivateKeyBlob = b64.Encode(abPrivateBlob)
    ' "BwIAAAAkAAD4eEFtuTwbjhCgyJLhAzMsX5uF75r6vP5ISv3eKF8ak8NmCp9YxB9LiMyuxvZkl6BSohXROUeRW4CvQzPNO7rnV6wrRFOCn0xX2RVKt3kh4Ers6WImC66l6QWjpotVzVIu6/3pa1zbw7DyixpCuHUhm2Vvt1wAs7EY+tEj4Sh4PYrM3GRJxbp6/7ff5Yi0hWdVntx64HI4K+YABG7ZAj3PEE2AXA4AB7Ghx7Fuqxxm3BNS5xy6XiLJw6FU0lLIe/K4vmqXR1nsQ5vLaHBDCVs3qf2IN9J2N/TgfrFbLRRensgLiGmI1e3POkpDq7IGP4SgNA2XqVPPj0eOqKWSdVJy6v11AuDQuBvischmtE63cg0L/S+qUUr77NLrMb8kH3Bf1as7Y5mCBXqF9W3sFBKG1bHkXMhWlYzEJSxldHA+9velfxsgtJAlfHRL+ItxRzUhUHqTLgDMX19xSXM4E9xNLbAI5mUJluzd7avkDlpOxdPTkP0MeaHC33b4AoQUN/AdTkcF+Cs/Y651kwMQOdY/DaD+rBCDdx3JNRcNzoQL+JFEqii1pDyXiW2Rqv3cNv6JRn7JrZMo1Hn5kRYOHDA/E4cIJdE1xMHu3pGLiuYX7Uu6U9yz/u9H6JASaJnZ8y0K4Atm1igVtl+wxD1tOvoeHcJ2JUTdmS12ae4f40c2uwhv/oD8Ecu6Oc1OuKIN+dlRHuYwvUfzBrcU7BqD63n215P5c5REEbw3ETItv7BEaHHpZ9RaLLjX73XCoXKTb5MV8cgtqLE7MxEqbB4="
    ' Added by Ismail
End Sub

Public Sub EncryptFile(SourceFile As String, DestFile As String)
Dim arrByteArray() As Byte
Dim abSKey() As Byte
Dim lngFileSize As Long
Dim intNextFreeFile As Integer
Dim sHead As String

    intNextFreeFile = FreeFile
    Open SourceFile For Binary As intNextFreeFile
    lngFileSize = LOF(intNextFreeFile)
    If lngFileSize Then
        ReDim arrByteArray(lngFileSize - 1)
        Get intNextFreeFile, , arrByteArray()
        Close intNextFreeFile
        EncryptDecrypt arrByteArray, abSKey, True
        On Error Resume Next
        Kill DestFile
        On Error GoTo 0
        sHead = CStr(lngFileSize) & ":" & CStr(UBound(arrByteArray)) & ":"
        ReDim abSKey(Len(sHead) - 1)
        For intNextFreeFile = 0 To Len(sHead) - 1
            abSKey(intNextFreeFile) = Asc(Mid(sHead, intNextFreeFile + 1, 1))
        Next intNextFreeFile
        intNextFreeFile = FreeFile
        Open DestFile For Binary As intNextFreeFile
        Put intNextFreeFile, , abSKey()
        Put intNextFreeFile, , arrByteArray()
        Close intNextFreeFile
    End If
End Sub

Public Function DecryptFile(SourceFile As String, DestFile As String)
Dim baCText() As Byte
Dim baPlain() As Byte
Dim sHead As String
Dim lngW As Long
Dim lngX As Long
Dim lngY As Long
Dim lFS As Long
Dim lAS As Long
Dim iFF As Integer

    iFF = FreeFile
    Open SourceFile For Binary Lock Read Write As iFF
    lFS = LOF(iFF)
    ReDim baCText(lFS - 1)
    Get iFF, , baCText()
    Close iFF
    
    sHead = vbNullString
    lngW = lFS - 1
    If lngW < 11 Then Exit Function
    For lngX = 0 To lngW ' Get the length of the source file
        If baCText(lngX) >= 48 And baCText(lngX) <= 58 Then
            If baCText(lngX) = 58 Then Exit For
            sHead = sHead & Chr(baCText(lngX))
        End If
    Next lngX
    If Len(sHead) > 0 And baCText(lngX) = 58 Then
        lFS = CLng(sHead)
        sHead = vbNullString
        lngY = lngX + 1 ' start of encrypted array length
        For lngX = lngY To lngW ' Get the length of the encrypted array
            If baCText(lngX) >= 48 And baCText(lngX) <= 58 Then
                If baCText(lngX) = 58 Then Exit For
                sHead = sHead & Chr(baCText(lngX))
            End If
        Next lngX
        If Len(sHead) > 0 And baCText(lngX) = 58 Then
            lAS = CLng(sHead)
            lngY = lngX + 1 ' start of encrypted array
            If lngW - lngX < lAS Then Exit Function ' Not enough data
            For lngX = lngY To lngW ' Move the encrypted data to the start of array and resize
                baCText(lngX - lngY) = baCText(lngX)
            Next lngX
            ReDim Preserve baCText(lAS)
        Else
            ' Invalid file format
        End If
    Else
        ' Invalid file format
    End If
    If bolUnicode Then
        sHead = StrConv(baCText, vbUnicode, 1033)
    Else
        sHead = baCText
    End If
    EncryptDecrypt baCText, baPlain, False
    If bolUnicode Then
        sHead = StrConv(baPlain, vbUnicode, 1033)
    Else
        sHead = baPlain
    End If
    On Error Resume Next
    Kill DestFile
    On Error GoTo 0
    iFF = FreeFile
    Open DestFile For Binary As iFF
    Put iFF, , baPlain()
    Close iFF
End Function

Public Function EncryptString2File(sPlain As String, sDestinationFile As String) As Boolean
Dim baCypher() As Byte
Dim baPlain() As Byte
Dim lndx As Long
Dim iFF As Integer
Dim sHead As String

    EncryptString2File = False
    ReDim baCypher(Len(sPlain) - 1)
    If bolUnicode Then baPlain() = StrConv(sPlain, vbFromUnicode, 1033) _
        Else baPlain() = sPlain
    EncryptDecrypt baPlain(), baCypher(), True
    sHead = CStr(UBound(baPlain)) & ":" & CStr(UBound(baCypher)) & ":"
    ReDim baPlain(Len(sHead) - 1)
    For lndx = 0 To Len(sHead) - 1
        baPlain(lndx) = Asc(Mid(sHead, lndx + 1, 1))
    Next lndx
    On Error Resume Next
    Kill sDestinationFile
    On Error GoTo 0
    iFF = FreeFile
    Open sDestinationFile For Binary As #iFF
    Put iFF, , baPlain()
    Put iFF, , baCypher()
    Close iFF
    EncryptString2File = True
End Function

Public Function DecryptFile2String(SourceFile As String, ByRef sPlain As String) As Boolean
Dim baCText() As Byte
Dim baPlain() As Byte
Dim sHead As String
Dim lngW As Long
Dim lngX As Long
Dim lngY As Long
Dim lFS As Long
Dim lAS As Long
Dim iFF As Integer

    'read the encrypted source file
    DecryptFile2String = False
    iFF = FreeFile
    Open SourceFile For Binary Lock Read Write As iFF
    lFS = LOF(iFF)
    If lFS < 1 Then
        Close iFF
        Exit Function
    End If
    ReDim baCText(lFS - 1)
    Get iFF, , baCText()
    Close iFF
    
    sHead = vbNullString
    lngW = lFS - 1
    If lngW < 11 Then Exit Function
    For lngX = 0 To lngW ' Get the length of the plain text
        If baCText(lngX) >= 48 And baCText(lngX) <= 58 Then
            If baCText(lngX) = 58 Then Exit For
            sHead = sHead & Chr(baCText(lngX))
        End If
    Next lngX
    If Len(sHead) > 0 And baCText(lngX) = 58 Then
        lFS = CLng(sHead)
        sHead = vbNullString
        lngY = lngX + 1 ' start of encrypted array length
        For lngX = lngY To lngW ' Get the length of the encrypted array
            If baCText(lngX) >= 48 And baCText(lngX) <= 58 Then
                If baCText(lngX) = 58 Then Exit For
                sHead = sHead & Chr(baCText(lngX))
            End If
        Next lngX
        If Len(sHead) > 0 And baCText(lngX) = 58 Then
            lAS = CLng(sHead)
            lngY = lngX + 1 ' start of encrypted array
            If lngW - lngX < lAS Then Exit Function ' Not enough data
            For lngX = lngY To lngW ' Move the encrypted data to the start of array and resize
                baCText(lngX - lngY) = baCText(lngX)
            Next lngX
            ReDim Preserve baPlain(lAS)
        Else
            ' Invalid file format
            Exit Function
        End If
    Else
        ' Invalid file format
        Exit Function
    End If
    EncryptDecrypt baPlain(), baCText(), False
    
    'resize the return string (byref)
    sPlain = Space(UBound(baPlain) + 1)
    If bolUnicode Then sPlain = StrConv(baPlain, vbUnicode, 1033) _
        Else sPlain = baPlain

    DecryptFile2String = True
End Function

Public Function EncryptString(sText As String) As String
Dim abData() As Byte
Dim abCyph() As Byte

    ReDim abData(Len(sText) - 1)
    If bolUnicode Then abData = StrConv(sText, vbFromUnicode, 1033) _
        Else abData = sText
    
    EncryptDecrypt abData, abCyph, True
    
    If UBound(abCyph) > 0 Then
        If bolUnicode Then EncryptString = StrConv(abCyph, vbUnicode, 1033) _
            Else EncryptString = abCyph
    Else
        EncryptString = vbNullString
    End If
End Function

Public Function DecryptString(sText As String) As String
Dim abData() As Byte
Dim abCyph() As Byte

    ReDim abCyph(Len(sText) - 1)
    If bolUnicode Then abCyph = StrConv(sText, vbFromUnicode, 1033) _
                      Else abCyph = sText

    EncryptDecrypt abData, abCyph, False

    If UBound(abData) > 0 Then
        If bolUnicode Then DecryptString = StrConv(abData, vbUnicode, 1033) _
            Else DecryptString = abData
    Else
        DecryptString = vbNullString
    End If
End Function

Private Sub EncryptDecrypt(ByRef abData() As Byte, ByRef abCText() As Byte, Encrypt As Boolean)
Dim abArray() As Byte
Dim lngLength As Long
Dim lngBuffer As Long
Dim lngReturnValue As Long
Dim sHeader As String
Dim sBuffer As String
Dim lngW As Long
Dim lngX As Long
Dim lngY As Long

    'check provider handle
    If hCryptProv <> 0 Then 'continue
    Else
        Exit Sub
    End If
    'release old key pair handle
    If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
    'release old session key handle
    If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
    
    'get a handle to the exchange key pair
    lngReturnValue = CryptGetUserKey(hCryptProv, AT_KEYEXCHANGE, hKeyPair)
        If lngReturnValue = 0 Then
            sLastErr = DecodeLastError(Err.LastDllError)
            If sLastErr = vbNullString Then sLastErr = CStr(Err.LastDllError) & Err.Description
            Debug.Print "DLL error: " & sLastErr & " Could not obtain handle to key pair"
            Err.Clear
            Exit Sub
        End If

    If Encrypt Then
        ' Generate a random Session key
        lngReturnValue = CryptGenKey(hCryptProv, CALG_RC4, CRYPT_EXPORTABLE, hSessionKey)
        If lngReturnValue = 0 Then _
            Debug.Print CStr(Err.LastDllError) & _
                " Could not create session key for use in key pair encryption"

        ' Export the Session Key Size (pass Null as the key blob and length is returned)
        lngReturnValue = CryptExportKey( _
            hSessionKey, hKeyPair, SIMPLEBLOB, 0, vbNull, lngLength)
        If lngLength < 1 Then
            Debug.Print CStr(Err.LastDllError) & _
                " Could not export session key size for use in export"
            Exit Sub
        End If
        ' this next line is because the last call didn't complete it's function
        ' it was suppose to return the key...all it did was return the size
        ' so you have to clear the Err.
        If lngReturnValue = 0 Then Err.Clear
        ReDim abSessionBlob(lngLength - 1)

        ' Export & (using key pair) encrypt the Session key
        lngReturnValue = CryptExportKey( _
            hSessionKey, hKeyPair, SIMPLEBLOB, 0, abSessionBlob(0), lngLength)
        If lngReturnValue = 0 Then _
            Debug.Print CStr(Err.LastDllError) & _
                "Could not export the session key during key pair encryption"
        ' the encrypted Session key is now available
        
        ' Pack the key length & ":" & the session key to abCText
        sHeader = CStr(lngLength) & Chr(58)
        lngW = Len(sHeader)
        ReDim abCText(lngW + lngLength - 1)
        lngX = UBound(abCText)
        For lngY = 0 To lngW - 1
            abCText(lngY) = Asc(Mid(sHeader, lngY + 1, 1))
        Next lngY
        For lngY = lngW To lngX
            abCText(lngY) = abSessionBlob(lngY - lngW)
        Next lngY
        lngSessionBlobLength = UBound(abCText) + 1
        
        ' Encrypt the data
        bReturnStat = False
        ReDim abArray(UBound(abData))
        lngLength = UBound(abData) + 1
        lngBuffer = lngLength
        lngReturnValue = CryptEncrypt(hSessionKey, 0, 1, 0, abData(0), lngLength, lngBuffer)
        If lngReturnValue = 0 Or lngBuffer <> lngLength Then _
            Debug.Print CStr(Err.LastDllError) & "Error during encryption." _
            Else bReturnStat = True
        
        ' Pack the length of the data & ":" & the encrypted data, to abCText
        sHeader = CStr(lngLength) & Chr(58)
        lngW = Len(sHeader)
        ReDim Preserve abCText(lngSessionBlobLength + lngW + lngLength - 1)
        lngX = UBound(abCText)
        For lngY = lngSessionBlobLength To lngSessionBlobLength + lngW - 1
            abCText(lngY) = Asc(Mid(sHeader, lngY - lngSessionBlobLength + 1, 1))
        Next lngY
        For lngY = lngSessionBlobLength + lngW To lngX
            abCText(lngY) = abData(lngY - lngSessionBlobLength - lngW)
        Next lngY
    Else
        ' Extract the parameters needed for decryption...
        sHeader = vbNullString
        lngW = UBound(abCText)
        For lngX = 0 To lngW ' Get the length of the session key
            If abCText(lngX) >= 48 And abCText(lngX) <= 58 Then
                If abCText(lngX) = 58 Then Exit For
                sHeader = sHeader & Chr(abCText(lngX))
            End If
        Next lngX
        ' Extract the session key
        If Len(sHeader) > 0 And abCText(lngX) = 58 Then
            lngX = lngX + 1 ' start byte of SessionBlob
            lngSessionBlobLength = CLng(sHeader)
            If lngSessionBlobLength And lngW > lngX + lngSessionBlobLength - 1 Then
                ReDim abSessionBlob(lngSessionBlobLength - 1)
                For lngY = lngX To lngX + lngSessionBlobLength - 1
                    abSessionBlob(lngY - lngX) = abCText(lngY)
                Next lngY
            End If
            lngY = lngX + lngSessionBlobLength ' start byte of cryptoText length
            ' get the length of the data
            sHeader = vbNullString
            For lngX = lngY To lngW
                If abCText(lngX) >= 48 And abCText(lngX) <= 58 Then
                    If abCText(lngX) = 58 Then Exit For
                    sHeader = sHeader & Chr(abCText(lngX))
                End If
            Next lngX
            ' get the data
            If Len(sHeader) > 0 And abCText(lngX) = 58 Then
                lngLength = CLng(sHeader)
                If lngLength > 0 And lngW >= lngX + lngLength Then
                    lngX = lngX + 1 ' start byte of cryptoText
                    ReDim abData(lngLength - 1)
                    For lngY = lngX To lngX + lngLength - 1
                        abData(lngY - lngX) = abCText(lngY)
                    Next lngY
                End If
            End If
        End If

        ' Import the Session Key and grab the handle
        lngLength = lngSessionBlobLength
        lngReturnValue = CryptImportKey( _
            hCryptProv, abSessionBlob(0), lngLength, 0, 0, hSessionKey)
        If lngReturnValue = 0 Then
            Debug.Print "DLL error: " & DecodeLastError(Err.LastDllError) & " Could not import key"
            Exit Sub
        End If
        
        ' Decrypt the Array
        bReturnStat = False
        lngLength = UBound(abData) + 1
        lngReturnValue = CryptDecrypt(hSessionKey, 0, 1, 0, abData(0), lngLength)
        If lngReturnValue = 0 Then _
            Debug.Print CStr(Err.LastDllError) & " Error during key pair CryptDecrypt decryption." _
            Else bReturnStat = True
    End If
End Sub

Public Function DecodeLastError(iErrCode As Long) As String

    Select Case iErrCode
    Case NTE_BAD_UID
        DecodeLastError = "NTE_BAD_UID"
    Case NTE_BAD_HASH
        DecodeLastError = "NTE_BAD_HASH"
    Case NTE_BAD_KEY
        DecodeLastError = "NTE_BAD_KEY"
    Case NTE_BAD_LEN
        DecodeLastError = "NTE_BAD_LEN"
    Case NTE_BAD_DATA
        DecodeLastError = "NTE_BAD_DATA"
    Case NTE_BAD_SIGNATURE
        DecodeLastError = "NTE_BAD_SIGNATURE"
    Case NTE_BAD_VER
        DecodeLastError = "NTE_BAD_VER"
    Case NTE_BAD_ALGID
        DecodeLastError = "NTE_BAD_ALGID"
    Case NTE_BAD_FLAGS
        DecodeLastError = "NTE_BAD_FLAGS"
    Case NTE_BAD_TYPE
        DecodeLastError = "NTE_BAD_TYPE"
    Case NTE_BAD_KEY_STATE
        DecodeLastError = "NTE_BAD_KEY_STATE"
    Case NTE_BAD_HASH_STATE
        DecodeLastError = "NTE_BAD_HASH_STATE"
    Case NTE_NO_KEY
        DecodeLastError = "NTE_NO_KEY"
    Case NTE_NO_MEMORY
        DecodeLastError = "NTE_NO_MEMORY"
    Case NTE_EXISTS
        DecodeLastError = "NTE_EXISTS"
    Case NTE_PERM
        DecodeLastError = "NTE_PERM"
    Case NTE_NOT_FOUND
        DecodeLastError = "NTE_NOT_FOUND"
    Case NTE_DOUBLE_ENCRYPT
        DecodeLastError = "NTE_DOUBLE_ENCRYPT"
    Case NTE_BAD_PROVIDER
        DecodeLastError = "NTE_BAD_PROVIDER"
    Case NTE_BAD_PROV_TYPE
        DecodeLastError = "NTE_BAD_PROV_TYPE"
    Case NTE_BAD_PUBLIC_KEY
        DecodeLastError = "NTE_BAD_PUBLIC_KEY"
    Case NTE_BAD_KEYSET
        DecodeLastError = "NTE_BAD_KEYSET"
    Case NTE_PROV_TYPE_NOT_DEF
        DecodeLastError = "NTE_PROV_TYPE_NOT_DEF"
    Case NTE_PROV_TYPE_ENTRY_BAD
        DecodeLastError = "NTE_PROV_TYPE_ENTRY_BAD"
    Case NTE_KEYSET_NOT_DEF
        DecodeLastError = "NTE_KEYSET_NOT_DEF"
    Case NTE_KEYSET_ENTRY_BAD
        DecodeLastError = "NTE_KEYSET_ENTRY_BAD"
    Case NTE_PROV_TYPE_NO_MATCH
        DecodeLastError = "NTE_PROV_TYPE_NO_MATCH"
    Case NTE_SIGNATURE_FILE_BAD
        DecodeLastError = "NTE_SIGNATURE_FILE_BAD"
    Case NTE_PROVIDER_DLL_FAIL
        DecodeLastError = "NTE_PROVIDER_DLL_FAIL"
    Case NTE_PROV_DLL_NOT_FOUND
        DecodeLastError = "NTE_PROV_DLL_NOT_FOUND"
    Case NTE_BAD_KEYSET_PARAM
        DecodeLastError = "NTE_BAD_KEYSET_PARAM"
    Case NTE_FAIL
        DecodeLastError = "NTE_FAIL"
    Case NTE_SYS_ERR
        DecodeLastError = "NTE_SYS_ERR"
    End Select
End Function

Public Function EncryptDecryptFileByPhrase(sPassPhrase As String, _
    sInFile As String, sOutFile As String, bEncrypt As Boolean) As Boolean
Dim intFF As Integer
Dim abPassPhrase() As Byte
Dim abText() As Byte
Dim lLen As Long

    intFF = FreeFile
    Open sInFile For Binary Access Read As #intFF
    lLen = LOF(intFF)
    If lLen > 0 Then
        ReDim abText(lLen - 1)
        Get #intFF, , abText
        Close #intFF
    Else
        Close #intFF
        Exit Function
    End If
    abPassPhrase = sPassPhrase
    If bolUnicode Then abPassPhrase = StrConv(sPassPhrase, vbFromUnicode, 1033)
    If EncryptDecryptPhrase(abPassPhrase(), abText(), lLen, bEncrypt) Then
        intFF = FreeFile
        Open sOutFile For Binary Access Write As #intFF
        Put #intFF, , abText
        Close #intFF
        EncryptDecryptFileByPhrase = True
    End If
End Function

Public Function EncryptDecryptStringByPhrase(sPassPhrase As String, _
    sInString As String, ByRef sOutString As String, bEncrypt As Boolean) As Boolean
Dim abPassPhrase() As Byte
Dim abText() As Byte
Dim lLen As Long

    lLen = Len(sInString)
    abPassPhrase = sPassPhrase
    abText = sInString
    If bolUnicode Then
        abPassPhrase = StrConv(sPassPhrase, vbFromUnicode, 1033)
        abText = StrConv(sInString, vbFromUnicode)
    End If
    If EncryptDecryptPhrase(abPassPhrase(), abText(), lLen, bEncrypt) Then
        sOutString = abText
        If bEncrypt Then
            If bolUnicode Then sOutString = StrConv(abText, vbUnicode, 1033)
            'sOutString = Left$(sOutString, lLen)
        Else
            If bolUnicode Then sOutString = StrConv(abText, vbUnicode, 1033)
            'sOutString = Left$(sOutString, lLen)
        End If
        EncryptDecryptStringByPhrase = True
    Else
        sOutString = vbNullString
    End If
End Function

Private Function EncryptDecryptPhrase(ByRef abPassPhrase() As Byte, _
    ByRef abData() As Byte, ByRef lLength As Long, Encrypt As Boolean) As Boolean
Dim abText() As Byte
Dim lngReturnValue As Long
Dim lngD As Long
Dim lngB As Long
Dim lngC As Long
Dim sLng As String

    'clear session key and hash
    If hSessionKey <> 0 Then CryptDestroyKey hSessionKey
    If hHash <> 0 Then CryptDestroyHash hHash
    
    'create the Hash Object
    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    If lngReturnValue = 0 Then Exit Function
    
    'Hash the data
    lngD = UBound(abPassPhrase) + 1
    lngReturnValue = CryptHashData(hHash, abPassPhrase(0), lngD, 0)
    If lngReturnValue = 0 Then Exit Function
    
    'create Session key from Hash
    lngReturnValue = CryptDeriveKey(hCryptProv, CALG_RC4, hHash, 0, hSessionKey)
    If lngReturnValue = 0 Then Exit Function
    lngReturnValue = CryptDestroyHash(hHash)
    If lngReturnValue = 0 Then Exit Function
    
    If Encrypt Then 'encrypt data
        abText = abData
        lngD = UBound(abData) + 1
        lngB = lngD
        lngC = lngB
        lngReturnValue = CryptEncrypt(hSessionKey, 0, 1, 0, abText(0), lngD, lngB)
        If lngD > lngC Then 'enlarged return array...fix it
            Err.Clear
            ReDim abText(lngD)
            abText = abData
            lngD = UBound(abData) + 1
            lngReturnValue = CryptEncrypt(hSessionKey, 0, 1, 0, abText(0), lngD, lngB)
        End If
        lLength = lngD
        sLng = CStr(lngD) & "|" & CStr(UBound(abData) + 1) & ":"
        lngC = Len(sLng) - 1
        ReDim abData(lngD + lngC)
        For lngB = 0 To lngC
            abData(lngB) = Asc(Mid$(sLng, lngB + 1, 1))
        Next lngB
        lngC = Len(sLng)
        For lngB = lngC To lngC + lngD - 1
            abData(lngB) = abText(lngB - lngC)
        Next lngB
    
    Else 'decrypt data
        lngB = UBound(abData)
        lngC = 0
        Do
            sLng = sLng & Chr(abData(lngC))
            lngC = lngC + 1
            If lngC > lngB Then Exit Function
        Loop Until abData(lngC) = Asc(":")
        If Val(sLng) > 0 Then
            lngB = lngC + 1
            lngD = CLng(Val(sLng))
            ReDim abText(lngD)
            If lngD + lngB > UBound(abData) Then
                For lngC = 0 To lngD - 1
                    abText(lngC) = abData(lngB + lngC)
                Next lngC
                lngReturnValue = CryptDecrypt(hSessionKey, 0, 1, 0, abText(0), lngD)
            End If
            lLength = CLng(Val(Mid$(sLng, InStr(1, sLng, "|", vbTextCompare) + 1)))
            abData = abText
            ReDim Preserve abData(lLength)
        End If
    End If
    
    If lngReturnValue = 0 Then Exit Function
    lngReturnValue = CryptDestroyKey(hSessionKey)
    If lngReturnValue = 0 Then Exit Function
    EncryptDecryptPhrase = True
End Function

Public Function SignString(ByRef sData As String, _
    Optional bReturnText As Boolean = True) As Boolean
Dim abData() As Byte
Dim abSigned() As Byte
Dim sText As String
Dim sSigned As String
Dim sReturn As String
Dim lngX As Long
Dim lngText As Long
Dim lngSign As Long

    SignString = False
    lngText = Len(sData)
    lngSign = 0
    If lngText < 1 Then Exit Function
    sText = sData
    ReDim abData(lngText - 1)
    ReDim abSigned(0)
    abSigned(0) = 0
    If bolUnicode Then
        abData = StrConv(sText, vbFromUnicode, 1033)
    Else
        abData = sText
    End If
    ReDim Preserve abData(lngText - 1)
    
    If SignValidate(abData, abSigned, True) = 0 Then
        sText = sData
        lngSign = UBound(abSigned) + 1
'        sSigned = Space(lngSign + Len(CStr(lngSign) & ":"))
'        sSigned = CStr(lngSign) & ":"
'        For lngX = 0 To lngSign - 1
'            sSigned = sSigned & Chr(abSigned(lngX))
'        Next lngX
'        sReturn = CStr(Len(sSigned)) & "|" & CStr(Len(sText)) & "|" _
'            & sSigned
'        If bReturnText Then sReturn = sReturn & sText
'        sData = Space(Len(sReturn))
'        sData = sReturn
        SignString = True

        Dim lenabSigned As Long, i As Long
        lenabSigned = UBound(abSigned)
        ReDim Preserve abSigned(lenabSigned + 2)
        abSigned(lenabSigned + 1) = 0
        abSigned(lenabSigned + 2) = 124 '|
        ReDim Preserve abSigned(lenabSigned + 2 + 1 + lngText - 1)
        For i = lenabSigned + 2 + 1 To lenabSigned + 2 + 1 + lngText - 1
            abSigned(i) = abData(i - lenabSigned - 2 - 1)
        Next
        
'        lenabSigned = lenabSigned + 2 + 1 + lngText - 1 + 1
'        ReDim Preserve abSigned(lenabSigned + 2)
'        abSigned(lenabSigned + 1) = 0
'        abSigned(lenabSigned + 2) = 58 ':
'        ReDim Preserve abSigned(19 + lenabSigned + 3)
'        i = lenabSigned + 3
'        abSigned(i + 0) = 49
'        abSigned(i + 1) = 0
'        abSigned(i + 2) = 0
'        abSigned(i + 3) = 0
'        abSigned(i + 4) = 48
'        abSigned(i + 5) = 0
'        abSigned(i + 6) = 0
'        abSigned(i + 7) = 0
'        abSigned(i + 8) = 48
'        abSigned(i + 9) = 0
'        abSigned(i + 10) = 0
'        abSigned(i + 11) = 0
'        abSigned(i + 12) = 48
'        abSigned(i + 13) = 0
'        abSigned(i + 14) = 0
'        abSigned(i + 15) = 0
'        abSigned(i + 16) = 48
'        abSigned(i + 17) = 0
'        abSigned(i + 18) = 0
'        abSigned(i + 19) = 0
'        Dim finalUBound As Long
'        finalUBound = UBound(abSigned)
'        Select Case Len(CStr(finalUBound))
'        Case 1
'            abSigned(i + 16) = finalUBound + 48
'        Case 2
'            abSigned(i + 12) = CLng(Mid(CStr(finalUBound), 1, 1)) + 48
'            abSigned(i + 16) = CLng(Mid(CStr(finalUBound), 2, 1)) + 48
'        Case 3
'            abSigned(i + 8) = CLng(Mid(CStr(finalUBound), 1, 1)) + 48
'            abSigned(i + 12) = CLng(Mid(CStr(finalUBound), 2, 1)) + 48
'            abSigned(i + 16) = CLng(Mid(CStr(finalUBound), 3, 1)) + 48
'        Case 4
'            abSigned(i + 4) = CLng(Mid(CStr(finalUBound), 1, 1)) + 48
'            abSigned(i + 8) = CLng(Mid(CStr(finalUBound), 2, 1)) + 48
'            abSigned(i + 12) = CLng(Mid(CStr(finalUBound), 3, 1)) + 48
'            abSigned(i + 16) = CLng(Mid(CStr(finalUBound), 4, 1)) + 48
'        Case 5
'            abSigned(i + 0) = CLng(Mid(CStr(finalUBound), 1, 1)) + 48
'            abSigned(i + 4) = CLng(Mid(CStr(finalUBound), 2, 1)) + 48
'            abSigned(i + 8) = CLng(Mid(CStr(finalUBound), 3, 1)) + 48
'            abSigned(i + 12) = CLng(Mid(CStr(finalUBound), 4, 1)) + 48
'            abSigned(i + 16) = CLng(Mid(CStr(finalUBound), 5, 1)) + 48
'        End Select
        
        ' Added by Ismail
        'Dim b64 As clsCryptoAPIBase64
        'Set b64 = New clsCryptoAPIBase64
        'sData = b64.encode(abSigned)
        'sData = b64.EncodeBytes64(abSigned)
        Dim b64 As clsCryptoAPIBase64
        Set b64 = New clsCryptoAPIBase64
        sData = b64.Encode(abSigned)
        ' Added by Ismail

    
''        ' Added by Ismail
'        Dim b64 As clsCryptoAPIBase64
'        Set b64 = New clsCryptoAPIBase64
'        sData = b64.encode(abSigned)
''        ' Added by Ismail
    End If
End Function

Public Function ValidateString(ByRef sText As String, _
    Optional sNoText As String = vbNullString) As Boolean
Dim abSigned() As Byte
Dim abText() As Byte
Dim strSigned As String
Dim strText As String
Dim lngX As Long
Dim lngL As Long
Dim lngY As Long
Dim lngM As Long, i As Long, j As Long
Dim newLimitabSigned As Long
Dim totalLen() As Byte
Dim uboundStr As String

' Added by Ismail
'Dim b64 As clsCryptoAPIBase64
'Set b64 = New clsCryptoAPIBase64
'abSigned = b64.decode(sText)
'abSigned = b64.DecodeBytes64(sText)
Dim b64 As clsCryptoAPIBase64
Set b64 = New clsCryptoAPIBase64
abSigned = b64.Decode(sText)
' Scan the byte array to find the ":"
On Error GoTo validateError

'For i = 0 To UBound(abSigned)
'    If abSigned(i) = 0 And abSigned(i + 1) = 58 Then
'        For j = i + 2 To UBound(abSigned)
'            If abSigned(j) <> 0 Then uboundStr = uboundStr & CStr(abSigned(j) - 48)
'        Next
'        If CLng(CLng(uboundStr) - 10000) <> UBound(abSigned) Then GoTo validateError
'        Exit For
'    End If
'Next
'ReDim Preserve abSigned(i - 2)

For i = 0 To UBound(abSigned)
    If abSigned(i) = 0 And abSigned(i + 1) = 124 Then
        ReDim abText(UBound(abSigned) - i - 2)
        For j = i + 2 To UBound(abSigned)
            abText(j - i - 2) = abSigned(j)
        Next
        Exit For
        newLimitabSigned = i - 2
    
    End If
Next
ReDim Preserve abSigned(i - 1)

' Added by Ismail

'' Added by Ismail
'Dim b64 As clsCryptoAPIBase64
'Set b64 = New clsCryptoAPIBase64
'sText = b64.DecodeToString(sText)
'' Added by Ismail
'
'ValidateString = False
'lngM = 0
'lngX = InStr(1, sText, "|")
'If lngX > 1 And lngX < 5 Then
'    lngL = CLng(Mid(sText, 1, lngX - 1)) ' lngL = length of Signature
'    If lngL > 0 Then
'        lngY = InStr(lngX + 2, sText, "|")
'        If lngY > lngX + 1 And lngY < lngX + 6 Then
'            lngM = CLng(Mid(sText, lngX + 1, lngY - lngX - 1)) ' lngM = length of Text
'            If lngM > 0 And Len(sText) + Len(sNoText) > lngL + lngM + lngY - 1 Then
'                strSigned = Mid(sText, lngY + 1, lngL)
'                If sNoText <> vbNullString Then
'                    strText = CStr(lngM) & ":" & sNoText
'                Else
'                    strText = CStr(lngM) & ":" & Mid(sText, lngY + 1 + lngL, lngM)
'                End If
'            Else
'                lngM = 0
'            End If
'        End If
'    End If
'End If
'
'lngY = 0
'lngX = InStr(1, strSigned, Chr(58))
'If lngX > 1 And lngX < 5 Then
'    lngL = CLng(Mid(strSigned, 1, lngX - 1)) ' lngL = length of Signature array
'    strSigned = Mid(strSigned, lngX + 1, lngL)
'    If lngL > 0 Then
'        ReDim Preserve abSigned(lngL - 1)
'        For lngY = 0 To lngL - 1
'            abSigned(lngY) = Asc(Mid(strSigned, lngY + 1, 1))
'        Next lngY
'    End If
'End If
'
'lngY = 0
'lngX = InStr(1, strText, Chr(58))
'If lngX > 1 And lngX < 5 Then
'    lngL = CLng(Mid(strText, 1, lngX - 1)) ' lngL = length of Text array
'    strText = Mid(strText, lngX + 1, lngL)
'    If lngL > 0 Then
'        ReDim Preserve abText(lngL - 1)
'        For lngY = 0 To lngL - 1
'            abText(lngY) = Asc(Mid(strText, lngY + 1, 1))
'        Next lngY
'    End If
'End If

'    If strText = "" Then Exit Function

If SignValidate(abText, abSigned, False) = 0 Then
    If bolUnicode Then
        sText = StrConv(abText, vbUnicode, 1033)
    Else
        sText = abText
    End If
    sText = strText
    ValidateString = True
Else
    ValidateString = False
End If
Exit Function

validateError:
ValidateString = False
End Function

Public Function GetHashVal(sData As String, ByRef abHashVal() As Byte) As Boolean
Dim hHash As Long
Dim lngReturnValue As Long
Dim lngHashLen As Long
Dim abText() As Byte
Dim strTxt As String
Dim lngW As Long
Dim lngX As Long
Dim lngY As Long

    GetHashVal = False
    'Create a hash object to sign/validate
    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    If lngReturnValue <> 0 Then
        'Hash the data
        ReDim abText(Len(sData))
        If bolUnicode Then
            abText = StrConv(sData, vbFromUnicode, 1033)
        Else
            abText = sData
        End If
        lngW = UBound(abText) + 1
        lngReturnValue = CryptHashData(hHash, abText(0), lngW, 0)
        If lngReturnValue <> 0 Then
            'Get the length of the Hash
            lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, 0, lngHashLen, 0)
            'Get the Hash value
            ReDim abHashVal(lngHashLen)
            lngReturnValue = CryptGetHashParam(hHash, HP_HASHVAL, abHashVal(0), lngHashLen, 0)
            If lngReturnValue <> 0 And lngHashLen > 0 Then
                GetHashVal = True
                ReDim Preserve abHashVal(lngHashLen - 1)
            End If
        End If
    End If
End Function

Private Function SignValidate(ByRef abData() As Byte, _
                            ByRef abSigned() As Byte, _
                            Optional bSigned As Boolean = True) As Long
Dim hHash As Long
Dim lngReturnValue As Long
Dim lngSigLen As Long
Dim abText() As Byte
Dim strTxt As String
Dim lngW As Long
Dim lngX As Long
Dim lngY As Long

    SignValidate = -1
    ReDim abText(UBound(abData))
    abText = abData
    
    'Create a hash object to sign/validate
    lngReturnValue = CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Could not create a Hash Object (CryptCreateHash API)"
    'Hash the data
    lngW = UBound(abText) + 1
    lngReturnValue = CryptHashData(hHash, abText(0), lngW, 0)
    If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Could not calculate a Hash Value (CryptHashData API)"
    
    If bSigned Then
        'release old key pair handle
        If hKeyPair <> 0 Then CryptDestroyKey hKeyPair
        'get a handle to the signature key pair
        lngReturnValue = CryptGetUserKey(hCryptProv, AT_SIGNATURE, hKeyPair)
        If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , _
            "DLL error code shown above. Could not obtain key pair"
        'Determine the size of the signature
        lngReturnValue = CryptSignHash(hHash, AT_SIGNATURE, "sign", 0, vbNull, lngSigLength)
        If lngSigLength > 0 Then ReDim abSig(lngSigLength - 1)
        'Sign the hash object
        lngReturnValue = CryptSignHash(hHash, AT_SIGNATURE, "sign", 0, abSig(0), lngSigLength)
        If lngReturnValue = 0 Then Err.Raise Err.LastDllError, , "DLL error code shown above. Could not sign the hash"
        ' the signature is now available
        ' size returned array to signature length
        ReDim abSigned(UBound(abSig))
        ' return the signature to the calling procedure
        abSigned = abSig
        SignValidate = 0
    Else
        lngSigLength = UBound(abSigned) + 1
        ReDim abSig(UBound(abSigned))
        abSig = abSigned ' load the Signature array
        MsgBox hHash & "  " & abSig(0) & "  " & lngSigLength & "  " & hKeyPair
        lngReturnValue = CryptVerifySignature(hHash, abSig(0), lngSigLength, hKeyPair, "sign", 0)
        If lngReturnValue = 0 Then 'some error occurred
            SignValidate = Err.LastDllError
            MsgBox "SignValidate returns " & SignValidate
        Else
            SignValidate = 0
        End If
    End If
End Function
