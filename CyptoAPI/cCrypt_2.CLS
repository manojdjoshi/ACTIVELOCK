Attribute VB_Name = "cCrypt"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0   'vbNone
  MTSTransactionMode = 0   'NotAnMTSObject
End
Option Explicit


'=============================================================================================================
'
' cCrypt Class Module
' -------------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Created On  : May 01, 2001
' Last Update : June 20, 2002
'
' VB Versions : 5.0 / 6.0
'
' Requires    : Windows 95 OSR2 or later (or Windows 95 with Internet Explorer 3.02 or later)
'
' Description : This class module was created to easily encrypt and decrypt strings using Hash encryption
'               via the Windows API.
'
' Example Use :
'
'  Private Crypt As cCrypt
'
'  Private Sub Form_Load()
'    Set Crypt = New cCrypt
'    Crypt.Password = "Password"
'  End Sub
'
'  Private Sub cmdDecrypt_Click()
'    txtOutput.Text = Crypt.Decrypt_String(txtInput.Text)
'  End Sub
'
'  Private Sub cmdEncrypt_Click()
'    txtOutput.Text = Crypt.Encrypt_String(txtInput.Text)
'  End Sub
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================


' Constants - CryptAcquireContext.dwProvType
Private Const PROV_RSA_FULL = 1        ' The PROV_RSA_FULL provider type supports both digital signatures and data encryption. It is considered a general purpose CSP. The RSA public-key algorithm is used for all public-key operations.
Private Const PROV_RSA_SIG = 2         ' The PROV_RSA_SIG provider type is a subset of PROV_RSA_FULL. It supports only those functions and algorithms required for hashes and digital signatures.
Private Const PROV_RSA_SCHANNEL = 12   ' The PROV_RSA_SCHANNEL provider type supports both RSA and Schannel protocols.
Private Const PROV_DSS = 3             ' The PROV_DSS provider type, like PROV_RSA_SIG, only supports hashes and digital signatures. The signature algorithm specified by the PROV_DSS provider type is the Digital Signature Algorithm (DSA).
Private Const PROV_DSS_DH = 13         ' The PROV_DSS_DH provider is a superset of the PROV_DSS provider type.
Private Const PROV_DH_SCHANNEL = 18    ' The PROV_DH_SCHANNEL provider type supports both Diffie-Hellman and Schannel protocols
Private Const PROV_FORTEZZA = 4        ' The PROV_FORTEZZA provider type contains a set of cryptographic protocols and algorithms owned by the National Institute of Standards and Technology (NIST).
Private Const PROV_MS_EXCHANGE = 5     ' The PROV_MS_EXCHANGE provider type is designed for the cryptographic needs of the Microsoft Exchange mail application and other applications compatible with Microsoft Mail. This provider type is preliminary.
Private Const PROV_SSL = 6             ' The PROV_SSL provider type supports the Secure Sockets Layer (SSL) protocol.

' Constants - CryptAcquireContext.pszProvider
' (The following Cryptographic Service Provider (CSP) names are defined in the Win32 API)
Private Const MS_DEF_PROV              As String = "Microsoft Base Cryptographic Provider v1.0" & vbNullChar
Private Const MS_ENHANCED_PROV         As String = "Microsoft Enhanced Cryptographic Provider" & vbNullChar
Private Const MS_STRONG_PROV           As String = "Microsoft Strong Cryptographic Provider" & vbNullChar
Private Const MS_DEF_RSA_SIG_PROV      As String = "Microsoft RSA Signature Cryptographic Provider" & vbNullChar
Private Const MS_DEF_RSA_SCHANNEL_PROV As String = "Microsoft RSA SChannel Cryptographic Provider" & vbNullChar
Private Const MS_DEF_DSS_PROV          As String = "Microsoft Base DSS Cryptographic Provider" & vbNullChar
Private Const MS_DEF_DSS_DH_PROV       As String = "Microsoft Base DSS and Diffie-Hellman Cryptographic Provider" & vbNullChar
Private Const MS_ENH_DSS_DH_PROV       As String = "Microsoft Enhanced DSS and Diffie-Hellman Cryptographic Provider" & vbNullChar
Private Const MS_DEF_DH_SCHANNEL_PROV  As String = "Microsoft DH SChannel Cryptographic Provider" & vbNullChar
Private Const MS_SCARD_PROV            As String = "Microsoft Base Smart Card Cryptographic Provider" & vbNullChar

' Algorithm Classes
Private Const ALG_CLASS_ANY = 0
Private Const ALG_CLASS_SIGNATURE = 8192     '(1 << 13)
Private Const ALG_CLASS_MSG_ENCRYPT = 16384  '(2 << 13)
Private Const ALG_CLASS_DATA_ENCRYPT = 24576 '(3 << 13)
Private Const ALG_CLASS_HASH = 32768         '(4 << 13)
Private Const ALG_CLASS_KEY_EXCHANGE = 40960 '(5 << 13)
Private Const ALG_CLASS_ALL = 57344          '(7 << 13)

' Algorithm Types
Private Const ALG_TYPE_ANY = 0
Private Const ALG_TYPE_DSS = 512            '(1 << 9)
Private Const ALG_TYPE_RSA = 1024           '(2 << 9)
Private Const ALG_TYPE_BLOCK = 1536         '(3 << 9)
Private Const ALG_TYPE_STREAM = 2048        '(4 << 9)
Private Const ALG_TYPE_DH = 2560            '(5 << 9)
Private Const ALG_TYPE_SECURECHANNEL = 3072 '(6 << 9)

' Sub-IDs (Windows Compatible)
Private Const ALG_SID_ANY = 0
Private Const ALG_SID_MD2 = 1
Private Const ALG_SID_RC2 = 2
Private Const ALG_SID_MD5 = 3
Private Const ALG_SID_SHA = 4
Private Const ALG_SID_SHA1 = 4
Private Const ALG_SID_MAC = 5
Private Const ALG_SID_SSL3SHAMD5 = 8
Private Const ALG_SID_HMAC = 9

' Stream Cipher Sub-IDs
Private Const ALG_SID_RC4 = 1
Private Const ALG_SID_SEAL = 2

' Constants - CryptAcquireContext.dwFlags
Private Const CRYPT_VERIFYCONTEXT = &HF0000000 ' The application has no access to the private keys, and the pszContainer parameter must be set to NULL.  This option is intended for applications that do not use private keys.  When CryptAcquireContext is called, many CSPs require input from the owning user before granting access to the private keys in the key container. For example, the private keys can be encrypted, requiring a password from the user before they can be used. However, if the CRYPT_VERIFYCONTEXT flag is specified, access to the private keys is not required and the user interface can be bypassed.
Private Const CRYPT_NEWKEYSET = &H8            ' A new key container is created with the name specified by pszContainer. If pszContainer is NULL, a key container with the default name is created.
Private Const CRYPT_MACHINE_KEYSET = &H20      ' By default, keys and key containers are stored as user keys. For Base Providers, this means that user key containers are stored in the user's profile. The CRYPT_MACHINE_KEYSET flag can be combined with all of the other flags to indicate that the key container of interest is a machine key container and the CSP treats it as such. For Base Providers, this means that the keys are stored locally on the computer that created the key container. If a key container is to be a machine container, the CRYPT_MACHINE_KEYSET flag must be used with all calls to CryptAcquireContext that reference the machine container. The CRYPT_MACHINE_KEYSET flag is useful when the user is accessing from a service or user account that did not log on interactively When key containers are created, most CSPs do not automatically create any public/private key pairs. These keys must be created as a separate step with the CryptGenKey function.
Private Const CRYPT_DELETEKEYSET = &H10        ' The key container specified by pszContainer is deleted. If pszContainer is NULL, the key container with the default name is deleted. All key pairs in the key container are also destroyed.  When this flag is set, the value returned in phProv is undefined, and thus, the CryptReleaseContext function need not be called afterwards.
Private Const CRYPT_SILENT = &H40              ' The application requests that the CSP not display any user interface (UI) for this context. If the CSP must display the UI to operate, the call fails and the NTE_SILENT_CONTEXT error code is set as the last error. In addition, if calls are made to CryptGenKey with the CRYPT_USER_PROTECTED flag with a context that has been acquired with the CRYPT_SILENT flag, the calls fail and the CSP sets NTE_SILENT_CONTEXT.  CRYPT_SILENT is intended for use with applications for which the UI cannot be displayed by the CSP.  This flag is supported with Microsoft® Windows® 2000 or later. It is not supported in Windows 98 or Microsoft® Internet Explorer version 5.0.

' Constants - CryptHashData.dwFlags
Private Const CRYPT_USERDATA = 1 'All Microsoft Cryptographic Providers ignore this parameter. For any CSP that does not ignore this parameter, if this flag is set, the CSP prompts the user to input data directly. This data is added to the hash. The application is not allowed access to the data. This flag can be used to allow the user to enter a PIN into the system.

' Constants - CryptDeriveKey.Algid
Private Const CALG_RC2 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_RC2) ' RC2 block cipher
Private Const CALG_RC4 = (ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM Or ALG_SID_RC4) 'RC4 stream cipher

' Constants - CryptDecrypt.dwFlags
Private Const CRYPT_OAEP = &H40 ' When set with the MS Enhanced Provider and RSA encryption/decryption causes PKCS #1 version 2 formatting to be used.

' Constants - CryptCreateHash.Algid
Private Const CALG_HMAC = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_HMAC)              ' HMAC, a keyed hash algorithm
Private Const CALG_MAC = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MAC)                ' MAC Message Authentication Code
Private Const CALG_MD2 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2)                ' MD2 hashing algorithm
Private Const CALG_MD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5)                ' MD5 hashing algorithm
Private Const CALG_SHA = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA)                ' US DSA Secure Hash Algorithm
Private Const CALG_SHA1 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1)              ' Same as CALG_SHA
Private Const CALG_SSL3_SHAMD5 = (ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SSL3SHAMD5) ' SSL3 client authentication

' Error Constants
Private Const ERROR_INVALID_HANDLE = 6&            ' One of the parameters specifies an invalid handle.
Private Const ERROR_INVALID_PARAMETER = 87         ' One of the parameters contains an invalid value. This is most often an illegal pointer.
Private Const ERROR_NOT_ENOUGH_MEMORY = 8          ' The operating system ran out of memory during the operation.
Private Const ERROR_BUSY = 170&                    ' The hash object specified by hHash is currently being used by another process.
Private Const NTE_BAD_ALGID = &H80090008           ' The Algid parameter specifies an algorithm that this CSP does not support.
Private Const NTE_BAD_DATA = &H80090005            ' The data to be encrypted is invalid. For example, when a block cipher is used and the Final flag is FALSE, the value specified by pdwDataLen must be a multiple of the block size.
Private Const NTE_BAD_FLAGS = &H80090009           ' The dwFlags parameter has an illegal value.
Private Const NTE_BAD_HASH = &H80090002            ' The hash object specified by the hHash parameter is invalid.
Private Const NTE_BAD_HASH_STATE = &H8009000C      ' An attempt was made to add data to a hash object that is already marked "finished".
Private Const NTE_BAD_KEY = &H80090003             ' A keyed hash algorithm (such as CALG_MAC) is specified by Algid and the hKey parameter is either zero or it specifies an invalid key handle. This error code will also be returned if the key is to a stream cipher, or if the cipher mode is anything other than CBC.
Private Const NTE_BAD_KEYSET = &H80090016          ' The Registry entry for the key container could not be opened and may not exist.
Private Const NTE_BAD_KEYSET_PARAM = &H8009001F    ' The pszContainer or pszProvider parameter is set to an illegal value.
Private Const NTE_BAD_LEN = &H80090004             ' The CRYPT_USERDATA flag is set and the dwDataLen parameter has a nonzero value.
Private Const NTE_BAD_PROV_TYPE = &H80090014       ' The value of the dwProvType parameter is out of range. All provider types must be from 1 to 999, inclusive.
Private Const NTE_BAD_SIGNATURE = &H80090006       ' The provider DLL signature did not verify correctly. Either the DLL or the digital signature has been tampered with.
Private Const NTE_BAD_UID = &H80090001             ' The CSP context that was specified when the hash object was created cannot be found.
Private Const NTE_DOUBLE_ENCRYPT = &H80090012      ' The application attempted to encrypt the same data twice.
Private Const NTE_EXISTS = &H8009000F              ' The dwFlags parameter is CRYPT_NEWKEYSET, but the key container already exists.
Private Const NTE_FAIL = &H80090020                ' The function failed in some unexpected way.
Private Const NTE_KEYSET_ENTRY_BAD = &H8009001A    ' The Registry entry for the pszContainer key container was found (in the HKEY_CURRENT_USER window), but is corrupt. See the section System Administration for details about CryptoAPI’s Registry usage.
Private Const NTE_KEYSET_NOT_DEF = &H80090019      ' No Registry entry exists in the HKEY_CURRENT_USER window for the key container specified by pszContainer.
Private Const NTE_NO_MEMORY = &H8009000E           ' The CSP ran out of memory during the operation.
Private Const NTE_PROV_DLL_NOT_FOUND = &H8009001E  ' The provider DLL file does not exist or is not on the current path.
Private Const NTE_PROV_TYPE_ENTRY_BAD = &H80090018 ' The Registry entry for the provider type specified by dwProvType is corrupt. This error may relate to either the user default CSP list or the machine default CSP list. See the section System Administration for details about CryptoAPI’s Registry usage.
Private Const NTE_PROV_TYPE_NO_MATCH = &H8009001B  ' The provider type specified by dwProvType does not match the provider type found in the Registry. Note that this error can only occur when pszProvider specifies an actual CSP name.
Private Const NTE_PROV_TYPE_NOT_DEF = &H80090017   ' No Registry entry exists for the provider type specified by dwProvType.
Private Const NTE_PROVIDER_DLL_FAIL = &H8009001D   ' The provider DLL file could not be loaded, and may not exist. If it exists, then the file is not a valid DLL.
Private Const NTE_SIGNATURE_FILE_BAD = &H8009001C  ' An error occurred while loading the DLL file image, prior to verifying its signature.

' Property Variables
Private p_Password   As String
Private p_CSP_String As String
Private p_CSP_Type   As Long

' Win32 API Declarations - General
Private Declare Function GetLastError Lib "KERNEL32" () As Long

' Win32 API Declarations - Encryption
Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As Long, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, ByVal pbData As String, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hBaseData As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long, ByVal dwBufLen As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, ByVal dwFlags As Long, ByVal pbData As String, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


Private Sub Class_Initialize()
  p_Password = "password"
  p_CSP_String = MS_DEF_PROV
  p_CSP_Type = PROV_RSA_FULL
End Sub


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


Public Property Get Password() As String
  Password = p_Password
End Property
Public Property Let Password(ByVal NewValue As String)
  p_Password = NewValue
End Property


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


' This procedure is the one that actually does the work of decrypting the data
Public Function Decrypt_String(ByVal StringToDecrypt As String, _
                               Optional ByVal ConvertFromNumbers As Boolean, _
                               Optional ByVal ConvertFromHEX As Boolean, _
                               Optional ByRef Return_String As String, _
                               Optional ByRef Return_ErrNum As Long, _
                               Optional ByRef Return_ErrDesc As String) As Boolean
On Error GoTo ErrorTrap
  
  Dim lngHash         As Long
  Dim lngKey          As Long
  Dim lngCryptProv    As Long
  Dim strCryptBuffer  As String
  Dim lngCryptBuffLen As Long
  
  ' Set default values
  Return_String = ""
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Make sure the user has specified a password
  If p_Password = "" Then
    Return_ErrNum = -1: Return_ErrDesc = "No password has been defined to decrypt with"
    Exit Function
  
  ' Make sure decrypt string is valid
  ElseIf StringToDecrypt = "" Then
    Return_ErrNum = -1: Return_ErrDesc = "No string specified to decrypt"
    Exit Function
  End If
  
  ' If the user is converting from a decimal number, it has to be numeric
  If ConvertFromNumbers = True Then
    StringToDecrypt = Trim(StringToDecrypt)
    If ContainsNonNumeric(StringToDecrypt) = True Then
      If ConvertFromHEX = False Then
        Return_ErrNum = -1: Return_ErrDesc = "Invalid decimal string specified to convert.  Value must be numeric ONLY."
        Exit Function
      Else
        ConvertFromNumbers = False
      End If
    End If
  ElseIf ConvertFromHEX = True Then
    StringToDecrypt = Trim(StringToDecrypt)
  End If
  
  ' If the user specified to convert to HEX or DEC then convert it here
  If ConvertFromHEX = True Then
    StringToDecrypt = Convert_HEXtoSTR(StringToDecrypt)
  ElseIf ConvertFromNumbers = True Then
    StringToDecrypt = Convert_NUMtoSTR(StringToDecrypt)
  End If
  
  ' Get handle to the default CSP.
  If CryptAcquireContext(lngCryptProv, 0, p_CSP_String, p_CSP_Type, CRYPT_MACHINE_KEYSET) = 0 Then
    ' Bad provider name or API error
    CheckGetLastError Err.LastDllError, Return_ErrNum, Return_ErrDesc, "CryptAcquireContext", False
    Exit Function
  End If
  
  ' Create a hash object
  If CryptCreateHash(lngCryptProv, CALG_MD5, 0, 0, lngHash) = 0 Then
    ' Error creating encrypt object or API error
    CheckGetLastError Err.LastDllError, Return_ErrNum, Return_ErrDesc, "CryptCreateHash", False
    GoTo CleanUp
  End If
  
  ' Hash in the password text
  If CryptHashData(lngHash, p_Password, Len(p_Password), 0) = 0 Then
    ' Error passing key or API error
    CheckGetLastError Err.LastDllError, Return_ErrNum, Return_ErrDesc, "CryptHashData", False
    GoTo CleanUp
  End If
  
  ' Create a session key from the hash object
  If CryptDeriveKey(lngCryptProv, CALG_RC4, lngHash, 0, lngKey) = 0 Then
    ' Error creating a session key or API error
    CheckGetLastError Err.LastDllError, Return_ErrNum, Return_ErrDesc, "CryptDeriveKey", False
    GoTo CleanUp
  End If
  
  ' Destroy the hash object.
  CryptDestroyHash lngHash
  lngHash = 0
  
  ' Prepare strCryptBuffer for CryptDecrypt
  lngCryptBuffLen = Len(StringToDecrypt) * 2
  strCryptBuffer = String(lngCryptBuffLen, vbNullChar)
  LSet strCryptBuffer = StringToDecrypt
  
  ' Decrypt data
  If CryptDecrypt(lngKey, 0, 1, 0, strCryptBuffer, lngCryptBuffLen) = 0 Then
    If CheckGetLastError(GetLastError, , , "CryptDecrypt", True) = False Then Err.Raise -1, "CryptDecrypt", "Error decrypting string or API error"
    GoTo CleanUp
  End If
  
  ' Setup output buffer with just decrypted data
  Return_String = Mid(strCryptBuffer, 1, lngCryptBuffLen / 2)
  Decrypt_String = True
  
CleanUp:
  
  ' Destroy session key
  If lngKey <> 0 Then CryptDestroyKey lngKey
    
  ' Destroy hash object
  If lngHash <> 0 Then CryptDestroyHash lngHash
  
  ' Release Context provider handle
  If lngCryptProv <> 0 Then CryptReleaseContext lngCryptProv, 0
  
  Exit Function
  
ErrorTrap:
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  GoTo CleanUp
  
End Function

' This method is the method that actually does the work of encrypting the data
Public Function Encrypt_String(ByVal StringToEncrypt As String, _
                               Optional ByVal ConvertToNumbers As Boolean, _
                               Optional ByVal ConvertToHEX As Boolean, _
                               Optional ByRef Return_String As String, _
                               Optional ByRef Return_ErrNum As Long, _
                               Optional ByRef Return_ErrDesc As String) As Boolean
On Error GoTo ErrorTrap
  
  Dim lngHash         As Long
  Dim lngKey          As Long
  Dim lngCryptProv    As Long
  Dim lngCryptLen     As Long
  Dim strCryptBuffer  As String
  Dim lngCryptBuffLen As Long
  
  ' Set default values
  Return_String = ""
  Return_ErrNum = 0
  Return_ErrDesc = ""
  
  ' Make sure the user has specified a password
  If p_Password = "" Then
    Return_ErrNum = -1: Return_ErrDesc = "No password has been defined to encrypt with"
    Exit Function
    
  ' Make sure encrypt string is valid
  ElseIf StringToEncrypt = "" Then
    Return_ErrNum = -1: Return_ErrDesc = "No string specified to encrypt"
    Exit Function
  End If
  
  ' Get handle to the default CSP
  If CryptAcquireContext(lngCryptProv, 0, p_CSP_String, p_CSP_Type, CRYPT_MACHINE_KEYSET) = 0 Then
    
    ' If there is no default key container then create one using Flags field
    CheckGetLastError Err.LastDllError, Return_ErrNum, Return_ErrDesc, "CryptAcquireContext", False
    If Return_ErrNum = -2146893802 Then
      Return_ErrNum = 0
      Return_ErrDesc = ""
      If CryptAcquireContext(lngCryptProv, 0, p_CSP_String, p_CSP_Type, CRYPT_NEWKEYSET Or CRYPT_MACHINE_KEYSET) = 0 Then
        CheckGetLastError Err.LastDllError, Return_ErrNum, Return_ErrDesc, "CryptAcquireContext", False
        Exit Function
      End If
    Else
      Exit Function
    End If
  End If
  
  ' Create a hash object
  If CryptCreateHash(lngCryptProv, CALG_MD5, 0, 0, lngHash) = 0 Then
    ' Error creating encrypt object or API error
    CheckGetLastError Err.LastDllError, Return_ErrNum, Return_ErrDesc, "CryptCreateHash", False
    GoTo CleanUp
  End If
  
  ' Hash in the password text
  If CryptHashData(lngHash, p_Password, Len(p_Password), 0) = 0 Then
    ' Error passing key or API error
    CheckGetLastError Err.LastDllError, Return_ErrNum, Return_ErrDesc, "CryptHashData", False
    GoTo CleanUp
  End If
  
  ' Create a session key from the hash object.
  If CryptDeriveKey(lngCryptProv, CALG_RC4, lngHash, 0, lngKey) = 0 Then
    ' Error creating session key or API error
    CheckGetLastError Err.LastDllError, Return_ErrNum, Return_ErrDesc, "CryptDeriveKey", False
    GoTo CleanUp
  End If
  
  ' Destroy the hash object.
  CryptDestroyHash lngHash
  lngHash = 0
  
  ' Create a buffer for the CryptEncrypt function
  lngCryptLen = Len(StringToEncrypt)
  lngCryptBuffLen = lngCryptLen * 2
  strCryptBuffer = String(lngCryptBuffLen, vbNullChar)
  LSet strCryptBuffer = StringToEncrypt
  
  ' Encrypt the text data
  If CryptEncrypt(lngKey, 0, 1, 0, strCryptBuffer, lngCryptLen, lngCryptBuffLen) = 0 Then
    ' Error encrypting the data or API error
    CheckGetLastError Err.LastDllError, Return_ErrNum, Return_ErrDesc, "CryptEncrypt", False
    GoTo CleanUp
  End If
  
  ' Return the results
  Return_String = Mid(strCryptBuffer, 1, lngCryptLen)
  If ConvertToHEX = True Then
    Return_String = Convert_STRtoHEX(Return_String)
  ElseIf ConvertToNumbers = True Then
    Return_String = Convert_STRtoNUM(Return_String)
  End If
  Encrypt_String = True
  
CleanUp:
  
  ' Destroy session key.
  If lngKey <> 0 Then CryptDestroyKey lngKey
  
  ' Destroy hash object
  If lngHash <> 0 Then CryptDestroyHash lngHash
  
  ' Release Context provider handle
  If lngCryptProv <> 0 Then CryptReleaseContext lngCryptProv, 0
  
  Exit Function
  
ErrorTrap:
  Return_ErrNum = Err.Number
  Return_ErrDesc = Err.Description
  Err.Clear
  GoTo CleanUp
  
End Function


'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX


' Function that takes a string and takes it letter by letter and converts it
' the ASCII numerical equivelent... and then returns a string of numbers
Private Function Convert_STRtoNUM(ByVal StringToConvert As String) As String
On Error Resume Next
  
  Dim MyCounter   As Long
  Dim CharLeft    As String
  Dim CharRight   As String
  Dim StringSoFar As String
  Dim OneChar     As String
  
  For MyCounter = 1 To Len(StringToConvert)
    CharLeft = Left(StringToConvert, MyCounter)
    CharRight = Right(CharLeft, 1)
    If CharRight <> "" Then
      OneChar = CStr(Asc(CharRight))
      If Len(OneChar) = 1 Then
        StringSoFar = StringSoFar & "00" & OneChar
      ElseIf Len(OneChar) = 2 Then
        StringSoFar = StringSoFar & "0" & OneChar
      Else
        StringSoFar = StringSoFar & OneChar
      End If
    End If
  Next
  
  Convert_STRtoNUM = StringSoFar
  
End Function

' Function that takes a string and takes it letter by letter and converts it
' the string equivelent... and then returns a string of characters
Private Function Convert_NUMtoSTR(ByVal NumbersToConvert As String) As String
On Error Resume Next
  
  Dim MyCounter   As Long
  Dim CharLeft    As String
  Dim CharRight   As String
  Dim StringSoFar As String
  Dim CharSoFar   As String
  Dim OneChar     As String
  
  For MyCounter = 1 To Len(NumbersToConvert)
    CharLeft = Left(NumbersToConvert, MyCounter)
    CharRight = Right(CharLeft, 1)
    CharSoFar = CharSoFar & CharRight
    If Len(CharSoFar) = 3 Then
      If IsNumeric(CharSoFar) = False Then
        MsgBox "Invalid value passed to convert", vbOKOnly + vbExclamation, "  Error"
        Exit Function
      End If
      OneChar = Chr(CharSoFar)
      StringSoFar = StringSoFar & OneChar
      CharSoFar = ""
    End If
  Next
  
  Convert_NUMtoSTR = StringSoFar
  
End Function

' Function that takes a string and takes it letter by letter and converts it
' the HEXIDECIMAL equivelent... and then returns a string of HEX values
Private Function Convert_STRtoHEX(ByVal StringToConvert As String) As String
On Error Resume Next
  
  Dim MyCounter   As Long
  Dim CharLeft    As String
  Dim CharRight   As String
  Dim StringSoFar As String
  Dim OneChar     As String
  
  For MyCounter = 1 To Len(StringToConvert)
    CharLeft = Left(StringToConvert, MyCounter)
    CharRight = Right(CharLeft, 1)
    If CharRight <> "" Then
      OneChar = CStr(Hex(Asc(CharRight)))
      If Len(OneChar) = 1 Then OneChar = "0" & OneChar
      StringSoFar = StringSoFar & OneChar
    End If
  Next
  
  Convert_STRtoHEX = StringSoFar
  
End Function

' Function that takes a string and takes it letter by letter and converts it
' the string value equivelent... and then returns the string of characters
Private Function Convert_HEXtoSTR(ByVal NumbersToConvert As String) As String
On Error Resume Next
  
  Dim MyCounter   As Long
  Dim CharLeft    As String
  Dim CharRight   As String
  Dim StringSoFar As String
  Dim CharSoFar   As String
  Dim OneChar     As String
  
  For MyCounter = 1 To Len(NumbersToConvert)
    CharLeft = Left(NumbersToConvert, MyCounter)
    CharRight = Right(CharLeft, 1)
    CharSoFar = CharSoFar & CharRight
    If Len(CharSoFar) = 2 Then
      OneChar = Chr(Val("&H" & CharSoFar))
      StringSoFar = StringSoFar & OneChar
      CharSoFar = ""
    End If
  Next
  
  Convert_HEXtoSTR = StringSoFar
  
End Function

' Generates a random password that is used to encrypt the data
Public Function GeneratePassword(ByVal PasswordLength As Integer) As String
On Error GoTo ErrorTrap
  
  Dim sUniquePassword As String
  Dim lCounter        As Long
  Dim lValue          As Long
  
  ' Restrict maximum password length
  If PasswordLength > 256 Then PasswordLength = 256
  
  ' Generate random password
  Randomize Timer
  sUniquePassword = ""
  For lCounter = 1 To PasswordLength
    lValue = (Rnd() * 36) + 65
    If lValue > 90 Then
      sUniquePassword = sUniquePassword & CStr(lValue - 90)
    Else
      sUniquePassword = sUniquePassword & Chr(lValue)
    End If
  Next lCounter
  
  GeneratePassword = sUniquePassword
  
  Exit Function
  
ErrorTrap:
  
  If Err.Number = 0 Then      ' No Error
    Resume Next
  ElseIf Err.Number = 20 Then ' Resume Without Error
    Resume Next
  Else                        ' Other Error
    MsgBox Err.Source & " caused the following error:" & Chr(13) & Chr(13) & "Error Number = " & CStr(Err.Number) & Chr(13) & "Error Description = " & Err.Description, vbOKOnly + vbExclamation, "  Error  -  " & Err.Description
    Err.Clear
  End If
  
End Function

' Function that checks if an error occured in the last crypt related API
' called, and if there was... an error message is displayed
Private Function CheckGetLastError(ByVal ErrorNumber As Long, _
                                   Optional ByRef Return_ErrNum As Long, _
                                   Optional ByRef Return_ErrDesc As String, _
                                   Optional ByVal NameOfLastAPICalled As String = "last", _
                                   Optional ByVal ShowErrorMsg As Boolean = True) As Boolean
On Error Resume Next
  
  ' Set the default return values
  Return_ErrNum = 0
  Return_ErrDesc = ""
  CheckGetLastError = False
  
  ' Check if an error number was passed.  If none was, check for an error occurance.
  If ErrorNumber = 0 Then
    ErrorNumber = GetLastError
    If ErrorNumber = 0 Then
      Err.Clear
      Exit Function
    End If
  End If
  
  ' An Error Occured
  CheckGetLastError = True
  Return_ErrNum = ErrorNumber
  Err.Clear
  
  ' Get the error text
  Select Case ErrorNumber
    Case 0:                       Exit Function
    Case ERROR_INVALID_HANDLE:    Return_ErrDesc = "One of the parameters specifies an invalid handle."
    Case ERROR_INVALID_PARAMETER: Return_ErrDesc = "One of the parameters contains an invalid value. This is most often an illegal pointer."
    Case ERROR_NOT_ENOUGH_MEMORY: Return_ErrDesc = "The operating system ran out of memory during the operation."
    Case ERROR_BUSY:              Return_ErrDesc = "The hash object specified by hHash is currently being used by another process."
    Case NTE_BAD_ALGID:           Return_ErrDesc = "The Algid parameter specifies an algorithm that this CSP does not support."
    Case NTE_BAD_DATA:            Return_ErrDesc = "The data to be encrypted is invalid. For example, when a block cipher is used and the Final flag is FALSE, the value specified by pdwDataLen must be a multiple of the block size."
    Case NTE_BAD_FLAGS:           Return_ErrDesc = "The dwFlags parameter has an illegal value."
    Case NTE_BAD_HASH:            Return_ErrDesc = "The hash object specified by the hHash parameter is invalid."
    Case NTE_BAD_HASH_STATE:      Return_ErrDesc = "An attempt was made to add data to a hash object that is already marked 'finished'."
    Case NTE_BAD_KEY:             Return_ErrDesc = "A keyed hash algorithm (such as CALG_MAC) is specified by Algid and the hKey parameter is either zero or it specifies an invalid key handle. This error code will also be returned if the key is to a stream cipher, or if the cipher mode is anything other than CBC."
    Case NTE_BAD_KEYSET:          Return_ErrDesc = "The Registry entry for the key container could not be opened and may not exist."
    Case NTE_BAD_KEYSET_PARAM:    Return_ErrDesc = "The pszContainer or pszProvider parameter is set to an illegal value."
    Case NTE_BAD_LEN:             Return_ErrDesc = "The CRYPT_USERDATA flag is set and the dwDataLen parameter has a nonzero value."
    Case NTE_BAD_PROV_TYPE:       Return_ErrDesc = "The value of the dwProvType parameter is out of range. All provider types must be from 1 to 999, inclusive."
    Case NTE_BAD_SIGNATURE:       Return_ErrDesc = "The provider DLL signature did not verify correctly. Either the DLL or the digital signature has been tampered with."
    Case NTE_BAD_UID:             Return_ErrDesc = "The CSP context that was specified when the hash object was created cannot be found."
    Case NTE_DOUBLE_ENCRYPT:      Return_ErrDesc = "The application attempted to encrypt the same data twice."
    Case NTE_EXISTS:              Return_ErrDesc = "The dwFlags parameter is CRYPT_NEWKEYSET, but the key container already exists."
    Case NTE_FAIL:                Return_ErrDesc = "The function failed in some unexpected way."
    Case NTE_KEYSET_ENTRY_BAD:    Return_ErrDesc = "The Registry entry for the pszContainer key container was found (in the HKEY_CURRENT_USER window), but is corrupt. See the section System Administration for details about CryptoAPI’s Registry usage."
    Case NTE_KEYSET_NOT_DEF:      Return_ErrDesc = "No Registry entry exists in the HKEY_CURRENT_USER window for the key container specified by pszContainer."
    Case NTE_NO_MEMORY:           Return_ErrDesc = "The CSP ran out of memory during the operation."
    Case NTE_PROV_DLL_NOT_FOUND:  Return_ErrDesc = "The provider DLL file does not exist or is not on the current path."
    Case NTE_PROV_TYPE_ENTRY_BAD: Return_ErrDesc = "The Registry entry for the provider type specified by dwProvType is corrupt. This error may relate to either the user default CSP list or the machine default CSP list. See the section System Administration for details about CryptoAPI’s Registry usage."
    Case NTE_PROV_TYPE_NO_MATCH:  Return_ErrDesc = "The provider type specified by dwProvType does not match the provider type found in the Registry. Note that this error can only occur when pszProvider specifies an actual CSP name."
    Case NTE_PROV_TYPE_NOT_DEF:   Return_ErrDesc = "No Registry entry exists for the provider type specified by dwProvType."
    Case NTE_PROVIDER_DLL_FAIL:   Return_ErrDesc = "The provider DLL file could not be loaded, and may not exist. If it exists, then the file is not a valid DLL."
    Case NTE_SIGNATURE_FILE_BAD:  Return_ErrDesc = "An error occurred while loading the DLL file image, prior to verifying its signature."
    Case Else:                    Return_ErrDesc = "Unknown Error"
  End Select
  
  ' Display the error
  If ShowErrorMsg = True Then
    MsgBox "The " & NameOfLastAPICalled & " Windows API caused the following error:" & Chr(13) & Chr(13) & "Error Number = " & CStr(ErrorNumber) & Chr(13) & "Error Description = " & Return_ErrDesc, vbOKOnly + vbExclamation, "  Windows API Error"
  End If
  
End Function

Private Function ContainsNonNumeric(ByVal StringToCheck As String) As Boolean
  
  Dim MyCounter As Long
  Dim CharLeft  As String
  
  For MyCounter = 1 To Len(StringToCheck)
    CharLeft = Left(StringToCheck, MyCounter)
    If IsNumeric(Right(CharLeft, 1)) = False Then
      ContainsNonNumeric = True
      Exit For
    End If
  Next
  
End Function
