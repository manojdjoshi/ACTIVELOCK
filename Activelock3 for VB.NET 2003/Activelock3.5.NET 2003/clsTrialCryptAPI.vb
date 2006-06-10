Option Strict Off
Option Explicit On
Friend Class clsCryptAPI
	'*   ActiveLock
	'*   Copyright 1998-2002 Nelson Ferraz
	'*   Copyright 2005 The ActiveLock Software Group (ASG)
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
	' Name: clsCryptAPI
	' Purpose: Encryption/Decryption Class
	' Remarks: 'Information concerning the CryptAPI encryption/decryption can probably
	' be found somewhere on Microsoft homepage http://www.microsoft.com/
	' <p>This class is used by the Trial Period/Runs feature
	' Functions:
	' Properties:
	' Methods:
    ' Started: 08.05.2005
    ' Modified: 03.24.2006
	'===============================================================================

	Private m_Key As String
	
	Private Declare Function CryptAcquireContext Lib "advapi32.dll"  Alias "CryptAcquireContextA"(ByRef phProv As Integer, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Integer, ByVal dwFlags As Integer) As Integer
	Private Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Integer, ByVal Algid As Integer, ByVal hKey As Integer, ByVal dwFlags As Integer, ByRef phHash As Integer) As Integer
	Private Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Integer, ByVal pbData As String, ByVal dwDataLen As Integer, ByVal dwFlags As Integer) As Integer
	Private Declare Function CryptDeriveKey Lib "advapi32.dll" (ByVal hProv As Integer, ByVal Algid As Integer, ByVal hBaseData As Integer, ByVal dwFlags As Integer, ByRef phKey As Integer) As Integer
	Private Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Integer) As Integer
	Private Declare Function CryptEncrypt Lib "advapi32.dll" (ByVal hKey As Integer, ByVal hHash As Integer, ByVal Final As Integer, ByVal dwFlags As Integer, ByVal pbData As String, ByRef pdwDataLen As Integer, ByVal dwBufLen As Integer) As Integer
	Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Integer) As Integer
	Private Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Integer, ByVal dwFlags As Integer) As Integer
	Private Declare Function CryptDecrypt Lib "advapi32.dll" (ByVal hKey As Integer, ByVal hHash As Integer, ByVal Final As Integer, ByVal dwFlags As Integer, ByVal pbData As String, ByRef pdwDataLen As Integer) As Integer
	
	Private Const SERVICE_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
	Private Const KEY_CONTAINER As String = "Metallica"
	Private Const PROV_RSA_FULL As Integer = 1
	Private Const CRYPT_NEWKEYSET As Integer = 8
	Private Const ALG_CLASS_DATA_ENCRYPT As Integer = 24576
	Private Const ALG_CLASS_HASH As Integer = 32768
	Private Const ALG_TYPE_ANY As Integer = 0
	Private Const ALG_TYPE_STREAM As Integer = 2048
	Private Const ALG_SID_RC4 As Integer = 1
	Private Const ALG_SID_MD5 As Integer = 3
	Private Const CALG_MD5 As Integer = ((ALG_CLASS_HASH Or ALG_TYPE_ANY) Or ALG_SID_MD5)
	Private Const CALG_RC4 As Integer = ((ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_STREAM) Or ALG_SID_RC4)
	Private Const ENCRYPT_ALGORITHM As Integer = CALG_RC4
    '===============================================================================
	' Name: Sub EncryptByte
	' Input:
	'   ByRef ByteArray As Byte - Input byte array
	'   ByRef password As String - Input password if used
	' Output: None
	' Purpose: Converts an array into a string, encrypts it and then converts it back to an array
	' Remarks: None
	'===============================================================================
	' @param ByteArray  Input byte array
	' @param password   Input optional password
	Public Sub EncryptByte(ByRef ByteArray() As Byte, Optional ByRef password As String = "")
		
        Dim myString As String
        Dim asciiEnc As System.Text.Encoding = System.Text.Encoding.ASCII
        Dim unicodeEnc As System.Text.Encoding = System.Text.Encoding.Unicode
        myString = System.Text.UnicodeEncoding.Unicode.GetString(ByteArray)
        Dim asciiChars(asciiEnc.GetCharCount(ByteArray, 0, myString.Length)) As Char
        asciiEnc.GetChars(ByteArray, 0, myString.Length, asciiChars, 0)
        Dim asciiString As New String(asciiChars)
        Dim encryptedString As String
        encryptedString = EncryptString(asciiString, password)
        Dim unicodebytes As Byte() = unicodeEnc.GetBytes(encryptedString)
        Dim asciiBytes As Byte() = System.Text.Encoding.Convert(unicodeEnc, asciiEnc, unicodebytes)
        ByteArray = asciiBytes

	End Sub
    '===============================================================================
	' Name: Function EncryptString
	' Input:
	'   ByRef Text As String - String to be encrypted
	'   ByRef password As String - Optional password used in encryption
	' Output:
	'   String - Encrypted string
	' Purpose: Encrypts a string with an optional password
	' Remarks: None
	'===============================================================================
	Public Function EncryptString(ByRef Text As String, Optional ByRef password As String = "") As String
		
		'Set the new key if any was sent to the function
		If (Len(password) > 0) Then Key = password
		
		'Return the encrypted data
		EncryptString = EncryptDecrypt(Text, True)
		
	End Function
    '===============================================================================
	' Name: Sub DecryptByte
	' Input:
	'   ByRef ByteArray As Byte - Byte array input
	'   ByRef password As String - Password if used any
	' Output: None
	' Purpose: Converts an array into a string, decrypts it and then converts it back to an array
	' Remarks: None
	'===============================================================================
	Public Sub DecryptByte(ByRef ByteArray() As Byte, Optional ByRef password As String = "")
		
        Dim myString As String
        Dim asciiDec As System.Text.Encoding = System.Text.Encoding.ASCII
        Dim unicodeDec As System.Text.Encoding = System.Text.Encoding.Unicode
        myString = System.Text.UnicodeEncoding.Unicode.GetString(ByteArray)
        Dim asciiChars(asciiDec.GetCharCount(ByteArray, 0, myString.Length)) As Char
        asciiDec.GetChars(ByteArray, 0, myString.Length, asciiChars, 0)
        Dim asciiString As New String(asciiChars)
        Dim decryptedString As String
        decryptedString = DecryptString(asciiString, password)
        Dim unicodebytes As Byte() = unicodeDec.GetBytes(decryptedString)
        Dim asciiBytes As Byte() = System.Text.Encoding.Convert(unicodeDec, asciiDec, unicodebytes)
        ByteArray = asciiBytes

	End Sub
    '===============================================================================
	' Name: Function DecryptString
	' Input:
	'   ByRef Text As String - Input text to be decrypted
	'   ByRef password As String - Input password if used
	' Output:
	'   String - Decrypted string
	' Purpose: Decrypts a given string and then converts it back to an array
	' Remarks: None
	'===============================================================================
	Public Function DecryptString(ByRef Text As String, Optional ByRef password As String = "") As String
		On Error GoTo ErrHandler
		
		'Set the new key if any was sent to the function
		If (Len(password) > 0) Then Key = password
		
		'Return the decrypted data
		DecryptString = EncryptDecrypt(Text, False)
		
ErrHandler: 
		If Err.Number <> 0 Then
			DecryptString = ""
		End If
	End Function
    '===============================================================================
	' Name: Sub EncryptFile
	' Input:
	'   ByRef SourceFile As String - Source file to be encrypted
	'   ByRef DestFile As String - Destination file name to store the encrypted data
	'   ByRef Key As String - Optional Key used in encryption
	' Output: None
	' Purpose: Encrypts the contents of a file and stores them into another file
	' Remarks: None
	'===============================================================================
	Public Sub EncryptFile(ByRef SourceFile As String, ByRef DestFile As String, Optional ByRef Key As String = "")
		
		Dim Filenr As Short
		Dim ByteArray() As Byte
		
		'Make sure the source file do exist
		If (Not fileExist(SourceFile)) Then
			'Call Err.Raise(vbObjectError, , "Error in Skipjack EncryptFile procedure (Source file does not exist).")
			Exit Sub
		End If
		
		'Open the source file and read the content
		'into a bytearray to pass onto encryption
		Filenr = FreeFile
		FileOpen(Filenr, SourceFile, OpenMode.Binary)
		ReDim ByteArray(LOF(Filenr) - 1)
        FileGet(Filenr, ByteArray)
		FileClose(Filenr)
		
		'Encrypt the bytearray
		Call EncryptByte(ByteArray, Key)
		
		'If the destination file already exist we need
		'to delete it since opening it for binary use
		'will preserve it if it already exist
		If (fileExist(DestFile)) Then Kill(DestFile)
		
		'Store the encrypted data in the destination file
		Filenr = FreeFile
		FileOpen(Filenr, DestFile, OpenMode.Binary)
        FilePut(Filenr, ByteArray)
		FileClose(Filenr)
		
	End Sub
    '===============================================================================
	' Name: Sub DecryptFile
	' Input:
	'   ByRef SourceFile As String - Full path and name of the source file
	'   ByRef DestFile As String - Full path and name of the destination file
	'   ByRef Key As String - Optional Key
	' Output: None
	' Purpose: Opens a source file, reads its contents into a byte array, decrypts it and saves it as another file
	' Remarks: None
	'===============================================================================
	Public Sub DecryptFile(ByRef SourceFile As String, ByRef DestFile As String, Optional ByRef Key As String = "")
		
		Dim Filenr As Short
		Dim ByteArray() As Byte
		
		'Make sure the source file do exist
		If (Not fileExist(SourceFile)) Then
			'Call Err.Raise(vbObjectError, , "Error in Skipjack EncryptFile procedure (Source file does not exist).")
			Exit Sub
		End If
		
		'Open the source file and read the content
		'into a bytearray to decrypt
		Filenr = FreeFile
		FileOpen(Filenr, SourceFile, OpenMode.Binary)
		ReDim ByteArray(LOF(Filenr) - 1)
        FileGet(Filenr, ByteArray)
		FileClose(Filenr)
		
		'Decrypt the bytearray
		Call DecryptByte(ByteArray, Key)
		
		'If the destination file already exist we need
		'to delete it since opening it for binary use
		'will preserve it if it already exist
		If (fileExist(DestFile)) Then Kill(DestFile)
		
		'Store the decrypted data in the destination file
		Filenr = FreeFile
		FileOpen(Filenr, DestFile, OpenMode.Binary)
        FilePut(Filenr, ByteArray)
		FileClose(Filenr)
		
	End Sub
    '===============================================================================
	' Name: Function EncryptDecrypt
	' Input:
	'   ByVal Text As String - String to be encrypted or decrypted
	'   ByRef Encrypt As Boolean - Encrypted if True, decrypted if False
	' Output:
	'   String - Encrypted or decrypted string
	' Purpose: Encrypts or Decrypts a string based on the specified parameter
	' Remarks: None
	'===============================================================================
	Private Function EncryptDecrypt(ByVal Text As String, ByRef Encrypt As Boolean) As String
		
		Dim hKey As Integer
		Dim hHash As Integer
		Dim lLength As Integer
		Dim hCryptProv As Integer
		
		'Get handle to CSP
		If (CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, CRYPT_NEWKEYSET) = 0) Then
			If (CryptAcquireContext(hCryptProv, KEY_CONTAINER, SERVICE_PROVIDER, PROV_RSA_FULL, 0) = 0) Then
				'Call Err.Raise(vbObjectError, , "Error during CryptAcquireContext for a new key container." & vbCrLf & "A container with this name probably already exists.")
			End If
		End If
		
		'Create a hash object to calculate a session
		'key from the password (instead of encrypting
		'with the actual key)
		If (CryptCreateHash(hCryptProv, CALG_MD5, 0, 0, hHash) = 0) Then
			'Call Err.Raise(vbObjectError, , "Could not create a Hash Object (CryptCreateHash API)")
		End If
		
		'Hash the password
		If (CryptHashData(hHash, m_Key, Len(m_Key), 0) = 0) Then
			'Call Err.Raise(vbObjectError, , "Could not calculate a Hash Value (CryptHashData API)")
		End If
		
		'Derive a session key from the hash object
		If (CryptDeriveKey(hCryptProv, ENCRYPT_ALGORITHM, hHash, 0, hKey) = 0) Then
			'Call Err.Raise(vbObjectError, , "Could not create a session key (CryptDeriveKey API)")
		End If
		
		'Encrypt or decrypt depending on the Encrypt parameter
		lLength = Len(Text)
		If (Encrypt) Then
			If (CryptEncrypt(hKey, 0, 1, 0, Text, lLength, lLength) = 0) Then
				'Call Err.Raise(vbObjectError, , "Error during CryptEncrypt.")
			End If
		Else
			If (CryptDecrypt(hKey, 0, 1, 0, Text, lLength) = 0) Then
				'Call Err.Raise(vbObjectError, , "Error during CryptDecrypt.")
			End If
		End If
		
		'Return the encrypted/decrypted data
		EncryptDecrypt = Left(Text, lLength)
		
		'Destroy the session key
		If (hKey <> 0) Then Call CryptDestroyKey(hKey)
		
		'Destroy the hash object
		If (hHash <> 0) Then Call CryptDestroyHash(hHash)
		
		'Release provider handle
		If (hCryptProv <> 0) Then Call CryptReleaseContext(hCryptProv, 0)
		
	End Function
    '===============================================================================
	' Name: Property Let Key
	' Input:
	'   Byref New_Value As String - Value to be set
	' Output: None
	' Purpose: Sets the key value used in encryption
	' Remarks: None
	'===============================================================================
	Public WriteOnly Property Key() As String
		Set(ByVal Value As String)
			
			'Do nothing if no change was made
			If (m_Key = Value) Then Exit Property
			
			'Set the new key
			m_Key = Value
			
		End Set
	End Property
End Class