Imports System.Security.Cryptography
Imports System.IO
Imports System.text
Imports System.Text.Encoding
Imports System.ComponentModel

Public Class Crypto

#Region "Class Variables"
    Public Enum KeySize As Integer
        RC2 = 64
        DES = 64
        TripleDES = 192
        AES = 128
        RSA = 2048
    End Enum


    Public Enum Algorithm As Integer
        SHA1 = 0
        SHA256 = 1
        SHA384 = 2
        Rijndael = 3
        TripleDES = 4
        RSA = 5
        RC2 = 6
        DES = 7
        'DSA = 8
        MD5 = 9
        RNG = 10
        'Base64 = 11
        SHA512 = 12
    End Enum

    Public Enum EncodingType As Integer
        HEX = 0
        BASE_64 = 1
    End Enum

    'Initialization Vectors that we will use for symmetric encryption/decryption. These
    'byte arrays are completely arbitrary, and you can change them to whatever you like.
    Private Shared IV_8 As Byte() = New Byte() {2, 63, 9, 36, 235, 174, 78, 12}
    Private Shared IV_16 As Byte() = New Byte() {15, 199, 56, 77, 244, 126, 107, 239, _
                                          9, 10, 88, 72, 24, 202, 31, 108}
    Private Shared IV_24 As Byte() = New Byte() {37, 28, 19, 44, 25, 170, 122, 25, _
                                          25, 57, 127, 5, 22, 1, 66, 65, _
                                          14, 155, 224, 64, 9, 77, 18, 251}
    Private Shared IV_32 As Byte() = New Byte() {133, 206, 56, 64, 110, 158, 132, 22, _
                                          99, 190, 35, 129, 101, 49, 204, 248, _
                                          251, 243, 13, 194, 160, 195, 89, 152, _
                                          149, 227, 245, 5, 218, 86, 161, 124}

    'Salt value used to encrypt a plain text key. Again, this can be whatever you like
    Private Shared SALT_BYTES As Byte() = New Byte() {162, 27, 98, 1, 28, 239, 64, 30, 156, 102, 223}

    'File names to be used for public and private keys
    Private Const KEY_PUBLIC As String = "public.key"
    Private Const KEY_PRIVATE As String = "private.key"

    'Values used for RSA-based asymmetric encryption
    Private Const RSA_BLOCKSIZE As Integer = 58
    Private Const RSA_DECRYPTBLOCKSIZE As Integer = 128

    'Error messages
    Private Const ERR_NO_KEY As String = "No encryption key was provided"
    Private Const ERR_NO_ALGORITHM As String = "No algorithm was specified"
    Private Const ERR_NO_CONTENT As String = "No content was provided"
    Private Const ERR_INVALID_PROVIDER As String = "An invalid cryptographic provider was specified for this method"
    Private Const ERR_NO_FILE As String = "The specified file does not exist"
    Private Const ERR_INVALID_FILENAME As String = "The specified filename is invalid"
    Private Const ERR_FILE_WRITE As String = "Could not create file"
    Private Const ERR_FILE_READ As String = "Could not read file"

    'Initialization variables
    Private Shared _key As String
    Private Shared _algorithm As Algorithm
    Private Shared _content As String
    Private Shared _exception As CryptographicException
    Private Shared _encodingType As EncodingType = EncodingType.HEX
#End Region

#Region "Public Functions"
    <Description("The key that is used to encrypt and decrypt data")> _
    Public Shared Property Key() As String
        Get
            Return _key
        End Get
        Set(ByVal value As String)
            _key = value
        End Set
    End Property

    <Description("The algorithm that will be used for encryption and decryption")> _
    Public Shared Property EncryptionAlgorithm() As Algorithm
        Get
            Return _algorithm
        End Get
        Set(ByVal value As Algorithm)
            _algorithm = value
        End Set
    End Property

    <Description("The format in which content is returned after encryption, or provided for decryption")> _
    Public Shared Property Encoding() As EncodingType
        Get
            Return _encodingType
        End Get
        Set(ByVal value As EncodingType)
            _encodingType = value
        End Set
    End Property

    <Description("Encrypted content to be retrieved after an encryption call, or provided for a decryption call")> _
    Public Shared Property Content() As String
        Get
            Return _content
        End Get
        Set(ByVal Value As String)
            _content = Value
        End Set
    End Property

    <Description("If an encryption or decryption call returns false, then this will contain the exception")> _
    Public Shared ReadOnly Property CryptoException() As CryptographicException
        Get
            Return _exception
        End Get
    End Property

    <Description("Determines whether the currently specified algorithm is a hash")> _
    Public Shared ReadOnly Property IsHashAlgorithm() As Boolean
        Get
            Select Case _algorithm
                Case Algorithm.MD5, Algorithm.SHA1, Algorithm.SHA256, Algorithm.SHA384, Algorithm.SHA512
                    Return True
                Case Else
                    Return False
            End Select
        End Get
    End Property

    <Description("Encryption of a string using the 'Key' and 'EncryptionAlgorithm' properties")> _
    Public Shared Function EncryptString(ByVal Content As String) As Boolean
        Dim cipherBytes() As Byte

        Try
            cipherBytes = _Encrypt(Content)
        Catch ex As CryptographicException
            _exception = New CryptographicException(ex.Message, ex.InnerException)
            Return False
        End Try

        If _encodingType = EncodingType.HEX Then
            _content = BytesToHex(cipherBytes)
        Else
            _content = System.Convert.ToBase64String(cipherBytes)
        End If

        Return True
    End Function

    Public Shared Function DecryptString() As Boolean
        Dim encText As Byte() = Nothing
        Dim clearText As Byte() = Nothing

        Try
            clearText = _Decrypt(_content)
        Catch ex As Exception
            _exception = New CryptographicException(ex.Message, ex.InnerException)
            Return False
        End Try

        _content = UTF8.GetString(clearText)

        Return True
    End Function

    Public Shared Function EncryptFile(ByVal Filename As String, ByVal Target As String) As Boolean
        If Not File.Exists(Filename) Then
            _exception = New CryptographicException(ERR_NO_FILE)
            Return False
        End If

        'Make sure the target file can be written
        Try
            Dim fs As FileStream = File.Create(Target)
            fs.Close()
            'fs.Dispose()  ' Works with VB 2005 only
            File.Delete(Target)
        Catch ex As Exception
            _exception = New CryptographicException(ERR_FILE_WRITE)
            Return False
        End Try

        Dim inStream() As Byte
        Dim cipherBytes() As Byte

        Try
            Dim objReader As StreamReader
            Dim objFS As FileStream
            Dim objEncoding As New System.Text.ASCIIEncoding
            objFS = New FileStream(Filename, FileMode.Open)
            objReader = New StreamReader(objFS)
            inStream = objEncoding.GetBytes(objReader.ReadToEnd)
            ' The following is the VB 2005 equivalent
            'inStream = File.ReadAllBytes(Filename)
        Catch ex As Exception
            _exception = New CryptographicException(ERR_FILE_READ)
            Return False
        End Try

        Try
            cipherBytes = _Encrypt(inStream)
        Catch ex As CryptographicException
            _exception = ex
            Return False
        End Try

        Dim encodedString As String = String.Empty

        If _encodingType = EncodingType.BASE_64 Then
            encodedString = System.Convert.ToBase64String(cipherBytes)
        Else
            encodedString = BytesToHex(cipherBytes)
        End If

        Dim encodedBytes() As Byte = UTF8.GetBytes(encodedString)

        'Create the encrypted file
        Dim outStream As FileStream = File.Create(Target)

        outStream.Write(encodedBytes, 0, encodedBytes.Length)
        outStream.Close()
        'outStream.Dispose()  ' Works with VB 2005 only

        Return True
    End Function

    Public Shared Function DecryptFile(ByVal Filename As String, ByVal Target As String) As Boolean
        If Not File.Exists(Filename) Then
            _exception = New CryptographicException(ERR_NO_FILE)
            Return False
        End If

        'Make sure the target file can be written
        Try
            Dim fs As FileStream = File.Create(Target)
            fs.Close()
            'fs.Dispose()  ' Works with VB 2005 only
            File.Delete(Target)
        Catch ex As Exception
            _exception = New CryptographicException(ERR_FILE_WRITE)
            Return False
        End Try

        Dim inStream() As Byte
        Dim clearBytes() As Byte

        Try
            Dim objReader As StreamReader
            Dim objFS As FileStream
            Dim objEncoding As New System.Text.ASCIIEncoding
            objFS = New FileStream(Filename, FileMode.Open)
            objReader = New StreamReader(objFS)
            inStream = objEncoding.GetBytes(objReader.ReadToEnd)
            ' The following is the VB 2005 equivalent
            'inStream = File.ReadAllBytes(Filename)
        Catch ex As Exception
            _exception = New CryptographicException(ERR_FILE_READ)
            Return False
        End Try

        Try
            clearBytes = _Decrypt(inStream)
        Catch ex As Exception
            _exception = New CryptographicException(ex.Message, ex.InnerException)
            Return False
        End Try

        'Create the decrypted file
        Dim outStream As FileStream = File.Create(Target)
        outStream.Write(clearBytes, 0, clearBytes.Length)
        outStream.Close()
        'outStream.Dispose() ' Works with VB 2005 only

        Return True
    End Function

    Public Shared Function GenerateHash(ByVal Content As String) As Boolean
        If Content Is Nothing OrElse Content.Equals(String.Empty) Then
            _exception = New CryptographicException(ERR_NO_CONTENT)
            Return False
        End If

        If _algorithm.Equals(-1) Then
            _exception = New CryptographicException(ERR_NO_ALGORITHM)
            Return False
        End If

        Dim hashAlgorithm As HashAlgorithm = Nothing

        Select Case _algorithm
            Case Algorithm.SHA1
                hashAlgorithm = New SHA1CryptoServiceProvider
            Case Algorithm.SHA256
                hashAlgorithm = New SHA256Managed
            Case Algorithm.SHA384
                hashAlgorithm = New SHA384Managed
            Case Algorithm.SHA512
                hashAlgorithm = New SHA512Managed
            Case Algorithm.MD5
                hashAlgorithm = New MD5CryptoServiceProvider
            Case Else
                _exception = New CryptographicException(ERR_INVALID_PROVIDER)
        End Select

        Try
            Dim hash() As Byte = ComputeHash(hashAlgorithm, Content)
            If _encodingType = EncodingType.HEX Then
                _content = BytesToHex(hash)
            Else
                _content = System.Convert.ToBase64String(hash)
            End If
            hashAlgorithm.Clear()
            Return True
        Catch ex As CryptographicException
            _exception = ex
            Return False
        Finally
            hashAlgorithm.Clear()
        End Try
    End Function

    Public Shared Sub Clear()
        _algorithm = Algorithm.SHA1
        _content = String.Empty
        _key = String.Empty
        _encodingType = EncodingType.HEX
        _exception = Nothing
    End Sub

#End Region

#Region "Shared Cryptographic Functions"

    Private Shared Function _Encrypt(ByVal Content As Byte()) As Byte()
        If Not IsHashAlgorithm AndAlso _key Is Nothing Then
            Throw New CryptographicException(ERR_NO_KEY)
        End If

        If _algorithm.Equals(-1) Then
            Throw New CryptographicException(ERR_NO_ALGORITHM)
        End If

        If Content Is Nothing OrElse Content.Equals(String.Empty) Then
            Throw New CryptographicException(ERR_NO_CONTENT)
        End If


        Dim cipherBytes() As Byte = Nothing
        Dim NumBytes As Integer = 0

        If _algorithm = Algorithm.RSA Then
            'This is an asymmetric call, which has to be treated differently
            Try
                cipherBytes = RSAEncrypt(Content)
            Catch ex As CryptographicException
                Throw ex
            End Try
        Else
            Dim provider As SymmetricAlgorithm

            Select Case _algorithm
                Case Algorithm.DES
                    provider = New DESCryptoServiceProvider
                    NumBytes = KeySize.DES
                Case Algorithm.TripleDES
                    provider = New TripleDESCryptoServiceProvider
                    NumBytes = KeySize.TripleDES
                Case Algorithm.Rijndael
                    provider = New RijndaelManaged
                    NumBytes = KeySize.AES
                Case Algorithm.RC2
                    provider = New RC2CryptoServiceProvider
                    NumBytes = KeySize.RC2
                Case Else
                    Throw New CryptographicException(ERR_INVALID_PROVIDER)
            End Select

            Try
                'Encrypt the string
                cipherBytes = SymmetricEncrypt(provider, Content, _key, NumBytes)
            Catch ex As CryptographicException
                Throw New CryptographicException(ex.Message, ex.InnerException)
            Finally
                'Free any resources held by the SymmetricAlgorithm provider
                provider.Clear()
            End Try
        End If

        Return cipherBytes
    End Function

    Private Shared Function _Encrypt(ByVal Content As String) As Byte()
        Return _Encrypt(UTF8.GetBytes(Content))
    End Function

    Private Shared Function _Decrypt(ByVal Content As Byte()) As Byte()
        If Not IsHashAlgorithm AndAlso _key Is Nothing Then
            Throw New CryptographicException(ERR_NO_KEY)
        End If

        If _algorithm.Equals(-1) Then
            Throw New CryptographicException(ERR_NO_ALGORITHM)
        End If

        If Content Is Nothing OrElse Content.Length.Equals(0) Then
            Throw New CryptographicException(ERR_NO_CONTENT)
        End If

        Dim encText As String = UTF8.GetString(Content)

        If _encodingType = EncodingType.BASE_64 Then
            'We need to convert the content to Hex before decryption
            encText = BytesToHex(System.Convert.FromBase64String(encText))
        End If

        Dim clearBytes() As Byte = Nothing
        Dim NumBytes As Integer = 0

        If _algorithm = Algorithm.RSA Then
            Try
                clearBytes = RSADecrypt(encText)
            Catch ex As CryptographicException
                Throw ex
            End Try
        Else
            Dim provider As SymmetricAlgorithm

            Select Case _algorithm
                Case Algorithm.DES
                    provider = New DESCryptoServiceProvider
                    NumBytes = KeySize.DES
                Case Algorithm.TripleDES
                    provider = New TripleDESCryptoServiceProvider
                    NumBytes = KeySize.TripleDES
                Case Algorithm.Rijndael
                    provider = New RijndaelManaged
                    NumBytes = KeySize.AES
                Case Algorithm.RC2
                    provider = New RC2CryptoServiceProvider
                    NumBytes = KeySize.RC2
                Case Else
                    Throw New CryptographicException(ERR_INVALID_PROVIDER)
            End Select

            Try
                clearBytes = SymmetricDecrypt(provider, encText, _key, NumBytes)
            Catch ex As CryptographicException
                Throw ex
            Finally
                'Free any resources held by the SymmetricAlgorithm provider
                provider.Clear()
            End Try
        End If

        'Now return the plain text content
        Return clearBytes
    End Function

    Private Shared Function _Decrypt(ByVal Content As String) As Byte()
        Return _Decrypt(UTF8.GetBytes(Content))
    End Function

    Private Shared Function ComputeHash(ByVal Provider As HashAlgorithm, ByVal plainText As String) As Byte()
        'All hashing mechanisms inherit from the HashAlgorithm base class so we can use that to cast the crypto service provider
        Dim hash As Byte() = Provider.ComputeHash(UTF8.GetBytes(plainText))
        Provider.Clear()
        Return hash
    End Function

    Private Shared Function SymmetricEncrypt(ByVal Provider As SymmetricAlgorithm, ByVal plainText As Byte(), ByVal key As String, ByVal keySize As Integer) As Byte()
        'All symmetric algorithms inherit from the SymmetricAlgorithm base class, to which we can cast from the original crypto service provider
        Dim ivBytes As Byte() = Nothing
        Select Case keySize / 8 'Determine which initialization vector to use
            Case 8
                ivBytes = IV_8
            Case 16
                ivBytes = IV_16
            Case 24
                ivBytes = IV_24
            Case 32
                ivBytes = IV_32
            Case Else
                'TODO: Throw an error because an invalid key length has been passed
        End Select

        Provider.KeySize = keySize

        'Generate a secure key based on the original password by using SALT
        Dim keyStream As Byte() = DerivePassword(key, CType(keySize / 8, Integer))  ' CType not necessary in VB 2005

        'Initialize our encryptor object
        Dim trans As ICryptoTransform = Provider.CreateEncryptor(keyStream, ivBytes)

        'Perform the encryption on the textStream byte array
        Dim result As Byte() = trans.TransformFinalBlock(plainText, 0, plainText.GetLength(0))

        'Release cryptographic resources
        Provider.Clear()
        trans.Dispose()

        Return result
    End Function

    Private Shared Function SymmetricDecrypt(ByVal Provider As SymmetricAlgorithm, ByVal encText As String, ByVal key As String, ByVal keySize As Integer) As Byte()
        'All symmetric algorithms inherit from the SymmetricAlgorithm base class, to which we can cast from the original crypto service provider
        Dim ivBytes As Byte() = Nothing
        Select Case keySize / 8 'Determine which initialization vector to use
            Case 8
                ivBytes = IV_8
            Case 16
                ivBytes = IV_16
            Case 24
                ivBytes = IV_24
            Case 32
                ivBytes = IV_32
            Case Else
                'TODO: Throw an error because an invalid key length has been passed
        End Select

        'Generate a secure key based on the original password by using SALT
        Dim keyStream As Byte() = DerivePassword(key, CType(keySize / 8, Integer)) ' CType not necessary in VB 2005

        'Convert our hex-encoded cipher text to a byte array
        Dim textStream As Byte() = HexToBytes(encText)
        Provider.KeySize = keySize

        'Initialize our decryptor object
        Dim trans As ICryptoTransform = Provider.CreateDecryptor(keyStream, ivBytes)

        'Initialize the result stream
        Dim result() As Byte = Nothing

        Try
            'Perform the decryption on the textStream byte array
            result = trans.TransformFinalBlock(textStream, 0, textStream.GetLength(0))
        Catch ex As Exception
            Throw New System.Security.Cryptography.CryptographicException("The following exception occurred during decryption: " & ex.Message)
        Finally
            'Release cryptographic resources
            Provider.Clear()
            trans.Dispose()
        End Try

        Return result
    End Function

    Private Shared Function RSAEncrypt(ByVal plainText As Byte()) As Byte()
        'Make sure that the public and private key exists
        ValidateRSAKeys()
        Dim publicKey As String = GetTextFromFile(KEY_PUBLIC)
        Dim privateKey As String = GetTextFromFile(KEY_PRIVATE)

        'The RSA algorithm works on individual blocks of unencoded bytes. In this case, the
        'maximum is 58 bytes. Therefore, we are required to break up the text into blocks and
        'encrypt each one individually. Each encrypted block will give us an output of 128 bytes.
        'If we do not break up the blocks in this manner, we will throw a "key not valid for use
        'in specified state" exception

        'Get the size of the final block
        Dim lastBlockLength As Integer = plainText.Length Mod RSA_BLOCKSIZE
        Dim blockCount As Integer = CType(Math.Floor(plainText.Length / RSA_BLOCKSIZE), Integer) ' CType not necessary in VB 2005
        Dim hasLastBlock As Boolean = False
        If Not lastBlockLength.Equals(0) Then
            'We need to create a final block for the remaining characters
            blockCount += 1
            hasLastBlock = True
        End If

        'Initialize the result buffer
        Dim result() As Byte = New Byte() {}

        'Initialize the RSA Service Provider with the public key
        Dim Provider As New RSACryptoServiceProvider(KeySize.RSA)
        Provider.FromXmlString(publicKey)

        'Break the text into blocks and work on each block individually
        For blockIndex As Integer = 0 To blockCount - 1
            Dim thisBlockLength As Integer

            'If this is the last block and we have a remainder, then set the length accordingly
            If blockCount.Equals(blockIndex + 1) AndAlso hasLastBlock Then
                thisBlockLength = lastBlockLength
            Else
                thisBlockLength = RSA_BLOCKSIZE
            End If
            Dim startChar As Integer = blockIndex * RSA_BLOCKSIZE

            'Define the block that we will be working on
            Dim currentBlock(thisBlockLength - 1) As Byte
            Array.Copy(plainText, startChar, currentBlock, 0, thisBlockLength)

            'Encrypt the current block and append it to the result stream
            Dim encryptedBlock() As Byte = Provider.Encrypt(currentBlock, False)
            Dim originalResultLength As Integer = result.Length

            ReDim Preserve result(originalResultLength + encryptedBlock.Length) ' This is for VB 2005
            'Array.Resize(result, originalResultLength + encryptedBlock.Length)

            encryptedBlock.CopyTo(result, originalResultLength)
        Next

        'Release any resources held by the RSA Service Provider
        Provider.Clear()

        Return result
    End Function

    Private Shared Function RSADecrypt(ByVal encText As String) As Byte()
        'Make sure that the public and private key exists
        ValidateRSAKeys()
        Dim publicKey As String = GetTextFromFile(KEY_PUBLIC)
        Dim privateKey As String = GetTextFromFile(KEY_PRIVATE)

        'When we encrypt a string using RSA, it works on individual blocks of up to
        '58 bytes. Each block generates an output of 128 encrypted bytes. Therefore, to
        'decrypt the message, we need to break the encrypted stream into individual
        'chunks of 128 bytes and decrypt them individually

        'Determine how many bytes are in the encrypted stream. The input is in hex format,
        'so we have to divide it by 2
        Dim maxBytes As Integer = CType(encText.Length / 2, Integer)  ' CType not necessary in VB 2005

        'Ensure that the length of the encrypted stream is divisible by 128
        If Not (maxBytes Mod RSA_DECRYPTBLOCKSIZE).Equals(0) Then
            Throw New System.Security.Cryptography.CryptographicException("Encrypted text is an invalid length")
            Return Nothing
        End If

        'Calculate the number of blocks we will have to work on
        Dim blockCount As Integer = CType(maxBytes / RSA_DECRYPTBLOCKSIZE, Integer)

        'Initialize the result buffer
        Dim result() As Byte = New Byte() {}

        'Initialize the RSA Service Provider
        Dim Provider As New RSACryptoServiceProvider(KeySize.RSA)
        Provider.FromXmlString(privateKey)

        'Iterate through each block and decrypt it
        For blockIndex As Integer = 0 To blockCount - 1
            'Get the current block to work on
            Dim currentBlockHex As String = encText.Substring(blockIndex * (RSA_DECRYPTBLOCKSIZE * 2), RSA_DECRYPTBLOCKSIZE * 2)
            Dim currentBlockBytes As Byte() = HexToBytes(currentBlockHex)

            'Decrypt the current block and append it to the result stream
            Dim currentBlockDecrypted() As Byte = Provider.Decrypt(currentBlockBytes, False)
            Dim originalResultLength As Integer = result.Length

            ReDim Preserve result(originalResultLength + currentBlockDecrypted.Length)
            'Array.Resize(result, originalResultLength + currentBlockDecrypted.Length) ' This is for VB 2005

            currentBlockDecrypted.CopyTo(result, originalResultLength)
        Next

        'Release all resources held by the RSA service provider
        Provider.Clear()

        Return result
    End Function

#End Region

#Region "Utility Functions"
    '********************************************************
    '* BytesToHex: Converts a byte array to a hex-encoded
    '*             string
    '********************************************************
    Private Shared Function BytesToHex(ByVal bytes() As Byte) As String
        Dim hex As New StringBuilder
        For n As Integer = 0 To bytes.Length - 1
            hex.AppendFormat("{0:X2}", bytes(n))
        Next
        Return hex.ToString
    End Function

    '********************************************************
    '* HexToBytes: Converts a hex-encoded string to a
    '*             byte array
    '********************************************************
    Private Shared Function HexToBytes(ByVal Hex As String) As Byte()
        Dim numBytes As Integer = CType(Hex.Length / 2, Integer)  ' CType not necessary in VB 2005
        Dim bytes(numBytes - 1) As Byte
        For n As Integer = 0 To numBytes - 1
            Dim hexByte As String = Hex.Substring(n * 2, 2)
            bytes(n) = CType(Integer.Parse(hexByte, Globalization.NumberStyles.HexNumber), Byte) ' CType not necessary with VB 2005
        Next
        Return bytes
    End Function

    '********************************************************
    '* ClearBuffer: Clears a byte array to ensure
    '*              that it cannot be read from memory
    '********************************************************
    Private Shared Sub ClearBuffer(ByVal bytes() As Byte)
        If bytes Is Nothing Then Exit Sub
        For n As Integer = 0 To bytes.Length - 1
            bytes(n) = 0
        Next
    End Sub

    '********************************************************
    '* GenerateSalt: No, this is not a culinary routine. This
    '*               generates a random salt value for
    '*               password generation
    '********************************************************
    Private Shared Function GenerateSalt(ByVal saltLength As Integer) As Byte()
        Dim salt() As Byte
        If saltLength > 0 Then
            salt = New Byte(saltLength) {}
        Else
            salt = New Byte(0) {}
        End If

        Dim seed As New RNGCryptoServiceProvider
        seed.GetBytes(salt)
        Return salt
    End Function

    '********************************************************
    '* DerivePassword: This takes the original plain text key
    '*                 and creates a secure key using SALT
    '********************************************************
    Private Shared Function DerivePassword(ByVal originalPassword As String, ByVal passwordLength As Integer) As Byte()

        ' The following section works with VB 2005 only
        'Dim derivedBytes As New Rfc2898DeriveBytes(originalPassword, SALT_BYTES, 5)
        'Return derivedBytes.GetBytes(passwordLength)

        Dim derivedBytes As New PasswordDeriveBytes(originalPassword, SALT_BYTES)
        Return derivedBytes.GetBytes(passwordLength)

    End Function

    '********************************************************
    '* ValidateRSAKeys: Checks for the existence of a public
    '*                  and private key file and creates them
    '*                  if they do not exist
    '********************************************************
    Private Shared Sub ValidateRSAKeys()
        If Not File.Exists(KEY_PRIVATE) OrElse Not File.Exists(KEY_PUBLIC) Then
            'Dim rsa As New RSACryptoServiceProvider
            Dim key As RSA = RSA.Create
            key.KeySize = KeySize.RSA
            Dim privateKey As String = key.ToXmlString(True)
            Dim publicKey As String = key.ToXmlString(False)
            Dim privateFile As StreamWriter = File.CreateText(KEY_PRIVATE)
            privateFile.Write(privateKey)
            privateFile.Close()
            'privateFile.Dispose()  ' Works with VB 2005 only
            Dim publicFile As StreamWriter = File.CreateText(KEY_PUBLIC)
            publicFile.Write(publicKey)
            publicFile.Close()
            'publicFile.Dispose()   ' Works with VB 2005 only
        End If
    End Sub

    '********************************************************
    '* GetTextFromFile: Reads the text from a file
    '********************************************************
    Private Shared Function GetTextFromFile(ByVal fileName As String) As String
        If File.Exists(fileName) Then
            Dim textFile As StreamReader = File.OpenText(fileName)
            Dim result As String = textFile.ReadToEnd
            textFile.Close()
            'textFile.Dispose() ' Works with VB 2005 only
            Return result
        Else
            Throw New IOException("Specified file does not exist")
            Return Nothing
        End If
    End Function
#End Region

End Class


