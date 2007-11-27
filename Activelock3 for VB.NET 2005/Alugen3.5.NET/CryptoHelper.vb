Imports System 
Imports System.IO 
Imports System.Text 
Imports System.Security.Cryptography 
Imports System.Security.Permissions 
Namespace NCrypto.Security.Cryptography 

 Public NotInheritable Class CryptoHelper 
   Private Shared _rng As RNGCryptoServiceProvider = New RNGCryptoServiceProvider 
   Private Shared _defaultCspParameters As CspParameters = New CspParameters 

   Private Sub New() 
   End Sub 

   Shared Sub New() 
     _defaultCspParameters.KeyContainerName = AppDomain.CurrentDomain.FriendlyName 
     If Environment.UserInteractive Then 
       _defaultCspParameters.Flags = CspProviderFlags.UseDefaultKeyContainer 
     Else 
       _defaultCspParameters.Flags = CspProviderFlags.UseMachineKeyStore 
     End If 
   End Sub 

   Public Shared ReadOnly Property DefaultCspParameters() As CspParameters 
     Get 
       Return _defaultCspParameters 
     End Get 
   End Property 

   Public Shared ReadOnly Property RsaInstance() As RSACryptoServiceProvider 
     Get 
                Call (AddressOf Sign).Demand()
       call (New SecurityPermission(SecurityPermissionFlag.UnmanagedCode)).Assert 
       Dim rsa As RSACryptoServiceProvider = New RSACryptoServiceProvider(_defaultCspParameters) 
       rsa.PersistKeyInCsp = True 
       Return rsa 
     End Get 
   End Property 

   Public Shared Function Encrypt(ByVal plainText As String) As String 
     Return ProtectedData.Protect(plainText) 
   End Function 

   Public Shared Function Encrypt(ByVal plainText As Byte()) As Byte() 
     Return ProtectedData.Protect(plainText) 
   End Function 

   Public Shared Function Encrypt(ByVal plainText As String, ByVal password As String) As String 
     Dim algorithm As SymmetricAlgorithm = New RijndaelManaged 
     Try 
       Return Encrypt(plainText, password, algorithm) 
     Finally 
       algorithm.Clear 
     End Try 
   End Function 

   Public Shared Function Encrypt(ByVal plainText As String, ByVal password As String, ByVal algorithm As SymmetricAlgorithm) As String 
     If plainText Is Nothing Then 
       Throw New ArgumentNullException("plainText") 
     End If 
     If password Is Nothing OrElse password.Length = 0 Then 
       Throw New ArgumentNullException("password") 
     End If 
     If algorithm Is Nothing Then 
       Throw New ArgumentNullException("algorithm") 
     End If 
     Dim userData As Byte() = Utility.DefaultEncoding.GetBytes(plainText) 
     Try 
       ' Using 
       Dim keys As DerivedKeys = New DerivedKeys(password, algorithm) 
       Try 
         algorithm.Key = keys.Key 
         algorithm.IV = keys.IV 
         Dim cipherText As Byte() = Encrypt(userData, algorithm) 
         Return Convert.ToBase64String(Utility.JoinArrays(keys.Salt, cipherText)) 
       Finally 
         CType(keys, IDisposable).Dispose() 
       End Try 
     Finally 
       Array.Clear(userData, 0, userData.Length) 
     End Try 
   End Function 

   Public Shared Function Encrypt(ByVal plainText As Byte(), ByVal algorithm As SymmetricAlgorithm) As Byte() 
     If plainText Is Nothing Then 
       Throw New ArgumentNullException("plainText") 
     End If 
     If algorithm Is Nothing Then 
       Throw New ArgumentNullException("algorithm") 
     End If 
     call (AddressOf Encrypt).Demand 
     Dim ms As MemoryStream = New MemoryStream 
     ' Using 
     Dim cs As CryptoStream = New CryptoStream(ms, algorithm.CreateEncryptor, CryptoStreamMode.Write) 
     Try 
       cs.Write(plainText, 0, plainText.Length) 
       cs.Close 
       Return ms.ToArray 
     Finally 
       CType(cs, IDisposable).Dispose() 
     End Try 
   End Function 

   Public Shared Function Encrypt(ByVal plainText As String, ByVal parameters As RSAParameters) As String 
     Return Utility.ToHexString(Encrypt(Utility.DefaultEncoding.GetBytes(plainText), parameters, False)) 
   End Function 

   Public Shared Function Encrypt(ByVal plainText As Byte(), ByVal parameters As RSAParameters, ByVal oaep As Boolean) As Byte() 
     If plainText Is Nothing Then 
       Throw New ArgumentNullException("plainText") 
     End If 
     If parameters.Modulus Is Nothing OrElse parameters.Exponent Is Nothing Then 
       Throw New ArgumentNullException("parameters.Modulus/parameters.Exponent") 
     End If 
     call (AddressOf Encrypt).Demand 
     Dim rsa As RSACryptoServiceProvider = RsaInstance 
     Try 
       rsa.ImportParameters(parameters) 
       Dim blockSize As Integer = (rsa.KeySize >> 3) - 11 
       If plainText.Length <= blockSize Then 
         Return rsa.Encrypt(plainText, oaep) 
       Else 
         Dim modulusSize As Integer = blockSize + 11 
         ' Using 
         Dim msin As MemoryStream = New MemoryStream(plainText) 
         Try 
           ' Using 
           Dim msout As MemoryStream = New MemoryStream(blockSize) 
           Try 
             Dim buffer(blockSize) As Byte 
             Dim bytesRead As Integer 
             Do 
               bytesRead = msin.Read(buffer, 0, blockSize) 
               If bytesRead = blockSize Then 
                 msout.Write(rsa.Encrypt(buffer, oaep), 0, modulusSize) 
               Else 
                 Dim final(bytesRead) As Byte 
                 Array.Copy(buffer, final, bytesRead) 
                 msout.Write(rsa.Encrypt(final, oaep), 0, modulusSize) 
               End If 
             Loop While bytesRead = blockSize 
             Return msout.ToArray 
           Finally 
             CType(msout, IDisposable).Dispose() 
           End Try 
         Finally 
           CType(msin, IDisposable).Dispose() 
         End Try 
       End If 
     Finally 
       rsa.Clear 
     End Try 
   End Function 

   Public Shared Function Decrypt(ByVal cipherText As String) As String 
     Return ProtectedData.Unprotect(cipherText) 
   End Function 

   Public Shared Function Decrypt(ByVal cipherText As Byte()) As Byte() 
     Return ProtectedData.Unprotect(cipherText, Nothing, DataProtectionScope.LocalMachine) 
   End Function 

   Public Shared Function Decrypt(ByVal cipherText As String, ByVal password As String) As String 
     Dim algorithm As SymmetricAlgorithm = New RijndaelManaged 
     Try 
       Return Decrypt(cipherText, password, algorithm) 
     Finally 
       algorithm.Clear 
     End Try 
   End Function 

   Public Shared Function Decrypt(ByVal cipherText As String, ByVal password As String, ByVal algorithm As SymmetricAlgorithm) As String 
     If cipherText Is Nothing OrElse cipherText.Length = 0 Then 
       Throw New ArgumentNullException("cipherText") 
     End If 
     If password Is Nothing Then 
       Throw New ArgumentNullException("password") 
     End If 
     If algorithm Is Nothing Then 
       Throw New ArgumentNullException("algorithm") 
     End If 
     If Not Utility.IsBase64Encoding(cipherText) Then 
       Throw New ArgumentException(Resource.ResourceManager(Resource.MessageKey.Base64EncodingException, cipherText)) 
     End If 
     Dim cipherBytes As Byte() = Convert.FromBase64String(cipherText) 
     Dim clearText As Byte() = Nothing 
     Try 
       ' Using 
       Dim keys As DerivedKeys = New DerivedKeys(password, algorithm, cipherBytes) 
       Try 
         algorithm.Key = keys.Key 
         algorithm.IV = keys.IV 
         clearText = Decrypt(cipherBytes, algorithm) 
         Return Utility.DefaultEncoding.GetString(clearText) 
       Finally 
         CType(keys, IDisposable).Dispose() 
       End Try 
     Finally 
       If Not (clearText Is Nothing) Then 
         Array.Clear(clearText, 0, clearText.Length) 
       End If 
     End Try 
   End Function 

   Public Shared Function Decrypt(ByVal cipherText As Byte(), ByVal algorithm As SymmetricAlgorithm) As Byte() 
     If cipherText Is Nothing OrElse cipherText.Length = 0 Then 
       Throw New ArgumentNullException("cipherText") 
     End If 
     If algorithm Is Nothing Then 
       Throw New ArgumentNullException("algorithm") 
     End If 
     call (AddressOf Decrypt).Demand 
     Dim ms As MemoryStream = New MemoryStream 
     ' Using 
     Dim cs As CryptoStream = New CryptoStream(ms, algorithm.CreateDecryptor, CryptoStreamMode.Write) 
     Try 
       cs.Write(cipherText, 0, cipherText.Length) 
       cs.Close 
       Return ms.ToArray 
     Finally 
       CType(cs, IDisposable).Dispose() 
     End Try 
   End Function 

   Public Shared Function Decrypt(ByVal cipherText As String, ByVal parameters As RSAParameters) As String 
     Return Utility.DefaultEncoding.GetString(Decrypt(Utility.FromHexString(cipherText), parameters, False)) 
   End Function 

   Public Shared Function Decrypt(ByVal cipherText As Byte(), ByVal parameters As RSAParameters, ByVal oaep As Boolean) As Byte() 
     If cipherText Is Nothing OrElse cipherText.Length = 0 Then 
       Throw New ArgumentNullException("cipherText") 
     End If 
     If parameters.Modulus Is Nothing OrElse parameters.Exponent Is Nothing Then 
       Throw New ArgumentNullException("parameters.Modulus/parameters.Exponent") 
     End If 
     call (AddressOf Decrypt).Demand 
     Dim rsa As RSACryptoServiceProvider = RsaInstance 
     Try 
       rsa.ImportParameters(parameters) 
       Dim modulusSize As Integer = rsa.KeySize >> 3 
       If cipherText.Length = modulusSize Then 
         Return rsa.Decrypt(cipherText, oaep) 
       Else 
         If Not (cipherText.Length Mod modulusSize = 0) Then 
           Throw New ArgumentOutOfRangeException("cipherText", Resource.ResourceManager(Resource.MessageKey.InvalidAsymmetricDataSize, modulusSize)) 
         Else 
           ' Using 
           Dim msin As MemoryStream = New MemoryStream(cipherText) 
           Try 
             ' Using 
             Dim msout As MemoryStream = New MemoryStream(modulusSize) 
             Try 
               Dim buffer(modulusSize) As Byte 
               Dim bytesRead As Integer 
               Do 
                 bytesRead = msin.Read(buffer, 0, modulusSize) 
                 If bytesRead > 0 Then 
                   Dim plain As Byte() = rsa.Decrypt(buffer, oaep) 
                   msout.Write(plain, 0, plain.Length) 
                   Array.Clear(plain, 0, plain.Length) 
                 End If 
               Loop While bytesRead > 0 
               Return msout.ToArray 
             Finally 
               CType(msout, IDisposable).Dispose() 
             End Try 
           Finally 
             CType(msin, IDisposable).Dispose() 
           End Try 
         End If 
       End If 
     Finally 
       rsa.Clear 
     End Try 
   End Function 

   Public Shared Function ComputeHash(ByVal value As String) As String 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     Dim algorithm As HashAlgorithm = New SHA1Managed 
     Try 
       Return Utility.ToHexString(ComputeHash(Utility.DefaultEncoding.GetBytes(value), algorithm)) 
     Finally 
       algorithm.Clear 
     End Try 
   End Function 

   Public Shared Function ComputeHash(ByVal value As String, ByVal algorithm As HashAlgorithm) As String 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     Return Utility.ToHexString(ComputeHash(Utility.DefaultEncoding.GetBytes(value), algorithm)) 
   End Function 

   Public Shared Function ComputeHash(ByVal value As Byte(), ByVal algorithm As HashAlgorithm) As Byte() 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     If algorithm Is Nothing Then 
       Throw New ArgumentNullException("algorithm") 
     End If 
     Return algorithm.ComputeHash(value) 
   End Function 

   Public Shared Function ComputeKeyedHash(ByVal value As String, ByVal key As Byte()) As String 
     Dim algorithm As KeyedHashAlgorithm = New HMACSHA1 
     Try 
       Return ComputeKeyedHash(value, algorithm, key) 
     Finally 
       algorithm.Clear 
     End Try 
   End Function 

   Public Shared Function ComputeKeyedHash(ByVal value As String, ByVal algorithm As KeyedHashAlgorithm, ByVal key As Byte()) As String 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     Return Utility.ToHexString(ComputeKeyedHash(Utility.DefaultEncoding.GetBytes(value), algorithm, key)) 
   End Function 

   Public Shared Function ComputeKeyedHash(ByVal value As Byte(), ByVal algorithm As KeyedHashAlgorithm) As Byte() 
     Return ComputeKeyedHash(value, algorithm, algorithm.Key) 
   End Function 

   Public Shared Function ComputeKeyedHash(ByVal value As Byte(), ByVal algorithm As KeyedHashAlgorithm, ByVal key As Byte()) As Byte() 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     If algorithm Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     If key Is Nothing OrElse key.Length = 0 Then 
       Throw New ArgumentNullException("key") 
     End If 
     Dim validKeySize As Integer = algorithm.Key.Length 
     If Not (key.Length = validKeySize) Then 
       Dim pm As PKCS1MaskGenerationMethod = New PKCS1MaskGenerationMethod 
       key = pm.GenerateMask(key, validKeySize) 
     End If 
     algorithm.Key = key 
     Return algorithm.ComputeHash(value) 
   End Function 

   Public Shared Function ComputeSaltedHash(ByVal value As String) As String 
     Dim algorithm As HashAlgorithm = New SHA1Managed 
     Try 
       Return ComputeSaltedHash(value, algorithm) 
     Finally 
       algorithm.Clear 
     End Try 
   End Function 

   Public Shared Function ComputeSaltedHash(ByVal value As String, ByVal algorithm As HashAlgorithm) As String 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     Return Utility.ToHexString(ComputeSaltedHash(Utility.DefaultEncoding.GetBytes(value), algorithm)) 
   End Function 

   Public Shared Function ComputeSaltedHash(ByVal value As Byte(), ByVal algorithm As HashAlgorithm) As Byte() 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     If algorithm Is Nothing Then 
       Throw New ArgumentNullException("algorithm") 
     End If 
     Dim hashSize As Integer = (algorithm.HashSize) >> 3 
     Dim salt As Byte() = ComputeRandomBytes(hashSize) 
     Return Utility.JoinArrays(salt, ComputeHash(Utility.JoinArrays(salt, value), algorithm)) 
   End Function 

   Public Shared Function VerifySaltedHash(ByVal value As String, ByVal saltedhash As String) As Boolean 
     Dim algorithm As HashAlgorithm = New SHA1Managed 
     Try 
       Return VerifySaltedHash(value, saltedhash, algorithm) 
     Finally 
       algorithm.Clear 
     End Try 
   End Function 

   Public Shared Function VerifySaltedHash(ByVal value As String, ByVal saltedhash As String, ByVal algorithm As HashAlgorithm) As Boolean 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     If saltedhash Is Nothing Then 
       Throw New ArgumentNullException("saltedhash") 
     End If 
     Return VerifySaltedHash(Utility.DefaultEncoding.GetBytes(value), Utility.FromHexString(saltedhash), algorithm) 
   End Function 

   Public Shared Function VerifySaltedHash(ByVal value As Byte(), ByVal saltedhash As Byte(), ByVal algorithm As HashAlgorithm) As Boolean 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     If algorithm Is Nothing Then 
       Throw New ArgumentNullException("algorithm") 
     End If 
     If saltedhash Is Nothing Then 
       Throw New ArgumentNullException("saltedhash") 
     End If 
     Dim hashSize As Integer = (algorithm.HashSize) >> 3 
     Dim saltAndHash As SplitArray = Utility.SplitArrays(saltedhash, hashSize) 
     Return Utility.CompareArrays(saltAndHash.SecondArray, ComputeHash(Utility.JoinArrays(saltAndHash.FirstArray, value), algorithm)) 
   End Function 

   Public Shared Function Sign(ByVal value As String) As String 
     Dim algorithm As AsymmetricAlgorithm = RsaInstance 
     Try 
       Return Sign(value, algorithm) 
     Finally 
       algorithm.Clear 
     End Try 
   End Function 

   Public Shared Function Sign(ByVal value As String, ByVal algorithm As AsymmetricAlgorithm) As String 
     Dim hash As HashAlgorithm = New SHA1Managed 
     Try 
       Return Convert.ToBase64String(Sign(value, algorithm, hash)) 
     Finally 
       hash.Clear 
     End Try 
   End Function 

   Public Shared Function Sign(ByVal value As String, ByVal algorithm As AsymmetricAlgorithm, ByVal hash As HashAlgorithm) As Byte() 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     If algorithm Is Nothing Then 
       Throw New ArgumentNullException("algorithm") 
     End If 
     If hash Is Nothing Then 
       Throw New ArgumentNullException("hash") 
     End If 
     call (AddressOf Sign).Demand 
     Dim userData As Byte() = Utility.DefaultEncoding.GetBytes(value) 
     Try 
       Dim asf As AsymmetricSignatureFormatter = GetAsymmetricFormatter(algorithm) 
       asf.SetHashAlgorithm(hash.ToString) 
       Return asf.CreateSignature(ComputeHash(userData, hash)) 
     Finally 
       Array.Clear(userData, 0, userData.Length) 
     End Try 
   End Function 

   Public Shared Function VerifySignature(ByVal value As String, ByVal signature As String) As Boolean 
     Dim algorithm As AsymmetricAlgorithm = RsaInstance 
     Try 
       Return VerifySignature(value, signature, algorithm) 
     Finally 
       algorithm.Clear 
     End Try 
   End Function 

   Public Shared Function VerifySignature(ByVal value As String, ByVal signature As String, ByVal algorithm As AsymmetricAlgorithm) As Boolean 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     If Not Utility.IsBase64Encoding(signature) Then 
       Throw New ArgumentException(Resource.ResourceManager(Resource.MessageKey.Base64EncodingException, signature)) 
     End If 
     Dim hash As HashAlgorithm = New SHA1Managed 
     Try 
       Return VerifySignature(Utility.DefaultEncoding.GetBytes(value), Convert.FromBase64String(signature), algorithm, hash) 
     Finally 
       hash.Clear 
     End Try 
   End Function 

   Public Shared Function VerifySignature(ByVal value As Byte(), ByVal signature As Byte(), ByVal algorithm As AsymmetricAlgorithm, ByVal hash As HashAlgorithm) As Boolean 
     call (AddressOf Sign).Demand 
     If value Is Nothing Then 
       Throw New ArgumentNullException("value") 
     End If 
     If signature Is Nothing Then 
       Throw New ArgumentNullException("signature") 
     End If 
     If algorithm Is Nothing Then 
       Throw New ArgumentNullException("algorithm") 
     End If 
     If hash Is Nothing Then 
       Throw New ArgumentNullException("hash") 
     End If 
     Dim asd As AsymmetricSignatureDeformatter = GetAsymmetricDeformatter(algorithm) 
     asd.SetHashAlgorithm(hash.ToString) 
     Return asd.VerifySignature(ComputeHash(value, hash), signature) 
   End Function 

   Public Shared Function ComputeRandomBytes(ByVal randomBytes As Integer) As Byte() 
     If randomBytes <= 0 Then 
       Throw New ArgumentOutOfRangeException("randomBytes") 
     End If 
     Dim data(randomBytes) As Byte 
     _rng.GetBytes(data) 
     Return CType(data.Clone, Byte()) 
   End Function 

   Private Shared Function GetAsymmetricFormatter(ByVal key As AsymmetricAlgorithm) As AsymmetricSignatureFormatter 
     Dim asf As AsymmetricSignatureFormatter 
            If TypeOf key Is RSA Then
                asf = New RSAPKCS1SignatureFormatter(key)
            Else
                If TypeOf key Is DSA Then
                    asf = New DSASignatureFormatter(key)
                Else
                    Throw New CryptographicUnexpectedOperationException(Resource.ResourceManager(Resource.MessageKey.UnknownAsymmetricFormatter))
                End If
            End If
            Return asf
   End Function 

   Private Shared Function GetAsymmetricDeformatter(ByVal key As AsymmetricAlgorithm) As AsymmetricSignatureDeformatter 
     Dim asd As AsymmetricSignatureDeformatter 
     If TypeOf key Is RSA Then 
       asd = New RSAPKCS1SignatureDeformatter(key) 
     Else 
       If TypeOf key Is DSA Then 
         asd = New DSASignatureDeformatter(key) 
       Else 
         Throw New CryptographicUnexpectedOperationException(Resource.ResourceManager(Resource.MessageKey.UnknownAsymmetricDeformatter)) 
       End If 
     End If 
     Return asd 
   End Function 
 End Class 
End Namespace
