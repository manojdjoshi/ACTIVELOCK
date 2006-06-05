Option Strict Off
Option Explicit On
Module modActiveGlobals
	
	
	' Performs RSA signing of <code>strData</code> using the specified key.
	' @param strPub     RSA Public key blob
	' @param strPriv    RSA Private key blob
	' @param strData    Data to be signed
	' @return           Signature string.
	' 05.13.05    - alkan  - Removed the modActiveLock references
	'
	Public Function RSASign(ByVal strPub As String, ByVal strPriv As String, ByVal strdata As String) As String
		Dim Key As RSAKey
        ReDim Key.data(32)
        ' create the key from the key blobs
		rsa_createkey(strPub, Len(strPub), strPriv, Len(strPriv), Key)
		' sign the data using the created key
		Dim sLen As Integer
		rsa_sign(Key, strdata, Len(strdata), vbNullString, sLen)
		Dim strSig As String : strSig = New String(Chr(0), sLen)
		rsa_sign(Key, strdata, Len(strdata), strSig, sLen)
		' throw away the key
		rsa_freekey(Key)
		RSASign = strSig
	End Function
	
	' Verifies an RSA signature.
	' @param strPub     Public key blob
	' @param strData    Data to be signed
	' @param strSig     Private key blob
	' @return           Zero if verification is successful; Non-zero otherwise.
	'
	Public Function RSAVerify(ByVal strPub As String, ByVal strdata As String, ByVal strSig As String) As Integer
		Dim Key As RSAKey
		Dim rc As Integer
		' create the key from the public key blob
		rsa_createkey(strPub, Len(strPub), vbNullString, 0, Key)
		' validate the key
		rc = rsa_verifysig(Key, strSig, Len(strSig), strdata, Len(strdata))
		' de-allocate memory used by the key
		rsa_freekey(Key)
		RSAVerify = rc
	End Function
End Module