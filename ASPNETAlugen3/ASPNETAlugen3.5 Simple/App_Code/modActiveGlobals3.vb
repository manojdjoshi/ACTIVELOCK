Option Strict Off
Option Explicit On

Namespace ASPNETAlugen3

    Public Module modActiveGlobals

        Public Const STRRSAERROR As String = "Internal RSA Error."
        Public Const RETVAL_ON_ERROR As Integer = -999
        Public Const STRWRONGIPADDRESS As String = "Wrong IP Address."

        ' The following constants and declares are used to Get/Set Locale Date format
        Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String, ByVal cchData As Integer) As Integer
        Private Declare Function SetLocaleInfo Lib "kernel32" Alias "SetLocaleInfoA" (ByVal Locale As Integer, ByVal LCType As Integer, ByVal lpLCData As String) As Boolean
        Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Short
        Const LOCALE_SSHORTDATE As Short = &H1FS
        Public regionalSymbol As String

        Public Sub Get_locale() ' Retrieve the regional setting
            Dim Symbol As String
            Dim iRet1 As Integer
            Dim iRet2 As Integer
            Dim lpLCDataVar As String = String.Empty
            Dim Pos As Short
            Dim Locale As Integer
            Locale = GetUserDefaultLCID()
            iRet1 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, lpLCDataVar, 0)
            Symbol = New String(Chr(0), iRet1)
            iRet2 = GetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol, iRet1)
            Pos = InStr(Symbol, Chr(0))
            If Pos > 0 Then
                Symbol = Left(Symbol, Pos - 1)
                If Symbol <> "yyyy/MM/dd" Then regionalSymbol = Symbol
            End If
        End Sub
        Public Sub Set_locale(Optional ByVal localSymbol As String = "") 'Change the regional setting
            Dim Symbol As String
            Dim iRet As Integer
            Dim Locale As Integer
            Locale = GetUserDefaultLCID() 'Get user Locale ID
            If localSymbol = "" Then
                Symbol = "yyyy/MM/dd" 'New character for the locale
            Else
                Symbol = localSymbol
            End If

            iRet = SetLocaleInfo(Locale, LOCALE_SSHORTDATE, Symbol)
        End Sub


        ' Performs RSA signing of <code>strData</code> using the specified key.
        ' @param strPub     RSA Public key blob
        ' @param strPriv    RSA Private key blob
        ' @param strData    Data to be signed
        ' @return           Signature string.
        ' 05.13.05    - alkan  - Removed the modActiveLock references
        '
        Public Function RSASign(ByVal strPub As String, ByVal strPriv As String, ByVal strdata As String, ByVal siteNameUrl As String) As String
            Dim Key As RSAKey
            Dim sLen As Integer
            Dim strSig As String

            ReDim Key.data(32)
            Select Case siteNameUrl
                Case "localhost"
                    ' create the key from the key blobs
                    rsa_createkey_local(strPub, Len(strPub), strPriv, Len(strPriv), Key)
                    ' sign the data using the created key
                    rsa_sign_local(Key, strdata, Len(strdata), vbNullString, sLen)
                    strSig = New String(Chr(0), sLen)
                    rsa_sign_local(Key, strdata, Len(strdata), strSig, sLen)
                    ' throw away the key
                    rsa_freekey_local(Key)
                Case Else
                    ' create the key from the key blobs
                    rsa_createkey(strPub, Len(strPub), strPriv, Len(strPriv), Key)
                    ' sign the data using the created key
                    rsa_sign(Key, strdata, Len(strdata), vbNullString, sLen)
                    strSig = New String(Chr(0), sLen)
                    rsa_sign(Key, strdata, Len(strdata), strSig, sLen)
                    ' throw away the key
                    rsa_freekey(Key)
            End Select
            RSASign = strSig
        End Function

        ' Verifies an RSA signature.
        ' @param strPub     Public key blob
        ' @param strData    Data to be signed
        ' @param strSig     Private key blob
        ' @return           Zero if verification is successful; Non-zero otherwise.
        '
        Public Function RSAVerify(ByVal strPub As String, ByVal strdata As String, ByVal strSig As String, ByVal siteNameUrl As String) As Integer
            Dim Key As RSAKey = Nothing
            Dim rc As Integer
            Select Case siteNameUrl
                Case "localhost"
                    ' create the key from the public key blob
                    rsa_createkey_local(strPub, Len(strPub), vbNullString, 0, Key)
                    ' validate the key
                    rc = rsa_verifysig_local(Key, strSig, Len(strSig), strdata, Len(strdata))
                    ' de-allocate memory used by the key
                    rsa_freekey_local(Key)
                Case Else
                    ' create the key from the public key blob
                    rsa_createkey(strPub, Len(strPub), vbNullString, 0, Key)
                    ' validate the key
                    rc = rsa_verifysig(Key, strSig, Len(strSig), strdata, Len(strdata))
                    ' de-allocate memory used by the key
                    rsa_freekey(Key)
            End Select
            RSAVerify = rc
        End Function
    End Module
End Namespace
