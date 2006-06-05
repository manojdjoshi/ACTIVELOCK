Option Strict Off
Option Explicit On 

Imports System
Imports System.Data
Imports DotNetNuke
Imports DotNetNuke.Framework
Imports FriendSoftware.DNN.Modules.Alugen3NET.Data
Imports ActiveLock3NET

Namespace FriendSoftware.DNN.Modules.Alugen3NET.Business
  Public Class aluGenerator
    Implements ActiveLock3NET._IALUGenerator


    Private MyActiveLock As _IActiveLock
    Private _ProductInfo As ProductInfo

    ' Returns IActiveLock interface
    '
    Private ReadOnly Property ActiveLockInterface() As _IActiveLock
      Get
        ActiveLockInterface = MyActiveLock
      End Get
    End Property


    Public Sub DeleteProduct(ByVal name As String, ByVal Ver As String) Implements ActiveLock3NET._IALUGenerator.DeleteProduct

    End Sub

    Public Function GenKey(ByRef Lic As ActiveLock3NET.ProductLicense, ByVal InstCode As String, ByVal ProductInfo As ProductInfo, Optional ByVal RegisteredLevel As String = "0") As String
      _ProductInfo = ProductInfo
      If Not IsNothing(RegisteredLevel) Then
        Return GenKey(Lic, InstCode, RegisteredLevel)
      Else
        Return GenKey(Lic, InstCode)
      End If

    End Function

    Public Function GenKey(ByRef Lic As ActiveLock3NET.ProductLicense, ByVal InstCode As String, Optional ByVal RegisteredLevel As String = "0") As String Implements ActiveLock3NET._IALUGenerator.GenKey
      ' Take request code and decrypt it.

      Dim strReq As String
      ' 05.13.05 - ialkan Modified to merge DLLs into one
      strReq = Base64_Decode(InstCode)

      ' strReq now contains the {LockCode + vbLf + User} string
      Dim strLock, strUser As String
      GetLockAndUserFromInstallCode(strReq, strLock, strUser)

      Lic.Licensee = strUser
      '  registration date
      Dim strRegDate As String
      ' registered level
      Lic.RegisteredLevel = RegisteredLevel
      strRegDate = Lic.RegisteredDate

      Dim strEncrypted As String
      ' @todo Rethink this bit about encrypting the dates.
      ' We need to keep in mind that the app does not have access to the private key, so and any decryption that requires private key
      ' would not be possible.
      ' Perhaps instead of encrypting, we could do MD5 hash of (regdate+lockcode)?
      'ActiveLockEventSink_ValidateValue strRegDate, strEncrypted
      ' hash it
      'strEncrypted = ActiveLock3.MD5Hash(strEncrypted)
      strEncrypted = strRegDate

      ' get software codes
      'Dim ProdInfo As ProductInfo
      'ProdInfo = IALUGenerator_RetrieveProduct(Lic.ProductName, Lic.ProductVer)
      Lic.ProductKey = _ProductInfo.VCode

      '@todo Check for "ProdInfo Is Nothing" and handle appropriately

      Dim strLic As String
      ' Encrypt Product license using the MAC
      strLic = Lic.ToString_Renamed() & vbLf & strLock
      System.Diagnostics.Debug.WriteLine("strLic: " & vbCrLf & strLic)

      ' sign it
      Dim strSig As String
      strSig = New String(Chr(0), 1024)
      ' 05.13.05 - ialkan Modified to merge DLLs into one. Moved RSASign into a module
      strSig = RSASign(_ProductInfo.VCode, _ProductInfo.GCode, strLic)

      ' Create liberation key.  This will be a base-64 encoded string of the whole license.
      Dim strLicKey As String
      ' 05.13.05 - ialkan Modified to merge DLLs into one
      strLicKey = Base64_Encode(strSig)
      ' update Lic with license key
      Lic.LicenseKey = strLicKey
      ' Print some info for debugging purposes
      System.Diagnostics.Debug.WriteLine("VCode: " & _ProductInfo.VCode)
      System.Diagnostics.Debug.WriteLine("Lic: " & strLic)
      System.Diagnostics.Debug.WriteLine("Lic hash: " & modMD5.Hash(strLic))
      System.Diagnostics.Debug.WriteLine("LicKey: " & strLicKey)
      System.Diagnostics.Debug.WriteLine("Sig: " & strSig)
      System.Diagnostics.Debug.WriteLine("Verify: " & RSAVerify(_ProductInfo.VCode, strLic, Base64_Decode(strLicKey)))
      System.Diagnostics.Debug.WriteLine("====================================================")

      ' Serialize it into a formatted string
      Dim strLibKey As String
      Lic.Save(strLibKey)
      GenKey = strLibKey
    End Function

    Public Function RetrieveProduct(ByVal name As String, ByVal Ver As String) As ActiveLock3NET.ProductInfo Implements ActiveLock3NET._IALUGenerator.RetrieveProduct

    End Function

    Public Function RetrieveProducts() As ActiveLock3NET.ProductInfo() Implements ActiveLock3NET._IALUGenerator.RetrieveProducts

    End Function

    Public Sub SaveProduct(ByRef ProdInfo As ActiveLock3NET.ProductInfo) Implements ActiveLock3NET._IALUGenerator.SaveProduct

    End Sub

    Public WriteOnly Property StoragePath() As String Implements ActiveLock3NET._IALUGenerator.StoragePath
      Set(ByVal Value As String)

      End Set
    End Property

    ' Retrieves lock string and user info from the request string
    ' @param strReq     strLock combined with user name.
    ' @param strLock    Generate Request code to Lock.
    ' @param strUser    User name.
    '
    Private Sub GetLockAndUserFromInstallCode(ByVal strReq As String, ByRef strLock As String, ByRef strUser As String)
      Dim Index, i As Short
      Index = 0 : i = 1
      ' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
      Do While i > 0
        i = InStr(Index + 1, strReq, vbLf)
        If i > 0 Then Index = i
      Loop

      If Index <= 0 Then Exit Sub
      ' lockcode is from beginning to Index-1
      strLock = Left(strReq, Index - 1)
      ' user name starts from Index+1 to the end
      strUser = Mid(strReq, Index + 1)
      strUser = modActiveLock.TrimNulls(strUser)
    End Sub
  End Class
End Namespace
