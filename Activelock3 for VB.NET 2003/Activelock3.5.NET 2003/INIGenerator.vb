Option Strict On
Option Explicit On

Friend Class INIGenerator
  Implements _IALUGenerator
  '*   ActiveLock
  '*   Copyright 1998-2002 Nelson Ferraz
  '*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
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
  ' Name: Generator
  ' Purpose: This is a concrete implementation of the IALUGenerator interface.
  ' Functions:
  ' Properties:
  ' Methods:
  ' Started: 08.15.2003
  ' Modified: 03.24.2006
  '===============================================================================
  ' @author activelock-admins
  ' @version 3.3.0
  ' @date 03.24.2006

  Private MyActiveLock As _IActiveLock
  Private MyIniFile As ActiveLock3_5NET.INIFile = New ActiveLock3_5NET.INIFile
  '===============================================================================
  ' Name: Property Get ActiveLockInterface
  ' Input: None
  ' Output: None
  ' Purpose: Returns IActiveLock interface
  ' Remarks: None
  '===============================================================================
  Private ReadOnly Property ActiveLockInterface() As _IActiveLock
    Get
      ActiveLockInterface = MyActiveLock
    End Get
  End Property
  '===============================================================================
  ' Name: Property Let IALUGenerator_StoragePath
  ' Input:
  '   ByVal RHS As String -
  ' Output: None
  ' Purpose: None
  ' Remarks: None
  '===============================================================================
  Private WriteOnly Property IALUGenerator_StoragePath() As String Implements _IALUGenerator.StoragePath
    Set(ByVal Value As String)
      Dim hFile As Integer
      If Not FileExists(Value) Then
        ' Create an empty licenses.ini file if it doesn't exists
        hFile = FreeFile()
        FileOpen(hFile, Value, OpenMode.Output)
        FileClose(hFile)
      End If
      MyIniFile.File = Value
    End Set
  End Property
  'Class_Initialize was upgraded to Class_Initialize_Renamed
  Private Sub Class_Initialize_Renamed()
    ' 05.13.05 - ialkan Modified to merge DLLs into one
    ' Initialize AL
    MyActiveLock = New IActiveLock
  End Sub
  Public Sub New()
    MyBase.New()
    Class_Initialize_Renamed()
  End Sub
  '===============================================================================
  ' Name: Sub LoadProdInfo
  ' Input:
  '   ByRef Section As String - Section Name that contains ProdName and ProdVer in order to be unique
  '   ByRef ProdInfo As ProductInfo - Object containing product information to be saved.
  ' Output: None
  ' Purpose: Loads Product Info from the specified INI section.
  ' Remarks: None
  '===============================================================================
  Private Sub LoadProdInfo(ByRef Section As String, ByRef ProdInfo As ProductInfo)
    With MyIniFile
      .Section = Section
      ProdInfo.Name = CStr(.Values("Name"))
      ProdInfo.Version = CStr(.Values("Version"))
      ProdInfo.VCode = CStr(.Values("VCode"))
      ProdInfo.GCode = CStr(.Values("GCode"))
    End With
  End Sub
  '===============================================================================
  ' Name: Sub IALUGenerator_SaveProduct
  ' Input:
  '   ByRef ProdInfo As ProductInfo - Object containing product information to be saved.
  ' Output: None
  ' Purpose: Saves product details in the store file
  ' Remarks: IALUGenerator Interface implementation
  '===============================================================================
  Private Sub IALUGenerator_SaveProduct(ByRef ProdInfo As ProductInfo) Implements _IALUGenerator.SaveProduct
    With MyIniFile
      ' Section name has to contain ProdName and ProdVer in order to be unique
      .Section = ProdInfo.Name & " " & ProdInfo.Version
      .Values("Name") = ProdInfo.Name
      .Values("Version") = ProdInfo.Version
      .Values("VCode") = ProdInfo.VCode '@todo Encrypt code1 and code2, possibly using something as simple as modCrypto.bas
      .Values("GCode") = ProdInfo.GCode
    End With
  End Sub
  '===============================================================================
  ' Name: IALUGenerator_RetrieveProducts
  ' Input: None
  ' Output:
  '   ProductInfo - Product info object
  ' Purpose: Retrieves all product information from INI.
  ' Remarks: Returns as an array.
  '===============================================================================
  Private Function IALUGenerator_RetrieveProducts() As ProductInfo() Implements _IALUGenerator.RetrieveProducts
    ' Retrieve all product information from INI.  Return as an array.
    On Error GoTo RetrieveProductsError
    IALUGenerator_RetrieveProducts = Nothing
    Dim arrProdInfos() As ProductInfo
    Dim Count As Integer
    Dim iniCount As Short
    Dim arrSections() As String = Nothing
    Count = 0
    iniCount = MyIniFile.EnumSections(arrSections)

    ' If there are no products in the INI file, exit gracefully
    If iniCount < 1 Then Exit Function
    For Count = 0 To iniCount - 1
      ReDim Preserve arrProdInfos(Count)
      arrProdInfos(Count) = New ProductInfo
      LoadProdInfo(arrSections(Count), arrProdInfos(Count))
    Next
    IALUGenerator_RetrieveProducts = CType(VB6.CopyArray(arrProdInfos), ProductInfo())
    Exit Function

RetrieveProductsError:
  End Function
  '===============================================================================
  ' Name: Function IALUGenerator_RetrieveProduct
  ' Input:
  '   ByVal name As String - Product name
  '   ByVal Ver As String - Product version
  ' Output: None
  ' Purpose: Retrieves product VCode and GCode from the store file
  ' Remarks: todo Error Handling - Need to return Nothing if store file doesn't contain the product
  '===============================================================================
  Private Function IALUGenerator_RetrieveProduct(ByVal Name As String, ByVal Ver As String) As ProductInfo Implements _IALUGenerator.RetrieveProduct
    '@todo Error Handling - Need to return Nothing if store file doesn't contain the product
    Dim ProdInfo As New ProductInfo
    ProdInfo.Name = Name
    ProdInfo.Version = Ver
    With MyIniFile
      .Section = Name & " " & Ver
      ProdInfo.Version = CStr(.Values("Version"))
      ProdInfo.VCode = CStr(.Values("VCode")) '@todo Decrypt code1 and code2
      ProdInfo.GCode = CStr(.Values("GCode"))
    End With
    If ProdInfo.VCode = "" Or ProdInfo.GCode = "" Then
      'ACTIVELOCKSTRING could be replaced by System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly.Location).ProductName
      Err.Raise(AlugenGlobals.alugenErrCodeConstants.alugenProdInvalid, ACTIVELOCKSTRING, "Product code set is invalid.")
    End If
    IALUGenerator_RetrieveProduct = ProdInfo
  End Function
  '===============================================================================
  ' Name: Sub IALUGenerator_DeleteProduct
  ' Input:
  '   ByVal name As String - Product name
  '   ByVal Ver As String - Product version
  ' Output: None
  ' Purpose: Removes the license keys section from a INI file, i.e. deletes product details in the license database
  ' Remarks: Removes a section from the INI file
  '===============================================================================
  Private Sub IALUGenerator_DeleteProduct(ByVal name As String, ByVal Ver As String) Implements _IALUGenerator.DeleteProduct
    ' Remove the section from INI file
    Call MyIniFile.DeleteSection(name & " " & Ver)
  End Sub
  '===============================================================================
  ' Name: Function IALUGenerator_GenKey
  ' Input:
  '   ByRef Lic As ActiveLock3.ProductLicense - Product license
  '   ByVal InstCode As String - Installation Code sent by the user
  '   ByVal RegisteredLevel As String - Registration Level for the license. Default is "0"
  ' Output:
  '   String - Liberation key for the license
  ' Purpose: Given the Installation Code, generates an Activelock license liberation key.
  ' Remarks: None
  '===============================================================================
  Private Function IALUGenerator_GenKey(ByRef Lic As ActiveLock3_5NET.ProductLicense, ByVal InstCode As String, Optional ByVal RegisteredLevel As String = "0") As String Implements _IALUGenerator.GenKey
    ' Take request code and decrypt it.

    Dim strReq As String
    ' 05.13.05 - ialkan Modified to merge DLLs into one
    strReq = Base64_Decode(InstCode)

    ' strReq now contains the {LockCode + vbLf + User} string
    Dim strLock As String = String.Empty
    Dim strUser As String = String.Empty
    GetLockAndUserFromInstallCode(strReq, strLock, strUser)

    Lic.Licensee = strUser
    ' registration date
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
    Dim ProdInfo As ProductInfo
    ProdInfo = IALUGenerator_RetrieveProduct(Lic.ProductName, Lic.ProductVer)
    Lic.ProductKey = ProdInfo.VCode

    '@todo Check for "ProdInfo Is Nothing" and handle appropriately

    Dim strLic As String
    strLic = Lic.ToString_Renamed() & vbLf & strLock
    System.Diagnostics.Debug.WriteLine("strLic: " & vbCrLf & strLic)

    ' sign it
    Dim strSig As String
    strSig = New String(Chr(0), 1024)
    ' 05.13.05 - ialkan Modified to merge DLLs into one. Moved RSASign into a module
    strSig = RSASign(ProdInfo.VCode, ProdInfo.GCode, strLic)

    ' Create liberation key.  This will be a base-64 encoded string of the whole license.
    Dim strLicKey As String
    ' 05.13.05 - ialkan Modified to merge DLLs into one
    strLicKey = Base64_Encode(strSig)
    ' update Lic with license key
    Lic.LicenseKey = strLicKey
    ' Print some info for debugging purposes
    System.Diagnostics.Debug.WriteLine("VCode: " & ProdInfo.VCode)
    System.Diagnostics.Debug.WriteLine("Lic: " & strLic)
    System.Diagnostics.Debug.WriteLine("Lic hash: " & modMD5.Hash(strLic))
    System.Diagnostics.Debug.WriteLine("LicKey: " & strLicKey)
    System.Diagnostics.Debug.WriteLine("Sig: " & strSig)
    System.Diagnostics.Debug.WriteLine("Verify: " & RSAVerify(ProdInfo.VCode, strLic, Base64_Decode(strLicKey)))
    System.Diagnostics.Debug.WriteLine("====================================================")

    ' Serialize it into a formatted string
    Dim strLibKey As String = String.Empty
    Lic.Save(strLibKey)
    IALUGenerator_GenKey = strLibKey
  End Function
  '===============================================================================
  ' Name: Sub GetLockAndUserFromInstallCode
  ' Input:
  '   ByVal strReq As String - strLock combined with user name.
  '   ByRef strLock As String - Generated request code to Lock
  '   ByRef strUser As String - User name
  ' Output: None
  ' Purpose: Retrieves lock string and user info from the request string
  ' Remarks: None
  '===============================================================================
  Private Sub GetLockAndUserFromInstallCode(ByVal strReq As String, ByRef strLock As String, ByRef strUser As String)
    Dim Index, i As Integer
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