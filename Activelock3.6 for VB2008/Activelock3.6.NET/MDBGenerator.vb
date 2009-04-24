Option Strict On
Option Explicit On
Imports System.Data
Imports System.Data.OleDb
Imports System.Security.Cryptography
Imports System.text

#Region "Copyright"
' This project is available from SVN on SourceForge.net under the main project, Activelock !
'
' ProjectPage: http://sourceforge.net/projects/activelock
' WebSite: http://www.activeLockSoftware.com
' DeveloperForums: http://forums.activelocksoftware.com
' ProjectManager: Ismail Alkan - http://activelocksoftware.com/simplemachinesforum/index.php?action=profile;u=1
' ProjectLicense: BSD Open License - http://www.opensource.org/licenses/bsd-license.php
' ProjectPurpose: Copy Protection, Software Locking, Anti Piracy
'
' //////////////////////////////////////////////////////////////////////////////////////////
' *   ActiveLock
' *   Copyright 1998-2002 Nelson Ferraz
' *   Copyright 2003-2009 The ActiveLock Software Group (ASG)
' *   All material is the property of the contributing authors.
' *
' *   Redistribution and use in source and binary forms, with or without
' *   modification, are permitted provided that the following conditions are
' *   met:
' *
' *     [o] Redistributions of source code must retain the above copyright
' *         notice, this list of conditions and the following disclaimer.
' *
' *     [o] Redistributions in binary form must reproduce the above
' *         copyright notice, this list of conditions and the following
' *         disclaimer in the documentation and/or other materials provided
' *         with the distribution.
' *
' *   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' *   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' *   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
' *   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
' *   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
' *   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
' *   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
' *   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' *   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' *   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' *   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' *
#End Region

Friend Class MDBGenerator
  Implements _IALUGenerator

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
  Private oConnect As OleDbConnection
  Private oRST As OleDbDataReader
  Private myCmd As OleDbCommand

  Private fileMDB As String

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
      If Not FileExists(Value) Then
      Else
        fileMDB = Value
      End If
    End Set
  End Property
  'Class_Initialize was upgraded to Class_Initialize_Renamed
  Private Sub Class_Initialize_Renamed()
    ' 05.13.05 - ialkan Modified to merge DLLs into one
    ' Initialize AL
    MyActiveLock = New IActiveLock
    oConnect = New OleDbConnection
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
  Private Sub LoadProdInfo(ByRef pDataReader As OleDbDataReader, ByRef ProdInfo As ProductInfo)
    ProdInfo.Name = CType(pDataReader("name"), String)
    ProdInfo.Version = CType(pDataReader("version"), String)
    ProdInfo.VCode = CType(pDataReader("vcode"), String)
    ProdInfo.GCode = CType(pDataReader("gcode"), String)
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
    'save product in database

    Dim strSQL As String
    strSQL = "INSERT INTO products (name, version, vcode, gcode) VALUES ('" & ProdInfo.Name & "', '" & ProdInfo.Version & "', '" & ProdInfo.VCode & "', '" & ProdInfo.GCode & "')"

    Try
      'open connection to MDB file
      oConnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source= " & fileMDB & ";"
      oConnect.Open()

      myCmd = New OleDbCommand
      myCmd.Connection = oConnect
      myCmd.CommandTimeout = 60
      myCmd.CommandText = strSQL
      myCmd.CommandType = CommandType.Text
      myCmd.ExecuteNonQuery()
    Finally
      oConnect.Close()
    End Try

  End Sub
  '===============================================================================
  ' Name: IALUGenerator_RetrieveProducts
  ' Input: None
  ' Output:
  '   ProductInfo - Product info object
  ' Purpose: Retrieves all product information from MDB.
  ' Remarks: Returns as an array.
  '===============================================================================
  Private Function IALUGenerator_RetrieveProducts() As ProductInfo() Implements _IALUGenerator.RetrieveProducts
    'Retrieve all product information from MDB file.  Return as an array.
        Dim arrProdInfos() As ProductInfo = Nothing
    Dim strSQL As String = "SELECT * FROM products"

    Try
      'open connection to MDB file
      oConnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source= " & fileMDB & ";"
      oConnect.Open()

      myCmd = New OleDbCommand
      myCmd.CommandText = strSQL
      myCmd.CommandType = CommandType.Text
      myCmd.CommandTimeout = 60
      myCmd.Connection = oConnect
      oRST = myCmd.ExecuteReader()

      Dim Count As Integer
      Count = 0
      Do While oRST.Read
        ReDim Preserve arrProdInfos(Count)
        arrProdInfos(Count) = New ProductInfo
        'load product info
        LoadProdInfo(oRST, arrProdInfos(Count))
        Count += 1
      Loop
    Finally
      'close dataset
      If Not oRST.IsClosed Then
        oRST.Close()
      End If
      'Close connection
      If Not oConnect.State = ConnectionState.Closed Then
        oConnect.Close()
      End If
    End Try
    IALUGenerator_RetrieveProducts = arrProdInfos
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
    Dim strSQL As String = "SELECT * FROM products WHERE name='" & Name & "' AND version='" & Ver & "'"

    Try
      'open connection to MDB file
      oConnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                "Data Source= " & fileMDB & ";"
      oConnect.Open()

      myCmd = New OleDbCommand
      myCmd.CommandText = strSQL
      myCmd.CommandType = CommandType.Text
      myCmd.CommandTimeout = 60
      myCmd.Connection = oConnect
      oRST = myCmd.ExecuteReader()

      If oRST.HasRows Then
        oRST.Read()
        ProdInfo.VCode = CType(oRST("vcode"), String)
        ProdInfo.GCode = CType(oRST("gcode"), String)
      End If

    Finally
      oRST.Close()
      oConnect.Close()
    End Try

    If ProdInfo.VCode = "" Or ProdInfo.GCode = "" Then
            '* Set_locale(regionalSymbol)
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
        ' Remove the section from MDB file
        Dim strSQL As String
        strSQL = "DELETE FROM products WHERE name='" & name & "' AND version='" & Ver & "'"

        Try
            'open connection to MDB file
            oConnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                  "Data Source= " & fileMDB & ";"
            oConnect.Open()

            myCmd = New OleDbCommand
            myCmd.Connection = oConnect
            myCmd.CommandTimeout = 60
            myCmd.CommandText = strSQL
            myCmd.CommandType = CommandType.Text
            myCmd.ExecuteNonQuery()
        Finally
            oConnect.Close()
        End Try
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
    Private Function IALUGenerator_GenKey(ByRef Lic As ActiveLock3_6NET.ProductLicense, ByVal InstCode As String, Optional ByVal RegisteredLevel As String = "0") As String Implements _IALUGenerator.GenKey
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
        '* strRegDate = Lic.RegisteredDate
        strRegDate = DateToDblString(Lic.RegisteredDate) '*

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
        'System.Diagnostics.Debug.WriteLine("strLic: " & vbCrLf & strLic)

        If strLeft(ProdInfo.VCode, 3) <> "RSA" Then
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
        Else

            Try
                Dim rsaCSP As New System.Security.Cryptography.RSACryptoServiceProvider
                Dim strPublicBlob, strPrivateBlob As String

                strPublicBlob = ProdInfo.VCode
                strPrivateBlob = ProdInfo.GCode

                If strLeft(ProdInfo.GCode, 6) = "RSA512" Then
                    strPrivateBlob = strRight(ProdInfo.GCode, Len(ProdInfo.GCode) - 6)
                Else
                    strPrivateBlob = strRight(ProdInfo.GCode, Len(ProdInfo.GCode) - 7)
                End If
                ' import private key params into instance of RSACryptoServiceProvider
                rsaCSP.FromXmlString(strPrivateBlob)
                Dim rsaPrivateParams As RSAParameters 'stores private key
                rsaPrivateParams = rsaCSP.ExportParameters(True)
                rsaCSP.ImportParameters(rsaPrivateParams)

                Dim userData As Byte() = Encoding.UTF8.GetBytes(strLic)
                Dim asf As AsymmetricSignatureFormatter = New RSAPKCS1SignatureFormatter(rsaCSP)
                Dim algorithm As HashAlgorithm = New SHA1Managed
                asf.SetHashAlgorithm(algorithm.ToString)
                Dim myhashedData() As Byte ' a byte array to store hash value
                Dim myhashedDataString As String
                myhashedData = algorithm.ComputeHash(userData)
                myhashedDataString = BitConverter.ToString(myhashedData).Replace("-", String.Empty)
                Dim mysignature As Byte() ' holds signatures
                mysignature = asf.CreateSignature(algorithm)
                Dim mySignatureBlock As String
                mySignatureBlock = Convert.ToBase64String(mysignature)
                Lic.LicenseKey = mySignatureBlock
            Catch ex As Exception
                '* Set_locale(regionalSymbol)
                Err.Raise(AlugenGlobals.alugenErrCodeConstants.alugenProdInvalid, ACTIVELOCKSTRING, ex.Message)
            End Try

        End If

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