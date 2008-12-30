Option Strict On
Option Explicit On 
Imports System.Xml
Imports System.Security.Cryptography
Imports System.text

Friend Class XMLGenerator
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
  Private fileXML As String
  Private MyXMLDoc As System.Xml.XmlDocument

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
                'setup xml doc
                Call SetupXmlDoc()
                'save the new created xml doc
                MyXMLDoc.Save(Value)
            Else
                MyXMLDoc = New Xml.XmlDocument
                'load existing xml doc
                MyXMLDoc.Load(Value)
            End If
            fileXML = Value
        End Set
  End Property

  Private Sub SetupXmlDoc()
    Dim Version As XmlProcessingInstruction  '<?xml version="1.0"?>

    'Create a new document object
    MyXMLDoc = New XmlDocument
    MyXMLDoc.preserveWhiteSpace = False
    'MyXMLDoc.async = False
    'tell the parser to automatically load externally defined DTD's or XSD's
    'MyXMLDoc.resolveExternals = True

    Version = MyXMLDoc.createProcessingInstruction("xml", "version=" & Chr(34) & "1.0" & Chr(34))
    'create the root element
    'and append it and the XML version to the document
    MyXMLDoc.appendChild(Version)

    MyXMLDoc.appendChild(MyXMLDoc.createElement("Alugen3"))

  End Sub

  Private Function BuildProductNodes(ByVal RootNode As XmlNode, ByVal ProdInfo As ProductInfo) As XmlNode
    'Build the document structure needed to add a new product

    Dim NameNode As XmlNode     'software name
    Dim VersionNode As XmlNode  'software version
    Dim VCodeNode As XmlNode    'software vcode
    Dim GCodeNode As XmlNode    'software gcode

    Dim productsNode As XmlNode

    'MyXMLDoc.namespaces
    productsNode = MyXMLDoc.CreateElement("Products")
    RootNode.AppendChild(productsNode)

    'Create all the new data nodes
    NameNode = MyXMLDoc.CreateElement("name")
    VersionNode = MyXMLDoc.CreateElement("version")
    VCodeNode = MyXMLDoc.CreateElement("vcode")
    GCodeNode = MyXMLDoc.CreateElement("gcode")

    'fill info
    NameNode.appendChild(MyXMLDoc.CreateTextNode(ProdInfo.Name))
    VersionNode.appendChild(MyXMLDoc.CreateTextNode(ProdInfo.Version))
    VCodeNode.appendChild(MyXMLDoc.CreateTextNode(ProdInfo.VCode))
    GCodeNode.appendChild(MyXMLDoc.CreateTextNode(ProdInfo.GCode))

    'Add the new nodes to the root "Products" node
    'Note that the append order is important
    productsNode.appendChild(NameNode)
    productsNode.appendChild(VersionNode)
    productsNode.appendChild(VCodeNode)
    productsNode.appendChild(GCodeNode)

    'Return the new root "Products" node
    BuildProductNodes = RootNode
  End Function


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
  Private Sub LoadProdInfo(ByRef pNode As XmlNode, ByRef ProdInfo As ProductInfo)
    With pNode
      ProdInfo.Name = pNode.SelectSingleNode("name").InnerText
      ProdInfo.Version = pNode.SelectSingleNode("version").InnerText
      ProdInfo.VCode = pNode.SelectSingleNode("vcode").InnerText
      ProdInfo.GCode = pNode.SelectSingleNode("gcode").InnerText
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
    Dim rootProducts As XmlNode
    Dim newProductNode As XmlNode
    'Dim schemaCache As New XMLSchemaCache40
    'Dim Result As IXMLDOMParseError

    rootProducts = MyXMLDoc.SelectSingleNode("Alugen3")
    newProductNode = BuildProductNodes(rootProducts, ProdInfo)

    'schema validation
    'schemaCache.Add "http://jclement.ca/xml/simple", "c:simple.xsd"
    'Set MyXMLDoc.schemas = schemaCache
    'Set result = MyXMLDoc.Validate
    'If result.errorCode = 0 Then
        '    Set_locale(regionalSymbol)
        '    Err.Raise ("XML Validation Error:" & result.reason)
    'End If
    MyXMLDoc.Save(fileXML)
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
    'Retrieve all product information from XML file.  Return as an array.
    On Error GoTo RetrieveProductsError
        Dim arrProdInfos() As ProductInfo = Nothing
    Dim Count As Integer
    Dim xmlCount As Integer
    Count = 0
    Dim rootProducts As XmlNodeList
    rootProducts = MyXMLDoc.GetElementsByTagName("Products")

    xmlCount = rootProducts.Count

    ' If there are no products in the XML file, exit gracefully
        If xmlCount < 1 Then Return Nothing
    For Count = 0 To xmlCount - 1
      ReDim Preserve arrProdInfos(Count)
      arrProdInfos(Count) = New ProductInfo
      LoadProdInfo(rootProducts(Count), arrProdInfos(Count))
    Next
    IALUGenerator_RetrieveProducts = arrProdInfos
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

    Dim xmlCount As Integer
    Dim Count As Integer
    Count = 0

    Dim rootProducts As XmlNodeList
    rootProducts = MyXMLDoc.GetElementsByTagName("Products")

    xmlCount = rootProducts.Count

    For Count = 0 To xmlCount - 1
      With rootProducts(Count)
        If .SelectSingleNode("name").InnerText = Name And _
              .SelectSingleNode("version").InnerText = Ver Then
          ProdInfo.VCode = .SelectSingleNode("vcode").InnerText
          ProdInfo.GCode = .SelectSingleNode("gcode").InnerText
          Exit For
        End If
      End With
    Next
    If ProdInfo.VCode = "" Or ProdInfo.GCode = "" Then
            'Set_locale(regionalSymbol)
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
    ' Remove the section from XML file
    Dim xmlCount As Integer
    Dim Count As Integer
    Count = 0

    Dim rootProducts As XmlNodeList
    rootProducts = MyXMLDoc.GetElementsByTagName("Products")

    xmlCount = rootProducts.Count

    For Count = 0 To xmlCount - 1
      If rootProducts(Count).SelectSingleNode("name").InnerText = name And _
            rootProducts(Count).SelectSingleNode("version").InnerText = Ver Then
        rootProducts(Count).ParentNode.RemoveChild(rootProducts(Count))
        Exit For
      End If
    Next

    MyXMLDoc.Save(fileXML)
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
            Catch ex As Exception
                'Set_locale(regionalSymbol)
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