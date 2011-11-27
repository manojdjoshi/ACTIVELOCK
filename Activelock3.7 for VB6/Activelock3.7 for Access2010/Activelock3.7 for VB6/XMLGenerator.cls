VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XMLGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Implements IALUGenerator

Private MyActiveLock As ActiveLock3.IActiveLock
Private MyXMLDoc As MSXML2.DOMDocument40
Private fileXML As String

'===============================================================================
' Name: Property Get ActiveLockInterface
' Input: None
' Output: None
' Purpose: Returns IActiveLock interface
' Remarks: None
'===============================================================================
Private Property Get ActiveLockInterface() As IActiveLock
  Set ActiveLockInterface = MyActiveLock
End Property


Private Sub Class_Initialize()
  ' Initialize AL
  Set MyActiveLock = New IActiveLock
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
Private Sub LoadProdInfo(pNode As MSXML2.IXMLDOMNode, ProdInfo As ProductInfo)
    With pNode
        ProdInfo.name = pNode.selectSingleNode("name").Text
        ProdInfo.Version = pNode.selectSingleNode("version").Text
        ProdInfo.VCode = pNode.selectSingleNode("vcode").Text
        ProdInfo.GCode = pNode.selectSingleNode("gcode").Text
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
Private Sub IALUGenerator_SaveProduct(ProdInfo As ProductInfo)
    Dim rootProducts As MSXML2.IXMLDOMNode
    Dim newProductNode As MSXML2.IXMLDOMNode
    Dim schemaCache As New XMLSchemaCache40
    Dim Result As IXMLDOMParseError
    
    Set rootProducts = MyXMLDoc.selectSingleNode("Alugen3")
    Set newProductNode = BuildProductNodes(rootProducts, ProdInfo)
    
    'schema validation
    'schemaCache.Add "http://jclement.ca/xml/simple", "c:simple.xsd"
    'Set MyXMLDoc.schemas = schemaCache
    'Set result = MyXMLDoc.Validate
    'If result.errorCode = 0 Then
    '    Set_locale regionalSymbol
    '    Err.Raise ("XML Validation Error:" & result.reason)
    'End If

    MyXMLDoc.Save (fileXML)
End Sub

'===============================================================================
' Name: IALUGenerator_RetrieveProducts
' Input: None
' Output:
'   ProductInfo - Product info object
' Purpose: Retrieves all product information from INI.
' Remarks: Returns as an array.
'===============================================================================
Private Function IALUGenerator_RetrieveProducts() As ProductInfo()
    'Retrieve all product information from XML file.  Return as an array.
    On Error GoTo RetrieveProductsError
    Dim arrProdInfos() As ProductInfo
    Dim Count As Integer
    Dim xmlCount As Integer
    Count = 0
    Dim rootProducts As MSXML2.IXMLDOMNodeList
    Set rootProducts = MyXMLDoc.getElementsByTagName("Products")
    
    xmlCount = rootProducts.Length
    
    ' If there are no products in the XML file, exit gracefully
    If xmlCount < 1 Then Exit Function
    For Count = 0 To xmlCount - 1
        ReDim Preserve arrProdInfos(Count)
        Set arrProdInfos(Count) = New ProductInfo
        LoadProdInfo rootProducts(Count), arrProdInfos(Count)
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
Private Function IALUGenerator_RetrieveProduct(ByVal name As String, ByVal Ver As String) As ProductInfo
    '@todo Error Handling - Need to return Nothing if store file doesn't contain the product
    Dim ProdInfo As New ProductInfo
    ProdInfo.name = name
    ProdInfo.Version = Ver
    
    Dim xmlCount As Integer
    Dim Count As Integer
    Count = 0
    
    Dim rootProducts As MSXML2.IXMLDOMNodeList
    Set rootProducts = MyXMLDoc.getElementsByTagName("Products")
    
    xmlCount = rootProducts.Length

    For Count = 0 To xmlCount - 1
      With rootProducts(Count)
          If .selectSingleNode("name").Text = name And _
                .selectSingleNode("version").Text = Ver Then
              ProdInfo.VCode = .selectSingleNode("vcode").Text
              ProdInfo.GCode = .selectSingleNode("gcode").Text
              Exit For
          End If
      End With
    Next
    If ProdInfo.VCode = "" Or ProdInfo.GCode = "" Then
        Set_locale regionalSymbol
        Err.Raise ALUGENErrCodeConstants.alugenProdInvalid, ACTIVELOCKSTRING, "Product code set is invalid."
    End If
    Set IALUGenerator_RetrieveProduct = ProdInfo
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
Private Sub IALUGenerator_DeleteProduct(ByVal name As String, ByVal Ver As String)
    ' Remove the section from XML file
    Dim xmlCount As Integer
    Dim Count As Integer
    Count = 0
    
    Dim rootProducts As MSXML2.IXMLDOMNodeList
    Set rootProducts = MyXMLDoc.getElementsByTagName("Products")

    xmlCount = rootProducts.Length

    For Count = 0 To xmlCount - 1
      If rootProducts(Count).selectSingleNode("name").Text = name And _
            rootProducts(Count).selectSingleNode("version").Text = Ver Then
          rootProducts(Count).parentNode.removeChild rootProducts(Count)
          Exit For
      End If
    Next
    
    MyXMLDoc.Save (fileXML)
End Sub

'===============================================================================
' Name: Property Let IALUGenerator_StoragePath
' Input:
'   ByVal RHS As String -
' Output: None
' Purpose: None
' Remarks: None
'===============================================================================
Private Property Let IALUGenerator_StoragePath(ByVal RHS As String)
    If Not FileExists(RHS) Then
      'setup xml doc
      Call SetupXmlDoc
      'save the new created xml doc
      MyXMLDoc.Save (RHS)
    Else
      Set MyXMLDoc = New MSXML2.DOMDocument40
      'load existing xml doc
      MyXMLDoc.Load (RHS)
    End If
    fileXML = RHS
End Property

Private Sub SetupXmlDoc()
    Dim Version As MSXML2.IXMLDOMProcessingInstruction  '<?xml version="1.0"?>
     
    'Create a new document object
    Set MyXMLDoc = New MSXML2.DOMDocument40
    MyXMLDoc.preserveWhiteSpace = False
    MyXMLDoc.async = False
    'tell the parser to automatically load externally defined DTD's or XSD's
    MyXMLDoc.resolveExternals = True
  
    Set Version = MyXMLDoc.createProcessingInstruction("xml", "version=" & Chr(34) & "1.0" & Chr(34))
    'create the root element
    'and append it and the XML version to the document
    MyXMLDoc.appendChild Version
    
    MyXMLDoc.appendChild MyXMLDoc.createElement("Alugen3")
  
End Sub

Private Function BuildProductNodes(RootNode As IXMLDOMNode, ProdInfo As ProductInfo) As IXMLDOMNode
    'Build the document structure needed to add a new product
    
    Dim NameNode As MSXML2.IXMLDOMNode     'software name
    Dim VersionNode As MSXML2.IXMLDOMNode    'software version
    Dim VCodeNode As MSXML2.IXMLDOMNode     'software vcode
    Dim GCodeNode As MSXML2.IXMLDOMNode     'software gcode
  
    Dim productsNode As MSXML2.IXMLDOMNode
  
    'MyXMLDoc.namespaces
    Set productsNode = MyXMLDoc.createElement("Products")
    RootNode.appendChild productsNode
    
    'Create all the new data nodes
    Set NameNode = MyXMLDoc.createElement("name")
    Set VersionNode = MyXMLDoc.createElement("version")
    Set VCodeNode = MyXMLDoc.createElement("vcode")
    Set GCodeNode = MyXMLDoc.createElement("gcode")
  
    'fill info
    NameNode.appendChild MyXMLDoc.createTextNode(ProdInfo.name)
    VersionNode.appendChild MyXMLDoc.createTextNode(ProdInfo.Version)
    VCodeNode.appendChild MyXMLDoc.createTextNode(ProdInfo.VCode)
    GCodeNode.appendChild MyXMLDoc.createTextNode(ProdInfo.GCode)
  
    'Add the new nodes to the root "Products" node
    'Note that the append order is important
    productsNode.appendChild NameNode
    productsNode.appendChild VersionNode
    productsNode.appendChild VCodeNode
    productsNode.appendChild GCodeNode
  
    'Return the new root "Products" node
    Set BuildProductNodes = RootNode
End Function


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
Private Function IALUGenerator_GenKey(Lic As ActiveLock3.ProductLicense, ByVal InstCode As String, Optional ByVal RegisteredLevel As String = "0") As String
Dim strReq As String
Dim strLock As String, strUser As String
Dim strRegDate As String
Dim ProdInfo As ProductInfo
Dim strLic As String
Dim strSig As String
Dim strLicKey As String
Dim strLibKey As String

' Take request code and decrypt it.

' 05.13.05 - ialkan Modified to merge DLLs into one
strReq = Base64_Decode(InstCode)
' strReq now contains the {LockCode + vbLf + User} string
GetLockAndUserFromInstallCode strReq, strLock, strUser
Lic.Licensee = strUser
' registered level
Lic.RegisteredLevel = RegisteredLevel
' registration date
'* strRegDate = Lic.RegisteredDate
strRegDate = DateToDblString(Lic.RegisteredDate)

' @todo Rethink this bit about encrypting the dates.
' We need to keep in mind that the app does not have access to the private key, so and any decryption that requires private key
' would not be possible.
' Perhaps instead of encrypting, we could do MD5 hash of (regdate+lockcode)?
'ActiveLockEventSink_ValidateValue strRegDate, strEncrypted
' hash it
'strEncrypted = ActiveLock3.MD5Hash(strEncrypted)
'strEncrypted = strRegDate

' get software codes
Set ProdInfo = IALUGenerator_RetrieveProduct(Lic.ProductName, Lic.ProductVer)
Lic.ProductKey = ProdInfo.VCode

'@todo Check for "ProdInfo Is Nothing" and handle appropriately

strLic = Lic.ToString() & vbLf & strLock
Debug.Print "strLic: " & vbCrLf & strLic

' ALCrypto Removal
'If Left(ProdInfo.VCode, 3) <> "RSA" Then
'    ' sign it
'    strSig = String(1024, 0)
'    ' 05.13.05 - ialkan Modified to merge DLLs into one. Moved RSASign into a module
'    strSig = RSASign(ProdInfo.VCode, ProdInfo.GCode, strLic)
'
'    ' Create liberation key.  This will be a base-64 encoded string of the whole license.
'    ' 05.13.05 - ialkan Modified to merge DLLs into one
'    strLicKey = Base64_Encode(strSig)
'    ' update Lic with license key
'    Lic.LicenseKey = strLicKey
'
'    ' Print some info for debugging purposes
'    Debug.Print "VCode: " & ProdInfo.VCode
'    Debug.Print "Lic: " & strLic
'    Debug.Print "Lic hash: " & modMD5.Hash(strLic)
'    Debug.Print "LicKey: " & strLicKey
'    Debug.Print "Sig: " & strSig
'    Debug.Print "Verify: " & RSAVerify(ProdInfo.VCode, strLic, Base64_Decode(strLicKey))
'    Debug.Print "===================================================="
'Else
    
    Dim MyGlobals As Globals
    Dim ok As String, strPrivateBlob As String, modulus As Long
    Set MyGlobals = New Globals
    'ok = MyGlobals.ContainerChange(Lic.ProductName & Lic.ProductVer)
    strPrivateBlob = ProdInfo.GCode
    ' Destroy Container
    ok = MyGlobals.CryptoAPIAction(2, Lic.ProductName & Lic.ProductVer, "", "", "", "", 0)
    
    ' Sign a string
    If Left(ProdInfo.GCode, 6) = "RSA512" Then
        strPrivateBlob = Right(ProdInfo.GCode, Len(ProdInfo.GCode) - 6)
        modulus = CInt(Mid(ProdInfo.VCode, 4, 3))
    Else
        strPrivateBlob = Right(ProdInfo.GCode, Len(ProdInfo.GCode) - 7)
        modulus = CInt(Mid(ProdInfo.VCode, 4, 4))
    End If
    ok = MyGlobals.CryptoAPIAction(4, Lic.ProductName & Lic.ProductVer, strLic, strSig, "", strPrivateBlob, modulus)
    ' Create liberation key.  This will be a base-64 encoded string of the whole license.
    ' 05.13.05 - ialkan Modified to merge DLLs into one
    strLicKey = Base64_Encode(strSig)
    ' update Lic with license key
    Lic.LicenseKey = strLicKey

' ALCrypto Removal
'End If

' Serialize it into a formatted string
Lic.Save strLibKey
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
    Dim Index%, i%
    Index = 0: i = 1
    ' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
    Do While i > 0
        i = InStr(Index + 1, strReq, vbLf)
        If i > 0 Then Index = i
    Loop

    If Index <= 0 Then Exit Sub
        ' lockcode is from beginning to Index-1
    strLock = Left(strReq, Index - 1)
    ' user name starts from Index+1 to the end
    strUser = Mid$(strReq, Index + 1)
    strUser = TrimNulls(strUser)
    
End Sub
