VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MDBGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 0
Implements IALUGenerator

Private MyActiveLock As ActiveLock3.IActiveLock

Dim oConnect As ADODB.Connection
Dim oRST As ADODB.Recordset
Dim myCmd As ADODB.Command

Private fileMDB As String

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
  Set oConnect = New ADODB.Connection
  Set oRST = New ADODB.Recordset
End Sub

'===============================================================================
' Name: Sub LoadProdInfo
' Input:
'   ByRef Section As String - Section Name that contains ProdName and ProdVer in order to be unique
'   ByRef ProdInfo As ProductInfo - Object containing product information to be saved.
' Output: None
' Purpose: Loads Product Info from the specified MDB database.
' Remarks: None
'===============================================================================
Private Sub LoadProdInfo(pRecordset As Recordset, ProdInfo As ProductInfo)
  ProdInfo.name = pRecordset("name")
  ProdInfo.Version = pRecordset("version")
  ProdInfo.VCode = pRecordset("vcode")
  ProdInfo.GCode = pRecordset("gcode")
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
  'save product in database
  
  Dim strSQL As String
  'open connection to MDB file
  oConnect.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
    "Data Source= " & fileMDB & ";"
  
  strSQL = "INSERT INTO products (name, version, vcode, gcode) VALUES ('" & ProdInfo.name & "', '" & ProdInfo.Version & "', '" & ProdInfo.VCode & "', '" & ProdInfo.GCode & "')"

  Set myCmd = New ADODB.Command
  myCmd.ActiveConnection = oConnect
  myCmd.CommandTimeout = 0
  myCmd.CommandText = strSQL
  myCmd.CommandType = adCmdText
  myCmd.Execute
  oConnect.Close

End Sub

'===============================================================================
' Name: IALUGenerator_RetrieveProducts
' Input: None
' Output:
'   ProductInfo - Product info object
' Purpose: Retrieves all product information from MDB.
' Remarks: Returns as an array.
'===============================================================================
Private Function IALUGenerator_RetrieveProducts() As ProductInfo()
    'Retrieve all product information from MDB file.  Return as an array.
    On Error GoTo RetrieveProductsError
    Dim arrProdInfos() As ProductInfo
    
    'open connection to MDB file
    oConnect.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source= " & fileMDB & ";"

    'open recordset
    oRST.Open "select * from products", oConnect
    
    Dim Count As Integer
    Count = 0
    Do Until oRST.EOF
      ReDim Preserve arrProdInfos(Count)
      Set arrProdInfos(Count) = New ProductInfo
      'load product info
      LoadProdInfo oRST, arrProdInfos(Count)
        
      Count = Count + 1
      oRST.MoveNext
    Loop
    'close recordset
    oRST.Close
    'Close connection
    oConnect.Close
    
    IALUGenerator_RetrieveProducts = arrProdInfos
    Exit Function
    
RetrieveProductsError:
    'close recordset
    oRST.Close
    'Close connection
    oConnect.Close
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
    
    'open connection to MDB file
    oConnect.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source= " & fileMDB & ";"

    'open recordset
    oRST.Open "select * from products where name='" & name & "' AND version='" & Ver & "'", oConnect
    
    If oRST.RecordCount > 0 Then
      ProdInfo.VCode = oRST("vcode")
      ProdInfo.GCode = oRST("gcode")
    End If
    
    oRST.Close
    oConnect.Close
    
    If ProdInfo.VCode = "" Or ProdInfo.GCode = "" Then
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
' Purpose: Removes the license keys section from a MDB database, i.e. deletes product details in the license database
'===============================================================================
Private Sub IALUGenerator_DeleteProduct(ByVal name As String, ByVal Ver As String)
    ' Remove the section from MDB file
    
    Dim strSQL As String
    'open connection to MDB file
    oConnect.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
      "Data Source= " & fileMDB & ";"
    
    strSQL = "DELETE FROM products WHERE name='" & name & "' AND version='" & Ver & "'"

    Set myCmd = New ADODB.Command
    myCmd.ActiveConnection = oConnect
    myCmd.CommandTimeout = 0
    myCmd.CommandText = strSQL
    myCmd.CommandType = adCmdText
    myCmd.Execute
    oConnect.Close

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
  Else
    fileMDB = RHS
  End If
End Property

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
    ' Take request code and decrypt it.
    Dim strReq As String
    ' 05.13.05 - ialkan Modified to merge DLLs into one
    strReq = Base64_Decode(InstCode)
    ' strReq now contains the {LockCode + vbLf + User} string
    Dim strLock$, strUser$
    GetLockAndUserFromInstallCode strReq, strLock, strUser
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
    Set ProdInfo = IALUGenerator_RetrieveProduct(Lic.ProductName, Lic.ProductVer)
    Lic.ProductKey = ProdInfo.VCode

    '@todo Check for "ProdInfo Is Nothing" and handle appropriately

    Dim strLic As String
    strLic = Lic.ToString() & vbLf & strLock
Debug.Print "strLic: " & vbCrLf & strLic

    ' sign it
    Dim strSig As String: strSig = String(1024, 0)
    ' 05.13.05 - ialkan Modified to merge DLLs into one. Moved RSASign into a module
    strSig = RSASign(ProdInfo.VCode, ProdInfo.GCode, strLic)

    ' Create liberation key.  This will be a base-64 encoded string of the whole license.
    Dim strLicKey As String
    ' 05.13.05 - ialkan Modified to merge DLLs into one
    strLicKey = Base64_Encode(strSig)
    ' update Lic with license key
    Lic.LicenseKey = strLicKey
' Print some info for debugging purposes
Debug.Print "VCode: " & ProdInfo.VCode
Debug.Print "Lic: " & strLic
Debug.Print "Lic hash: " & modMD5.Hash(strLic)
Debug.Print "LicKey: " & strLicKey
Debug.Print "Sig: " & strSig
Debug.Print "Verify: " & RSAVerify(ProdInfo.VCode, strLic, Base64_Decode(strLicKey))
Debug.Print "===================================================="

    ' Serialize it into a formatted string
    Dim strLibKey As String
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
