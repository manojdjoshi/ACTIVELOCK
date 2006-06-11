Option Strict Off
Option Explicit On
Public Interface _IALUGenerator
	WriteOnly Property StoragePath As String
	Sub SaveProduct(ByRef ProdInfo As ProductInfo)
	Function RetrieveProduct(ByVal name As String, ByVal Ver As String) As ProductInfo
	Function RetrieveProducts() As ProductInfo()
	Sub DeleteProduct(ByVal name As String, ByVal Ver As String)
    Function GenKey(ByRef Lic As ActiveLock3_5NET.ProductLicense, ByVal InstCode As String, Optional ByVal RegisteredLevel As String = "0") As String
End Interface
<System.Runtime.InteropServices.ProgId("IALUGenerator_NET.IALUGenerator")> Public Class IALUGenerator
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
    ' Name: IALUGenerator
    ' Purpose: Interface for the ActiveLock Universal Generator (ALUGEN)
    ' Functions:
    ' Properties:
    ' Methods:
    ' Started: 08.15.2003
    ' Modified: 03.23.2006
    '===============================================================================
    ' @author activelock-admins
    ' @version 3.0.0
    ' @date 03.23.2006
    '
    Private mstrProductFile As String
    Private mINIFile As New INIFile
    '===============================================================================
    ' Name: Property Let StoragePath
    ' Input:
    '   ByVal strPath As String - INI file path
    ' Output: None
    ' Purpose: Specifies the path where information about the products is stored.
    ' Remarks: None
    '===============================================================================
    Public WriteOnly Property StoragePath() As String Implements _IALUGenerator.StoragePath
        Set(ByVal Value As String)
            mstrProductFile = Value
            mINIFile.File = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Sub SaveProduct
    ' Input:
    '   ProdInfo As ProductInfo - Object containing product information to be saved.
    ' Output: None
    ' Purpose: Saves a new product information to the product store.
    ' Remarks: Raises error if product already exists.
    '===============================================================================
    Public Sub SaveProduct(ByRef ProdInfo As ProductInfo) Implements _IALUGenerator.SaveProduct

    End Sub
    '===============================================================================
    ' Name: Function RetrieveProduct
    ' Input:
    '   ByVal name As String - Product name
    '   ByVal Ver As String - Product version
    ' Output:
    '   ProductInfo - Object containing product information.
    ' Purpose: Retrieves product information.
    ' Remarks: None
    '===============================================================================
    Public Function RetrieveProduct(ByVal name As String, ByVal Ver As String) As ProductInfo Implements _IALUGenerator.RetrieveProduct
        RetrieveProduct = Nothing
    End Function
    '===============================================================================
    ' Name: Function RetrieveProducts
    ' Input: None
    ' Output:
    '   ProductInfo - Array of ProductInfo objects.
    ' Purpose: Retrieves all product infos.
    ' Remarks: None
    '===============================================================================
    Public Function RetrieveProducts() As ProductInfo() Implements _IALUGenerator.RetrieveProducts
        RetrieveProducts = Nothing
    End Function
    '===============================================================================
    ' Name: Sub DeleteProduct
    ' Input:
    '   ByVal name As String - Product name
    '   ByVal Ver As String - Product version
    ' Output: None
    ' Purpose:Removes a product from the store.
    ' Remarks: Raises Error if problems removing the product.
    '===============================================================================
    Public Sub DeleteProduct(ByVal name As String, ByVal Ver As String) Implements _IALUGenerator.DeleteProduct

    End Sub
    '===============================================================================
    ' Name: Function GenKey
    ' Input:
    '   Lic As ActiveLock3.ProductLicense - License object for which to generate the liberation key.
    '   ByVal InstCode As String - User installation code
    '   ByVal RegisteredLevel As String - Level for which the user is allowed
    ' Output:
    '   String - Generated Liberation Key
    ' Purpose: Generates a liberation key for the specified product.
    ' Remarks: None
    '===============================================================================
    Public Function GenKey(ByRef Lic As ActiveLock3_5NET.ProductLicense, ByVal InstCode As String, Optional ByVal RegisteredLevel As String = "0") As String Implements _IALUGenerator.GenKey
        GenKey = String.Empty
    End Function
End Class