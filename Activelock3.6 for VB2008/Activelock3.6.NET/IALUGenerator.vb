Option Strict Off
Option Explicit On

#Region "Copyright"
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
#End Region

''' <summary>
''' _IALUGenerator - Interface
''' </summary>
''' <remarks></remarks>
Public Interface _IALUGenerator
	WriteOnly Property StoragePath As String
	Sub SaveProduct(ByRef ProdInfo As ProductInfo)
	Function RetrieveProduct(ByVal name As String, ByVal Ver As String) As ProductInfo
	Function RetrieveProducts() As ProductInfo()
	Sub DeleteProduct(ByVal name As String, ByVal Ver As String)
    Function GenKey(ByRef Lic As ActiveLock3_6NET.ProductLicense, ByVal InstCode As String, Optional ByVal RegisteredLevel As String = "0") As String
End Interface

''' <summary>
''' Interface for the ActiveLock Universal Generator (ALUGEN)
''' </summary>
''' <remarks></remarks>
<System.Runtime.InteropServices.ProgId("IALUGenerator_NET.IALUGenerator")> Public Class IALUGenerator
    Implements _IALUGenerator
	' Started: 08.15.2003
	' Modified: 03.23.2006
	'===============================================================================
	' @author activelock-admins
	' @version 3.0.0
	' @date 03.23.2006

	''' <summary>
	''' Private Variable - mstrProductFile
	''' </summary>
	''' <remarks></remarks>
	Private mstrProductFile As String

	''' <summary>
	''' Private Variable - mINIFile
	''' </summary>
	''' <remarks></remarks>
	Private mINIFile As New INIFile

	''' <summary>
	''' StoragePath - Write Only - Specifies the path where information about the products is stored.
	''' </summary>
	''' <value>ByVal strPath As String - INI file path</value>
	''' <remarks></remarks>
    Public WriteOnly Property StoragePath() As String Implements _IALUGenerator.StoragePath
        Set(ByVal Value As String)
            mstrProductFile = Value
            mINIFile.File = Value
        End Set
	End Property

	''' <summary>
	''' SaveProduct - Saves a new product information to the product store.
	''' </summary>
	''' <param name="ProdInfo">ProdInfo As ProductInfo - Object containing product information to be saved.</param>
	''' <remarks>Raises error if product already exists.</remarks>
    Public Sub SaveProduct(ByRef ProdInfo As ProductInfo) Implements _IALUGenerator.SaveProduct

	End Sub

	''' <summary>
	''' RetrieveProduct - Retrieves product information.
	''' </summary>
	''' <param name="name">ByVal name As String - Product name</param>
	''' <param name="Ver">ByVal Ver As String - Product version</param>
	''' <returns>ProductInfo - Object containing product information.</returns>
	''' <remarks></remarks>
    Public Function RetrieveProduct(ByVal name As String, ByVal Ver As String) As ProductInfo Implements _IALUGenerator.RetrieveProduct
        RetrieveProduct = Nothing
	End Function

	''' <summary>
	''' RetrieveProducts - Retrieves all product infos.
	''' </summary>
	''' <returns>ProductInfo - Array of ProductInfo objects.</returns>
	''' <remarks></remarks>
    Public Function RetrieveProducts() As ProductInfo() Implements _IALUGenerator.RetrieveProducts
        RetrieveProducts = Nothing
	End Function

	''' <summary>
	''' DeleteProduct - Removes a product from the store.
	''' </summary>
	''' <param name="name">ByVal name As String - Product name</param>
	''' <param name="Ver">Ver As String - Product version</param>
	''' <remarks></remarks>
    Public Sub DeleteProduct(ByVal name As String, ByVal Ver As String) Implements _IALUGenerator.DeleteProduct

	End Sub

	''' <summary>
	''' GenKey - Generates a liberation key for the specified product.
	''' </summary>
	''' <param name="Lic">Lic As ActiveLock3.ProductLicense - License object for which to generate the liberation key.</param>
	''' <param name="InstCode">ByVal InstCode As String - User installation code</param>
	''' <param name="RegisteredLevel">ByVal RegisteredLevel As String - Level for which the user is allowed</param>
	''' <returns>String - Generated Liberation Key</returns>
	''' <remarks></remarks>
    Public Function GenKey(ByRef Lic As ActiveLock3_6NET.ProductLicense, ByVal InstCode As String, Optional ByVal RegisteredLevel As String = "0") As String Implements _IALUGenerator.GenKey
        GenKey = String.Empty
    End Function
End Class