VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IALUGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
'===============================================================================
' Name: IALUGenerator
' Purpose: Interface for the ActiveLock Universal Generator (ALUGEN)
' Functions:
' Properties:
' Methods:
' Started: 08.15.2003
' Modified: 08.08.2005
'===============================================================================
' @author activelock-admins
' @version 3.0.0
' @date 20050808
'
'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////

'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' @history
' <pre>
' 08.15.03 - th2tran       - Created
' 09.21.03 - th2tran       - Changed GenKey() interface signature from
'                               GenKey(ByVal Name As String, ByVal Ver As String, ByVal ReqCode As String, LicType As Long, ExpireDate As String)
'                            to
'                               GenKey(objActiveLock As IActiveLock, ByVal ReqCode As String, ExpireDate As String)
' 10.13.03 - th2tran       - Renamed Code1 and Code2 to VCode and GCode respectively.
' 11.02.03 - th2tran       - GenKey interface changed (again) from
'                                GenKey(objActiveLock As IActiveLock, ByVal ReqCode As String, ExpireDate As String) As String
'                             to
'                                GenKey(Lic As ActiveLock3.ProductLicense, ByVal InstCode As String, ByVal RegisteredLevel As String) As String
' 06.28.04 - th2tran       - Added ByVal qualifier to IALUGENerator_StoragePath parameter for better compatibility /w VC++
' 05.13.05 - ialkan        - Modified to add the Register to Level X feature
' </pre>

'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit

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
Public Property Let StoragePath(ByVal strPath As String)
Attribute StoragePath.VB_Description = "Specifies the path where information about the products is stored."
    mstrProductFile = strPath
    mINIFile.File = strPath
End Property


'===============================================================================
' Name: Sub SaveProduct
' Input:
'   ProdInfo As ProductInfo - Object containing product information to be saved.
' Output: None
' Purpose: Saves a new product information to the product store.
' Remarks: Raises error if product already exists.
'===============================================================================
Public Sub SaveProduct(ProdInfo As ProductInfo)
Attribute SaveProduct.VB_Description = "Saves a new product information to the product store."

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
Public Function RetrieveProduct(ByVal name As String, ByVal Ver As String) As ProductInfo
Attribute RetrieveProduct.VB_Description = "Retrieves product information."

End Function


'===============================================================================
' Name: Function RetrieveProducts
' Input: None
' Output:
'   ProductInfo - Array of ProductInfo objects.
' Purpose: Retrieves all product infos.
' Remarks: None
'===============================================================================
Public Function RetrieveProducts() As ProductInfo()
Attribute RetrieveProducts.VB_Description = "Retrieves all information about all products in the store."

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
Public Sub DeleteProduct(ByVal name As String, ByVal Ver As String)
Attribute DeleteProduct.VB_Description = "Removes a product from the store."

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
Public Function GenKey(Lic As ActiveLock3.ProductLicense, ByVal InstCode As String, Optional ByVal RegisteredLevel As String = "0") As String
Attribute GenKey.VB_Description = "Generates liberation key for the specified product."

End Function
