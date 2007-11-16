VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RegistryKeyStoreProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2005 The ActiveLock Software Group (ASG)
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
' Name: RegistryKeyStoreProvider
' Purpose: This IKeyStoreProvider implementation is used to  maintain the license keys in the registry.
' Functions:
' Properties:
' Methods:
' Started: 04.21.2005
' Modified: 08.15.2005
'===============================================================================
' @author activelock-admins
' @version 3.0.0
' @date 20050815
'
'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
'
'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
'
'   07.07.03 - mcrute   - Updated the header comments for this file.
'
'
'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit

Implements IKeyStoreProvider

'===============================================================================
' Name: Function IKeyStoreProvider_Retrieve
' Input:
'   ProductCode As String - Product (software) code
'    mLicenseFileType As ALLicenseFileTypes - License File Type
' Output:
'   Productlicense - Product license object
' Purpose:  Not implemented yet
' Remarks: None
'===============================================================================
Private Function IKeyStoreProvider_Retrieve(ProductCode As String, mLicenseFileType As ALLicenseFileTypes) As ProductLicense
' TODO: Implement Me
End Function

'===============================================================================
' Name: Property Let IKeyStoreProvider_KeyStorePath
' Input:
'   RHS As String - Key store file path
' Output: None
' Purpose:  Not implemented yet
' Remarks: None
'===============================================================================
Private Property Let IKeyStoreProvider_KeyStorePath(RHS As String)
' TODO: Implement Me
End Property
'===============================================================================
' Name: Sub IKeyStoreProvider_Store
' Input:
'    Lic As ProductLicense - Product license object
'    mLicenseFileType As ALLicenseFileTypes - License File Type
' Output: None
' Purpose: Not implemented yet
' Remarks: None
'===============================================================================
Private Sub IKeyStoreProvider_Store(Lic As ProductLicense, mLicenseFileType As ALLicenseFileTypes)
' TODO: Implement Me
End Sub
