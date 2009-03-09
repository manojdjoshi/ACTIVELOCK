VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IKeyStoreProvider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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
' Name: IKeyStoreProvider
' Purpose: This is the interface for a class that facilitates storing and
' retrieving of product license keys.
' Functions:
' Properties:
' Methods:
' Started: 21.04.2005
' Modified: 08.10.2005
'===============================================================================
' @author activelock-admins
' @version 3.0.0
' @date 20050810
'
'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
'
'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' @history
' <pre>
'   07.07.03 - mcrute    - Updated the header comments for this file.
'   08.03.03 - th2tran   - VBDox&#39;ed this class.
' </pre>

'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit


'===============================================================================
' Name: Property Let KeyStorePath
' Input:
'   ByRef path As String - Key store path
' Output: None
' Purpose: Specifies the path under which the keys are stored.
' Remarks: Example: path to a license file, or path to the Windows Registry hive
'===============================================================================
Public Property Let KeyStorePath(path As String)

End Property


'===============================================================================
' Name: Function Retrieve
' Input:
'   ByRef ProductName As String - Product Name
'   ByRef mLicenseFileType As ALLicenseFileTypes - License File Type
' Output:
'   ProductLicense - ProductLicense object matching the specified product name.
'   If no license found, then <code>Nothing</code> is returned.
' Purpose: Retrieves license info for the specified product name.
' Remarks: None
'===============================================================================
Public Function Retrieve(ProductName As String, mLicenseFileType As ALLicenseFileTypes) As ProductLicense

End Function

'===============================================================================
' Name: Sub Store
' Input:
'   ByRef Lic As ProductLicense - Product license
'   ByRef mLicenseFileType As ALLicenseFileTypes - License File Type
' Output: None
' Purpose: Stores a license.
' Remarks: None
'===============================================================================
Public Sub Store(Lic As ProductLicense, mLicenseFileType As ALLicenseFileTypes)

End Sub
