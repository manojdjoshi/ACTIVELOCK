VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AlugenGlobals"
Attribute VB_GlobalNameSpace = True
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
' Name: AlugenGlobals
' Purpose: Global Accessors to ALUGENLib
' Remarks:
' Functions:
' Properties:
' Methods:
' Started: 08.15.2003
' Modified: 08.05.2005
'===============================================================================
'
' @author activelock-admins
' @version 3.0.0
' @date 08.05.2005
'
' * ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////

' * ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' @history
' <pre>
' 08.15.03 - th2tran       - Created
' 07.09.04 - th2tran       - Added ALUGENErrCodeConstants Enum
'
' </pre>

'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit



' ActiveLock Error Codes.
' These error codes are used for <code>Err.Number</code> whenever ActiveLock raises an error.
'
' @param alugenOK               No error.  Operation was successful.
' @param alugenProdInvalid      Product Info is invalid
Public Enum ALUGENErrCodeConstants
    alugenOK = 0   ' successful
    alugenProdInvalid = &H80040100         ' vbObjectError (&H80040000) + &H100
End Enum



'===============================================================================
' Name: Function GeneratorInstance
' Input: None
' Output:
'   IALUGenerator - New Generator instance
' Purpose: Returns a new Generator instance
' Remarks: None
'===============================================================================
Public Function GeneratorInstance(pProductStorageType As ProductsStoreType) As IALUGenerator
Select Case pProductStorageType
Case ProductsStoreType.alsINIFile
    Set GeneratorInstance = New INIGenerator
Case ProductsStoreType.alsXMLFile
    Set GeneratorInstance = New XMLGenerator
Case ProductsStoreType.alsMDBFile
    Set GeneratorInstance = New MDBGenerator
  'TODO - MSSQLGenerator
  'Case ProductsStoreType.alsMSSQL
  '  Set GeneratorInstance = New MSSQLGenerator
Case Else
    Set_locale regionalSymbol
    Err.Raise ActiveLockErrCodeConstants.alerrNotImplemented, ACTIVELOCKSTRING, STRNOTIMPLEMENTED
End Select
End Function

'===============================================================================
' Name: Function CreateProductInfo
' Input:
'   ByVal name As String - Product name
'   ByVal Ver As String - Product version
'   ByVal VCode As String - Product VCODE (public key)
'   ByVal GCode As String - Product GCODE (private key)
' Output:
'   ProductInfo - Product information
' Purpose: Instantiates a new ProductInfo object
' Remarks: None
'===============================================================================
Public Function CreateProductInfo(ByVal name As String, ByVal Ver As String, ByVal VCode As String, ByVal GCode As String) As ProductInfo
    Dim ProdInfo As New ProductInfo
    With ProdInfo
        .name = name
        .Version = Ver
        .VCode = VCode
        .GCode = GCode
    End With
    Set CreateProductInfo = ProdInfo
End Function

