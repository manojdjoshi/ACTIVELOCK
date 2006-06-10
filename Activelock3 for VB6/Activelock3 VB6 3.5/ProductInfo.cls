VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProductInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2005 The ActiveLock Software Group (ASG)
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
' Name: ProductInfo
' Purpose: This class encapsulates information about a product maintained by ALUGEN.
' Functions:
' Properties:
' Methods:
' Started: 07.07.2003
' Modified: 08.10.2005
'===============================================================================
' @author activelock-admins
' @version 3.0.0
' @date 20050810
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
' 09.21.03 - th2tran       - Fixed bug: Name() property wasn't implemented.
'                            Originally, the methods were there but had no code.
' 10.13.03 - th2tran       - Renamed Code1 and Code2 to VCode and GCode respectively.
' </pre>

'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit

Private mstrName As String
Private mstrVer As String
Private mstrCode1 As String
Private mstrCode2 As String


'===============================================================================
' Name: Property Get name
' Input: None
' Output:
'   String - Product name string
' Purpose: Gets product name
' Remarks: None
'===============================================================================
Public Property Get name() As String
    name = mstrName
End Property

'===============================================================================
' Name: Property Let name
' Input:
'   ByVal sName As String - Product name
' Output: None
' Purpose: [INTERNAL] Specifies Product Name
' Remarks: None
'===============================================================================
Friend Property Let name(ByVal sName As String)
    mstrName = sName
End Property


'===============================================================================
' Name: Property Get Version
' Input: None
' Output:
'   String - Product version string
' Purpose: Gets product version
' Remarks: None
'===============================================================================
Public Property Get Version() As String
    Version = mstrVer
End Property

'===============================================================================
' Name: Property Let Version
' Input:
'   ByVal sVer As String - Product version string
' Output: None
' Purpose: [INTERNAL] Specifies Product Version
' Remarks: None
'===============================================================================
Friend Property Let Version(ByVal sVer As String)
    mstrVer = sVer
End Property


'===============================================================================
' Name: Property Get VCode
' Input: None
' Output:
'   String - Product VCode string
' Purpose: Gets the public encryption key used to verify product license keys.
' Remarks: None
'===============================================================================
Public Property Get VCode() As String
    VCode = mstrCode1
End Property

'===============================================================================
' Name: Property Let VCode
' Input:
'   ByRef sCode As String - Product VCode string
' Output: None
' Purpose: [INTERNAL] Specifies the public encryption key used to generate product license keys.
' Remarks: None
'===============================================================================
Friend Property Let VCode(sCode As String)
    mstrCode1 = sCode
End Property


'===============================================================================
' Name: Property Get GCode
' Input: None
' Output:
'   String - Code string
' Purpose: Returns the private encryption key used to generate product license keys.
' Remarks: None
'===============================================================================
Public Property Get GCode() As String
    GCode = mstrCode2
End Property

'===============================================================================
' Name: Property Let GCode
' Input:
'   ByRef sCode As String - Product GCode
' Output: None
' Purpose: [INTERNAL] Specifies the private encryption key used to generate product license keys.
' Remarks: None
'===============================================================================
Friend Property Let GCode(sCode As String)
    mstrCode2 = sCode
End Property
