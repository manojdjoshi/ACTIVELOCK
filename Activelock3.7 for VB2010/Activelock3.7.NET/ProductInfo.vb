Option Strict Off
Option Explicit On

#Region "Copyright"
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
' *   Copyright 2003-2009 The Activelock - Ismail Alkan (ASG)
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
#End Region

<System.Runtime.InteropServices.ProgId("ProductInfo_NET.ProductInfo")> Public Class ProductInfo

    '===============================================================================
    ' Name: ProductInfo
    ' Purpose: This class encapsulates information about a product maintained by ALUGEN.
    ' Functions:
    ' Properties:
    ' Methods:
    ' Started: 07.07.2003
    ' Modified: 03.25.2006
    '===============================================================================
    ' @author activelock-admins
    ' @version 3.3.0
    ' @date 03.25.2006
    '
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
    '===============================================================================
	' Name: Property Let name
	' Input:
	'   ByVal sName As String - Product name
	' Output: None
	' Purpose: [INTERNAL] Specifies Product Name
	' Remarks: None
	'===============================================================================
    Public Property Name() As String
        Get
            Return mstrName
        End Get
        Set(ByVal Value As String)
            mstrName = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Get Version
    ' Input: None
    ' Output:
    '   String - Product version string
    ' Purpose: Gets product version
    ' Remarks: None
    '===============================================================================
    '===============================================================================
    ' Name: Property Let Version
    ' Input:
    '   ByVal sVer As String - Product version string
    ' Output: None
    ' Purpose: [INTERNAL] Specifies Product Version
    ' Remarks: None
    '===============================================================================
    Public Property Version() As String
        Get
            Return mstrVer
        End Get
        Set(ByVal Value As String)
            mstrVer = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Get VCode
    ' Input: None
    ' Output:
    '   String - Product VCode string
    ' Purpose: Gets the public encryption key used to verify product license keys.
    ' Remarks: None
    '===============================================================================
    '===============================================================================
    ' Name: Property Let VCode
    ' Input:
    '   ByRef sCode As String - Product VCode string
    ' Output: None
    ' Purpose: [INTERNAL] Specifies the public encryption key used to generate product license keys.
    ' Remarks: None
    '===============================================================================
    Public Property VCode() As String
        Get
            Return mstrCode1
        End Get
        Set(ByVal Value As String)
            mstrCode1 = Value
        End Set
    End Property
    '===============================================================================
    ' Name: Property Get GCode
    ' Input: None
    ' Output:
    '   String - Code string
    ' Purpose: Returns the private encryption key used to generate product license keys.
    ' Remarks: None
    '===============================================================================
    '===============================================================================
    ' Name: Property Let GCode
    ' Input:
    '   ByRef sCode As String - Product GCode
    ' Output: None
    ' Purpose: [INTERNAL] Specifies the private encryption key used to generate product license keys.
    ' Remarks: None
    '===============================================================================
    Public Property GCode() As String
        Get
            Return mstrCode2
        End Get
        Set(ByVal Value As String)
            mstrCode2 = Value
        End Set
    End Property
End Class