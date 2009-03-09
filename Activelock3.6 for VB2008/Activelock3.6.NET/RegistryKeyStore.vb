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
#End Region

Friend Class RegistryKeyStoreProvider
	Implements _IKeyStoreProvider

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
    '   07.07.03 - mcrute   - Updated the header comments for this file.
    '
    '
    '  ///////////////////////////////////////////////////////////////////////
    '  /                MODULE CODE BEGINS BELOW THIS LINE                   /
    '  ///////////////////////////////////////////////////////////////////////

    '===============================================================================
    ' Name: Function IKeyStoreProvider_Retrieve
    ' Input:
    '   ProductCode As String - Product (software) code
    ' Output:
    '   Productlicense - Product license object
    ' Purpose:  Not implemented yet
    ' Remarks: None
    '===============================================================================
    Private Function IKeyStoreProvider_Retrieve(ByRef ProductCode As String, ByVal mLicenseFileType As IActiveLock.ALLicenseFileTypes) As ProductLicense Implements _IKeyStoreProvider.Retrieve
        ' TODO: Implement Me
        Return Nothing
    End Function

    '===============================================================================
    ' Name: Property Let IKeyStoreProvider_KeyStorePath
    ' Input:
    '   RHS As String - Key store file path
    ' Output: None
    ' Purpose:  Not implemented yet
    ' Remarks: None
    '===============================================================================
    Private WriteOnly Property IKeyStoreProvider_KeyStorePath() As String Implements _IKeyStoreProvider.KeyStorePath
        Set(ByVal Value As String)
            ' TODO: Implement Me
        End Set
    End Property
    '===============================================================================
    ' Name: Sub IKeyStoreProvider_Store
    ' Input:
    '    Lic As ProductLicense - Product license object
    ' Output: None
    ' Purpose: Not implemented yet
    ' Remarks: None
    '===============================================================================
    Private Sub IKeyStoreProvider_Store(ByRef Lic As ProductLicense, ByVal mLicenseFileType As IActiveLock.ALLicenseFileTypes) Implements _IKeyStoreProvider.Store
        ' TODO: Implement Me
    End Sub
End Class