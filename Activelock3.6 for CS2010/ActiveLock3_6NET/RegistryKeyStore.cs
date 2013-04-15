using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;
using ActiveLock3_6NET;
 // ERROR: Not supported in C#: OptionDeclaration
internal class RegistryKeyStoreProvider : _IKeyStoreProvider
{
	//*   ActiveLock
	//*   Copyright 1998-2002 Nelson Ferraz
	//*   Copyright 2006 The ActiveLock Software Group (ASG)
	//*   All material is the property of the contributing authors.
	//*
	//*   Redistribution and use in source and binary forms, with or without
	//*   modification, are permitted provided that the following conditions are
	//*   met:
	//*
	//*     [o] Redistributions of source code must retain the above copyright
	//*         notice, this list of conditions and the following disclaimer.
	//*
	//*     [o] Redistributions in binary form must reproduce the above
	//*         copyright notice, this list of conditions and the following
	//*         disclaimer in the documentation and/or other materials provided
	//*         with the distribution.
	//*
	//*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
	//*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
	//*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
	//*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
	//*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
	//*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
	//*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
	//*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
	//*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
	//*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
	//*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
	//*
	//*
	//===============================================================================
	// Name: RegistryKeyStoreProvider
	// Purpose: This IKeyStoreProvider implementation is used to  maintain the license keys in the registry.
	// Functions:
	// Properties:
	// Methods:
	// Started: 04.21.2005
	// Modified: 08.15.2005
	//===============================================================================
	// @author activelock-admins
	// @version 3.0.0
	// @date 20050815
	//
	//* ///////////////////////////////////////////////////////////////////////
	//  /                        MODULE TO DO LIST                            /
	//  ///////////////////////////////////////////////////////////////////////
	//
	//  ///////////////////////////////////////////////////////////////////////
	//  /                        MODULE CHANGE LOG                            /
	//  ///////////////////////////////////////////////////////////////////////
	//
	//   07.07.03 - mcrute   - Updated the header comments for this file.
	//
	//
	//  ///////////////////////////////////////////////////////////////////////
	//  /                MODULE CODE BEGINS BELOW THIS LINE                   /
	//  ///////////////////////////////////////////////////////////////////////


	//===============================================================================
	// Name: Function IKeyStoreProvider_Retrieve
	// Input:
	//   ProductCode As String - Product (software) code
	// Output:
	//   Productlicense - Product license object
	// Purpose:  Not implemented yet
	// Remarks: None
	//===============================================================================
	private ProductLicense IKeyStoreProvider_Retrieve(ref string ProductCode, IActiveLock.ALLicenseFileTypes mLicenseFileType)
	{
		// TODO: Implement Me
		return null;
	}
	ProductLicense _IKeyStoreProvider.Retrieve(ref string ProductCode, IActiveLock.ALLicenseFileTypes mLicenseFileType)
	{
		return IKeyStoreProvider_Retrieve(ProductCode, mLicenseFileType);
	}

	//===============================================================================
	// Name: Property Let IKeyStoreProvider_KeyStorePath
	// Input:
	//   RHS As String - Key store file path
	// Output: None
	// Purpose:  Not implemented yet
	// Remarks: None
	//===============================================================================
	private string IKeyStoreProvider_KeyStorePath {
			// TODO: Implement Me
		set { }
	}
	string _IKeyStoreProvider.KeyStorePath {
		set { IKeyStoreProvider_KeyStorePath = value; }
	}
	//===============================================================================
	// Name: Sub IKeyStoreProvider_Store
	// Input:
	//    Lic As ProductLicense - Product license object
	// Output: None
	// Purpose: Not implemented yet
	// Remarks: None
	//===============================================================================
	private void IKeyStoreProvider_Store(ref ProductLicense Lic, IActiveLock.ALLicenseFileTypes mLicenseFileType)
	{
		// TODO: Implement Me
	}
	void _IKeyStoreProvider.Store(ref ProductLicense Lic, IActiveLock.ALLicenseFileTypes mLicenseFileType)
	{
		IKeyStoreProvider_Store(Lic, mLicenseFileType);
	}
}
