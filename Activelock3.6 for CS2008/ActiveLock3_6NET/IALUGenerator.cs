using Microsoft.VisualBasic;
using Microsoft.VisualBasic.Compatibility;
using System;
using System.Collections;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Windows.Forms;

#region "Copyright"
//*   ActiveLock
//*   Copyright 1998-2002 Nelson Ferraz
//*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
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
#endregion

/// <summary>
/// _IALUGenerator - Interface
/// </summary>
/// <remarks></remarks>
 // ERROR: Not supported in C#: OptionDeclaration
public interface _IALUGenerator
{
	string StoragePath {
		set;
	}
	void SaveProduct(ref ProductInfo ProdInfo);
	ProductInfo RetrieveProduct(string name, string Ver);
	ProductInfo[] RetrieveProducts();
	void DeleteProduct(string name, string Ver);
	string GenKey(ref ActiveLock3_6NET.ProductLicense Lic, string InstCode, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute("0")]  // ERROR: Optional parameters aren't supported in C#
string RegisteredLevel);
}

/// <summary>
/// Interface for the ActiveLock Universal Generator (ALUGEN)
/// </summary>
/// <remarks></remarks>
[System.Runtime.InteropServices.ProgId("IALUGenerator_NET.IALUGenerator")]
public class IALUGenerator : _IALUGenerator
{
	// Started: 08.15.2003
	// Modified: 03.23.2006
	//===============================================================================
	// @author activelock-admins
	// @version 3.0.0
	// @date 03.23.2006

	/// <summary>
	/// Private Variable - mstrProductFile
	/// </summary>
	/// <remarks></remarks>

	private string mstrProductFile;
	/// <summary>
	/// Private Variable - mINIFile
	/// </summary>
	/// <remarks></remarks>

	private INIFile mINIFile = new INIFile();
	/// <summary>
	/// StoragePath - Write Only - Specifies the path where information about the products is stored.
	/// </summary>
	/// <value>ByVal strPath As String - INI file path</value>
	/// <remarks></remarks>
	public string StoragePath {
		set {
			mstrProductFile = value;
			mINIFile.File = value;
		}
	}

	/// <summary>
	/// SaveProduct - Saves a new product information to the product store.
	/// </summary>
	/// <param name="ProdInfo">ProdInfo As ProductInfo - Object containing product information to be saved.</param>
	/// <remarks>Raises error if product already exists.</remarks>

	public void SaveProduct(ref ProductInfo ProdInfo)
	{
	}

	/// <summary>
	/// RetrieveProduct - Retrieves product information.
	/// </summary>
	/// <param name="name">ByVal name As String - Product name</param>
	/// <param name="Ver">ByVal Ver As String - Product version</param>
	/// <returns>ProductInfo - Object containing product information.</returns>
	/// <remarks></remarks>
	public ProductInfo RetrieveProduct(string name, string Ver)
	{
		return null;
	}

	/// <summary>
	/// RetrieveProducts - Retrieves all product infos.
	/// </summary>
	/// <returns>ProductInfo - Array of ProductInfo objects.</returns>
	/// <remarks></remarks>
	public ProductInfo[] RetrieveProducts()
	{
		return null;
	}

	/// <summary>
	/// DeleteProduct - Removes a product from the store.
	/// </summary>
	/// <param name="name">ByVal name As String - Product name</param>
	/// <param name="Ver">Ver As String - Product version</param>
	/// <remarks></remarks>

	public void DeleteProduct(string name, string Ver)
	{
	}

	/// <summary>
	/// GenKey - Generates a liberation key for the specified product.
	/// </summary>
	/// <param name="Lic">Lic As ActiveLock3.ProductLicense - License object for which to generate the liberation key.</param>
	/// <param name="InstCode">ByVal InstCode As String - User installation code</param>
	/// <param name="RegisteredLevel">ByVal RegisteredLevel As String - Level for which the user is allowed</param>
	/// <returns>String - Generated Liberation Key</returns>
	/// <remarks></remarks>
	public string GenKey(ref ActiveLock3_6NET.ProductLicense Lic, string InstCode, [System.Runtime.InteropServices.OptionalAttribute, System.Runtime.InteropServices.DefaultParameterValueAttribute("0")]  // ERROR: Optional parameters aren't supported in C#
string RegisteredLevel)
	{
		return string.Empty;
	}
}
