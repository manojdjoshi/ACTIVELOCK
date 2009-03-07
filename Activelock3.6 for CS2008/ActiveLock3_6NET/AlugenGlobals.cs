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
/// Global Accessors to ALUGENLib
/// </summary>
/// <remarks>Class instancing was changed to public.</remarks>
[System.Runtime.InteropServices.ProgId("AlugenGlobals_NET.AlugenGlobals")]
public class AlugenGlobals
{

	// Started: 08.15.2003
	// Modified: 03.23.2006
	//===============================================================================
	//
	// @author activelock-admins
	// @version 3.3.0
	// @date 03.23.2006

	/// <summary>
	/// <para>ActiveLock Error Codes.</para>
	/// <para>These error codes are used for <code>Err.Number</code> whenever ActiveLock raises an error</para>
	/// </summary>
	/// <remarks></remarks>
	public enum alugenErrCodeConstants
	{
		/// <summary>
		/// No error.  Operation was successful.
		/// </summary>
		/// <remarks></remarks>
		alugenOk = 0,
		// successful
		/// <summary>
		/// Product Info is invalid
		/// </summary>
		/// <remarks></remarks>
		alugenProdInvalid = 0x80040100
		// vbObjectError (&H80040000) + &H100
	}

	/// <summary>
	/// Returns a new Generator instance
	/// </summary>
	/// <param name="pProductStorageType">IActiveLock.ProductsStoreType - Storage Type!</param>
	/// <returns>IALUGenerator - New Generator instance</returns>
	/// <remarks></remarks>

	public _IALUGenerator GeneratorInstance(IActiveLock.ProductsStoreType pProductStorageType)
	{
		_IALUGenerator functionReturnValue = null;
		switch (pProductStorageType) {
			case IActiveLock.ProductsStoreType.alsINIFile:
				functionReturnValue = new INIGenerator();
				break;
			case IActiveLock.ProductsStoreType.alsXMLFile:
				functionReturnValue = new XMLGenerator();
				break;
			case IActiveLock.ProductsStoreType.alsMDBFile:
				functionReturnValue = new MDBGenerator();
				break;
			//TODO - MSSQLGenerator
			//Case ProductsStoreType.alsMSSQL
			//  Set GeneratorInstance = New MSSQLGenerator
			default:
				modActiveLock.Set_Locale(modActiveLock.regionalSymbol);
				Err().Raise(Globals.ActiveLockErrCodeConstants.AlerrNotImplemented, modTrial.ACTIVELOCKSTRING, modActiveLock.STRNOTIMPLEMENTED);
				functionReturnValue = null;
				break;
		}
		return functionReturnValue;
	}

	/// <summary>
	/// Instantiates a new ProductInfo object
	/// </summary>
	/// <param name="Name">String - Product name</param>
	/// <param name="Ver">String - Product version</param>
	/// <param name="VCode">String - Product VCODE (public key)</param>
	/// <param name="GCode">String - Product GCODE (private key)</param>
	/// <returns>ProductInfo - Product information</returns>
	/// <remarks></remarks>
	public ProductInfo CreateProductInfo(string Name, string Ver, string VCode, string GCode)
	{
		ProductInfo ProdInfo = new ProductInfo();
		{
			ProdInfo.Name = Name;
			ProdInfo.Version = Ver;
			ProdInfo.VCode = VCode;
			ProdInfo.GCode = GCode;
		}
		return ProdInfo;
	}
}
