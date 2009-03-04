Option Strict On
Option Explicit On

#Region "Copyright"
'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
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
#End Region

''' <summary>
''' Global Accessors to ALUGENLib
''' </summary>
''' <remarks>Class instancing was changed to public.</remarks>
<System.Runtime.InteropServices.ProgId("AlugenGlobals_NET.AlugenGlobals")> Public Class AlugenGlobals

    ' Started: 08.15.2003
    ' Modified: 03.23.2006
    '===============================================================================
    '
    ' @author activelock-admins
    ' @version 3.3.0
    ' @date 03.23.2006

    ''' <summary>
    ''' <para>ActiveLock Error Codes.</para>
    ''' <para>These error codes are used for <code>Err.Number</code> whenever ActiveLock raises an error</para>
    ''' </summary>
    ''' <remarks></remarks>
    Public Enum alugenErrCodeConstants
        ''' <summary>
        ''' No error.  Operation was successful.
        ''' </summary>
        ''' <remarks></remarks>
        alugenOk = 0 ' successful
        ''' <summary>
        ''' Product Info is invalid
        ''' </summary>
        ''' <remarks></remarks>
        alugenProdInvalid = &H80040100 ' vbObjectError (&H80040000) + &H100
    End Enum

    ''' <summary>
    ''' Returns a new Generator instance
    ''' </summary>
    ''' <param name="pProductStorageType">IActiveLock.ProductsStoreType - Storage Type!</param>
    ''' <returns>IALUGenerator - New Generator instance</returns>
    ''' <remarks></remarks>
  Public Function GeneratorInstance(ByVal pProductStorageType As IActiveLock.ProductsStoreType) As _IALUGenerator

    Select Case pProductStorageType
      Case IActiveLock.ProductsStoreType.alsINIFile
        GeneratorInstance = New INIGenerator
      Case IActiveLock.ProductsStoreType.alsXMLFile
        GeneratorInstance = New XMLGenerator
      Case IActiveLock.ProductsStoreType.alsMDBFile
        GeneratorInstance = New MDBGenerator
        'TODO - MSSQLGenerator
        'Case ProductsStoreType.alsMSSQL
        '  Set GeneratorInstance = New MSSQLGenerator
      Case Else
                Set_Locale(regionalSymbol)
                Err.Raise(Globals.ActiveLockErrCodeConstants.AlerrNotImplemented, ACTIVELOCKSTRING, STRNOTIMPLEMENTED)
                GeneratorInstance = Nothing
        End Select
  End Function

    ''' <summary>
    ''' Instantiates a new ProductInfo object
    ''' </summary>
    ''' <param name="Name">String - Product name</param>
    ''' <param name="Ver">String - Product version</param>
    ''' <param name="VCode">String - Product VCODE (public key)</param>
    ''' <param name="GCode">String - Product GCODE (private key)</param>
    ''' <returns>ProductInfo - Product information</returns>
    ''' <remarks></remarks>
  Public Function CreateProductInfo(ByVal Name As String, ByVal Ver As String, ByVal VCode As String, ByVal GCode As String) As ProductInfo
    Dim ProdInfo As New ProductInfo
    With ProdInfo
      .Name = Name
      .Version = Ver
      .VCode = VCode
      .GCode = GCode
    End With
    CreateProductInfo = ProdInfo
  End Function
End Class