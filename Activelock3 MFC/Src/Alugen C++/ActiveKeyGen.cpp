/**
*   ActiveLock 3
*   Copyright 2003-2005 The ActiveLock Software Group (ASG)
*
*   All material is the property of the contributing authors.
*
*   Redistribution and use in source and binary forms, with or without
*   modification, are permitted provided that the following conditions are
*   met:
*
*     [o] Redistributions of source code must retain the above copyright
*         notice, this list of conditions and the following disclaimer.
*
*     [o] Redistributions in binary form must reproduce the above
*         copyright notice, this list of conditions and the following
*         disclaimer in the documentation and/or other materials provided
*         with the distribution.
*
*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*
*/


/*
* ALUGENC - Console version of alugen.exe
*
* New Authors take over: 4/21/05 - ASG Group admins Scott Nelson (sentax), Ismail Alkin (ialkin), Terrence Kuchel (tkuchel)
* Visit http://activelock.sourceforge.net for support
*
* 
*
*/

/**********************************************************************************************
* Change Log
* ==========
*
* Date (MM/DD/YY)   Author      Description
* ---------------  ----------- --------------------------------------------------------------
* 10/01/05			   David      - All ActiveLock stuff in this file
*
***********************************************************************************************/


#include "stdafx.h"

#include <vector>
using namespace std;

#include "direct.h"
#import "C:\\windows\\system32\\ActiveLock3.5.dll" 
using namespace ActiveLock3; 


inline CString S(char *c)
{
  return c ? c : "";
}

void ActiveLockHandleException(_com_error ce, BOOL bShow) 
{ 
  const TCHAR* pCause = ce.ErrorMessage(); 
  if(bShow){ 
    AfxMessageBox(pCause);
  }
} 

/////////////////////////////////////////////////////////////////////////////
// alugenc.exe main
//
int ActveLockKeyGen(
                    CString strProdName,
                    CString strProdVer,
                    enum ALLicType licType,
                    CString strRegDate,
                    CString strExpire,
                    CString strInstCode,
                    CString strLicensee,
                    CString strLevel,
                    CString& strKey 
                    )
{
  HRESULT hr;
  SYSTEMTIME st;        // Registered Date
  GetLocalTime(&st);
  strRegDate.Format("%04d/%02d/%02d", st.wYear, st.wMonth, st.wDay);


  // Obtain ALUGEN instance via AlugenGlobal
  _AlugenGlobalsPtr alugenGlobals;
  hr = alugenGlobals.CreateInstance(_T("ActiveLock3.AlugenGlobals"),NULL);
  _IALUGeneratorPtr alugen;
  alugen = alugenGlobals->GeneratorInstance(0); // new parameter - this is a cop out 

  // finding products.ini is the biggest weakness of this program
  // currently this program must be placed in Src_v3 to find products.ini
  CString strPath;
  int index;
  // Calculate full path to products.ini based on the exe's directory
  char    cPath[255];
  GetModuleFileName(NULL, cPath, 255);
  strPath = cPath;
  index = strPath.ReverseFind('\\'); // find last index of '\'
  if (index > 0) {
    strPath = strPath.Left(index);
  }
  strPath.Insert(strPath.GetLength(), _T("\\products.ini"));
  // Make sure the file exists. If not, report an error and then fail.
  CFileFind        FF;
  if(!FF.FindFile(strPath)) {  
    strKey = _T("Failed: Products.ini not found at - ") + strPath ;
    AfxMessageBox(strKey); 
    return 1;
  }

  // Set StorePath 
  alugen->PutStoragePath(_bstr_t(strPath));
  // Generate key
  _ProductLicensePtr  Lic;
  CString licKey;
  try {
    // ActiveLock initialization
    _GlobalsPtr alGlobals; 
    // Obtain Global Instance
    hr = alGlobals.CreateInstance("ActiveLock3.Globals", NULL); 

    // using the old interface method the following were some of the parameters that neede replacing
    // keep as a record in case _bstr_t("") is not a good replacement for NULL
    //      NULL,           // product/software code    next procedure gets this from products.ini
    //      NULL,           // license key              *** we are trying to create this so at moment we do not know it
    //      NULL,           // Hash1                    no extra comment = I do not know what this is  

    Lic = alGlobals->CreateProductLicense(
      _bstr_t(strProdName), // product or appl. name    "excel"   used in next procedure to access products.ini
      _bstr_t(strProdVer),  // version No.              "3.2"     used in next procedure to access products.ini
      _bstr_t(""),          // product/software code    next procedure gets this from products.ini
      alfSingle,            // LicFlags:                alSingle=0 
      licType,              // ALLicType license type   allicTimeLocked = 3
      _bstr_t(strLicensee), // licensee registered user "David Weatherall"
      _bstr_t(strLevel),    // level                    3
      _bstr_t(strExpire),   // expiration date          "2006.4.3"
      _bstr_t(""),          // license key              *** we are trying to create this so at moment it is unknown
      _bstr_t(strRegDate),  // date of registration     "2006.1.1"
      _bstr_t(""),          // Hash1                     no extra comment = I do not know what this is  
      1,                    // max users                1
      ""                    // now believe this is installation code. This is provided in next routine maybe this should be strInstCode
      );

    struct _ProductLicense * pLic = Lic;  // this one line took me 8 hours to find, &Lic below is not the same

    licKey = S(alugen->GenKey(            // ** the final answer the total license 
      &pLic,                              // not the same as &Lic - lots of operators overridden
      _bstr_t(strInstCode),               // installation code        "Nz678T99985"
      _bstr_t(strLevel)                   // registration level       "1"
      ));


  } catch (_com_error ce) {
    const TCHAR* pCause = ce.ErrorMessage();
    CString mes("Failed: ");
    mes += CString(pCause);
    AfxMessageBox(mes);
    strKey = pCause;
    return 1;
  } catch (...) {
    strKey = "Failed - Unknown exception";
    AfxMessageBox(strKey); 
    return 1;
  }

  strKey = licKey ;
  return 0;
}
