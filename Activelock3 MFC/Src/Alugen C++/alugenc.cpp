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
* 04/21/05			sentax		 - Upgraded source to reflect ActiveLock3 version
* 06/28/04         th2tran      Created
* 07/09/04         th2tran      - Replaced ce->ReportException (which displays an error messagebox)
*								   with cerr output to keep the program free of any GUI
*								 - products.ini is now expected to reside in the same directory as alugenc.exe.
*								   Previously, it was assumed to be in the working directory.
* 07/10/04         th2tran      - Fixed bug: upon error, should be returning 1 instead of FALSE 
* 07/27/04         th2tran      - Forgot to set RegisteredDate for the license. Woops!
*
***********************************************************************************************/

#include "stdafx.h"

#include <vector>
using namespace std;

#include "direct.h"
#import "C:\\windows\\system32\\ActiveLock3.4.dll" 
using namespace ActiveLock3; 

#include "activekeygen.h"
#include "alugenc.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// The one and only application object

CWinApp theApp;


/////////////////////////////////////////////////////////////////////////////
// alugenc.exe main
//
int _tmain(int argc, TCHAR* argv[], TCHAR* envp[])
{
  int sortOfDebug = 0;

  // Print syntax
  if (argc < 15) {
    cerr << _T("ALUGENC - ActiveLock Universal GENerator Command-Line Interface\nSyntax: alugenc.exe -p <productname> -v <productversion> -t <licensetype> -x <expiration_date> -i <installation_code> -u <user> -l <level>") << endl;
    cerr << _T("Example: alugenc.exe -p TestApp -v 1.0 -t 0 -x 2005/10/27 -i YXRpb24gVXNlcg -u \"Eval User\" -l \"0\"") << endl;
    return 1;
  }

  // initialize MFC and print and error on failure
  if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0))
  {
    cerr << _T("Fatal Error: MFC initialization failed") << endl;
    return 1;
  }

  // Initialize OLE libraries
  if (!AfxOleInit())
  {
    cerr << _T("Fatal Error: OLE initialization failed") << endl;
    return 1;
  }

  CoInitialize(NULL);

  // Obtain ProdName, ProdVer, lictype, regdate, expiredate, instcode from command-line arguments
  CString strProdName;
  CString strProdVer;
  ALLicType licType = allicTimeLocked;
  CString strRegDate;
  CString strExpire;
  CString strInstCode;
  CString strLicensee;
  CString strLevel;
  CString strKey;
  CString strPassThru;

  // Parse command-line arguments
  CString temp;
  for (int i=1; i<argc; i++) {
    temp = argv[i];
    if (temp.CompareNoCase(_T("-p")) == 0) {
      strProdName = argv[i+1];
      i++;
    } else if (temp.CompareNoCase(_T("-v")) == 0) {
      strProdVer = argv[i+1];
      i++;
    } else if (temp.CompareNoCase(_T("-t")) == 0) {
      int t = atoi(argv[i+1]);
      licType = t==0 ? allicNone : ( t==1 ? allicPeriodic :( t==2 ? allicPermanent : allicTimeLocked ));
      i++;
    } else if (temp.CompareNoCase(_T("-x")) == 0) {
      strExpire = argv[i+1];
      i++;
    } else if (temp.CompareNoCase(_T("-i")) == 0) {
      strInstCode = argv[i+1];
      i++;
    } else if (temp.CompareNoCase(_T("-u")) == 0) {
      strLicensee = argv[i+1];
      i++;
    } else if (temp.CompareNoCase(_T("-l")) == 0) {
      strLevel = argv[i+1];
      i++;
    } else if (temp.CompareNoCase(_T("-d")) == 0) {
      strPassThru = argv[i+1];
      i++;
    }
  }

  // all activelock stuff is in this proc
  int result =  ActveLockKeyGen(
    strProdName,
    strProdVer,
    licType,
    strRegDate,
    strExpire,
    strInstCode,
    strLicensee,
    strLevel,
    strKey
    );
  if(result){
    cerr << "Error: " << strKey << endl;
    if(sortOfDebug){ // while debugging
      char fini[255];
      cin >> fini;
    }
    return 1;
  }

  cout << (LPCTSTR)strPassThru << endl;
  cout << (LPCTSTR)strKey << endl;

  if(sortOfDebug){ // while debugging
    char fini[255];
    cin >> fini;
  }

  return 0;
}