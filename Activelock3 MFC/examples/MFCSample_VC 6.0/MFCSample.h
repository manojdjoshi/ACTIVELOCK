/**
 *   ActiveLock 3
 *   Copyright 2004 The ActiveLock Software Group (ASG)
 *   Portions Copyright by Simon Tatham and the PuTTY project.
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

/**********************************************************************************************
 * Change Log
 * ==========
 *
 * Date (MM/DD/YY)	Author		Description				
 * ---------------  ----------- --------------------------------------------------------------
 * 04/11/04			th2tran		Created
 *
 ***********************************************************************************************/

// MFCSample.h : main header file for the MFCSAMPLE application
//

#if !defined(AFX_MFCSAMPLE_H__96D514BA_5632_4EAB_B077_BF8616FB909A__INCLUDED_)
#define AFX_MFCSAMPLE_H__96D514BA_5632_4EAB_B077_BF8616FB909A__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif
#include <iostream>
#include <strstream>
#include <iomanip>
#include <string>
#include <vector>
using namespace std;

#include "resource.h"       // main symbols

#import "C:\\windows\\system32\\ActiveLock3.5.dll" 
using namespace ActiveLock3; 
#include "..\..\common\active\ActiveLockUtil.h"
#include "..\..\common\active\ActiveLockMFC.h"

/////////////////////////////////////////////////////////////////////////////
// CMFCSampleApp:
// See MFCSample.cpp for the implementation of this class
//

class CMFCSampleApp : public CWinApp
{
public:
	CMFCSampleApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CMFCSampleApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL
protected:
// Implementation
  CActiveLockMFC*   m_pActiveLockMfc;  // use a pointer - this should make it more compatable / interchangeable with Jereons class
  CActiveLockUtil util;
  CString strApp ;
  CString strIni ;
  enum ALLockTypes lockType;
  COleTemplateServer m_server;
  
  
  CString ProgramPath();
  BOOL CrcsEtc(int debug);
 
// Server object for document creation
	//{{AFX_MSG(CMFCSampleApp)
	afx_msg void OnAppAbout();
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
  virtual int ExitInstance();
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_MFCSAMPLE_H__96D514BA_5632_4EAB_B077_BF8616FB909A__INCLUDED_)
