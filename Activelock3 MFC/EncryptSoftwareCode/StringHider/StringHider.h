// StringHider.h : main header file for the StringHider application
//
#pragma once

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"       // main symbols


// CStringHiderApp:
// See StringHider.cpp for the implementation of this class
//

class CStringHiderApp : public CWinApp
{
public:
	CStringHiderApp();


// Overrides
public:
	virtual BOOL InitInstance();

// Implementation
	afx_msg void OnAppAbout();
	DECLARE_MESSAGE_MAP()
};

extern CStringHiderApp theApp;