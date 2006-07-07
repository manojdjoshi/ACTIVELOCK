// fixtures.h : main header file for the fixtures application
//
#pragma once

#ifndef __AFXWIN_H__
#error include 'stdafx.h' before including this file for PCH
#endif

#include "ConnectionAdvisor.h"
#include "eventsink.h"

// CfixturesApp:
// See fixtures.cpp for the implementation of this class
//

class FindDirectory
{
public:
  CString Application();
  CString System();
};


struct CNameAndCrc
{
  CString name;
  CString msg;
  DWORD   crc;
};


typedef vector<CNameAndCrc>                 VectorNameAndCrc;
typedef vector<CNameAndCrc>::iterator       VectorNameAndCrcIterator;


class CActiveLockUtil 
{
public:
	CActiveLockUtil();
	~CActiveLockUtil();

  // Overrides
public:

  void HandleException(CException* ce, int reportMkr=0);
  BOOL CheckACrc(CString& strActive, DWORD alcryptoCrc, CString msg);
  BOOL  CheckAllCrcs(BOOL debug, VectorNameAndCrc& vNameCrc);
};

