// fixtures.h : main header file for the fixtures application
//
#pragma once

#ifndef __AFXWIN_H__
#error include 'stdafx.h' before including this file for PCH
#endif

#include "license.h"
#include "ConnectionAdvisor.h"
#include "eventsink.h"
#include "hwid.h"

// CfixturesApp:
// See fixtures.cpp for the implementation of this class
//

class FindDirectory
{
public:
  CString Application();
  CString System();
};

class CActiveLockUtil 
{
public:
	CActiveLockUtil();
	~CActiveLockUtil();

	_IActiveLockPtr         AL;
	CActiveLockEventSink    ALES;

  License   lic;
  Hardware  hw;
  string    m_Email;
  enum ALLockTypes alLockType;
  // Overrides
public:
int xCrcs(bool debug, DWORD activelock3Crc, DWORD alcryptoCrc);

int readHW(string user, VectorDevice& devices, int index);
  void ReadHardware(vector<int> hwIndices);


	int  FindAndCheckLicense(string& softwareName, 
		string& version, 
		bool debug, 
		DWORD activelock3Crc,
    DWORD alcryptoCrc,
    string KeyStorePath,
		string AutoRegisterKeyPath ,
		string email,
    enum ALTrialTypes trialType,
 	  int trialNo
		);

	void Exiting();
	void HandleException(CException* ce, int reportMkr=0);
  int CRCCheck(string& strActive, DWORD alcryptoCrc, int& crcFails, string msg);

	void GetInfo(
		CString &            installCode,
		CString &        	  registeredDate  ,
		CString &        	  expiryDate      ,
		long &              usedDays        ,
		CString &        	  registeredUser  ,
		CString &        	  status  ,
		CString &        	  explain,
		CString &        	  licinfo,
		CString          	  explainIfNeeded
		);

};

