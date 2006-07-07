// about.h : header file
//

#pragma once
#include "afxwin.h"

#import "C:\\windows\\system32\\ActiveLock3.4.dll" 
using namespace ActiveLock3; 
#include "..\..\common\active\ActiveLockUtil.h"
#include "..\..\common\active\ActiveLockMfc.h"

class CAboutDlg : public CDialog
{
public:
	CAboutDlg(CActiveLockMFC& _mfc);

// Dialog Data
	enum { IDD = IDD_ABOUTBOX };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Implementation
protected:
	DECLARE_MESSAGE_MAP()
public:
  CString             installCode;
  CString         	  registeredDate  ;
  CString         	  expiryDate      ;
	long                usedDays        ;
  CString         	  registeredUser  ;
  CString         	  status  ;
  CString         	  explain;
  CString         	  licinfo;
  int                 lockIndex;
  ALLockTypes         lockType;

  BOOL                InstallationCodeFromHardware;
  BOOL                GenerateInstallCode;

	CActiveLockMFC& mfc;

  void OnBnClickedLicBtn();
  void BnSwitcher();

  afx_msg void OnBnClickedLicNone();
  afx_msg void OnBnClickedLicWindows();
  afx_msg void OnBnClickedLicCompName();
  afx_msg void OnBnClickedLicDiskOld();
  afx_msg void OnBnClickedLicMAC();
  afx_msg void OnBnClickedLicBIOS();
  afx_msg void OnBnClickedLicIP();
  afx_msg void OnBnClickedLicMotherboard();
  afx_msg void OnBnClickedLicDisk();

  afx_msg void OnShowWindow(BOOL bShow, UINT nStatus);
  afx_msg void OnBnClickedGenerate();
};



