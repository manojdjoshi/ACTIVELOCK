// about.cpp : implementation file
//

#include "stdafx.h"
#include <string>
#include <vector>
#include <algorithm>
#include <strstream>
#include <iomanip>

using namespace std;

#include "resource.h"
#include "about.h"
#include ".\about.h"

// #define _CRTDBG_MAP_ALLOC 

const int lockSize = 14;
int ALLockTypesToIndex(ALLockTypes alt)
{
  const int locks[] = {
    lockNone ,
      lockMAC ,
      lockComp ,
      lockHD ,
      lockHDFirmware ,
      lockWindows ,
      lockBIOS ,
      lockMotherboard ,
      lockIP ,
      lockFingerprint,
      lockMemory,
      lockCPUID,
      lockBaseboardID,
      lockvideoID
  };
  for(int i=0; i<lockSize; i++)
    if(alt == locks[i])
      return i;
  return -1;
}

// CAboutDlg dialog used for App About

CAboutDlg::CAboutDlg(CActiveLockMFC& _mfc) : CDialog(CAboutDlg::IDD)
, registeredUser(_T(""))
, installCode(_T(""))
, expiryDate(_T("")) 
, usedDays(0)
, registeredDate(_T(""))
, status(_T(""))
, explain(_T(""))
, licinfo(_T(""))
, lockIndex(8)   //DW 8
, GenerateInstallCode(TRUE)
, mfc(_mfc)
{
  if(mfc.RegStatus() && mfc.LicenseStatus() == "Registered" ){
    InstallationCodeFromHardware = FALSE;
  } else {
    InstallationCodeFromHardware = TRUE;
  }
}

void CAboutDlg::DoDataExchange(CDataExchange* pDX)
{
  CDialog::DoDataExchange(pDX);

  CString explainIfNeeded = "Obtain an installation code by selecting either network or disk. \
Copy Installation Code to an e-mail.\
Tip: to avoid mistakes, use mouse to select code then copy with Control C \
Inside the email use Control V to paste. \
Send to ActiveLock@ActiveLock.com";

  explain						= "";
  registeredDate 		= "" ;
  status 						= mfc.LicenseStatus();
  registeredUser    = mfc.RegisteredUser();
  if(InstallationCodeFromHardware){
    explain = explainIfNeeded;
  }  else {
    // everything is in license
    registeredDate 		= mfc.GetRegisteredDate();
  }
  DDX_Text(pDX, IDC_LIC_USERNAME,    registeredUser);
  DDX_Text(pDX, IDC_LIC_INSTALLCODE, installCode);
  DDX_Text(pDX, IDC_LIC_EXPIRYDATE,  expiryDate);
  DDX_Text(pDX, IDC_LIC_USEDDAYS,    usedDays);
  DDX_Text(pDX, IDC_LIC_STATUS, 		 status);
  DDX_Text(pDX, IDC_LIC_REGDATE, 		 registeredDate);
  DDX_Text(pDX, IDC_LIC_EXPLAIN, 		 explain);
  DDX_Text(pDX, IDC_LICINFO, 		     licinfo);
  if(mfc.RegStatus()){
  }else{
  }

}
/*
Public Enum ALLockTypes
    lockNone = 0
    lockMAC = 1             '8
    lockComp = 2
    lockHD = 4
    lockHDFirmware = 8      '256
    lockWindows = 16        '1
    lockBIOS = 32           '16
    lockMotherboard = 64
    lockIP = 128            '32
    lockExternalIP = 256    '128
    lockFingerprint = 512
    lockMemory = 1024
    lockCPUID = 2048
    lockBaseboardID = 4096
    lockvideoID = 8192
End Enum

*/

void CAboutDlg::OnBnClickedLicBtn()
{
  if(GenerateInstallCode)
    installCode =  mfc.GetInstallationCode(registeredUser);

  UpdateData(FALSE);
  GenerateInstallCode = FALSE;
}

void CAboutDlg::OnBnClickedLicNone()
{
  lockIndex = 0;    
  mfc.PutLockType(lockType = lockNone);
  OnBnClickedLicBtn();
}

void CAboutDlg::OnBnClickedLicWindows()
{
  lockIndex = 5;
  mfc.PutLockType(lockType = lockWindows);
  OnBnClickedLicBtn();
}

void CAboutDlg::OnBnClickedLicCompName()
{
  lockIndex = 2;
  mfc.PutLockType(lockType = lockComp);
  OnBnClickedLicBtn();
}

void CAboutDlg::OnBnClickedLicDiskOld()
{
  lockIndex = 3;
  mfc.PutLockType(lockType = lockHD);
  OnBnClickedLicBtn();
}

void CAboutDlg::OnBnClickedLicMAC()
{
  lockIndex = 1;
  mfc.PutLockType(lockType = lockMAC);
  OnBnClickedLicBtn();
}

void CAboutDlg::OnBnClickedLicBIOS()
{
  lockIndex = 6;
  mfc.PutLockType(lockType = lockBIOS);
  OnBnClickedLicBtn();
}

void CAboutDlg::OnBnClickedLicIP()
{
  lockIndex = 8;
  mfc.PutLockType(lockType = lockIP);
  OnBnClickedLicBtn();
}

void CAboutDlg::OnBnClickedLicMotherboard()
{
  lockIndex = 7;
  mfc.PutLockType(lockType = lockMotherboard);
  OnBnClickedLicBtn();
}

void CAboutDlg::OnBnClickedLicDisk()
{
  lockIndex = 4;
  mfc.PutLockType(lockType = lockHDFirmware);
  OnBnClickedLicBtn();
}


void CAboutDlg::OnShowWindow(BOOL bShow, UINT nStatus)
{
  try
  {
    expiryDate        = mfc.GetExpirationDate();
    usedDays          = mfc.GetUsedDays();
  }
  catch(...)
  {
  }

  registeredUser    = mfc.RegisteredUser();
  CDialog::OnShowWindow(bShow, nStatus);
  if(bShow){
    if(!InstallationCodeFromHardware){
      // this should obtain all the information from the license itself not the hardware
      installCode =  mfc.GetInstallationCode(registeredUser);
      lockType = mfc.GetLockType();
      lockIndex = ALLockTypesToIndex(lockType);
      CButton* pBut = (CButton*) GetDlgItem(IDC_GENERATE);
      if(pBut)
        pBut->EnableWindow(FALSE);
        for(int i=0; i<lockSize; i++){ 
          if(i != lockIndex){
            int item = IDC_LIC_NONE + i;
            CButton* pBut = (CButton*) GetDlgItem(item);
            if(pBut)
              pBut->EnableWindow(FALSE);
          }
        }
    } else {
      // user has possibility to set lockCode and program gets instllation code from hardware
      lockType = mfc.GetLockType();
      lockIndex = ALLockTypesToIndex(lockType);
    }
    int item = IDC_LIC_NONE + lockIndex;
    CButton* pBut = (CButton*) GetDlgItem(item);
    if(pBut){
      pBut->SetState(TRUE);
      pBut->SetCheck(BST_CHECKED);
    }
    if(mfc.RegStatus() && mfc.LicenseStatus() != "Trial License" ){
    }
  } 
  BnSwitcher();
}

void CAboutDlg::BnSwitcher()
{
  switch(lockIndex)
  {
    /*
    From VS 2003 you can use the commeneted out lines
    but this does not work in VC 6.0
    so use the really crap version below
    */


    //  case Hardware::indNone:
  case 0:
    OnBnClickedLicNone();
    break;
    //  case Hardware::indWindows:
  case 1:
    OnBnClickedLicMAC();
    break;
    //  case Hardware::indComp:
  case 2:
    OnBnClickedLicCompName();
    break;
    //  case Hardware::indHD:
  case 3:
    OnBnClickedLicDiskOld();
    break;
    //  case Hardware::indMAC:
  case 4:
    OnBnClickedLicDisk();
    break;
  case 5:
    OnBnClickedLicWindows();
    break;
  case 6:
      // BIOS
      break;
  case 7:
    OnBnClickedLicMotherboard();
    break;
  case 8:
    OnBnClickedLicIP();
    break;
  default:
  }
}



BEGIN_MESSAGE_MAP(CAboutDlg, CDialog)
  ON_WM_SHOWWINDOW()
  ON_BN_CLICKED(IDC_LIC_NONE, OnBnClickedLicNone)
  ON_BN_CLICKED(IDC_LIC_WINDOWS, OnBnClickedLicWindows)
  ON_BN_CLICKED(IDC_LIC_CPU, OnBnClickedLicCompName)
  ON_BN_CLICKED(IDC_LIC_DISK_OLD, OnBnClickedLicDiskOld)
  ON_BN_CLICKED(IDC_LIC_NETWORK, OnBnClickedLicMAC)
  ON_BN_CLICKED(IDC_LIC_DISK, OnBnClickedLicDisk)
  ON_BN_CLICKED(IDC_GENERATE, OnBnClickedGenerate)
END_MESSAGE_MAP()



void CAboutDlg::OnBnClickedGenerate()
{
  CEdit* pEd = (CEdit*) GetDlgItem(IDC_LIC_USERNAME);
  if(pEd){
    pEd->GetWindowText(registeredUser);
    mfc.SetRegisteredUser(registeredUser);
  }
  GenerateInstallCode = TRUE;
  OnBnClickedLicBtn();
}
