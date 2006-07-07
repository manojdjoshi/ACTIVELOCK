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

const int lockSize = 9;
int ALLockTypesToIndex(ALLockTypes alt)
{
  const int locks[] = {
    lockNone ,
      lockWindows ,
      lockComp ,
      lockHD ,
      lockMAC ,
      lockBIOS ,
      lockIP ,
      lockMotherboard ,
      lockHDFirmware 
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
,GenerateInstallCode(TRUE)
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
  //installCode       = "" ;
  expiryDate        = "" ;
  usedDays          = 0  ;
  status 						= "" ;
  registeredDate 		= "" ;
  status 						= mfc.LicenseStatus();
  registeredUser    = mfc.RegisteredUser();

  //DW  doubt that the following 2 lines are correct
  // if( mfc.m_strInitAnswer != "")      licinfo           = mfc.m_strInitAnswer;
  if( mfc.m_strAcquireAnswer != "")   licinfo           = mfc.m_strAcquireAnswer;

  if(InstallationCodeFromHardware){
    explain = explainIfNeeded;
  }  else {
    // everything is in license
    registeredDate 		= mfc.GetRegisteredDate();
  }
/*
  if(mfc.RegStatus() && mfc.LicenseStatus() != "Trial License" ){
    registeredDate 		= mfc.RegisteredDate();
    installCode       = mfc.InstallationCode();
  }
  else{
    //if(lic.m_strInstallationCode != ""){
    if(!mfc.InstallationCode().IsEmpty()){
      installCode       = mfc.InstallationCode();
    } else {
      installCode       = mfc.InstallationCode();  
    }
    explain = explainIfNeeded;
  }
*/
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
  lockIndex = 1;
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
  lockIndex = 4;
  mfc.PutLockType(lockType = lockMAC);
  OnBnClickedLicBtn();
}

void CAboutDlg::OnBnClickedLicBIOS()
{
  lockIndex = 5;
  mfc.PutLockType(lockType = lockBIOS);
  OnBnClickedLicBtn();
}

void CAboutDlg::OnBnClickedLicIP()
{
  lockIndex = 6;
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
  lockIndex = 8;
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
    OnBnClickedLicWindows();
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
    OnBnClickedLicMAC();
    break;
  case 5:
    OnBnClickedLicIP();
    break;
  case 6:
    OnBnClickedLicMotherboard();
    break;
  case 7:
    OnBnClickedLicMAC();
    break;
    //  case Hardware::indHDFirmware:
  case 8:
    OnBnClickedLicDisk();
    break;
  }
}

/*
lockNone = 0,
lockWindows = 1,
lockComp = 2,
lockHD = 4,
lockMAC = 8,
lockBIOS = 16,
lockIP = 32,
lockMotherboard = 64,
lockHDFirmware = 256

*/


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
  // TODO: Add your control notification handler code here
  CEdit* pEd = (CEdit*) GetDlgItem(IDC_LIC_USERNAME);
  if(pEd){
    pEd->GetWindowText(registeredUser);
    mfc.SetRegisteredUser(registeredUser);
  }
  GenerateInstallCode = TRUE;
  OnBnClickedLicBtn();
}
