#include "stdafx.h" 
#include <vector>
#include <string>
using namespace std;

// in the distant past - com_util was faulty - this was the fix
// considering this was many years ago this must have been fixed but keep in case problems
//#include "comutilfix.h"  
//using namespace _com_util_fix; 

#include <comutil.h>
using namespace _com_util;


#import "C:\\Windows\\SysWOW64\\ActiveLock3.6.dll"   // ******* this maybe somewhere else - check *************
using namespace ActiveLock3; 

#include "ConnectionAdvisor.h"   
#include "EventSink.h" 
#include "ActiveLockmfc.h" 

IID iid_IActiveLockEventSink(__uuidof(__ActiveLockEventNotifier) );

void ActiveLockHandleException(_com_error ce, BOOL bShow = FALSE);
string GetIt();

void ActiveLockHandleException(_com_error ce, BOOL bShow) 
{ 
  _bstr_t bDesc = ce.Description();
  _bstr_t bContext = ce.HelpContext();
  _bstr_t bFile = ce.HelpFile();
  CString  Desc( ConvertBSTRToString(bDesc)); 
  CString  Context( ConvertBSTRToString(bContext)); 
  CString  File( ConvertBSTRToString(bFile)); 
  CString Cause = ce.ErrorMessage();
  CString mes = "Descr. " + Desc + " Context: " + Context + " File: " + File + " Cause: " + Cause;

  if(bShow){ 
    AfxMessageBox(LPCSTR(mes));
  }
} 

CString& CActiveLockWrap::BSTRToCString(BSTR& bs) 
{ 
  char *pc = ConvertBSTRToString(bs); 
  CString* pstr = new CString(pc); 
  delete pc; 
  SysFreeString(bs); 
  return *pstr; 
} 

const CString CActiveLockWrap::GetInstallationCode(const CString& strUsn) const
{ 
  CString strInstCode = _T(""); 
  try 
  { 
    _ProductLicense* pProductLicense = 0;
    strInstCode = S(m_ActiveLock->GetInstallationCode(LPCTSTR(strUsn), &pProductLicense)); 
  } 
  catch (_com_error ce) 
  { 
    ActiveLockHandleException(ce,TRUE); 
    return _T("Not Valid"); 
  } 
  return strInstCode; 
} 

// next 2 procs bury details of providing an "out" string to acquire.

const CString CActiveLockWrap::InitBSTR() 
{ 
  BSTR autoLicString = SysAllocStringByteLen( "", 10000); // license could be large and this is freed below
  m_ActiveLock->Init(&autoLicString); 
  return BSTRToCString(autoLicString); 
} 


const CString CActiveLockWrap::AcquireBSTR(        
        CString & strRemainingTrialDays,
        CString & strRemainingTrialRuns,
        CString & strTrialLength,
        CString & strUsedDays,
        CString & strExpirationDate,
        CString & strRegisteredUser,
        CString & strRegisteredLevel,
        CString & strLicenseClass,
        CString & strMaxCount,
        CString & strLicenseFileType,
        CString & strLicenseType,
        CString & strUsedLockType )
{ 
  BSTR bstrAcquireAns = SysAllocStringByteLen( "", 100); 
  BSTR bstrRemainingTrialDays = SysAllocStringByteLen( "", 100); 
  BSTR bstrRemainingTrialRuns = SysAllocStringByteLen( "", 100); 
  BSTR bstrTrialLength = SysAllocStringByteLen( "", 100); 
  BSTR bstrUsedDays = SysAllocStringByteLen( "", 100); 
  BSTR bstrExpirationDate = SysAllocStringByteLen( "", 100); 
  BSTR bstrRegisteredUser = SysAllocStringByteLen( "", 100); 
  BSTR bstrRegisteredLevel = SysAllocStringByteLen( "", 100); 
  BSTR bstrLicenseClass = SysAllocStringByteLen( "", 100); 
  BSTR bstrMaxCount = SysAllocStringByteLen( "", 100); 
  BSTR bstrLicenseFileType = SysAllocStringByteLen( "", 100); 
  BSTR bstrLicenseType = SysAllocStringByteLen( "", 100); 
  BSTR bstrUsedLockType = SysAllocStringByteLen( "", 100); 
  
  //AfxMessageBox(_T("Starting Acquire...")); 
     // Check to see if we have a valid license 
  m_ActiveLock->Acquire(&bstrAcquireAns,
        &bstrRemainingTrialDays,
        &bstrRemainingTrialRuns,
        &bstrTrialLength,
        &bstrUsedDays,
        &bstrExpirationDate,
        &bstrRegisteredUser,
        &bstrRegisteredLevel,
        &bstrLicenseClass,
        &bstrMaxCount,
        &bstrLicenseFileType,
        &bstrLicenseType,
        &bstrUsedLockType );
  //AfxMessageBox(_T("Finished Acquire...")); 

  strRemainingTrialDays = BSTRToCString(bstrRemainingTrialDays); 
  strRemainingTrialRuns = BSTRToCString(bstrRemainingTrialRuns); 
  strTrialLength = BSTRToCString(bstrTrialLength); 
  strUsedDays = BSTRToCString(bstrUsedDays); 
  strExpirationDate = BSTRToCString(bstrExpirationDate); 
  strRegisteredUser = BSTRToCString(bstrRegisteredUser); 
  strRegisteredLevel = BSTRToCString(bstrRegisteredLevel); 
  strLicenseClass = BSTRToCString(bstrLicenseClass); 
  strMaxCount = BSTRToCString(bstrMaxCount); 
  strLicenseFileType = BSTRToCString(bstrLicenseFileType); 
  strLicenseType = BSTRToCString(bstrLicenseType); 
  strUsedLockType = BSTRToCString(bstrUsedLockType); 

  return BSTRToCString(bstrAcquireAns);; 
} 

void CActiveLockWrap::DisconnectEvent() 
{ 
  m_ActiveLockEventSink.Unadvise(iid_IActiveLockEventSink); // disconnect events before activelock is destroyed 
} 


void CActiveLockWrap::DestroyActiveLock() 
{ 
  m_ActiveLock = 0;                                 // destroy activelock - this simple line is correct and also very complex 
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
enum ALLockTypes CActiveLockWrapEx::LockTypeFromInt(int i)      const 
  {
    switch(i){
      case 0:
        return lockNone;
      case 1:
        return lockMAC;
      case 2:
        return lockComp;
      case 4:
        return lockHD;
      case 8:
        return lockHDFirmware;
      case 16:
        return lockWindows;
      case 32:
        return lockBIOS;
      case 64:
        return lockMotherboard;
      case 128:
        return lockIP;
      case 256:
        return lockExternalIP;
        break;
      case 512:
        return lockFingerprint;
      case 1024:
        return lockMemory;
        break;
      case 2048:
        return lockCPUID;
      case 4096:
        return lockBaseboardID;
      case 8192:
        return lockvideoID;
      default:
        break;
    }
    return lockNone;
  }





CActiveLockMFC::CActiveLockMFC() 
{ 
  Init();
} 

CActiveLockMFC::~CActiveLockMFC() 
{ 
} 


void CActiveLockMFC::Init() 
{ 
  m_bRegStatus          = FALSE; 
  m_bLimited            = FALSE; 
  m_bTrial               = FALSE; 
  m_strUsedDays          = ""; 

  m_strExpirationDate.Empty(); 
  m_strRegisteredDate.Empty(); 
  m_strRegisteredUser.Empty(); 
  m_strRegisteredLevel.Empty(); 
  m_strLicenseStatus   = _T("Not registered"); 
} 

void CActiveLockMFC::Create() 
{ 
  _GlobalsPtr alGlobals; 
  HRESULT hr;
  // Initialize OLE libraries
  if (!AfxOleInit()){
    AfxMessageBox("OLE initialization failed.");
    return;
  }
  try 
  {
    hr = alGlobals.CreateInstance("ActiveLock3.Globals", NULL); 
    m_ActiveLock = alGlobals->NewInstance(); 
  } 
  catch (_com_error ce) 
  { 
    ActiveLockHandleException(ce, TRUE); 
    throw 1;
  } 
} 


BOOL CActiveLockMFC::Initialize() 
{ 
  m_strInitAnswer.Empty(); 
  CString strTmp; 

  try 
  { 
    _ActiveLockEventNotifier* alen = GetEventNotifier(); 
    BOOL bRes = m_ActiveLockEventSink.Advise(alen, iid_IActiveLockEventSink); 
    if(!bRes) 
      AfxMessageBox(_T("m_ActiveLockEventSink.Advise(...) == FALSE")); 

    // Init and Acquire return a CString, so you need the following 

    m_strInitAnswer = InitBSTR(); 

    if(0) 
    { //DEBUG TESTING set to 1 to simulate no license 
      AfxThrowUserException(); 
    } 
  } 
  catch (_com_error ce) 
  { 
    ActiveLockHandleException(ce); 
    throw 1;
    // return FALSE; 
  } 
  return TRUE; 
} 


BOOL CActiveLockMFC::CheckLicense() 
{ 
  m_bRegStatus = FALSE;                                           // assume the worst - no valid license
  m_bTrial = FALSE; 
  m_strAcquireAnswer.Empty(); 

  try 
  { 

    // if you have problems with the trial feature you may have to uncomment the 
    // following line and run the program and then reinsert the comments 
    //ResetTrial();  // in case things go wrong 
    //ResetTrial();  // a post said I should do it twice-easy enough-can not do any harm 

    CString strAns = AcquireBSTR(
        m_strRemainingTrialDays,
        m_strRemainingTrialRuns,
        m_strTrialLength,
        m_strUsedDays,             //
        m_strExpirationDate,      //
        m_strRegisteredUser,
        m_strRegisteredLevel,     //
        m_strLicenseClass,
        m_strMaxCount,
        m_strLicenseFileType,
        m_strLicenseType,
        m_strUsedLockType ); 

    // license fails - exception thrown - caught in catch below
    // license ok gets to here with strAns as empty string
    // license is a valid trial license gets to here with strAns blah blah trial license 16 days of 31 days left 

    m_strAcquireAnswer.Format(_T("%s"),strAns); 
    if(strAns != "")                                        // trial license
    {
      // a trial license 
      m_bRegStatus         = TRUE; 
      m_bTrial             = TRUE; 
      m_strLicenseStatus   = _T("Trial License"); 
    }
    else
    {                                                                               // successfully registered
      try 
      {                                                                             // get all info
        m_bRegStatus         = TRUE; 
        m_strLicenseStatus   = _T("Registered"); 
        m_strRegisteredDate  = GetRegisteredDate(); 
        m_lLockType          = m_ActiveLock->GetLockType();                         // Activelock index 
        
        if(m_strRegisteredLevel.Left(7).CompareNoCase(_T("Limited")) == 0)          // I do not use RegisterLevel - so following is just a play
          m_bLimited = TRUE; 
        else 
          m_bLimited = FALSE; 
      } 
      catch (_com_error ce) 
      { 
      } 

    } 
  } 
  catch (_com_error ce) 
  { 
    // default is set to faulty so not much to do - print some info to help sort this out
    ActiveLockHandleException(ce); 
    throw 1;
  } 
  return TRUE; 
} 


