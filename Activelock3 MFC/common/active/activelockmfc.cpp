#include "stdafx.h" 
#include <string>

#include "comutilfix.h"  
using namespace _com_util_fix; 


#import "C:\\windows\\system32\\ActiveLock3.4.dll" 
using namespace ActiveLock3; 

using namespace std;

#include "ConnectionAdvisor.h"   
#include "EventSink.h" 
#include "ActiveLockmfc.h" 

IID iid_IActiveLockEventSink(__uuidof(__ActiveLockEventNotifier) );


void ActiveLockHandleException(_com_error ce, BOOL bShow = FALSE);
string GetIt();


void ActiveLockHandleException(_com_error ce, BOOL bShow) 
{ 
  const TCHAR* pCause = ce.ErrorMessage(); 
  if(bShow){ 
    AfxMessageBox(pCause);
  }
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
  catch (...) 
  { 
    AfxMessageBox(_T("Unknown exception during retrive of installation code, please report authors about this problem")); 
    return _T("Not Valid"); 
  } 

  return strInstCode; 
} 


const CString CActiveLockWrap::LockCode() const
{ 

  CString strInstCode = _T(""); 
  try 
  { 
    _ProductLicense* pProductLicense = 0;
    strInstCode = S(m_ActiveLock->LockCode(&pProductLicense)); 
  } 
  catch (_com_error ce) 
  { 
    ActiveLockHandleException(ce,TRUE); 
    return _T("Not Valid"); 
  } 
  catch (...) 
  { 
    AfxMessageBox(_T("Unknown exception during retrive of installation code, please report authors about this problem")); 
    return _T("Not Valid"); 
  } 

  return strInstCode; 
} 

// next 2 procs bury details of providing an "out" string to acquire.
const CString CActiveLockWrap::InitBSTR() 
{ 
  _ActiveLockEventNotifier* alen = GetEventNotifier(); 
  BOOL bRes = m_ActiveLockEventSink.Advise(alen, iid_IActiveLockEventSink); 
  if(!bRes) 
    AfxMessageBox(_T("m_ActiveLockEventSink.Advise(...) == FALSE")); 

  BSTR autoLicString = SysAllocStringByteLen( "", 1000); // license could be large 
  m_ActiveLock->Init(&autoLicString); 
  char *pci = _com_util_fix::ConvertBSTRToString(autoLicString); 
  CString strInitAns(pci); 
  delete pci; 
  SysFreeString(autoLicString); 

  return strInitAns; 
} 


const CString CActiveLockWrap::AcquireBSTR() 
{ 
  BSTR bstrAcquireAns = SysAllocStringByteLen( "", 100); 
  //AfxMessageBox(_T("Starting Acquire...")); 
  m_ActiveLock->Acquire(&bstrAcquireAns); // Check to see if we have a valid license 
  //AfxMessageBox(_T("Finished Acquire...")); 
  char *pc = _com_util_fix::ConvertBSTRToString(bstrAcquireAns); 
  CString strAcquireAns(pc); 
  delete pc; 
  SysFreeString(bstrAcquireAns); 
  return strAcquireAns; 
} 

void CActiveLockWrap::DisconnectEvent() 
{ 
  m_ActiveLockEventSink.Unadvise(iid_IActiveLockEventSink); // disconnect events before activelock is destroyed 
} 


void CActiveLockWrap::DestroyActiveLock() 
{ 
  m_ActiveLock = 0;    // destroy activelock - this simple line is correct and also very complex 
} 





enum ALLockTypes CActiveLockWrapEx::LockTypeFromInt(int i)      const 
  {
    switch(i){
      case 0:
        return lockNone;
        break;
      case 1:
        return lockWindows;
        break;
      case 2:
        return lockComp;
        break;
      case 4:
        return lockHD;
        break;
      case 8:
        return lockMAC;
        break;
      case 16:
        return lockBIOS;
        break;
      case 32:
        return lockIP;
        break;
      case 64:
        return lockMotherboard;
        break;
      case 256:
        return lockHDFirmware;
        break;
    }
    return lockNone;
  }







CActiveLockMFC::CActiveLockMFC() 
{ 
  Init();
} 

CActiveLockMFC::CActiveLockMFC(CString&                 softwareName, 
                               CString&                 version, 
                               enum LicStoreType        licStoreType,
                               enum ALLicType           LicType,
                               enum ALLockTypes         LockType,
                               CString                  KeyStorePath,
                               CString                  AutoRegisterKeyPath,
                               enum ALTrialTypes        trialType,
                               int                      trialNo,
                               enum ALTrialHideTypes    HideType,
                               CString                  RegisteredLevel
                               )
{
  Init();
  m_strSoftwareName             = softwareName; 
  m_strSoftwareVersion          = version; 
  m_licStoreType                = licStoreType;
  m_strKeyStorePath             = KeyStorePath; 
  m_strAutoRegisterKeyPath      = AutoRegisterKeyPath; 
  m_strRegisteredLevel          = RegisteredLevel; 

  m_lLockType                   = LockType; 
  m_lLicType                    = LicType; 
  m_lTrialType                  = trialType;; 
  m_lTrialLength                = trialNo; 
  m_lTrialHideType              = HideType; 
}


CActiveLockMFC::~CActiveLockMFC() 
{ 
} 


void CActiveLockMFC::Init() 
{ 
  m_bCreated            = FALSE; 
  m_bInitialized        = FALSE; 
  m_bRegStatus          = FALSE; 
  m_bLimited            = FALSE; 
} 


void CActiveLockMFC::ResetCollectionData(void) 
{ 
  m_bTrial               = FALSE; 
  m_lUsedDays            = 0; 
  m_strExpirationDate.Empty(); 
  m_strRegisteredDate.Empty(); 
  m_strRegisteredUser.Empty(); 
  m_strRegisteredLevel.Empty(); 
  m_strLicenseStatus   = _T("Not registered"); 
} 


BOOL CActiveLockMFC::CollectLicenseData() 
{ 
  //DW? ResetCollectionData(); 

  // we have a license 1. a real one or 2. a trial license 
  // a trial license will throw an exception on some of the following 
  // a real one will work 
  // could parse strAcquireAns, but I am lazy 
  try 
  { 
    m_lUsedDays            = GetUsedDays();
    m_strRegisteredUser    = GetRegisteredUser(); 
    m_strRegisteredDate    = GetRegisteredDate(); 
    m_strRegisteredLevel   = GetRegisteredLevel(); 
    m_lLockType            = m_ActiveLock->GetLockType();                       // Activelock index 
    if(m_strRegisteredLevel.Left(7).CompareNoCase(_T("Limited")) == 0) 
      m_bLimited = TRUE; 
    else 
      m_bLimited = FALSE; 
    CString strVersion   = GetSoftwareVersion(); 
		// I do not believe version is in the license - so next test is rubbish. I could be wrong
    /* remove
    if(m_strSoftwareVersion != strVersion) 
    { 
      AfxMessageBox(_T("Wrong software version")); 
      AfxThrowUserException(); 
      //         return FALSE; 
    } 
    */
    // only a proper license will get here
    m_lLockType = GetUsedLockType();    // lockType used in the license 
    // The above could be a combination of codes - this case is not handled here - future job
    // in fact a combination can not be a legal enum type-my fault should report to forum
    PutLockType(m_lLockType);                         // make them all agree

    m_bRegStatus         = TRUE; 
    m_strLicenseStatus   = _T("Registered"); 
  } 
  catch (_com_error ce) 
  { 
    // must (I hope) be a trial license 
    m_bRegStatus         = TRUE; 
    m_bTrial               = TRUE; 
    m_strLicenseStatus   = _T("Trial License"); 
  } 

  return TRUE; 
} 


BOOL CActiveLockMFC::Create() 
{ 
  _GlobalsPtr alGlobals; 
  HRESULT hr;
  if(!m_bCreated) 
  { 
    // Initialize OLE libraries
    if (!AfxOleInit()){
      AfxMessageBox("OLE initialization failed.");
      return 3;
    }
    try 
    {
      hr = alGlobals.CreateInstance("ActiveLock3.Globals", NULL); 
      m_ActiveLock = alGlobals->NewInstance(); 

      m_bCreated = TRUE; 
    } 
    catch (_com_error ce) 
    { 
      ActiveLockHandleException(ce, TRUE); 
      throw 1;
      return FALSE; 
    } 
    catch (...) 
    { 
      AfxMessageBox(_T("Failed to create protection instance, please report authors about this problem")); 
      throw 2;
      return FALSE; 
    } 
  } 

  return TRUE; 
} 


BOOL CActiveLockMFC::Initialize() 
{ 
  m_strInitAnswer.Empty(); 
  CString strTmp; 

  try 
  { 
    if(!m_bInitialized) 
    { 
      // Init and Acquire return a CString, so you need the following 

      strTmp = InitBSTR(); 
      m_strInitAnswer.Format(_T("Initial answer: %s"),strTmp); 

      if(0) 
      { //DEBUG TESTING set to 1 to simulate no license 
        AfxThrowUserException(); 
      } 

      m_bInitialized = TRUE; 
    } 
  } 
  catch (_com_error ce) 
  { 
    ActiveLockHandleException(ce); 
    throw 1;
    // return FALSE; 
  } 
  catch (...) 
  { 
    AfxMessageBox(_T("Unknown exception during initialization of protection instance, please report authors about this problem")); 
    throw 2;
    return FALSE; 
  } 

  return TRUE; 
} 


BOOL CActiveLockMFC::CheckLicense() 
{ 
  m_bRegStatus = FALSE; 
  m_bTrial = FALSE; 
  m_strAcquireAnswer.Empty(); 

  try 
  { 

    // if you have problems with the trial feature you may have to uncomment the 
    // following line and run the program and then reinsert the comments 
    //ResetTrial();  // in case things go wrong 
    //ResetTrial();  // a post said I should do it twice-easy enough-can not do any harm 

    CString strTmp = AcquireBSTR(); 
    m_strAcquireAnswer.Format(_T("%s"),strTmp); 
    CollectLicenseData(); 
  } 
  catch (_com_error ce) 
  { 
    ActiveLockHandleException(ce); 
    throw 1;
    //      return FALSE; 
  } 
  catch (...) 
  { 
    AfxMessageBox(_T("Unknown exception during put of acquire/collecting license data, please report authors about this problem")); 
    throw 2;
    return FALSE; 
  } 

  return TRUE; 
} 




void CActiveLockMFC::PutSoftwareData() 
{ 
  // Initialize ActiveLock properties 
  PutSoftwareName(m_strSoftwareName); 
  PutSoftwareVersion(m_strSoftwareVersion); 
  PutLockType(m_lLockType); 
  PutSoftwareCode(CString(GetIt().c_str())); 
  PutKeyStoreType(m_licStoreType);  
  PutTrialType(m_lTrialType);
  PutTrialHideType(m_lTrialHideType);
  PutTrialLength(m_lTrialLength);

  if(!m_strKeyStorePath.IsEmpty()) 
    PutKeyStorePath(m_strKeyStorePath); 
  if(!m_strAutoRegisterKeyPath.IsEmpty()) 
    PutAutoRegisterKeyPath(m_strAutoRegisterKeyPath); 
} 

BOOL CActiveLockMFC::CheckLicenseMFC() 
{ 
  Create(); 
  PutSoftwareData();

  Initialize(); 
  CheckLicense(); 
  //   if(!m_pActiveLockMfc->RegStatus()  ||  m_pActiveLockMfc->IsTrialLicense()) 
  //      ActiveLockRegister(); 
  //DWif(RegStatus()) 
  //DW  bLimitedLicense = IsLimitedLicense(); 
  return (RegStatus() &&  !IsTrialLicense()); 
  //return TRUE; 
} 

void CActiveLockMFC::DestroyAll() 
{ 
  DisconnectEvent();
  DestroyActiveLock();

  m_bCreated       = FALSE; 
  m_bInitialized   = FALSE; 
} 

