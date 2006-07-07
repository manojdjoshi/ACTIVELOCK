// fixtures.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "afxctl.h"

#include <vector>
#include <string>
#include <algorithm>
#include <strstream>
#include <iomanip>
#include "comutil.h"  

using namespace _com_util;

#import "C:\\windows\\system32\\ActiveLock3.4.dll" 
using namespace ActiveLock3; 

using namespace std;

#include "ConnectionAdvisor.h"   
#include "EventSink.h"
#include "license.h"
#include "hwid.h"
#include "comutil.h"
#include "crc32static.h"
#include "ActiveLockUtil.h"

IID iid_IActiveLockEventSink(__uuidof(__ActiveLockEventNotifier) );

string GetIt();

inline string S(char *c)
{
  return c ? c : "";
}












CString FindDirectory::Application()
{
  // the exe's directory
  CString strPath("");
  {
    char cPath[999];  cPath[0] = '\0';
    GetModuleFileName(NULL, cPath, 999);
    strPath = cPath;
  }
  int index = strPath.ReverseFind('\\'); // find last index of '\'
  if (index > 0) {
    strPath = strPath.Left(index);
  }
  return strPath;
}

CString FindDirectory::System()
{
    char windowsSystem[255];
    GetSystemDirectory(windowsSystem,255);
    return  windowsSystem;
}



// CActiveLockUtil construction

CActiveLockUtil::CActiveLockUtil()
{
    m_Email = "";
}

CActiveLockUtil::~CActiveLockUtil()
{
}

int CActiveLockUtil::CRCCheck(string& strActive, DWORD alcryptoCrc, int& crcFails, string msg)
{
  DWORD dwCrc32;
  DWORD dwAns;
  dwAns = CCrc32Static::FileCrc32Streams(strActive.c_str(), dwCrc32) ;
  if( dwAns 
    || dwCrc32 != alcryptoCrc)
  {
    crcFails += 1;
    // should I tell him or not - my preference is not
    //AfxMessageBox("ActiveLock dll has been tampered with");
    // but a bit more diplomatically
    AfxMessageBox(msg.c_str());
    return 1;
  }
  return 0;
}


int CActiveLockUtil::xCrcs(bool debug, DWORD activelock3Crc, DWORD alcryptoCrc)
{
  // has activelock.dll been tampered with
  int crcFails=0; // ok
  if(debug){				// should we check CRC of dll
    // crcFails = 0; 
  } else {
    string dllmsg[2];
    string dlls[2];
    DWORD Crcs[2];
    char windowsSystem[255];

    dllmsg[0] = "Possible wrong version of activelock3.dll";
    dllmsg[1] = "Possible wrong version of alcrypto3.dll";
    Crcs[0] = activelock3Crc;
    Crcs[1] = alcryptoCrc;
    GetSystemDirectory(windowsSystem,255);
    dlls[0] =  windowsSystem;
    dlls[0] += string("\\activelock3.4.dll");

    dlls[1] =  windowsSystem;
    dlls[1] += string("\\alcrypto3.dll");
    for(int i=0; i<2; i++){
      if(CRCCheck(dlls[i], Crcs[i], crcFails, dllmsg[i])){
        return 1;
      }
    }
  }

}

int CActiveLockUtil::FindAndCheckLicense(string& softwareName, 
                                         string& version, 
                                         bool debug, 
                                         DWORD activelock3Crc,
                                         DWORD alcryptoCrc,
                                         string KeyStorePath,
                                         string AutoRegisterKeyPath,
                                         string email,
                                         enum ALTrialTypes trialType,
                                         int trialNo
                                         )
{
  m_Email = email;
  // the following could be a proc parameter but we have enough, so try following
  // maybe test only needs to be on trialNo --- leave for now
  bool wantTrial = trialType == trialNone && trialNo == 0 ? false : true;   // seems a reasonable bit of logic

  // has activelock.dll been tampered with
  int crcFails=0; // ok
  if(debug){				// should we check CRC of dll
    // crcFails = 0; 
  } else {
    string dllmsg[2];
    string dlls[2];
    DWORD Crcs[2];
    char windowsSystem[255];

    dllmsg[0] = "Possible wrong version of activelock3.dll";
    dllmsg[1] = "Possible wrong version of alcrypto3.dll";
    Crcs[0] = activelock3Crc;
    Crcs[1] = alcryptoCrc;
    GetSystemDirectory(windowsSystem,255);
    dlls[0] =  windowsSystem;
    dlls[0] += string("\\activelock3.4.dll");

    dlls[1] =  windowsSystem;
    dlls[1] += string("\\alcrypto3.dll");
    for(int i=0; i<2; i++){
      if(CRCCheck(dlls[i], Crcs[i], crcFails, dllmsg[i])){
        return 1;
      }
    }
  }

  // the exe's directory
  CString strPath("");
  {
    char cPath[999];  cPath[0] = '\0';
    GetModuleFileName(NULL, cPath, 999);
    strPath = cPath;
  }
  int index = strPath.ReverseFind('\\'); // find last index of '\'
  if (index > 0) {
    strPath = strPath.Left(index);
  }
  string::iterator itA = find(AutoRegisterKeyPath.begin(), AutoRegisterKeyPath.end(), '\\');
  if(itA != AutoRegisterKeyPath.end()){
  } else {
    AutoRegisterKeyPath = AutoRegisterKeyPath + string("\\") + string(LPCTSTR(strPath));  // not tested
  }


  // Initialize OLE libraries
  if (!AfxOleInit()){
    AfxMessageBox("OLE initialization failed.");
    return 3;
  }
  // The application will run as an automation as long as object is active. 
  // Unlock is in ExitInstance
  //  AfxOleLockApp();


  lic.status    = "Not Registered";
  lic.hwStatus  = false;
  lic.regStatus = false;
  HRESULT hr; 

  _GlobalsPtr alGlobals;
  try 
  {
    // Create Global and then obtain ActiveLock instance and initialize its properties
    hr=alGlobals.CreateInstance("ActiveLock3.Globals", NULL); 
    AL=alGlobals->NewInstance(); 
  } catch (CException* ce){
    HandleException(ce, 1);
  } catch (...){
    string msg("1. Failed to access ActiveLock3.Globals, please send an email ");
    msg += email;
    AfxMessageBox(msg.c_str());
    return 4;
  }

  // we try to access activelock. Only to see if it is set up and works
  // we do no want to set up put_KeyStoreType here
  try 
  {
    AL->put_KeyStoreType(alsFile);  //  just a test access
  } catch (CException* ce){
    HandleException(ce, 1);
  } catch (...){
    string msg("1a. Failed to access ActiveLock3, suspect Activelock installation problem, please send an email ");
    msg += email;
    AfxMessageBox(msg.c_str());
    return 5;
  }
  // AL is set up


  try 
  {
    // Initialize ActiveLock properties
    AL->PutSoftwareName(softwareName.c_str());
    AL->PutSoftwareVersion(version.c_str()); // version as string
    AL->PutSoftwareCode(GetIt().c_str());
    AL->PutKeyStoreType(alsFile);                   // license in file
    AL->PutKeyStorePath(KeyStorePath.c_str());  // license file name
    AL->PutAutoRegisterKeyPath( AutoRegisterKeyPath.c_str());// use the auto register facility from this file

    // Connect ActiveLockEvents to my sink procedure
    BOOL Res = TRUE;
    _ActiveLockEventNotifier* alen;
    alen = AL->GetEventNotifier();
    Res = ALES.Advise(alen, iid_IActiveLockEventSink);
    // any problems here, first check that IID_IActiveLockEventSink in eventsink.cpp is correct - it changes occasionally
    ASSERT(Res == TRUE);

    // Init and Acquire return a string, so you need the following
    string strInitAns = InitBSTR();

    if(debug && 0){ //TESTING set to 1 to simulate no license
      AfxThrowUserException();
    }
    if(wantTrial){
      // Enum ALTrialHideTypes   trialSteganography = 1  trialHiddenFolder = 2  trialRegistry = 4 
      AL->PutTrialHideType(trialSteganography);
      // Enum ALTrialTypes  trialNone = 0  trialDays = 1  trialRuns = 2
      AL->PutTrialType(trialType);
      AL->PutTrialLength(trialNo); 
      // if you have problems with the trial feature you may have to uncomment the
      // following line and run the program and then reinsert the comments
      //AL->ResetTrial();  // in case things go wrong
      //AL->ResetTrial();  // in case things go wrong
    }

    string strAcquirAns = AcquireBSTR();


    // we have a license 1. a real one or 2. a trial license
    // a trial license will throw an exception on some of the following
    // a real one will work
    // could parse strAcquireAns, but I am lazy
    try 
    {
      lic.expiryDate      = S(AL->GetExpirationDate());
      lic.usedDays        = AL->GetUsedDays();
      lic.expiryDate      = S(AL->GetExpirationDate());
      lic.registeredDate  = S(AL->GetRegisteredDate());
      lic.registeredUser  = S(AL->GetRegisteredUser());
      lic.version         = S(AL->GetSoftwareVersion());
      //lic.level         = AL->GetRegisteredLevel(); // must create level first

      alLockType = AL->GetLockType();   // Activelock index
      lic.lockIndex = Hardware::lockToIndex(alLockType);       // my index

      if(lic.version != version){
        AfxThrowUserException();
      }
      lic.regStatus = true;
      lic.status = "Registered";
      lic.licinfo = "";

    } catch (CException* ce){
      // must (I hope) be a trial license
      lic.regStatus = true;
      lic.status    = "Trial License";
      lic.hwStatus  = false;     
      //int trialPeriod  = AL->get_TrialLength();
      lic.licinfo = strAcquirAns;

      ce->Delete();
    }
  } catch (CException* ce){
    HandleException(ce);
  } catch (_com_error cee){
    //printf("Error(%ld): %s\n", e.Error(), GetErrorString((enum ActiveLockErrCodeConstants)e.Error()).c_str()); 
    string msg("3. Unknown exception, please send an email ");
    msg += email;
    AfxMessageBox(msg.c_str());
    return 6;
  } catch (...){
    string msg("3. Unknown exception, please send an email ");
    msg += email;
    AfxMessageBox(msg.c_str());
    return 6;
  }

  // ok
  return 0;
}


void CActiveLockUtil::Exiting()
{
  //  AfxOleUnlockApp();
  hw.devices.clear();
  ALES.Unadvise(iid_IActiveLockEventSink); // disconnect events before activelock is destroyed
  AL = 0;    // destroy activelock
  OleUninitialize();
}

// this is the logic in about.cpp
// maybe I have overdone this seperation of activelock code from mfc
// but lets give it a go

void CActiveLockUtil::HandleException(CException* ce, int reportMkr){
  char szCause[255];
  ce->GetErrorMessage(szCause, 255);
  if(reportMkr) ce->ReportError();
  ce->Delete();
}

void CActiveLockUtil::GetInfo(
                              CString &            installCode,
                              CString &        	  registeredDate  ,
                              CString &        	  expiryDate      ,
                              long  &             usedDays        ,
                              CString &        	  registeredUser  ,
                              CString &        	  status  ,
                              CString &        	  explain,
                              CString &        	  licinfo,
                              CString          	  explainIfNeeded
                              )
{
  explain						= "";
  installCode       = "" ;
  expiryDate        = "" ;
  usedDays          = 0  ;
  status 						= "" ;
  registeredDate 		= "" ;
  status 						= CString(lic.status.c_str());
  registeredUser    = CString(lic.registeredUser.c_str());
  expiryDate        = CString(lic.expiryDate.c_str());
  usedDays          = lic.usedDays;
  licinfo          = lic.licinfo.c_str();
  if(lic.regStatus && lic.status != "Trial License" ){
    registeredDate 		= CString(lic.registeredDate.c_str());
    installCode       = CString(lic.getHardwareAndType().c_str());
  }
  else{
    //if(lic.hwLockCode != ""){
    if(lic.hwLockCode.size()){
      installCode       = CString(lic.getHardwareAndType().c_str());
    } else {
      installCode       = lic.lockCode.c_str();  
    }
    explain = explainIfNeeded;
  }
}


 void CActiveLockUtil::ReadHardware(vector<int> hwIndices)
{
 if(lic.regStatus && lic.status == "Trial License" 
    || !lic.regStatus
    ){
    // accessing installation code is time consuming - not necessary if program registered
    AL->put_LockType(alLockType); 
  try 
  {
    for(vector<int>::iterator it = hwIndices.begin(); it != hwIndices.end(); it++){
      readHW(lic.registeredUser, hw.devices, *it);
    }
  } catch (CException* ce){
    HandleException(ce, 1);
  } catch (...){
    string msg("2. Problems in accessing hardware, please send an email ");
    msg += m_Email;
    AfxMessageBox(msg.c_str());
    //return 7;
  }
  // readHW will have been setting LockType so reset
  AL->put_LockType(alLockType); 
  }
}


int CActiveLockUtil::readHW(string user, VectorDevice& devices, int index)
{
  const ALLockTypes locks[6] = {
    lockNone        ,
      lockWindows     ,
      lockComp        ,
      lockHD          ,
      lockMAC         ,
      lockHDFirmware  };


    if(index < 0 && index >= Hardware::indEnd) return -1;
    _ProductLicense* pProductLicense = 0;
    devices[index].serialNo = "";       // assume the worst
    try{
      AL->PutLockType(locks[index]); 
      string cs1 = AL->LockCode(&pProductLicense); // look at this in debugger only
      string cs = devices[index].serialNo = AL->GetInstallationCode(user.c_str(), &pProductLicense); 
    } catch (CException* ce){
      HandleException(ce);
      devices[index].serialNo = "";
    } catch(...){
      devices[index].serialNo = "";
    }
    return 0;
}
