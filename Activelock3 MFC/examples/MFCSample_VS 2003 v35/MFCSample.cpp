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


/*
* Sample MFC Application.
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

#include "stdafx.h"

#include "MFCSample.h"
#include "about.h"

#include "MainFrm.h"
#include "MFCSampleDoc.h"
#include "LeftView.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

BOOL oneApproach = TRUE;

extern IID iid_IActiveLockEventSink;

// this procedure is produced by the StringHider Utility as output to file temp.c
string GetIt(){
  typedef unsigned int integers;
  const int charsPerInt = 4;
  const int licSize     = 200;
  const int splitSize   = licSize/charsPerInt;
  const int ileft       = 0;
  int untwist[splitSize] ={
    32 , 37 , 28 , 49 , 16 , 44 , 42 , 2 , 46 , 39
      , 23 , 24 , 7 , 13 , 36 , 0 , 18 , 22 , 3 , 33
      , 21 , 19 , 29 , 40 , 35 , 9 , 41 , 31 , 38 , 5
      , 25 , 45 , 47 , 15 , 6 , 43 , 4 , 11 , 27 , 14
      , 20 , 1 , 17 , 10 , 30 , 12 , 48 , 34 , 8 , 26
  };
  integers licInt[splitSize] ={
    -1796840732 , -857840472 , -1771797410 , -1461425950 , -2105350516 , -587950492 , -1331368782 , -1328876346 , -1259966844 , 1858521774
      , -623850282 , 1851945120 , -294606716 , -596588320 , 1725870740 , -758856462 , -2071821694 , -2036045148 , 1722609250 , -526480240
      , -460663122 , 1753797252 , 1650757868 , -962550616 , -560276786 , -628698924 , 2054857312 , 1721538720 , -228423998 , -2070100778
      , 1617715440 , 1621927570 , -2105376126 , -357397792 , -2107188004 , -1902866300 , 1756520684 , -191076732 , -764634400 , -1901679004
      , -523721562 , 1892854484 , 1887736450 , -191968552 , -2105367916 , -1970902298 , -758980946 , -1534020888 , -488726334 , -2104859450
  };
  string stLic;
  stLic.reserve(licSize);
  for(int* cit = &untwist[0]; cit != &untwist[splitSize]; cit++)
  {
    char cstr[charsPerInt+1];
    int index = *cit;
    unsigned int i = licInt[index];
    i = _rotr( i, 1 );
    const unsigned int * pi = &i;
    const char * pc = reinterpret_cast<const char*>(pi);
    for(int ic=0; ic < charsPerInt; ic++){
      cstr[ic] = *pc++;
    }
    cstr[charsPerInt] = 0;
    stLic += cstr;
    const char* pTest = stLic.c_str();  // test only
  }
  if(ileft){ stLic = stLic.substr(0, (int)stLic.length() - (charsPerInt-ileft)); }
  return stLic;
  //  return     "AAAAB3NzaC1yc2EAAAABJQAAAIB8/B2KWoai2WSGTRPcgmMoczeXpd8nv0Y4r1sJ1wV3vH21q4rTpEYuBiD4HFOpkbNBSRdpBHJGWec7jUi8ISV0pM6i2KznjhCms5CEtYHRybbiYvRXleGzFsAAP817PLN3JYo3WkErT2ofR5RCkfhmx060BT8waPoqnn3AB7sZ0Q==";
    //          AAAAB3NzaC1yc2EAAAABJQAAAIB8/B2KWoai2WSGTRPcgmMoczeXpd8nv0Y4r1sJ1wV3vH21q4rTpEYuBiD4HFOpkbNBSRdpBHJGWec7jUi8ISV0pM6i2KznjhCms5CEtYHRybbiYvRXleGzFsAAP817PLN3JYo3WkErT2ofR5RCkfhmx060BT8waPoqnn3AB7sZ0Q==
}

//



/////////////////////////////////////////////////////////////////////////////
// CMFCSampleApp

BEGIN_MESSAGE_MAP(CMFCSampleApp, CWinApp)
  //{{AFX_MSG_MAP(CMFCSampleApp)
  ON_COMMAND(ID_APP_ABOUT, OnAppAbout)
  // NOTE - the ClassWizard will add and remove mapping macros here.
  //    DO NOT EDIT what you see in these blocks of generated code!
  //}}AFX_MSG_MAP
  // Standard file based document commands
  ON_COMMAND(ID_FILE_NEW, CWinApp::OnFileNew)
  ON_COMMAND(ID_FILE_OPEN, CWinApp::OnFileOpen)
  // Standard print setup command
  ON_COMMAND(ID_FILE_PRINT_SETUP, CWinApp::OnFilePrintSetup)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CMFCSampleApp construction

CMFCSampleApp::CMFCSampleApp()
{
  strApp = "MFCSample";
  strIni = "MFCSample.ini";

  // TODO: add construction code here,
  // Place all significant initialization in InitInstance

}

/////////////////////////////////////////////////////////////////////////////
// The one and only CMFCSampleApp object

CMFCSampleApp theApp;

// This identifier was generated to be statistically unique for your app.
// You may change it if you prefer to choose a specific identifier.

// {7800A88E-6BCD-4BF7-B830-C8B3002AADC7}
static const CLSID clsid =
{ 0x7800a88e, 0x6bcd, 0x4bf7, { 0xb8, 0x30, 0xc8, 0xb3, 0x0, 0x2a, 0xad, 0xc7 } };

/////////////////////////////////////////////////////////////////////////////
// CMFCSampleApp initialization

BOOL CMFCSampleApp::InitInstance()
{
  AfxEnableControlContainer();

  // Change the registry key under which our settings are stored.
  // TODO: You should modify this string to be something appropriate
  // such as the name of your company or organization.
  SetRegistryKey(_T("Local AppWizard-Generated Applications"));

  LoadStdProfileSettings();  // Load standard INI file options (including MRU)

  // Register the application's document templates.  Document templates
  //  serve as the connection between documents, frame windows and views.

  CSingleDocTemplate* pDocTemplate;
  pDocTemplate = new CSingleDocTemplate(
    IDR_MAINFRAME,
    RUNTIME_CLASS(CMFCSampleDoc),
    RUNTIME_CLASS(CMainFrame),       // main SDI frame window
    RUNTIME_CLASS(CLeftView));
  AddDocTemplate(pDocTemplate);

  // Connect the COleTemplateServer to the document template.
  //  The COleTemplateServer creates new documents on behalf
  //  of requesting OLE containers by using information
  //  specified in the document template.
  m_server.ConnectTemplate(clsid, pDocTemplate, TRUE);
  // Note: SDI applications register server objects only if /Embedding
  //   or /Automation is present on the command line.

  // Parse command line for standard shell commands, DDE, file open
  CCommandLineInfo cmdInfo;
  ParseCommandLine(cmdInfo);

  // Check to see if launched as OLE server
  if (cmdInfo.m_bRunEmbedded || cmdInfo.m_bRunAutomated)
  {
    // Register all OLE server (factories) as running.  This enables the
    //  OLE libraries to create objects from other applications.
    COleTemplateServer::RegisterAll();

    // Application was run with /Embedding or /Automation.  Don't show the
    //  main window in this case.
    return TRUE;
  }

  // When a server application is launched stand-alone, it is a good idea
  //  to update the system registry in case it has been damaged.
  m_server.UpdateRegistry(OAT_DISPATCH_OBJECT);
  COleObjectFactory::UpdateRegistryAll();

  // Dispatch commands specified on the command line
  if (!ProcessShellCommand(cmdInfo))
    return FALSE;

  // The one and only window has been initialized, so show and update it.
  m_pMainWnd->ShowWindow(SW_SHOW);
  m_pMainWnd->UpdateWindow();


  // ActiveLock Section
  BOOL debug = 1; // best supplied in debug mode as crc values are usually wrong
  if( !CrcsEtc(debug))
  {
    return FALSE; // stop dead
  }
  try
  {
    char chLockType[20];
    char UserName[200];
    GetPrivateProfileString (LPCTSTR(strApp), "User", "", UserName, 200, LPCTSTR(strIni)); 
    GetPrivateProfileString (LPCTSTR(strApp), "lockType", "", chLockType, 200, LPCTSTR(strIni)); 

    if(oneApproach){
      m_pActiveLockMfc = new CActiveLockMFC();
      m_pActiveLockMfc->SetRegisteredUser(CString(UserName));
      lockType = m_pActiveLockMfc->LockTypeFromInt(atoi(chLockType));
      m_pActiveLockMfc->SetLockType(lockType);

      CheckLicenseWrapperLevel();
    } else {
      // the other approach
      CString KeyStorePath				            = _T("TestApp.lic");  // license file name
      CString AutoRegisterKeyPath             = _T("c:\\TestApp.all");// use the auto register facility from this file
      CString softwareName				            = _T("TestApp");
      CString version						              = _T("1.0"); // license does not contain version number !!!!
      enum LicStoreType licStoreType          = alsFile;
      enum ALLicType    LicType               = allicTimeLocked;
      enum ALLockTypes  LockType              = lockType;
      enum ALTrialTypes trialType             = trialRuns;
      int  trialNo                            = 31;
      enum ALTrialHideTypes   HideType        = trialSteganography;
      CString  RegisteredLevel                = "Level 3";


      m_pActiveLockMfc = new   CActiveLockMFC( softwareName, 
        version,
        licStoreType,
        LicType,
        LockType,
        KeyStorePath,
        AutoRegisterKeyPath,
        trialType,
        trialNo,
        HideType,
        RegisteredLevel
        );
      m_pActiveLockMfc->SetRegisteredUser(CString(UserName));
      lockType = m_pActiveLockMfc->LockTypeFromInt(atoi(chLockType));
      m_pActiveLockMfc->SetLockType(lockType);
      CheckLicenseMFCLevel();
    }
  }
  catch(int e)
  {
    switch(e)
    {
    case 1:
      break;
    case 2:
      break;
    default:
      break;
    }
      
  }

  return TRUE;
}


int CMFCSampleApp::ExitInstance()
{
  if(oneApproach){
    m_pActiveLockMfc->DisconnectEvent();
    m_pActiveLockMfc->DestroyActiveLock();
  } else {
    m_pActiveLockMfc->DestroyAll();
  }
  delete m_pActiveLockMfc;
  return CWinApp::ExitInstance();
}

// App command to run the dialog
void CMFCSampleApp::OnAppAbout()
{
  CAboutDlg aboutDlg(*m_pActiveLockMfc);
  aboutDlg.DoModal();
  m_pActiveLockMfc->SetRegisteredUser(aboutDlg.registeredUser);
  m_pActiveLockMfc->SetLockType(aboutDlg.lockType);

  BOOL bAns;
  bAns =WritePrivateProfileString(LPCTSTR(strApp), "User", LPCTSTR(m_pActiveLockMfc->RegisteredUser()), LPCTSTR(strIni));
  int i = m_pActiveLockMfc->LockTypeToInt(m_pActiveLockMfc->LockType());

  enum ALLockTypes tmp = m_pActiveLockMfc->LockType();
 i = m_pActiveLockMfc->LockTypeToInt(tmp);

  CString cs;
  cs.Format(_T("%i"),i); 
  bAns =WritePrivateProfileString(LPCTSTR(strApp), "lockType", LPCTSTR(cs), LPCTSTR(strIni));
}

CString CMFCSampleApp::ProgramPath() 
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


BOOL CMFCSampleApp::CheckLicenseWrapperLevel() 
{ 
  BOOL bLimitedLicense;   // maybe this should be a member variable

  m_pActiveLockMfc->Create(); 

  m_pActiveLockMfc->PutSoftwareVersion(CString("1.0")); 
  m_pActiveLockMfc->PutSoftwareName(CString("TestApp")); 
  m_pActiveLockMfc->PutKeyStoreType(alsFile); // 1 = File 
  m_pActiveLockMfc->PutLockType(lockType);
  m_pActiveLockMfc->PutSoftwareCode(CString(GetIt().c_str())); 

  // license in application directory
  //CString strTmpLic = ProgramPath() + _T("//testapp.lic"); 

  // license in windows directory
  CString strTmpLic = _T("testapp.lic"); 

  m_pActiveLockMfc->PutKeyStorePath(strTmpLic); 

  //CString strTmpAll = ProgramPath() + _T("\\testapp.all"); 
  CString strTmpAll = _T("c:\\testapp.all"); 
  m_pActiveLockMfc->PutAutoRegisterKeyPath(strTmpAll); 

  // trialDays is more sensible but with runs you can see it change each activation, better than waiting 24 hours
  m_pActiveLockMfc->PutTrialType(trialRuns);
  m_pActiveLockMfc->PutTrialLength(31);
  m_pActiveLockMfc->PutTrialHideType(trialSteganography);

  m_pActiveLockMfc->Initialize(); 

  m_pActiveLockMfc->CheckLicense(); 

  //   if(!m_pActiveLockMfc->RegStatus()  ||  m_pActiveLockMfc->IsTrialLicense()) 
  //      ActiveLockRegister(); 

  if(m_pActiveLockMfc->RegStatus()) 
    bLimitedLicense = m_pActiveLockMfc->IsLimitedLicense(); 

  return (m_pActiveLockMfc->RegStatus() &&  !m_pActiveLockMfc->IsTrialLicense()); 
} 


BOOL CMFCSampleApp::CheckLicenseMFCLevel() 
{ 
  BOOL bLimitedLicense;

  //FindProgramPath(); 

  m_pActiveLockMfc->CheckLicenseMFC(); 

  //   if(!m_pActiveLockMfc->RegStatus()  ||  m_pActiveLockMfc->IsTrialLicense()) 
  //      ActiveLockRegister(); 

  if(m_pActiveLockMfc->RegStatus()) 
    bLimitedLicense = m_pActiveLockMfc->IsLimitedLicense(); 

  return (m_pActiveLockMfc->RegStatus() &&  !m_pActiveLockMfc->IsTrialLicense()); 
} 


BOOL  CMFCSampleApp::CrcsEtc(int debug)
{
  if(debug) return TRUE;
  VectorNameAndCrc vNameCrc;
  DWORD activelock3Crc			= 603433813;
  DWORD alcryptoCrc         = 2581601961;

  char windowsSystem[255];
  {
    GetSystemDirectory(windowsSystem,255);
    struct CNameAndCrc cnac;
    cnac.crc = activelock3Crc;
    cnac.msg = "Possible wrong version of activelock3.dll";
    cnac.name =  windowsSystem;
    cnac.name += CString("\\activelock3.4.dll");
    vNameCrc.push_back(cnac);
  }

  {
    struct CNameAndCrc cnac;
    cnac.crc = alcryptoCrc;
    cnac.msg = "Possible wrong version of alcrypto3.dll";
    cnac.name =  windowsSystem;
    cnac.name += CString("\\alcrypto3.dll");
    vNameCrc.push_back(cnac);
  }
  return util.CheckAllCrcs(debug, vNameCrc);
}

