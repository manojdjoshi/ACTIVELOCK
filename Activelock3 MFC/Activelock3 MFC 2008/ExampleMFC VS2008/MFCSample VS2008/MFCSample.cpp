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
#include <direct.h>
#include <shfolder.h>
#include <shlwapi.h>

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

extern IID iid_IActiveLockEventSink;

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



CString CMFCSampleApp::CurrentDirectory() 
{ 
  char path[_MAX_PATH];
  _getcwd( path, _MAX_PATH );
  CString* spCurr = new CString(path);
  return *spCurr;
}

/////////////////////////////////////////////////////////////////////////////
// CMFCSampleApp initialization

BOOL CMFCSampleApp::InitInstance()
{
  AfxEnableControlContainer();

  TCHAR szPath[MAX_PATH];
  if(SUCCEEDED( SHGetFolderPath(NULL, 
    CSIDL_COMMON_APPDATA|CSIDL_FLAG_CREATE, 
    NULL, 
    0, 
    szPath))) 
  {
    PathAppend(szPath, TEXT("MFCSample"));
    strDataFolder = szPath;
    // strDataFolder =  CString("C:\\Dave\\Lic");     // *********************************simplifed life at beginning
    _chdir((LPCTSTR)strDataFolder);
    CString test = CurrentDirectory();                // *********************************check it is set, had lots of trouble, set up in installation package
  }

  // Change the registry key under which our settings are stored.
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


  // A c t i v e L o c k   S e c t i o n

  BOOL debug = 1; // best supplied in debug mode as crc values are usually wrong 
  // but you should set this to 0 and debug CrcsEtc to contain correc values of CRCs - the values are certainly wrong
  if( !CrcsEtc(debug))
  {
    return FALSE; // stop dead
  }
  // nearly all errors from Activelock are reported by throwing exceptions
  try
  {
    // username and locktype are stored in ini file
    // locktype stored as integer and then converted to enum
    char chLockType[20];
    char UserName[200];
    GetPrivateProfileString (LPCTSTR(strApp), "User", "", UserName, 200, LPCTSTR(strIni)); 
    GetPrivateProfileString (LPCTSTR(strApp), "lockType", "", chLockType, 200, LPCTSTR(strIni)); 
    lockType = m_pActiveLockMfc->LockTypeFromInt(atoi(chLockType));

    m_pActiveLockMfc = new CActiveLockMFC();                        // wrapper class on Activelock dll
    m_pActiveLockMfc->SetRegisteredUser(CString(UserName));
    m_pActiveLockMfc->SetLockType(lockType);

    m_pActiveLockMfc->Create();                                     // 1. First common fail point: links to dll 

    m_pActiveLockMfc->PutSoftwarePassword("MyPassWord");
    m_pActiveLockMfc->PutSoftwareVersion(CString("3.6")); 
    m_pActiveLockMfc->PutSoftwareName(CString("MFCSample")); 
    m_pActiveLockMfc->PutKeyStoreType(alsFile); // 1 = File 
    m_pActiveLockMfc->PutLockType(lockType);

    // V Code from Alugen
    m_pActiveLockMfc->PutSoftwareCode( "RSA1024BgIAAAAkAABSU0ExAAQAAAEAAQDluCjG8CD+nHXnmhm56CUA2MP8/KPwOMXsWZhCnnniM4xALB2LJpkASP0ojz72kHRVY47TtgMzq1/53fdAmncFEZ6O20APo/5/mY5nsWucwmi70vhE+yrpHZ0vohR5Ux3fEnmJ/vMqUhFyGhgeiBhFD25C3NAZEbf87XDVfk6QxA==");
                                         
    // license in Current Directory 
    CString strTmpLic = CurrentDirectory() + _T("\\MFCSample3.6.lic"); 
    m_pActiveLockMfc->PutKeyStorePath(strTmpLic); 

    // liberation file in Current Directory
    CString strTmpAll = CurrentDirectory() + _T("\\MFCSample3.6.all"); 
    m_pActiveLockMfc->PutAutoRegisterKeyPath(strTmpAll); 

    // trialDays is more sensible but with runs you can see it change each activation, better than waiting 24 hours
    m_pActiveLockMfc->SetTrialType(trialRuns);
    //m_pActiveLockMfc->SetTrialType(trialDays);
    m_pActiveLockMfc->PutTrialLength(31);
    m_pActiveLockMfc->PutTrialHideType(trialSteganography);

    // All activelock Parameters set - now the real work begins
    m_pActiveLockMfc->Initialize();                                  // 2. Second common fail point: calls Activelock Init procedure

    m_pActiveLockMfc->CheckLicense();                                // 3. Third common fail point: calls Activelock Acquire procedure
    // Proper License or Trial License
  }
  catch(int e)  // License not valid
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
  m_pActiveLockMfc->DisconnectEvent();         // the activelock callback procedure is disconnected
  m_pActiveLockMfc->DestroyActiveLock();       // kill
  delete m_pActiveLockMfc;                     // possibly overkill but does no harm
  return CWinApp::ExitInstance();
}

// App command to run the dialog
void CMFCSampleApp::OnAppAbout()
{
  CAboutDlg aboutDlg(*m_pActiveLockMfc);
  aboutDlg.DoModal();                                                   // display about dialogue
  m_pActiveLockMfc->SetRegisteredUser(aboutDlg.registeredUser);
  m_pActiveLockMfc->SetLockType(aboutDlg.lockType);

  // write Username and Locktype to ini file
  BOOL bAns;
  bAns =WritePrivateProfileString(LPCTSTR(strApp), "User", LPCTSTR(m_pActiveLockMfc->RegisteredUser()), LPCTSTR(strIni));
  int i = m_pActiveLockMfc->LockTypeToInt(m_pActiveLockMfc->LockType());

  enum ALLockTypes tmp = m_pActiveLockMfc->LockType();
 i = m_pActiveLockMfc->LockTypeToInt(tmp);

  CString cs;
  cs.Format(_T("%i"),i); 
  bAns =WritePrivateProfileString(LPCTSTR(strApp), "lockType", LPCTSTR(cs), LPCTSTR(strIni));
}


BOOL  CMFCSampleApp::CrcsEtc(BOOL debug)
{
  if(debug) return TRUE;
  VectorNameAndCrc vNameCrc;
  DWORD activelock3Crc			= 603433813;    // ***********************these 2 values are wrong
  DWORD alcryptoCrc         = 2581601961;   // ***********************Use debugger to get correct values, dwAns in CheckACrc

  char windowsSystem[255];
  {
    GetSystemDirectory(windowsSystem,255);
    struct CNameAndCrc cnac;
    cnac.crc = activelock3Crc;
    cnac.msg = "Possible wrong version of activelock3.dll";
    cnac.name =  windowsSystem;
    cnac.name += CString("\\activelock3.6.dll");
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

