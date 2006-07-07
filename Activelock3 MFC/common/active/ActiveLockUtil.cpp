// fixtures.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "afxctl.h"

#include <vector>
#include <string>
#include <algorithm>
#include <strstream>
#include <iomanip>

using namespace std;

#include "ConnectionAdvisor.h"   
#include "EventSink.h"
#include "crc32static.h"
#include "ActiveLockUtil.h"



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
}

CActiveLockUtil::~CActiveLockUtil()
{
}

BOOL  CActiveLockUtil::CheckACrc(CString& strActive, DWORD alcryptoCrc, CString msg)
{
  DWORD dwCrc32;
  DWORD dwAns;
  dwAns = CCrc32Static::FileCrc32Streams(LPCTSTR(strActive), dwCrc32) ;
  if( dwAns 
    || dwCrc32 != alcryptoCrc)
  {
    // should I tell him or not - my preference is not
    //AfxMessageBox("ActiveLock dll has been tampered with");
    // but a bit more diplomatically
    AfxMessageBox(LPCTSTR(msg));
    return FALSE;
  }
  return TRUE;
}

BOOL  CActiveLockUtil::CheckAllCrcs(BOOL debug, VectorNameAndCrc& vNameCrc)
{
  // has *.dlls been tampered with
  if(!debug){				// should we check CRC of dll
    for(VectorNameAndCrcIterator it = vNameCrc.begin(); it != vNameCrc.end(); it++){
      if(CheckACrc(it->name, it->crc, it->msg)){
        return FALSE;
      }
    }
  }
  return TRUE;
}


