/* ------------------------------------------------- */
/*                   WinIo v1.2                      */
/*  Direct Hardware Access Under Windows 9x/NT/2000  */
/*        Copyright 1998-2000 Yariv Kaplan           */
/*           http://www.internals.com                */
/* ------------------------------------------------- */

#include <windows.h>

#include "alcrypto3.h"
#include "winio.h"

HANDLE hDriver;
bool   IsNT;
bool   IsWinIoInitialized = false;

static bool IsWinNT();

/*---------------------------------------------------------------------------*/
/* IsWinNT                                                                   */
/*---------------------------------------------------------------------------*/
static bool IsWinNT()
{
  OSVERSIONINFO OSVersionInfo;
  
  OSVersionInfo.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);
  
  GetVersionEx(&OSVersionInfo);
  
  return OSVersionInfo.dwPlatformId == VER_PLATFORM_WIN32_NT;
} /* IsWinNT */

/*---------------------------------------------------------------------------*/
/* InitializeWinIo                                                           */
/*---------------------------------------------------------------------------*/
bool InitializeWinIo()
{
  char szExePath[MAX_PATH];
  PSTR pszSlash;
  
  IsNT = IsWinNT();
  
  if (IsNT) {
    if (!GetModuleFileName(GetModuleHandle(NULL), szExePath, sizeof(szExePath))) {
      return false;
    }
    
    pszSlash = strrchr(szExePath, '\\');
    if (pszSlash) {
      pszSlash[1] = 0;
    }
    else {
      return false;
    }
    
    strcat(szExePath, "winio.sys");

#ifdef DONT_USE
    UnloadDeviceDriver("WINIO");
    if (!LoadDeviceDriver("WINIO", szExePath, &hDriver)) {
      return false;
    }
#endif
  }
  
  IsWinIoInitialized = true;

  return true;
} /* InitializeWinIo */

/*---------------------------------------------------------------------------*/
/* ShutdownWinIo                                                             */
/*---------------------------------------------------------------------------*/
void ShutdownWinIo()
{
#ifdef DONT_USE
  if (IsNT) {
    UnloadDeviceDriver("WINIO");
  }
#endif
} /* ShutdownWinIo */
