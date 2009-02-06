/**
 *   ActiveLock Cryptographic Library
 *   Copyright 2005 The ActiveLock Software Group (ASG)
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
 *  PuTTY License
 *  =============
 *
 *  PuTTY is copyright 1997-2001 Simon Tatham.
 *
 *  Portions copyright Robert de Bath, Joris van Rantwijk, Delian
 *  Delchev, Andreas Schultz, Jeroen Massar, Wez Furlong, Nicolas Barry,
 *  Justin Bradford, and CORE SDI S.A.
 *
 *  Permission is hereby granted, free of charge, to any person
 *  obtaining a copy of this software and associated documentation files
 *  (the "Software"), to deal in the Software without restriction,
 *  including without limitation the rights to use, copy, modify, merge,
 *  publish, distribute, sublicense, and/or sell copies of the Software,
 *  and to permit persons to whom the Software is furnished to do so,
 *  subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be
 *  included in all copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 *  EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 *  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 *  NONINFRINGEMENT.  IN NO EVENT SHALL THE COPYRIGHT HOLDERS BE LIABLE
 *  FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
 *  CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
 *  WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 *
 */

/*
 *     Module Name       : WindowsVersion.cpp
 *
 *     Type              : Functionality concerning Windows versions
 *                         
 *
 *     Author/Location   : J.R.F. De Maeijer, Nieuwegein
 *
 * ----------------------------------------------------------------------------
 *                            MODIFICATION HISTORY
 * ----------------------------------------------------------------------------
 * DATE        REASON                                                    AUTHOR
 * ----------------------------------------------------------------------------
 * 02-Jul-2002 Initial Release                                           J.D.M.
 * ----------------------------------------------------------------------------
 */

#include <windows.h>
#include <windowsx.h>
#include <tchar.h>
#include <stdio.h>
#pragma hdrstop

#include "WindowsVersion.h"
#include "alcrypto3.h"

static short WindowsVersion=WIN_UNKNOWN_VERSION;

static BOOL  Init();

/*---------------------------------------------------------------------------*/
/* ALGetWindowsVersion                                                       */
/*---------------------------------------------------------------------------*/
short ALGetWindowsVersion() {
  static short FirstTime=TRUE;

  if(FirstTime==TRUE) {
    FirstTime=FALSE;
    Init();
  }

  return WindowsVersion;
} /* ALGetWindowsVersion */

/*---------------------------------------------------------------------------*/
/* Init                                                                      */
/*---------------------------------------------------------------------------*/
static BOOL Init() {
  OSVERSIONINFOEX osvi;
  BOOL bOsVersionInfoEx;
  static short FirstTime=TRUE;

  // Try calling GetVersionEx using the OSVERSIONINFOEX structure,
  // which is supported on Windows 2000.
  //
  // If that fails, try using the OSVERSIONINFO structure.

  ZeroMemory(&osvi, sizeof(OSVERSIONINFOEX));
  osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);

  bOsVersionInfoEx = GetVersionEx ((OSVERSIONINFO *) &osvi);
#ifndef NOT_YET_SUPPORTED
  if( bOsVersionInfoEx ) { /* J.D.M.: Force use of OSVERSIONINFO */
    bOsVersionInfoEx = !bOsVersionInfoEx;
  }
#endif

  if( !bOsVersionInfoEx ) {
    // If OSVERSIONINFOEX doesn't work, try OSVERSIONINFO.
    osvi.dwOSVersionInfoSize = sizeof (OSVERSIONINFO);
    if (! GetVersionEx ( (OSVERSIONINFO *) &osvi) ) {
      return FALSE;
    }
  }

  switch (osvi.dwPlatformId) {
    case VER_PLATFORM_WIN32_NT:
      // Test for the product.
      if ( osvi.dwMajorVersion <= 4 ) {
        //_stprintf(_T("Microsoft Windows NT "));
        WindowsVersion=WIN_NT;
      }
      if ( osvi.dwMajorVersion == 5 ) {
        //_stprintf(_T("Microsoft Windows 2000 "));
        WindowsVersion=WIN_2000;
      }
      // Test for workstation versus server.
      if( bOsVersionInfoEx ) {
#ifdef NOT_YET_SUPPORTED
        if ( osvi.wProductType == VER_NT_WORKSTATION ) {
          //_stprintf(_T("Professional "));
          switch(WindowsVersion) {
            case WIN_NT:
              WindowsVersion=WIN_NT_WORKSTATION;
              break;
            case WIN_2000:
              WindowsVersion=WIN_2000_PROFESSIONAL;
              break;
            default:
              break;
          }
        }
        if ( osvi.wProductType == VER_NT_SERVER ) {
          //_stprintf"_T(Server "));
          switch(WindowsVersion) {
            case WIN_NT:
              WindowsVersion=WIN_NT_SERVER;
              break;
            case WIN_2000:
              WindowsVersion=WIN_2000_SERVER;
              break;
            default:
              break;
          }
        }
        if ( osvi.wProductType == VER_NT_DOMAIN_CONTROLLER ) {
          switch(WindowsVersion) {
            case WIN_NT:
              break;
            case WIN_2000:
              WindowsVersion=WIN_2000_DOMAIN_CONTROLLER;
              break;
            default:
              break;
          }
        }
#else
        //_stprintf(_T("???????? "));
#endif
      }
      else {
        HKEY hKey;
        char szProductType[80];
        DWORD dwBufLen;
        RegOpenKeyEx( HKEY_LOCAL_MACHINE,
                      _T("SYSTEM\\CurrentControlSet\\Control\\ProductOptions"),
                      0, KEY_QUERY_VALUE, &hKey );
        RegQueryValueEx( hKey, _T("ProductType"), NULL, NULL,
                        (LPBYTE) szProductType, &dwBufLen);
        RegCloseKey( hKey );
        if ( lstrcmpi( _T("WINNT"), (_TCHAR*)szProductType) == 0 ) {
          //_stprintf(_T("Workstation "));
          switch(WindowsVersion) {
            case WIN_NT:
              WindowsVersion=WIN_NT_WORKSTATION;
              break;
            case WIN_2000:
              WindowsVersion=WIN_2000_PROFESSIONAL;
              break;
            default:
              break;
          }
        }
        if ( lstrcmpi( _T("SERVERNT"), (_TCHAR*)szProductType) == 0 ) {
          //_stprintf(_T("Server "));
          switch(WindowsVersion) {
            case WIN_NT:
              WindowsVersion=WIN_NT_SERVER;
              break;
            case WIN_2000:
              WindowsVersion=WIN_2000_SERVER;
              break;
            default:
              break;
          }
        }
        //if ( lstrcmpi( _T("??DOMAINCONTROLLER??"), szProductType) == 0 ) {
        //  switch(WindowsVersion) {
        //    case WIN_NT:
        //      break;
        //    case WIN_2000:
        //      WindowsVersion=WIN_2000_DOMAIN_CONTROLLER;
        //      break;
        //    default:
        //      break;
        //  }
        //}
      }

      // Display version, service pack (if any), and build number.
      //_stprintf(_T("version %d.%d %s (Build %d)\n",
      //              osvi.dwMajorVersion,
      //              osvi.dwMinorVersion,
      //              osvi.szCSDVersion,
      //              osvi.dwBuildNumber & 0xFFFF));
      break;
    case VER_PLATFORM_WIN32_WINDOWS:
      if ((osvi.dwMajorVersion > 4) || 
         ((osvi.dwMajorVersion == 4) && (osvi.dwMinorVersion > 0))) {
        //_stprintf(_T("Microsoft Windows 98 "));
        WindowsVersion=WIN_98;
      } 
      else {
        //_stprintf(_T("Microsoft Windows 95 "));
        WindowsVersion=WIN_95;
      }
      break;
     case VER_PLATFORM_WIN32s:
      //_stprintf(_T("Microsoft Win32s "));
      WindowsVersion=WIN_31;
      break;
  }

  return TRUE; 
} /* Init */

#ifdef WINVER_EXECUTABLE
/*---------------------------------------------------------------------------*/
/* main                                                                      */
/*---------------------------------------------------------------------------*/
void main() {
  switch(ALGetWindowsVersion()) {
    case WIN_31:
      _tprintf(_T("WIN_31\n"));
      break;
    case WIN_95:
      _tprintf(_T("WIN_95\n"));
      break;
    case WIN_98:
      _tprintf(_T("WIN_98\n"));
      break;
    case WIN_NT:
      _tprintf(_T("WIN_NT\n"));
      break;
    case WIN_NT_WORKSTATION:
      _tprintf(_T("WIN_NT_WORKSTATION\n"));
      break;
    case WIN_NT_SERVER:
      _tprintf(_T("WIN_NT_SERVER\n"));
      break;
    case WIN_2000:
      _tprintf(_T("WIN_2000\n"));
      break;
    case WIN_2000_PROFESSIONAL:
      _tprintf(_T("WIN_2000_PROFESSIONAL\n"));
      break;
    case WIN_2000_DOMAIN_CONTROLLER:
      _tprintf(_T("WIN_2000_DOMAIN_CONTROLLER\n"));
      break;
    case WIN_2000_SERVER:
      _tprintf(_T("WIN_2000_DOMAIN_SERVER\n"));
      break;
    case WIN_UNKNOWN_VERSION:
    default:
      _tprintf(_T("WIN_UNKNOWN_VERSION\n"));
      break;
  }
}
#endif
