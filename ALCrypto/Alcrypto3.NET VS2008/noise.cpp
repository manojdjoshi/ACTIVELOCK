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
 * Noise generation for PuTTY's cryptographic random number
 * generator.
 */

/**********************************************************************************************
 * Change Log
 * ==========
 *
 * Date (MM/DD/YY)  Author      Description       
 * ---------------  ----------- --------------------------------------------------------------
 * 04/21/05         sentax      Upgraded source to version 3
 * 07/27/03         th2tran     Adapted from PuTTY project for used by ActiveLock project.            
 * 03/13/06         J.D.M.      Ported to C++
 *
 **********************************************************************************************/

#include <windows.h>
#include <stdio.h>

#include "alcrypto3.h"
#include "noise.h"
#include "memory.h"
#include "rand.h"

/*
 * GetSystemPowerStatus function.
 */
typedef BOOL(WINAPI * gsps_t) (LPSYSTEM_POWER_STATUS);
static gsps_t gsps;

/*
 * This function is called once, at PuTTY startup, and will do some
 * seriously silly things like listing directories and getting disk
 * free space and a process snapshot.
 */

static char seedpath[2 * MAX_PATH + 10] = "\0";

static void random_save_seed(void);
static void get_seedpath(void);

/*---------------------------------------------------------------------------*/
/* noise_get_heavy                                                           */
/*---------------------------------------------------------------------------*/
void noise_get_heavy(void (*func) (void *, int))
{
  HANDLE          srch;
  WIN32_FIND_DATA finddata;
  char            winpath[MAX_PATH + 3];
  HMODULE         mod;
  
  GetWindowsDirectory(winpath, sizeof(winpath));
  strcat(winpath, "\\*");
  srch = FindFirstFile(winpath, &finddata);
  if (srch != INVALID_HANDLE_VALUE) {
    do {
      func(&finddata, sizeof(finddata));
    } while (FindNextFile(srch, &finddata));
    FindClose(srch);
  }
  
  read_random_seed(func);
  /* Update the seed immediately, in case another instance uses it. */
  random_save_seed();
  
  gsps = NULL;
  mod = GetModuleHandle("KERNEL32");
  if (mod) {
    gsps = (gsps_t) GetProcAddress(mod, "GetSystemPowerStatus");
  }
} /* noise_get_heavy */

/*---------------------------------------------------------------------------*/
/* noise_get_light                                                           */
/*---------------------------------------------------------------------------*/
/* This function is called every time the random pool needs                  */
/* stirring, and will acquire the system time in all available               */
/* forms and the battery status.                                             */
/*---------------------------------------------------------------------------*/
void noise_get_light(void (*func) (void *, int))
{
  SYSTEMTIME          systime;
  DWORD               adjust[2];
  BOOL                rubbish;
  SYSTEM_POWER_STATUS pwrstat;
  
  GetSystemTime(&systime);
  func(&systime, sizeof(systime));
  
  GetSystemTimeAdjustment(&adjust[0], &adjust[1], &rubbish);
  func(&adjust, sizeof(adjust));
  
  /*
  * Call GetSystemPowerStatus if present.
  */
  if (gsps) {
    if (gsps(&pwrstat)) {
      func(&pwrstat, sizeof(pwrstat));
    }
  }
} /* noise_get_light */

/*---------------------------------------------------------------------------*/
/* noise_regular                                                             */
/*---------------------------------------------------------------------------*/
/* This function is called on a timer, and it will monitor                   */
/* frequently changing quantities such as the state of physical and          */
/* virtual memory, the state of the process's message queue, which           */
/* window is in the foreground, which owns the clipboard, etc.               */
/*---------------------------------------------------------------------------*/
void noise_regular(void)
{
  HWND         w;
  DWORD        z;
  POINT        pt;
  MEMORYSTATUS memstat;
  FILETIME     times[4];
  
  w = GetForegroundWindow();
  random_add_noise(&w, sizeof(w));
  w = GetCapture();
  random_add_noise(&w, sizeof(w));
  w = GetClipboardOwner();
  random_add_noise(&w, sizeof(w));
  z = GetQueueStatus(QS_ALLEVENTS);
  random_add_noise(&z, sizeof(z));
  
  GetCursorPos(&pt);
  random_add_noise(&pt, sizeof(pt));
  
  GlobalMemoryStatus(&memstat);
  random_add_noise(&memstat, sizeof(memstat));
  
  GetThreadTimes(GetCurrentThread(), times, times + 1, times + 2, times + 3);
  random_add_noise(&times, sizeof(times));
  GetProcessTimes(GetCurrentProcess(), times, times + 1, times + 2, times + 3);
  random_add_noise(&times, sizeof(times));
} /* noise_regular */

/*---------------------------------------------------------------------------*/
/* noise_ultralight                                                          */
/*---------------------------------------------------------------------------*/
/* This function is called on every keypress or mouse move, and              */
/* will add the current Windows time and performance monitor                 */
/* counter to the noise pool. It gets the scan code or mouse                 */
/* position passed in.                                                       */
/*---------------------------------------------------------------------------*/
void noise_ultralight(DWORD data)
{
  DWORD wintime;
  LARGE_INTEGER perftime;
  
  random_add_noise(&data, sizeof(DWORD));
  
  wintime = GetTickCount();
  random_add_noise(&wintime, sizeof(DWORD));
  
  if (QueryPerformanceCounter(&perftime)) {
    random_add_noise(&perftime, sizeof(perftime));
  }
} /* noise_ultralight */

/*---------------------------------------------------------------------------*/
/* get_seedpath                                                              */
/*---------------------------------------------------------------------------*/
/* Find the random seed file path and store it in `seedpath'.                */
/*---------------------------------------------------------------------------*/
static void get_seedpath(void)
{
  HKEY  rkey;
  DWORD type, size;

  size = sizeof(seedpath);

  if (RegOpenKey(HKEY_CURRENT_USER, REG_HOME, &rkey) == ERROR_SUCCESS) {
    int ret = RegQueryValueEx(rkey, "RandSeedFile", 0, &type, (unsigned char*)seedpath, &size);

    if (ret != ERROR_SUCCESS || type != REG_SZ) {
      seedpath[0] = '\0';
    }
    RegCloseKey(rkey);
  }
  else {
    seedpath[0] = '\0';
  }
  
  if (!seedpath[0]) {
    int len, ret;

    len = GetEnvironmentVariable("HOMEDRIVE", seedpath, sizeof(seedpath));
    ret = GetEnvironmentVariable("HOMEPATH", seedpath + len, sizeof(seedpath) - len);
    if (ret == 0) { /* probably win95; store in \WINDOWS */
      GetWindowsDirectory(seedpath, sizeof(seedpath));
      len = strlen(seedpath);
    }
    else {
      len += ret;
    }
    strcpy(seedpath + len, "\\ACTIVELOCK.RND");
  }
} /* get_seedpath */

/*---------------------------------------------------------------------------*/
/* read_random_seed                                                          */
/*---------------------------------------------------------------------------*/
void read_random_seed(noise_consumer_t consumer)
{
  HANDLE seedf;

  if (!seedpath[0]) {
    get_seedpath();
  }

  seedf = CreateFile(seedpath, GENERIC_READ,
                     FILE_SHARE_READ | FILE_SHARE_WRITE,
                     NULL, OPEN_EXISTING, 0, NULL);

  if (seedf != INVALID_HANDLE_VALUE) {
    while (1) {
      char  buf[1024];
      DWORD len;

      if (ReadFile(seedf, buf, sizeof(buf), &len, NULL) && len) {
        consumer(buf, len);
      }
      else {
        break;
      }
    }
    CloseHandle(seedf);
  }
} /* read_random_seed */

/*---------------------------------------------------------------------------*/
/* write_random_seed                                                         */
/*---------------------------------------------------------------------------*/
void write_random_seed(void *data, int len)
{
  HANDLE seedf;

  if (!seedpath[0]) {
    get_seedpath();
  }

  seedf = CreateFile(seedpath, GENERIC_WRITE, 0,
                     NULL, CREATE_ALWAYS, FILE_ATTRIBUTE_NORMAL, NULL);

  if (seedf != INVALID_HANDLE_VALUE) {
    DWORD lenwritten;
    
    WriteFile(seedf, data, len, &lenwritten, NULL);
    CloseHandle(seedf);
  }
} /* write_random_seed */

/*---------------------------------------------------------------------------*/
/* random_save_seed                                                          */
/*---------------------------------------------------------------------------*/
static void random_save_seed(void)
{
  int  len;
  void *data;
  
  if (random_active) {
    random_get_savedata(&data, &len);
    write_random_seed(data, len);
    sfree(data);
  }
} /* random_save_seed */
