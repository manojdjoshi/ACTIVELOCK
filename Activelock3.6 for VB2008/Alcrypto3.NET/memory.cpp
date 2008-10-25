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

/**********************************************************************************************
 * Change Log
 * ==========
 *
 * Date (MM/DD/YY)  Author      Description       
 * ---------------  ----------- --------------------------------------------------------------
 * 03/13/06         J.D.M.      Ported to C++
 *
 **********************************************************************************************/

/*---------------------------------------------------------------------------*/
/* My own versions of malloc, realloc and free. Because I want               */
/* malloc and realloc to bomb out and exit the program if they run           */
/* out of memory, realloc to reliably call malloc if passed a NULL           */
/* pointer, and free to reliably do nothing if passed a NULL                 */
/* pointer. We can also put trace printouts in, if we need to; and           */
/* we can also replace the allocator with an ElectricFence-like              */
/* one.                                                                      */
/*---------------------------------------------------------------------------*/

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <assert.h>
#include <stddef.h>          /* for size_t */
#include <string.h>          /* for memcpy() */
#include <malloc.h>

#include "alcrypto3.h"
#include "memory.h"

#ifdef MINEFIELD /* Minefield - a Windows equivalent for Electric Fence */
#define PAGESIZE 4096
/*
* Design:
* 
* We start by reserving as much virtual address space as Windows
* will sensibly (or not sensibly) let us have. We flag it all as
* invalid memory.
* 
* Any allocation attempt is satisfied by committing one or more
* pages, with an uncommitted page on either side. The returned
* memory region is jammed up against the _end_ of the pages.
* 
* Freeing anything causes instantaneous decommitment of the pages
* involved, so stale pointers are caught as soon as possible.
*/
static int  minefield_initialised = 0;
static void *minefield_region = NULL;
static long minefield_size = 0;
static long minefield_npages = 0;
static long minefield_curpos = 0;
static unsigned short *minefield_admin = NULL;
static void *minefield_pages = NULL;

static void minefield_admin_hide(int hide);
static void minefield_init(void);
static void minefield_bomb(void);
static void *minefield_alloc(int size);
static void minefield_free(void *ptr);
static int  minefield_get_size(void *ptr);
static void *minefield_c_malloc(size_t size);
static void minefield_c_free(void *p);
static void *minefield_c_realloc(void *p, size_t size);
#endif

#ifdef MALLOC_LOG
static FILE *fp = NULL;
static char *mlog_file = NULL;
static int  mlog_line = 0;

static void mlog(char *, int);
#endif

static short DoShowMessageBoxOnError = FALSE;
static short DoExitOnError = FALSE;
short DoCatchExceptions = TRUE;

/*---------------------------------------------------------------------------*/
/* ALInitALCrypto                                                            */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI ALInitALCrypto(short ShowMessageBoxOnError, short ExitOnError, short CatchExceptions)
{
  DoShowMessageBoxOnError = ShowMessageBoxOnError;
  DoExitOnError = ExitOnError;
  DoCatchExceptions = CatchExceptions;

  return 0;
} /* ALInitALCrypto */

/*---------------------------------------------------------------------------*/
/* cleanup_exit                                                              */
/*---------------------------------------------------------------------------*/
void cleanup_exit(int code)
{
  if(DoExitOnError==TRUE) {
    exit(code);
  }
  else {
    throw RETVAL_ON_ERROR;
  }
} /* cleanup_exit */

#ifdef MINEFIELD
/*---------------------------------------------------------------------------*/
/* minefield_admin_hide                                                      */
/*---------------------------------------------------------------------------*/
static void minefield_admin_hide(int hide)
{
  int access = hide ? PAGE_NOACCESS : PAGE_READWRITE;

  VirtualProtect(minefield_admin, minefield_npages * 2, access, NULL);
} /* minefield_admin_hide */

/*---------------------------------------------------------------------------*/
/* minefield_init                                                            */
/*---------------------------------------------------------------------------*/
static void minefield_init(void)
{
  int size;
  int admin_size;
  int i;

  for (size = 0x40000000; size > 0; size = ((size >> 3) * 7) & ~0xFFF) {
    minefield_region = VirtualAlloc(NULL, size, MEM_RESERVE, PAGE_NOACCESS);
    if (minefield_region) {
      break;
    }
  }
  minefield_size = size;

  /* Firstly, allocate a section of that to be the admin block. */
  /* We'll need a two-byte field for each page.                 */
  minefield_admin = (unsigned short*)minefield_region;
  minefield_npages = minefield_size / PAGESIZE;
  admin_size = (minefield_npages * 2 + PAGESIZE - 1) & ~(PAGESIZE - 1);
  minefield_npages = (minefield_size - admin_size) / PAGESIZE;
  minefield_pages = (char *) minefield_region + admin_size;

  /* Commit the admin region. */
  VirtualAlloc(minefield_admin, minefield_npages * 2, MEM_COMMIT, PAGE_READWRITE);

  /* Mark all pages as unused (0xFFFF). */
  for (i = 0; i < minefield_npages; i++) {
    minefield_admin[i] = 0xFFFF;
  }

  /* Hide the admin region. */
  minefield_admin_hide(1);

  minefield_initialised = 1;
} /* minefield_init */

/*---------------------------------------------------------------------------*/
/* minefield_bomb                                                            */
/*---------------------------------------------------------------------------*/
static void minefield_bomb(void)
{
  div(1, *(int *) minefield_pages);
} /* minefield_bomb */

/*---------------------------------------------------------------------------*/
/* minefield_alloc                                                           */
/*---------------------------------------------------------------------------*/
static void *minefield_alloc(int size)
{
  int npages;
  int pos, lim, region_end, region_start;
  int start;
  int i;

  npages = (size + PAGESIZE - 1) / PAGESIZE;

  minefield_admin_hide(0);

  /* Search from current position until we find a contiguous */
  /* bunch of npages+2 unused pages.                         */
  pos = minefield_curpos;
  lim = minefield_npages;
  while (1) {
    /* Skip over used pages. */
    while (pos < lim && minefield_admin[pos] != 0xFFFF) {
      pos++;
    }
    /* Count unused pages. */
    start = pos;
    while (pos < lim && pos - start < npages + 2 && minefield_admin[pos] == 0xFFFF) {
      pos++;
    }
    if (pos - start == npages + 2) {
      break;
    }
    /* If we've reached the limit, reset the limit or stop. */
    if (pos >= lim) {
      if (lim == minefield_npages) {
        /* go round and start again at zero */
        lim = minefield_curpos;
        pos = 0;
      }
      else {
        minefield_admin_hide(1);

        return NULL;
      }
    }
  }

  minefield_curpos = pos - 1;

  /* We have npages+2 unused pages starting at start. We leave */
  /* the first and last of these alone and use the rest.       */
  region_end = (start + npages + 1) * PAGESIZE;
  region_start = region_end - size;
  /* FIXME: could align here if we wanted */

  /* Update the admin region. */  
  for (i = start + 2; i < start + npages + 1; i++) {
    minefield_admin[i] = 0xFFFE;   /* used but no region starts here */
  }
  minefield_admin[start + 1] = region_start % PAGESIZE;

  minefield_admin_hide(1);

  VirtualAlloc((char *) minefield_pages + region_start, size,
               MEM_COMMIT, PAGE_READWRITE);

  return (char *) minefield_pages + region_start;
} /* minefield_alloc */

/*---------------------------------------------------------------------------*/
/* minefield_free                                                            */
/*---------------------------------------------------------------------------*/
static void minefield_free(void *ptr)
{
  int region_start, i, j;

  minefield_admin_hide(0);

  region_start = (char *) ptr - (char *) minefield_pages;
  i = region_start / PAGESIZE;
  if (i < 0 || i >= minefield_npages || minefield_admin[i] != region_start % PAGESIZE) {
    minefield_bomb();
  }
  for (j = i; j < minefield_npages && minefield_admin[j] != 0xFFFF; j++) {
    minefield_admin[j] = 0xFFFF;
  }

  VirtualFree(ptr, j * PAGESIZE - region_start, MEM_DECOMMIT);

  minefield_admin_hide(1);
} /* minefield_free */

/*---------------------------------------------------------------------------*/
/* minefield_get_size                                                        */
/*---------------------------------------------------------------------------*/
static int minefield_get_size(void *ptr)
{
  int region_start, i, j;

  minefield_admin_hide(0);

  region_start = (char *) ptr - (char *) minefield_pages;
  i = region_start / PAGESIZE;
  if (i < 0 || i >= minefield_npages || minefield_admin[i] != region_start % PAGESIZE) {
    minefield_bomb();
  }
  for (j = i; j < minefield_npages && minefield_admin[j] != 0xFFFF; j++) {
    ;
  }

  minefield_admin_hide(1);

  return j * PAGESIZE - region_start;
} /* minefield_get_size */

/*---------------------------------------------------------------------------*/
/* minefield_c_malloc                                                        */
/*---------------------------------------------------------------------------*/
static void *minefield_c_malloc(size_t size)
{
  if (!minefield_initialised) {
    minefield_init();
  }

  return minefield_alloc(size);
} /* minefield_c_malloc */

/*---------------------------------------------------------------------------*/
/* minefield_c_free                                                          */
/*---------------------------------------------------------------------------*/
static void minefield_c_free(void *p)
{
  if (!minefield_initialised) {
    minefield_init();
  }
  minefield_free(p);
} /* minefield_c_free */

/*---------------------------------------------------------------------------*/
/* minefield_c_realloc                                                       */
/*---------------------------------------------------------------------------*/
/* realloc _always_ moves the chunk, for rapid detection of code             */
/* that assumes it won't.                                                    */
/*---------------------------------------------------------------------------*/
static void *minefield_c_realloc(void *p, size_t size)
{
  size_t oldsize;
  void   *q;

  if (!minefield_initialised) {
    minefield_init();
  }
  q = minefield_alloc(size);
  oldsize = minefield_get_size(p);
  memcpy(q, p, (oldsize < size ? oldsize : size));
  minefield_free(p);

  return q;
} /* minefield_c_realloc */
#endif /* MINEFIELD */

/*---------------------------------------------------------------------------*/
/* mlog                                                                      */
/*---------------------------------------------------------------------------*/
static void mlog(char *file, int line)
{
#ifdef MALLOC_LOG
  mlog_file = file;
  mlog_line = line;
  if (!fp) {
    fp = fopen("alcrypto_mem.log", "w");
    setvbuf(fp, NULL, _IONBF, BUFSIZ);
  }
  if (fp) {
    fprintf(fp, "%s:%d: ", file, line);
  }
#endif
} /* mlog */

/*---------------------------------------------------------------------------*/
/* safemalloc                                                                */
/*---------------------------------------------------------------------------*/
void *safemalloc(size_t size)
{
  void *p;

#ifdef MINEFIELD
  p = minefield_c_malloc(size);
#else
  p = malloc(size);
#endif
  if (!p) {
    char str[200];

#ifdef MALLOC_LOG
    sprintf(str, "Out of memory! (%s:%d, size=%d)", mlog_file, mlog_line, size);
    fprintf(fp, "*** %s\n", str);
    fclose(fp);
#else
    strcpy(str, "Out of memory!");
#endif
    if(DoShowMessageBoxOnError==TRUE) {
      MessageBox(NULL, str, "ALCrypto Fatal Error", MB_SYSTEMMODAL | MB_ICONERROR | MB_OK);
    }
    cleanup_exit(1);
  }
#ifdef MALLOC_LOG
  if (fp) {
    fprintf(fp, "malloc(%d) returns %p\n", size, p);
  }
#endif

  return p;
} /* safemalloc */

/*---------------------------------------------------------------------------*/
/* saferealloc                                                               */
/*---------------------------------------------------------------------------*/
void *saferealloc(void *ptr, size_t size)
{
  void *p;

  if (!ptr) {
#ifdef MINEFIELD
    p = minefield_c_malloc(size);
#else
    p = malloc(size);
#endif
  }
  else {
#ifdef MINEFIELD
    p = minefield_c_realloc(ptr, size);
#else
    p = realloc(ptr, size);
#endif
  }
  if (!p) {
    char str[200];

#ifdef MALLOC_LOG
    sprintf(str, "Out of memory! (%s:%d, size=%d)", mlog_file, mlog_line, size);
    fprintf(fp, "*** %s\n", str);
    fclose(fp);
#else
    strcpy(str, "Out of memory!");
#endif
    if(DoShowMessageBoxOnError==TRUE) {
      MessageBox(NULL, str, "ALCrypto Fatal Error", MB_SYSTEMMODAL | MB_ICONERROR | MB_OK);
    }
    cleanup_exit(1);
  }
#ifdef MALLOC_LOG
  if (fp) {
    fprintf(fp, "realloc(%p,%d) returns %p\n", ptr, size, p);
  }
#endif

  return p;
} /* saferealloc */

/*---------------------------------------------------------------------------*/
/* saferealloc                                                               */
/*---------------------------------------------------------------------------*/
void safefree(void *ptr)
{
  if (ptr) {
#ifdef MALLOC_LOG
    if (fp) {
      fprintf(fp, "free(%p)\n", ptr);
    }
#endif
#ifdef MINEFIELD
    minefield_c_free(ptr);
#else
    free(ptr);
#endif
  }
#ifdef MALLOC_LOG
  else if (fp) {
    fprintf(fp, "freeing null pointer - no action taken\n");
  }
#endif
} /* safefree */
