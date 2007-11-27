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

/**
 * Miscellaneous utility functions.
 */

/**********************************************************************************************
 * Change Log
 * ==========
 *
 * Date (MM/DD/YY)  Author      Description       
 * ---------------  ----------- --------------------------------------------------------------
 * 04/21/05     sentax    Upgraded source to version 3
 * 07/27/03     th2tran   adapted from PuTTY project for used by ActiveLock project.            
 * 03/13/06         J.D.M.      Ported to C++
 *
 **********************************************************************************************/

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <assert.h>
#include <stddef.h>          /* for size_t */
#include <string.h>          /* for memcpy() */

#include "alcrypto3.h"
#include "misc.h"
#include "memory.h"

#define BUFFER_GRANULE  512

struct bufchain_granule {
  struct bufchain_granule *next;
  int buflen, bufpos;
  char buf[BUFFER_GRANULE];
};

static char hex[17] = "0123456789ABCDEF";

#ifdef _DEBUG
static FILE *debug_fp = NULL;
static int debug_got_console = 0;
#endif

static char *dupstr(char *s);
static char *dupcat(char *s1, ...);

/*---------------------------------------------------------------------------*/
/* String handling routines.                                                 */
/*---------------------------------------------------------------------------*/

/*---------------------------------------------------------------------------*/
/* dupstr                                                                    */
/*---------------------------------------------------------------------------*/
static char *dupstr(char *s)
{
  int  len = strlen(s);
  char *p = (char*)smalloc(len + 1);

  strcpy(p, s);

  return p;
} /* dupstr */

/*---------------------------------------------------------------------------*/
/* dupcat                                                                    */
/*---------------------------------------------------------------------------*/
/* Allocate the concatenation of N strings. Terminate arg list with NULL.    */
/*---------------------------------------------------------------------------*/
static char *dupcat(char *s1, ...)
{
  int     len;
  char    *p, *q, *sn;
  va_list ap;

  len = strlen(s1);
  va_start(ap, s1);
  while (1) {
    sn = va_arg(ap, char *);
    if (!sn) {
      break;
    }
    len += strlen(sn);
  }
  va_end(ap);

  p = (char*)smalloc(len + 1);
  strcpy(p, s1);
  q = p + strlen(p);

  va_start(ap, s1);
  while (1) {
    sn = va_arg(ap, char *);
    if (!sn) {
      break;
    }
    strcpy(q, sn);
    q += strlen(q);
  }
  va_end(ap);

  return p;
} /* dupcat */

/*---------------------------------------------------------------------------*/
/* base64_encode_atom                                                        */
/*---------------------------------------------------------------------------*/
/* Base64 encoding routine. This is required in public-key writing           */
/* but also in HTTP proxy handling, so it's centralised here.                */
/*---------------------------------------------------------------------------*/
void base64_encode_atom(unsigned char *data, int n, char *out)
{
  static const char base64_chars[] =
    "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";
  unsigned          word;

  word = data[0] << 16;
  if (n > 1) {
    word |= data[1] << 8;
  }
  if (n > 2) {
    word |= data[2];
  }
  out[0] = base64_chars[(word >> 18) & 0x3F];
  out[1] = base64_chars[(word >> 12) & 0x3F];
  if (n > 1) {
    out[2] = base64_chars[(word >> 6) & 0x3F];
  }
  else {
    out[2] = '=';
  }
  if (n > 2) {
    out[3] = base64_chars[word & 0x3F];
  }
  else {
    out[3] = '=';
  }
} /* base64_encode_atom */

/*---------------------------------------------------------------------------*/
/* base64_decode_atom                                                        */
/*---------------------------------------------------------------------------*/
int base64_decode_atom(char *atom, unsigned char *out)
{
  int vals[4];
  int i, v, len;
  unsigned word;
  char c;

  for (i = 0; i < 4; i++) {
    c = atom[i];
    if (c >= 'A' && c <= 'Z') {
      v = c - 'A';
    }
    else if (c >= 'a' && c <= 'z') {
      v = c - 'a' + 26;
    }
    else if (c >= '0' && c <= '9') {
      v = c - '0' + 52;
    }
    else if (c == '+') {
      v = 62;
    }
    else if (c == '/') {
      v = 63;
    }
    else if (c == '=') {
      v = -1;
    }
    else {
      return 0;          /* invalid atom */
    }
    vals[i] = v;
  }

  if (vals[0] == -1 || vals[1] == -1) {
    return 0;
  }
  if (vals[2] == -1 && vals[3] != -1) {
    return 0;
  }

  if (vals[3] != -1) {
    len = 3;
  }
  else if (vals[2] != -1) {
    len = 2;
  }
  else {
    len = 1;
  }

  word = ((vals[0] << 18) | (vals[1] << 12) | ((vals[2] & 0x3F) << 6) | (vals[3] & 0x3F));
  out[0] = (word >> 16) & 0xFF;
  if (len > 1) {
    out[1] = (word >> 8) & 0xFF;
  }
  if (len > 2) {
    out[2] = word & 0xFF;
  }

  return len;
} /* base64_decode_atom */

/*---------------------------------------------------------------------------*/
/* Generic routines to deal with send buffers: a linked list of              */
/* smallish blocks, with the operations                                      */
/*                                                                           */
/* - add an arbitrary amount of data to the end of the list                  */
/* - remove the first N bytes from the list                                  */
/* - return a (pointer,length) pair giving some initial data in              */
/*   the list, suitable for passing to a send or write system                */
/*   call                                                                    */
/* - retrieve a larger amount of initial data from the list                  */
/* - return the current size of the buffer chain in bytes                    */
/*---------------------------------------------------------------------------*/

/*---------------------------------------------------------------------------*/
/* bufchain_init                                                             */
/*---------------------------------------------------------------------------*/
void bufchain_init(bufchain *ch)
{
  ch->head = ch->tail = NULL;
  ch->buffersize = 0;
} /* bufchain_init */

/*---------------------------------------------------------------------------*/
/* bufchain_clear                                                            */
/*---------------------------------------------------------------------------*/
void bufchain_clear(bufchain *ch)
{
  struct bufchain_granule *b;

  while (ch->head) {
    b = ch->head;
    ch->head = ch->head->next;
    sfree(b);
  }
  ch->tail = NULL;
  ch->buffersize = 0;
} /* bufchain_clear */

/*---------------------------------------------------------------------------*/
/* bufchain_size                                                             */
/*---------------------------------------------------------------------------*/
int bufchain_size(bufchain *ch)
{
  return ch->buffersize;
} /* bufchain_size */

/*---------------------------------------------------------------------------*/
/* bufchain_add                                                              */
/*---------------------------------------------------------------------------*/
void bufchain_add(bufchain *ch, void *data, int len)
{
  char *buf = (char *)data;

  ch->buffersize += len;

  if (ch->tail && ch->tail->buflen < BUFFER_GRANULE) {
    int copylen = min(len, BUFFER_GRANULE - ch->tail->buflen);

    memcpy(ch->tail->buf + ch->tail->buflen, buf, copylen);
    buf += copylen;
    len -= copylen;
    ch->tail->buflen += copylen;
  }
  while (len > 0) {
    int    grainlen = min(len, BUFFER_GRANULE);
    struct bufchain_granule *newbuf;

    newbuf = (struct bufchain_granule*)smalloc(sizeof(struct bufchain_granule));
    newbuf->bufpos = 0;
    newbuf->buflen = grainlen;
    memcpy(newbuf->buf, buf, grainlen);
    buf += grainlen;
    len -= grainlen;
    if (ch->tail) {
      ch->tail->next = newbuf;
    }
    else {
      ch->head = ch->tail = newbuf;
    }
    newbuf->next = NULL;
    ch->tail = newbuf;
  }
} /* bufchain_add */

/*---------------------------------------------------------------------------*/
/* bufchain_consume                                                          */
/*---------------------------------------------------------------------------*/
void bufchain_consume(bufchain *ch, int len)
{
  struct bufchain_granule *tmp;

  assert(ch->buffersize >= len);
  while (len > 0) {
    int remlen = len;

    assert(ch->head != NULL);
    if (remlen >= ch->head->buflen - ch->head->bufpos) {
      remlen = ch->head->buflen - ch->head->bufpos;
      tmp = ch->head;
      ch->head = tmp->next;
      sfree(tmp);
      if (!ch->head) {
        ch->tail = NULL;
      }
    }
    else {
      ch->head->bufpos += remlen;
    }
    ch->buffersize -= remlen;
    len -= remlen;
  }
} /* bufchain_consume */

/*---------------------------------------------------------------------------*/
/* bufchain_prefix                                                           */
/*---------------------------------------------------------------------------*/
void bufchain_prefix(bufchain *ch, void **data, int *len)
{
  *len = ch->head->buflen - ch->head->bufpos;
  *data = ch->head->buf + ch->head->bufpos;
} /* bufchain_prefix */

/*---------------------------------------------------------------------------*/
/* bufchain_fetch                                                            */
/*---------------------------------------------------------------------------*/
void bufchain_fetch(bufchain *ch, void *data, int len)
{
  struct bufchain_granule *tmp;
  char   *data_c = (char *)data;

  tmp = ch->head;

  assert(ch->buffersize >= len);
  while (len > 0) {
    int remlen = len;

    assert(tmp != NULL);
    if (remlen >= tmp->buflen - tmp->bufpos) {
      remlen = tmp->buflen - tmp->bufpos;
    }
    memcpy(data_c, tmp->buf + tmp->bufpos, remlen);

    tmp = tmp->next;
    len -= remlen;
    data_c += remlen;
  }
} /* bufchain_fetch */

/*---------------------------------------------------------------------------*/
/* Debugging routines.                                                       */
/*---------------------------------------------------------------------------*/
#ifdef _DEBUG

/*---------------------------------------------------------------------------*/
/* dputs                                                                     */
/*---------------------------------------------------------------------------*/
static void dputs(char *buf)
{
  DWORD dw;

  if (!debug_got_console) {
    AllocConsole();
    debug_got_console = 1;
  }
  if (!debug_fp) {
    debug_fp = fopen("debug.log", "w");
  }

  WriteFile(GetStdHandle(STD_OUTPUT_HANDLE), buf, strlen(buf), &dw, NULL);
  fputs(buf, debug_fp);
  fflush(debug_fp);
} /* dputs */

/*---------------------------------------------------------------------------*/
/* dprintf                                                                   */
/*---------------------------------------------------------------------------*/
void dprintf(char *fmt, ...)
{
  char    buf[2048];
  va_list ap;

  va_start(ap, fmt);
  vsprintf(buf, fmt, ap);
  dputs(buf);
  va_end(ap);
} /* dprintf */

/*---------------------------------------------------------------------------*/
/* debug_memdump                                                             */
/*---------------------------------------------------------------------------*/
void debug_memdump(void *buf, int len, int L)
{
  int           i;
  unsigned char *p = (unsigned char*)buf;
  char          foo[17];

  if (L) {
    int delta;

    dprintf("\t%d (0x%x) bytes:\n", len, len);
    delta = 15 & (int) p;
    p -= delta;
    len += delta;
  }
  for (; 0 < len; p += 16, len -= 16) {
    dputs("  ");
    if (L) {
      dprintf("%p: ", p);
    }
    strcpy(foo, "................");  /* sixteen dots */
    for (i = 0; i < 16 && i < len; ++i) {
      if (&p[i] < (unsigned char *) buf) {
        dputs("   ");        /* 3 spaces */
        foo[i] = ' ';
      }
      else {
        dprintf("%c%02.2x", &p[i] != (unsigned char *) buf && i % 4 ? '.' : ' ', p[i]);
        if (p[i] >= ' ' && p[i] <= '~') {
          foo[i] = (char) p[i];
        }
      }
    }
    foo[i] = '\0';
    dprintf("%*s%s\n", (16 - i) * 3 + 2, "", foo);
  }
} /* debug_memdump */
#endif /* DEBUG */
