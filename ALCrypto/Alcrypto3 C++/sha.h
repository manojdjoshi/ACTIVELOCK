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

#ifndef ALCRYPTO_SHA_H
#define ALCRYPTO_SHA_H

#ifdef __cplusplus
extern "C" {
#endif

struct ssh_mac {
  void (*setcskey) (unsigned char *key);
  void (*setsckey) (unsigned char *key);
  void (*generate) (unsigned char *blk, int len, unsigned long seq);
  int  (*verify) (unsigned char *blk, int len, unsigned long seq);
  char *name;
  int  len;
};

typedef struct {
  uint32 h[5];
  unsigned char block[64];
  int blkused;
  uint32 lenhi, lenlo;
} SHA_State;

extern void SHATransform(word32 * digest, word32 * block);
extern void SHA_Simple(void *p, int len, unsigned char *output);

#ifdef __cplusplus
}
#endif

#endif