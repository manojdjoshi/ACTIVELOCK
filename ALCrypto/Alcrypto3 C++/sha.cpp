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
 * SHA1 hash algorithm. Used in SSH2 as a MAC, and the transform is
 * also used as a `stirring' function for the PuTTY random number
 * pool. Implemented directly from the specification by Simon
 * Tatham.
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

#include "alcrypto3.h"
#include "sha.h"

static SHA_State sha1_cs_mac_s1, sha1_cs_mac_s2;
static SHA_State sha1_sc_mac_s1, sha1_sc_mac_s2;

static void SHA_Core_Init(uint32 h[5]);
static void SHA_Init(SHA_State * s);
static void SHA_Bytes(SHA_State * s, void *p, int len);
static void SHA_Final(SHA_State * s, unsigned char *output);
static void sha1_key(SHA_State * s1, SHA_State * s2, unsigned char *key, int len);
static void sha1_cskey(unsigned char *key);
static void sha1_sckey(unsigned char *key);
static void sha1_cskey_buggy(unsigned char *key);
static void sha1_sckey_buggy(unsigned char *key);
static void sha1_do_hmac(SHA_State * s1, SHA_State * s2, unsigned char *blk, int len, unsigned long seq, unsigned char *hmac);
static void sha1_generate(unsigned char *blk, int len, unsigned long seq);
static int  sha1_verify(unsigned char *blk, int len, unsigned long seq);
static void hmac_sha1_simple(void *key, int keylen, void *data, int datalen, unsigned char *output);

const struct ssh_mac ssh_sha1 = {
  sha1_cskey,
  sha1_sckey,
  sha1_generate,
  sha1_verify,
  "hmac-sha1",
  20
};

const struct ssh_mac ssh_sha1_buggy = {
  sha1_cskey_buggy,
  sha1_sckey_buggy,
  sha1_generate,
  sha1_verify,
  "hmac-sha1",
  20
};

/*---------------------------------------------------------------------------*/
/* Core SHA algorithm: processes 16-word blocks into a message digest.       */
/*---------------------------------------------------------------------------*/

#define rol(x,y) ( ((x) << (y)) | (((uint32)x) >> (32-y)) )

/*---------------------------------------------------------------------------*/
/* SHA_Core_Init                                                             */
/*---------------------------------------------------------------------------*/
static void SHA_Core_Init(uint32 h[5])
{
  h[0] = 0x67452301;
  h[1] = 0xefcdab89;
  h[2] = 0x98badcfe;
  h[3] = 0x10325476;
  h[4] = 0xc3d2e1f0;
} /* SHA_Core_Init */

/*---------------------------------------------------------------------------*/
/* SHATransform                                                              */
/*---------------------------------------------------------------------------*/
void SHATransform(word32 * digest, word32 * block)
{
  word32 w[80];
  word32 a, b, c, d, e;
  int    t;

  for (t = 0; t < 16; t++) {
    w[t] = block[t];
  }

  for (t = 16; t < 80; t++) {
    word32 tmp = w[t - 3] ^ w[t - 8] ^ w[t - 14] ^ w[t - 16];
    w[t] = rol(tmp, 1);
  }

  a = digest[0];
  b = digest[1];
  c = digest[2];
  d = digest[3];
  e = digest[4];

  for (t = 0; t < 20; t++) {
    word32 tmp = rol(a, 5) + ((b & c) | (d & ~b)) + e + w[t] + 0x5a827999;
    e = d;
    d = c;
    c = rol(b, 30);
    b = a;
    a = tmp;
  }
  for (t = 20; t < 40; t++) {
    word32 tmp = rol(a, 5) + (b ^ c ^ d) + e + w[t] + 0x6ed9eba1;
    e = d;
    d = c;
    c = rol(b, 30);
    b = a;
    a = tmp;
  }
  for (t = 40; t < 60; t++) {
    word32 tmp = rol(a, 5) + ((b & c) | (b & d) | (c & d)) + e + w[t] + 0x8f1bbcdc;
    e = d;
    d = c;
    c = rol(b, 30);
    b = a;
    a = tmp;
  }
  for (t = 60; t < 80; t++) {
    word32 tmp = rol(a, 5) + (b ^ c ^ d) + e + w[t] + 0xca62c1d6;
    e = d;
    d = c;
    c = rol(b, 30);
    b = a;
    a = tmp;
  }

  digest[0] += a;
  digest[1] += b;
  digest[2] += c;
  digest[3] += d;
  digest[4] += e;
} /* SHATransform */

/*---------------------------------------------------------------------------*/
/* Outer SHA algorithm: take an arbitrary length byte string,                */
/* convert it into 16-word blocks with the prescribed padding at             */
/* the end, and pass those blocks to the core SHA algorithm.                 */
/*---------------------------------------------------------------------------*/

/*---------------------------------------------------------------------------*/
/* SHA_Init                                                                  */
/*---------------------------------------------------------------------------*/
static void SHA_Init(SHA_State * s)
{
  SHA_Core_Init(s->h);
  s->blkused = 0;
  s->lenhi = s->lenlo = 0;
} /* SHA_Init */

/*---------------------------------------------------------------------------*/
/* SHA_Bytes                                                                 */
/*---------------------------------------------------------------------------*/
static void SHA_Bytes(SHA_State * s, void *p, int len)
{
  unsigned char *q = (unsigned char *) p;
  uint32        wordblock[16];
  uint32        lenw = len;
  int           i;

  /* Update the length field. */
  s->lenlo += lenw;
  s->lenhi += (s->lenlo < lenw);

  if (s->blkused && s->blkused + len < 64) {
    /* Trivial case: just add to the block. */
    memcpy(s->block + s->blkused, q, len);
    s->blkused += len;
  }
  else {
    /* We must complete and process at least one block. */
    while (s->blkused + len >= 64) {
      memcpy(s->block + s->blkused, q, 64 - s->blkused);
      q += 64 - s->blkused;
      len -= 64 - s->blkused;
      /* Now process the block. Gather bytes big-endian into words */
      for (i = 0; i < 16; i++) {
        wordblock[i] = (((uint32) s->block[i * 4 + 0]) << 24) |
                       (((uint32) s->block[i * 4 + 1]) << 16) |
                       (((uint32) s->block[i * 4 + 2]) << 8) |
                       (((uint32) s->block[i * 4 + 3]) << 0);
      }
      SHATransform(s->h, wordblock);
      s->blkused = 0;
    }
    memcpy(s->block, q, len);
    s->blkused = len;
  }
} /* SHA_Bytes */

/*---------------------------------------------------------------------------*/
/* SHA_Final                                                                 */
/*---------------------------------------------------------------------------*/
static void SHA_Final(SHA_State * s, unsigned char *output)
{
  int           i;
  int           pad;
  unsigned char c[64];
  uint32        lenhi, lenlo;

  if (s->blkused >= 56) {
    pad = 56 + 64 - s->blkused;
  }
  else {
    pad = 56 - s->blkused;
  }

  lenhi = (s->lenhi << 3) | (s->lenlo >> (32 - 3));
  lenlo = (s->lenlo << 3);

  memset(c, 0, pad);
  c[0] = 0x80;
  SHA_Bytes(s, &c, pad);

  c[0] = (lenhi >> 24) & 0xFF;
  c[1] = (lenhi >> 16) & 0xFF;
  c[2] = (lenhi >> 8) & 0xFF;
  c[3] = (lenhi >> 0) & 0xFF;
  c[4] = (lenlo >> 24) & 0xFF;
  c[5] = (lenlo >> 16) & 0xFF;
  c[6] = (lenlo >> 8) & 0xFF;
  c[7] = (lenlo >> 0) & 0xFF;

  SHA_Bytes(s, &c, 8);

  for (i = 0; i < 5; i++) {
    output[i * 4] = (s->h[i] >> 24) & 0xFF;
    output[i * 4 + 1] = (s->h[i] >> 16) & 0xFF;
    output[i * 4 + 2] = (s->h[i] >> 8) & 0xFF;
    output[i * 4 + 3] = (s->h[i]) & 0xFF;
  }
} /* SHA_Final */

/*---------------------------------------------------------------------------*/
/* SHA_Simple                                                                */
/*---------------------------------------------------------------------------*/
void SHA_Simple(void *p, int len, unsigned char *output)
{
  SHA_State s;

  SHA_Init(&s);
  SHA_Bytes(&s, p, len);
  SHA_Final(&s, output);
} /* SHA_Simple */

/*---------------------------------------------------------------------------*/
/* The above is the SHA-1 algorithm itself. Now we implement the             */
/* HMAC wrapper on it.                                                       */
/*---------------------------------------------------------------------------*/

/*---------------------------------------------------------------------------*/
/* sha1_key                                                                  */
/*---------------------------------------------------------------------------*/
static void sha1_key(SHA_State * s1, SHA_State * s2, unsigned char *key, int len)
{
  unsigned char foo[64];
  int           i;

  memset(foo, 0x36, 64);
  for (i = 0; i < len && i < 64; i++) {
    foo[i] ^= key[i];
  }
  SHA_Init(s1);
  SHA_Bytes(s1, foo, 64);

  memset(foo, 0x5C, 64);
  for (i = 0; i < len && i < 64; i++) {
    foo[i] ^= key[i];
  }
  SHA_Init(s2);
  SHA_Bytes(s2, foo, 64);

  memset(foo, 0, 64);          /* burn the evidence */
} /* sha1_key */

/*---------------------------------------------------------------------------*/
/* sha1_cskey                                                                */
/*---------------------------------------------------------------------------*/
static void sha1_cskey(unsigned char *key)
{
  sha1_key(&sha1_cs_mac_s1, &sha1_cs_mac_s2, key, 20);
} /* sha1_cskey */

/*---------------------------------------------------------------------------*/
/* sha1_sckey                                                                */
/*---------------------------------------------------------------------------*/
static void sha1_sckey(unsigned char *key)
{
  sha1_key(&sha1_sc_mac_s1, &sha1_sc_mac_s2, key, 20);
} /* sha1_sckey */

/*---------------------------------------------------------------------------*/
/* sha1_cskey_buggy                                                          */
/*---------------------------------------------------------------------------*/
static void sha1_cskey_buggy(unsigned char *key)
{
  sha1_key(&sha1_cs_mac_s1, &sha1_cs_mac_s2, key, 16);
} /* sha1_cskey_buggy */

/*---------------------------------------------------------------------------*/
/* sha1_sckey_buggy                                                          */
/*---------------------------------------------------------------------------*/
static void sha1_sckey_buggy(unsigned char *key)
{
  sha1_key(&sha1_sc_mac_s1, &sha1_sc_mac_s2, key, 16);
} /* sha1_sckey_buggy */

/*---------------------------------------------------------------------------*/
/* sha1_do_hmac                                                              */
/*---------------------------------------------------------------------------*/
static void sha1_do_hmac(SHA_State * s1, SHA_State * s2,
                         unsigned char *blk, int len, unsigned long seq,
                         unsigned char *hmac)
{
  SHA_State     s;
  unsigned char intermediate[20];

  intermediate[0] = (unsigned char) ((seq >> 24) & 0xFF);
  intermediate[1] = (unsigned char) ((seq >> 16) & 0xFF);
  intermediate[2] = (unsigned char) ((seq >> 8) & 0xFF);
  intermediate[3] = (unsigned char) ((seq) & 0xFF);

  s = *s1; /* structure copy */
  SHA_Bytes(&s, intermediate, 4);
  SHA_Bytes(&s, blk, len);
  SHA_Final(&s, intermediate);

  s = *s2; /* structure copy */
  SHA_Bytes(&s, intermediate, 20);
  SHA_Final(&s, hmac);
} /* sha1_do_hmac */

/*---------------------------------------------------------------------------*/
/* sha1_generate                                                             */
/*---------------------------------------------------------------------------*/
static void sha1_generate(unsigned char *blk, int len, unsigned long seq)
{
  sha1_do_hmac(&sha1_cs_mac_s1, &sha1_cs_mac_s2, blk, len, seq, blk + len);
} /* sha1_generate */

/*---------------------------------------------------------------------------*/
/* sha1_verify                                                               */
/*---------------------------------------------------------------------------*/
static int sha1_verify(unsigned char *blk, int len, unsigned long seq)
{
  unsigned char correct[20];

  sha1_do_hmac(&sha1_sc_mac_s1, &sha1_sc_mac_s2, blk, len, seq, correct);

  return !memcmp(correct, blk + len, 20);
} /* sha1_verify */

/*---------------------------------------------------------------------------*/
/* hmac_sha1_simple                                                          */
/*---------------------------------------------------------------------------*/
static void hmac_sha1_simple(void *key, int keylen, void *data, int datalen,
                             unsigned char *output) {
  SHA_State     s1, s2;
  unsigned char intermediate[20];

  sha1_key(&s1, &s2, (unsigned char*)key, keylen);
  SHA_Bytes(&s1, data, datalen);
  SHA_Final(&s1, intermediate);

  SHA_Bytes(&s2, intermediate, 20);
  SHA_Final(&s2, output);
} /* hmac_sha1_simple */
