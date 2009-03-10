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
 * PuTTY is copyright 1997-2001 Simon Tatham.
 *
 * Portions copyright Robert de Bath, Joris van Rantwijk, Delian
 * Delchev, Andreas Schultz, Jeroen Massar, Wez Furlong, Nicolas Barry,
 * Justin Bradford, and CORE SDI S.A.
 *
 * Permission is hereby granted, free of charge, to any person
 * obtaining a copy of this software and associated documentation files
 * (the "Software"), to deal in the Software without restriction,
 * including without limitation the rights to use, copy, modify, merge,
 * publish, distribute, sublicense, and/or sell copies of the Software,
 * and to permit persons to whom the Software is furnished to do so,
 * subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT.  IN NO EVENT SHALL THE COPYRIGHT HOLDERS BE LIABLE
 * FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
 * CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 *
 */


/*
 * RSA implementation just sufficient for ssh client-side
 * initialisation step
 *
 * Rewritten for more speed by Joris van Rantwijk, Jun 1999.
 */

/**********************************************************************************************
 * Change Log
 * ==========
 *
 * Date (MM/DD/YY)  Author      Description       
 * ---------------  ----------- --------------------------------------------------------------
 * 04/21/05         sentax      Upgraded source to version 3
 * 07/27/03         th2tran     Adapted from PuTTY project for used by the ActiveLock project.
 * 09/14/03         th2tran     Fixed bug in byte count calculation for key blob generation that
 *                              resulted in crashing when the blobs are used to recreate the key.
 * 04/10/04         th2tran     Fixed bug in byte count calculation that resulted in signature
 *                              verification mismatch.
 * 03/13/06         J.D.M.      Ported to C++
 *
 **********************************************************************************************/

#include <windows.h>
#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <assert.h>

#include "alcrypto3.h"
#include "bignum.h"
#include "md5.h"
#include "rsa.h"
#include "misc.h"
#include "memory.h"
#include "sha.h"

struct t_progress {
    int phases;
    char* extra;
};

#define GET_32BIT(cp) \
    (((unsigned long)(unsigned char)(cp)[0] << 24) | \
    ((unsigned long)(unsigned char)(cp)[1] << 16) | \
    ((unsigned long)(unsigned char)(cp)[2] << 8) | \
    ((unsigned long)(unsigned char)(cp)[3]))

#define PUT_32BIT(cp, value) { \
    (cp)[0] = (unsigned char)((value) >> 24); \
    (cp)[1] = (unsigned char)((value) >> 16); \
    (cp)[2] = (unsigned char)((value) >> 8); \
    (cp)[3] = (unsigned char)(value); }

/*
* This is the magic ASN.1/DER prefix that goes in the decoded
* signature, between the string of FFs and the actual SHA hash
* value. The meaning of it is:
*
* 00 -- this marks the end of the FFs; not part of the ASN.1 bit itself
*
* 30 21 -- a constructed SEQUENCE of length 0x21
*    30 09 -- a constructed sub-SEQUENCE of length 9
*       06 05 -- an object identifier, length 5
*          2B 0E 03 02 1A -- object id { 1 3 14 3 2 26 }
*                            (the 1,3 comes from 0x2B = 43 = 40*1+3)
*       05 00 -- NULL
*    04 14 -- a primitive OCTET STRING of length 0x14
*       [0x14 bytes of hash data follows]
*
* The object id in the middle there is listed as `id-sha1' in
* ftp://ftp.rsasecurity.com/pub/pkcs/pkcs-1/pkcs-1v2-1d2.asn (the
* ASN module for PKCS #1) and its expanded form is as follows:
*
* id-sha1                OBJECT IDENTIFIER ::= {
*    iso(1) identified-organization(3) oiw(14) secsig(3)
*    algorithms(2) 26 }
*/
static unsigned char asn1_weird_stuff[] = {
  0x00, 0x30, 0x21, 0x30, 0x09, 0x06, 0x05, 0x2B,
  0x0E, 0x03, 0x02, 0x1A, 0x05, 0x00, 0x04, 0x14,
};

#define ASN1_LEN ( (int) sizeof(asn1_weird_stuff) )

static int           makekey(unsigned char *data, struct RSAKey *result, unsigned char **keystr, int order);
static int           makeprivate(unsigned char *data, struct RSAKey *result);
static void          rsaencrypt(int type, unsigned char *data, int *length, struct RSAKey *key);
static Bignum        rsadecrypt(int type, Bignum input, struct RSAKey *key);
static int           rsastr_len(struct RSAKey *key);
static void          rsastr_fmt(char *str, struct RSAKey *key);
static void          rsa_fingerprint(char *str, int len, struct RSAKey *key);
static int           rsa_verify(struct RSAKey *key);
static unsigned char *rsa_public_blob(struct RSAKey *key, int *len);
static int           rsa_public_blob_len(void *data);
static void          freersakey(struct RSAKey *key);
static void          getstring(char **data, int *datalen, char **p, int *length);
static Bignum        getmp(char **data, int *datalen);
static void          *rsa2_newkey(unsigned char *data, int len);
static void          rsa2_freekey(void *key);
static char          *rsa2_fmtkey(void *key);
static void          rsa2_public_blob_internal(void *key, unsigned char *blob);
static void          base64_encode_blob(unsigned char *in, unsigned char *out, int blobLen);
static void          base64_decode_blob(unsigned char *in, unsigned char *out, int len);
static void          rsa2_public_blob(void *key, unsigned char *blob);
static int           rsa2_public_blob_len(void *key);
static void          rsa2_private_blob_internal(void *key, unsigned char *blob);
static void          rsa2_private_blob(void *key, unsigned char *blob);
static int           rsa2_private_blob_len(void *key);
static void          *rsa2_createkey_internal(unsigned char *pub_blob, int pub_len, unsigned char *priv_blob, int priv_len);
static void          rsa2_createkey(unsigned char *pub_blob, int pub_len, unsigned char *priv_blob, int priv_len, struct RSAKey *key);
static void          *rsa2_openssh_createkey(unsigned char **blob, int *len);
static char          *rsa2_fingerprint(void *key);
static int           rsa2_verifysig(void *key, char *sig, int siglen, char *data, int datalen);
static unsigned char *rsa2_sign(void *key, char *data, int datalen, int *siglen);

/*---------------------------------------------------------------------------*/
/* makekey                                                                   */
/*---------------------------------------------------------------------------*/
static int makekey(unsigned char *data, struct RSAKey *result,
                   unsigned char **keystr, int order)
{
  unsigned char *p = data;
  int           i;
  
  if (result) {
    result->bits = 0;
    for (i = 0; i < 4; i++) {
      result->bits = (result->bits << 8) + *p++;
    }
  }
  else {
    p += 4;
  }
  
  /* order=0 means exponent then modulus (the keys sent by the */
  /* server). order=1 means modulus then exponent (the keys    */
  /* stored in a keyfile).                                     */
  if (order == 0) {
    p += ssh1_read_bignum(p, result ? &result->exponent : NULL);
  }
  if (result) {
    result->bytes = (((p[0] << 8) + p[1]) + 7) / 8;
  }
  if (keystr) {
    *keystr = p + 2;
  }
  p += ssh1_read_bignum(p, result ? &result->modulus : NULL);
  if (order == 1) {
    p += ssh1_read_bignum(p, result ? &result->exponent : NULL);
  }
  
  return p - data;
} /* makekey */

/*---------------------------------------------------------------------------*/
/* makeprivate                                                               */
/*---------------------------------------------------------------------------*/
static int makeprivate(unsigned char *data, struct RSAKey *result)
{
  return ssh1_read_bignum(data, &result->private_exponent);
} /* makeprivate */

/*---------------------------------------------------------------------------*/
/* rsaencrypt                                                                */
/*---------------------------------------------------------------------------*/
static void rsaencrypt(int type, unsigned char *data, int *length, struct RSAKey *key)
{
  Bignum        b1, b2;
  int           i;
  unsigned char *p;

#ifdef DONT_USE /* commented out the random byte padding */
  memmove(data + key->bytes - length, data, length);
  data[0] = 0;
  data[1] = 2;
  
  for (i = 2; i < key->bytes - length - 1; i++) {
    do {
      data[i] = random_byte();
    } while (data[i] == 0);
  }
#endif
  /*data[key->bytes - *length - 1] = '\0';*/
  data[*length] = '\0';
  
  b1 = bignum_from_bytes(data, *length+1);
#ifdef _DEBUG
  /*diagbn("un-encrypted bn: ", b1);*/
#endif
  if (type == 0) {
    /* public encrypt */
    b2 = modpow(b1, key->exponent, key->modulus);
  }
  else { /* private encrypt */
    b2 = modpow(b1, key->private_exponent, key->modulus);
  }

#ifdef _DEBUG
  /*diagbn("encrypted bn: ", b2);*/
  /*debug(("encrypted bn bitcount: %d\n", bignum_bitcount(b2)));*/
#endif
  
  p = data;
  for (i = key->bytes; i--;) {
    *p++ = bignum_byte(b2, i);
  }
  
  /* calculate encrypted length */
  *length = (bignum_bitcount(b2) + 8)/8;
  
  freebn(b1);
  freebn(b2);
} /* rsaencrypt */

/*---------------------------------------------------------------------------*/
/* rsadecrypt                                                                */
/*---------------------------------------------------------------------------*/
static Bignum rsadecrypt(int type, Bignum input, struct RSAKey *key)
{
  Bignum ret;
#ifdef _DEBUG
  /*diagbn("encrypted input: ", input);*/
  /*debug(("encrypted input bitcount: %d\n", bignum_bitcount(input)));*/
#endif
  if (type == 0) {
    /* public decrypt */
    ret = modpow(input, key->exponent, key->modulus);
  }
  else {
    ret = modpow(input, key->private_exponent, key->modulus);
  }
#ifdef _DEBUG
  /*diagbn("decrypted bn: ", ret);*/
#endif

  return ret;
} /* rsadecrypt */

/*---------------------------------------------------------------------------*/
/* rsastr_len                                                                */
/*---------------------------------------------------------------------------*/
static int rsastr_len(struct RSAKey *key)
{
  Bignum md, ex;
  int    mdlen, exlen;
  
  md = key->modulus;
  ex = key->exponent;
  mdlen = (bignum_bitcount(md) + 15) / 16;
  exlen = (bignum_bitcount(ex) + 15) / 16;

  return 4 * (mdlen + exlen) + 20;
} /* rsastr_len */

/*---------------------------------------------------------------------------*/
/* rsastr_fmt                                                                */
/*---------------------------------------------------------------------------*/
static void rsastr_fmt(char *str, struct RSAKey *key)
{
  Bignum            md, ex;
  int               len = 0, i, nibbles;
  static const char hex[] = "0123456789abcdef";
  
  md = key->modulus;
  ex = key->exponent;
  
  len += sprintf(str + len, "0x");
  
  nibbles = (3 + bignum_bitcount(ex)) / 4;
  if (nibbles < 1) {
    nibbles = 1;
  }
  for (i = nibbles; i--;) {
    str[len++] = hex[(bignum_byte(ex, i / 2) >> (4 * (i % 2))) & 0xF];
  }
  
  len += sprintf(str + len, ",0x");
  
  nibbles = (3 + bignum_bitcount(md)) / 4;
  if (nibbles < 1) {
    nibbles = 1;
  }
  for (i = nibbles; i--;) {
    str[len++] = hex[(bignum_byte(md, i / 2) >> (4 * (i % 2))) & 0xF];
  }
  
  str[len] = '\0';
} /* rsastr_fmt */

/*---------------------------------------------------------------------------*/
/* rsa_fingerprint                                                           */
/*---------------------------------------------------------------------------*/
/* Generate a fingerprint string for the key. Compatible with the            */
/* OpenSSH fingerprint code.                                                 */
/*---------------------------------------------------------------------------*/
static void rsa_fingerprint(char *str, int len, struct RSAKey *key)
{
  struct MD5Context md5c;
  unsigned char     digest[16];
  char              buffer[16 * 3 + 40];
  int               numlen, slen, i;
  
  MD5Init(&md5c);
  numlen = ssh1_bignum_length(key->modulus) - 2;
  for (i = numlen; i--;) {
    unsigned char c = bignum_byte(key->modulus, i);

    MD5Update(&md5c, &c, 1);
  }
  numlen = ssh1_bignum_length(key->exponent) - 2;
  for (i = numlen; i--;) {
    unsigned char c = bignum_byte(key->exponent, i);

    MD5Update(&md5c, &c, 1);
  }
  MD5Final(digest, &md5c);
  
  sprintf(buffer, "%d ", bignum_bitcount(key->modulus));
  for (i = 0; i < 16; i++) {
    sprintf(buffer + strlen(buffer), "%s%02x", i ? ":" : "", digest[i]);
  }
  strncpy(str, buffer, len);
  str[len - 1] = '\0';
  slen = strlen(str);
  if (key->comment && slen < len - 1) {
    str[slen] = ' ';
    strncpy(str + slen + 1, key->comment, len - slen - 1);
    str[len - 1] = '\0';
  }
} /* rsa_fingerprint */

/*---------------------------------------------------------------------------*/
/* rsa_verify                                                                */
/*---------------------------------------------------------------------------*/
/* Verify that the public data in an RSA key matches the private             */
/* data. We also check the private data itself: we ensure that p > q         */
/* and that iqmp really is the inverse of q mod p.                           */
/*---------------------------------------------------------------------------*/
static int rsa_verify(struct RSAKey *key)
{
  Bignum n, ed, pm1, qm1;
  int    cmp;
  
  /* n must equal pq. */
  n = bigmul(key->p, key->q);
  cmp = bignum_cmp(n, key->modulus);
  freebn(n);
  if (cmp != 0) {
    return 0;
  }
  
  /* e * d must be congruent to 1, modulo (p-1) and modulo (q-1). */
  pm1 = copybn(key->p);
  decbn(pm1);
  ed = modmul(key->exponent, key->private_exponent, pm1);
  cmp = bignum_cmp(ed, One);
  sfree(ed);
  if (cmp != 0) {
    return 0;
  }
  
  qm1 = copybn(key->q);
  decbn(qm1);
  ed = modmul(key->exponent, key->private_exponent, qm1);
  cmp = bignum_cmp(ed, One);
  sfree(ed);
  if (cmp != 0) {
    return 0;
  }
  
  /* Ensure p > q. */
  if (bignum_cmp(key->p, key->q) <= 0) {
    return 0;
  }
  
  /* Ensure iqmp * q is congruent to 1, modulo p. */
  n = modmul(key->iqmp, key->q, key->p);
  cmp = bignum_cmp(n, One);
  sfree(n);
  if (cmp != 0) {
    return 0;
  }
  
  return 1;
} /* rsa_verify */

/*---------------------------------------------------------------------------*/
/* rsa_public_blob                                                           */
/*---------------------------------------------------------------------------*/
/* Public key blob as used by Pageant: exponent before modulus.              */
/*---------------------------------------------------------------------------*/
static unsigned char *rsa_public_blob(struct RSAKey *key, int *len)
{
  int           length, pos;
  unsigned char *ret;
  
  length = (ssh1_bignum_length(key->modulus) + ssh1_bignum_length(key->exponent) + 4);
  ret = (unsigned char*)smalloc(length);
  
  PUT_32BIT(ret, bignum_bitcount(key->modulus));
  pos = 4;
  pos += ssh1_write_bignum(ret + pos, key->exponent);
  pos += ssh1_write_bignum(ret + pos, key->modulus);
  
  *len = length;

  return ret;
} /* rsa_public_blob */

/*---------------------------------------------------------------------------*/
/* rsa_public_blob_len                                                       */
/*---------------------------------------------------------------------------*/
/* Given a public blob, determine its length.                                */
/*---------------------------------------------------------------------------*/
static int rsa_public_blob_len(void *data)
{
  unsigned char *p = (unsigned char *)data;
  
  p += 4;                         /* length word */
  p += ssh1_read_bignum(p, NULL); /* exponent */
  p += ssh1_read_bignum(p, NULL); /* modulus */
  
  return p - (unsigned char *)data;
} /* rsa_public_blob_len */

/*---------------------------------------------------------------------------*/
/* freersakey                                                                */
/*---------------------------------------------------------------------------*/
static void freersakey(struct RSAKey *key)
{
  if (key->modulus) {
    freebn(key->modulus);
  }
  if (key->exponent) {
    freebn(key->exponent);
  }
  if (key->private_exponent) {
    freebn(key->private_exponent);
  }
  if (key->comment) {
    sfree(key->comment);
  }
} /* freersakey */

/*---------------------------------------------------------------------------*/
/* Implementation of the ssh-rsa signing key type.                           */
/*---------------------------------------------------------------------------*/

/*---------------------------------------------------------------------------*/
/* getstring                                                                 */
/*---------------------------------------------------------------------------*/
static void getstring(char **data, int *datalen, char **p, int *length)
{
  *p = NULL;
  if (*datalen < 4) {
    return;
  }
  *length = GET_32BIT(*data);
  if(*length<0) {
    throw RETVAL_ON_ERROR;
  }
  *datalen -= 4;
  *data += 4;
  if (*datalen < *length) {
    return;
  }
  *p = *data;
  *data += *length;
  *datalen -= *length;
} /* getstring */

/*---------------------------------------------------------------------------*/
/* getmp                                                                     */
/*---------------------------------------------------------------------------*/
static Bignum getmp(char **data, int *datalen)
{
  char   *p;
  int    length;
  Bignum b;
  
  getstring(data, datalen, &p, &length);
  if (!p) {
    return NULL;
  }
  b = bignum_from_bytes((const unsigned char*)p, length);

  return b;
} /* getmp */

/*---------------------------------------------------------------------------*/
/* rsa2_newkey                                                               */
/*---------------------------------------------------------------------------*/
static void *rsa2_newkey(unsigned char *data, int len)
{
  char          *p;
  int           slen;
  struct RSAKey *rsa;
  
  rsa = (struct RSAKey*)smalloc(sizeof(struct RSAKey));
  if (!rsa) {
    return NULL;
  }
  getstring((char**)&data, &len, &p, &slen);
  
  if (!p || slen != 7 || memcmp(p, "ssh-rsa", 7)) {
    sfree(rsa);
    
    return NULL;
  }
  rsa->exponent = getmp((char**)&data, &len);
  rsa->modulus = getmp((char**)&data, &len);
  rsa->private_exponent = NULL;
  rsa->comment = NULL;
  rsa->bits = bignum_bitcount(rsa->modulus)+1;
  rsa->bytes = rsa->bits/8;

  return rsa;
} /* rsa2_newkey */

/*---------------------------------------------------------------------------*/
/* rsa2_freekey                                                              */
/*---------------------------------------------------------------------------*/
static void rsa2_freekey(void *key)
{
  struct RSAKey *rsa = (struct RSAKey *) key;

  freersakey(rsa);
  sfree(rsa);
} /* rsa2_freekey */

/*---------------------------------------------------------------------------*/
/* rsa2_fmtkey                                                               */
/*---------------------------------------------------------------------------*/
static char *rsa2_fmtkey(void *key)
{
  struct RSAKey *rsa = (struct RSAKey *) key;
  char          *p;
  int           len;
  
  len = rsastr_len(rsa);
  p = (char*)smalloc(len);
  rsastr_fmt(p, rsa);

  return p;
} /* rsa2_fmtkey */

/*---------------------------------------------------------------------------*/
/* rsa2_public_blob_internal                                                 */
/*---------------------------------------------------------------------------*/
static void rsa2_public_blob_internal(void *key, unsigned char *blob)
{
  struct RSAKey *rsa = (struct RSAKey *) key;
  int           elen, mlen, bloblen;
  int           i;
  unsigned char *p;
  
  elen = (bignum_bitcount(rsa->exponent) + 7) / 8;
  mlen = (bignum_bitcount(rsa->modulus) + 7) / 8;
  
  /* string "ssh-rsa", mpint exp, mpint mod. Total 19+elen+mlen. */
  /* (three length fields, 12+7=19).                             */
  bloblen = 19 + elen + mlen;
  
  p = blob;
  PUT_32BIT(p, 7);
  p += 4;
  memcpy(p, "ssh-rsa", 7);
  p += 7;
  PUT_32BIT(p, elen);
  p += 4;
  for (i = elen; i--;) {
    *p++ = bignum_byte(rsa->exponent, i);
  }
  PUT_32BIT(p, mlen);
  p += 4;
  for (i = mlen; i--;) {
    *p++ = bignum_byte(rsa->modulus, i);
  }
  assert(p == blob + bloblen);
} /* rsa2_public_blob_internal */

/*---------------------------------------------------------------------------*/
/* base64_encode_blob                                                        */
/*---------------------------------------------------------------------------*/
static void base64_encode_blob(unsigned char *in, unsigned char *out, int blobLen)
{
  char *p;
  int  i;

  i = 0;
  p = (char *)out;
  while (i < blobLen) {
    int n = (blobLen - i < 3 ? blobLen - i : 3);

    base64_encode_atom(in + i, n, p);
    i += n;
    p += 4;
  }
} /* base64_encode_blob */

/*---------------------------------------------------------------------------*/
/* base64_decode_blob                                                        */
/*---------------------------------------------------------------------------*/
static void base64_decode_blob(unsigned char *in, unsigned char *out, int len)
{
  int i=0, j, k;
  unsigned char *blob;

  blob = out;
  for (j = 0; j < len; j += 4) {
    k = base64_decode_atom((char*)(in + j), blob + i);
    i += k;
    if (!k) {
      return; /* invalid */
    }
  }
} /* base64_decode_blob */

/*---------------------------------------------------------------------------*/
/* rsa2_public_blob                                                          */
/*---------------------------------------------------------------------------*/
static void rsa2_public_blob(void *key, unsigned char *blob)
{
  int           blobLen;
  unsigned char *buffer; /* holds the public key blob */
  unsigned char *buffer2;

  /* calculate the blob length */
  blobLen = rsa2_public_blob_len(key);
  
  buffer = (unsigned char *)smalloc(blobLen);
  buffer2 = (unsigned char *)smalloc(blobLen);
  rsa2_public_blob_internal(key, buffer);
  
  /* base-64 encode the blob */
  base64_encode_blob(buffer, blob, blobLen);
  base64_decode_blob(blob, buffer2, blobLen);
  sfree(buffer);
  sfree(buffer2);
} /* rsa2_public_blob */

/*---------------------------------------------------------------------------*/
/* rsa2_public_blob_len                                                      */
/*---------------------------------------------------------------------------*/
static int rsa2_public_blob_len(void *key)
{
  struct RSAKey *rsa = (struct RSAKey *) key;
  int           elen, mlen, bloblen;
  
  elen = (bignum_bitcount(rsa->exponent) + 7) / 8;
  mlen = (bignum_bitcount(rsa->modulus) + 7) / 8;
  
  /* string "ssh-rsa", mpint exp, mpint mod. Total 19+elen+mlen. */
  /* (three length fields, 12+7=19).                             */
  bloblen = 19 + elen + mlen;

  return bloblen;
} /* rsa2_public_blob_len */

/*---------------------------------------------------------------------------*/
/* rsa2_private_blob_internal                                                */
/*---------------------------------------------------------------------------*/
static void rsa2_private_blob_internal(void *key, unsigned char *blob)
{
  struct RSAKey *rsa = (struct RSAKey *) key;
  int           dlen, plen, qlen, ulen, bloblen;
  int           i;
  unsigned char *p;
 
  dlen = (bignum_bitcount(rsa->private_exponent) + 8) / 8;
  plen = (bignum_bitcount(rsa->p) + 7) / 8;
  qlen = (bignum_bitcount(rsa->q) + 7) / 8;
  ulen = (bignum_bitcount(rsa->iqmp) + 7) / 8;
  
  /* mpint private_exp, mpint p, mpint q, mpint iqmp. Total 16 + sum of lengths. */
  bloblen = 16 + dlen + plen + qlen + ulen;
  
  p = blob;
  PUT_32BIT(p, dlen);
  p += 4;
  for (i = dlen; i--;) {
    *p++ = bignum_byte(rsa->private_exponent, i);
  }
  PUT_32BIT(p, plen);
  p += 4;
  for (i = plen; i--;) {
    *p++ = bignum_byte(rsa->p, i);
  }
  PUT_32BIT(p, qlen);
  p += 4;
  for (i = qlen; i--;) {
    *p++ = bignum_byte(rsa->q, i);
  }
  PUT_32BIT(p, ulen);
  p += 4;
  for (i = ulen; i--;) {
    *p++ = bignum_byte(rsa->iqmp, i);
  }
  assert(p == blob + bloblen);
} /* rsa2_private_blob_internal */

/*---------------------------------------------------------------------------*/
/* rsa2_private_blob                                                         */
/*---------------------------------------------------------------------------*/
static void rsa2_private_blob(void *key, unsigned char *blob)
{
  int           blobLen;
  unsigned char *buffer; /* holds the private key blob */

  /* calculate the blob length */
  blobLen = rsa2_private_blob_len(key);
  
  buffer = (unsigned char *)smalloc(blobLen);
  rsa2_private_blob_internal(key, buffer);
  
  /* base-64 encode the blob */
  base64_encode_blob(buffer, blob, blobLen);
  sfree(buffer);
} /* rsa2_private_blob */

/*---------------------------------------------------------------------------*/
/* rsa2_private_blob_len                                                     */
/*---------------------------------------------------------------------------*/
static int rsa2_private_blob_len(void *key)
{
  struct RSAKey *rsa = (struct RSAKey *) key;
  int           dlen, plen, qlen, ulen, bloblen;
  
  dlen = (bignum_bitcount(rsa->private_exponent) + 7) / 8;
  plen = (bignum_bitcount(rsa->p) + 7) / 8;
  qlen = (bignum_bitcount(rsa->q) + 7) / 8;
  ulen = (bignum_bitcount(rsa->iqmp) + 7) / 8;
  
  /* mpint private_exp, mpint p, mpint q, mpint iqmp. Total 16 + sum of lengths. */
  bloblen = 16 + dlen + plen + qlen + ulen;

  return bloblen;
} /* rsa2_private_blob_len */

/*---------------------------------------------------------------------------*/
/* rsa2_createkey_internal                                                   */
/*---------------------------------------------------------------------------*/
static void *rsa2_createkey_internal(unsigned char *pub_blob, int pub_len,
                                     unsigned char *priv_blob, int priv_len)
{
  struct RSAKey *rsa;
  unsigned char *pb = priv_blob;
  
  rsa = (struct RSAKey*)rsa2_newkey(pub_blob, pub_len);
#ifdef _DEBUG
  /*diagbn("exponent: ", rsa->exponent);*/
  /*diagbn("modulus: ", rsa->modulus);*/
#endif
  if (pb != NULL) {
    rsa->private_exponent = getmp((char**)&pb, &priv_len);
  }

#ifdef DONT_USE
  rsa->p = getmp(&pb, &priv_len);
  rsa->q = getmp(&pb, &priv_len);
  rsa->iqmp = getmp(&pb, &priv_len);
  
  if (!rsa_verify(rsa)) {
    rsa2_freekey(rsa);

    return NULL;
  }
#endif

  return rsa;
} /* rsa2_createkey_internal */

/*---------------------------------------------------------------------------*/
/* rsa2_createkey                                                            */
/*---------------------------------------------------------------------------*/
static void rsa2_createkey(unsigned char *pub_blob, int pub_len,
                           unsigned char *priv_blob, int priv_len, struct RSAKey *key)
{
  /* base64-decode the blobs */
  unsigned char *pub_blob_decoded = NULL, *priv_blob_decoded = NULL; /* holds the private key blob */
  int           pub_decoded_len, priv_decoded_len=0;
  struct RSAKey *key2;

  /* calculate the blob length */
  /* encoded_length = 4 * ((decoded_length + 2) / 3) */
  if (pub_blob != NULL) {
    pub_decoded_len = (pub_len * 3)/4 - 2;
    pub_blob_decoded = (unsigned char *)smalloc(pub_decoded_len);
    /* decode the blob */
    base64_decode_blob(pub_blob, pub_blob_decoded, pub_len);
  }
  if (priv_blob != NULL) {
    priv_decoded_len = (priv_len * 3)/4 - 2;
    priv_blob_decoded = (unsigned char *)smalloc(priv_decoded_len);
    /* decode the blob */
    base64_decode_blob(priv_blob, priv_blob_decoded, priv_decoded_len);
  }
  
  key2 = (struct RSAKey *)rsa2_createkey_internal(pub_blob_decoded, pub_decoded_len, priv_blob_decoded, priv_decoded_len);
  /* TODO CHECK COMPILER WARNING LEVEL4: 'pub_decoded_len' may be used without having been initialized */
  /* TODO CHECK COMPILER WARNING LEVEL4: 'priv_decoded_len' may be used without having been initialized */
  memcpy((void *)key, (void *)key2, sizeof(struct RSAKey));
  /* release temp memory */
  sfree((void *)key2);
  sfree(pub_blob_decoded);
  sfree(priv_blob_decoded);
} /* rsa2_createkey */

/*---------------------------------------------------------------------------*/
/* rsa2_openssh_createkey                                                    */
/*---------------------------------------------------------------------------*/
static void *rsa2_openssh_createkey(unsigned char **blob, int *len)
{
  char          **b = (char **) blob;
  struct RSAKey *rsa;
  
  rsa = (struct RSAKey*)smalloc(sizeof(struct RSAKey));
  if (!rsa) {
    return NULL;
  }
  rsa->comment = NULL;
  
  rsa->modulus = getmp(b, len);
  rsa->exponent = getmp(b, len);
  rsa->private_exponent = getmp(b, len);
  rsa->iqmp = getmp(b, len);
  rsa->p = getmp(b, len);
  rsa->q = getmp(b, len);
  
  if (!rsa->modulus || !rsa->exponent || !rsa->private_exponent ||
    !rsa->iqmp || !rsa->p || !rsa->q) {
    sfree(rsa->modulus);
    sfree(rsa->exponent);
    sfree(rsa->private_exponent);
    sfree(rsa->iqmp);
    sfree(rsa->p);
    sfree(rsa->q);
    sfree(rsa);

    return NULL;
  }
  
  return rsa;
} /* rsa2_openssh_createkey */

/*---------------------------------------------------------------------------*/
/* rsa2_fingerprint                                                          */
/*---------------------------------------------------------------------------*/
static char *rsa2_fingerprint(void *key)
{
  struct RSAKey *rsa = (struct RSAKey *) key;
  struct MD5Context md5c;
  unsigned char digest[16], lenbuf[4];
  char buffer[16 * 3 + 40];
  char *ret;
  int numlen, i;
  
  MD5Init(&md5c);
  MD5Update(&md5c, (const unsigned char*)"\0\0\0\7ssh-rsa", 11);
  
/* start definition ADD_BIGNUM */
#define ADD_BIGNUM(bignum) \
  numlen = (bignum_bitcount(bignum)+8)/8; \
  PUT_32BIT(lenbuf, numlen); \
  MD5Update(&md5c, lenbuf, 4); \
  for (i = numlen; i-- ;) { \
    unsigned char c = bignum_byte(bignum, i); \
    \
    MD5Update(&md5c, &c, 1); \
  }
/* end definition ADD_BIGNUM */

  ADD_BIGNUM(rsa->exponent);
  ADD_BIGNUM(rsa->modulus);
  
  MD5Final(digest, &md5c);
  
  sprintf(buffer, "ssh-rsa %d ", bignum_bitcount(rsa->modulus));
  for (i = 0; i < 16; i++) {
    sprintf(buffer + strlen(buffer), "%s%02x", i ? ":" : "", digest[i]);
  }
  ret = (char*)smalloc(strlen(buffer) + 1);
  if (ret) {
    strcpy(ret, buffer);
  }

  return ret;
} /* rsa2_fingerprint */

/*---------------------------------------------------------------------------*/
/* rsa2_verifysig                                                            */
/*---------------------------------------------------------------------------*/
/* Verify a signature.                                                       */
/* Returns 0 if verification passes; 1 otherwise.                            */
/*---------------------------------------------------------------------------*/
static int rsa2_verifysig(void *key, char *sig, int siglen,
                          char *data, int datalen)
{
  struct RSAKey *rsa = (struct RSAKey *) key;
  Bignum        in, out;
  char          *p;
  int           slen;
  int           bytes, i, j, ret;
  unsigned char hash[20];
  
  getstring(&sig, &siglen, &p, &slen);
  if (!p || slen != 7 || memcmp(p, "ssh-rsa", 7)) {
    return 1;
  }
  in = getmp(&sig, &siglen);
  out = modpow(in, rsa->exponent, rsa->modulus);
  freebn(in);
  
  ret = 1;
  
  bytes = (bignum_bitcount(rsa->modulus) + 7)/ 8;
  /* Top (partial) byte should be zero. */
  if (bignum_byte(out, bytes - 1) != 0) {
    goto exit_label;
  }
  /* First whole byte should be 1. */
  if (bignum_byte(out, bytes - 2) != 1) {
    goto exit_label;
  }
  /* Most of the rest should be FF. */
  for (i = bytes - 3; i >= 20 + ASN1_LEN; i--) {
    if (bignum_byte(out, i) != 0xFF) {
      goto exit_label;
    }
  }
  /* Then we expect to see the asn1_weird_stuff. */
  for (i = 20 + ASN1_LEN - 1, j = 0; i >= 20; i--, j++) {
    if (bignum_byte(out, i) != asn1_weird_stuff[j]) {
      goto exit_label;
    }
  }
  /* Finally, we expect to see the SHA-1 hash of the signed data. */
  SHA_Simple(data, datalen, hash);
  for (i = 19, j = 0; i >= 0; i--, j++) {
    if (bignum_byte(out, i) != hash[j]) {
      goto exit_label;
    }
  }
  /* all invalid possibilities exhausted. Return success! */
  ret = 0;

exit_label:
  freebn(out);

  return ret;
} /* rsa2_verifysig */

/*---------------------------------------------------------------------------*/
/* rsa2_sign                                                                 */
/*---------------------------------------------------------------------------*/
static unsigned char *rsa2_sign(void *key, char *data, int datalen, int *siglen)
{
  struct RSAKey *rsa = (struct RSAKey *) key;
  unsigned char *bytes;
  int           nbytes;
  unsigned char hash[20];
  Bignum        in, out;
  int           i, j;
  
  SHA_Simple(data, datalen, hash);
  
  nbytes = (bignum_bitcount(rsa->modulus) - 1) / 8;
  bytes = (unsigned char*)smalloc(nbytes);
  
  bytes[0] = 1;
  for (i = 1; i < nbytes - 20 - ASN1_LEN; i++) {
    bytes[i] = 0xFF;
  }
  for (i = nbytes - 20 - ASN1_LEN, j = 0; i < nbytes - 20; i++, j++) {
    bytes[i] = asn1_weird_stuff[j];
  }
  for (i = nbytes - 20, j = 0; i < nbytes; i++, j++) {
    bytes[i] = hash[j];
  }
  
  in = bignum_from_bytes(bytes, nbytes);
  sfree(bytes);
  
  out = modpow(in, rsa->private_exponent, rsa->modulus);
  freebn(in);
  
  nbytes = (bignum_bitcount(out) + 7) / 8;
  bytes = (unsigned char*)smalloc(4 + 7 + 4 + nbytes);
  PUT_32BIT(bytes, 7);
  memcpy(bytes + 4, "ssh-rsa", 7);
  PUT_32BIT(bytes + 4 + 7, nbytes);
  for (i = 0; i < nbytes; i++) {
    bytes[4 + 7 + 4 + i] = bignum_byte(out, nbytes - 1 - i);
  }
  freebn(out);
  
  *siglen = 4 + 7 + 4 + nbytes;

  return bytes;
} /* rsa2_sign */

/*---------------------------------------------------------------------------*/
/* AL_rsa_get_public_blob                                                    */
/*---------------------------------------------------------------------------*/
/* Retrieves the public key blob.                                            */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_get_public_blob(struct RSAKey *key, char *blob, int *len)
{
  unsigned char *public_blob;

  public_blob = rsa_public_blob(key, len);
  if (blob == NULL) {
    return 0;
  }
  memcpy(blob, public_blob, *len);
  sfree(public_blob);

  return 0;
} /* AL_rsa_get_public_blob */

/*---------------------------------------------------------------------------*/
/* AL_rsa_encrypt                                                            */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_encrypt(int type, unsigned char *data,
                              int *len, struct RSAKey *key)
{
  rsaencrypt(type, data, len, key);

  return 0;
} /* AL_rsa_encrypt */

/*---------------------------------------------------------------------------*/
/* AL_rsa_decrypt                                                            */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_decrypt(int type, unsigned char *data,
                              int *outlen, struct RSAKey *key)
{
  Bignum input, output;
    
  input = bignum_from_bytes(data, key->bytes);
  output = rsadecrypt(type, input, key);
  bignum_to_bytes(output, data);
  *outlen = strlen((const char*)data);
  freebn(input);
  freebn(output);

  return 0;
} /* AL_rsa_decrypt */

/*---------------------------------------------------------------------------*/
/* AL_rsa_public_key_blob                                                    */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_public_key_blob(struct RSAKey *key,
                                      unsigned char *pub_blob,
                                      int *len)
{
  int blobLen;

  if (pub_blob == NULL) {
    /* calculate the blob length */
    blobLen = rsa2_public_blob_len(key);
    /* calculate base-64 encoded length */
    *len = 4 * ((blobLen + 2) / 3);
    return 0;
  }
  rsa2_public_blob(key, pub_blob);

  return 0;
} /* AL_rsa_public_key_blob */

/*---------------------------------------------------------------------------*/
/* AL_rsa_private_key_blob                                                   */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_private_key_blob(struct RSAKey *key,
                                       unsigned char *priv_blob,
                                       int *len)
{
  int blobLen;

  if (priv_blob == NULL) {
    /* calculate the blob length */
    blobLen = rsa2_private_blob_len(key);
    /* calculate base-64 encoded length */
    *len = 4 * ((blobLen + 2) / 3);

    return 0;
  }
  rsa2_private_blob(key, priv_blob);

  return 0;
} /* AL_rsa_private_key_blob */

/*---------------------------------------------------------------------------*/
/* AL_rsa_createkey                                                          */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_createkey(unsigned char *pub_blob, int pub_len,
                                unsigned char *priv_blob, int priv_len,
                                struct RSAKey *key)
{
  if(pub_len%4||priv_len%4) {
    throw RETVAL_ON_ERROR;
  }

  rsa2_createkey(pub_blob, pub_len, priv_blob, priv_len, key);

  return 0;
} /* AL_rsa_createkey */

/*---------------------------------------------------------------------------*/
/* AL_rsa_freekey                                                            */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_freekey(struct RSAKey *key)
{
  freersakey(key);

  return 0;
} /* AL_rsa_freekey */

/*---------------------------------------------------------------------------*/
/* AL_rsa_sign                                                               */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_sign(struct RSAKey *key, char *data, int datalen,
                           char *sig, int *siglen)
{
  char *sigtemp;

  sigtemp = (char*)rsa2_sign(key, data, datalen, siglen);
  if (sig == NULL) {
    sfree(sigtemp);
    return 0;
  }
  memcpy(sig, sigtemp, *siglen);
  sig[*siglen] = '\0';
  sfree((void *)sigtemp);

  return 0;
} /* AL_rsa_sign */

/*---------------------------------------------------------------------------*/
/* AL_rsa_verifysig                                                          */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_verifysig(void *key, char *sig, int siglen,
                                char *data, int datalen)
{
  return rsa2_verifysig(key, sig, siglen, data, datalen);
} /* AL_rsa_verifysig */
