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

/* 
 * Bignum routines for RSA and DH and stuff.
 */

#ifndef ALCRYPTO_BIGNUM_H
#define ALCRYPTO_BIGNUM_H

#ifdef __cplusplus
extern "C" {
#endif

typedef struct {
  unsigned long hi, lo;
} uint64, int64;

#if defined __GNUC__ && defined __i386__

typedef unsigned long BignumInt;
typedef unsigned long long BignumDblInt;
#define BIGNUM_INT_MASK  0xFFFFFFFFUL
#define BIGNUM_TOP_BIT   0x80000000UL
#define BIGNUM_INT_BITS  32
#define MUL_WORD(w1, w2) ((BignumDblInt)w1 * w2)
#define DIVMOD_WORD(q, r, hi, lo, w) \
        __asm__("div %2" : \
        "=d" (r), "=a" (q) : \
        "r" (w), "d" (hi), "a" (lo))

#else

typedef unsigned short BignumInt;
typedef unsigned long BignumDblInt;
#define BIGNUM_INT_MASK  0xFFFFU
#define BIGNUM_TOP_BIT   0x8000U
#define BIGNUM_INT_BITS  16
#define MUL_WORD(w1, w2) ((BignumDblInt)w1 * w2)
#define DIVMOD_WORD(q, r, hi, lo, w) \
        do { \
          BignumDblInt n = (((BignumDblInt)hi) << BIGNUM_INT_BITS) | lo; \
          q = n / w; \
          r = n % w; \
        } while (0)

#endif

#define BIGNUM_INT_BYTES (BIGNUM_INT_BITS / 8)

#define BIGNUM_INTERNAL

typedef BignumInt *Bignum;

extern Bignum Zero, One;

extern void   bn_restore_invariant(Bignum b);
extern Bignum copybn(Bignum orig);
extern void   freebn(Bignum b);
extern Bignum bn_power_2(int n);
extern Bignum modpow(Bignum base, Bignum exp, Bignum mod);
extern Bignum modmul(Bignum p, Bignum q, Bignum mod);
extern void   decbn(Bignum bn);
extern Bignum bignum_from_bytes(const unsigned char *data, int nbytes);
extern void   bignum_to_bytes(Bignum bn, unsigned char *data);
extern int    ssh1_read_bignum(const unsigned char *data, Bignum *result);
extern int    bignum_bitcount(Bignum bn);
extern int    ssh1_bignum_length(Bignum bn);
extern int    bignum_byte(Bignum bn, int i);
extern int    bignum_bit(Bignum bn, int i);
extern void   bignum_set_bit(Bignum bn, int bitnum, int value);
extern int    ssh1_write_bignum(void *data, Bignum bn);
extern int    bignum_cmp(Bignum a, Bignum b);
extern Bignum bignum_rshift(Bignum a, int shift);
extern Bignum bigmuladd(Bignum a, Bignum b, Bignum addend);
extern Bignum bigmul(Bignum a, Bignum b);
extern Bignum bignum_from_long(unsigned long nn);
extern Bignum bignum_add_long(Bignum number, unsigned long addendx);
extern unsigned short bignum_mod_short(Bignum number, unsigned short modulus);
#ifdef _DEBUG
extern void   diagbn(char *prefix, Bignum md);
#endif
extern Bignum modinv(Bignum number, Bignum modulus);

#ifdef __cplusplus
}
#endif

#endif
