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
 * RSA key generation.
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
#include "rsag.h"
#include "bignum.h"
#include "prime.h"
#include "rand.h"
#include "rsa.h"

#define RSA_EXPONENT 37          /* we like this prime */

extern void random_init(void); 

/*---------------------------------------------------------------------------*/
/* AL_rsa_generate                                                           */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_generate(struct RSAKey *key, int bits, progfn_t pfn, void *pfnparam)
{
  Bignum pm1, qm1, phi_n;

 /*
  * Set up the phase limits for the progress report. We do this
  * by passing minus the phase number.
  *
  * For prime generation: our initial filter finds things
  * coprime to everything below 2^16. Computing the product of
  * (p-1)/p for all prime p below 2^16 gives about 20.33; so
  * among B-bit integers, one in every 20.33 will get through
  * the initial filter to be a candidate prime.
  *
  * Meanwhile, we are searching for primes in the region of 2^B;
  * since pi(x) ~ x/log(x), when x is in the region of 2^B, the
  * prime density will be d/dx pi(x) ~ 1/log(B), i.e. about
  * 1/0.6931B. So the chance of any given candidate being prime
  * is 20.33/0.6931B, which is roughly 29.34 divided by B.
  *
  * So now we have this probability P, we're looking at an
  * exponential distribution with parameter P: we will manage in
  * one attempt with probability P, in two with probability
  * P(1-P), in three with probability P(1-P)^2, etc. The
  * probability that we have still not managed to find a prime
  * after N attempts is (1-P)^N.
  * 
  * We therefore inform the progress indicator of the number B
  * (29.34/B), so that it knows how much to increment by each
  * time. We do this in 16-bit fixed point, so 29.34 becomes
  * 0x1D.57C4.
  */
  pfn(pfnparam, PROGFN_PHASE_EXTENT, 1, 0x10000);
  pfn(pfnparam, PROGFN_EXP_PHASE, 1, -0x1D57C4 / (bits / 2));
  pfn(pfnparam, PROGFN_PHASE_EXTENT, 2, 0x10000);
  pfn(pfnparam, PROGFN_EXP_PHASE, 2, -0x1D57C4 / (bits - bits / 2));
  pfn(pfnparam, PROGFN_PHASE_EXTENT, 3, 0x4000);
  pfn(pfnparam, PROGFN_LIN_PHASE, 3, 5);
  pfn(pfnparam, PROGFN_READY, 0, 0);

  /* We don't generate e; we just use a standard one always. */
  key->exponent = bignum_from_long(RSA_EXPONENT);

 /*
  * Generate p and q: primes with combined length `bits', not
  * congruent to 1 modulo e. (Strictly speaking, we wanted (p-1)
  * and e to be coprime, and (q-1) and e to be coprime, but in
  * general that's slightly more fiddly to arrange. By choosing
  * a prime e, we can simplify the criterion.)
  */
  random_init();
  key->p = primegen(bits / 2, RSA_EXPONENT, 1, NULL, 1, pfn, pfnparam);
#ifdef _DEBUG
  /*diagbn("key p: ", key->p);*/
#endif
  key->q = primegen(bits - bits / 2, RSA_EXPONENT, 1, NULL, 2, pfn, pfnparam);
#ifdef _DEBUG
  /*diagbn("key q: ", key->q);*/
#endif

  /* Ensure p > q, by swapping them if not. */
  if (bignum_cmp(key->p, key->q) < 0) {
    Bignum t = key->p;
    key->p = key->q;
    key->q = t;
  }

 /*
  * Now we have p, q and e. All we need to do now is work out
  * the other helpful quantities: n=pq, d=e^-1 mod (p-1)(q-1),
  * and (q^-1 mod p).
  */
  pfn(pfnparam, PROGFN_PROGRESS, 3, 1);
  key->modulus = bigmul(key->p, key->q);
  pfn(pfnparam, PROGFN_PROGRESS, 3, 2);
  pm1 = copybn(key->p);
  decbn(pm1);
  qm1 = copybn(key->q);
  decbn(qm1);
  phi_n = bigmul(pm1, qm1);
  pfn(pfnparam, PROGFN_PROGRESS, 3, 3);
  freebn(pm1);
  freebn(qm1);
  key->private_exponent = modinv(key->exponent, phi_n);
  pfn(pfnparam, PROGFN_PROGRESS, 3, 4);
  key->iqmp = modinv(key->q, key->p);
  pfn(pfnparam, PROGFN_PROGRESS, 3, 5);
  key->bits = bits;
  key->bytes = bits/8;
#ifdef _DEBUG
  /*diagbn("key modulus: ", key->modulus);*/
  /*diagbn("key exponent: ", key->exponent);*/
  /*diagbn("key private exponent: ", key->private_exponent);*/
#endif
  /* Clean up temporary numbers. */
  freebn(phi_n);

  return 1;
} /* AL_rsa_generate */

/*---------------------------------------------------------------------------*/
/* AL_rsa_generate2                                                          */
/*---------------------------------------------------------------------------*/
/* This is the second version of rsa_generate without any progress updates   */
/* in the key generation utility, alugen                                     */
/* This function is needed for the VB.NET version of alugen                  */
/* ialkan 8-2-2005.                                                          */
/*---------------------------------------------------------------------------*/
LRESULT WINAPI AL_rsa_generate2(struct RSAKey *key, int bits)
{
  Bignum pm1, qm1, phi_n;

  /* We don't generate e; we just use a standard one always. */
  key->exponent = bignum_from_long(RSA_EXPONENT);

 /*
  * Generate p and q: primes with combined length `bits', not
  * congruent to 1 modulo e. (Strictly speaking, we wanted (p-1)
  * and e to be coprime, and (q-1) and e to be coprime, but in
  * general that's slightly more fiddly to arrange. By choosing
  * a prime e, we can simplify the criterion.)
  */
  random_init();
  key->p = primegen2(bits / 2, RSA_EXPONENT, 1, NULL, 1);
#ifdef _DEBUG
  /*diagbn("key p: ", key->p);*/
#endif
  key->q = primegen2(bits - bits / 2, RSA_EXPONENT, 1, NULL, 2);
#ifdef _DEBUG
  /*diagbn("key q: ", key->q);*/
#endif
    
  /* Ensure p > q, by swapping them if not. */
  if (bignum_cmp(key->p, key->q) < 0) {
    Bignum t = key->p;
    key->p = key->q;
    key->q = t;
  }

 /*
  * Now we have p, q and e. All we need to do now is work out
  * the other helpful quantities: n=pq, d=e^-1 mod (p-1)(q-1),
  * and (q^-1 mod p).
  */
  key->modulus = bigmul(key->p, key->q);
  pm1 = copybn(key->p);
  decbn(pm1);
  qm1 = copybn(key->q);
  decbn(qm1);
  phi_n = bigmul(pm1, qm1);
  freebn(pm1);
  freebn(qm1);
  key->private_exponent = modinv(key->exponent, phi_n);
  key->iqmp = modinv(key->q, key->p);
  key->bits = bits;
  key->bytes = bits/8;
#ifdef _DEBUG
  /*diagbn("key modulus: ", key->modulus);*/
  /*diagbn("key exponent: ", key->exponent);*/
  /*diagbn("key private exponent: ", key->private_exponent);*/
#endif
  /* Clean up temporary numbers. */
  freebn(phi_n);

  return 1;
} /* AL_rsa_generate2 */
