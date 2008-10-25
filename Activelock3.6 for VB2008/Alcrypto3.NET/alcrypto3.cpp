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

/* This source contains only wrapper functions. The dll should have it's interface clearly */
/* specified. This is the main reason of the existance of this source.                     */

#include <windows.h>

#include "alcrypto3.h"
#include "WindowsVersion.h"
#include "diskid32.h"
#include "md5.h"
#include "bignum.h"
#include "rsa.h"
#include "rsag.h"
#include "memory.h"

/*---------------------------------------------------------------------------*/
/* InitALCrypto                                                              */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI InitALCrypto(short ShowMessageBoxOnError, short ExitOnError, short CatchExceptions) {
  try {
    return ALInitALCrypto(ShowMessageBoxOnError, ExitOnError, CatchExceptions);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* InitALCrypto */

/*---------------------------------------------------------------------------*/
/* rsa_get_public_blob                                                       */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_get_public_blob(struct RSAKey *key, char *blob, int *len)
{
  try {
    return AL_rsa_get_public_blob(key, blob, len);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_get_public_blob */

/*---------------------------------------------------------------------------*/
/* rsa_generate                                                              */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_generate(struct RSAKey *key, int bits, progfn_t pfn, void *pfnparam)
{
  try {
    return AL_rsa_generate(key, bits, pfn, pfnparam);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_generate */

/*---------------------------------------------------------------------------*/
/* rsa_generate2                                                             */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_generate2(struct RSAKey *key, int bits)
{
  try {
    return AL_rsa_generate2(key, bits);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_generate2 */

/*---------------------------------------------------------------------------*/
/* rsa_encrypt                                                               */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_encrypt(int type, unsigned char *data, int *len, struct RSAKey *key)
{
  try {
    return AL_rsa_encrypt(type, data, len, key);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_encrypt */

/*---------------------------------------------------------------------------*/
/* rsa_decrypt                                                               */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_decrypt(int type, unsigned char *data, int *outlen, struct RSAKey *key)
{
  try {
    return AL_rsa_decrypt(type, data, outlen, key);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_decrypt */

/*---------------------------------------------------------------------------*/
/* rsa_public_key_blob                                                       */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_public_key_blob(struct RSAKey *key, unsigned char *pub_blob, int *len)
{
  try {
    return AL_rsa_public_key_blob(key, pub_blob, len);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_public_key_blob */

/*---------------------------------------------------------------------------*/
/* rsa_private_key_blob                                                      */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_private_key_blob(struct RSAKey *key, unsigned char *priv_blob, int *len)
{
  try {
    return AL_rsa_private_key_blob(key, priv_blob, len);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_private_key_blob */

/*---------------------------------------------------------------------------*/
/* rsa_createkey                                                             */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_createkey(unsigned char *pub_blob, int pub_len, unsigned char *priv_blob,
                                          int priv_len, struct RSAKey *key)
{
  try {
    return AL_rsa_createkey(pub_blob, pub_len, priv_blob, priv_len, key);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_createkey */

/*---------------------------------------------------------------------------*/
/* rsa_sign                                                                  */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_sign(struct RSAKey *key, char *data, int datalen, char *sig, int *siglen)
{
  try {
    return AL_rsa_sign(key, data, datalen, sig, siglen);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_sign */

/*---------------------------------------------------------------------------*/
/* rsa_verifysig                                                             */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_verifysig(void *key, char *sig, int siglen, char *data, int datalen)
{
  try {
    return AL_rsa_verifysig(key, sig, siglen, data, datalen);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_verifysig */

/*---------------------------------------------------------------------------*/
/* rsa_freekey                                                               */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI rsa_freekey(struct RSAKey *key)
{
  try {
    return AL_rsa_freekey(key);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* rsa_freekey */

/*---------------------------------------------------------------------------*/
/* md5_hash                                                                  */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI md5_hash(unsigned char *in, int len, unsigned char *out)
{
  try {
    return AL_md5_hash(in, len, out);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* md5_hash */

/*---------------------------------------------------------------------------*/
/* getHardDriveFirmware                                                      */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI getHardDriveFirmware(mUDT* pU)
{
  try {
    return ALGetHardDriveFirmware(pU);
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* getHardDriveFirmware */

/*---------------------------------------------------------------------------*/
/* GetWindowsVersion                                                         */
/*---------------------------------------------------------------------------*/
ALCRYPTO_API LRESULT WINAPI GetWindowsVersion()
{
  try {
    return ALGetWindowsVersion();
  }
  catch(...) {
    if(DoCatchExceptions==FALSE) {
      throw;
    }
    return RETVAL_ON_ERROR; /* Check on this return value in activelock dll */
  }
} /* GetWindowsVersion */

/*---------------------------------------------------------------------------*/
/* DllMain                                                                   */
/*---------------------------------------------------------------------------*/
BOOL APIENTRY DllMain(HANDLE hModule, DWORD ul_reason_for_call, LPVOID lpReserved)
{
  switch (ul_reason_for_call) {
    case DLL_PROCESS_ATTACH:
    case DLL_THREAD_ATTACH:
    case DLL_THREAD_DETACH:
    case DLL_PROCESS_DETACH:
      break;
  }

  return TRUE;
} /* DllMain */
