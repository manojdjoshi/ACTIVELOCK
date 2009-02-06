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

/* THIS FILE DEFINES THE COMPLETE INTERFACE TO ALCRYPTO */

#ifndef ALCRYPTO_ALCRYPTO3_H
#define ALCRYPTO_ALCRYPTO3_H

#ifdef __cplusplus
extern "C" {
#endif

// The following ifdef block is the standard way of creating macros which make exporting 
// from a DLL simpler. All files within this DLL are compiled with the ALCRYPTO_EXPORTS
// symbol defined on the command line. this symbol should not be defined on any project
// that uses this DLL. This way any other project whose source files include this file see 
// RSA_API functions as being imported from a DLL, wheras this DLL sees symbols
// defined with this macro as being exported.
#ifdef ALCRYPTO_EXPORTS
#define ALCRYPTO_API __declspec(dllexport) 
#else
#define ALCRYPTO_API __declspec(dllimport) 
#endif

typedef unsigned int word32;
typedef unsigned int uint32;

#ifndef FALSE
#define FALSE 0
#endif
#ifndef TRUE
#define TRUE 1
#endif

enum VBenum
{
  wf1 = 1,
  wf8 = 8,
  wf1e = 0x1E
};

typedef struct
{
  char  cArray[30];
  VBenum  lformat;
} mUDT;

typedef void (__stdcall *progfn_t) (void *param, int action, int phase, int progress);

#define REG_HOME "Software\\ActiveLock Software Group\\ActiveLock"

/*---------------------------------------------------------------------------*/
/* INTERFACE BELOW HERE: */
/* Functions all return RETVAL_ON_ERROR when an exception occurs */
#define RETVAL_ON_ERROR -999

ALCRYPTO_API LRESULT WINAPI rsa_get_public_blob(struct RSAKey *key, char *blob, int *len);
ALCRYPTO_API LRESULT WINAPI rsa_generate(struct RSAKey *key, int bits, progfn_t pfn, void *pfnparam);
ALCRYPTO_API LRESULT WINAPI rsa_generate2(struct RSAKey *key, int bits);
ALCRYPTO_API LRESULT WINAPI rsa_encrypt(int type, unsigned char *data, int *len, struct RSAKey *key);
ALCRYPTO_API LRESULT WINAPI rsa_decrypt(int type, unsigned char *data, int *outlen, struct RSAKey *key);
ALCRYPTO_API LRESULT WINAPI rsa_public_key_blob(struct RSAKey *key, unsigned char *pub_blob, int *len);
ALCRYPTO_API LRESULT WINAPI rsa_private_key_blob(struct RSAKey *key, unsigned char *priv_blob, int *len);
ALCRYPTO_API LRESULT WINAPI rsa_createkey(unsigned char *pub_blob, int pub_len, unsigned char *priv_blob, int priv_len, struct RSAKey *key);
ALCRYPTO_API LRESULT WINAPI rsa_sign(struct RSAKey *key, char *data, int datalen, char *sig, int *siglen);
ALCRYPTO_API LRESULT WINAPI rsa_verifysig(void *key, char *sig, int siglen, char *data, int datalen);
ALCRYPTO_API LRESULT WINAPI rsa_freekey(struct RSAKey *key);
ALCRYPTO_API LRESULT WINAPI md5_hash(unsigned char *in, int len, unsigned char *out);
ALCRYPTO_API LRESULT WINAPI getHardDriveFirmware (mUDT* pU);
ALCRYPTO_API LRESULT WINAPI InitALCrypto(short ShowMessageBoxOnError=FALSE, short ExitOnError=FALSE, short CatchExceptions=TRUE);

enum WIN_VERSION {
/* 00 */  WIN_UNKNOWN_VERSION
/* 01 */ ,WIN_31
/* 02 */ ,WIN_95
/* 03 */ ,WIN_98
/* 04 */ ,WIN_NT
/* 05 */ ,WIN_NT_WORKSTATION
/* 06 */ ,WIN_NT_SERVER
/* 07 */ ,WIN_2000
/* 08 */ ,WIN_2000_PROFESSIONAL
/* 09 */ ,WIN_2000_DOMAIN_CONTROLLER
/* 10 */ ,WIN_2000_SERVER
};
ALCRYPTO_API LRESULT WINAPI GetWindowsVersion();
/* INTERFACE ABOVE HERE: */

#ifdef __cplusplus
}
#endif

#endif
