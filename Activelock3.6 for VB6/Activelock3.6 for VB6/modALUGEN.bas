Attribute VB_Name = "modALUGEN"
' This project is available from SVN on SourceForge.net under the main project, Activelock !
'
' ProjectPage: http://sourceforge.net/projects/activelock
' WebSite: http://www.activeLockSoftware.com
' DeveloperForums: http://forums.activelocksoftware.com
' ProjectManager: Ismail Alkan - http://activelocksoftware.com/simplemachinesforum/index.php?action=profile;u=1
' ProjectLicense: BSD Open License - http://www.opensource.org/licenses/bsd-license.php
' ProjectPurpose: Copy Protection, Software Locking, Anti Piracy
'
' //////////////////////////////////////////////////////////////////////////////////////////
' *   ActiveLock
' *   Copyright 1998-2002 Nelson Ferraz
' *   Copyright 2003-2009 The ActiveLock Software Group (ASG)
' *   All material is the property of the contributing authors.
' *
' *   Redistribution and use in source and binary forms, with or without
' *   modification, are permitted provided that the following conditions are
' *   met:
' *
' *     [o] Redistributions of source code must retain the above copyright
' *         notice, this list of conditions and the following disclaimer.
' *
' *     [o] Redistributions in binary form must reproduce the above
' *         copyright notice, this list of conditions and the following
' *         disclaimer in the documentation and/or other materials provided
' *         with the distribution.
' *
' *   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
' *   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
' *   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
' *   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
' *   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
' *   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
' *   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
' *   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' *   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' *   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' *   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
' *
'===============================================================================
' Name: modActiveLock
' Purpose: Module used by ALUGEN.
' Functions:
' Properties:
' Methods:
' Started: 04.21.2005
' Modified: 08.15.2005
'===============================================================================
' @author activelock-admins
' @version 3.0.0
' @date 20050421
'
' * ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////

' * ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' @history
' <pre>
' 08.15.03 - th2tran       - Created
' 05.13.05 - ialkan        - Added this module to activelock3.dll commented out duplicate entries
' </pre>

'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////
Option Explicit
Option Base 0

Public Const ACTIVELOCKSTRING As String = "Activelock3"

' RSA Key Structure
' @param bits   Key length in bits
' @param data   Key data
Type RSAKey  ' 36-byte structure
    bits As Long
    data(32) As Byte
End Type


' ALCrypto Removal
'' Generates an RSA key set
'' @param ptrKey RSA key structure
'' @param bits   key length in bits
'' @param pfn    TBD
'' @param pfnparam    TBD
'Public Declare Function rsa_generate Lib "ALCrypto3" (ptrKey As RSAKey, ByVal bits As Long, ByVal pfn As Long, ByVal pfnparam As Long) As Long
'
'
'' Generates an RSA key set without showing any progress
'' needed for the VB.NET build
'' @param ptrKey RSA key structure
'' @param bits   key length in bits
'Public Declare Function rsa_generate2 Lib "ALCrypto3" (ptrKey As RSAKey, ByVal bits As Long) As Long
'
'
'' Returns the public key blob for the specified key.
'' @param ptrKey RSA key structure
'' @param blob   [Output] Key bob to be returned
'' @param blobLen    Length of the key blob, in bytes
'Public Declare Function rsa_public_key_blob Lib "ALCrypto3" (ptrKey As RSAKey, ByVal blob As String, blobLen As Long) As Long
'
'
'' Returns the private key blob for the specified key.
'' @param ptrKey RSA key structure
'' @param blob   [Output] Key bob to be returned
'' @param blobLen    Length of the key blob, in bytes
'Public Declare Function rsa_private_key_blob Lib "ALCrypto3" (ptrKey As RSAKey, ByVal blob As String, blobLen As Long) As Long
'
'
'' Creates a new RSAKey from the specified key blobs.
'' @param pub_blob   Public key blob
'' @param pub_len    Length of public key blob, in bytes
'' @param priv_blob  Private key blob
'' @param priv_len   Length of private key blob, in bytes
'' @param ptrKey     [out] RSA key to be returned.
'Public Declare Function rsa_createkey Lib "ALCrypto3" (ByVal pub_blob As String, ByVal pub_len As Long, ByVal priv_blob As String, ByVal priv_len As Long, ptrKey As RSAKey) As Long
'
'' Release memory allocated by rsa_createkey() to store the key.
'' @param ptrKey     RSA key
'Public Declare Function rsa_freekey Lib "ALCrypto3" (ptrKey As RSAKey) As Long
'
'' Signs the data using the specified RSA private key.
'' @param ptrKey Key to be used for signing
'' @param data   Data to be signed
'' @param dLen   Data length
'' @param sig    [out] Signature
'' @param sLen   Signature length
'Public Declare Function rsa_sign Lib "ALCrypto3" (ByRef ptrKey As RSAKey, ByVal data As String, ByVal dLen As Long, ByVal sig As String, ByRef sLen As Long) As Long
''
'' Verifies an RSA signature.
'' @param ptrKey Key to be used for signing
'' @param sig    [out] Signature
'' @param sLen   Signature length
'' @param data   Data with which to verify
'' @param dLen   Data length
'Public Declare Function rsa_verifysig Lib "ALCrypto3" (ByRef ptrKey As RSAKey, ByVal sig As String, ByVal sLen As Long, ByVal data As String, ByVal dLen As Long) As Long

Public Type PhaseType
    exponential As Byte
    startpoint As Byte
    total As Byte
    param As Byte
    current As Byte
    N As Byte    ' if exponential */
    Mult As Byte             ' if linear */
End Type

Public Const MAXPHASE& = 5

Public Type ProgressType
    nphases As Long
    phases(0 To MAXPHASE - 1) As PhaseType
    total As Byte
    divisor As Byte
    range As Byte
    hwndProgbar As Long
End Type
