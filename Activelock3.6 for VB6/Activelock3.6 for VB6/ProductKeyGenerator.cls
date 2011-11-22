VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ProductKeyGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'  ///////////////////////////////////////////////////////////////////////
'  / Filename:  ProductKeyGeberator.cls                                  /
'  / Version:   2.0.0.1                                                  /
'  / Purpose:   Manages product encryption keys                          /
'  / ActiveLock Software Group (ASG)                                     /
'  /                                                                     /
'  / Date Created:         June 16, 2003 - THT                           /
'  / Date Last Modified:   July 07, 2003 - MEC                           /
'  /                                                                     /
'  / This software is released under the license detailed below and is   /
'  / subject to said license. Neither this header nor the licese below   /
'  / may be removed from this module.                                    /
'  ///////////////////////////////////////////////////////////////////////
'
'
'  ///////////////////////////////////////////////////////////////////////
'  /                       MODULE DESCRIPTION                            /
'  ///////////////////////////////////////////////////////////////////////
'
'   This class manages product key.  Product keys are keys used to encrypt
'   license information for a particular product.
'
'
'  ///////////////////////////////////////////////////////////////////////
'  /                        SOFTWARE LICENSE                             /
'  ///////////////////////////////////////////////////////////////////////
'
'   ActiveLock
'   Copyright 1998-2002 Nelson Ferraz
'   Copyright 2003 The ActiveLock Software Group (ASG)
'   All material is the property of the contributing authors.
'
'   Redistribution and use in source and binary forms, with or without
'   modification, are permitted provided that the following conditions are
'   met:
'
'     [o] Redistributions of source code must retain the above copyright
'         notice, this list of conditions and the following disclaimer.
'
'     [o] Redistributions in binary form must reproduce the above
'         copyright notice, this list of conditions and the following
'         disclaimer in the documentation and/or other materials provided
'         with the distribution.
'
'   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
'   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
'   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
'   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
'   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
'   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
'   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
'   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'
'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
'
'   [ ] Make some code to go in the module!
'
'
'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
'
'   07.07.03 - MEC - Updated the header comments for this file.
'
'
'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////


Option Explicit

