VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ActiveLockEventNotifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
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
' Name: ActivelockEventNotifier
' Purpose: This class handles ActiveLock COM event notifications to the interested observers.
' <p>It is simply a wrapper containing public events.<p>These events should
' really belong in IActiveLock, but since VB doesn&#39;t support inheritance
' of events, we have to do it this way.
' Remarks:
' Functions:
' Properties:
' Methods:
' Started: 21.04.2005
' Modified: 08.05.2005
'===============================================================================
'
' @author activelock-admins
' @version 3.0.0
' @date 08.05.2005
'
'* ///////////////////////////////////////////////////////////////////////
'  /                        MODULE TO DO LIST                            /
'  ///////////////////////////////////////////////////////////////////////
'   [ ] TODO Item 1
'   [ ] TODO Item 2
'
'  ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' @history
' <pre>
'   07.20.03 - th2tran - Created
'   08.03.03 - th2tran  - VBDox&#39;ed this interface.
'   10.13.03 - th2tran  - ValidateValue() event signature changed from 2 parameters
'                         to 1 parameter, for simplicity.
' </pre>

'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////

Option Explicit


' (Optional) Product License Property Value validation event allows the client application
' to return the encrypted version of a license property value (such as LastRunDate).
' <p>An example, of when <code>ValidateValue</code> event would be used,
' can be observed for the <code>LastRunDate</code> property.
' For readability, this property is saved in the KeyStore in plain-text format. However, to prevent hackers from
' changing this value, an accompanying Hash Code for this value, <code>Hash1</code>, is also stored. This Hash Code
' is an MD5 hash of the (possibly) encrypted value of <code>LastRunDate</code>.  The encrypted value is
' is user application specific, and is obtained from the user application via the <code>ValidateValue</code> event.
' The client will receive this event, encrypt <code>Value</code> using its own encryption algorithm,
' and store the result back in <code>Value</code> to be returned to ActiveLock.
' <p>Handling of this event is OPTIONAL.  If not handled, it simply means there will be no encryption for
' the stored property values.
'
' @param Value  Property value.
Public Event ValidateValue(ByRef Value As String)

'===============================================================================
' Name: Sub Notify
' Input:
'   ByVal EventName As String - Event name
'   ByRef ParamArray Args As Variant - Parametric array arguments
' Output: None
' Purpose: Handles ActiveLock COM event notifications to the interested observers
' Remarks: None
'===============================================================================
Friend Sub Notify(ByVal EventName As String, ParamArray Args())
    If EventName = "ValidateValue" Then
        Dim Result As String
        Result = Args(0)
        RaiseEvent ValidateValue(Result)
        Args(0) = Result ' assign value back to the result
    End If
End Sub

