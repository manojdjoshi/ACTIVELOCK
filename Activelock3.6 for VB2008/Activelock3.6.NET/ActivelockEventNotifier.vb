Option Strict Off
Option Explicit On
'Class instancing was changed to public
<System.Runtime.InteropServices.ProgId("ActiveLockEventNotifier_NET.ActiveLockEventNotifier")> Public Class ActiveLockEventNotifier
	'*   ActiveLock
	'*   Copyright 1998-2002 Nelson Ferraz
    '*   Copyright 2006 The ActiveLock Software Group (ASG)
	'*   All material is the property of the contributing authors.
	'*
	'*   Redistribution and use in source and binary forms, with or without
	'*   modification, are permitted provided that the following conditions are
	'*   met:
	'*
	'*     [o] Redistributions of source code must retain the above copyright
	'*         notice, this list of conditions and the following disclaimer.
	'*
	'*     [o] Redistributions in binary form must reproduce the above
	'*         copyright notice, this list of conditions and the following
	'*         disclaimer in the documentation and/or other materials provided
	'*         with the distribution.
	'*
	'*   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
	'*   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
	'*   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
	'*   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT
	'*   OWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
	'*   SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT
	'*   LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,
	'*   DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
	'*   THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
	'*   (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
	'*   OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
	'*
	'*
	
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
    ' Modified: 03.23.2006
	'===============================================================================
	'
	' @author activelock-admins
    ' @version 3.3.0
    ' @date 03.23.2006
	'
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
    'ParamArray Args was changed from ByRef to ByVal, and made Args a string
    Friend Sub Notify(ByVal EventName As String, ByRef Args As String)
        Dim Result As String
        If EventName = "ValidateValue" Then
            Result = Args
            RaiseEvent ValidateValue(Result)
            Args = Result ' assign value back to the result
        End If
    End Sub
End Class