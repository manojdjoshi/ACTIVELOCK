/*
'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2008 The ActiveLock Software Group (ASG)
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
*/
function CopyToClipboard(mytxtAreaID)
{
	var mytxtArea = document.getElementById(mytxtAreaID);
	mytxtArea.focus();
  mytxtArea.select();
  if(document.selection.createRange)
  {
		CopiedTxt = document.selection.createRange();
		CopiedTxt.execCommand("Copy");
	}
	else
	{
		//Firefox
	}
}

function PasteFromClipboard(mytxtAreaID)
{ 
	var mytxtArea = document.getElementById(mytxtAreaID);
  mytxtArea.focus();
  mytxtArea.select();
  if(mytxtArea.createTextRange)
  {
		PastedText = mytxtArea.createTextRange();
		PastedText.execCommand("Paste");
	}
	else
	{
		//??? FireFox
	}
} 

function PrintLicenseKey(content, mWidth, mHeight)
{
  var generator=window.open('','name','height=' + mHeight + ',width=' + mWidth);
  generator.document.write(content);
  generator.document.close();
}

// copyright 1999 Idocs, Inc. http://www.idocs.com
// Distribute this script freely but keep this notice in place
function letternumber(e)
{
var key;
var keychar;

if (window.event)
   key = window.event.keyCode;
else if (e)
   key = e.which;
else
   return true;
keychar = String.fromCharCode(key);
keychar = keychar.toLowerCase();

// control keys
if ((key==null) || (key==0) || (key==8) || 
    (key==9) || (key==13) || (key==27) )
   return true;

// alphas and numbers
else if ((("abcdefghijklmnopqrstuvwxyz0123456789.").indexOf(keychar) > -1))
   return true;
else
   return false;
}


function fnchkLockBIOS_CheckedChanged() {
    if (document.getElementById("chkLockBIOS").checked) {
        document.getElementById("chkLockBIOS").checked = true;
    }
    else {
        document.getElementById("chkLockBIOS").checked = false;
    }
}