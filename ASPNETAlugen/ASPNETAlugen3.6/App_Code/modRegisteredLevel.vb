'*   ActiveLock
'*   Copyright 1998-2002 Nelson Ferraz
'*   Copyright 2003-2006 The ActiveLock Software Group (ASG)
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
Option Strict On
Option Explicit On 


Namespace ASPNETAlugen3


Public Module modRegisteredLevel
	
	Public strRegisteredLevelDBName As String
	
	Public Function AddBackSlash(ByVal sPath As String) As String
    'Returns sPath with a trailing backslash if sPath does not
		'already have a trailing backslash. Otherwise, returns sPath.
		sPath = Trim(sPath)
		
    If sPath.Length > 0 Then
      If sPath.Substring(sPath.Length - 1) <> "\" Then
        sPath = sPath & "\"
      End If
    End If

    Return sPath
  End Function

  Public Sub SaveComboBox(ByRef FileName As String, ByRef Control As ListControl, Optional ByRef IsItemData As Boolean = False)
    Dim FileNum As Integer
    Dim Teller As Integer
    Dim mList As ListItem

    FileNum = FreeFile()
    FileOpen(FileNum, FileName, OpenMode.Output)
    PrintLine(FileNum, IsItemData)

    For Teller = 0 To Control.items.count - 1
      mList = Control.Items(Teller)
      PrintLine(FileNum, mList.Text)
      PrintLine(FileNum, Convert.ToString(mList.Value))
    Next Teller

    FileClose(FileNum)
  End Sub

  Public Sub LoadComboBox(ByRef FileName As String, ByRef Control As ListControl, ByRef IsItemData As Boolean)
    Dim FileNum As Integer
    Dim Text As String, Text1 As String, Text2 As String

    FileNum = FreeFile()
    FileOpen(FileNum, FileName, OpenMode.Input)

    If LOF(FileNum) = 0 Then FileClose(FileNum) : Exit Sub
    Text = LineInput(FileNum)
    IsItemData = CBool(Text)

    Do While Not EOF(FileNum)
      If IsItemData = True Then
        Text1 = LineInput(FileNum)
        Text2 = LineInput(FileNum)
        Control.Items.Add(New ListItem(Text1, Text2))
      End If
    Loop

    FileClose(FileNum)
  End Sub

End Module
End Namespace
