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
'*
Option Strict On
Option Explicit On 

Module modRegisteredLevel
	
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

  Public Sub SaveComboBox(ByRef FileName As String, ByRef myEnum As IEnumerator, Optional ByRef IsItemData As Boolean = False)
    Dim FileNum As Integer
    Dim mList As Mylist
   
    FileNum = FreeFile()
    FileOpen(FileNum, FileName, OpenMode.Output)
    PrintLine(FileNum, IsItemData)

    Do While myEnum.MoveNext
      mList = CType(myEnum.Current, Mylist)
      PrintLine(FileNum, mList.Name)
      PrintLine(FileNum, Convert.ToString(mList.ItemData))
    Loop
    'For Teller = 0 To ListItems.Count - 1 ' Control.Items.Count
    '  mList = ListItems(Teller) 'Control.Items(Teller)
    '  PrintLine(FileNum, mList.Name)
    '  PrintLine(FileNum, Convert.ToString(mList.ItemData))
    'Next Teller

    FileClose(FileNum)
  End Sub

  Public Sub LoadComboBox(ByRef FileName As String, ByRef Control As ComboBox, ByRef IsItemData As Boolean)
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
        Control.Items.Add(New Mylist(Text1, Convert.ToInt32(Text2)))
      End If
    Loop

    FileClose(FileNum)
  End Sub

  Public Sub LoadListBox(ByRef FileName As String, ByRef Control As ListBox, ByRef IsItemData As Boolean)
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
        Control.Items.Add(New Mylist(Text1, Convert.ToInt32(Text2)))
      End If
    Loop

    FileClose(FileNum)
  End Sub

  Public Class Mylist
    Private sName As String
    ' You can also declare this as String,bitmap or almost anything. 
    ' If you change this delcaration you will also need to change the Sub New 
    ' to reflect any change. Also the ItemData Property will need to be updated. 
    Private iID As Integer

    ' Default empty constructor. 
    Public Sub New()
      sName = ""
      ' This would need to be changed if you modified the declaration above. 
      iID = 0
    End Sub

    Public Sub New(ByVal Name As String, ByVal ID As Integer)
      sName = Name
      iID = ID
    End Sub

    Public Property Name() As String
      Get
        Return sName
      End Get
      Set(ByVal sValue As String)
        sName = sValue
      End Set
    End Property

    ' This is the property that holds the extra data. 
    Public Property ItemData() As Int32
      Get
        Return iID
      End Get
      Set(ByVal iValue As Int32)
        iID = iValue
      End Set
    End Property

    ' This is neccessary because the ListBox and ComboBox rely 
    ' on this method when determining the text to display. 
    Public Overrides Function ToString() As String
      Return sName
    End Function

  End Class
End Module