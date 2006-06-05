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
Imports System
Imports System.Web.UI
Imports System.Web.UI.WebControls

Public Class LimitColumn
  Inherits BoundColumn
  Private _characterLimit As Integer = 0

  Public Property CharacterLimit() As Integer
    Get
      Return _characterLimit
    End Get
    Set(ByVal Value As Integer)
      _characterLimit = Value
    End Set
  End Property

  Protected Overloads Overrides Function FormatDataValue(ByVal dataValue As Object) As String
    Return Truncate(dataValue.ToString)
  End Function

  Function Truncate(ByVal input As String) As String
    Dim output As String = input
    If output.Length > _characterLimit AndAlso _characterLimit > 0 Then
      output = output.Substring(0, _characterLimit)
      If Not (input.Substring(output.Length, 1) = " ") Then
        Dim LastSpace As Integer = output.LastIndexOf(" ")
        If Not (LastSpace = -1) Then
          output = output.Substring(0, LastSpace)
        End If
      End If
      output += "..."
    End If
    Return output
  End Function
End Class
