Option Strict Off
Option Explicit On
Module modSHA1
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
	'===============================================================================
	' Name: modSHA1
	' Purpose: SHA1 Encryption Module
	' (C) 1998 Ian Lynagh
	' Date Created:         June 16, 1998 - IL
	' Date Last Modified:   July 07, 2003 - MEC
	' Functions:
	' Properties:
	' Methods:
	' Started: 06.16.1998
    ' Modified: 03.25.2006
	'===============================================================================
    '===============================================================================
	' Name: Function BigAA1Add
	' Input:
	'   ByVal value1 As String
	'   ByVal value2 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA1Add(ByVal value1 As String, ByVal value2 As String) As String
        Dim valueans As String = String.Empty
        Dim loopit, tempnum As Short
		
		tempnum = Len(value1) - Len(value2)
		If tempnum < 0 Then
			value1 = Space(System.Math.Abs(tempnum)) & value1
		ElseIf tempnum > 0 Then 
			value2 = Space(System.Math.Abs(tempnum)) & value2
		End If
		
		tempnum = 0
		For loopit = Len(value1) To 1 Step -1
			tempnum = tempnum + Val("&H" & Mid(value1, loopit, 1)) + Val("&H" & Mid(value2, loopit, 1))
			valueans = Hex(tempnum Mod 16) & valueans
			tempnum = Int(tempnum / 16)
		Next loopit
		If tempnum <> 0 Then
			valueans = Hex(tempnum) & valueans
		End If
		
		BigAA1Add = valueans
		
	End Function
    '===============================================================================
	' Name: Function BigAA1AND
	' Input:
	'   ByVal value1 As String
	'   ByVal value2 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA1AND(ByVal value1 As String, ByVal value2 As String) As String
        Dim valueans As String = String.Empty
        Dim loopit, tempnum As Short
		
		tempnum = Len(value1) - Len(value2)
		If tempnum < 0 Then
			value2 = Mid(value2, System.Math.Abs(tempnum) + 1)
		ElseIf tempnum > 0 Then 
			value1 = Mid(value1, tempnum + 1)
		End If
		
		For loopit = 1 To Len(value1)
			valueans = valueans & Hex(Val("&H" & Mid(value1, loopit, 1)) And Val("&H" & Mid(value2, loopit, 1)))
		Next loopit
		
		BigAA1AND = valueans
		
	End Function
    '===============================================================================
	' Name: Function BigAA1Mod32Add
	' Input:
	'   ByVal value1 As String
	'   ByVal value2 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA1Mod32Add(ByVal value1 As String, ByVal value2 As String) As String
		BigAA1Mod32Add = Right(BigAA1Add(value1, value2), 8)
	End Function
    '===============================================================================
	' Name: Function BigAA1NOT
	' Input:
	'   ByVal value1 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA1NOT(ByVal value1 As String) As String
        Dim valueans As String = String.Empty
        Dim loopit As Short
		
		value1 = Right(value1, 8)
		value1 = New String("0", 8 - Len(value1)) & value1
		For loopit = 1 To 8
			valueans = valueans & Hex(15 Xor Val("&H" & Mid(value1, loopit, 1)))
		Next loopit
		
		BigAA1NOT = valueans
		
	End Function
    '===============================================================================
	' Name: Function BigAA1OR
	' Input:
	'   ByVal value1 As String
	'   ByVal value2 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA1OR(ByVal value1 As String, ByVal value2 As String) As String
        Dim valueans As String = Nothing
		Dim loopit, tempnum As Short
		
		tempnum = Len(value1) - Len(value2)
		If tempnum < 0 Then
			valueans = Left(value2, System.Math.Abs(tempnum))
			value2 = Mid(value2, System.Math.Abs(tempnum) + 1)
		ElseIf tempnum > 0 Then 
			valueans = Left(value1, System.Math.Abs(tempnum))
			value1 = Mid(value1, tempnum + 1)
		End If
		
		For loopit = 1 To Len(value1)
			valueans = valueans & Hex(Val("&H" & Mid(value1, loopit, 1)) Or Val("&H" & Mid(value2, loopit, 1)))
		Next loopit
		
		BigAA1OR = valueans
		
	End Function
    '===============================================================================
	' Name: Function BigAA1RotLeft
	' Input:
	'   ByVal value1 As String
	'   ByVal rots As Integer
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA1RotLeft(ByVal value1 As String, ByVal rots As Short) As String
		Dim tempstr As String
		Dim loopinner, loopit, tempnum As Short
		
		rots = rots Mod 32
		If rots = 0 Then
			BigAA1RotLeft = value1
			Exit Function
		End If
		value1 = Right(value1, 8)
		tempstr = New String("0", 8 - Len(value1)) & value1
		value1 = ""
		
		' Convert to binary
		For loopit = 1 To 8
			tempnum = Val("&H" & Mid(tempstr, loopit, 1))
			For loopinner = 3 To 0 Step -1
				If tempnum And 2 ^ loopinner Then
					value1 = value1 & "1"
				Else
					value1 = value1 & "0"
				End If
			Next loopinner
		Next loopit
		
		tempstr = Mid(value1, rots + 1) & Left(value1, rots)
		
		' And convert back to hex
		value1 = ""
		For loopit = 0 To 7
			tempnum = 0
			For loopinner = 0 To 3
				If Val(Mid(tempstr, 4 * loopit + loopinner + 1, 1)) Then
					tempnum = tempnum + 2 ^ (3 - loopinner)
				End If
			Next loopinner
			value1 = value1 & Hex(tempnum)
		Next loopit
		
		BigAA1RotLeft = value1
		
	End Function
    '===============================================================================
	' Name: Function BigAA1XOR
	' Input:
	'   ByVal value1 As String
	'   ByVal value2 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA1XOR(ByVal value1 As String, ByVal value2 As String) As String
        Dim valueans As String = String.Empty
        Dim loopit, tempnum As Short
		
		tempnum = Len(value1) - Len(value2)
		If tempnum < 0 Then
			valueans = Left(value2, System.Math.Abs(tempnum))
			value2 = Mid(value2, System.Math.Abs(tempnum) + 1)
		ElseIf tempnum > 0 Then 
			valueans = Left(value1, System.Math.Abs(tempnum))
			value1 = Mid(value1, tempnum + 1)
		End If
		
		For loopit = 1 To Len(value1)
			valueans = valueans & Hex(Val("&H" & Mid(value1, loopit, 1)) Xor Val("&H" & Mid(value2, loopit, 1)))
		Next loopit
		
		BigAA1XOR = valueans
		
	End Function
    '===============================================================================
	' Name: Function BigAA2Add
	' Input:
	'   ByVal value1 As String
	'   ByVal value2 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA2Add(ByVal value1 As String, ByVal value2 As String) As String
        Dim valueans As String = String.Empty
        Dim temps1, temps2, tempstr As String
        Dim tempnum As Integer

		tempnum = Len(value1) - Len(value2)
		If tempnum < 0 Then
			value1 = New String("0", System.Math.Abs(tempnum)) & value1
		ElseIf tempnum > 0 Then 
			value2 = New String("0", tempnum) & value2
		End If
		
		tempnum = 0
		While Len(value1) > 5
			temps1 = Right(value1, 6)
			temps2 = Right(value2, 6)
			tempnum = tempnum + Val("&H" & temps1 & "&") + Val("&H" & temps2 & "&")
			tempstr = Hex(tempnum Mod 16777216)
			valueans = New String("0", 6 - Len(tempstr)) & tempstr & valueans
			tempnum = Int(tempnum / 16777216)
			value1 = Left(value1, Len(value1) - 6)
			value2 = Left(value2, Len(value2) - 6)
		End While
		tempnum = tempnum + Val("&H" & value1 & "&") + Val("&H" & value2 & "&")
		valueans = Hex(tempnum Mod 16777216) & valueans
		
		BigAA2Add = valueans
		
	End Function
    '===============================================================================
	' Name: Function BigAA2AND
	' Input:
	'   ByVal value1 As String
	'   ByVal value2 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA2AND(ByVal value1 As String, ByVal value2 As String) As String
        Dim valueans As String = String.Empty
        Dim tempstr As String
        Dim tempnum As Short

		tempnum = Len(value1) - Len(value2)
		If tempnum < 0 Then
			value2 = Mid(value2, System.Math.Abs(tempnum) + 1)
		ElseIf tempnum > 0 Then 
			value1 = Mid(value1, tempnum + 1)
		End If
		
		While Len(value1) > 7
			tempstr = Hex(Val("&H" & Left(value1, 7) & "&") And Val("&H" & Left(value2, 7) & "&"))
			valueans = valueans & New String("0", 7 - Len(tempstr)) & tempstr
			value1 = Mid(value1, 8)
			value2 = Mid(value2, 8)
		End While
		tempnum = Len(value1)
		tempstr = Hex(Val("&H" & value1 & "&") And Val("&H" & value2 & "&"))
		valueans = valueans & New String("0", tempnum - Len(tempstr)) & tempstr
		
		BigAA2AND = valueans
		
	End Function
    '===============================================================================
	' Name: Function BigAA2Mod32Add
	' Input:
	'   ByVal value1 As String
	'   ByVal value2 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA2Mod32Add(ByVal value1 As String, ByVal value2 As String) As String
		BigAA2Mod32Add = BigAA1Mod32Add(value1, value2)
	End Function
    '===============================================================================
	' Name: Function BigAA2NOT
	' Input:
	'   ByVal value1 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA2NOT(ByVal value1 As String) As String
		Dim valueans As String

		value1 = Right(value1, 8)
		value1 = New String("0", 8 - Len(value1)) & value1
		valueans = Hex(65535 Xor Val("&H" & Right(value1, 4) & "&"))
		valueans = New String("0", 4 - Len(valueans)) & valueans
		valueans = Hex(65535 Xor Val("&H" & Left(value1, 4) & "&")) & valueans
		valueans = New String("0", 8 - Len(valueans)) & valueans
		
		BigAA2NOT = valueans
		
	End Function
    '===============================================================================
	' Name: Function BigAA2OR
	' Input:
	'   ByVal value1 As String
	'   ByVal value2 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA2OR(ByVal value1 As String, ByVal value2 As String) As String
        Dim valueans As String = String.Empty
        Dim tempstr As String
        Dim tempnum As Short

		tempnum = Len(value1) - Len(value2)
		If tempnum < 0 Then
			valueans = Left(value2, System.Math.Abs(tempnum))
			value2 = Mid(value2, System.Math.Abs(tempnum) + 1)
		ElseIf tempnum > 0 Then 
			valueans = Left(value1, System.Math.Abs(tempnum))
			value1 = Mid(value1, tempnum + 1)
		End If
		
		While Len(value1) > 7
			tempstr = Hex(Val("&H" & Left(value1, 7) & "&") Or Val("&H" & Left(value2, 7) & "&"))
			valueans = valueans & New String("0", 7 - Len(tempstr)) & tempstr
			value1 = Mid(value1, 8)
			value2 = Mid(value2, 8)
		End While
		tempnum = Len(value1)
		tempstr = Hex(Val("&H" & value1 & "&") Or Val("&H" & value2 & "&"))
		valueans = valueans & New String("0", tempnum - Len(tempstr)) & tempstr
		
		BigAA2OR = valueans
		
	End Function
    '===============================================================================
	' Name: Function BigAA2RotLeft
	' Input:
	'   ByVal value1 As String
	'   ByVal rots As Integer
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA2RotLeft(ByVal value1 As String, ByVal rots As Short) As String
		BigAA2RotLeft = BigAA1RotLeft(value1, rots)
	End Function
    '===============================================================================
	' Name: Function BigAA2XOR
	' Input:
	'   ByVal value1 As String
	'   ByVal value2 As String
	' Output: String
	' Purpose: [INTERNAL] SHA function
	' Remarks: None
	'===============================================================================
	Function BigAA2XOR(ByVal value1 As String, ByVal value2 As String) As String
        Dim valueans As String = String.Empty
        Dim tempstr As String
        Dim tempnum As Short

		tempnum = Len(value1) - Len(value2)
		If tempnum < 0 Then
			valueans = Left(value2, System.Math.Abs(tempnum))
			value2 = Mid(value2, System.Math.Abs(tempnum) + 1)
		ElseIf tempnum > 0 Then 
			valueans = Left(value1, System.Math.Abs(tempnum))
			value1 = Mid(value1, tempnum + 1)
		End If
		
		While Len(value1) > 7
			tempstr = Hex(Val("&H" & Left(value1, 7) & "&") Xor Val("&H" & Left(value2, 7) & "&"))
			valueans = valueans & New String("0", 7 - Len(tempstr)) & tempstr
			value1 = Mid(value1, 8)
			value2 = Mid(value2, 8)
		End While
		tempnum = Len(value1)
		tempstr = Hex(Val("&H" & value1 & "&") Xor Val("&H" & value2 & "&"))
		valueans = valueans & New String("0", tempnum - Len(tempstr)) & tempstr
		
		BigAA2XOR = valueans
		
	End Function
    '===============================================================================
	' Name: Function SHA1AA1Hash
	' Input:
	'   ByVal hashthis As String - Input string to be hashed
	' Output: String
	' Purpose: SHA Hash function
	' Remarks: None
	'===============================================================================
	Function SHA1AA1Hash(ByVal hashthis As String) As String
		Dim buf(4) As String
		Dim in_(79) As String
		Dim loopouter, tempnum2, tempnum, loopit, loopinner As Short
		Dim d, b, a, c, e As String
        Dim tempstr As String = String.Empty

		' Add padding
		tempnum = 8 * Len(hashthis)
		hashthis = hashthis & Chr(128) 'Add binary 10000000
		tempnum2 = 56 - Len(hashthis) Mod 64
		If tempnum2 < 0 Then
			tempnum2 = 64 + tempnum2
		End If
		hashthis = hashthis & New String(Chr(0), tempnum2)
		For loopit = 1 To 8
			tempstr = Chr(tempnum Mod 256) & tempstr
			tempnum = tempnum - tempnum Mod 256
			tempnum = tempnum / 256
		Next loopit
		hashthis = hashthis & tempstr
		
		' Set magic numbers
		buf(0) = "67452301"
		buf(1) = "efcdab89"
		buf(2) = "98badcfe"
		buf(3) = "10325476"
		buf(4) = "c3d2e1f0"
		
		' For each 512 bit section
		For loopouter = 0 To Len(hashthis) / 64 - 1
			a = buf(0)
			b = buf(1)
			c = buf(2)
			d = buf(3)
			e = buf(4)
			
			' Get the 512 bits
			For loopit = 0 To 15
				in_(loopit) = ""
				For loopinner = 4 To 1 Step -1
					in_(loopit) = Hex(Asc(Mid(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) & in_(loopit)
					If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" & in_(loopit)
				Next loopinner
			Next loopit
			
			For loopit = 16 To 79
				in_(loopit) = BigAA1RotLeft(BigAA1XOR(BigAA1XOR(BigAA1XOR(in_(loopit - 3), in_(loopit - 8)), in_(loopit - 14)), in_(loopit - 16)), 1)
			Next loopit
			
			For loopit = 0 To 19
				tempstr = BigAA1OR(BigAA1AND(b, c), BigAA1AND(BigAA1NOT(b), d))
				tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_(loopit), "5A827999"))))
				e = d
				d = c
				c = BigAA1RotLeft(b, 30)
				b = a
				a = tempstr
			Next loopit
			
			For loopit = 20 To 39
				tempstr = BigAA1XOR(BigAA1XOR(b, c), d)
				tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_(loopit), "6ED9EBA1"))))
				e = d
				d = c
				c = BigAA1RotLeft(b, 30)
				b = a
				a = tempstr
			Next loopit
			
			For loopit = 40 To 59
				tempstr = BigAA1OR(BigAA1OR(BigAA1AND(b, c), BigAA1AND(b, d)), BigAA1AND(c, d))
				tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_(loopit), "8F1BBCDC"))))
				e = d
				d = c
				c = BigAA1RotLeft(b, 30)
				b = a
				a = tempstr
			Next loopit
			
			For loopit = 60 To 79
				tempstr = BigAA1XOR(BigAA1XOR(b, c), d)
				tempstr = BigAA1Mod32Add(BigAA1RotLeft(a, 5), BigAA1Mod32Add(tempstr, BigAA1Mod32Add(e, BigAA1Mod32Add(in_(loopit), "CA62C1D6"))))
				e = d
				d = c
				c = BigAA1RotLeft(b, 30)
				b = a
				a = tempstr
			Next loopit
			
			buf(0) = BigAA1Mod32Add(buf(0), a)
			buf(1) = BigAA1Mod32Add(buf(1), b)
			buf(2) = BigAA1Mod32Add(buf(2), c)
			buf(3) = BigAA1Mod32Add(buf(3), d)
			buf(4) = BigAA1Mod32Add(buf(4), e)
			
		Next loopouter
		
		' Extract MD5Hash
		hashthis = ""
		For loopit = 0 To 4
			For loopinner = 0 To 3
				hashthis = hashthis & Mid(buf(loopit), 1 + 2 * loopinner, 2)
			Next loopinner
		Next loopit
		
		' And return it
		SHA1AA1Hash = hashthis
		
	End Function
    '===============================================================================
	' Name: Function SHA1AA2Hash
	' Input:
	'   ByVal hashthis As String - Input string to be hashed
	' Output: String
	' Purpose: SHA Hash function
	' Remarks: None
	'===============================================================================
	Function SHA1AA2Hash(ByVal hashthis As String) As String
		Dim buf(4) As String
		Dim in_(79) As String
		Dim loopouter, tempnum2, tempnum, loopit, loopinner As Short
		Dim d, b, a, c, e As String
        Dim tempstr As String = String.Empty

		' Add padding
		tempnum = 8 * Len(hashthis)
		hashthis = hashthis & Chr(128) 'Add binary 10000000
		tempnum2 = 56 - Len(hashthis) Mod 64
		If tempnum2 < 0 Then
			tempnum2 = 64 + tempnum2
		End If
		hashthis = hashthis & New String(Chr(0), tempnum2)
		For loopit = 1 To 8
			tempstr = Chr(tempnum Mod 256) & tempstr
			tempnum = tempnum - tempnum Mod 256
			tempnum = tempnum / 256
		Next loopit
		hashthis = hashthis & tempstr
		
		' Set magic numbers
		buf(0) = "67452301"
		buf(1) = "efcdab89"
		buf(2) = "98badcfe"
		buf(3) = "10325476"
		buf(4) = "c3d2e1f0"
		
		' For each 512 bit section
		For loopouter = 0 To Len(hashthis) / 64 - 1
			a = buf(0)
			b = buf(1)
			c = buf(2)
			d = buf(3)
			e = buf(4)
			
			' Get the 512 bits
			For loopit = 0 To 15
				in_(loopit) = ""
				For loopinner = 4 To 1 Step -1
					in_(loopit) = Hex(Asc(Mid(hashthis, 64 * loopouter + 4 * loopit + loopinner, 1))) & in_(loopit)
					If Len(in_(loopit)) Mod 2 Then in_(loopit) = "0" & in_(loopit)
				Next loopinner
			Next loopit
			
			For loopit = 16 To 79
				in_(loopit) = BigAA2RotLeft(BigAA2XOR(BigAA2XOR(BigAA2XOR(in_(loopit - 3), in_(loopit - 8)), in_(loopit - 14)), in_(loopit - 16)), 1)
			Next loopit
			
			For loopit = 0 To 19
				tempstr = BigAA2OR(BigAA2AND(b, c), BigAA2AND(BigAA2NOT(b), d))
				tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_(loopit), "5A827999"))))
				e = d
				d = c
				c = BigAA2RotLeft(b, 30)
				b = a
				a = tempstr
			Next loopit
			
			For loopit = 20 To 39
				tempstr = BigAA2XOR(BigAA2XOR(b, c), d)
				tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_(loopit), "6ED9EBA1"))))
				e = d
				d = c
				c = BigAA2RotLeft(b, 30)
				b = a
				a = tempstr
			Next loopit
			
			For loopit = 40 To 59
				tempstr = BigAA2OR(BigAA2OR(BigAA2AND(b, c), BigAA2AND(b, d)), BigAA2AND(c, d))
				tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_(loopit), "8F1BBCDC"))))
				e = d
				d = c
				c = BigAA2RotLeft(b, 30)
				b = a
				a = tempstr
			Next loopit
			
			For loopit = 60 To 79
				tempstr = BigAA2XOR(BigAA2XOR(b, c), d)
				tempstr = BigAA2Mod32Add(BigAA2RotLeft(a, 5), BigAA2Mod32Add(tempstr, BigAA2Mod32Add(e, BigAA2Mod32Add(in_(loopit), "CA62C1D6"))))
				e = d
				d = c
				c = BigAA2RotLeft(b, 30)
				b = a
				a = tempstr
			Next loopit
			
			buf(0) = BigAA2Mod32Add(buf(0), a)
			buf(1) = BigAA2Mod32Add(buf(1), b)
			buf(2) = BigAA2Mod32Add(buf(2), c)
			buf(3) = BigAA2Mod32Add(buf(3), d)
			buf(4) = BigAA2Mod32Add(buf(4), e)
			
		Next loopouter
		
		' Extract MD5Hash
		hashthis = ""
		For loopit = 0 To 4
			For loopinner = 0 To 3
				hashthis = hashthis & Mid(buf(loopit), 1 + 2 * loopinner, 2)
			Next loopinner
		Next loopit
		
		' And return it
		SHA1AA2Hash = hashthis
		
	End Function
End Module