Attribute VB_Name = "modCryptoAPI"
Option Explicit

'******************************************************************************
'*
'* Module MCommon
'*
'* Assisting module to CCipher, CDataBlock and CHash. Implements checks and error handling
'* helpers used by these classes.
'*
'* (c) 2004 Ulrich Korndörfer proSource software development
'*          www.prosource.de
'*          German site with VB articles (in english) and code (comments in english)
'*
'* Precautions: None. May be compiled to native code (which is strongly recommended),
'*              with all extended options selected.
'*
'* External dependencies: MVarInfo
'*
'* Version history
'*
'* Version 1.1 from 2004.09.23 Now uses MVarInfo. This allows to use the extended
'*                             options when compiling to native code.
'* Version 1.0 from 2004.04.01
'*
'*  Disclaimer:
'*  All code in this class is for demonstration purposes only.
'*  It may be used in a production environment, as it is thoroughly tested,
'*  but do not hold us responsible for any consequences resulting of using it.
'*
'******************************************************************************


'******************************************************************************
'* Private consts
'******************************************************************************

Private Const cRANGEFROM As Long = 1
Private Const cRANGETO As Long = (17 * 17 - 1) '288
'Private Const cRANGETO As Long = (29 * 29 - 1) '840

Public publicPrivateKeyBlob As String
Public publicPublicKeyBlob As String

'******************************************************************************
'*
'* This module implements one method called gGetVarInfo, which delivers information
'* about variables, especially array variables, without the need of provoking errors.
'*
'* Use it to decide: is it an array var
'*                   is it an undimmed array (aka does not have a SA descriptor)
'*                   is it a zombie array (dimmed with UBound < LBound)
'*                   how many dimensions has the array
'*                   the bounds of the array's dimensions (stored in *reverse*
'*                   order, see below)
'*
'* Note: can not be used with arrays of basetype fixed string
'*       if used with UDTs, the UDT has to be defined in a public object module
'*
'*
'* (c) 2004 Ulrich Korndörfer proSource software development
'*          www.prosource.de
'*          German site with VB articles (in english) and code (comments in english)
'*
'* Precautions: None. May be compiled to native code (which is strongly recommended),
'*              with all extended options selected.
'*
'* External dependencies: some APIs (see below)
'*
'* Version history
'*
'* Version 1.0 from 2004.09.23
'*
'*  Disclaimer:
'*  All code in this class is for demonstration purposes only.
'*  It may be used in a production environment, as it is thoroughly tested,
'*  but do not hold us responsible for any consequences resulting of using it.
'*
'******************************************************************************


'******************************************************************************
'* API declarations and consts
'******************************************************************************

Private Declare Sub CopyMem4 Lib "msvbvm60.dll" Alias "GetMem4" _
       (ByVal FromAddr As Long, ByRef ToAddr As Any)
Private Declare Sub CopyMem8 Lib "msvbvm60.dll" Alias "GetMem8" _
       (ByVal FromAddr As Long, ByRef ToAddr As Any)

Private Const VARIANT_DATA_OFFSET As Long = 8
Private Const SAD_LEN As Long = 16
Private Const SADBOUND_LEN As Long = 8


'******************************************************************************
'* Types
'******************************************************************************

'TVarInfo_SADBOUND is translated to VB from the SAFEARRAYBOUND structure as defined in MSDN
'Note: HighBound =  UBound(Arr) = LowBound + ElementCount - 1

Private Type TVarInfo_SADBOUND
  ElementCount As Long 'How many elements this dimension has
  LowBound As Long     'The minimum index of this dimension (equals LBound(Arr))
End Type

'TVarInfo holds all gathered information about the variable.
'The part starting with SAD_DimCount and ending with SAD_DataAdress is translated
'from the SAFEARRAY structure as defined in MSDN.

Public Type TVarInfo

  'Variable's type
  
  Var_BaseType As VbVarType 'Base type like vbSingle etc. The vbArray part has been removed.
  Var_IsArray As Boolean    'Is it an array var (vbArray).
  
  'If VarIsArray = True and a SA descriptor has been set for it
  '(if the array has been "dimmed"), then Var_SADAdress <> 0.
  
  Var_SADAdress As Long   'Adress of the SA descriptor
  
  'If Var_SADAdress <> 0, the SA descriptors data follow here.
  'The following is a 1:1 translation of the SAFEARRAY struct as defined in the MSDN.
  'It has a length of 16 bytes.
  
  SAD_DimCount As Integer 'How many dimensions the array has
  SAD_Features As Integer 'Features flags
  SAD_ElementSize As Long 'The size of the arrays element type
  SAD_Locks As Long       'How many locks are applied
  SAD_DataAddress As Long 'Address of the memory location, where the data starts
  
  'If SAD_DimCount > 0, the bounds of the dimensions follow here.
  'Note: a VB array with a SA descriptor set (that is with Var_SADAdress <> 0)
  'always has a SAD_DimCount > 0
  'Note: the dims are stored in *reverse* order!
  '      Dim 1 has index SAD_DimCount - 1, Dim n has index 0!
  
  SAD_Bounds() As TVarInfo_SADBOUND

End Type



'******************************************************************************
'* Public array checking methods
'******************************************************************************

'Arr must be initialized as one-dimensional array and have at least one byte (no zombie).
'Arr may be static or dynamic.

'If given, Low or High both must be valid indexes to Arr (including type and magnitude)
'If both Low and High are given, Low must be lower equal High
'Sets L, H and Length

Public Sub gCheckArray(ByRef Arr() As Byte, _
                       ByRef Low As Variant, _
                       ByRef High As Variant, _
                       ByRef l As Long, _
                       ByRef h As Long, _
                       ByRef Length As Long, _
                       ByRef ClassName As String, _
                       ByRef MethodName As String, _
                       ByRef Description As String)
Dim lh As Long

On Error GoTo MethodError

With gGetVarInfo(Arr)
  If .SAD_DimCount <> 1 Then GoTo ErrorExit 'Either not initialized or wrong dimension
  With .SAD_Bounds(0)
    If .ElementCount = 0 Then GoTo ErrorExit 'Is zombie
    lh = .LowBound + .ElementCount - 1
    'Get boundaries. Throws error if Low or High can not be converted to a long value
    If IsMissing(Low) Then l = .LowBound Else l = Low
    If IsMissing(High) Then h = lh Else h = High
    'Check, if L and H are valid indexes
    If l < .LowBound Or l > lh Then GoTo ErrorExit
    If h < .LowBound Or h > lh Then GoTo ErrorExit
    'Check, if H >= L
    Length = h - l + 1
    If (Length < 1) Then GoTo ErrorExit
  End With
End With

Exit Sub

ErrorExit:
  On Error GoTo 0
MethodError:
  gRaiseError ClassName, MethodName, Description, 5

End Sub

'Arr must be initialized as one-dimensional array and have at least one byte (no zombie),
'starting at Index 0.
'Arr may be static or dynamic.

'Pos must be a valid index to Arr.
'Sets Pos and Length

Public Sub gSimpleCheckArray(ByRef Arr() As Byte, _
                             ByRef Length As Long, _
                             ByVal Pos As Long, _
                             ByRef ClassName As String, _
                             ByRef MethodName As String, _
                             ByRef Description As String)

On Error GoTo MethodError

With gGetVarInfo(Arr)
  If .SAD_DimCount <> 1 Then GoTo ErrorExit 'Either not initialized or wrong dimension
  With .SAD_Bounds(0)
    If .ElementCount = 0 Or .LowBound <> 0 Then GoTo ErrorExit 'Is zombie or has wrong base index
    If Pos < 0 Or Pos >= .ElementCount Then GoTo ErrorExit 'Pos is invalid index
    Length = .ElementCount
  End With
End With

Exit Sub

ErrorExit:
  On Error GoTo 0
MethodError:
  gRaiseError ClassName, MethodName, Description, 5

End Sub


'******************************************************************************
'* Public methods
'******************************************************************************

'Not allowed array types are:

'- Arrays of fixed strings, eg. Dim FSA() As String * 10.
'  Those can not be wrapped into a Variant!

'Noteworthy array types are:

'- Arrays of object references, eg. Dim OA() As CMyClass. Those can be used,
'  but note: an *undimmed* array of any object type always is (by VB)
'  created as zombie array!

'- Arrays of UDTs, eg. Dim UA() As TMyType. Those can be used only if TMyType is
'  declared as public in a public object module! Also note that, like object arrays,
'  an *undimmed* array of any UDT-type always is (by VB) created as zombie array!

Public Function gGetVarInfo(ByRef theVar As Variant) As TVarInfo
Dim i As Long

With gGetVarInfo
  
  .Var_BaseType = VarType(theVar)
  .Var_IsArray = (.Var_BaseType And vbArray) <> 0
  .Var_BaseType = .Var_BaseType And Not vbArray
  
  If .Var_IsArray Then
    
    'Get the adress of SAD (the array var's safearray descriptor)
    
    CopyMem4 UAdd(VarPtr(theVar), VARIANT_DATA_OFFSET), .Var_SADAdress
    CopyMem4 .Var_SADAdress, .Var_SADAdress
    
    'If it has none, it is "undimmed". Exit then as there is no more data available
    
    If .Var_SADAdress = 0 Then Exit Function

    'Copy the content of the array's SAD into our type
    
    CopyMem8 .Var_SADAdress, .SAD_DimCount
    CopyMem8 UAdd(.Var_SADAdress, 8), .SAD_Locks
    
    'If the array has no dimensions, there also are no bounds. So exit then.
    
    If .SAD_DimCount = 0 Then Exit Function
  
    'Copy the array bounds into our type.
    
    ReDim .SAD_Bounds(0 To .SAD_DimCount - 1)
    For i = 0 To .SAD_DimCount - 1
      CopyMem8 UAdd(.Var_SADAdress, SAD_LEN + i * 8), .SAD_Bounds(i)
    Next i
    
  End If

End With

End Function

'******************************************************************************
'* Private helpers
'******************************************************************************

'UAdd does an unsigned add of Base and Offset.

'Base and Offset are treated as *unsigned* longs and are added unsigned, with:

'&H0 <= Base        <= &HFFFFFFFF and
'&H0 <= Offset      <= &H7FFFFFFF and of course for the unsigned sum
'&H0 <= Base+Offset <= &HFFFFFFFF

'This would be the code for an unrestricted unsigned add:

'&H0 <= Base        <= &HFFFFFFFF and
'&H0 <= Offset      <= &HFFFFFFFF and
'&H0 <= Base+Offset <= &HFFFFFFFF

'UAdd = (((Base Xor &H80000000) + (Offset And &H7FFFFFFF)) Xor &H80000000) Or (Offset And &H80000000)

'In both cases overflow behaviour is complex and has to be investigated.

Private Function UAdd(ByVal Base As Long, ByVal Offset As Long) As Long
UAdd = ((Base Xor &H80000000) + Offset) Xor &H80000000
End Function


'******************************************************************************
'* Public others
'******************************************************************************

'gCheckPrimeInRange checks:

'- Is Value in the range cRANGEFROM to cRANGETO (including the borders)
'- Is Value a prime number from this range.

'gPrintPrimes is a testroutine for gCheckPrimeInRange, also callable directly
'in the debug window, which prints 62 found primes between 1 and 288:

'    1    2    3    5    7   11   13   17   19   23
'   29   31   37   41   43   47   53   59   61   67
'   71   73   79   83   89   97  101  103  107  109
'  113  127  131  137  139  149  151  157  163  167
'  173  179  181  191  193  197  199  211  223  227
'  229  233  239  241  251  257  263  269  271  277
'  281  283

'Further primes are:

'            293  307  311  313  317  331  337  347
'  349  353  359  367  373  379  383  389  397  401
'  409  419  421  431  433  439  443  449  457  461
'  463  467  479  487  491  499  503  509  521  523
'  541  547  557  563  569  571  577  587  593  599
'  601  607  613  617  619  631  641  643  647  653
'  659  661  673  677  683  691  701  709  719  727
'  733  739  743  751  757  761  769  773  787  797
'  809  811  821  823  827  829  839

'Between and including 17 and 251 there are 48 primes.

#If TESTMODE = 1 Then

  Public Sub gPrintPrimes()
  Dim i As Long, s As String, C As Long
  
  For i = cRANGEFROM To cRANGETO
    If gCheckPrimeInRange(i) Then
      s = s & FInt(i, 5): C = C + 1
      If C Mod 10 = 0 Then Debug.Print s: s = vbNullString: C = 0
    End If
  Next i
  If C <> 0 Then Debug.Print s
  End Sub

  Private Function FInt(ByVal Val As String, ByVal PadLen As Long) As String
  Dim l As Long
  l = Len(Val): If l >= PadLen Then FInt = Val Else FInt = Space$(PadLen - l) & Val
  End Function


#End If

Public Function gCheckPrimeInRange(ByVal Value As Long) As Boolean

Select Case True
  Case Value < cRANGEFROM
  Case Value > cRANGETO
  Case (Value Mod 2) = 0:  gCheckPrimeInRange = (Value = 2)
  Case (Value Mod 3) = 0:  gCheckPrimeInRange = (Value = 3)
  Case (Value Mod 5) = 0:  gCheckPrimeInRange = (Value = 5)
  Case (Value Mod 7) = 0:  gCheckPrimeInRange = (Value = 7)
  Case (Value Mod 11) = 0: gCheckPrimeInRange = (Value = 11)
  Case (Value Mod 13) = 0: gCheckPrimeInRange = (Value = 13) 'For cRANGETO = 288
'  Case (Value Mod 17) = 0: gCheckPrimeInRange = (Value = 17)
'  Case (Value Mod 19) = 0: gCheckPrimeInRange = (Value = 19)
'  Case (Value Mod 23) = 0: gCheckPrimeInRange = (Value = 23) 'For cRANGETO = 840
  Case Else:               gCheckPrimeInRange = True
End Select
End Function


'******************************************************************************
'* Public error handling helpers
'******************************************************************************

Public Sub gCheckCond(ByVal ErrCond As Boolean, _
                      ByRef ClassName As String, _
                      ByRef MethodName As String, _
                      ByRef Description As String)
If ErrCond Then gRaiseError ClassName, MethodName, Description, 5
End Sub

Public Sub gRaiseError(ByRef ClassName As String, _
                       ByRef MethodName As String, _
                       Optional ByRef Description As String, _
                       Optional ByVal Number As Long)

If Err.Number = 0 Then
    Err.Raise Number, ClassName & "." & MethodName, Description
Else
    If Len(Description) = 0 Then
        Err.Raise Err.Number, ClassName & "." & MethodName & vbCrLf & Err.Source, Err.Description
    Else
        Err.Raise Err.Number, ClassName & "." & MethodName & vbCrLf & Err.Source, Description & vbCrLf & Err.Description
    End If
End If

End Sub

