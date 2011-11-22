VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
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
' Name: INIFile
' Purpose: Stores and retrieves product keys
' <p>An "object-oriented" approach to using Windows INI files, with some
'   useful additions.
' <p>Klaus H. Probst [kprobst@vbbox.com]
' Functions:
' Properties:
' Methods:
' Started: 07.07.2003
' Modified: 08.07.2005
'===============================================================================

'  ///////////////////////////////////////////////////////////////////////
'  / Filename:  INIFile.cls                                              /
'  / Version:   1.0.0.1                                                  /
'  / Purpose:   Stores and retrieves product keys                         /
'  / Klaus H. Probst [kprobst@vbbox.com]                                 /
'  /                                                                     /
'  / Date Created:         ???? ??, ???? - KHP                           /
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
'  ///////////////////////////////////////////////////////////////////////
'  /                        SOFTWARE LICENSE                             /
'  ///////////////////////////////////////////////////////////////////////
'
'   If you intend to distribute the file(s) that make up this sample to
'   any WWW site, online service, electronic bulletin board system (BBS),
'   CD or any other electronic or physical media, you must notify me in
'   advance to obtain my express permission.
'
' * ///////////////////////////////////////////////////////////////////////
'  /                        MODULE CHANGE LOG                            /
'  ///////////////////////////////////////////////////////////////////////
' @history
' <pre>
' 11.02.03 - th2tran       - Allow key values to have max length of 2048 bytes
' 05.13.05 - ialkan        - Eliminated multiple INIFile.cls files and merged into one
' </pre>
'
'  ///////////////////////////////////////////////////////////////////////
'  /                MODULE CODE BEGINS BELOW THIS LINE                   /
'  ///////////////////////////////////////////////////////////////////////

Option Explicit
DefLng A-Z

'Private Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As Any, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSectionNames Lib "kernel32" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private m_sFileName As String
Private m_sSection As String

Private m_LastError As String ' Last error message

Private Const MAX_BUFFER_SIZE As Long = 4096


'===============================================================================
' Name: Function GetValue
' Input:
'   ByVal ValueName As String - Passed key
'   ByVal Default As String - Default return value to be used if the call fails
'   ByRef Failed As Boolean - This will be set to False if everything went well, or True of something went wrong
' Output:
'   Variant - This will be set to False if everything went well, or True of something went wrong, and the return is
' the value passed in the Default argument.
' Purpose: Retrieves a value from the passed key (ValueName) and returns it as a variant
' (String subtype). This proc is useful if your requirements go above that of the
' Values Get/Let pair. You can specify a Default return value in case the call fails,
' and you can pass a Boolean variable in the Fail argument.
' Remarks: None
'===============================================================================
Public Function GetValue(ByVal ValueName As String, Optional ByVal Default As String, Optional ByRef Failed As Boolean) As Variant
    Dim sBuffer As String
    Dim lReturn As Long
    
    m_LastError = "" ' reset error message
    
    sBuffer = String$(MAX_BUFFER_SIZE - 1, vbNullChar)
    lReturn = GetPrivateProfileString(m_sSection, ValueName, Default, sBuffer, Len(sBuffer), m_sFileName)

    If lReturn > 0 Then
        Failed = False
        GetValue = Left$(sBuffer, lReturn)
    Else
        Failed = True
        m_LastError = WinError(lReturn)

        GetValue = Default
    End If

End Function



'===============================================================================
' Name: Sub AddSection
' Input:
'   ByVal name As String - Section name to be added
'   ByRef SetAsCurrent As Boolean - If True, sets the added section as the current section name
' Output: None
' Purpose: Adds an empty section to the current file. If you're wondering
' why this is here, know that I use it a lot, but I suppose it's
' of limited value in most cases.
' Remarks: None
'===============================================================================
Public Sub AddSection(ByVal name As String, Optional SetAsCurrent As Boolean = False)
    On Error Resume Next
    '// to add an empty section, we have to write a dummy
    '// key into it (which automatically creates the section), and
    '// then delete the dummy key. This leaves the section intact
    Call WritePrivateProfileString(name, "Dummy", "Val", m_sFileName)
    Call WritePrivateProfileString(name, "Dummy", 0&, m_sFileName)

    If SetAsCurrent = True Then m_sSection = name
End Sub

'===============================================================================
' Name: Function EnumSectionKeys
' Input:
'   ByRef ArrayResult As Variant - Returned array
' Output:
'   Integer - Returns 0 if failure
' Purpose: Enumerates the keys (not the Key-Value pairs) under the current section,
' copies the array of keys into the ArrayResult argument, and returns the
' number of keys enumerated.
' <p>For how the class handles buffer sizes on INI calls, please see the Notes
' section on the [Declarations] section of the class module.
' <p>If the method fails, the return will be zero, the ArrayResult argument
' will be set to Null and you will hit an assert.
' Remarks: None
'===============================================================================
Public Function EnumSectionKeys(ByRef ArrayResult As Variant) As Integer
    Dim sBuffer As String
    Dim lReturn As Long
    Dim sValuePair As String
    Dim iCounter As Integer
    Dim asEnumKeys() As String
    Dim lBuffSize As Long
    Dim iDelim As Integer
    Dim iTries As Integer

    On Error GoTo catch
    iCounter = -1         '// preset this
    lBuffSize = MAX_BUFFER_SIZE   '// start out with 4096 bytes
    Do
        lBuffSize = (lBuffSize * 2)
        lBuffSize = lBuffSize - 1       '// otherwise we overflow 32767
        sBuffer = String$(lBuffSize, vbNullChar)
        lReturn = GetPrivateProfileSection(m_sSection, _
                                           sBuffer, _
                                           lBuffSize, _
                                           m_sFileName)
        If lReturn = (lBuffSize - 2) Then
            iTries = iTries + 1
            lBuffSize = lBuffSize + 1           '// readjust to pair boundary
        Else
            Exit Do
        End If
        sBuffer = ""      '// clear if this is not the first time
    Loop While iTries < 4

    If Len(sBuffer) > 3 Then
        '// Trim the resulting buffer
        sBuffer = Left$(sBuffer, lReturn + 1)
        Do
            If iCounter = -1 Then
                iCounter = 0
                ReDim asEnumKeys(0 To 0)
            Else
                ReDim Preserve asEnumKeys(0 To iCounter)
            End If
        
            'pull up to the first Null
            sValuePair = Mid$(sBuffer, 1, InStr(sBuffer, vbNullChar) - 1)
            If Len(sValuePair) > 0 Then
                iDelim = InStr(sValuePair, "=")
                If iDelim > 0 Then
                    asEnumKeys(iCounter) = Left$(sValuePair, iDelim - 1)
                End If
                        
            End If
            
            'trim the string to the next null
            sBuffer = Mid$(sBuffer, InStr(sBuffer, vbNullChar) + 1)
            iCounter = iCounter + 1
        Loop While Len(sBuffer) > 3
    End If
finally:
    '// See if we have success
    If iCounter > -1 Then
        ArrayResult = asEnumKeys()
    Else
        ArrayResult = Null
    End If
    EnumSectionKeys = iCounter
    Exit Function
catch:
    '// Ah, hello :)
    Debug.Assert 0
    iCounter = -1
    Resume finally
End Function

'===============================================================================
' Name: Function EnumSectionValues
' Input:
'   ByRef ArrayResult As Variant - Returned array
' Output:
'   Integer - Returns 0 if failure
' Purpose: This proc will enumerate just the values contained under a given section
' and will return them in an array.
' For how the class handles buffer sizes on INI calls, please see the Notes
' section on the [Declarations] section of the class module.
' If the method fails, the return will be zero, the ArrayResult argument
' will be set to Null and you will hit an assert.
' Remarks: None
'===============================================================================
Public Function EnumSectionValues(ByRef ArrayResult As Variant) As Integer

    Dim sBuffer As String
    Dim lReturn As Long
    Dim sValuePair As String
    Dim iCounter As Integer
    Dim asEnumValues() As String
    Dim lBuffSize As Long
    Dim iDelim As Integer
    Dim iTries As Integer

    'Debug.Assert 0
    On Error GoTo catch
    iCounter = -1
    lBuffSize = MAX_BUFFER_SIZE   '// start out with 4096 bytes
    Do
        lBuffSize = (lBuffSize * 2)
        lBuffSize = lBuffSize - 1       '// otherwise we overflow 32767
        sBuffer = String$(lBuffSize, vbNullChar)
        lReturn = GetPrivateProfileSection(m_sSection, _
                                           sBuffer, _
                                           lBuffSize, _
                                           m_sFileName)
        If lReturn = (lBuffSize - 2) Then
            iTries = iTries + 1
            lBuffSize = lBuffSize + 1           '// readjust to pair boundary
        Else
            Exit Do
        End If
        sBuffer = ""      '// clear if this is not the first time
    Loop While iTries < 4
    If ((Len(sBuffer) > 0) And (lReturn <> 0)) Then
        sBuffer = Left$(sBuffer, lReturn - 1)
        sBuffer = sBuffer & vbNullChar
        Do
            If iCounter = -1 Then
                iCounter = 0
                ReDim asEnumValues(0 To 0)
            Else
                ReDim Preserve asEnumValues(0 To iCounter)
            End If
            'pull up to the first Null
            sValuePair = Mid$(sBuffer, 1, InStr(sBuffer, vbNullChar) - 1)
            If Len(sValuePair) > 0 Then
                'Parse the string
                iDelim = InStr(sValuePair, "=")
                If iDelim > 0 Then
                    asEnumValues(iCounter) = Right$(sValuePair, (Len(sValuePair) - iDelim))
                End If
            End If
            
            'trim the string to the next null
            sBuffer = Mid$(sBuffer, InStr(sBuffer, vbNullChar) + 1)
            iCounter = iCounter + 1
        Loop While Len(sBuffer) > 3
    End If
finally:

    '// See if we have success
    If iCounter > -1 Then
        ArrayResult = asEnumValues()
    Else
        ArrayResult = Null
    End If
    EnumSectionValues = iCounter
    Exit Function
catch:
    '// Ah, hello :)
    Debug.Assert 0
    iCounter = -1
    Resume finally
End Function
'===============================================================================
' Name: Function EnumSections
' Input:
'   ByRef ArrayResult As Variant - Returned array
' Output:
'   Integer - Returns 0 if failure
' Purpose: This proc will enumerate all the names of the sections of the current INI
' file and return them in an array.
' <p>For how the class handles buffer sizes on INI calls, please see the Notes
' section on the [Declarations] section of the class module.
' <p>If the method fails, the return will be zero, the ArrayResult argument
' will be set to Null and you will hit an assert.
' Remarks: None
'===============================================================================
Public Function EnumSections(ByRef ArrayResult As Variant) As Integer
    On Error GoTo catch
    Dim sBuffer As String
    Dim lReturn As Long
    Dim sValue As String
    Dim iCounter As Integer
    Dim asEnumSections() As String
    Dim lBuffSize As Long
    Dim iTries As Integer

    'Debug.Assert 0
    On Error GoTo catch
    iCounter = -1
    lBuffSize = MAX_BUFFER_SIZE   '// start out with 4096 bytes
    Do
        lBuffSize = (lBuffSize * 2)
        lBuffSize = lBuffSize - 1       '// otherwise we overflow 32767
        sBuffer = String$(lBuffSize, vbNullChar)
        lReturn = GetPrivateProfileSectionNames(sBuffer, _
                                                lBuffSize, _
                                                m_sFileName)
        If lReturn = (lBuffSize - 2) Then
            iTries = iTries + 1
            lBuffSize = lBuffSize + 1           '// readjust to pair boundary
        Else
            Exit Do
        End If
        sBuffer = ""      '// clear if this is not the first time
    Loop While iTries < 4

    If ((Len(sBuffer) > 0) And (lReturn <> 0)) Then
        sBuffer = Left$(sBuffer, lReturn + 1)     '// Must allow for an extra NULL
        Do
            If iCounter = -1 Then
                iCounter = 0
                ReDim asEnumSections(0 To 0)
            Else
                ReDim Preserve asEnumSections(0 To iCounter)
            End If
        
            'pull up to the first Null
            sValue = Mid$(sBuffer, 1, InStr(sBuffer, vbNullChar) - 1)
            If Len(sValue) > 0 Then
                asEnumSections(iCounter) = sValue
                        
            End If
            
            'trim the string to the next null
            sBuffer = Mid$(sBuffer, InStr(sBuffer, vbNullChar) + 1)
            iCounter = iCounter + 1
        Loop While Len(sBuffer) > 3
    Else
        EnumSections = 0
        Exit Function
    End If
finally:

    '// See if we have success
    If iCounter > -1 Then
        ArrayResult = asEnumSections()
    Else
        ArrayResult = Null
    End If

    EnumSections = iCounter
    Exit Function
catch:
    '// Ah, hello :)
    Debug.Assert 0
    iCounter = -1
    Resume finally
End Function

'===============================================================================
' Name: Sub DeleteSection
' Input:
'   ByVal SectionName As String - Name of the section to be deleted
' Output: None
' Purpose: Deletes a section from the current INI file
' Remarks: None
'===============================================================================
Public Sub DeleteSection(ByVal SectionName As String)
    Call WritePrivateProfileString(SectionName, 0&, 0&, m_sFileName)
End Sub
'===============================================================================
' Name: Sub DeleteKey
' Input:
'   ByVal KeyName As String - Name of the key to be deleted
' Output: None
' Purpose: Deletes a key (a value pair) from the INI file
' Remarks: None
'===============================================================================
Public Sub DeleteKey(ByVal KeyName As String)
    Call WritePrivateProfileString(m_sSection, KeyName, 0&, m_sFileName)
End Sub
'===============================================================================
' Name: Function EnumSectionValuePairs
' Input:
'   ByRef ArrayResult As Variant - Returned array
' Output:
'   Integer - Returns 0 if failure
' Purpose: This proc will enumerate a given section's Key=Value and place them
' in the passed array as two different arrays. That is:
' <p>    Array(Array(Keys),Array(Values))
' <p>For how the class handles buffer sizes on INI calls, please see the Notes
' section on the [Declarations] section of the class module.
' <p>If the method fails, the return will be zero, the ArrayResult argument
' will be set to Null and you will hit an assert.
' Remarks: None
'===============================================================================
Public Function EnumSectionValuePairs(ByRef ArrayResult As Variant) As Integer
    On Error GoTo catch
    
    Dim sBuffer As String
    Dim lReturn As Long
    Dim sValuePair As String
    Dim iCounter As Integer
    Dim arrEnumValues() As String
    Dim arrEnumKeys() As String
    Dim lBuffSize As Long
    Dim iDelim As Integer
    Dim iTries As Integer

    'Debug.Assert 0
    On Error GoTo catch
    iCounter = -1
    lBuffSize = MAX_BUFFER_SIZE   '// start out with 4096 bytes
    Do
        lBuffSize = (lBuffSize * 2)
        lBuffSize = lBuffSize - 1       '// otherwise we overflow 32767
        sBuffer = String$(lBuffSize, vbNullChar)
        lReturn = GetPrivateProfileSection(m_sSection, _
                                           sBuffer, _
                                           lBuffSize, _
                                           m_sFileName)
        If lReturn = (lBuffSize - 2) Then
            iTries = iTries + 1
            lBuffSize = lBuffSize + 1           '// readjust to pair boundary
        Else
            Exit Do
        End If
        sBuffer = ""      '// clear if this is not the first time
    Loop While iTries < 4

    If ((Len(sBuffer) > 0) And (lReturn <> 0)) Then
        sBuffer = Left$(sBuffer, lReturn)
        Do
            If iCounter = -1 Then
                iCounter = 0
                ReDim arrEnumValues(0 To 0)
                ReDim arrEnumKeys(0 To 0)
            Else
                'top-redim the array
                ReDim Preserve arrEnumKeys(0 To iCounter)
                ReDim Preserve arrEnumValues(0 To iCounter)
            End If
            'pull up to the first Null
            sValuePair = Mid$(sBuffer, 1, InStr(sBuffer, vbNullChar) - 1)
            If Len(sValuePair) > 0 Then
                'Parse the string
                iDelim = InStr(sValuePair, "=")
                If iDelim > 0 Then
                    arrEnumKeys(iCounter) = Left$(sValuePair, iDelim - 1)
                    arrEnumValues(iCounter) = Right$(sValuePair, (Len(sValuePair) - iDelim))
                End If
            End If
            
            'trim the string to the next null
            sBuffer = Mid$(sBuffer, InStr(sBuffer, vbNullChar) + 1)
            iCounter = iCounter + 1
        Loop While Len(sBuffer) > 3
    End If
finally:

    '// See if we have success
    If iCounter > -1 Then
        ArrayResult = Array(arrEnumKeys(), arrEnumValues())
    Else
        ArrayResult = Array() ' return empty array
    End If
    EnumSectionValuePairs = iCounter
    Exit Function
catch:
    '// Ah, hello :)
    Debug.Print Err.Number, Err.Description
    Debug.Assert 0
    iCounter = -1
    Resume finally
End Function
'===============================================================================
' Name: Sub Flush
' Input: None
' Output: None
' Purpose: Flushes the INI file cache for the current file. <p>Note that
' the INI cache in Win32 (and Win16) is notoriously kranky. This flush
' is a logical one, not a physical flush. So if you need this type of
' thing seriously use the registry instead.
' Remarks: None
'===============================================================================
Public Sub Flush()
    Call WritePrivateProfileString(0&, 0&, 0&, m_sFileName)
End Sub


'===============================================================================
' Name: Property Get Values
' Input:
'   ByVal ValueName As String - Key name
' Output:
'   Variant - Key value
' Purpose: Retrieves a value from a key as a variant (string subtype) from the current section.
' Remarks: None
'===============================================================================
Public Property Get Values(ByVal ValueName As String) As Variant
    Dim szBuffer As String
    Dim lReturn As Long
    Dim sDefault As String

    szBuffer = String$(MAX_BUFFER_SIZE - 1, vbNullChar)
    sDefault = ""
    lReturn = GetPrivateProfileString(m_sSection, ValueName, sDefault, szBuffer, Len(szBuffer), m_sFileName)

    If lReturn > 0 Then
        Values = Left$(szBuffer, lReturn)
    Else
        Values = sDefault
    End If
End Property



'===============================================================================
' Name: Property Let Values
' Input:
'   ByVal ValueName As String - Key name
'   ByVal rValue As Variant - Key value
' Output: None
' Purpose: Sets the value of a key in the current section. Note that the passed value
' (a variant) is converted to a string before saving. This means that, for
' example, if you're passing a boolean value and you don't want a '-1' converted
' to a "True" literal, you'll have to convert the boolean to an integer or
' long before calling this method.
' Remarks: None
'===============================================================================
Public Property Let Values(ByVal ValueName As String, ByVal rValue As Variant)
    Call WritePrivateProfileString(m_sSection, ValueName, CStr(rValue), m_sFileName)
End Property


'===============================================================================
' Name: Property Get File
' Input: None
' Output:
'   String - Current file name
' Purpose: Returns the current filename
' Remarks: None
'===============================================================================
Public Property Get File() As String
    File = m_sFileName
End Property


'===============================================================================
' Name: Property Let File
' Input:
'   ByVal rValue As String - File name used
' Output: None
' Purpose: Sets the filename that all methods will use
' by default. Note that if the passed string is not a FQP,
' the INI file must be in the Windows search path somewhere.
' Remarks: None
'===============================================================================
Public Property Let File(ByVal rValue As String)
    ' TODO: Check if File exists.
    ' If file not found:
    ' Set_locale regionalSymbol
    ' Err.Raise vbObjectError, , "File not found: " & rValue
    m_sFileName = rValue
    m_sSection = ""       '// clear the section
End Property

'===============================================================================
' Name: Property Get Section
' Input: None
' Output:
'   String - Name of the current section
' Purpose: Retrieves the name of the current section
' Remarks: None
'===============================================================================
Public Property Get Section() As String
    Section = m_sSection
End Property


'===============================================================================
' Name: Property Let Section
' Input:
'    ByVal rValue As String - Current section name
' Output: None
' Purpose: Sets the name of the current section
' Remarks: None
'===============================================================================
Public Property Let Section(ByVal rValue As String)
    m_sSection = rValue
End Property



'===============================================================================
' Name: Sub WriteSection
' Input:
'   ByVal rValue As Variant - variant array
' Output: None
' Purpose: Accepts a variant array and writes the contents to
' the current section. Note that this will overwrite
' all of the existing key-value pairs under that section.
' The array must be structured as follows:
' <p>    rValue(n) = "KeyName = Value"
' <p>where n is a given index of the array. The "=" literal between the
' KeyName and Value *must* be present, or the call will fail.
' Remarks: None
'===============================================================================
Public Sub WriteSection(ByVal rValue As Variant)
    Dim sBuffer As String
    Dim a&

    If Not IsArray(rValue) Then Exit Sub

    For a& = LBound(rValue) To UBound(rValue)
        sBuffer = sBuffer & rValue(a&) & vbNullChar
    Next a&

    sBuffer = sBuffer & String$(2, vbNullChar)
    Call WritePrivateProfileSection(m_sSection, sBuffer, m_sFileName)
End Sub
