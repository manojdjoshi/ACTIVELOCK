' *   ActiveLock
' *   Copyright 1998-2002 Nelson Ferraz
' *   Copyright 2003-2008 The ActiveLock Software Group (ASG)
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
' *
'===============================================================================
' Name: modActiveLock
' Purpose: Module used by ALUGEN.
' Functions:
' Properties:
' Methods:
' Started: 04.21.2005
' Modified: 08.15.2005
'===============================================================================
'
Option Strict On
Option Explicit On 

Imports System.Runtime.InteropServices
Imports ActiveLock37Net

Module modALUGEN

    Public Const ACTIVELOCKSTRING As String = "Activelock3"
    Public Const ACTIVELOCKSOFTWAREWEB As String = "http://www.activelocksoftware.com"

    Public ActiveLock3AlugenGlobals_definst As New AlugenGlobals
    Public ActiveLock3Globals_definst As New Globals
    Public storage() As storageType


    Public Structure PhaseType
        Dim exponential As Byte
        Dim startpoint As Byte
        Dim total As Byte
        Dim param As Byte
        Dim current As Byte
        Dim N As Byte ' if exponential */
        Dim Mult As Byte ' if linear */
    End Structure

    Public Const MAXPHASE As Integer = 5
    <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)> Public Structure ProgressType
        Dim nphases As Integer
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=MAXPHASE - 1)> Dim phases() As PhaseType
        Dim total As Byte
        Dim divisor As Byte
        Dim range As Byte
        Dim hwndProgbar As Integer

        Public Sub Initialize()
            ReDim phases(MAXPHASE - 1)
        End Sub
    End Structure

    Structure storageType
        <VBFixedArray(10)> Dim D As String()
        Public Sub Initialize()
            ReDim D(10)
        End Sub
    End Structure

    Public Function AppPath() As String
        Return System.Windows.Forms.Application.StartupPath
    End Function


    Public Class ProductInfoItem
        Private _ProductName As String
        Private _productVersion As String

        'constructors
        Public Sub New()
            _ProductName = ""
            _productVersion = ""
        End Sub
        Public Sub New(ByVal ProductName As String, ByVal ProductVersion As String)
            _ProductName = ProductName
            _productVersion = ProductVersion
        End Sub

        Public Property ProductName() As String
            Get
                Return _ProductName
            End Get
            Set(ByVal Value As String)
                _ProductName = Value
            End Set
        End Property

        Public Property ProductVersion() As String
            Get
                Return _productVersion
            End Get
            Set(ByVal Value As String)
                _productVersion = Value
            End Set
        End Property

        Public ReadOnly Property ProductNameVersion() As String
            Get
                Dim mProductname As String
                mProductname = _ProductName
                If _productVersion.Length > 0 Then
                    mProductname = mProductname & " (" & _productVersion & ")"
                End If
                Return mProductname
            End Get
        End Property
    End Class
    Public Function strLeft(ByVal vString As String, ByVal vLength As Integer) As String
        strLeft = vString.Substring(0, vLength)
    End Function
    Public Function strRight(ByVal vString As String, ByVal vLength As Integer) As String
        strRight = vString.Substring(vString.Length - vLength)
    End Function
    Public Function Is64Bit() As Boolean
        Dim x64 As Boolean = System.Environment.Is64BitOperatingSystem
        If x64 Then
            Return True
        Else
            Return False
        End If
    End Function

End Module