VERSION 5.00
Begin VB.Form frmHDDSerialVB6 
   Caption         =   "Form1"
   ClientHeight    =   2475
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2475
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Get HD Firmware"
      Height          =   1095
      Left            =   450
      TabIndex        =   0
      Top             =   360
      Width           =   3750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   405
      TabIndex        =   1
      Top             =   1620
      Width           =   3795
   End
End
Attribute VB_Name = "frmHDDSerialVB6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Dim sDHNo As String
    sDHNo = GetSerialNumberFromWMI("Win32_PhysicalMedia")
    If sDHNo = "" Then sDHNo = GetSerialNumberFromWMI("Win32_DiskDrive")
    Label1.Caption = sDHNo
End Sub

Private Function GetSerialNumberFromWMI(wmi_selection As String) As String
    Dim o As Integer
    Dim sHDNoHex, reversedStr As String
    Dim sHDNoHexToChar As String
    Dim str1, str2 As String
    Dim jj As Integer
    Dim svc As Object
    Dim objEnum As WbemScripting.SWbemObjectSet
    Dim obj As WbemScripting.SWbemObject
    Set svc = GetObject("winmgmts:root\cimv2")
    Set objEnum = svc.ExecQuery("select * from " & wmi_selection)
    
    ' Check SerialNumber property
    For Each obj In objEnum
        Dim I As WbemScripting.SWbemProperty

        For Each I In obj.Properties_
            If IsNull2(I.Name, "") = "SerialNumber" Then
                'Debug.Print i.Value
                sHDNoHex = IsNull2(I.Value, "")
            End If
        Next I
    Next obj
    If Len(sHDNoHex) > 0 Then
        sHDNoHexToChar = ""
        For o = 1 To Len(sHDNoHex) Step 2
            sHDNoHexToChar = sHDNoHexToChar & Chr(CDec(("&H" & Trim(Mid(sHDNoHex, o, 2)))))
        Next
        reversedStr = ""
        For jj = 0 To Len(sHDNoHexToChar) / 2
            str1 = Mid(sHDNoHexToChar, jj * 2 + 1, 1)
            str2 = Mid(sHDNoHexToChar, jj * 2 + 2, 1)
            reversedStr = reversedStr & str2 & str1
        Next
        GetSerialNumberFromWMI = StripControlChars(Trim(reversedStr), False)
    Else
        GetSerialNumberFromWMI = sHDNoHex
    End If

    Dim mPos  As Integer
    Dim k As Integer
    Dim mChar As String
    Dim mChars As String
    Dim mSerial As String
    ' Check PNPDevideID property
    If GetSerialNumberFromWMI = "" Then
        For Each obj In objEnum
            Dim ii As WbemScripting.SWbemProperty
    
            For Each ii In obj.Properties_
                If IsNull2(ii.Name, "") = "PNPDeviceID" Then
                    'Debug.Print ii.Value
                    If Left(ii.Value, 3) = "IDE" Or Left(ii.Value, 4) = "SCSI" Then
                        sHDNoHex = IsNull2(ii.Value, "")
                        Exit For
                    End If
                End If
            Next ii
        Next obj
        
        mPos = InStrRev(sHDNoHex, "\")
        If mPos > 0 Then
            sHDNoHex = Mid(sHDNoHex, mPos + 1)

            ' Strip & characters
            Dim Index As Long
            Dim Bytes() As Byte
            
            Bytes() = sHDNoHex
            For Index = 0 To UBound(Bytes) Step 2
                If Bytes(Index) = 38 Then
                    Bytes(Index) = 0
                End If
            Next
            sHDNoHex = VBA.Replace(Bytes(), vbNullChar, "")
            mSerial = sHDNoHex

        End If
                
        GetSerialNumberFromWMI = StripControlChars(Trim(mSerial), False)
    End If
End Function

Function StripControlChars(Source As String, Optional KeepCRLF As Boolean = True) As String
    Dim Index As Long
    Dim Bytes() As Byte
    
    Bytes() = Source
    For Index = 0 To UBound(Bytes) Step 2
        If Bytes(Index) < 32 And Bytes(Index + 1) = 0 Then
            If Not KeepCRLF Or (Bytes(Index) <> 13 And Bytes(Index) <> 10) Then
                Bytes(Index) = 0
            End If
        End If
    Next
    StripControlChars = VBA.Replace(Bytes(), vbNullChar, "")
End Function

Private Function IsNull2(vValue As Variant, vReturnValue As Variant) As Variant
    If IsNull(vValue) = True Then
        IsNull2 = vReturnValue
    Else
        IsNull2 = vValue
    End If
End Function

