Attribute VB_Name = "Module1"


Public Declare Function GetVersionExA Lib "kernel32" _
               (lpVersionInformation As OSVERSIONINFO) As Integer
 
Public Type OSVERSIONINFO
               dwOSVersionInfoSize As Long
               dwMajorVersion As Long
               dwMinorVersion As Long
               dwBuildNumber As Long
               dwPlatformId As Long
               szCSDVersion As String * 128
End Type

'Public Declare Function ShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" (ByVal hWnd As Long, _
    ByVal lpOperation As String, ByVal lpFile As String, _
    ByVal lpParameters As String, ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) As Long

Public Function getwinOS() As Long
    Dim osinfo As OSVERSIONINFO
               Dim retvalue As Integer
 
               osinfo.dwOSVersionInfoSize = 148
               osinfo.szCSDVersion = Space$(128)
               retvalue = GetVersionExA(osinfo)
 
               With osinfo
               Select Case .dwPlatformId
 
                Case 1
                
                    Select Case .dwMinorVersion
                        Case 0
                            getversion = "Windows 95"
                        Case 10
                            getversion = "Windows 98"
                        Case 90
                            getversion = "Windows Millennium"
                    End Select
    
                Case 2
                    Select Case .dwMajorVersion
                        Case 3
                            getversion = "Windows NT 3.51"
                        Case 4
                            getversion = "Windows NT 4.0"
                        Case 5
                            If .dwMinorVersion = 0 Then
                                getversion = "Windows 2000"
                            Else
                                getversion = "Windows XP"
                            End If
                        Case 6
                            If .dwMinorVersion = 0 Then
                                getversion = "Windows Vista"
                            Else
                                getversion = ""
                            End If
                    End Select
    
                Case Else
                   getversion = "Failed"
                End Select
 
               End With
getwinOS = osinfo.dwMajorVersion
' getwinOS = UCase(getversion)
End Function





