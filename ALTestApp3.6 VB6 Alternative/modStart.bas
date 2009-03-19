Attribute VB_Name = "modStart"
Option Explicit

Public objReg As New clsReg

Public Sub Main()
    objReg.Init
End Sub


Private Sub Command1_Click()
    Dim strText As String
    Dim i As Integer
    Dim iLen As Integer
    Dim strOutput As String
    
    strText = Text1.Text
    
    iLen = Len(strText)
    
    For i = 1 To iLen
                
       strOutput = strOutput & "Chr(" & Asc(Mid(strText, i, i)) & ")" & " & "
    Next
    
    Text2.Text = strOutput
End Sub

