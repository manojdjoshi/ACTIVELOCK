Attribute VB_Name = "modRegisteredLevel"
Option Explicit

Public strRegisteredLevelDBName As String

Public Function AddBackSlash(ByVal sPath As String) As String

    'Returns sPath with a trailing backslash if sPath does not
    'already have a trailing backslash. Otherwise, returns sPath.
    sPath = Trim$(sPath)

    If Len(sPath) > 0 Then

        sPath = sPath & IIf(Right$(sPath, 1) <> "\", "\", "")

    End If

    AddBackSlash = sPath

End Function

Public Function FileExists(sFileName As String) As Boolean

    Dim fs As Object
    Set fs = CreateObject("Scripting.FileSystemObject")
    FileExists = fs.FileExists(sFileName)
    Set fs = Nothing

End Function

Public Sub SaveComboBox(FileName As String, _
                        Control As Object, _
                        Optional IsItemData As Boolean)
    Dim FileNum As Byte
    Dim Teller As Integer
    FileNum = FreeFile
    Open FileName For Output As #FileNum
    Print #FileNum, IsItemData

    For Teller = 0 To Control.ListCount - 1
        Print #FileNum, Control.List(Teller)

        If IsItemData = True Then 'Als Itemdata gebruikt is
            Print #FileNum, Control.ItemData(Teller)
        End If

    Next Teller

    Close #FileNum
End Sub

Public Sub LoadComboBox(FileName As String, _
                        Control As Object, _
                        IsItemData As Boolean)
    Dim FileNum As Byte
    Dim Teller As Integer
    Dim Text As String
    FileNum = FreeFile
    Open FileName For Input As #FileNum

    If LOF(FileNum) = 0 Then Close #FileNum: Exit Sub
    Line Input #FileNum, Text
    IsItemData = Text

    Do While Not EOF(FileNum)
        Line Input #FileNum, Text
        Control.AddItem Text

        If IsItemData = True Then
            Line Input #FileNum, Text
            Control.ItemData(Control.NewIndex) = Val(Text)
        End If

    Loop

    Close #FileNum
End Sub
