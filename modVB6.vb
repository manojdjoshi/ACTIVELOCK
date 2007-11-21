Module modVB6
    Dim TheFileName As String = "modALVB6.bas"
    Dim FilePath As String = Nothing
    Public Sub CreateVB6Module()
        Dim MyDialogResult As DialogResult
        frmMain.FolderBrowserDialog.SelectedPath = Application.StartupPath
        frmMain.FolderBrowserDialog.Description = "Please Select a folder To Save The File To"
        MyDialogResult = frmMain.FolderBrowserDialog.ShowDialog
        If MyDialogResult = Windows.Forms.DialogResult.OK Then
            FilePath = frmMain.FolderBrowserDialog.SelectedPath
        Else
            MsgBox("You Have To Select A Folder To Save The File To First", MsgBoxStyle.Critical)
            Exit Sub
        End If
        CreateFile()
        WriteToFile(My.Resources.VB6FormHeader)
        WriteToFile(My.Resources.VB6FormTop)
        CreateRoutineInitActivelock()
        WriteToFile(My.Resources.VB6FormRoutines)
        MsgBox("File Created Successfully, The New FileName Is: " & TheFileName, MsgBoxStyle.Information)
    End Sub
    Private Sub CreateFile()
        My.Computer.FileSystem.WriteAllText(FilePath & "\" & TheFileName, String.Empty, False)
    End Sub
    Private Sub WriteToFile(ByVal Data As String)
        My.Computer.FileSystem.WriteAllText(FilePath & "\" & TheFileName, Data, True)
    End Sub
    Private Sub CreateRoutineInitActivelock()

    End Sub
End Module
