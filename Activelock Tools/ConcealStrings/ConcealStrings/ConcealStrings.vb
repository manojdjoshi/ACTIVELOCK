Public Class ConcealStrings
    Dim systemEvent As Boolean
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Dim i As Integer, j As Integer
        Dim conceal As String = String.Empty
        Dim s As String
        If systemEvent Then Exit Sub
        s = TextBox1.Text
        For i = 1 To Len(s)
            j = Asc(Mid$(s, i, 1))
            conceal = conceal & "Chr(" & j & ")"
            If i < s.Length Then conceal = conceal & " & "
        Next i
        systemevent = True
        TextBox2.Text = conceal
        systemevent = False

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        On Error GoTo exitHere
        Dim i As Integer, j As Integer
        Dim reveal As String = String.Empty
        Dim s As String = String.Empty
        If systemEvent Then Exit Sub
        s = TextBox2.Text
        If s.Length < 6 Then Exit Sub
        If s.ToLower.Substring(0, 3) <> "chr" Then Exit Sub
        Dim a As Object
        s = Replace(s.ToLower, "chr(", ")")
        s = Replace(s, "&", "")
        a = Split(s, ")")
        For i = 0 To UBound(a)
            j = 0
            If a(i) <> "" And Val(a(i)) >= 32 And Val(a(i)) <= 126 Then j = CType(a(i), Integer)
            If j >= 32 And j <= 126 Then reveal = reveal & Chr(j)
        Next i
        systemevent = True
        TextBox1.Text = reveal
        systemEvent = False
exitHere:
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TextBox1.Text = String.Empty
        TextBox2.Text = String.Empty
    End Sub
End Class
