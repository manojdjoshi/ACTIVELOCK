Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Ok As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.Ok = New System.Windows.Forms.Button
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(48, 64)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(208, 16)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Ok will activate test.  To end close !"
        '
        'Ok
        '
        Me.Ok.Location = New System.Drawing.Point(88, 96)
        Me.Ok.Name = "Ok"
        Me.Ok.TabIndex = 2
        Me.Ok.Text = "Ok"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(32, 8)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(240, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "1. Test Template (development) version"
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(32, 32)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(240, 16)
        Me.Label3.TabIndex = 4
        Me.Label3.Text = "2. Test real Code from temp.c"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(280, 130)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Ok)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)

    End Sub

#End Region
    ' All procedure beginning with template are development of the code to be produced
    ' so we test the development code
    ' then we test the code produced by StringHider in temp.c, should be the same as template code
    ' but lets test it

    ' This version uses MFCSample = testapp 1.0

    ' Rotate a Long to the right the specified number of times
    ' Treats long as a psudo unsigned int

    Function TemplateRotateRight(ByVal value As Long, ByVal times As Long) As Long
        Dim i As Long, signBits As Long

        ' no need to rotate more times than required
        times = times Mod 32
        ' return the number if it's a multiple of 32
        If times = 0 Then TemplateRotateRight = value : Exit Function

        For i = 1 To times
            ' remember the sign bit and bit 0
            signBits = value And &H80000001
            ' clear those bits and shift to the right by one position
            value = (value And &H7FFFFFFE) \ 2
            ' if the number was negative, then re-insert the bit
            If signBits And &H80000000 Then
                value = value Or &H40000000
            End If
            ' if bit 0 was set, then set the sign bit
            If signBits And 1 Then
                value = value Or &H80000000
            End If
        Next
        TemplateRotateRight = value
    End Function


    Function TemplateGetIt() As String

        ' The idea is not to have the following 2 lines in the program
        ' so remove them. Only here too check that the procedure works. When satisfied remove them or convert to comment
        Dim strLicOrig As String
        strLicOrig = "AAAAB3NzaC1yc2EAAAABJQAAAIB8/B2KWoai2WSGTRPcgmMoczeXpd8nv0Y4r1sJ1wV3vH21q4rTpEYuBiD4HFOpkbNBSRdpBHJGWec7jUi8ISV0pM6i2KznjhCms5CEtYHRybbiYvRXleGzFsAAP817PLN3JYo3WkErT2ofR5RCkfhmx060BT8waPoqnn3AB7sZ0Q=="
        ' end of remove section

        Const charsPerInt = 4
        Const licSize = 200
        Const splitSize = licSize / charsPerInt
        Const ileft = 0
        Dim untwist As Integer() = { _
      2, 37, 29, 25, 21, 4, 40, 13, 49, 14 _
    , 42, 44, 30, 0, 19, 22, 36, 27, 18, 10 _
    , 17, 9, 43, 15, 24, 6, 47, 7, 26, 5 _
    , 11, 33, 46, 41, 12, 31, 35, 38, 16, 45 _
    , 23, 3, 8, 48, 39, 32, 20, 1, 34, 28 _
   }
        Dim licInt As Long() = { _
      -596588320, -2107188004, -2105376126, -857840472, -2105367916, -587950492, 1858521774, 1621927570, -2036045148, -526480240 _
    , -357397792, -628698924, -1331368782, -1771797410, -1901679004, -523721562, 1721538720, 1753797252, -1461425950, 1756520684 _
    , -488726334, -2071821694, -1796840732, -460663122, -1902866300, -2104859450, -764634400, 1650757868, 2054857312, -228423998 _
    , -1328876346, -191968552, -294606716, -1970902298, -1259966844, -2105350516, 1722609250, -191076732, 1851945120, 1617715440 _
    , 1887736450, -758856462, -962550616, -2070100778, -560276786, 1725870740, -1534020888, 1892854484, -623850282, -758980946 _
   }

        Dim stLic As String
        Dim I As Integer
        Dim Iv As Long
        For I = 0 To splitSize - 1
            Iv = untwist(I)
            Iv = licInt(Iv)
            Iv = RotateRight(Iv, 1)
            Dim Ic As Integer
            For Ic = 0 To charsPerInt - 1
                stLic &= Chr(Iv And 255)
                Iv = Iv >> 8
            Next Ic
        Next I
        If ileft Then
            stLic = Microsoft.VisualBasic.Left(stLic, Len(stLic) - (charsPerInt - ileft))
        End If

        ' remove following section when all working
        If StrComp(stLic, strLicOrig) Then
            MsgBox("Failed to prroduce original license string. Two strings are in next box!")
            MsgBox(stLic & vbCrLf & strLicOrig)
        End If
        ' end of remove section

        TemplateGetIt = stLic
    End Function
    ' code here is all in one lump. Seperate and Hide ...
    ' Advice. If you are new to ActiveLock, then get it working by just using the long license key first
    ' When that is working implement this as its own stage
    ' Remember if you produce a new key, eg a new version YOU MUST redo this step

    ' Rotate a Long to the right the specified number of times               
    ' Treats long as a psudo unsigned int                                    

    Function RotateRight(ByVal value As Long, ByVal times As Long) As Long
        Dim i As Long, signBits As Long

        ' no need to rotate more times than required                         
        times = times Mod 32
        ' return the number if it's a multiple of 32                         
        If times = 0 Then RotateRight = value : Exit Function

        For i = 1 To times
            ' remember the sign bit and bit 0                                
            signBits = value And &H80000001
            ' clear those bits and shift to the right by one position        
            value = (value And &H7FFFFFFE) \ 2
            ' if the number was negative, then re-insert the bit             
            If signBits And &H80000000 Then
                value = value Or &H40000000
            End If
            ' if bit 0 was set, then set the sign bit                        
            If signBits And 1 Then
                value = value Or &H80000000
            End If
        Next
        RotateRight = value
    End Function
    Function GetIt() As String
        ' The idea is not to have the following 2 lines in the program                          
        ' so remove them. Only here too check that the procedure works. When satisfied remove them or convert to comment
        Dim strLicOrig As String
        strLicOrig = "AAAAB3NzaC1yc2EAAAABJQAAAIB8/B2KWoai2WSGTRPcgmMoczeXpd8nv0Y4r1sJ1wV3vH21q4rTpEYuBiD4HFOpkbNBSRdpBHJGWec7jUi8ISV0pM6i2KznjhCms5CEtYHRybbiYvRXleGzFsAAP817PLN3JYo3WkErT2ofR5RCkfhmx060BT8waPoqnn3AB7sZ0Q=="
        ' end of remove section
        Const charsPerInt = 4
        Const licSize = 200
        Const splitSize = licSize / charsPerInt
        Const ileft = 0
        Dim untwist As Integer() = { _
      2, 37, 29, 25, 21, 4, 40, 13, 49, 14 _
    , 42, 44, 30, 0, 19, 22, 36, 27, 18, 10 _
    , 17, 9, 43, 15, 24, 6, 47, 7, 26, 5 _
    , 11, 33, 46, 41, 12, 31, 35, 38, 16, 45 _
    , 23, 3, 8, 48, 39, 32, 20, 1, 34, 28 _
   }
        Dim licInt As Long() = { _
      -596588320, -2107188004, -2105376126, -857840472, -2105367916, -587950492, 1858521774, 1621927570, -2036045148, -526480240 _
    , -357397792, -628698924, -1331368782, -1771797410, -1901679004, -523721562, 1721538720, 1753797252, -1461425950, 1756520684 _
    , -488726334, -2071821694, -1796840732, -460663122, -1902866300, -2104859450, -764634400, 1650757868, 2054857312, -228423998 _
    , -1328876346, -191968552, -294606716, -1970902298, -1259966844, -2105350516, 1722609250, -191076732, 1851945120, 1617715440 _
    , 1887736450, -758856462, -962550616, -2070100778, -560276786, 1725870740, -1534020888, 1892854484, -623850282, -758980946 _
   }
        Dim stLic As String
        Dim I As Integer
        Dim Iv As Long
        For I = 0 To splitSize - 1
            Iv = untwist(I)
            Iv = licInt(Iv)
            Iv = RotateRight(Iv, 1)
            Dim Ic As Integer
            For Ic = 0 To charsPerInt - 1
                stLic &= Chr(Iv And 255)
                Iv = Iv >> 8
            Next Ic
        Next I
        If ileft Then
            stLic = Microsoft.VisualBasic.Left(stLic, Len(stLic) - (charsPerInt - ileft))
        End If
        ' remove following section when all working   
        If StrComp(stLic, strLicOrig) Then
            MsgBox("Failed to produce original license string")
            MsgBox(stLic & vbCrLf & strLicOrig)
        End If
        ' end of remove section                   
        GetIt = stLic
    End Function


    Private Sub Ok_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Ok.Click
        MsgBox("**** Ok For Template  ****" & vbCrLf & TemplateGetIt())
        MsgBox("**** Ok For Real Test ****" & vbCrLf & GetIt())
    End Sub
End Class
