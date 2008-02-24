Imports System.IO
Imports System.Diagnostics
Imports System.Text
Imports System.Runtime.InteropServices
Imports ActiveLock3_5NET

Namespace ASPNETAlugen3

Partial Class Form1
    Inherits System.Web.UI.Page

    Public ActiveLock3AlugenGlobals_definst As New AlugenGlobals
        Public ActiveLock3Globals_definst As New Globals
    Public GeneratorInstance As _IALUGenerator
    Public AL As ActiveLock3_5NET._IActiveLock
    Public mKeyStoreType As IActiveLock.LicStoreType
    Public mProductsStoreType As IActiveLock.ProductsStoreType
    Public mProductsStoragePath As String

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub


    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Not IsPostBack Then
            With cmbLicType
                .ClearSelection()
                .Items.Add("Time Locked")
                .Items.Add("Periodic")
                .Items.Add("Permanent")
            End With
            With cmbRegisteredLevel
                .ClearSelection()
                .Items.Add("0: Level 0")
                .Items.Add("1: Level 1")
                .Items.Add("2: Level 2")
                .Items.Add("3: Level 3")
                .Items.Add("4: Level 4")
                .Items.Add("5: Level 5")
                .Items.Add("6: Level 6")
                .Items.Add("7: Level 7")
                .Items.Add("8: Level 8")
                .Items.Add("9: Level 9")
                .Items.Add("10: Level 10")
                .Items.Add("11: Level 11")
                .Items.Add("12: Level 12")
                .Items.Add("13: Level 13")
                .Items.Add("14: Level 14")
                .Items.Add("15: Level 15")
                .Items.Add("16: Level 16")
                .Items.Add("17: Level 17")
                .Items.Add("18: Level 18")
                .Items.Add("19: Level 19")
                .Items.Add("20: Level 20")
                .Items.Add("21: Level 21")
                .Items.Add("22: Level 22")
                .Items.Add("23: Level 23")
                .SelectedIndex = 0
            End With
            txtDays.ReadOnly = False
            txtDays.BackColor = System.Drawing.SystemColors.Window
            lblExpiry.Text = "Expires on Date"
            Dim dtExpire As Date
            dtExpire = DateAdd(DateInterval.Day, 30, Now())
            txtDays.Text = dtExpire.Year & "/" & dtExpire.Month & "/" & dtExpire.Day()
            lblDays.Text = "YYYY/MM/DD"
            InitUI()

            End If
    End Sub

    Private Function Make64ByteChunks(ByVal strData As String) As String
        Dim i As Long
        Dim Count As Long
        Count = Len(strData)
        Dim sResult As String
        sResult = Strings.Left(strData, 64)
        i = 65
        While i <= Count
            sResult = sResult & vbCrLf & Mid$(strData, i, 64)
            i = i + 64
        End While
        Make64ByteChunks = sResult
    End Function
    Private Sub InitUI()
        ' Init Default license class
        cmbLicType.SelectedIndex = 0
        'With gridProds
        '    .Clear()
        '    .Rows = 1
        '    .FormatString = "Name                             |Version          | VCode                                          |GCode                                                                                  "
        '    .ColAlignment(0) = flexAlignLeftCenter
        '    .ColAlignment(1) = flexAlignLeftCenter
        '    .ColAlignment(2) = flexAlignLeftCenter
        '    .ColAlignment(3) = flexAlignLeftCenter
        'End With
        ' Populate Product List on Product Code Generator tab
            ' and Key Gen tab with product info from licenses.ini
        Dim arrProdInfos() As ProductInfo
        Dim GeneratorInstance As ActiveLock3_5NET._IALUGenerator
        Dim MyGen As New ActiveLock3_5NET.AlugenGlobals
        GeneratorInstance = MyGen.GeneratorInstance(mProductsStoreType)
            GeneratorInstance.StoragePath = AppPath() & "\licenses.ini"

        arrProdInfos = GeneratorInstance.RetrieveProducts()
        If arrProdInfos.Length = 0 Then Exit Sub

        Dim i As Short
        For i = LBound(arrProdInfos) To UBound(arrProdInfos)
            PopulateUI(arrProdInfos(i))
        Next
    End Sub
    Private Sub PopulateUI(ByVal ProdInfo As ProductInfo)
        With ProdInfo
            AddRow(.Name, .Version, .VCode, .GCode, False)
        End With
    End Sub
    Private Sub AddRow(ByRef Name As String, ByRef Ver As String, ByRef Code1 As String, ByRef Code2 As String, Optional ByRef fUpdateStore As Boolean = True)
        ' Call ActiveLock3_5NET.IALUGenerator to add product
        Dim ProdInfo As ActiveLock3_5NET.ProductInfo
        ProdInfo = ActiveLock3AlugenGlobals_definst.CreateProductInfo(Name, Ver, Code1, Code2)
        If fUpdateStore Then
            Call GeneratorInstance.SaveProduct(ProdInfo)
        End If
        cmbProds.Items.Add(Name & " - " & Ver)
        'cmdRemove.Enabled = True
    End Sub

    Private Sub cmbLicType_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbLicType.SelectedIndexChanged
        ' enable the days edit box
        If cmbLicType.SelectedValue = "Periodic" Or cmbLicType.SelectedValue = "Time Locked" Then
            txtDays.ReadOnly = False
            txtDays.BackColor = System.Drawing.SystemColors.Window
        Else
            txtDays.ReadOnly = True
            txtDays.BackColor = System.Drawing.SystemColors.Control
        End If
        If cmbLicType.SelectedValue = "Time Locked" Then
            lblExpiry.Text = "Expires on Date"
            Dim dtExpire As Date
            dtExpire = DateAdd(DateInterval.Day, 30, Now())
            txtDays.Text = dtExpire.Year & "/" & dtExpire.Month & "/" & dtExpire.Day()
            lblDays.Text = "YYYY/MM/DD"
        Else
            lblExpiry.Text = "   Expires after"
            txtDays.Text = "30"
            lblDays.Text = "Day(s)"
        End If

    End Sub
    Private Function ActiveLockDateFormat(ByVal dt As Date) As String
        ' Normalizes date format to yyyy/mm/dd
        ActiveLockDateFormat = Year(dt) & "/" & Month(dt) & "/" & Day(dt)
    End Function

    Private Sub cmbProds_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        UpdateKeyGenButtonStatus()
    End Sub
    Private Sub UpdateKeyGenButtonStatus()
        If txtReqCodeIn.Text = "" Then
            cmdKeyGen.Enabled = False
        Else
            If cmbProds.SelectedValue <> "" Then
                cmdKeyGen.Enabled = True
            End If
        End If
    End Sub

    Private Sub cmdKeyGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdKeyGen.Click

        AL = ActiveLock3Globals_definst.NewInstance()
        Dim arrProdVer() As String
        arrProdVer = Split(cmbProds.SelectedValue, "-")
        Dim strName, strVer As String
        strName = Trim(arrProdVer(0))
        strVer = Trim(arrProdVer(1))

        Dim varLicType As ActiveLock3_5NET.ProductLicense.ALLicType
        If cmbLicType.SelectedValue = "Periodic" Then
            varLicType = ActiveLock3_5NET.ProductLicense.ALLicType.allicPeriodic
        ElseIf cmbLicType.SelectedValue = "Permanent" Then
            varLicType = ActiveLock3_5NET.ProductLicense.ALLicType.allicPermanent
        ElseIf cmbLicType.SelectedValue = "Time Locked" Then
            varLicType = ActiveLock3_5NET.ProductLicense.ALLicType.allicTimeLocked
        Else
            varLicType = ActiveLock3_5NET.ProductLicense.ALLicType.allicNone
        End If

        'With AL
        '    .SoftwareName = strName
        '    .SoftwareVersion = strVer
        '    .SoftwareCode = "AAAAB3NzaC1yc2EAAAABJQAAAIB8/B2KWoai2WSGTRPcgmMoczeXpd8nv0Y4r1sJ1wV3vH21q4rTpEYuBiD4HFOpkbNBSRdpBHJGWec7jUi8ISV0pM6i2KznjhCms5CEtYHRybbiYvRXleGzFsAAP817PLN3JYo3WkErT2ofR5RCkfhmx060BT8waPoqnn3AB7sZ0Q=="
        '    .LockType = 1   'lockNone
        '    .AutoRegisterKeyPath = AppPath() & "\ASPNETAlugen3.all"
        'End With

        AL.KeyStoreType = ActiveLock3_5NET.IActiveLock.LicStoreType.alsFile
            txtLibFile.Text = AppPath() & "\" & strName & "_" & strVer & ".all"

        AL.Init(AppPath() & "\bin")

            'Dim MyAL As New Globals
            'Dim Generator As New AlugenGlobals

            Dim MyGen As New AlugenGlobals
            GeneratorInstance = MyGen.GeneratorInstance(IActiveLock.ProductsStoreType.alsINIFile)
            GeneratorInstance.StoragePath = AppPath() & "\licenses.ini"

            ' Get the current date format and save it to regionalSymbol variable
            Get_locale()
            ' Use this trick to temporarily set the date format to "yyyy/MM/dd"
            Set_locale("")

            Dim strExpire As String
        strExpire = GetExpirationDate()
        Dim strRegDate As String
            strRegDate = Date.UtcNow.ToString("yyyy/MM/dd")

        Dim Lic As ActiveLock3_5NET.ProductLicense
        Lic = ActiveLock3Globals_definst.CreateProductLicense(strName, strVer, "", ActiveLock3_5NET.ProductLicense.LicFlags.alfSingle, varLicType, "", cmbRegisteredLevel.SelectedValue, strExpire, , strRegDate)
        Dim alugenproduct As ActiveLock3_5NET.ProductInfo
        alugenproduct = GeneratorInstance.RetrieveProduct(strName, strVer)
        Dim LibKey As String
        LibKey = GeneratorInstance.GenKey(Lic, txtReqCodeIn.Text, cmbRegisteredLevel.SelectedValue)
            txtLibKey.Text = Make64ByteChunks(LibKey)

            Set_locale(regionalSymbol)

    End Sub

    Private Sub cmdProductCodeGenerator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdProductCodeGenerator.Click
        'retrieve url info
        Dim siteNameUrl As String
        siteNameUrl = Request.ServerVariables.Get("SERVER_NAME")
        Select Case siteNameUrl
            Case "localhost"
                    Response.Redirect("../ASPNETAlugen3.5 Simple/ASPNETAlugen3_1.aspx")
            Case Else
                Response.Redirect("ASPNETAlugen3_1.aspx")
        End Select
    End Sub
    Private Function GetExpirationDate() As String
            If cmbLicType.SelectedValue = "Time Locked" Then 'Time Locked
                If txtDays.Text.Trim.Length = 0 Then txtDays.Text = Date.UtcNow.AddDays(30).ToString("yyyy/MM/dd")
                GetExpirationDate = CType(txtDays.Text, DateTime).ToString("yyyy/MM/dd")
            Else
                If txtDays.Text.Trim.Length = 0 Then txtDays.Text = "30"
                GetExpirationDate = Date.UtcNow.AddDays(CShort(txtDays.Text)).ToString("yyyy/MM/dd")
            End If
        End Function
    Public Function AppPath() As String
            'Dim siteNameUrl As String
        AppPath = System.IO.Path.GetDirectoryName(Server.MapPath("ASPNETAlugen3_2.aspx"))
    End Function

        Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
            ' save the license key
            SaveLicenseKey(txtLibKey.Text, txtLibFile.Text)
        End Sub
        Private Sub SaveLicenseKey(ByVal sLibKey As String, ByVal sFileName As String)
            Dim hFile As Integer
            hFile = FreeFile()
            FileOpen(hFile, sFileName, OpenMode.Output)
            PrintLine(hFile, sLibKey)
            FileClose(hFile)

            Response.Clear()
            Response.ContentType = "text/plain"
            Response.AddHeader("content-disposition", "attachment; filename=" + IO.Path.GetFileName(sFileName))

            Response.TransmitFile(sFileName)
            Response.Flush()
            Response.End()
        End Sub

    End Class

End Namespace

