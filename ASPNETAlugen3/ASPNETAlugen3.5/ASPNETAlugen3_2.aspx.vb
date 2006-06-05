Imports System.IO
Imports System.Diagnostics
Imports System.Text
Imports System.Runtime.InteropServices
Imports ActiveLock3_5NET

Public Class Form1
  Inherits System.Web.UI.Page
  Public ActiveLock3AlugenGlobals_definst As New AlugenGlobals
  Public ActiveLock3Globals_definst As New Globals_Renamed
  Public GeneratorInstance As _IALUGenerator

  Protected WithEvents Button2 As System.Web.UI.WebControls.Button
  Protected WithEvents Label10 As System.Web.UI.WebControls.Label
  Protected WithEvents Label11 As System.Web.UI.WebControls.Label
  Protected WithEvents Button3 As System.Web.UI.WebControls.Button
  Protected WithEvents txtLibFile As System.Web.UI.WebControls.TextBox
  Protected WithEvents Label9 As System.Web.UI.WebControls.Label
  Protected WithEvents Label1 As System.Web.UI.WebControls.Label
  Protected WithEvents txtLibKey As System.Web.UI.WebControls.TextBox
  Protected WithEvents cmdKeyGen As System.Web.UI.WebControls.Button
  Protected WithEvents txtUser As System.Web.UI.WebControls.TextBox
  Protected WithEvents Label2 As System.Web.UI.WebControls.Label
  Protected WithEvents txtReqCodeIn As System.Web.UI.WebControls.TextBox
  Protected WithEvents Label3 As System.Web.UI.WebControls.Label
  Protected WithEvents lblDays As System.Web.UI.WebControls.Label
  Protected WithEvents txtDays As System.Web.UI.WebControls.TextBox
  Protected WithEvents lblExpiry As System.Web.UI.WebControls.Label
  Protected WithEvents cmbLicType As System.Web.UI.WebControls.DropDownList
  Protected WithEvents Label6 As System.Web.UI.WebControls.Label
  Protected WithEvents cmbRegisteredLevel As System.Web.UI.WebControls.DropDownList
  Protected WithEvents Label8 As System.Web.UI.WebControls.Label
  Protected WithEvents cmbProds As System.Web.UI.WebControls.DropDownList
  Protected WithEvents Label7 As System.Web.UI.WebControls.Label
  Protected WithEvents cmdProductCodeGenerator As System.Web.UI.WebControls.Button
  'Public GeneratorInstance As ActiveLock3.IALUGenerator
  Public AL As _IActiveLock

    'Private WithEvents ActiveLockEventSink As ActiveLock3.ActiveLockEventNotifier

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

  End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

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
    ' and Key Gen tab with product info from products.ini
    Dim arrProdInfos() As ProductInfo
    Dim GeneratorInstance As _IALUGenerator
    Dim MyGen As New AlugenGlobals
    GeneratorInstance = MyGen.GeneratorInstance()
    GeneratorInstance.StoragePath = AppPath() & "\products.ini"

    arrProdInfos = GeneratorInstance.RetrieveProducts()
    If arrProdInfos.Length = 0 Then Exit Sub

    Dim i As Short
    For i = LBound(arrProdInfos) To UBound(arrProdInfos)
      PopulateUI(arrProdInfos(i))
    Next
  End Sub
  Private Sub PopulateUI(ByVal ProdInfo As ProductInfo)
    With ProdInfo
      AddRow(.name, .Version, .VCode, .GCode, False)
    End With
  End Sub
  Private Sub AddRow(ByRef Name As String, ByRef Ver As String, ByRef Code1 As String, ByRef Code2 As String, Optional ByRef fUpdateStore As Boolean = True)
    ' Call ActiveLock3NET.IALUGenerator to add product
    Dim ProdInfo As ProductInfo
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

    Dim varLicType As ProductLicense.ALLicType
    If cmbLicType.SelectedValue = "Periodic" Then
      varLicType = ProductLicense.ALLicType.allicPeriodic
    ElseIf cmbLicType.SelectedValue = "Permanent" Then
      varLicType = ProductLicense.ALLicType.allicPermanent
    ElseIf cmbLicType.SelectedValue = "Time Locked" Then
      varLicType = ProductLicense.ALLicType.allicTimeLocked
    Else
      varLicType = ProductLicense.ALLicType.allicNone
    End If

    'With AL
    '    .SoftwareName = strName
    '    .SoftwareVersion = strVer
    '    .SoftwareCode = "AAAAB3NzaC1yc2EAAAABJQAAAIB8/B2KWoai2WSGTRPcgmMoczeXpd8nv0Y4r1sJ1wV3vH21q4rTpEYuBiD4HFOpkbNBSRdpBHJGWec7jUi8ISV0pM6i2KznjhCms5CEtYHRybbiYvRXleGzFsAAP817PLN3JYo3WkErT2ofR5RCkfhmx060BT8waPoqnn3AB7sZ0Q=="
    '    .LockType = 1   'lockNone
    '    .AutoRegisterKeyPath = AppPath() & "\ASPNETAlugen3.all"
    'End With

    AL.KeyStoreType = IActiveLock.LicStoreType.alsFile
    txtLibFile.Text = AppPath() & "\ASPNETAlugen3.all"

    AL.Init(AppPath() & "\bin")

    Dim MyAL As New Globals_Renamed
    Dim Generator As New AlugenGlobals

    GeneratorInstance = Generator.GeneratorInstance()
    GeneratorInstance.StoragePath = AppPath() & "\products.ini"

    Dim strExpire As String
    strExpire = GetExpirationDate()
    Dim strRegDate As String
    strRegDate = ActiveLockDateFormat(Now)

    Dim Lic As ProductLicense
    Lic = ActiveLock3Globals_definst.CreateProductLicense(strName, strVer, "", ProductLicense.LicFlags.alfSingle, varLicType, "", cmbRegisteredLevel.SelectedValue, strExpire, , strRegDate)
    Dim alugenproduct As ProductInfo
    alugenproduct = GeneratorInstance.RetrieveProduct(strName, strVer)
    Dim LibKey As String
    LibKey = GeneratorInstance.GenKey(Lic, txtReqCodeIn.Text, cmbRegisteredLevel.SelectedValue)
    txtLibKey.Text = Make64ByteChunks(LibKey)
  End Sub

  Private Sub cmdProductCodeGenerator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdProductCodeGenerator.Click
    'retrieve url info
    Dim siteNameUrl As String
    siteNameUrl = Request.ServerVariables.Get("SERVER_NAME")
    Select Case siteNameUrl
      Case "localhost"
        Response.Redirect("../ASPNETAlugen3/ASPNETAlugen3_1.aspx")
      Case Else
        Response.Redirect("ASPNETAlugen3_1.aspx")
    End Select
  End Sub
  Private Function GetExpirationDate() As String
    If cmbLicType.SelectedValue = "Time Locked" Then
      GetExpirationDate = txtDays.Text
    Else
      GetExpirationDate = ActiveLockDateFormat(System.DateTime.FromOADate(Now.ToOADate + CShort(txtDays.Text)))
    End If
  End Function
  Public Function AppPath() As String
    Dim siteNameUrl As String
    AppPath = System.IO.Path.GetDirectoryName(Server.MapPath("ASPNETAlugen3_2.aspx"))
  End Function

  Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

  End Sub
End Class

