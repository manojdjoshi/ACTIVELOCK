Imports ActiveLock3_5NET

Public Class ASPNETAlugen3_1
  Inherits System.Web.UI.Page
  Public ActiveLock3AlugenGlobals_definst As New AlugenGlobals
  Public ActiveLock3Globals_definst As New Globals_Renamed
  Public GeneratorInstance As _IALUGenerator
  Public ActiveLock As _IActiveLock
  Protected WithEvents lblUpdateStatus As System.Web.UI.WebControls.Label
  Protected WithEvents gridProds As System.Web.UI.WebControls.DataGrid
  Public dt As New DataTable
  Protected WithEvents ltlAlert As System.Web.UI.WebControls.Literal

  Private Declare Ansi Function WritePrivateProfileString _
      Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As String, _
      ByVal lpKeyName As String, ByVal lpString As String, _
      ByVal lpFileName As String) As Integer
  Private Declare Ansi Function FlushPrivateProfileString _
      Lib "kernel32.dll" Alias "WritePrivateProfileStringA" _
      (ByVal lpApplicationName As Integer, _
      ByVal lpKeyName As Integer, ByVal lpString As Integer, _
      ByVal lpFileName As String) As Integer

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

  End Sub
    Protected WithEvents Label11 As System.Web.UI.WebControls.Label
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button
    Protected WithEvents Label10 As System.Web.UI.WebControls.Label
    Protected WithEvents cmdLicenseKeyGenerator As System.Web.UI.WebControls.Button
    Protected WithEvents Label7 As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents txtName As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents Label4 As System.Web.UI.WebControls.Label
    Protected WithEvents txtCode1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCode2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents cmdAdd As System.Web.UI.WebControls.Button
    Protected WithEvents txtCode2_2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCode1_2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label5 As System.Web.UI.WebControls.Label
    Protected WithEvents Label6 As System.Web.UI.WebControls.Label
    Protected WithEvents Label8 As System.Web.UI.WebControls.Label
    Protected WithEvents Label9 As System.Web.UI.WebControls.Label
    Protected WithEvents cmdCodeGen As System.Web.UI.WebControls.Button
    Protected WithEvents cmdValidate As System.Web.UI.WebControls.Button
    Protected WithEvents cmdRemove As System.Web.UI.WebControls.Button
    Protected WithEvents txtVer As System.Web.UI.WebControls.TextBox

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
        'Put user code to initialize the page here
        ' Populate Product List on Product Code Generator tab
        ' and Key Gen tab with product info from products.ini

        Dim arrProdInfos() As ProductInfo
    Dim MyGen As New AlugenGlobals
        GeneratorInstance = MyGen.GeneratorInstance()
        GeneratorInstance.StoragePath = AppPath() & "\products.ini"

        arrProdInfos = GeneratorInstance.RetrieveProducts()
        If arrProdInfos.Length = 0 Then Exit Sub

        'Populate the DataGrid control

        'dt = New DataTable
        Dim dc As New DataColumn("Name", System.Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("Version", System.Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("VCode", System.Type.GetType("System.String"))
        dt.Columns.Add(dc)
        dc = New DataColumn("GCode", System.Type.GetType("System.String"))
        dt.Columns.Add(dc)

        Dim i As Short
        For i = LBound(arrProdInfos) To UBound(arrProdInfos)
            PopulateUI(arrProdInfos(i))
        Next

        gridProds.AutoGenerateColumns = False

        Dim dgc As New BoundColumn
        dgc.HeaderText = "Name"
        dgc.DataField = "Name"
        gridProds.Columns.Add(dgc)

        dgc = New BoundColumn
        dgc.HeaderText = "Version"
        dgc.DataField = "Version"
        dgc.ItemStyle.Wrap = True
        gridProds.Columns.Add(dgc)

        dgc = New BoundColumn
        dgc.HeaderText = "VCode"
        dgc.DataField = "VCode"
        dgc.ItemStyle.Wrap = True
        'dgc.ItemStyle.Width = Unit.Pixel(40)
        gridProds.Columns.Add(dgc)

        dgc = New BoundColumn
        dgc.HeaderText = "GCode"
        dgc.DataField = "GCode"
        dgc.ItemStyle.Wrap = True
        gridProds.Columns.Add(dgc)

        gridProds.DataSource = dt
        gridProds.DataBind()

        If Not IsPostBack Then
            gridProds.SelectedIndex = dt.Rows.Count - 1
            gridProds.SelectedItemStyle.BackColor = Color.CadetBlue
            txtName.Text = arrProdInfos(i - 1).name
            txtVer.Text = arrProdInfos(i - 1).Version
            txtCode1_2.Text = arrProdInfos(i - 1).VCode
            txtCode2_2.Text = arrProdInfos(i - 1).GCode

            ' The following code is for forcing a postback when the contents of a 
            ' textbox gets changed and it loses focus
            'Dim js As String
            'js = "javascript:" & Page.GetPostBackEventReference(Me, "@@@@@buttonPostBack") & ";"
            'txtName.Attributes.Add("onchange", js)
            'txtVer.Attributes.Add("onchange", js)

            ' The following Javascript KeyUp code will allow us disable the
            ' Generate button when there's no Product Name or Version entered
            txtName.Attributes.Add("OnKeyUp", "document.all." & cmdCodeGen.ClientID & ".disabled=(this.value=='')")
            txtVer.Attributes.Add("OnKeyUp", "document.all." & cmdCodeGen.ClientID & ".disabled=(this.value=='')")

        End If

    End Sub

    Private Sub cmdLicenseKeyGenerator_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdLicenseKeyGenerator.Click
        'retrieve url info
        Dim siteNameUrl As String
        siteNameUrl = Request.ServerVariables.Get("SERVER_NAME")
        Select Case siteNameUrl
            Case "localhost"
                Response.Redirect("../ASPNETAlugen3/ASPNETAlugen3_2.aspx")
            Case Else
                Response.Redirect("ASPNETAlugen3_2.aspx")
        End Select
    End Sub

    Private Sub cmdAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdAdd.Click
        If txtCode1.Text = "" Then
            Say("New product VCode is missing.")
            Exit Sub
        ElseIf txtCode2.Text = "" Then
            Say("New product GCode is missing.")
            Exit Sub
        End If
        If CheckDuplicate(txtName.Text, txtVer.Text) Then
            Say("Selected Product Name and Version already exist in the Product Listbox.")
            Exit Sub
        End If
        'This button add the new product into the list and into the INI file
        AddRow(txtName.Text, txtVer.Text, txtCode1.Text, txtCode2.Text, True)
        txtCode1.Text = ""
        txtCode2.Text = ""
        cmdValidate.Enabled = True
        cmdRemove.Enabled = True
    End Sub

    ' Add a Product Row to the GUI.
    ' If fUpdateStore is True, then product info is also saved to the store.
    Private Sub AddRow(ByRef Name As String, ByRef Ver As String, ByRef Code1 As String, ByRef Code2 As String, Optional ByRef fUpdateStore As Boolean = True)
    ' Update the view
    Call CreateDataSource(Name, Ver, Code1, Code2)

    Dim ProdInfo As ProductInfo
    ProdInfo = ActiveLock3AlugenGlobals_definst.CreateProductInfo(Name, Ver, Code1, Code2)
    If fUpdateStore Then
      Call GeneratorInstance.SaveProduct(ProdInfo)
      gridProds.DataSource = dt
      gridProds.DataBind()
      gridProds.SelectedIndex = dt.Rows.Count - 1
    End If
  End Sub
  Function CreateDataSource(ByVal Name As String, ByVal Ver As String, ByVal Code1 As String, ByVal Code2 As String) As ICollection

    Dim dr As DataRow = dt.NewRow()

    'Create new row and add data to it
    dr(0) = Name
    dr(1) = Ver
    dr(2) = Code1
    dr(3) = Code2

    ''Sample assignments that can be used elsewhere
    ''dr(0) = i
    ''dr(1) = "Item " + i.ToString()
    ''dr(2) = DateTime.Now.ToShortTimeString
    ''If (i Mod 2 <> 0) Then
    ''    dr(3) = True
    ''Else
    ''    dr(3) = False
    ''End If
    ''dr(4) = 1.23 * (i + 1)

    'add the row to the datatable
    dt.Rows.Add(dr)

    ''return a DataView to the DataTable
    ''CreateDataSource = New DataView(dt)

  End Function

  Private Sub cmdCodeGen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCodeGen.Click
    If txtName.Text = "" Then
      Say("Product Name is empty.")
      Exit Sub
    ElseIf txtVer.Text = "" Then
      Say("Product Version is empty.")
      Exit Sub
    ElseIf CheckDuplicate(txtName.Text, txtVer.Text) Then
      Say("Selected Product Name and Version already exist in the Product Listbox.")
      Exit Sub
    End If
    lblUpdateStatus.Text = ""
    txtCode1.Text = ""
    txtCode2.Text = ""
    On Error GoTo Done
    Dim Key As New RSAKey
    ReDim Key.data(32)
    'Adding a delegate for AddressOf CryptoProgressUpdate did not work
    'for now. Modified alcrypto3NET.dll to create a second generate function
    'rsa_generate2 that does not deal with progress monitoring - ialkan
    modALUGEN.rsa_generate2(Key, 1024)

    ' extract private and public key blobs
    Dim strBlob As String
    Dim blobLen As Integer
    rsa_public_key_blob(Key, vbNullString, blobLen)
    If blobLen > 0 Then
      strBlob = New String(Chr(0), blobLen)
      rsa_public_key_blob(Key, strBlob, blobLen)
      'System.Diagnostics.Debug.WriteLine("Public blob: " & strBlob)
      txtCode1.Text = strBlob
    End If

    'FUTURE RSA IMPLEMENTATION USING NATIVE VB.NET FUNCTIONS
    'Dim rsa As New RSACryptoServiceProvider
    'Dim xmlPublicKey As String
    'Dim xmlPrivateKey As String
    'xmlPublicKey = rsa.ToXmlString(False)
    'xmlPrivateKey = rsa.ToXmlString(True)
    'rsa.FromXmlString(xmlPublicKey)
    'txtCode1.Text = xmlPublicKey
    'txtCode2.Text = xmlPrivateKey

    modALUGEN.rsa_private_key_blob(Key, vbNullString, blobLen)
    If blobLen > 0 Then
      strBlob = New String(Chr(0), blobLen)
      modALUGEN.rsa_private_key_blob(Key, strBlob, blobLen)
      'System.Diagnostics.Debug.WriteLine("Private blob: " & strBlob)
      txtCode2.Text = strBlob
    End If

    ' done with the key - throw it away
    modALUGEN.rsa_freekey(Key)

    ' Test generated key for correctness by recreating it from the blobs
    ' Note:
    ' ====
    ' Due to an outstanding bug in ALCrypto3NET.dll, sometimes this step will crash the app because
    ' the generated keyset is bad.
    ' The work-around for the time being is to keep regenerating the keyset until eventually
    ' you'll get a valid keyset that no longer crashes.
    Dim strdata As String : strdata = "This is a test string to be encrypted."
    modALUGEN.rsa_createkey(txtCode1.Text, Len(txtCode1.Text), txtCode2.Text, Len(txtCode2.Text), Key)
    ' It worked! We're all set to go.
    modALUGEN.rsa_freekey(Key)
Done:

  End Sub

  Private Sub cmdValidate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdValidate.Click
    If txtCode1_2.Text = "" Then
      Say("Product VCode field is empty.")
      Exit Sub
    ElseIf txtVer.Text = "" Then
      Say("Product GCode field is empty.")
      Exit Sub
    ElseIf txtCode1_2.Text = "" And txtCode2_2.Text = "" Then
      UpdateStatus("Product VCode and GCode fields are blank. Nothing to validate.")
      Exit Sub ' nothing to validate
    End If
    ' Validate to keyset to make sure it's valid.
    UpdateStatus("Validating keyset...")
    Dim Key As RSAKey
    ReDim Key.data(32)
    Dim strdata As String
    strdata = "This is a test string to be signed."
    Dim strSig As String
    Dim txt1 As String, txt2 As String
    Dim len1 As Integer, len2 As Integer
    txt1 = txtCode1_2.Text
    len1 = Len(txtCode1_2.Text)
    txt2 = txtCode2_2.Text
    len2 = Len(txtCode2_2.Text)
    modALUGEN.rsa_createkey(txt1, len1, txt2, len2, Key)
    ' sign it
    strSig = RSASign(txtCode1_2.Text, txtCode2_2.Text, strdata)
    Dim rc As Integer
    rc = RSAVerify(txtCode1_2.Text, strdata, strSig)
    Dim dr As DataRow
    dr = dt.Rows(gridProds.SelectedIndex)
    If rc = 0 Then
      UpdateStatus(dr(0) & " (" & dr(1) & ") validated successfully.")
    Else
      UpdateStatus(dr(0) & " (" & dr(1) & ") GCode-VCode mismatch!")
    End If
    ' It worked! We're all set to go.
    modALUGEN.rsa_freekey(Key)

  End Sub

  Private Sub cmdRemove_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRemove.Click

    Dim dr As DataRow = dt.Rows(gridProds.SelectedItem.ItemIndex)
    dt.Rows.Remove(dr)
    Dim strName, strVer As String
    strName = gridProds.SelectedItem.Cells(1).Text
    strVer = gridProds.SelectedItem.Cells(2).Text

    GeneratorInstance.DeleteProduct(strName, strVer)
    gridProds.DataSource = dt
    gridProds.DataBind()

    gridProds.SelectedIndex = dt.Rows.Count - 1
    txtName.Text = gridProds.SelectedItem.Cells(1).Text
    txtVer.Text = gridProds.SelectedItem.Cells(2).Text

    txtCode1.Text = ""
    txtCode2.Text = ""
    cmdValidate.Enabled = True
  End Sub

  Private Sub UpdateStatus(ByRef Msg As String)
    lblUpdateStatus.Text = Msg
  End Sub

  Public Function AppPath() As String
    Dim siteNameUrl As String
    AppPath = System.IO.Path.GetDirectoryName(Server.MapPath("ASPNETAlugen3_2.aspx"))
  End Function

  Private Sub PopulateUI(ByVal ProdInfo As ProductInfo)
    With ProdInfo
      AddRow(.name, .Version, .VCode, .GCode, False)
    End With
  End Sub
  Private Sub DataGrid1_ItemCreated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles gridProds.ItemCreated
    If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Or e.Item.ItemType = ListItemType.SelectedItem Then
      e.Item.Attributes.Add("onmouseover", "this.style.backgroundColor='beige';this.style.cursor='hand'")
      e.Item.Attributes.Add("onmouseout", "this.style.backgroundColor='white';this.style.color='black'")
      e.Item.Attributes.Add("onclick", "javascript:__doPostBack('" & "gridProds:" & "_ctl" & (e.Item.ItemIndex + 2) & ":_ctl0','')")
    End If
  End Sub

  Private Sub DataGrid1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridProds.SelectedIndexChanged
    lblUpdateStatus.Text = ""
    gridProds.SelectedItem.Attributes.Item("onmouseover") = "this.style.cursor='hand'"
    gridProds.SelectedItem.Attributes.Remove("onmouseout")
    txtName.Text = gridProds.SelectedItem.Cells(1).Text
    txtVer.Text = gridProds.SelectedItem.Cells(2).Text
    txtCode1_2.Text = gridProds.SelectedItem.Cells(3).Text
    txtCode2_2.Text = gridProds.SelectedItem.Cells(4).Text
  End Sub

  Private Sub txtName_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtName.TextChanged

  End Sub

  Private Sub txtVer_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtVer.TextChanged

  End Sub
  Private Function UpdateButtonStatus() As Boolean
    If txtName.Text = "" Or txtVer.Text = "" Then
      Say("No Product Name or Version Number has been specified.")
      UpdateButtonStatus = True
    ElseIf CheckDuplicate(txtName.Text, txtVer.Text) Then
      Say("Selected Product Name and Version already exist in the Product Listbox.")
      UpdateButtonStatus = True
    End If
  End Function

  ' Validate and enable Add/Change button as appropriate
  Private Function CheckDuplicate(ByRef Name As String, ByRef Ver As String) As Boolean
    CheckDuplicate = False
    Dim dr As DataRow
    dr = dt.Rows(gridProds.SelectedIndex)
    Dim i As Short
    For i = 0 To dt.Rows.Count - 1
      If dr(0) = Name Then
        If dr(1) = Ver Then
          CheckDuplicate = True
          Exit Function
        End If
      End If
    Next
  End Function

  Private Sub Say(ByVal Message As String)
    ' Format string properly
    Message = Message.Replace("'", "\'")
    Message = Message.Replace(Convert.ToChar(10), "\n")
    Message = Message.Replace(Convert.ToChar(13), "")
    ' Display as JavaScript alert
    ltlAlert.Text = "alert('" & Message & "')"
  End Sub

End Class
