'
' FRIEND SOFTWARE SRC - http://www.friendsoftware.ro
' Copyright (©) 2006
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated 
' documentation files (the "Software"), to deal in the Software without restriction, including without limitation 
' the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and 
' to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or substantial portions 
' of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED 
' TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL 
' THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF 
' CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
' DEALINGS IN THE SOFTWARE.
'
Imports System.Xml
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports FriendSoftware.DNN.Modules.Alugen3NET.Business
Imports DotNetNuke
Imports DotNetNuke.UI.Utilities
Imports DotNetNuke.Entities.Modules
Imports ActiveLock3NET
Imports Nini.Ini


Namespace FriendSoftware.DNN.Modules.Alugen3NET

  Public MustInherit Class AlugenMain
    Inherits Entities.Modules.PortalModuleBase
    Implements Entities.Modules.IActionable
    Implements Entities.Modules.IPortable
    Implements Entities.Modules.ISearchable
    Implements IClientAPICallbackEventHandler


#Region "Controls"
    Protected WithEvents pnlCredits As System.Web.UI.WebControls.Panel
    Protected WithEvents pnlMessage As System.Web.UI.WebControls.Panel
    Protected WithEvents lblMessage As System.Web.UI.WebControls.Label
    Protected WithEvents imgFriendSoftware As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents imgActiveLock As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents iniProductFile As System.Web.UI.HtmlControls.HtmlInputFile
    Protected WithEvents cmdImportProducts As System.Web.UI.WebControls.LinkButton
    Protected WithEvents txtImport As System.Web.UI.WebControls.TextBox
    Protected WithEvents lblImportProducts As System.Web.UI.WebControls.Label
    Protected WithEvents productstable As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents customerstable As System.Web.UI.HtmlControls.HtmlTable
    Protected WithEvents codestable As System.Web.UI.HtmlControls.HtmlTable
#End Region

#Region "Private Members"
    Private lastPage As String
#End Region

#Region "Event Handlers"
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Try
        Dim objModules As New ModuleController
        Dim mysettings As New Hashtable
        mysettings = objModules.GetModuleSettings(ModuleId)

        Dim objAlugen3NETController As New AlugenProductsController
        Dim objAlugen3NETInfo As New AlugenProductsInfo
        Dim objAlugenCustomersController As New AlugenCustomersController
        Dim objAlugenCustomersInfo As New AlugenCustomersInfo
        Dim objAlugenCodeController As New AlugenCodeController
        Dim objAlugenCodeInfo As New AlugenCodeInfo

        ' Determine lastpageID
        If Not (Request.Params("lastPage") Is Nothing) Then
          lastPage = Request.Params("lastPage").ToString()
          If txtLastPage.Value.Length = 0 Then
            txtLastPage.Value = lastPage
          End If
        Else
          lastPage = String.Empty
        End If

        'hide Credits
        pnlCredits.Visible = Not CType(mysettings("HideCredits"), Boolean)
        If pnlCredits.Visible Then
          Me.imgFriendSoftware.Attributes.Add("src", ModulePath + "friendsoftware_logo.JPG")
          Me.imgActiveLock.Attributes.Add("src", ModulePath + "activelock_logo.gif")
        End If

        ' Determine Message and IsError
        pnlMessage.Visible = False
        If Not Request.Params("Message") Is Nothing Then
          If Not Request.Params("MsgModule") Is Nothing AndAlso Request.Params("MsgModule") = ModuleId.ToString Then
            pnlMessage.Visible = True
            lblMessage.Text = XmlConvert.DecodeName(Request.Params("Message"))

            lblMessage.CssClass = "messagevalid"
            If Not Request.Params("IsError") Is Nothing Then
              If (CType(Request.Params("IsError"), Boolean)) Then
                lblMessage.CssClass = "messageerror"
              End If
            End If
          End If
        End If

        Me.imgRefresh.Attributes.Add("src", ModulePath + "refresh.gif")
        Me.chkFilterProducts.Attributes.Add("onclick", "enableFilter('" & Me.chkFilterProducts.ClientID & "', '" & Me.pnlProducts.ClientID & "');")
        Me.chkFilterCustomers.Attributes.Add("onclick", "enableFilter('" & Me.chkFilterCustomers.ClientID & "', '" & Me.pnlCustomers.ClientID & "');")
        Me.chkFilterCodes.Attributes.Add("onclick", "enableFilter('" & Me.chkFilterCodes.ClientID & "', '" & Me.pnlCodes.ClientID & "');")

        'these won't be necessary in next release after 3.2.0
        If ClientAPI.BrowserSupportsFunctionality(ClientAPI.ClientFunctionality.XMLHTTP) _
          AndAlso ClientAPI.BrowserSupportsFunctionality(ClientAPI.ClientFunctionality.XML) Then
          ClientAPI.RegisterClientReference(Me.Page, ClientAPI.ClientNamespaceReferences.dnn_xml)
          ClientAPI.RegisterClientReference(Me.Page, ClientAPI.ClientNamespaceReferences.dnn_xmlhttp)

          'Only this line will be necessary after 3.2
          'Me.lnkRefresh.Attributes.Add("onclick", "this.style.cursor='" & ModulePath & "dnnwork_arrow.ani';document.body.style.cursor='" & ModulePath & "dnnwork_arrow.ani';" & ClientAPI.GetCallbackEventReference(Me, "'refresh' + '|' + dnn.dom.getById('" & Me.pnlProducts.ClientID & "').style.display + '|' + dnn.dom.getById('" & Me.pnlCustomers.ClientID & "').style.display + '|' + dnn.dom.getById('" & Me.pnlCodes.ClientID & "').style.display", "successFuncRefresh", "'" & Me.ClientID & "'", "errorFunc"))
        End If

        'setting default state of filters
        Me.pnlProducts.Attributes("mytag") = "0"
        Me.pnlCustomers.Attributes("mytag") = "0"
        Me.pnlCodes.Attributes("mytag") = "0"
        If CType(mysettings("EnableFiltersByDefault"), Boolean) Then
          Me.pnlProducts.Attributes("mytag") = "1"
          Me.pnlCustomers.Attributes("mytag") = "1"
          Me.pnlCodes.Attributes("mytag") = "1"

          Me.chkFilterProducts.Checked = True
          Me.chkFilterCustomers.Checked = True
          Me.chkFilterCodes.Checked = True
        End If

        'inject JS call into body load
        Dim mbody As HtmlGenericControl = CType(Page.FindControl("body"), HtmlGenericControl)
        'only one filter init
        If Not mbody.Attributes("onload") Is Nothing Then
          If Not mbody.Attributes("onload").IndexOf("filters_init()") > 0 Then
            mbody.Attributes("onload") += ";filters_init();"
          End If
        Else
          mbody.Attributes.Add("onload", ";filters_init();")
        End If

        If Me.Page.IsClientScriptBlockRegistered("alugen3netMain.js") = False Then
          Me.Page.RegisterClientScriptBlock("alugen3netMain.js", "<script language=javascript src=""" & Me.ModulePath & "alugen3netMain.js""></script>")
        End If
        If Me.Page.IsClientScriptBlockRegistered("sorttable.js") = False Then
          Me.Page.RegisterClientScriptBlock("sorttable.js", "<script language=javascript src=""" & Me.ModulePath & "sorttable.js""></script>")
        End If
        If Me.Page.IsClientScriptBlockRegistered("tablefilter.js") = False Then
          Me.Page.RegisterClientScriptBlock("tablefilter.js", "<script language=javascript src=""" & Me.ModulePath & "tablefilter.js""></script>")
        End If

        Select Case lastPage
          Case "c" 'customers page
            pnlProducts.Style("display") = "none"
            tdbtnProducts.Attributes.Add("class", "multipage")

            pnlCustomers.Style("display") = "block"
            tdbtnCustomers.Attributes.Add("class", "multipageactive")

            pnlCodes.Style("display") = "none"
            tdbtnCodes.Attributes.Add("class", "multipage")
          Case "d" 'codes pages
            pnlProducts.Style("display") = "none"
            tdbtnProducts.Attributes.Add("class", "multipage")

            pnlCustomers.Style("display") = "none"
            tdbtnCustomers.Attributes.Add("class", "multipage")

            pnlCodes.Style("display") = "block"
            tdbtnCodes.Attributes.Add("class", "multipageactive")
          Case Else 'and "p" 'product pages
            pnlProducts.Style("display") = "block"
            tdbtnProducts.Attributes.Add("class", "multipageactive")

            pnlCustomers.Style("display") = "none"
            tdbtnCustomers.Attributes.Add("class", "multipage")

            pnlCodes.Style("display") = "none"
            tdbtnCodes.Attributes.Add("class", "multipage")
        End Select

        If Not Page.IsPostBack Then
          'adding events to client elements
          tdbtnProducts.Attributes.Add("onClick", "ShowProducts('" & Me.ClientID & "');")
          tdbtnCustomers.Attributes.Add("onClick", "ShowCustomers('" & Me.ClientID & "');")
          tdbtnCodes.Attributes.Add("onClick", "ShowCodes('" & Me.ClientID & "');")

          'filling repeaters
          rptAlugenProducts.DataSource = objAlugen3NETController.List(PortalId, ModuleId)
          rptAlugenProducts.DataBind()
          rptAlugenCustomers.DataSource = objAlugenCustomersController.List(PortalId, ModuleId)
          rptAlugenCustomers.DataBind()
          rptAlugenCodes.DataSource = objAlugenCodeController.ListFull(PortalId, ModuleId)
          rptAlugenCodes.DataBind()

        End If

      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub
#End Region

#Region "Optional Interfaces"
    Public ReadOnly Property ModuleActions() As Entities.Modules.Actions.ModuleActionCollection Implements Entities.Modules.IActionable.ModuleActions
      Get
        Dim Actions As New Entities.Modules.Actions.ModuleActionCollection
        'Actions.Add(GetNextActionID, Localization.GetString(Entities.Modules.Actions.ModuleActionType.AddContent, LocalResourceFile), Entities.Modules.Actions.ModuleActionType.AddContent, "", "", EditUrl(), False, DotNetNuke.Security.SecurityAccessLevel.Edit, True, False)

        Actions.Add(GetNextActionID, Localization.GetString("AddCode.Action", LocalResourceFile), _
          Entities.Modules.Actions.ModuleActionType.AddContent, "", "", EditUrl("EditCode"), False, Security.SecurityAccessLevel.Edit, True, False)

        Actions.Add(GetNextActionID, Localization.GetString("AddCustomer.Action", LocalResourceFile), _
          Entities.Modules.Actions.ModuleActionType.AddContent, "", "", EditUrl("EditCustomer"), False, Security.SecurityAccessLevel.Edit, True, False)

        Actions.Add(GetNextActionID, Localization.GetString("AddProduct.Action", LocalResourceFile), _
          Entities.Modules.Actions.ModuleActionType.AddContent, "", "", EditUrl("EditProduct"), False, Security.SecurityAccessLevel.Edit, True, False)

        Return Actions
      End Get
    End Property

    Public Function ExportModule(ByVal ModuleID As Integer) As String Implements Entities.Modules.IPortable.ExportModule
      ' included as a stub only so that the core knows this module Implements Entities.Modules.IPortable
    End Function

    Public Sub ImportModule(ByVal ModuleID As Integer, ByVal Content As String, ByVal Version As String, ByVal UserID As Integer) Implements Entities.Modules.IPortable.ImportModule
      ' included as a stub only so that the core knows this module Implements Entities.Modules.IPortable
    End Sub

    Public Function GetSearchItems(ByVal ModInfo As Entities.Modules.ModuleInfo) As Services.Search.SearchItemInfoCollection Implements Entities.Modules.ISearchable.GetSearchItems
      ' included as a stub only so that the core knows this module Implements Entities.Modules.ISearchable
    End Function

#End Region

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents lstFAQs As System.Web.UI.WebControls.DataList
    Protected WithEvents pnlProducts As System.Web.UI.WebControls.Panel
    Protected WithEvents pnlCustomers As System.Web.UI.WebControls.Panel
    Protected WithEvents pnlCodes As System.Web.UI.WebControls.Panel
    Protected WithEvents rptAlugen3NEtProducts As System.Web.UI.WebControls.Repeater
    Protected WithEvents txtProductName As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtProductVersion As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtProductVCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtProductGCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtProductPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnGenerateNew As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents btnClearCodes As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents btnValidate As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents btnCopyVCode As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents btnCopyGCode As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents rptAlugen3NETCustomers As System.Web.UI.WebControls.Repeater
    Protected WithEvents txtCustomerName As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCustomerAddress As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCustomerContactPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCustomerPhone As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCustomerEmail As System.Web.UI.WebControls.TextBox
    Protected WithEvents cmdUpdateProduct As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdCancelProduct As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdDeleteProduct As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdUpdateCustomer As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdCancelCustomer As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdDeleteCustomer As System.Web.UI.WebControls.LinkButton
    Protected WithEvents rptAlugen3NETCodes As System.Web.UI.WebControls.Repeater
    Protected WithEvents imgRefresh As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents lnkRefresh As System.Web.UI.WebControls.LinkButton
    Protected WithEvents txtLastPage As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents tdbtnCodes As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents tdbtnCustomers As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents tdbtnProducts As System.Web.UI.HtmlControls.HtmlTableCell
    Protected WithEvents lblProducts As System.Web.UI.WebControls.Label
    Protected WithEvents chkFilterProducts As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents rptAlugenProducts As System.Web.UI.WebControls.Repeater
    Protected WithEvents rptAlugenCustomers As System.Web.UI.WebControls.Repeater
    Protected WithEvents rptAlugenCodes As System.Web.UI.WebControls.Repeater
    Protected WithEvents chkFilterCustomers As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents chkFilterCodes As System.Web.UI.HtmlControls.HtmlInputCheckBox

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
      'CODEGEN: This method call is required by the Web Form Designer
      'Do not modify it using the code editor.
      InitializeComponent()
    End Sub

#End Region


    Public Function MyDecFormat() As String
      'decimal price format
      Dim decimals As Short = 0
      Dim PriceDecimals As String = CType(PortalSettings.GetModuleSettings(ModuleId)("PriceDecimals"), String)
      If Not Null.IsNull(PriceDecimals) Then
        decimals = CShort(PriceDecimals)
      End If
      Return "{0:N" & decimals.ToString() & "}"
    End Function


    Public Function GetLicenseType(ByVal mValue As String) As String
      'return license type
      Dim mReturn As String
      Select Case Int32.Parse(mValue)
        Case 0
          mReturn = "Time locked"
        Case 1
          mReturn = "Periodic"
        Case 2
          mReturn = "Permanent"
        Case Else
          mReturn = "(undefined)"
      End Select
      Return mReturn
    End Function

    Public Function RaiseClientAPICallbackEvent(ByVal eventArgument As String) As String Implements DotNetNuke.UI.Utilities.IClientAPICallbackEventHandler.RaiseClientAPICallbackEvent
      Try
        Dim msg As String = String.Empty

        Dim strSplit As String()
        strSplit = Split(eventArgument, "|")
        If strSplit(0) = "refresh" Then
        Else
          msg = ""
        End If
        Return msg
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Function

    Private Sub lnkRefresh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lnkRefresh.Click
      'redirecting
      Response.Redirect(NavigateURL(TabId, "", "lastPage=" & txtLastPage.Value), True)
    End Sub

    Private Sub cmdImportProducts_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdImportProducts.Click
      Dim strMessage As String
      Try

        Dim impFile As HttpPostedFile = iniProductFile.PostedFile
        If Not Null.IsNull(impFile) AndAlso iniProductFile.PostedFile.FileName.Trim.Length <> 0 Then
          'import
          Dim objIniDoc As New IniDocument
          impFile.InputStream.Position = 0
          objIniDoc.Load(impFile.InputStream)
          Dim objSection As IniSection
          Dim objAlugenProductsController As New AlugenProductsController
          Dim x As Integer
          For x = 0 To objIniDoc.Sections.Count - 1
            objSection = objIniDoc.Sections(x)
            'For Each objSection In objIniDoc.Sections
            Dim objAlugenProductsInfo As New AlugenProductsInfo
            objAlugenProductsInfo.PortalID = PortalId
            objAlugenProductsInfo.ModuleID = ModuleId
            'real info
            objAlugenProductsInfo.ProductName = objSection.GetValue("Name")
            objAlugenProductsInfo.ProductVersion = objSection.GetValue("Version")
            objAlugenProductsInfo.ProductVCode = objSection.GetValue("VCode")
            objAlugenProductsInfo.ProductGCode = objSection.GetValue("GCode")
            objAlugenProductsInfo.ProductPrice = 0
            objAlugenProductsController.Add(objAlugenProductsInfo)
          Next
        Else
          Throw New System.Exception("File not valid!")
        End If
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
        'succes
        strMessage = "Products import has fail!<br>" & exc.Message
        Response.Redirect(NavigateURL(TabId, "", "lastPage=" & txtLastPage.Value & "&Message=" & XmlConvert.EncodeName(strMessage) & "&IsError=true&MsgModule=" & ModuleId.ToString()), True)
        Exit Sub
      End Try

      'succes
      strMessage = "Products imported succesfully!"
      Response.Redirect(NavigateURL(TabId, "", "lastPage=" & txtLastPage.Value & "&Message=" & XmlConvert.EncodeName(strMessage) & "&MsgModule=" & ModuleId.ToString()), True)

    End Sub
  End Class
End Namespace

