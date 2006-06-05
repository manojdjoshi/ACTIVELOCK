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
Imports System
Imports System.Xml
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports FriendSoftware.DNN.Modules.Alugen3NET.Business
Imports DotNetNuke
Imports DotNetNuke.Entities.Modules
Imports DotNetNuke.UI.Utilities
Imports ActiveLock3NET

Namespace FriendSoftware.DNN.Modules.Alugen3NET

  Public MustInherit Class AlugenProductsEdit
    Inherits Entities.Modules.PortalModuleBase
    Implements IClientAPICallbackEventHandler


#Region "Controls"
    Protected WithEvents cmdUpdate As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdCancel As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdDelete As System.Web.UI.WebControls.LinkButton
    Protected WithEvents txtProductName As System.Web.UI.WebControls.TextBox
    Protected WithEvents plProductName As DotNetNuke.UI.UserControls.LabelControl
    Protected WithEvents plProductNameVersion As DotNetNuke.UI.UserControls.LabelControl
    Protected WithEvents plProductVCode As DotNetNuke.UI.UserControls.LabelControl
    Protected WithEvents plProductGCode As DotNetNuke.UI.UserControls.LabelControl
    Protected WithEvents plProductPrice As DotNetNuke.UI.UserControls.LabelControl
    Protected WithEvents txtProductVersion As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtProductPrice As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnGenerateNew As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents txtProductVCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtProductGCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnClearCodes As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents btnValidate As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents imgCopyVCode As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents imgCopyGCode As System.Web.UI.HtmlControls.HtmlImage
#End Region

#Region "Private Members"
    Private productID As Integer
    Protected WithEvents pnlCredits As System.Web.UI.WebControls.Panel
    Protected WithEvents imgFriendSoftware As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents imgActiveLock As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents chkDirectCodeEdit As System.Web.UI.WebControls.CheckBox
    Private strMessage As String
#End Region

#Region "Event Handlers"
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Try
        Dim objModules As New ModuleController
        Dim mysettings As New Hashtable
        mysettings = objModules.GetModuleSettings(ModuleId)

        Dim objAlugen3NETController As New AlugenProductsController
        Dim objAlugen3NETInfo As New AlugenProductsInfo

        ' Determine ProductID
        If Not (Request.Params("ProductID") Is Nothing) Then
          productID = Int32.Parse(Request.Params("ProductID"))
        Else
          productID = Null.NullInteger()
        End If

        'hide Credits
        pnlCredits.Visible = Not CType(mysettings("HideCredits"), Boolean)
        If pnlCredits.Visible Then
          Me.imgFriendSoftware.Attributes.Add("src", ModulePath + "friendsoftware_logo.JPG")
          Me.imgActiveLock.Attributes.Add("src", ModulePath + "activelock_logo.gif")
        End If

        'inject JS call into body load
        Dim mbody As HtmlGenericControl = CType(Page.FindControl("body"), HtmlGenericControl)
        mbody.Attributes("onload") += ";inimodulePath('" & Me.ModulePath & "');"

        If Request.Browser.Browser = "Netscape" Then
          'copy command not working on FireFox, Netscape
          Me.imgCopyVCode.Attributes.Add("src", ModulePath + "copy_disabled.gif")
          Me.imgCopyGCode.Attributes.Add("src", ModulePath + "copy_disabled.gif")
        Else
          Me.imgCopyVCode.Attributes.Add("src", ModulePath + "copy.gif")
          Me.imgCopyGCode.Attributes.Add("src", ModulePath + "copy.gif")
        End If

        Me.btnClearCodes.Attributes.Add("onclick", "ClearCodes('" & Me.ClientID & "');createPopup('" & Me.btnClearCodes.ClientID & "', 'Clearing confirmation', 'Clearing was done succesfully!<br><br><img border=0 src=" & Me.ModulePath & "button_ok.gif style=\'cursor:pointer;cursor:hand;\' onclick=\'hidebox();return false\'>', 250, 200);")
        Me.imgCopyVCode.Attributes.Add("onclick", "CopyToClipboard('" & Me.txtProductVCode.ClientID & "')")
        Me.imgCopyGCode.Attributes.Add("onclick", "CopyToClipboard('" & Me.txtProductGCode.ClientID & "')")
        Me.chkDirectCodeEdit.Attributes.Add("onclick", "ChangeCodeEditProduct('" & Me.ClientID & "');")

        'these won't be necessary in next release after 3.2.0
        If ClientAPI.BrowserSupportsFunctionality(ClientAPI.ClientFunctionality.XMLHTTP) _
          AndAlso ClientAPI.BrowserSupportsFunctionality(ClientAPI.ClientFunctionality.XML) Then
          ClientAPI.RegisterClientReference(Me.Page, ClientAPI.ClientNamespaceReferences.dnn_xml)
          ClientAPI.RegisterClientReference(Me.Page, ClientAPI.ClientNamespaceReferences.dnn_xmlhttp)

          'Only this line will be necessary after 3.2
          Me.btnGenerateNew.Attributes.Add("onclick", "DisableAllControlsProduct('" & Me.ClientID & "');this.style.cursor = 'url(" & ModulePath & "dnnwork_arrow.ani),wait;';document.body.style.cursor='url(" & ModulePath & "dnnwork_arrow.ani),wait;';" & ClientAPI.GetCallbackEventReference(Me, "'generate'" & "+'|'+dnn.dom.getById('" & Me.txtProductName.ClientID & "').value+'|'+dnn.dom.getById('" & Me.txtProductVersion.ClientID & "').value", "successFuncGenerate", "'" & Me.ClientID & "'", "errorFunc") & "; dnn.dom.getById('" & Me.txtProductVCode.ClientID & "').value='..generating codes - please wait..'; dnn.dom.getById('" & Me.txtProductGCode.ClientID & "').value='..generating codes - please wait..';")
          Me.btnValidate.Attributes.Add("onclick", "DisableAllControlsProduct('" & Me.ClientID & "');this.style.cursor='url(" & ModulePath & "dnnwork_arrow.ani),wait;';document.body.style.cursor='url(" & ModulePath & "dnnwork_arrow.ani),wait;';" & ClientAPI.GetCallbackEventReference(Me, "'validate'" + "+'|' + encodeURIComponent(dnn.dom.getById('" & Me.txtProductVCode.ClientID & "').value)+'|'+encodeURIComponent(dnn.dom.getById('" & Me.txtProductGCode.ClientID & "').value)", "successFuncValidate", "'" & Me.ClientID & "'", "errorFunc"))
        End If

        'setting the js file
        If Me.Page.IsClientScriptBlockRegistered("alugen3netMain.js") = False Then
          Me.Page.RegisterClientScriptBlock("alugen3netMain.js", "<script language=javascript src=""" & Me.ModulePath & "alugen3netMain.js""></script>")
        End If
        If Me.Page.IsClientScriptBlockRegistered("popup.js") = False Then
          Me.Page.RegisterClientScriptBlock("popup.js", "<script language=javascript src=""" & Me.ModulePath & "popup.js""></script>")
        End If

        If Not Page.IsPostBack Then
          cmdDelete.Attributes.Add("onClick", "javascript:return confirm('" & Localization.GetString("DeleteItem.Text", LocalResourceFile) & "');")
          If Not Null.IsNull(productID) Then
            objAlugen3NETInfo = objAlugen3NETController.Get(productID)
            If Not objAlugen3NETInfo Is Nothing Then
              'Load data
              txtProductName.Text = objAlugen3NETInfo.ProductName
              txtProductVersion.Text = objAlugen3NETInfo.ProductVersion
              txtProductVCode.Text = objAlugen3NETInfo.ProductVCode
              txtProductGCode.Text = objAlugen3NETInfo.ProductGCode
              txtProductPrice.Text = objAlugen3NETInfo.ProductPrice.ToString()

              'setting controls
              If txtProductVCode.Text.Length > 0 Or txtProductVCode.Text.Length > 0 Then
                btnClearCodes.Disabled = False
                btnGenerateNew.Disabled = True
              Else
                btnClearCodes.Disabled = True
                btnGenerateNew.Disabled = False
              End If
              cmdDelete.Visible = True
            Else ' security violation attempt to access item not related to this Module
              Response.Redirect(NavigateURL(), True)
            End If

          End If
        End If
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdUpdate.Click
      Try
        ' Only Update if the Entered Data is Valid
        If Page.IsValid = True Then
          If txtProductName.Text.Trim.Length > 0 Then
            Dim objAlugen3NETInfo As New AlugenProductsInfo
            Dim strAction As String

            objAlugen3NETInfo = CType(CBO.InitializeObject(objAlugen3NETInfo, GetType(AlugenProductsInfo)), AlugenProductsInfo)

            'bind values to object
            objAlugen3NETInfo.ProductID = productID
            objAlugen3NETInfo.ModuleID = ModuleId
            objAlugen3NETInfo.PortalID = PortalId
            objAlugen3NETInfo.ProductGCode = txtProductGCode.Text
            objAlugen3NETInfo.ProductVCode = txtProductVCode.Text
            objAlugen3NETInfo.ProductName = txtProductName.Text
            objAlugen3NETInfo.ProductVersion = txtProductVersion.Text
            Try
              objAlugen3NETInfo.ProductPrice = Math.Max(0, CType(txtProductPrice.Text, Decimal))
            Catch
              objAlugen3NETInfo.ProductPrice = 0
            End Try
            objAlugen3NETInfo.UserName = UserInfo.Username

            'updating data
            Dim objCtlAlugen3NET As New AlugenProductsController
            If Null.IsNull(productID) Then
              strAction = "added"
              objCtlAlugen3NET.Add(objAlugen3NETInfo)
            Else
              strAction = "updated"
              objCtlAlugen3NET.Update(objAlugen3NETInfo)
            End If

            ' Redirect back to the portal home page
            strMessage = "Product " & strAction & " succesfully!"
            Response.Redirect(NavigateURL(TabId, "", "lastPage=p&Message=" & XmlConvert.EncodeName(strMessage) & "&MsgModule=" & ModuleId.ToString()), True)
          End If
        Else
          'name is null (zero lenght)
        End If
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdCancel.Click
      Try
        Response.Redirect(NavigateURL(TabId, "", "lastPage=p"), True)
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdDelete.Click
      Dim mystrMessage As String
      Try
        If Not Null.IsNull(productID) Then
          Dim objAlugen3NETController As New AlugenProductsController
          Dim objAlugenProductsInfo As New AlugenProductsInfo
          objAlugenProductsInfo = objAlugen3NETController.Get(productID)

          'verify if product is already used
          Dim objAlugenCodeController As New AlugenCodeController
          Dim myListByProduct As ArrayList = objAlugenCodeController.ListByProductID(PortalId, ModuleId, productID)

          If myListByProduct.Count > 0 Then
            mystrMessage = "Product is in use!"
            Throw New System.Exception("Product " & objAlugenProductsInfo.ProductName & " - " & objAlugenProductsInfo.ProductVersion & " cannot be deleted beacause it have generated codes! Plese delete those codes first")
          End If

          objAlugen3NETController.Delete(productID)
        End If

      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
        strMessage = "Product deletion fail!"
        strMessage = strMessage & "<br>" & mystrMessage
        Response.Redirect(NavigateURL(TabId, "", "lastPage=p&Message=" & XmlConvert.EncodeName(strMessage) & "&IsError=true" & "&MsgModule=" & ModuleId.ToString()), True)
        Exit Sub
      End Try

      ' Redirect back to the portal home page
      strMessage = "Product deleted succesfully!"
      Response.Redirect(NavigateURL(TabId, "", "lastPage=p&Message=" & XmlConvert.EncodeName(strMessage) & "&MsgModule=" & ModuleId.ToString()), True)

    End Sub
#End Region

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

    Public Function RaiseClientAPICallbackEvent(ByVal eventArgument As String) As String Implements DotNetNuke.UI.Utilities.IClientAPICallbackEventHandler.RaiseClientAPICallbackEvent
      Try
        Dim msg As String = String.Empty
        Dim m_txtProductVCode As String = String.Empty
        Dim m_txtProductGCode As String = String.Empty

        Dim strSplit As String()
        strSplit = Split(eventArgument, "|")
        If strSplit(0) = "generate" Then
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
            m_txtProductVCode = strBlob
          End If

          rsa_private_key_blob(Key, vbNullString, blobLen)
          If blobLen > 0 Then
            strBlob = New String(Chr(0), blobLen)
            modALUGEN.rsa_private_key_blob(Key, strBlob, blobLen)
            m_txtProductGCode = strBlob
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
          modALUGEN.rsa_createkey(m_txtProductVCode, Len(m_txtProductVCode), m_txtProductGCode, Len(m_txtProductGCode), Key)
          ' It worked! We're all set to go.
          modALUGEN.rsa_freekey(Key)
          msg = strSplit(1) & "|" & strSplit(2) & "|" & m_txtProductVCode & "|" & m_txtProductGCode
        ElseIf strSplit(0) = "validate" Then
          If strSplit(1).Length > 0 And strSplit(2).Length > 0 Then
            msg = "false"
            Try
              Dim Key As RSAKey
              ReDim Key.data(32)
              Dim strdata As String
              strdata = "This is a test string to be signed."
              Dim strSig As String
              Dim txt1 As String, txt2 As String
              Dim len1 As Integer, len2 As Integer
              txt1 = strSplit(1)
              len1 = Len(txt1)
              txt2 = strSplit(2)
              len2 = Len(txt2)
              rsa_createkey(txt1, len1, txt2, len2, Key)
              ' sign it
              strSig = RSASign(txt1, txt2, strdata)
              Dim rc As Integer
              rc = RSAVerify(txt1, strdata, strSig)
              ' It worked! We're all set to go.
              rsa_freekey(Key)

              If rc = 0 Then
                'validated successfull
                msg = "true"
              Else
                'GCode-VCode mismatch!
                msg = "false"
              End If
            Catch
            Finally
            End Try
          Else
            msg = "false"
          End If
        Else
          'other non-senses cases
          msg = ""
        End If

        Return msg
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Function
  End Class

End Namespace
