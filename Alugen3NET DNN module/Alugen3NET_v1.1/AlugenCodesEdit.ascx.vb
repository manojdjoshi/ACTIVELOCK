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
'Imports System
'Imports System.Web
Imports System.Xml
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports FriendSoftware.DNN.Modules.Alugen3NET.Business
Imports DotNetNuke
Imports DotNetNuke.Entities.Modules
Imports DotNetNuke.UI.Utilities
Imports ActiveLock3NET
Imports VB = Microsoft.VisualBasic
Imports System.Uri
Imports DotNetNuke.Services.Mail

Namespace FriendSoftware.DNN.Modules.Alugen3NET

  Public MustInherit Class AlugenCodesEdit
    Inherits Entities.Modules.PortalModuleBase
    Implements IClientAPICallbackEventHandler

#Region "Controls"
    Protected WithEvents cmdUpdate As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdCancel As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdDelete As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cboProduct As System.Web.UI.WebControls.DropDownList
    Protected WithEvents cboCustomer As System.Web.UI.WebControls.DropDownList
    Protected WithEvents cboLicenseType As System.Web.UI.WebControls.DropDownList
    Protected WithEvents txtExpireDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtInstalationCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtUserName As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtActivationCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnGenerate As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents cboRegisteredLevel As System.Web.UI.WebControls.DropDownList
    Protected WithEvents imgCopyActivationCode As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents imgPasteInstallCode As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents imgExpireDateCalendar As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents plExpireDate As DotNetNuke.UI.UserControls.LabelControl
    Protected WithEvents plPeriodicDays As DotNetNuke.UI.UserControls.LabelControl
    Protected WithEvents pnlExpireDate As System.Web.UI.WebControls.Panel
    Protected WithEvents pnlPeriodicDays As System.Web.UI.WebControls.Panel
    Protected WithEvents imgEmail As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents chkSendEmailNotification As System.Web.UI.WebControls.CheckBox
    Protected WithEvents pnlCredits As System.Web.UI.WebControls.Panel
    Protected WithEvents imgFriendSoftware As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents imgActiveLock As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents chkDirectCodeEdit As System.Web.UI.WebControls.CheckBox
    Protected WithEvents cmdPasteCode As System.Web.UI.HtmlControls.HtmlAnchor
    Protected WithEvents chkRegisterByCustomer As System.Web.UI.WebControls.CheckBox
    Protected WithEvents pnlGenByCustomer As System.Web.UI.WebControls.Panel
#End Region

#Region "Private Members"
    Private codeId As Integer
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
        Dim objAlugenCustomersController As New AlugenCustomersController
        Dim objAlugenCustomersInfo As New AlugenCustomersInfo
        Dim objAlugenCodeController As New AlugenCodeController
        Dim objAlugenCodeInfo As New AlugenCodeInfo

        ' Determine ItemId
        If Not (Request.Params("CodeId") Is Nothing) Then
          codeId = Int32.Parse(Request.Params("CodeId"))
        Else
          codeId = Null.NullInteger()
        End If

        'hide Credits
        pnlCredits.Visible = Not CType(mysettings("HideCredits"), Boolean)
        If pnlCredits.Visible Then
          Me.imgFriendSoftware.Attributes.Add("src", ModulePath + "friendsoftware_logo.JPG")
          Me.imgActiveLock.Attributes.Add("src", ModulePath + "activelock_logo.gif")
        End If

        'default registration type
        pnlExpireDate.Style("display") = "block"
        pnlPeriodicDays.Style("display") = "none"
        txtExpireDate.Style("display") = "block"
        imgExpireDateCalendar.Style("display") = "block"

        'inject JS call into body load
        Dim mbody As HtmlGenericControl = CType(Page.FindControl("body"), HtmlGenericControl)
        mbody.Attributes("onload") += ";inimodulePath('" & Me.ModulePath & "');"

        If Request.Browser.Browser = "Netscape" Then
          'copy command not working on FireFox, Netscape
          Me.imgCopyActivationCode.Attributes.Add("src", ModulePath + "copy_disabled.gif")
          Me.imgPasteInstallCode.Attributes.Add("src", ModulePath + "paste_disabled.gif")
        Else
          Me.imgCopyActivationCode.Attributes.Add("src", ModulePath + "copy.gif")
          Me.imgPasteInstallCode.Attributes.Add("src", ModulePath + "paste.gif")
        End If

        Me.imgCopyActivationCode.Attributes.Add("onclick", "CopyToClipboard('" & Me.txtActivationCode.ClientID & "')")
        Me.imgPasteInstallCode.Attributes.Add("onclick", "PasteFromClipboard('" & Me.txtInstalationCode.ClientID & "')")
        Me.cboLicenseType.Attributes.Add("onchange", "ShowLicenseType('" & Me.cboLicenseType.ClientID & "','" & Me.ClientID & "')")
        Me.imgExpireDateCalendar.Attributes.Add("src", ModulePath + "calendar.gif")
        Me.imgEmail.Attributes.Add("src", ModulePath + "email.gif")
        Me.chkDirectCodeEdit.Attributes.Add("onclick", "ChangeCodeEditCode('" & Me.ClientID & "');")
        Me.chkRegisterByCustomer.Attributes.Add("onclick", "GenByCustomer('" & Me.chkRegisterByCustomer.ClientID & "','" & Me.ClientID & "')")

        'these won't be necessary in next release after 3.2.0
        If ClientAPI.BrowserSupportsFunctionality(ClientAPI.ClientFunctionality.XMLHTTP) _
          AndAlso ClientAPI.BrowserSupportsFunctionality(ClientAPI.ClientFunctionality.XML) Then
          ClientAPI.RegisterClientReference(Me.Page, ClientAPI.ClientNamespaceReferences.dnn_xml)
          ClientAPI.RegisterClientReference(Me.Page, ClientAPI.ClientNamespaceReferences.dnn_xmlhttp)

          'Only this line will be necessary after 3.2
          '1 - ProductID, 2 - CustomerID, 3 - License type, 4 - Registered level, 5 - Install code, 6 - Days / Date expire
          Me.btnGenerate.Attributes.Add("onclick", "DisableAllControlsCode('" & Me.ClientID & "');this.style.cursor='url(" & ModulePath & "dnnwork_arrow.ani),wait;';document.body.style.cursor='url(" & ModulePath & "dnnwork_arrow.ani),wait;';" & _
            ClientAPI.GetCallbackEventReference(Me, "'activate'" & "+'|'+ dnn.dom.getById('" & Me.cboProduct.ClientID & "').value+'|'+dnn.dom.getById('" & Me.cboCustomer.ClientID & "').value+'|'+dnn.dom.getById('" & Me.cboLicenseType.ClientID & "').value+'|'+dnn.dom.getById('" & Me.cboRegisteredLevel.ClientID & "').value+'|'+dnn.dom.getById('" & Me.txtInstalationCode.ClientID & "').value+'|'+dnn.dom.getById('" & Me.txtExpireDate.ClientID & "').value", "successFuncActivate", "'" & Me.ClientID & "'", "errorFunc") & "; dnn.dom.getById('" & Me.txtActivationCode.ClientID & "').value='..generating activation license code - please wait..';")
          Me.txtInstalationCode.Attributes.Add("onkeyup", ClientAPI.GetCallbackEventReference(Me, "'getuser'" & "+'|'+dnn.dom.getById('" & Me.txtInstalationCode.ClientID & "').value", "successFuncGetUser", "'" & Me.ClientID & "'", "errorFunc"))
          Me.txtInstalationCode.Attributes.Add("onBlur", ClientAPI.GetCallbackEventReference(Me, "'getuser'" & "+'|'+dnn.dom.getById('" & Me.txtInstalationCode.ClientID & "').value", "successFuncGetUser", "'" & Me.ClientID & "'", "errorFunc"))
          Me.imgEmail.Attributes.Add("onclick", ClientAPI.GetCallbackEventReference(Me, "'email'" & "+'|'+dnn.dom.getById('" & Me.cboCustomer.ClientID & "').value+'|'+dnn.dom.getById('" & Me.txtUserName.ClientID & "').value+'|'+encodeURIComponent(dnn.dom.getById('" & Me.txtInstalationCode.ClientID & "').value)+'|'+encodeURIComponent(dnn.dom.getById('" & Me.txtActivationCode.ClientID & "').value)+'|'+encodeURIComponent(dnn.dom.getById('" & Me.cboProduct.ClientID & "').value)", "successFuncEmail", "'" & Me.ClientID & "'", "errorFunc"))

        End If

        'setting the js file
        If Me.Page.IsClientScriptBlockRegistered("alugen3netMain.js") = False Then
          Me.Page.RegisterClientScriptBlock("alugen3netMain.js", "<script language=javascript src=""" & Me.ModulePath & "alugen3netMain.js""></script>")
        End If
        If Me.Page.IsClientScriptBlockRegistered("popup.js") = False Then
          Me.Page.RegisterClientScriptBlock("popup.js", "<script language=javascript src=""" & Me.ModulePath & "popup.js""></script>")
        End If

        If Not Page.IsPostBack Then
          'fill combos
          cboProduct.DataSource = objAlugen3NETController.ListFull(PortalId, ModuleId)
          cboProduct.DataTextField = "ProductNameVersion"
          cboProduct.DataValueField = "ProductID"
          cboProduct.DataBind()
          cboCustomer.DataSource = objAlugenCustomersController.List(PortalId, ModuleId)
          cboCustomer.DataTextField = "CustomerName"
          cboCustomer.DataValueField = "CustomerID"
          cboCustomer.DataBind()

          txtExpireDate.Text = Date.Now.ToString("d") '"yyyy/MM/dd"

          cmdDelete.Attributes.Add("onClick", "javascript:return confirm('" & Localization.GetString("DeleteItem.Text", LocalResourceFile) & "');")
          If Not Null.IsNull(codeId) Then
            objAlugenCodeInfo = objAlugenCodeController.Get(codeId)
            If Not objAlugenCodeInfo Is Nothing Then
              'Load data
              txtActivationCode.Text = objAlugenCodeInfo.CodeActivationCode
              txtUserName.Text = objAlugenCodeInfo.CodeUserName
              txtInstalationCode.Text = objAlugenCodeInfo.CodeInstalationCode

              cboCustomer.SelectedValue = objAlugenCodeInfo.CustomerID.ToString
              cboProduct.SelectedValue = objAlugenCodeInfo.ProductID.ToString
              cboLicenseType.SelectedValue = objAlugenCodeInfo.CodeLicenseType.ToString
              If objAlugenCodeInfo.CodeRegisteredLevel >= 0 And objAlugenCodeInfo.CodeRegisteredLevel < 100 Then
                cboRegisteredLevel.SelectedValue = objAlugenCodeInfo.CodeRegisteredLevel.ToString
              Else
                cboRegisteredLevel.SelectedIndex = -1
              End If
              chkRegisterByCustomer.Checked = objAlugenCodeInfo.GenByCustomer
              If chkRegisterByCustomer.Checked Then
                pnlGenByCustomer.Style("display") = "none"
              Else
                pnlGenByCustomer.Style("display") = "block"
              End If

              Select Case objAlugenCodeInfo.CodeLicenseType
                Case 0 'time locked
                  txtExpireDate.Text = objAlugenCodeInfo.CodeExpireDate.ToString("d") '"yyyy/MM/dd"
                  pnlExpireDate.Style("display") = "block"
                  pnlPeriodicDays.Style("display") = "none"
                  txtExpireDate.Style("display") = "block"
                  imgExpireDateCalendar.Style("display") = "block"
                Case 1 'periodic
                  txtExpireDate.Text = objAlugenCodeInfo.CodeDays.ToString()
                  pnlExpireDate.Style("display") = "none"
                  pnlPeriodicDays.Style("display") = "block"
                  txtExpireDate.Style("display") = "block"
                  imgExpireDateCalendar.Style("display") = "none"
                Case 2 'permanent
                  txtExpireDate.Text = "0"
                  pnlExpireDate.Style("display") = "none"
                  pnlPeriodicDays.Style("display") = "none"
                  txtExpireDate.Style("display") = "none"
                  imgExpireDateCalendar.Style("display") = "none"
                Case Else
              End Select

              cmdDelete.Visible = True
            Else ' security violation attempt to access item not related to this Module
              Response.Redirect(NavigateURL(), True)
            End If

          End If

          'this needs to execute always to the client script code is registred in InvokePopupCal
          Me.imgExpireDateCalendar.Attributes.Add("onclick", Common.Utilities.Calendar.InvokePopupCal(txtExpireDate))

        End If
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

    Private Sub cmdUpdate_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdUpdate.Click
      Try
        Dim objModules As New ModuleController
        Dim mysettings As New Hashtable
        mysettings = objModules.GetModuleSettings(ModuleId)

        ' Only Update if the Entered Data is Valid
        If Page.IsValid = True Then
          If cboCustomer.SelectedValue <> "" And cboProduct.SelectedValue <> "" Then
            Dim objAlugenCustomersController As New AlugenCustomersController
            Dim objAlugenCustomersInfo As New AlugenCustomersInfo
            objAlugenCustomersInfo = objAlugenCustomersController.Get(Int32.Parse(cboCustomer.SelectedValue))
            Dim objAlugenCodeInfo As New AlugenCodeInfo

            objAlugenCodeInfo = CType(CBO.InitializeObject(objAlugenCodeInfo, GetType(AlugenCodeInfo)), AlugenCodeInfo)

            'bind text values to object
            objAlugenCodeInfo.CodeID = codeId
            objAlugenCodeInfo.PortalID = PortalId
            objAlugenCodeInfo.ModuleID = ModuleId
            objAlugenCodeInfo.CustomerID = Int32.Parse(cboCustomer.SelectedValue)
            objAlugenCodeInfo.ProductID = Int32.Parse(cboProduct.SelectedValue)
            objAlugenCodeInfo.CodeLicenseType = Int32.Parse(cboLicenseType.SelectedValue)
            objAlugenCodeInfo.CodeRegisteredLevel = Int32.Parse(cboRegisteredLevel.SelectedValue)
            objAlugenCodeInfo.CodeDays = 0
            objAlugenCodeInfo.CodeExpireDate = DateTime.Now
            Try
              Select Case objAlugenCodeInfo.CodeLicenseType
                Case 0 'time locked
                  objAlugenCodeInfo.CodeExpireDate = DateTime.Parse(txtExpireDate.Text)
                Case 1 'periodic
                  objAlugenCodeInfo.CodeDays = Integer.Parse(txtExpireDate.Text)
                  objAlugenCodeInfo.CodeExpireDate = System.DateTime.FromOADate(Now.ToOADate + CShort(objAlugenCodeInfo.CodeDays))
                Case 2 'permanent
                Case Else
              End Select
            Catch
            End Try
            objAlugenCodeInfo.GenByCustomer = chkRegisterByCustomer.Checked
            objAlugenCodeInfo.CodeInstalationCode = txtInstalationCode.Text
            objAlugenCodeInfo.CodeUserName = txtUserName.Text
            objAlugenCodeInfo.CodeActivationCode = txtActivationCode.Text

            objAlugenCodeInfo.CreatedByUser = UserInfo.Username
            objAlugenCodeInfo.CreatedDate = DateAndTime.Now

            Dim objAlugenCodeController As New AlugenCodeController
            Dim message As String
            Dim strAction As String
            If Null.IsNull(codeId) Then
              objAlugenCodeController.Add(objAlugenCodeInfo)
              strAction = "added"
            Else
              objAlugenCodeController.Update(objAlugenCodeInfo)
              strAction = "updated"
            End If
            'send email notification
            If chkSendEmailNotification.Checked Then
              message = "<html>Dear <strong><i>" & objAlugenCustomersInfo.CustomerContactPerson & "</i></strong><br>" & vbCr
              message = message & "from <strong><i>" & objAlugenCustomersInfo.CustomerName & "</i></strong><br>" & vbCrLf
              message = message & "<br>" & vbCrLf
              message = message & "We will like to inform you that following valid activation code has been succsefully " & strAction & " into our license database." & "<br>" & vbCrLf
              message = message & "<br>" & vbCrLf
              message = message & "Product: <strong>" & cboProduct.SelectedItem.ToString() & "</strong><br>" & vbCrLf
              Dim strLicensetype As String
              Select Case objAlugenCodeInfo.CodeLicenseType
                Case 0 'time locked
                  strLicensetype = "Time Locked"
                Case 1 'period
                  strLicensetype = "Periodic"
                Case 2 'permanent
                  strLicensetype = "Permanent"
              End Select
              message = message & "License type: <strong>" & strLicensetype & "</strong><br>" & vbCrLf
              Select Case objAlugenCodeInfo.CodeLicenseType
                Case 0, 1
                  message = message & "Expire date: <strong>" & objAlugenCodeInfo.CodeExpireDate.ToString("d") & "</strong><br>" & vbCrLf
                Case 2
              End Select
              message = message & "Registered level: <strong>" & objAlugenCodeInfo.CodeRegisteredLevel.ToString() & "</strong><br>" & vbCrLf
              message = message & "<br>" & vbCrLf
              message = message & "User name: <br>" & vbCrLf
              message = message & "---------------------------------------------<br>" & vbCrLf
              message = message & "<strong><font face='Courier'>" & Server.HtmlEncode(objAlugenCodeInfo.CodeUserName.ToString()) & "</font></strong><br>" & vbCrLf
              message = message & "---------------------------------------------<br>" & vbCrLf
              message = message & "Installation code: <br>"
              message = message & "---------------------------------------------<br>" & vbCrLf
              message = message & "<strong><font face='Courier'>" & Server.HtmlEncode(objAlugenCodeInfo.CodeInstalationCode.ToString()) & "</font></strong><br>" & vbCrLf
              message = message & "---------------------------------------------<br>" & vbCrLf
              message = message & "Activation code: <br>---------------------------------------------<br>" & vbCrLf
              message = message & "<strong><font face='Courier'>" & objAlugenCodeInfo.CodeActivationCode.Replace(vbCrLf, "<BR>") & "</font></strong><br>" & vbCrLf
              message = message & "---------------------------------------------<br>" & vbCrLf
              message = message & "<br>" & vbCrLf
              message = message & "Kind redards," & "<br>" & vbCrLf
              message = message & PortalSettings.PortalName & " sales departament.</html>" & vbCrLf

              Dim strSendTo As String

              'send email to Alugen3NET administrator
              strSendTo = CType(mysettings("AdminEmail"), String)
              If strSendTo.Trim.Trim.Length > 0 Then
                SendEmail("Code " & strAction & " succesfully!", message, strSendTo)
              End If
              'get sendto customer email adress
              If objAlugenCustomersInfo.UseUserEmail And objAlugenCustomersInfo.AssociatedUser <> "(none)" Then
                Dim objUsers As New UserController
                Dim objUserInfo As UserInfo
                objUserInfo = objUsers.GetUserByUsername(PortalId, objAlugenCustomersInfo.AssociatedUser)
                If Not Null.IsNull(objUserInfo) Then
                  strSendTo = objUserInfo.Membership.Email
                Else
                  strSendTo = objAlugenCustomersInfo.CustomerEmail
                End If
              Else
                strSendTo = objAlugenCustomersInfo.CustomerEmail
              End If
              'send email to customer (associater user) adress
              If strSendTo.Trim.Trim.Length > 0 Then
                SendEmail("Code " & strAction & " succesfully!", message, strSendTo)
              End If
            End If

            ' Redirect back to the portal home page
            strMessage = "Code " & strAction & " succesfully!"
            Response.Redirect(NavigateURL(TabId, "", "lastPage=d&Message=" & XmlConvert.EncodeName(strMessage) & "&MsgModule=" & ModuleId.ToString()), True)
          Else
            'null value (zero lenght)
          End If
        End If
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdCancel.Click
      Try
        Response.Redirect(NavigateURL(TabId, "", "lastPage=d"), True)
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdDelete.Click
      Try
        Dim objModules As New ModuleController
        Dim mysettings As New Hashtable
        mysettings = objModules.GetModuleSettings(ModuleId)

        If Not Null.IsNull(codeId) Then
          Dim objAlugenCodeController As New AlugenCodeController
          Dim objAlugenCodeInfo As New AlugenCodeInfo
          objAlugenCodeInfo = objAlugenCodeController.Get(codeId)
          Dim objAlugenCustomersController As New AlugenCustomersController
          Dim objAlugenCustomersInfo As New AlugenCustomersInfo
          objAlugenCustomersInfo = objAlugenCustomersController.Get(objAlugenCodeInfo.CustomerID)

          Dim strAction As String = "deleted"
          Dim message As String = "Code " & strAction & " succesfully!"

          'delete code
          objAlugenCodeController.Delete(codeId)
          'send email notification
          If chkSendEmailNotification.Checked Then
            message = "<html>Dear <strong><i>" & objAlugenCustomersInfo.CustomerContactPerson & "</i></strong><br>" & vbCr
            message = message & "from <strong><i>" & objAlugenCustomersInfo.CustomerName & "</i></strong><br>" & vbCrLf
            message = message & "<br>" & vbCrLf
            message = message & "We will like to inform you that following valid activation code has been succsefully " & strAction & " into our license database." & "<br>" & vbCrLf
            message = message & "<br>" & vbCrLf
            message = message & "Product: <strong>" & cboProduct.SelectedItem.ToString() & "</strong><br>" & vbCrLf
            Dim strLicensetype As String
            Select Case objAlugenCodeInfo.CodeLicenseType
              Case 0 'time locked
                strLicensetype = "Time Locked"
              Case 1 'period
                strLicensetype = "Periodic"
              Case 2 'permanent
                strLicensetype = "Permanent"
            End Select
            message = message & "License type: <strong>" & strLicensetype & "</strong><br>" & vbCrLf
            Select Case objAlugenCodeInfo.CodeLicenseType
              Case 0, 1
                message = message & "Expire date: <strong>" & objAlugenCodeInfo.CodeExpireDate.ToString("d") & "</strong><br>" & vbCrLf
              Case 2
            End Select
            message = message & "Registered level: <strong>" & objAlugenCodeInfo.CodeRegisteredLevel.ToString() & "</strong><br>" & vbCrLf
            message = message & "<br>" & vbCrLf
            message = message & "User name: <br>" & vbCrLf
            message = message & "---------------------------------------------<br>" & vbCrLf
            message = message & "<strong><font face='Courier'>" & Server.HtmlEncode(objAlugenCodeInfo.CodeUserName.ToString()) & "</font></strong><br>" & vbCrLf
            message = message & "---------------------------------------------<br>" & vbCrLf
            message = message & "Installation code: <br>"
            message = message & "---------------------------------------------<br>" & vbCrLf
            message = message & "<strong><font face='Courier'>" & Server.HtmlEncode(objAlugenCodeInfo.CodeInstalationCode.ToString()) & "</font></strong><br>" & vbCrLf
            message = message & "---------------------------------------------<br>" & vbCrLf
            message = message & "Activation code: <br>---------------------------------------------<br>" & vbCrLf
            message = message & "<strong><font face='Courier'>" & objAlugenCodeInfo.CodeActivationCode.Replace(vbCrLf, "<BR>") & "</font></strong><br>" & vbCrLf
            message = message & "---------------------------------------------<br>" & vbCrLf
            message = message & "<br>" & vbCrLf
            message = message & "Kind redards," & "<br>" & vbCrLf
            message = message & Me.PortalSettings.PortalName & " sales departament.</html>" & vbCrLf
            Dim strSendTo As String
            'send email to Alugen3NET administrator
            strSendTo = CType(mysettings("AdminEmail"), String)
            If strSendTo.Trim.Trim.Length > 0 Then
              SendEmail("Code " & strAction & " succesfully!", message, strSendTo)
            End If
            'get sendto customer email adress
            If objAlugenCustomersInfo.UseUserEmail And objAlugenCustomersInfo.AssociatedUser <> "(none)" Then
              Dim objUsers As New UserController
              Dim objUserInfo As UserInfo
              objUserInfo = objUsers.GetUserByUsername(PortalId, objAlugenCustomersInfo.AssociatedUser)
              If Not Null.IsNull(objUserInfo) Then
                strSendTo = objUserInfo.Membership.Email
              Else
                strSendTo = objAlugenCustomersInfo.CustomerEmail
              End If
            Else
              strSendTo = objAlugenCustomersInfo.CustomerEmail
            End If
            'send email to customer (associater user) adress
            If strSendTo.Trim.Trim.Length > 0 Then
              SendEmail("Code " & strAction & " succesfully!", message, strSendTo)
            End If
          End If
        End If

      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
        Exit Sub
      End Try

      ' Redirect back to the portal home page
      strMessage = "Code deleted succesfully!"
      Response.Redirect(NavigateURL(TabId, "", "lastPage=d&Message=" & XmlConvert.EncodeName(strMessage) & "&MsgModule=" & ModuleId.ToString()), True)

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
        Dim strSplit As String()
        strSplit = Split(eventArgument, "|")
        If strSplit(0) = "activate" Then
          Dim myProductInfo As New ProductInfo
          'generate activation code
          Dim objAlugen3NETController As New AlugenProductsController
          Dim objAlugen3NETInfo As New AlugenProductsInfo
          objAlugen3NETInfo = objAlugen3NETController.Get(Int32.Parse(strSplit(1))) '1 - ProductID
          If Not Null.IsNull(objAlugen3NETInfo) Then
            'extract name and version
            Dim strName, strVer As String
            strName = objAlugen3NETInfo.ProductName
            myProductInfo.name = objAlugen3NETInfo.ProductName
            strVer = objAlugen3NETInfo.ProductVersion
            myProductInfo.Version = objAlugen3NETInfo.ProductVersion
            myProductInfo.VCode = objAlugen3NETInfo.ProductVCode
            myProductInfo.GCode = objAlugen3NETInfo.ProductGCode

            '2 - CustomerID

            Dim varLicType As ActiveLock3NET.ProductLicense.ALLicType
            Dim iLicType As Integer = Int32.Parse(strSplit(3)) '3 - License type
            Select Case iLicType
              Case 0 'Time locked
                varLicType = ActiveLock3NET.ProductLicense.ALLicType.allicTimeLocked
              Case 1 'Periodic
                varLicType = ActiveLock3NET.ProductLicense.ALLicType.allicPeriodic
              Case 2 'Permanent
                varLicType = ActiveLock3NET.ProductLicense.ALLicType.allicPermanent
              Case Else
                varLicType = ActiveLock3NET.ProductLicense.ALLicType.allicNone
            End Select

            Dim strExpire As String
            Dim sDays As String = strSplit(6) ' 6 - Days / Date expire
            strExpire = GetExpirationDate(iLicType, sDays)
            Dim strRegDate As String
            strRegDate = ActiveLockDateFormat(Now)

            Dim Lic As ActiveLock3NET.ProductLicense
            Dim sRegLevel As String = strSplit(4) ' 4 - Registered level
            Dim sInstallCode As String = strSplit(5) ' 5 - Install code
            ' Create a product license object without the product key or license key
            Lic = ActiveLock3Globals_definst.CreateProductLicense(strName, strVer, sInstallCode, ActiveLock3NET.ProductLicense.LicFlags.alfSingle, varLicType, "", sRegLevel, strExpire, , strRegDate)

            Dim LibKey As String
            Dim myLicGen As New aluGenerator
            LibKey = myLicGen.GenKey(Lic, sInstallCode, myProductInfo, sRegLevel)
            msg = Make64ByteChunks(LibKey)
            If msg.Trim.Length = 0 Then
              msg = "Activation code generation failed!"
              Throw New System.Exception(msg)
            End If
          End If
        ElseIf strSplit(0) = "getuser" Then
          'get user on key press
          Dim strInstCode As String
          strInstCode = strSplit(1)
          strInstCode = ActiveLock3Globals_definst.Base64Decode(strInstCode)
          Dim Index, i As Integer
          Index = 0 : i = 1
          ' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
          Do While i > 0
            i = InStr(Index + 1, strInstCode, vbLf)
            If i > 0 Then Index = i
          Loop
          ' user name starts from Index+1 to the end
          msg = Mid(strInstCode, Index + 1)
        ElseIf strSplit(0) = "email" Then
          'msg will contain email string body
          'mailto:strTO?subject=MySubject&body=MyBody
          Dim intIDCustomer As Integer = Integer.Parse(strSplit(1)) '1 - CustomerID
          Dim strUserCode As String = strSplit(2)
          Dim strInstallCode As String = strSplit(3)
          Dim strActivateCode As String = strSplit(4)
          Dim strProductId As String = strSplit(5)
          Dim objAlugenProductsController As New AlugenProductsController
          Dim objAlugenProductsInfo As New AlugenProductsInfo
          objAlugenProductsInfo = objAlugenProductsController.Get(Int32.Parse(strProductId)) '1 - ProductID

          Dim strSubject1 As String = "Activation code for customer ["
          Dim strSubject2 As String = "] user ["
          Dim strSubject3 As String = "] product ["
          Dim strSubject4 As String = "]"
          Dim bodymsg As String = String.Empty

          Dim objAlugenCustomersController As New AlugenCustomersController
          Dim objAlugenCustomersInfo As New AlugenCustomersInfo
          objAlugenCustomersInfo = objAlugenCustomersController.Get(intIDCustomer)
          If Not Null.IsNull(objAlugenCustomersInfo) Then
            msg = "mailto:" & objAlugenCustomersInfo.CustomerEmail.ToString & "?subject=" & strSubject1 & objAlugenCustomersInfo.CustomerName & strSubject2 & strUserCode & strSubject3 & objAlugenProductsInfo.ProductName & " - " & objAlugenProductsInfo.ProductVersion & strSubject4
            msg = msg & "&body="
            bodymsg = "Installcode:" & "%0D%0A" & strInstallCode & "%0D%0A" & "%0D%0A" & "Activation code:" & "%0D%0A" & HttpUtility.UrlEncode(strActivateCode)
            msg = msg & bodymsg
          End If
        End If
        Return msg
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Function

    ' Breaks a long string into chunks of 64-byte lines.
    Private Function Make64ByteChunks(ByRef strdata As String) As String
      Dim i As Integer
      Dim Count As Integer
      Count = Len(strdata)
      Dim sResult As String
      sResult = VB.Left(strdata, 64)
      i = 65
      While i <= Count
        sResult = sResult & vbCrLf & Mid(strdata, i, 64)
        i = i + 64
      End While
      Make64ByteChunks = sResult
    End Function

    Private Function GetExpirationDate(ByVal piLicType As Integer, ByVal psDays As String) As String
      Select Case piLicType
        Case 0 'time locked
          Dim expDate As DateTime
          expDate = DateTime.Parse(psDays)
          GetExpirationDate = expDate.Year.ToString & "/" & expDate.Month.ToString & "/" & expDate.Day.ToString
        Case 1 'Periodic
          GetExpirationDate = ActiveLockDateFormat(System.DateTime.FromOADate(Now.ToOADate + CShort(psDays)))
        Case 2 'Permanent
          GetExpirationDate = ActiveLockDateFormat(System.DateTime.Now)
        Case Else
          'trow error1
          Throw New System.Exception("License type not valid.")
      End Select
    End Function

    ' Normalizes date format to yyyy/mm/dd
    Private Function ActiveLockDateFormat(ByRef dt As Date) As String
      ActiveLockDateFormat = Year(dt) & "/" & Month(dt) & "/" & VB.Day(dt)
    End Function

    Public Function AppPath() As String
      Dim siteNameUrl As String
      AppPath = System.IO.Path.GetDirectoryName(Server.MapPath("AlugenCodesEdit.aspx"))
    End Function


    Private Sub SendEmail(ByVal strSubject As String, ByVal strMessage As String, ByVal toEmail As String)
      Try
        If PortalSettings.Email <> "" And strMessage.Length > 0 Then

          Dim ReturnMsg As String = Mail.SendMail(PortalSettings.Email, toEmail, "", _
                strSubject, strMessage, _
                "", "html", "", "", "", "")
          'Dim ReturnMsg As String = Mail.SendMail(PortalSettings.Email, toEmail, "", _
          'strSubject, Server.HtmlEncode(strMessage), _
          '"", "html", "", "", "", "")
          If ReturnMsg = "" Then
            'email success
          Else
            'email failure
          End If
        End If
      Catch exc As Exception    'Module failed to load
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub
  End Class
End Namespace

