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

  Public MustInherit Class AlugenCodesClientEdit
    Inherits Entities.Modules.PortalModuleBase
    Implements IClientAPICallbackEventHandler

#Region "Controls"
    Protected WithEvents cmdCancel As System.Web.UI.WebControls.LinkButton
    Protected WithEvents txtExpireDate As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtInstalationCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtUserName As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtActivationCode As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnGenerate As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents imgCopyActivationCode As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents imgPasteInstallCode As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents plExpireDate As DotNetNuke.UI.UserControls.LabelControl
    Protected WithEvents plPeriodicDays As DotNetNuke.UI.UserControls.LabelControl
    Protected WithEvents pnlExpireDate As System.Web.UI.WebControls.Panel
    Protected WithEvents pnlPeriodicDays As System.Web.UI.WebControls.Panel
    Protected WithEvents imgEmail As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents pnlCredits As System.Web.UI.WebControls.Panel
    Protected WithEvents imgFriendSoftware As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents imgActiveLock As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents cmdPasteCode As System.Web.UI.HtmlControls.HtmlAnchor
    Protected WithEvents txtProduct As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCustomer As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtLicenseType As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtRegisteredLevel As System.Web.UI.WebControls.TextBox
    Protected WithEvents idProduct As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents idCustomer As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents pnlMessage As System.Web.UI.WebControls.Panel
    Protected WithEvents pnlClientGeneration As System.Web.UI.WebControls.Panel
    Protected WithEvents lblMessage As System.Web.UI.WebControls.Label
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

        If Not Null.IsNull(codeId) Then
          objAlugenCodeInfo = objAlugenCodeController.Get(codeId)
          If Not objAlugenCodeInfo Is Nothing Then
            If UserInfo.Username <> objAlugenCustomersController.Get(objAlugenCodeInfo.CustomerID).AssociatedUser And Not UserInfo.IsSuperUser Then
              'security violation
              lblMessage.Text = "Unauthorized access!"
              lblMessage.Font.Bold = True
              lblMessage.ForeColor = Color.Red

              pnlClientGeneration.Visible = False
              pnlMessage.Visible = True
              'logging the attempt
              'Throw New System.Exception("Unauthorized access!")
              Exit Sub
            End If
          End If
        End If

        'default registration type
        txtExpireDate.Text = Date.Now.ToString("d") '"yyyy/MM/dd"
        pnlExpireDate.Style("display") = "block"
        pnlPeriodicDays.Style("display") = "none"
        txtExpireDate.Style("display") = "block"

        'inject JS call into body load
        Dim mbody As HtmlGenericControl = CType(Page.FindControl("body"), HtmlGenericControl)
        mbody.Attributes("onload") += ";inimodulePath('" & Me.ModulePath & "');"
        'add events
        Me.imgCopyActivationCode.Attributes.Add("onclick", "CopyToClipboard('" & Me.txtActivationCode.ClientID & "')")
        Me.imgPasteInstallCode.Attributes.Add("onclick", "PasteFromClipboard('" & Me.txtInstalationCode.ClientID & "')")
        Me.imgEmail.Attributes.Add("src", ModulePath + "email.gif")

        If Request.Browser.Browser = "Netscape" Then
          'copy command not working on FireFox, Netscape
          Me.imgCopyActivationCode.Attributes.Add("src", ModulePath + "copy_disabled.gif")
          Me.imgPasteInstallCode.Attributes.Add("src", ModulePath + "paste_disabled.gif")
        Else
          Me.imgCopyActivationCode.Attributes.Add("src", ModulePath + "copy.gif")
          Me.imgPasteInstallCode.Attributes.Add("src", ModulePath + "paste.gif")
        End If

        'these won't be necessary in next release after 3.2.0
        If ClientAPI.BrowserSupportsFunctionality(ClientAPI.ClientFunctionality.XMLHTTP) _
          AndAlso ClientAPI.BrowserSupportsFunctionality(ClientAPI.ClientFunctionality.XML) Then
          ClientAPI.RegisterClientReference(Me.Page, ClientAPI.ClientNamespaceReferences.dnn_xml)
          ClientAPI.RegisterClientReference(Me.Page, ClientAPI.ClientNamespaceReferences.dnn_xmlhttp)

          'Only this line will be necessary after 3.2
          Me.txtInstalationCode.Attributes.Add("onkeyup", ClientAPI.GetCallbackEventReference(Me, "'getuser'" & "+'|'+dnn.dom.getById('" & Me.txtInstalationCode.ClientID & "').value", "successFuncGetUser", "'" & Me.ClientID & "'", "errorFunc"))
          Me.txtInstalationCode.Attributes.Add("onBlur", ClientAPI.GetCallbackEventReference(Me, "'getuser'" & "+'|'+dnn.dom.getById('" & Me.txtInstalationCode.ClientID & "').value", "successFuncGetUser", "'" & Me.ClientID & "'", "errorFunc"))

        End If

        'setting the js file
        If Me.Page.IsClientScriptBlockRegistered("alugen3netMain.js") = False Then
          Me.Page.RegisterClientScriptBlock("alugen3netMain.js", "<script language=javascript src=""" & Me.ModulePath & "alugen3netMain.js""></script>")
        End If
        If Me.Page.IsClientScriptBlockRegistered("popup.js") = False Then
          Me.Page.RegisterClientScriptBlock("popup.js", "<script language=javascript src=""" & Me.ModulePath & "popup.js""></script>")
        End If


        If Not Page.IsPostBack Then

          If Not Null.IsNull(codeId) Then
            'objAlugenCodeInfo = objAlugenCodeController.Get(codeId)
            If Not objAlugenCodeInfo Is Nothing Then
              'Load data
              txtProduct.Text = objAlugen3NETController.Get(objAlugenCodeInfo.ProductID).ProductName & " - " & objAlugen3NETController.Get(objAlugenCodeInfo.ProductID).ProductVersion
              idProduct.Value = objAlugenCodeInfo.ProductID.ToString
              txtCustomer.Text = objAlugenCustomersController.Get(objAlugenCodeInfo.CustomerID).CustomerName
              idCustomer.Value = objAlugenCodeInfo.CustomerID.ToString
              txtRegisteredLevel.Text = objAlugenCodeInfo.CodeRegisteredLevel.ToString

              Select Case objAlugenCodeInfo.CodeLicenseType
                Case 0 'time locked
                  txtLicenseType.Text = "Time Locked"
                  txtExpireDate.Text = objAlugenCodeInfo.CodeExpireDate.ToString("d") '"yyyy/MM/dd"
                  pnlExpireDate.Style("display") = "block"
                  pnlPeriodicDays.Style("display") = "none"
                  txtExpireDate.Style("display") = "block"
                Case 1 'periodic
                  txtLicenseType.Text = "Periodic"
                  txtExpireDate.Text = objAlugenCodeInfo.CodeDays.ToString()
                  pnlExpireDate.Style("display") = "none"
                  pnlPeriodicDays.Style("display") = "block"
                  txtExpireDate.Style("display") = "block"
                Case 2 'permanent
                  txtLicenseType.Text = "Permanent"
                  txtExpireDate.Text = "0"
                  pnlExpireDate.Style("display") = "none"
                  pnlPeriodicDays.Style("display") = "none"
                  txtExpireDate.Style("display") = "none"
                Case Else
              End Select

              If objAlugenCodeInfo.GenByCustomer Then
                txtInstalationCode.Text = ""
                txtUserName.Text = ""
                txtActivationCode.Text = ""
                txtInstalationCode.ReadOnly = False
                imgPasteInstallCode.Visible = True
                btnGenerate.Visible = True
                'imgEmail.Visible = False
                'imgCopyActivationCode.Visible = False
              Else
                txtInstalationCode.Text = objAlugenCodeInfo.CodeInstalationCode
                txtUserName.Text = objAlugenCodeInfo.CodeUserName
                txtActivationCode.Text = objAlugenCodeInfo.CodeActivationCode
                txtInstalationCode.ReadOnly = True
                txtInstalationCode.ForeColor = Color.FromName("#424242")
                txtInstalationCode.BackColor = Color.FromName("#DBDBDB")
                imgPasteInstallCode.Visible = False
                btnGenerate.Visible = False
                'imgEmail.Visible = True
                'imgCopyActivationCode.Visible = True
              End If

              'these won't be necessary in next release after 3.2.0
              If ClientAPI.BrowserSupportsFunctionality(ClientAPI.ClientFunctionality.XMLHTTP) _
                AndAlso ClientAPI.BrowserSupportsFunctionality(ClientAPI.ClientFunctionality.XML) Then
                '1 - CodeID, 2 - Install code
                Me.btnGenerate.Attributes.Add("onclick", "if(!confirm('" & Localization.GetString("GenerateLicenseCode.Text", LocalResourceFile) & "',2)){return;};DisableAllControlsCodeClient('" & Me.ClientID & "');this.style.cursor='url(" & ModulePath & "dnnwork_arrow.ani),wait;';document.body.style.cursor='url(" & ModulePath & "dnnwork_arrow.ani),wait;';" & _
                  ClientAPI.GetCallbackEventReference(Me, "'activate'" & "+'|'+ '" & objAlugenCodeInfo.CodeID.ToString() & "'+'|'+encodeURIComponent(dnn.dom.getById('" & Me.txtInstalationCode.ClientID & "').value)", "successFuncActivateClient", "'" & Me.ClientID & "'", "errorFunc") & "; dnn.dom.getById('" & Me.txtActivationCode.ClientID & "').value='..generating activation license code - please wait..';")
                Me.imgEmail.Attributes.Add("onclick", ClientAPI.GetCallbackEventReference(Me, "'email'" & "+'|'+'" & objAlugenCodeInfo.CodeID.ToString & "'", "successFuncEmail", "'" & Me.ClientID & "'", "errorFunc"))
              End If

            Else ' security violation attempt to access item not related to this Module
              Response.Redirect(NavigateURL(), True)
            End If

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
        Dim objModules As New ModuleController
        Dim mysettings As New Hashtable
        mysettings = objModules.GetModuleSettings(ModuleId)

        Dim msg As String = String.Empty
        Dim strSplit As String()
        strSplit = Split(eventArgument, "|")
        If strSplit(0) = "activate" Then
          '1 - CodeID, 2 - Install code
          Dim objAlugenCodeController As New AlugenCodeController
          Dim objAlugenCodeInfo As AlugenCodeInfo
          Dim objAlugenCustomersController As New AlugenCustomersController
          Dim objAlugenCustomersInfo As New AlugenCustomersInfo
          Dim objAlugenProductsController As New AlugenProductsController
          Dim objAlugenProductsInfo As New AlugenProductsInfo

          objAlugenCodeInfo = objAlugenCodeController.Get(Int32.Parse(strSplit(1))) '1 - CodeID

          If Not objAlugenCodeInfo Is Nothing Then
            If UserInfo.Username <> objAlugenCustomersController.Get(objAlugenCodeInfo.CustomerID).AssociatedUser And Not UserInfo.IsSuperUser Then
              'security violation
              lblMessage.Text = "Unauthorized access!"
              pnlClientGeneration.Visible = False
              lblMessage.Font.Bold = True
              lblMessage.ForeColor = Color.Red

              pnlMessage.Visible = True
              'logging the attempt
              'Throw New System.Exception("Unauthorized access!")
              'Response.Redirect(NavigateURL(), True)
              Exit Function
            End If
          Else
            'ivalid codeid
            Throw New System.Exception("Invalid code!")
          End If

          Dim myProductInfo As New ProductInfo
          'generate activation code
          objAlugenProductsInfo = objAlugenProductsController.Get(objAlugenCodeInfo.ProductID)
          objAlugenCustomersInfo = objAlugenCustomersController.Get(objAlugenCodeInfo.CustomerID)

          If Not Null.IsNull(objAlugenProductsInfo) Then
            'extract name and version
            Dim strName, strVer As String
            strName = objAlugenProductsInfo.ProductName
            myProductInfo.name = objAlugenProductsInfo.ProductName
            strVer = objAlugenProductsInfo.ProductVersion
            myProductInfo.Version = objAlugenProductsInfo.ProductVersion
            myProductInfo.VCode = objAlugenProductsInfo.ProductVCode
            myProductInfo.GCode = objAlugenProductsInfo.ProductGCode

            Dim sDays As String
            Dim varLicType As ActiveLock3NET.ProductLicense.ALLicType
            Dim iLicType As Integer = objAlugenCodeInfo.CodeLicenseType
            Select Case iLicType
              Case 0 'Time locked
                varLicType = ActiveLock3NET.ProductLicense.ALLicType.allicTimeLocked
                sDays = objAlugenCodeInfo.CodeExpireDate.ToString("d")
              Case 1 'Periodic
                varLicType = ActiveLock3NET.ProductLicense.ALLicType.allicPeriodic
                sDays = objAlugenCodeInfo.CodeDays.ToString
              Case 2 'Permanent
                varLicType = ActiveLock3NET.ProductLicense.ALLicType.allicPermanent
                sDays = "0"
              Case Else
                varLicType = ActiveLock3NET.ProductLicense.ALLicType.allicNone
                sDays = "0"
            End Select

            Dim strExpire As String
            strExpire = GetExpirationDate(iLicType, sDays)
            Dim strRegDate As String
            strRegDate = ActiveLockDateFormat(Now)

            Dim Lic As ActiveLock3NET.ProductLicense
            Dim sRegLevel As String = objAlugenCodeInfo.CodeRegisteredLevel.ToString  'Registered level
            objAlugenCodeInfo.CodeInstalationCode = strSplit(2) '2 - Install code ????Base64Decode
            Dim strInstCode As String = ActiveLock3Globals_definst.Base64Decode(objAlugenCodeInfo.CodeInstalationCode)
            Dim Index, i As Integer
            Index = 0 : i = 1
            ' Get to the last vbLf, which denotes the ending of the lock code and beginning of user name.
            Do While i > 0
              i = InStr(Index + 1, strInstCode, vbLf)
              If i > 0 Then Index = i
            Loop
            ' user name starts from Index+1 to the end
            Dim strUser As String = Mid(strInstCode, Index + 1)
            objAlugenCodeInfo.CodeUserName = strUser

            Dim sInstallCode As String = objAlugenCodeInfo.CodeInstalationCode
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

            'update code
            objAlugenCodeInfo.CodeActivationCode = LibKey
            objAlugenCodeInfo.GenByCustomer = False 'mark to generated
            objAlugenCodeInfo.CreatedDate = DateTime.Now
            objAlugenCodeController.Update(objAlugenCodeInfo)

            'send email notification
            Dim strAction As String = "generated by customer"
            Dim message As String = "Code " & strAction & " succesfully!"
            Dim EmailNotification As Boolean = CType(mysettings("EmailNotification"), Boolean)
            If EmailNotification Then
              message = "<html>Dear <strong><i>" & objAlugenCustomersInfo.CustomerContactPerson & "</i></strong><br>" & vbCr
              message = message & "from <strong><i>" & objAlugenCustomersInfo.CustomerName & "</i></strong><br>" & vbCrLf
              message = message & "<br>" & vbCrLf
              message = message & "We will like to inform you that following valid activation code has been succsefully " & strAction & " into our license database." & "<br>" & vbCrLf
              message = message & "<br>" & vbCrLf
              message = message & "Product: <strong>" & objAlugenProductsInfo.ProductName & " - " & objAlugenProductsInfo.ProductVersion & "</strong><br>" & vbCrLf
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

              Dim objModuleController As New DotNetNuke.Entities.Modules.ModuleController
              Dim strSendTo As String
              strSendTo = CType(objModuleController.GetModuleSettings(CType(mysettings("AlugenMainModule"), Integer))("AdminEmail"), String)
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

          Dim objAlugenCodeController As New AlugenCodeController
          Dim objAlugenCodeInfo As AlugenCodeInfo
          objAlugenCodeInfo = objAlugenCodeController.Get(Int32.Parse(strSplit(1))) '1 - CodeID
          Dim objAlugenProductsController As New AlugenProductsController
          Dim objAlugenProductsInfo As New AlugenProductsInfo
          objAlugenProductsInfo = objAlugenProductsController.Get(objAlugenCodeInfo.ProductID)
          Dim objAlugenCustomersController As New AlugenCustomersController
          Dim objAlugenCustomersInfo As New AlugenCustomersInfo
          objAlugenCustomersInfo = objAlugenCustomersController.Get(objAlugenCodeInfo.CustomerID)

          Dim intIDCustomer As Integer = objAlugenCodeInfo.CustomerID
          Dim strUserCode As String = objAlugenCodeInfo.CodeUserName
          Dim strInstallCode As String = objAlugenCodeInfo.CodeInstalationCode
          Dim strActivateCode As String = objAlugenCodeInfo.CodeActivationCode

          Dim strSubject1 As String = "Activation code for customer ["
          Dim strSubject2 As String = "] user ["
          Dim strSubject3 As String = "] product ["
          Dim strSubject4 As String = "]"
          Dim bodymsg As String = String.Empty

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

