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

Namespace FriendSoftware.DNN.Modules.Alugen3NET

  Public MustInherit Class AlugenCustomersEdit
    Inherits Entities.Modules.PortalModuleBase


#Region "Controls"
    Protected WithEvents cmdUpdate As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdCancel As System.Web.UI.WebControls.LinkButton
    Protected WithEvents cmdDelete As System.Web.UI.WebControls.LinkButton
    Protected WithEvents txtCustomerName As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCustomerAddress As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCustomerContactPerson As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCustomerPhone As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtCustomerEmail As System.Web.UI.WebControls.TextBox
    Protected WithEvents chkUseUserEmail As System.Web.UI.WebControls.CheckBox
    Protected WithEvents pnlCredits As System.Web.UI.WebControls.Panel
    Protected WithEvents imgFriendSoftware As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents imgActiveLock As System.Web.UI.HtmlControls.HtmlImage
    Protected WithEvents cboAssociatedUser As System.Web.UI.WebControls.DropDownList
#End Region

#Region "Private Members"
    Private customerId As Integer
    Private strMessage As String
    Private Const NoUserString As String = "(none)"
#End Region

#Region "Event Handlers"
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
      Try
        Dim objModules As New ModuleController
        Dim mysettings As New Hashtable
        mysettings = objModules.GetModuleSettings(ModuleId)

        Dim objAlugenCustomersController As New AlugenCustomersController
        Dim objAlugenCustomersInfo As New AlugenCustomersInfo
        Dim objUsers As New UserController

        'hide Credits
        pnlCredits.Visible = Not CType(mysettings("HideCredits"), Boolean)
        If pnlCredits.Visible Then
          Me.imgFriendSoftware.Attributes.Add("src", ModulePath + "friendsoftware_logo.JPG")
          Me.imgActiveLock.Attributes.Add("src", ModulePath + "activelock_logo.gif")
        End If

        ' Determine ItemId
        If Not (Request.Params("CustomerId") Is Nothing) Then
          customerId = Int32.Parse(Request.Params("CustomerId"))
        Else
          customerId = Null.NullInteger()
        End If

        If Not Page.IsPostBack Then

          'fill users combo
          cboAssociatedUser.Items.Clear()
          cboAssociatedUser.DataSource = objUsers.GetUsers(PortalId, False, False)
          cboAssociatedUser.DataTextField = "Username"
          cboAssociatedUser.DataValueField = "Username"
          cboAssociatedUser.DataBind()
          'no associated user
          cboAssociatedUser.Items.Insert(0, NoUserString)

          cmdDelete.Attributes.Add("onClick", "javascript:return confirm('" & Localization.GetString("DeleteItem.Text", LocalResourceFile) & "');")
          If Not Null.IsNull(customerId) Then
            objAlugenCustomersInfo = objAlugenCustomersController.Get(customerId)
            If Not objAlugenCustomersInfo Is Nothing Then

              'Load data
              txtCustomerName.Text = objAlugenCustomersInfo.CustomerName
              txtCustomerAddress.Text = objAlugenCustomersInfo.CustomerAddress
              txtCustomerContactPerson.Text = objAlugenCustomersInfo.CustomerContactPerson
              txtCustomerPhone.Text = objAlugenCustomersInfo.CustomerPhone
              txtCustomerEmail.Text = objAlugenCustomersInfo.CustomerEmail
              If Not cboAssociatedUser.Items.FindByValue(objAlugenCustomersInfo.AssociatedUser) Is Nothing Then
                cboAssociatedUser.Items.FindByValue(objAlugenCustomersInfo.AssociatedUser).Selected = True
              End If
              'If Not Null.IsNull(objAlugenCustomersInfo.AssociatedUser) Then
              '  'cboAssociatedUser.SelectedValue = objAlugenCustomersInfo.AssociatedUser
              'Else
              '  'cboAssociatedUser.SelectedValue = NoUserString
              'End If
              chkUseUserEmail.Checked = objAlugenCustomersInfo.UseUserEmail

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
          If txtCustomerName.Text.Trim.Length > 0 Then
            Dim objAlugenCustomersInfo As New AlugenCustomersInfo
            Dim strAction As String

            objAlugenCustomersInfo = CType(CBO.InitializeObject(objAlugenCustomersInfo, GetType(AlugenCustomersInfo)), AlugenCustomersInfo)

            'bind text values to object
            objAlugenCustomersInfo.CustomerID = customerId
            objAlugenCustomersInfo.PortalID = PortalId
            objAlugenCustomersInfo.ModuleID = ModuleId
            objAlugenCustomersInfo.CustomerName = txtCustomerName.Text
            objAlugenCustomersInfo.CustomerContactPerson = txtCustomerContactPerson.Text
            objAlugenCustomersInfo.CustomerAddress = txtCustomerAddress.Text
            objAlugenCustomersInfo.CustomerPhone = txtCustomerPhone.Text
            objAlugenCustomersInfo.CustomerEmail = txtCustomerEmail.Text
            If Not (cboAssociatedUser.SelectedValue = NoUserString) Then
              objAlugenCustomersInfo.AssociatedUser = cboAssociatedUser.SelectedValue
            End If
            objAlugenCustomersInfo.UseUserEmail = chkUseUserEmail.Checked

            Dim objCtlAlugenCustomers As New AlugenCustomersController
            If Null.IsNull(customerId) Then
              strAction = "added"
              objCtlAlugenCustomers.Add(objAlugenCustomersInfo)
            Else
              strAction = "updated"
              objCtlAlugenCustomers.Update(objAlugenCustomersInfo)
            End If

            ' Redirect back to the portal home page
            strMessage = "Customer " & strAction & " succesfully!"
            Response.Redirect(NavigateURL(TabId, "", "lastPage=c&Message=" & XmlConvert.EncodeName(strMessage) & "&MsgModule=" & ModuleId.ToString()), True)
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
        Response.Redirect(NavigateURL(TabId, "", "lastPage=c"), True)
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

    Private Sub cmdDelete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdDelete.Click
      Dim mystrMessage As String
      Try
        If Not Null.IsNull(customerId) Then
          Dim objAlugenCustomersController As New AlugenCustomersController
          Dim objAlugenCustomersInfo As New AlugenCustomersInfo
          objAlugenCustomersInfo = objAlugenCustomersController.Get(customerId)

          'check if customer have been used in code generation
          Dim objAlugenCodeController As New AlugenCodeController
          Dim myListByCustomer As ArrayList = objAlugenCodeController.ListByCustomerID(PortalId, ModuleId, customerId)

          If myListByCustomer.Count > 0 Then
            mystrMessage = "Customer is in use!"
            Throw New System.Exception("Customer " & objAlugenCustomersInfo.CustomerName & " cannot be deleted beacause it have generated codes! Plese delete those codes first")
          End If

          Dim objCtlAlugenCustomers As New AlugenCustomersController
          objCtlAlugenCustomers.Delete(customerId)
        End If

      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
        strMessage = "Customer deletion fail!"
        strMessage = strMessage & "<br>" & mystrMessage
        Response.Redirect(NavigateURL(TabId, "", "lastPage=c&Message=" & XmlConvert.EncodeName(strMessage) & "&IsError=true" & "&MsgModule=" & ModuleId.ToString()), True)
        Exit Sub
      End Try

      ' Redirect back to the portal home page
      strMessage = "Customer deleted succesfully!"
      Response.Redirect(NavigateURL(TabId, "", "lastPage=c&Message=" & XmlConvert.EncodeName(strMessage) & "&MsgModule=" & ModuleId.ToString()), True)

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

  End Class

End Namespace
