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
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports DotNetNuke
Imports FriendSoftware.DNN.Modules.Alugen3NET.Business

Namespace FriendSoftware.DNN.Modules.Alugen3NET

  Public MustInherit Class Settings
    Inherits Entities.Modules.ModuleSettingsBase

#Region "Controls"
#End Region

    Public Overrides Sub LoadSettings()
      Try
        If Not Page.IsPostBack Then
          ' Load settings from TabModuleSettings: specific to this instance
          Dim PriceDecimals As String = CType(ModuleSettings("PriceDecimals"), String)
          cboProductsPriceDecimals.SelectedValue = PriceDecimals

          Dim EnableFiltersByDefault As Boolean = CType(ModuleSettings("EnableFiltersByDefault"), Boolean)
          chkEnableFiltersByDefault.Checked = EnableFiltersByDefault

          Dim AdminEmail As String = CType(ModuleSettings("AdminEmail"), String)
          txtAdminEmail.Text = AdminEmail

          Dim HideCredits As Boolean = CType(ModuleSettings("HideCredits"), Boolean)
          chkHideCredits.Checked = HideCredits

        End If
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

    Public Overrides Sub UpdateSettings()
      Try
        Dim objModules As New Entities.Modules.ModuleController

        ' Update TabModuleSettings
        'objModules.UpdateTabModuleSetting(TabModuleId, "PriceDecimals", cboProductsPriceDecimals.SelectedValue.ToString())
        'objModules.UpdateTabModuleSetting(TabModuleId, "EnableFiltersByDefault", chkEnableFiltersByDefault.Checked.ToString())
        'objModules.UpdateTabModuleSetting(TabModuleId, "AdminEmail", txtAdminEmail.Text.Trim)
        'objModules.UpdateTabModuleSetting(TabModuleId, "HideCredits", chkHideCredits.Checked.ToString())
        'Update ModuleSttings
        objModules.UpdateModuleSetting(ModuleId, "PriceDecimals", cboProductsPriceDecimals.SelectedValue.ToString())
        objModules.UpdateModuleSetting(ModuleId, "EnableFiltersByDefault", chkEnableFiltersByDefault.Checked.ToString())
        objModules.UpdateModuleSetting(ModuleId, "AdminEmail", txtAdminEmail.Text.Trim)
        objModules.UpdateModuleSetting(ModuleId, "HideCredits", chkHideCredits.Checked.ToString())

      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents cboProductsPriceDecimals As System.Web.UI.WebControls.DropDownList
    Protected WithEvents plProductsPriceDecimals As DotNetNuke.UI.UserControls.LabelControl
    Protected WithEvents plProductsPriceDecimals2 As System.Web.UI.WebControls.Label
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents chkHideCredits As System.Web.UI.WebControls.CheckBox
    Protected WithEvents chkEnableFiltersByDefault As System.Web.UI.WebControls.CheckBox
    Protected WithEvents txtAdminEmail As System.Web.UI.WebControls.TextBox

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
