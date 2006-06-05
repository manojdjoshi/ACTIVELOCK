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
Imports DotNetNuke.Entities.Modules
Imports FriendSoftware.DNN.Modules.Alugen3NET.Business

Namespace FriendSoftware.DNN.Modules.Alugen3NET

  Public MustInherit Class SettingsClient
    Inherits Entities.Modules.ModuleSettingsBase

#Region "Controls"
#End Region

    Public Overrides Sub LoadSettings()
      Try
        If Not Page.IsPostBack Then
          Dim objModuleController As New ModuleController
          'Dim xxx As ModuleInfo
          'xxx.g.ModuleID()
          cboAlugenMainModule.DataSource = objModuleController.GetModulesByDefinition(PortalId, "Alugen3NET")
          cboAlugenMainModule.DataTextField = "ModuleTitle"
          cboAlugenMainModule.DataValueField = "ModuleID"
          cboAlugenMainModule.DataBind()
          'no associate module
          cboAlugenMainModule.Items.Insert(0, New ListItem("(none)", "0"))

          'Load settings from TabModuleSettings: specific to this instance
          Dim AlugenMainModule As String = CType(ModuleSettings("AlugenMainModule"), String)
          cboAlugenMainModule.SelectedValue = AlugenMainModule

          Dim HideCredits As Boolean = CType(ModuleSettings("HideCredits"), Boolean)
          chkHideCredits.Checked = HideCredits

          Dim EmailNotification As Boolean = CType(ModuleSettings("EmailNotification"), Boolean)
          chkSendEmailNotification.Checked = EmailNotification

        End If
      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

    Public Overrides Sub UpdateSettings()
      Try
        Dim objModules As New Entities.Modules.ModuleController

        ' Update TabModuleSettings
        objModules.UpdateModuleSetting(ModuleId, "AlugenMainModule", cboAlugenMainModule.SelectedValue.ToString())
        objModules.UpdateModuleSetting(ModuleId, "HideCredits", chkHideCredits.Checked.ToString())
        objModules.UpdateModuleSetting(ModuleId, "EmailNotification", chkSendEmailNotification.Checked.ToString())

      Catch exc As Exception
        ProcessModuleLoadException(Me, exc)
      End Try
    End Sub

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents chkHideCredits As System.Web.UI.WebControls.CheckBox
    Protected WithEvents cboAlugenMainModule As System.Web.UI.WebControls.DropDownList
    Protected WithEvents chkSendEmailNotification As System.Web.UI.WebControls.CheckBox

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
