<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Control Language="vb" AutoEventWireup="false" Codebehind="SettingsClient.ascx.vb" Inherits="FriendSoftware.DNN.Modules.Alugen3NET.SettingsClient" %>
<table cellPadding="0" width="100%" border="0">
	<tr vAlign="top">
		<td vAlign="top" align="left" width="175" class="SubHead">
			<dnn:label id="plAlugenMainModule" suffix=":" controlname="cboAlugenMainModule" runat="server"></dnn:label>
		</td>
		<td vAlign="top" class="Normal">
			<asp:dropdownlist id="cboAlugenMainModule" runat="server" Height="200px"></asp:dropdownlist>&nbsp;</td>
	</tr>
	<tr vAlign="top">
		<td vAlign="top" align="left" width="175" class="SubHead">
			<dnn:label id="plSendEmailNotification" suffix=":" controlname="chkSendEmailNotification" runat="server"></dnn:label>
		</td>
		<td vAlign="top" class="Normal">
			<asp:CheckBox id="chkSendEmailNotification" controlname="chkSendEmailNotification" resourcekey="chkSendEmailNotification.Text"
				runat="server" Width="100%"></asp:CheckBox>
		</td>
	</tr>
	<tr vAlign="top">
		<td vAlign="top" align="left" width="175" class="SubHead">
			<dnn:label id="plHideCredits" suffix=":" controlname="chkHideCredits" runat="server"></dnn:label>
		</td>
		<td vAlign="top" class="Normal">
			<asp:CheckBox id="chkHideCredits" controlname="chkHideCredits" resourcekey="chkHideCredits" runat="server"
				Width="100%"></asp:CheckBox>
		</td>
	</tr>
</table>
