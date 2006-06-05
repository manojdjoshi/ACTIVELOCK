<%@ Control Language="vb" AutoEventWireup="false" Codebehind="Settings.ascx.vb" Inherits="FriendSoftware.DNN.Modules.Alugen3NET.Settings" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<table cellPadding="0" width="100%" border="0">
	<tr vAlign="top">
		<td vAlign="top" align="left" width="175" class="SubHead">
			<dnn:label id="plProductsPriceDecimals" suffix=":" controlname="cboProductsPriceDecimals" runat="server"></dnn:label>
		</td>
		<td vAlign="top" class="Normal">
			<asp:dropdownlist id="cboProductsPriceDecimals" runat="server" Height="100px">
				<asp:ListItem Value="0">0</asp:ListItem>
				<asp:ListItem Value="1">1</asp:ListItem>
				<asp:ListItem Value="2">2</asp:ListItem>
				<asp:ListItem Value="3">3</asp:ListItem>
				<asp:ListItem Value="4">4</asp:ListItem>
				<asp:ListItem Value="5">5</asp:ListItem>
				<asp:ListItem Value="6">6</asp:ListItem>
			</asp:dropdownlist>&nbsp;
			<asp:Label id="plProductsPriceDecimals2" runat="server" resourcekey="plProductsPriceDecimals2"></asp:Label></td>
	</tr>
	<tr>
		<td vAlign="top" align="left" width="175" class="SubHead">
			<dnn:label id="plEnableFiltersByDefault" suffix=":" controlname="chkEnableFiltersByDefault"
				runat="server"></dnn:label>
		</td>
		<td vAlign="top" class="Normal">
			<asp:CheckBox id="chkEnableFiltersByDefault" controlname="chkEnableFiltersByDefault" resourcekey="chkEnableFiltersByDefault"
				runat="server" Width="100%"></asp:CheckBox>
		</td>
	</tr>
	<tr vAlign="top">
		<td vAlign="top" align="left" width="175" class="SubHead">
			<dnn:label id="plAdminEmail" suffix=":" controlname="txtAdminEmail" runat="server"></dnn:label>
		</td>
		<td vAlign="top" class="Normal">
			<asp:TextBox id="txtAdminEmail" runat="server" Width="100%"></asp:TextBox>
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
