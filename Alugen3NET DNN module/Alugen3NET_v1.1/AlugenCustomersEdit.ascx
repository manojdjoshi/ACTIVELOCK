<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Control Language="vb" AutoEventWireup="false" Codebehind="AlugenCustomersEdit.ascx.vb" Inherits="FriendSoftware.DNN.Modules.Alugen3NET.AlugenCustomersEdit" %>
<table cellSpacing="0" cellPadding="0" width="100%" border="0">
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plCustomerName" suffix=":" controlname="txtCustomerName" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top"><asp:textbox id="txtCustomerName" controlname="txtCustomerName" runat="server" Width="100%"></asp:textbox></td>
	</tr>
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plCustomerAddress" suffix=":" controlname="txtCustomerAddress" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top"><asp:textbox id="txtCustomerAddress" controlname="txtCustomerAddress" runat="server" Width="100%"
				Rows="4" TextMode="MultiLine"></asp:textbox></td>
	</tr>
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plCustomerContactPerson" suffix=":" controlname="txtCustomerContactPerson" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top"><asp:textbox id="txtCustomerContactPerson" controlname="txtCustomerContactPerson" runat="server"
				Width="100%"></asp:textbox></td>
	</tr>
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plCustomerPhone" suffix=":" controlname="txtCustomerPhone" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top"><asp:textbox id="txtCustomerPhone" controlname="txtCustomerPhone" runat="server" Width="100%"></asp:textbox></td>
	</tr>
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plAssociatedUser" suffix=":" controlname="cboAssociatedUser" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top">
			<asp:dropdownlist id="cboAssociatedUser" controlname="cboAssociatedUser" runat="server"></asp:dropdownlist></td>
	</tr>
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plUseUserEmail" suffix=":" controlname="chkUseUserEmail" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top" align="left"><asp:checkbox id="chkUseUserEmail" runat="server" Text="Should we use associated user email address?"
				resourcekey="chkUseUserEmail"></asp:checkbox></td>
	</tr>
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plCustomerEmail" suffix=":" controlname="txtCustomerEmail" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top"><asp:textbox id="txtCustomerEmail" controlname="txtCustomerEmail" runat="server" Width="100%"></asp:textbox></td>
	</tr>
</table>
<asp:linkbutton id="cmdUpdate" runat="server" resourcekey="cmdUpdate" BorderStyle="None" CssClass="CommandButton">Update</asp:linkbutton>
&nbsp;
<asp:linkbutton id="cmdCancel" runat="server" resourcekey="cmdCancel" BorderStyle="None" CssClass="CommandButton"
	CausesValidation="False">Cancel</asp:linkbutton>
&nbsp;
<asp:linkbutton id="cmdDelete" runat="server" resourcekey="cmdDelete" BorderStyle="None" CssClass="CommandButton"
	CausesValidation="False" Visible="False">Delete</asp:linkbutton>
<br>
<br>
<asp:panel class="Normal" id="pnlCredits" runat="server">
	<TABLE id="Table3" cellSpacing="0" cellPadding="0" width="100%" border="0">
		<TR vAlign="top">
			<TD class="normal" vAlign="top" align="left">Created by <A href="http://www.friendsoftware.ro">
					Friend Software</A><BR>
				copyright © 2006<BR>
				<A href="http://www.friendsoftware.ro"><IMG id="imgFriendSoftware" alt="Friend Software" src="http://www.friendsoftware.ro/images/friendsoftware_logo.JPG"
						align="absMiddle" border="0" runat="server"></A>
			</TD>
			<TD class="normal" vAlign="top" align="right">Using <A href="http://www.activelocksoftware.com">
					Activelock Software Group</A> software protection technology<BR>
				copyright © 2001-2006<BR>
				<A href="http://www.activelocksoftware.com"><IMG id="imgActiveLock" alt="ActiveLockSoftware" src="http://www.friendsoftware.ro/images/activelock_logo.gif"
						align="absMiddle" border="0" runat="server"> </A>
			</TD>
		</TR>
	</TABLE>
</asp:panel>
