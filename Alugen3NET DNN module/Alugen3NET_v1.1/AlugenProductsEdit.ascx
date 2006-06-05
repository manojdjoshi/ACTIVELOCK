<%@ Control Language="vb" AutoEventWireup="false" Codebehind="AlugenProductsEdit.ascx.vb" Inherits="FriendSoftware.DNN.Modules.Alugen3NET.AlugenProductsEdit" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<table border="0" cellspacing="0" cellpadding="0" width="100%">
	<tr>
		<td class="SubHead" width="125">
			<dnn:label id="plProductName" runat="server" controlname="txtProductName" suffix=":"></dnn:label>
		</td>
		<td>
			<asp:TextBox id="txtProductName" controlname="txtProductName" runat="server" Width="100%"></asp:TextBox>
		</td>
	</tr>
	<TR>
		<td class="SubHead" width="125">
			<dnn:label id="plProductVersion" runat="server" controlname="txtProductVersion" suffix=":"></dnn:label>
		</td>
		<td>
			<asp:TextBox id="txtProductVersion" controlname="txtProductVersion" runat="server" Width="100%"></asp:TextBox>
		</td>
	</TR>
	<tr>
		<td class="SubHead" width="125"></td>
		<td valign="top"><INPUT id="btnGenerateNew" type="button" value="Generate new codes" name="btnGenerateNew"
				runat="server">&nbsp;<INPUT id="btnClearCodes" type="button" value="Clear codes" name="btnClearCodes" runat="server">&nbsp;<INPUT id="btnValidate" type="button" value="Validate" name="btnValidate" runat="server">&nbsp;</td>
	</tr>
	<tr>
		<td class="SubHead" width="125" height="20"></td>
		<td valign="top" height="20" class="Normal">
			<asp:CheckBox id="chkDirectCodeEdit" name="chkDirectCodeEdit" resourcekey="chkDirectCodeEdit.Text" runat="server" Text="Enable direct code edit"></asp:CheckBox>
		</td>
	</tr>
	<TR valign="top">
		<td class="SubHead" width="125" valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td valign="top" class="SubHead">
						<dnn:label id="plProductVCode" runat="server" controlname="txtProductVCode" suffix=":"></dnn:label>
					</td>
					<td align="right" valign="top">
						<IMG id="imgCopyVCode" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: pointer; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
							height="16" alt="Copy VCode" width="16" border="1" name="imgCopyVCode" runat="server">
					</td>
				</tr>
			</table>
		</td>
		<td valign="top">
			<asp:TextBox id="txtProductVCode" runat="server" TextMode="MultiLine" Rows="3" ReadOnly="True"
				ForeColor="#424242" BackColor="#DBDBDB" Width="600"></asp:TextBox>
		</td>
	</TR>
	<TR valign="top">
		<td class="SubHead" width="125" valign="top">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td class="SubHead" valign="top">
						<dnn:label id="plProductGCode" runat="server" controlname="txtProductGCode" suffix=":"></dnn:label>
					</td>
					<td align="right" valign="top">
						<IMG id="imgCopyGCode" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: pointer; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
							height="16" alt="Copy GCode" width="16" border="1" name="imgCopyGCode" runat="server">
					</td>
				</tr>
			</table>
		</td>
		<td valign="top">
			<asp:TextBox id="txtProductGCode" controlname="txtProductGCode" runat="server" Width="600px"
				TextMode="MultiLine" Rows="3" ReadOnly="True" ForeColor="#424242" BackColor="#DBDBDB"></asp:TextBox>
		</td>
	</TR>
	<TR>
		<td class="SubHead" width="125">
			<dnn:label id="plProductPrice" runat="server" controlname="txtProductPrice" suffix=":"></dnn:label>
		</td>
		<td>
			<asp:TextBox id="txtProductPrice" controlname="txtProductPrice" runat="server" Width="100%"></asp:TextBox>
		</td>
	</TR>
</table>
<asp:linkbutton id="cmdUpdate" CssClass="CommandButton" runat="server" BorderStyle="None" resourcekey="cmdUpdate">Update</asp:linkbutton>
<asp:linkbutton id="cmdCancel" CssClass="CommandButton" runat="server" CausesValidation="False"
	BorderStyle="None" resourcekey="cmdCancel">Cancel</asp:linkbutton>
<asp:linkbutton id="cmdDelete" CssClass="CommandButton" runat="server" CausesValidation="False"
	BorderStyle="None" Visible="False" resourcekey="cmdDelete">Delete</asp:linkbutton>
<br>
<br>
<asp:Panel class="Normal" id="pnlCredits" runat="server">
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
</asp:Panel>
