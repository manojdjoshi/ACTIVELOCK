<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Control Language="vb" AutoEventWireup="false" Codebehind="AlugenCodesEdit.ascx.vb" Inherits="FriendSoftware.DNN.Modules.Alugen3NET.AlugenCodesEdit" %>
<table cellSpacing="0" cellPadding="0" width="100%" border="0">
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plProduct" suffix=":" controlname="cboProduct" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top" width="200"><asp:dropdownlist id="cboProduct" controlname="cboProduct" runat="server" Width="100%"></asp:dropdownlist></td>
		<td width="25"></td>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plCustomer" suffix=":" controlname="cboCustomer" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top" width="200"><asp:dropdownlist id="cboCustomer" controlname="cboCustomer" runat="server" Width="100%"></asp:dropdownlist></td>
	</tr>
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plLicenseType" suffix=":" controlname="cboLicenseType" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top" width="200"><asp:dropdownlist id="cboLicenseType" controlname="cboLicenseType" runat="server" Width="100%">
				<asp:ListItem Value="0">Time locked</asp:ListItem>
				<asp:ListItem Value="1">Periodic</asp:ListItem>
				<asp:ListItem Value="2">Permanent</asp:ListItem>
			</asp:dropdownlist></td>
		<td width="25"></td>
		<td class="SubHead" vAlign="top" width="125">
			<table width="100%" cellpadding="0" cellspacing="border=0">
				<tr>
					<td valign="top" class="SubHead">
						<asp:Panel class="SubHead" id="pnlExpireDate" name="pnlExpireDate" runat="server" Width="100%"
							Height="100%">
							<dnn:label id="plExpireDate" runat="server" controlname="txtExpireDate" suffix=":"></dnn:label>
						</asp:Panel>
					</td>
					<td valign="top" class="SubHead">
						<asp:Panel class="SubHead" id="pnlPeriodicDays" name="pnlPeriodicDays" runat="server" Width="100%"
							Height="100%">
							<dnn:label id="plPeriodicDays" runat="server" controlname="txtExpireDate"></dnn:label>
						</asp:Panel>
					</td>
				</tr>
			</table>
		</td>
		<td class="Normal" vAlign="top" width="200">
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr>
					<td valign="top" Width="100%">
						<asp:textbox id="txtExpireDate" controlname="txtExpireDate" runat="server" Width="100%"></asp:textbox>
					</td>
					<td valign="top">
						<IMG id="imgExpireDateCalendar" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: pointer; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
							height="16" alt="Expire Date" width="16" border="0" name="imgExpireDateCalendar" runat="server">
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plRegisteredLevel" suffix=":" controlname="cboRegisteredLevel" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top" width="200"><asp:dropdownlist id="cboRegisteredLevel" controlname="cboRegisteredLevel" runat="server" Width="100%">
				<asp:ListItem Value="0">Level 0</asp:ListItem>
				<asp:ListItem Value="1">Level 1</asp:ListItem>
				<asp:ListItem Value="2">Level 2</asp:ListItem>
				<asp:ListItem Value="3">Level 3</asp:ListItem>
				<asp:ListItem Value="4">Level 4</asp:ListItem>
				<asp:ListItem Value="5">Level 5</asp:ListItem>
				<asp:ListItem Value="6">Level 6</asp:ListItem>
				<asp:ListItem Value="7">Level 7</asp:ListItem>
				<asp:ListItem Value="8">Level 8</asp:ListItem>
				<asp:ListItem Value="9">Level 9</asp:ListItem>
				<asp:ListItem Value="10">Level 10</asp:ListItem>
				<asp:ListItem Value="11">Level 11</asp:ListItem>
				<asp:ListItem Value="12">Level 12</asp:ListItem>
				<asp:ListItem Value="13">Level 13</asp:ListItem>
				<asp:ListItem Value="14">Level 14</asp:ListItem>
				<asp:ListItem Value="15">Level 15</asp:ListItem>
				<asp:ListItem Value="16">Level 16</asp:ListItem>
				<asp:ListItem Value="17">Level 17</asp:ListItem>
				<asp:ListItem Value="18">Level 18</asp:ListItem>
				<asp:ListItem Value="19">Level 19</asp:ListItem>
				<asp:ListItem Value="20">Level 20</asp:ListItem>
				<asp:ListItem Value="21">Level 21</asp:ListItem>
				<asp:ListItem Value="22">Level 22</asp:ListItem>
				<asp:ListItem Value="23">Level 23</asp:ListItem>
			</asp:dropdownlist></td>
		<td width="25"></td>
		<td class="SubHead" vAlign="top" width="125">
			<dnn:label id="plRegisterByCustomer" suffix=":" controlname="chkRegisterByCustomer" runat="server"></dnn:label>
		</td>
		<td class="Normal" vAlign="top" width="200">
			<asp:CheckBox id="chkRegisterByCustomer" runat="server" name="chkRegisterByCustomer" resourcekey="chkRegisterByCustomer.Text"></asp:CheckBox>
		</td>
	</tr>
</table>
<asp:Panel id="pnlGenByCustomer" controlname="pnlGenByCustomer" runat="server">
<table cellSpacing="0" cellPadding="0" width="100%" border="0">
	<tr>
		<td class="SubHead" vAlign="top" width="125">
			<table width="100%" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td class="SubHead" vAlign="top" width="125">
						<dnn:label id="plInstalationCode" suffix=":" controlname="txtInstalationCode" runat="server"></dnn:label>
					</td>
				</tr>
				<tr>
					<td valign="top" class="SubHead" align="right">
						<IMG id="imgPasteInstallCode" align="right" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: pointer; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
							height="16" alt="Paste Install Code" width="16" border="1" name="imgPasteInstallCode"
							runat="server">
					</td>
				</tr>
			</table>
		</td>
		<td class="Normal" vAlign="top" width="550"><asp:textbox id="txtInstalationCode" controlname="txtInstalationCode" runat="server" Width="550"
				Rows="4" TextMode="MultiLine"></asp:textbox>
		</td>
	</tr>
	<tr>
		<td class="SubHead" vAlign="top" width="125"><dnn:label id="plUserName" suffix=":" controlname="txtUserName" runat="server"></dnn:label></td>
		<td class="Normal" vAlign="top" width="550"><asp:textbox id="txtUserName" controlname="txtUserName" runat="server" Width="465" ReadOnly="True"
				ForeColor="#424242" BackColor="#DBDBDB"></asp:textbox>
			<INPUT id="btnGenerate" type="button" value="Generate" name="btnGenerate" runat="server">
			<BR>
			<asp:CheckBox id="chkDirectCodeEdit" runat="server" name="chkDirectCodeEdit" Text="Enable direct code edit"
				resourcekey="chkDirectCodeEdit.Text"></asp:CheckBox></td>
	</tr>
	<tr>
		<td class="SubHead" vAlign="top" width="125">
			<table width="125" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td valign="top" class="SubHead">
						<dnn:label id="plActivationCode" suffix=":" controlname="txtActivationCode" runat="server"></dnn:label>
					</td>
				</tr>
				<tr>
					<td valign="top" class="SubHead" align="right">
						<IMG id="imgCopyActivationCode" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: pointer; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
							height="16" alt="Copy Activation Code" width="16" border="1" name="imgCopyActivationCode"
							runat="server">
					</td>
				</tr>
				<tr>
					<td valign="top" class="SubHead" align="right">
						<IMG id="imgEmail" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: pointer; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
							height="16" alt="Email Activation Code" width="16" border="1" name="imgEmail" runat="server">
					</td>
				</tr>
			</table>
		</td>
		<td class="Normal" vAlign="top" width="550">
			<asp:textbox id="txtActivationCode" controlname="txtActivationCode" runat="server" Width="550"
				ReadOnly="True" TextMode="MultiLine" Rows="6" ForeColor="#424242" BackColor="#DBDBDB"></asp:textbox>
		</td>
	</tr>
</table>
</asp:Panel>
<table cellSpacing="0" cellPadding="0" width="100%" border="0">
	<tr>
		<td class="SubHead" vAlign="top" width="125">
			<dnn:label id="plSendEmail" suffix=":" controlname="chkSendEmailNotification" runat="server"></dnn:label>
		</td>
		<td class="Normal" valign="top" width="550">
			<asp:CheckBox id="chkSendEmailNotification" resourcekey="chkSendEmailNotification" runat="server"
				Text="Send Email Notification ?"></asp:CheckBox></td>
	</tr>
</table>
<br>
<asp:linkbutton id="cmdUpdate" runat="server" CssClass="CommandButton" BorderStyle="None" resourcekey="cmdUpdate">Update</asp:linkbutton>
<asp:linkbutton id="cmdCancel" runat="server" CssClass="CommandButton" BorderStyle="None" resourcekey="cmdCancel"
	CausesValidation="False">Cancel</asp:linkbutton>
<asp:linkbutton id="cmdDelete" runat="server" CssClass="CommandButton" BorderStyle="None" resourcekey="cmdDelete"
	CausesValidation="False" Visible="False">Delete</asp:linkbutton>
<br>
<br>
<asp:Panel class="Normal" id="pnlCredits" runat="server" Width="100%">
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
