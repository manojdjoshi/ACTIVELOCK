<%@ Control Language="vb" AutoEventWireup="false" Codebehind="AlugenCodesClientEdit.ascx.vb" Inherits="FriendSoftware.DNN.Modules.Alugen3NET.AlugenCodesClientEdit" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<asp:panel class="Normal" id="pnlMessage" runat="server" Width="100%" Visible="False">
	<CENTER>
		<asp:Label id="lblMessage" runat="server"></asp:Label></CENTER>
</asp:panel><asp:panel class="Normal" id="pnlClientGeneration" runat="server" Width="100%">
	<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
		<TR>
			<TD class="SubHead" vAlign="top" width="125">
				<dnn:label id="plProduct" runat="server" controlname="txtProduct" suffix=":"></dnn:label></TD>
			<TD class="Normal" vAlign="top" width="200">
				<asp:textbox id="txtProduct" runat="server" controlname="txtProduct" ForeColor="#424242" BackColor="#DBDBDB"
					ReadOnly="True" Width="100%"></asp:textbox></TD>
			<TD width="25"></TD>
			<TD class="SubHead" vAlign="top" width="125">
				<dnn:label id="plCustomer" runat="server" controlname="txtCustomer" suffix=":"></dnn:label></TD>
			<TD class="Normal" vAlign="top" width="200">
				<asp:textbox id="txtCustomer" runat="server" controlname="txtCustomer" ForeColor="#424242" BackColor="#DBDBDB"
					ReadOnly="True" Width="100%"></asp:textbox></TD>
		<TR>
			<TD class="SubHead" vAlign="top" width="125">
				<dnn:label id="plLicenseType" runat="server" controlname="txtLicenseType" suffix=":"></dnn:label></TD>
			<TD class="Normal" vAlign="top" width="200">
				<asp:textbox id="txtLicenseType" runat="server" controlname="txtLicenseType" ForeColor="#424242"
					BackColor="#DBDBDB" ReadOnly="True" Width="100%"></asp:textbox>
			<TD width="25"></TD>
			<TD class="SubHead" vAlign="top" width="125">
				<TABLE cellSpacing="border=0" cellPadding="0" width="100%">
					<TR>
						<TD class="SubHead" vAlign="top">
							<asp:Panel class="SubHead" id="pnlExpireDate" runat="server" Width="100%" Height="100%" name="pnlExpireDate">
								<dnn:label id="plExpireDate" runat="server" controlname="txtExpireDate" suffix=":"></dnn:label>
							</asp:Panel></TD>
						<TD class="SubHead" vAlign="top">
							<asp:Panel class="SubHead" id="pnlPeriodicDays" runat="server" Width="100%" Height="100%" name="pnlPeriodicDays">
								<dnn:label id="plPeriodicDays" runat="server" controlname="txtExpireDate"></dnn:label>
							</asp:Panel></TD>
					</TR>
				</TABLE>
			</TD>
			<TD class="Normal" vAlign="top" width="200">
				<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD vAlign="top" width="100%">
							<asp:textbox id="txtExpireDate" runat="server" controlname="txtExpireDate" ForeColor="#424242"
								BackColor="#DBDBDB" ReadOnly="True" Width="100%"></asp:textbox></TD>
					</TR>
				</TABLE>
			</TD>
		</TR>
		<TR>
			<TD class="SubHead" vAlign="top" width="125">
				<dnn:label id="plRegisteredLevel" runat="server" controlname="txtRegisteredLevel" suffix=":"></dnn:label></TD>
			<TD class="Normal" vAlign="top" width="200">
				<asp:textbox id="txtRegisteredLevel" runat="server" controlname="txtRegisteredLevel" ForeColor="#424242"
					BackColor="#DBDBDB" ReadOnly="True" Width="100%"></asp:textbox></TD>
			<TD width="25"></TD>
			<TD class="SubHead" vAlign="top" width="125"></TD>
			<TD class="Normal" vAlign="top" width="200"></TD>
		</TR>
	</TABLE>
	<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
		<TR>
			<TD class="SubHead" vAlign="top" width="125">
				<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
					<TR>
						<TD class="SubHead" vAlign="top" width="125">
							<dnn:label id="plInstalationCode" runat="server" controlname="txtInstalationCode" suffix=":"></dnn:label></TD>
					</TR>
					<TR>
						<TD class="SubHead" vAlign="top" align="right"><IMG id="imgPasteInstallCode" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
								height="16" alt="Paste Install Code" width="16" align="right" border="1" name="imgPasteInstallCode" runat="server">
						</TD>
					</TR>
				</TABLE>
			</TD>
			<TD class="Normal" vAlign="top" width="550">
				<asp:textbox id="txtInstalationCode" runat="server" controlname="txtInstalationCode" Width="550"
					Rows="4" TextMode="MultiLine"></asp:textbox></TD>
		</TR>
		<TR>
			<TD class="SubHead" vAlign="top" width="125">
				<dnn:label id="plUserName" runat="server" controlname="txtUserName" suffix=":"></dnn:label></TD>
			<TD class="Normal" vAlign="top" width="550">
				<asp:textbox id="txtUserName" runat="server" controlname="txtUserName" ForeColor="#424242" BackColor="#DBDBDB"
					ReadOnly="True" Width="465"></asp:textbox><INPUT id="btnGenerate" type="button" value="Generate" name="btnGenerate" runat="server">
			</TD>
		</TR>
		<TR>
			<TD class="SubHead" vAlign="top" width="125">
				<TABLE cellSpacing="0" cellPadding="0" width="125" border="0">
					<TR>
						<TD class="SubHead" vAlign="top">
							<dnn:label id="plActivationCode" runat="server" controlname="txtActivationCode" suffix=":"></dnn:label></TD>
					</TR>
					<TR>
						<TD class="SubHead" vAlign="top" align="right"><IMG id="imgCopyActivationCode" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
								height="16" alt="Copy Activation Code" width="16" border="1" name="imgCopyActivationCode" runat="server">
						</TD>
					</TR>
					<TR>
						<TD class="SubHead" vAlign="top" align="right"><IMG id="imgEmail" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
								height="16" alt="Email Activation Code" width="16" border="1" name="imgEmail" runat="server">
						</TD>
					</TR>
				</TABLE>
			</TD>
			<TD class="Normal" vAlign="top" width="550">
				<asp:textbox id="txtActivationCode" runat="server" controlname="txtActivationCode" ForeColor="#424242"
					BackColor="#DBDBDB" ReadOnly="True" Width="550" Rows="6" TextMode="MultiLine"></asp:textbox></TD>
		</TR>
	</TABLE>
</asp:panel><br>
<asp:linkbutton id="cmdCancel" runat="server" CausesValidation="False" resourcekey="cmdCancel" BorderStyle="None"
	CssClass="CommandButton">Cancel</asp:linkbutton><br>
<br>
<asp:panel class="Normal" id="pnlCredits" runat="server" Width="100%">
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
</asp:panel><INPUT id="idProduct" type="hidden" name="idProduct" runat="server">
<INPUT id="idCustomer" type="hidden" name="idCustomer" runat="server">
