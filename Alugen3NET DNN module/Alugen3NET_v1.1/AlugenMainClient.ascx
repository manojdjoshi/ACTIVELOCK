<%@ Control Language="vb" AutoEventWireup="false" Codebehind="AlugenMainClient.ascx.vb" Inherits="FriendSoftware.DNN.Modules.Alugen3NET.AlugenMainClient" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<asp:panel class="Normal" id="pnlMessage" runat="server">
	<TABLE class="tablemessage" id="TableMessage" cellSpacing="0" cellPadding="10" width="100%">
		<TR>
			<TD class="Normal">
				<asp:Label id="lblMessage" runat="server"></asp:Label></TD>
		</TR>
	</TABLE>
</asp:panel>
<table id="Table1" cellSpacing="0" cellPadding="0" width="100%" border="0">
	<tr>
		<td class="tdspace" vAlign="middle" align="left"><asp:linkbutton id="lnkRefresh" runat="server">
				<div style="text-decoration: none;">
					<img id="imgRefresh" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
						height="16" alt="Refresh tables" width="16" border="0" name="imgRefresh" runat="server"></img>
					&nbsp;Refresh</div>
			</asp:linkbutton></td>
		<td class="multipageactive" id="tdbtnCodes" onmouseover="if(this.className!='multipageactive'){this.className='multipagehover'}"
			onmouseout="if(this.className!='multipageactive'){this.className='multipage'}" width="15%"
			runat="server" name="tdbtnCodes">Codes</td>
	</tr>
</table>
<table class="bodytable" id="Table2" cellSpacing="0" cellPadding="1" width="100%">
	<tr>
		<td><asp:panel class="Normal" id="pnlCodes" runat="server" name="pnlCodes" Width="100%" Height="100%"><INPUT id="chkFilterCodes" type="checkbox" name="chkFilterCodes" runat="server">Enable Filter<BR>
      <TABLE class="sortable filtered" id="codestable" cellSpacing="0" cellPadding="0" width="100%"
					border="1" name="codestable">
					<TR>
						<td class="Subhead thbody" vAlign="top" align="left" width="5%">Generated</td>
						<TD class="Subhead thbody" vAlign="top" align="left" width="20%">Product</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="10%">Version</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="27%">Customer</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="10%">User Name</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="15%">License<br>type</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="12%">Generation<br>Date</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="12%">Expiration<br>Date</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="5%">Level</TD>
					</TR>
					<asp:Repeater id="rptAlugenCodes" runat="server">
						<ItemTemplate>
							<tr>
								<td align="center" valign="middle" class="normal tdbody" width="0%">
									<asp:HyperLink NavigateUrl='<%# EditURL("CodeID",DataBinder.Eval(Container.DataItem,"CodeID"), "EditCodeClient", "") %>' Visible='<%# DataBinder.Eval(Container, "DataItem.GenByCustomer") %>' runat="server" ID="Hyperlink1" BorderWidth=0>
									Generate Now
									</asp:HyperLink>
									<asp:HyperLink NavigateUrl='<%# EditURL("CodeID",DataBinder.Eval(Container.DataItem,"CodeID"), "EditCodeClient", "") %>' Visible='<%# Not DataBinder.Eval(Container, "DataItem.GenByCustomer") %>' runat="server" ID="Hyperlink2" BorderWidth=0>
									View
									</asp:HyperLink>
								</td>
								<td align="left" valign="middle" class="normal tdbody">
									<%# DataBinder.Eval(Container, "DataItem.ProductName") %>
								</td>
								<td align="left" valign="middle" class="Normal tdbody"><%# DataBinder.Eval(Container, "DataItem.ProductVersion") %></td>
								<td align="left" valign="middle" class="Normal tdbody">
									<%# DataBinder.Eval(Container, "DataItem.CustomerName") %>
								</td>
								<td align="left" valign="middle" class="Normal tdbody">
									<%# DataBinder.Eval(Container, "DataItem.CodeUserName") %>
								</td>
								<td align="left" valign="middle" class="Normal tdbody">
									<%# GetLicenseType(DataBinder.Eval(Container, "DataItem.CodeLicenseType")) %>
								</td>
								<td align="center" valign="middle" class="Normal tdbody">
									<%# Format(DataBinder.Eval(Container, "DataItem.CreatedDate"), "d") %>
								</td>
								<td align="center" valign="middle" class="Normal tdbody">
									<%# Format(DataBinder.Eval(Container, "DataItem.CodeExpireDate"), "d") %>
								</td>
								<td align="center" valign="middle" class="Normal tdbody">
									<%# DataBinder.Eval(Container, "DataItem.CodeRegisteredLevel") %>
								</td>
							</tr>
						</ItemTemplate>
					</asp:Repeater></TABLE></asp:panel></td>
	</tr>
</table>
<asp:panel class="Normal" id="pnlCredits" runat="server">
	<br>
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
