<%@ Control Language="vb" AutoEventWireup="false" Codebehind="AlugenMain.ascx.vb" Inherits="FriendSoftware.DNN.Modules.Alugen3NET.AlugenMain" %>
<%@ Register TagPrefix="dnn" TagName="TextEditor" Src="~/controls/TextEditor.ascx"%>
<%@ Register TagPrefix="Portal" TagName="Audit" Src="~/controls/ModuleAuditControl.ascx" %>
<%@ Register TagPrefix="dnn" TagName="Label" Src="~/controls/LabelControl.ascx" %>
<asp:Panel class="Normal" id="pnlMessage" runat="server">
	<TABLE class="tablemessage" id="TableMessage" cellSpacing="0" cellPadding="10" width="100%">
		<TR>
			<TD class="Normal">
				<asp:Label id="lblMessage" runat="server"></asp:Label></TD>
		</TR>
	</TABLE>
</asp:Panel>
<table width="100%" cellpadding="0" cellspacing="0" border="0" id="Table1">
	<tr>
		<td class="tdspace" align="left" valign="middle"><asp:linkbutton id="lnkRefresh" runat="server">
				<div style="text-decoration: none;">
					<img id="imgRefresh" style="BORDER-LEFT-COLOR: #666665; BORDER-BOTTOM-COLOR: #666665; CURSOR: pointer; CURSOR: hand; BORDER-TOP-COLOR: #666665; BORDER-RIGHT-COLOR: #666665"
						height="16" alt="Refresh tables" width="16" border="0" name="imgRefresh" runat="server"></img>
					&nbsp;Refresh</div>
			</asp:linkbutton></td>
		<td width="15%" id="tdbtnProducts" name="tdbtnProducts" onmouseover="if(this.className!='multipageactive'){this.className='multipagehover'}"
			onmouseout="if(this.className!='multipageactive'){this.className='multipage'}" runat="server"
			class="NormalBold">
			<asp:Label id="lblProducts" runat="server">Products</asp:Label></td>
		<td class="tdspace" width="1">&nbsp;</td>
		<td width="15%" id="tdbtnCustomers" name="tdbtnCustomers" onmouseover="if(this.className!='multipageactive'){this.className='multipagehover'}"
			onmouseout="if(this.className!='multipageactive'){this.className='multipage'}" runat="server">Customers</td>
		<td class="tdspace" width="1">&nbsp;</td>
		<td width="15%" id="tdbtnCodes" name="tdbtnCodes" onmouseover="if(this.className!='multipageactive'){this.className='multipagehover'}"
			onmouseout="if(this.className!='multipageactive'){this.className='multipage'}" runat="server">Codes</td>
	</tr>
</table>
<table width="100%" cellpadding="1" cellspacing="0" class="bodytable" id="Table2">
	<tr>
		<td><asp:panel class="Normal" id="pnlProducts" runat="server" Height="100%" Width="100%" name="pnlProducts"><INPUT id="chkFilterProducts" type="checkbox" name="chkFilterProducts" runat="server">Enable Filter<BR>
      <TABLE class="sortable filtered" id="productstable" cellSpacing="0" cellPadding="0" width="100%"
					border="1" name="productstable">
					<TR class="noFilter">
						<TD class="Subhead thbody" vAlign="top" align="left" width="0%"></TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="65%">Product Name
						</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="20%">Product Version
						</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="15%">Product Price
						</TD>
					</TR>
					<asp:Repeater id="rptAlugenProducts" runat="server">
						<ItemTemplate>
							<tr>
								<td align="left" valign="top" class="normal tdbody">
									<asp:HyperLink NavigateUrl='<%# EditURL("ProductID",DataBinder.Eval(Container.DataItem,"ProductID"), "EditProduct", "") %>' runat="server" ID="Hyperlink1" BorderWidth=0>
										<asp:Image ID="Hyperlink1Image" Runat="server" ImageUrl="~/images/edit.gif" AlternateText="Edit" Visible="<%#IsEditable%>" resourcekey="Edit" />
									</asp:HyperLink>
								</td>
								<td align="left" valign="top" class="normal tdbody">
									<%# DataBinder.Eval(Container, "DataItem.ProductName") %>
								</td>
								<td align="left" valign="top" class="Normal tdbody">
									<%# DataBinder.Eval(Container, "DataItem.ProductVersion") %>
								</td>
								<td align="right" valign="top" class="Normal tdbody">
									<%# DataBinder.Eval(Container, "DataItem.ProductPrice", MyDecFormat()) %>
								</td>
							</tr>
						</ItemTemplate>
					</asp:Repeater></TABLE><BR>
<asp:Label id="lblImportProducts" runat="server" resourcekey="lblImportProducts.Text"></asp:Label>&nbsp;<INPUT id="iniProductFile" type="file" accept="text/*" size="30" name="iniProductFile"
					runat="server" Width="300"> 
<asp:LinkButton id="cmdImportProducts" runat="server" resourcekey="cmdImportProducts.Text"></asp:LinkButton>
			</asp:panel><asp:panel class="Normal" id="pnlCustomers" runat="server" Height="100%" Width="100%" name="pnlCustomers">
				<TABLE class="sortable filtered" id="customerstable" cellSpacing="0" cellPadding="0" width="100%"
					border="1" name="customerstable">
					<INPUT id="chkFilterCustomers" type="checkbox" name="chkFilterCustomers" runat="server">Enable 
					Filter<BR>
					<TR>
						<TD class="Subhead thbody" vAlign="top" align="left" width="0%"></TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="50%">Customer Name
						</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="25%">Customer Phone
						</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="25%">Customer Email
						</TD>
					</TR>
					<asp:Repeater id="rptAlugenCustomers" runat="server">
						<ItemTemplate>
							<tr>
								<td align="left" valign="top" class="normal tdbody">
									<asp:HyperLink NavigateUrl='<%# EditURL("CustomerID",DataBinder.Eval(Container.DataItem,"CustomerID"), "EditCustomer", "") %>' Visible="<%# IsEditable %>" runat="server" ID="Hyperlink2" BorderWidth=0>
										<asp:Image ID="Image1" Runat=server ImageUrl="~/images/edit.gif" AlternateText="Edit" Visible="<%#IsEditable%>" resourcekey="Edit"/>
									</asp:HyperLink>
								</td>
								<td align="left" valign="top" class="normal tdbody">
									<%# DataBinder.Eval(Container, "DataItem.CustomerName") %>
								</td>
								<td align="left" valign="top" class="normal tdbody">
									<%# DataBinder.Eval(Container, "DataItem.CustomerPhone") %>
								</td>
								<td align="left" valign="top" class="normal tdbody">
									<%# DataBinder.Eval(Container, "DataItem.CustomerEmail") %>
								</td>
							</tr>
						</ItemTemplate>
					</asp:Repeater></TABLE>
			</asp:panel><asp:panel class="Normal" id="pnlCodes" runat="server" Height="100%" Width="100%" name="pnlCodes"><INPUT id="chkFilterCodes" type="checkbox" name="chkFilterCodes" runat="server">Enable Filter<BR>
      <TABLE class="sortable filtered" id="codestable" cellSpacing="0" cellPadding="0" width="100%" border="1"
					name="codestable">
					<TR>
						<TD class="Subhead thbody" vAlign="top" align="left" width="0%"></TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="27%">Product</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="10%">Version</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="27%">Customer</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="20%">User Name</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="15%">License<br>type</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="12%">Generation<br>Date</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="12%">Expiration<br>Date</TD>
						<TD class="Subhead thbody" vAlign="top" align="left" width="5%">Level</TD>
					</TR>
					<asp:Repeater id="rptAlugenCodes" runat="server">
						<ItemTemplate>
							<tr>
								<td align="left" valign="middle" class="normal tdbody">
									<asp:HyperLink NavigateUrl='<%# EditURL("CodeID",DataBinder.Eval(Container.DataItem,"CodeID"), "EditCode", "") %>' Visible="<%# IsEditable %>" runat="server" ID="Hyperlink3" BorderWidth=0>
										<asp:Image ID="Image2" Runat=server ImageUrl="~/images/edit.gif" AlternateText="Edit" Visible="<%#IsEditable%>" resourcekey="Edit" />
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
					</asp:Repeater></TABLE>
			</asp:panel></td>
	</tr>
</table>
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
<INPUT type="hidden" id="txtLastPage" name="txtLastPage" runat="server">
