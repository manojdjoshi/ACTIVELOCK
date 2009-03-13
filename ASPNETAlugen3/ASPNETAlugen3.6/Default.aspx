<%@ Page Language="vb" AutoEventWireup="false" Inherits="ASPNETAlugen3.ASPNETAlugen3" CodeFile="~/Default.aspx.vb" %>
<%@ Register TagPrefix="ajax" Namespace="MagicAjax.UI.Controls" Assembly="MagicAjax" %>
<%@ Register TagPrefix="cc1" Namespace="msWebControlsLibrary" Assembly="msWebControlsLibrary" %>
<%@ Register TagPrefix="uc1" TagName="ImageTextButton" Src="ImageTextButton.ascx" %>

<HTML>
  <HEAD>
		<title>ASPNETAlugen3.5</title>
		<LINK href="Styles.css" type="text/css" rel="STYLESHEET">
			<script src="ASPNETAlugen3.js" type="text/javascript"></script>
</HEAD>
	<body>
		<form id="frmASPNETAlugen3" method="post" runat="server">
			<table cellSpacing="0" cellPadding="0" width="660" border="0">
				<tr>
					<td vAlign="middle" align="center" style="width: 59px"><IMG height="31" src="images/I_Trust_AL_small.gif" width="68" border="0"></td>
					<td vAlign="middle" align="center" style="text-align: left"><b>Alugen - ActiveLock Key Generator for 
							ASP.NET v3.6</b></td>
				</tr>
				<tr>
					<td colSpan="2">&nbsp;</td>
				</tr>
				<tr>
					<td colSpan="2"><ajax:ajaxpanel id="pnlMagicAjax" runat="server">
      <TABLE cellSpacing=0 cellPadding=0 width=660 border=0>
        <TR>
          <TD vAlign=top><asp:Button id=cmdProducts runat="server" Text="Products"></asp:Button>&nbsp; 
<asp:Button id=cmdLicenses runat="server" Text="Licenses"></asp:Button></TD></TR>
        <TR>
          <TD><asp:Panel id=pnlProducts runat="server">
            <TABLE cellSpacing=0 cellPadding=0 width=660 border=1>
              <TR>
                <TD class=rowHeader vAlign=top width=150>Name</TD>
                <TD colSpan=2><asp:TextBox id=txtProductName runat="server" Width="550" AutoPostBack="True"></asp:TextBox></TD></TR>
              <TR>
                <TD class=rowHeader vAlign=top width=150 style="height: 35px">Version</TD>
                <TD vAlign=top style="height: 35px"><asp:TextBox id=txtProductVersion runat="server" Width="432px" AutoPostBack="True"></asp:TextBox></TD>
                <TD vAlign=top align=center width=115 style="height: 35px">
                  <TABLE cellSpacing=1 cellPadding=0 width="100%" border=0>
                    <TR>
                      <TD vAlign=top><cc1:ExImageButton id=cmdGenerateCode runat="server" Text="Generate new codes" ToolTip="Generate new codes" ImageUrl="images/generate_codes.gif" DisableImageURL="images/generate_codes_dis.gif"></cc1:ExImageButton></TD></TR>
                    <TR>
                      <TD vAlign=top><cc1:ExImageButton id=cmdValidateCode runat="server" Text="Validate codes" ToolTip="Validate codes" ImageUrl="images/validate_codes.gif" DisableImageURL="images/validate_codes_dis.gif"></cc1:ExImageButton></TD></TR></TABLE></TD></TR>
                <tr>
                    <td class="rowHeader" valign="top" width="150">
                        Strength</td>
                    <td valign="top">
                        <asp:RadioButton ID="optStrenght0" runat="server" Checked="True" Font-Size="XX-Small"
                            Style="z-index: 100; left: 128px; position: absolute; top: 160px" Text="ALCrypto 1024-bit" GroupName="strength" Font-Bold="False" Width="137px" />
                        <asp:RadioButton ID="optStrenght1" runat="server" Font-Size="XX-Small" Style="z-index: 101;
                            left: 231px; position: absolute; top: 160px" Text="4096-bit" GroupName="strength" Font-Bold="False" Width="86px" />
                        <asp:RadioButton ID="optStrenght2" runat="server" Font-Size="XX-Small" Style="z-index: 102;
                            left: 287px; position: absolute; top: 160px" Text="2048-bit" GroupName="strength" Font-Bold="False" Width="86px" />
                        <asp:RadioButton ID="optStrenght3" runat="server" Font-Size="XX-Small" Style="z-index: 103;
                            left: 343px; position: absolute; top: 160px" Text="1536-bit" GroupName="strength" Font-Bold="False" Width="86px" />
                        <asp:RadioButton ID="optStrenght4" runat="server" Font-Size="XX-Small" Style="z-index: 104;
                            left: 399px; position: absolute; top: 160px" Text="1024-bit" Width="77px" GroupName="strength" Font-Bold="False" />
                        <asp:RadioButton ID="optStrenght5" runat="server" Font-Size="XX-Small" Style="z-index: 106;
                            left: 458px; position: absolute; top: 160px" Text="512-bit" Width="72px" GroupName="strength" Font-Bold="False" />
                    </td>
                    <td align="center" valign="top" width="115">
                    </td>
                </tr>
              <TR>
                <TD class=rowHeader vAlign=top width=150><asp:Image id=imgVCode runat="server" ImageUrl="images/keys.gif"></asp:Image>&nbsp; 
                  VCode (PUB_KEY)</TD>
                <TD vAlign=top><asp:TextBox id=txtVCode runat="server" Width="432px" TextMode="MultiLine" Rows="3"></asp:TextBox></TD>
                <TD vAlign=middle align=center width=115><IMG 
                  class=htmlimagebuttons id=cmdCopyVCode alt="Copy VCode" 
                  src="images/copy_vcode.gif" border=0 runat="server"></TD></TR>
              <TR>
                <TD class=rowHeader vAlign=top width=150><asp:Image id=imgGCode runat="server" ImageUrl="images/keys.gif"></asp:Image>&nbsp; 
                  GCode (PRV_KEY)</TD>
                <TD vAlign=top><asp:TextBox id=txtGCode runat="server" Width="432px" TextMode="MultiLine" Rows="3"></asp:TextBox></TD>
                <TD vAlign=middle align=center width=115><IMG 
                  class=htmlimagebuttons id=cmdCopyGCode alt="Copy GCode" 
                  src="images/copy_gcode.gif" border=0 runat="server"></TD></TR>
              <TR>
                <TD class=rowHeader vAlign=top width=150></TD>
                <TD class=rowHeader vAlign=top colSpan=2><cc1:ExImageButton id=cmdAddProduct runat="server" Text="Add product to list" ToolTip="Add product to list" ImageUrl="images/add_to_list.gif" DisableImageURL="images/add_to_list_dis.gif"></cc1:ExImageButton>&nbsp; 
<cc1:ExImageButton id=cmdRemoveProduct runat="server" Text="Remove product from list" ToolTip="Remove product from list" ImageUrl="images/remove_from_list.gif" DisableImageURL="images/remove_from_list_dis.gif"></cc1:ExImageButton></TD></TR>
              <TR>
                <TD colSpan=3>
                  <DIV style="OVERFLOW: auto; WIDTH: 660px; HEIGHT: 200px"><asp:DataGrid id=grdProducts runat="server" Width="644px" BorderWidth="1px" BorderColor="#CCCCCC" HorizontalAlign="Left" Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" CellPadding="3" AllowSorting="True">
																<SelectedItemStyle BackColor="#C8DCF0"></SelectedItemStyle>
																<AlternatingItemStyle Height="10px" VerticalAlign="Top"></AlternatingItemStyle>
																<HeaderStyle Font-Size="XX-Small" Font-Bold="True" ForeColor="White" BackColor="#646464"></HeaderStyle>
																<Columns>
																	<asp:ButtonColumn Visible="False" Text="Select" CommandName="Select"></asp:ButtonColumn>
																</Columns>
															</asp:DataGrid></DIV></TD></TR>
              <TR>
                <TD colSpan=3></TD></TR></TABLE></asp:Panel><asp:Panel id=pnlLicenses runat="server">
            <TABLE cellSpacing=0 cellPadding=0 width=660 border=1>
              <TR>
                <TD class=rowHeader vAlign=top height=7 style="width: 150px">Product</TD>
                <TD vAlign=top height=7><asp:DropDownList id=cboProduct runat="server" Width="180px"></asp:DropDownList></TD>
                <TD width=25 height=7>&nbsp;</TD>
                <TD class=rowHeader vAlign=top width=144 height=7>Registered 
                  level</TD>
                <TD height=7><asp:DropDownList id=cboRegLevel runat="server" Width="200px"></asp:DropDownList></TD></TR>
              <TR>
                <TD class=rowHeader vAlign=top height=1 style="width: 150px">License 
                type</TD>
                <TD vAlign=top height=1><asp:DropDownList id=cboLicenseType runat="server" Width="180px" AutoPostBack="True">
															<asp:ListItem Value="0">Time locked</asp:ListItem>
															<asp:ListItem Value="1">Periodic</asp:ListItem>
															<asp:ListItem Value="2">Permanent</asp:ListItem>
														</asp:DropDownList></TD>
                <TD width=25 height=1>&nbsp; </TD>
                <TD class=rowHeader vAlign=top width=144 height=1>&nbsp;</TD>
                <TD vAlign=top height=1><asp:CheckBox id=chkUseItemData runat="server" Text="Use item data for code" CssClass="rowHeader"></asp:CheckBox></TD></TR>
              <TR>
                <TD class=rowHeader vAlign=top style="width: 150px"><asp:Label id=lblExpiry runat="server">Expire on date</asp:Label></TD>
                <TD vAlign=top><asp:TextBox id=txtDays runat="server" Width="160px"></asp:TextBox></TD>
                <TD align=center width=25><asp:ImageButton id=cmdSelectExpireDate runat="server" ToolTip="Select expire date" ImageUrl="images/calendar.gif" BorderWidth="0px" BorderColor="#808080"></asp:ImageButton></TD>
                <TD class=rowHeader vAlign=top width=144>&nbsp; 
                  <asp:PlaceHolder id=plhDate runat="server"></asp:PlaceHolder></TD>
                <TD vAlign=top>&nbsp;</TD></TR>
              <TR>
                <TD class=rowHeader vAlign=top style="width: 150px">Install code</TD>
                <TD vAlign=top width=330 colSpan=3><asp:TextBox id=txtInstallCode runat="server" Width="330px" AutoPostBack="True"></asp:TextBox></TD>
                <TD class=rowHeader vAlign=top align=left><IMG 
                  class=htmlimagebuttons id=cmdPasteInstallCode 
                  alt="Paste installation code" 
                  src="images/paste_install_code.gif" border=0 
              runat="server"></TD></TR>
              <TR>
                <TD class=rowHeader vAlign=top style="width: 150px">User name</TD>
                <TD vAlign=top colSpan=4><asp:TextBox id=txtUserName runat="server" Width="545px"></asp:TextBox></TD></TR>
                <tr>
                    <td class="rowHeader" style="width: 150px; height: 178px" valign="top">
                    </td>
                    <td class="rowheader" colspan="4" style="height: 178px" valign="top">
                        <asp:CheckBox ID="chkLockComputer" runat="server" CssClass="rowheader" Font-Names="Microsoft Sans Serif"
                            Font-Size="XX-Small" Style="z-index: 100; left: 127px; position: absolute; top: 792px"
                            Text="Lock to Computer Name" Width="427px" />
                        <asp:CheckBox ID="chkLockHD" runat="server" CssClass="rowheader" Font-Names="Microsoft Sans Serif"
                            Font-Size="XX-Small" Style="z-index: 101; left: 127px; position: absolute; top: 812px"
                            Text="Lock to HDD Volume Serial" Width="427px" />
                        <asp:CheckBox ID="chkLockHDfirmware" runat="server" CssClass="rowheader" Font-Names="Microsoft Sans Serif"
                            Font-Size="XX-Small" Style="z-index: 102; left: 127px; position: absolute; top: 832px"
                            Text="Lock to HDD Firmware Serial" Width="427px" />
                        <asp:CheckBox ID="chkLockWindows" runat="server" CssClass="rowheader" Font-Names="Microsoft Sans Serif"
                            Font-Size="XX-Small" Style="z-index: 103; left: 127px; position: absolute; top: 852px"
                            Text="Lock to Windows Serial" Width="427px" />
                        <asp:CheckBox ID="chkLockBIOS" runat="server" CssClass="rowheader" Font-Names="Microsoft Sans Serif"
                            Font-Size="XX-Small" Style="z-index: 104; left: 127px; position: absolute; top: 872px"
                            Text="Lock to BIOS Version" Width="427px" />
                        <asp:CheckBox ID="chkLockMotherboard" runat="server" CssClass="rowheader" Font-Names="Microsoft Sans Serif"
                            Font-Size="XX-Small" Style="z-index: 105; left: 127px; position: absolute; top: 892px"
                            Text="Lock to Motherboard Serial" Width="427px" />
                        <asp:CheckBox ID="chkLockIP" runat="server" CssClass="rowheader" Font-Names="Microsoft Sans Serif"
                            Font-Size="XX-Small" Style="z-index: 106; left: 127px; position: absolute; top: 912px"
                            Text="Lock to IP Number" Width="427px" />
                        <asp:CheckBox ID="chkLockMACaddress" runat="server" CssClass="rowheader" Font-Names="Microsoft Sans Serif"
                            Font-Size="XX-Small" Style="z-index: 108; left: 127px; position: absolute; top: 772px"
                            Text="Lock to MAC Address" Width="427px" />
                    </td>
                </tr>
              <TR>
                <TD vAlign=top style="width: 150px">
                  <TABLE cellSpacing=0 cellPadding=0 width="100%" border=0>
                    <TR>
                      <TD class=rowHeader vAlign=top>License key </TD></TR>
                    <TR>
                      <TD class=rowHeader vAlign=top align=right><IMG 
                        height=25 alt="License key" src="images/KeyLock.gif" 
                        width=26 border=0></TD></TR>
                    <TR>
                      <TD class=rowHeader vAlign=top><cc1:ExImageButton id=cmdGenerateLicenseKey runat="server" Text="Generate license key" ToolTip="Generate license key" ImageUrl="images/generate_license.gif" DisableImageURL="images/generate_license_dis.gif" Enabled="False"></cc1:ExImageButton></TD></TR>
                    <TR>
                      <TD class=rowHeader vAlign=top><IMG 
                        class=htmlimagebuttons id=cmdCopyLicenseKey 
                        alt="Copy license key" src="images/copy_license.gif" 
                        runat="server"></TD></TR>
                    <TR>
                      <TD class=rowHeader vAlign=top><cc1:ExImageButton id=cmdPrintLicenseKey runat="server" Text="Print license key" ToolTip="Print license key" ImageUrl="images/print_license.gif" DisableImageURL="images/print_license_dis.gif" Enabled="False"></cc1:ExImageButton></TD></TR>
                    <TR>
                      <TD class=rowHeader vAlign=top><cc1:ExImageButton id=cmdEmailLicenseKey runat="server" Text="Email license key" ToolTip="Email license key" ImageUrl="images/email_license.gif" DisableImageURL="images/email_license_dis.gif" Enabled="False"></cc1:ExImageButton></TD></TR>
                    <TR>
                      <TD class=rowHeader vAlign=top style="height: 22px"><cc1:ExImageButton id=cmdSaveLicenseFile runat="server" Text="Save license key" ToolTip="Save license key" ImageUrl="images/save_license.gif" DisableImageURL="images/save_license_dis.gif" Enabled="False" ></cc1:ExImageButton></TD></TR></TABLE></TD>
                <TD vAlign=top colSpan=4><asp:TextBox id=txtLicenseKey runat="server" Width="545px" TextMode="MultiLine" Rows="9" Height="155px"></asp:TextBox></TD></TR></TABLE></asp:Panel></TD></TR></TABLE><INPUT 
      id=sortExpression type=hidden name=sortExpression runat="server"> <INPUT 
      id=sortOrder type=hidden name=sortOrder runat="server"> <asp:PlaceHolder id=plhSay runat="server"></asp:PlaceHolder><BR>
      <DIV id=myCalendar runat="server"><asp:Calendar id=Calendar1 runat="server" Width="200px" BorderColor="#999999" Font-Size="8pt" Font-Names="Verdana" CellPadding="4" Height="180px" BackColor="White" ForeColor="Black" Visible="False">
									<TodayDayStyle ForeColor="Black" BackColor="#CCCCCC"></TodayDayStyle>
									<SelectorStyle BackColor="#CCCCCC"></SelectorStyle>
									<NextPrevStyle VerticalAlign="Bottom"></NextPrevStyle>
									<DayHeaderStyle Font-Size="7pt" Font-Bold="True" BackColor="#CCCCCC"></DayHeaderStyle>
									<SelectedDayStyle Font-Bold="True" ForeColor="White" BackColor="#666666"></SelectedDayStyle>
									<TitleStyle Font-Bold="True" BorderColor="Black" BackColor="#999999"></TitleStyle>
									<WeekendDayStyle BackColor="#FFFFCC"></WeekendDayStyle>
									<OtherMonthDayStyle ForeColor="Gray"></OtherMonthDayStyle>
								</asp:Calendar></DIV>
						</ajax:ajaxpanel></td>
				</tr>
			</table>
		</form>
		<script>
		<asp:Literal id="ltlAlert" runat="server"
						EnableViewState="False">
		</asp:Literal>
		</script>
	</body>
</HTML>
