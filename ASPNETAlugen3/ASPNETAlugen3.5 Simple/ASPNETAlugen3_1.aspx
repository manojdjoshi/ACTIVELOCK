<%@ Page Language="vb" AutoEventWireup="false" Inherits="ASPNETAlugen3.ASPNETAlugen3_1" CodeFile="ASPNETAlugen3_1.aspx.vb" EnableEventValidation="false" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
  <HEAD>
		<title>ASPNETAlugen3_1</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
  
  <SCRIPT TYPE="text/javascript">
<!--
// copyright 1999 Idocs, Inc. http://www.idocs.com
// Distribute this script freely but keep this notice in place
function letternumber(e)
{
var key;
var keychar;

if (window.event)
   key = window.event.keyCode;
else if (e)
   key = e.which;
else
   return true;
keychar = String.fromCharCode(key);
keychar = keychar.toLowerCase();

// control keys
if ((key==null) || (key==0) || (key==8) || 
    (key==9) || (key==13) || (key==27) )
   return true;

// alphas and numbers
else if ((("abcdefghijklmnopqrstuvwxyz0123456789.").indexOf(keychar) > -1))
   return true;
else
   return false;
}
//-->
</SCRIPT>
</HEAD>
	<body>
		<form id="Form1" method="post" runat="server">
			<asp:label id="Label10" style="Z-INDEX: 100; LEFT: 58px; POSITION: absolute; TOP: 15px; text-align: center;" runat="server"
				Width="422px" Height="21px" Font-Size="Small" Font-Names="Microsoft Sans Serif" BackColor="Blue"
				Font-Bold="True" ForeColor="White">Alugen - Activelock Key Generator for ASP.NET v3.6</asp:label><asp:label id="lblUpdateStatus" style="Z-INDEX: 101; LEFT: 132px; POSITION: absolute; TOP: 348px"
				runat="server" Width="338px" Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" ForeColor="Red"></asp:label><asp:button id="Button1" style="Z-INDEX: 102; LEFT: 64px; POSITION: absolute; TOP: 48px" runat="server"
				Width="189px" Height="23px" ForeColor="ControlDark" BorderStyle="Inset" Text="Product Code Generator"></asp:button><asp:button id="cmdLicenseKeyGenerator" style="Z-INDEX: 103; LEFT: 284px; POSITION: absolute; TOP: 48px"
				runat="server" Width="189px" Height="23px" Text="License Key Generator"></asp:button><IMG style="Z-INDEX: 132; LEFT: 528px; POSITION: absolute; TOP: 13px" alt="" src="small-logo.bmp">
			<asp:label id="Label11" style="Z-INDEX: 104; LEFT: 520px; POSITION: absolute; TOP: 68px" runat="server"
				Width="68px" Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" Font-Bold="True" ForeColor="Blue">Activelock 3</asp:label><asp:label id="Label7" style="Z-INDEX: 105; LEFT: 18px; POSITION: absolute; TOP: 80px" runat="server"
				Width="9px" Font-Size="X-Small" Font-Names="Microsoft Sans Serif">Name</asp:label><asp:label id="Label1" style="Z-INDEX: 106; LEFT: 9px; POSITION: absolute; TOP: 103px" runat="server"
				Width="9px" Font-Size="X-Small" Font-Names="Microsoft Sans Serif">Version</asp:label>
            <asp:Label ID="Label12" runat="server" Font-Names="Microsoft Sans Serif" Font-Size="X-Small"
                Style="z-index: 133; left: 6px; position: absolute; top: 127px" Width="9px">Strength</asp:Label>
            <asp:label id="Label2" style="Z-INDEX: 108; LEFT: 22px; POSITION: absolute; TOP: 150px" runat="server"
				Width="9px" Font-Size="X-Small" Font-Names="Microsoft Sans Serif">Code</asp:label><asp:textbox id="txtName" style="Z-INDEX: 109; LEFT: 64px; POSITION: absolute; TOP: 77px" runat="server"
				Width="410px" Height="21px"></asp:textbox><asp:textbox id="txtVer" style="Z-INDEX: 110; LEFT: 64px; POSITION: absolute; TOP: 101px" runat="server"
				Width="190" Height="21px"></asp:textbox><asp:label id="Label3" style="Z-INDEX: 111; LEFT: 64px; POSITION: absolute; TOP: 150px" runat="server"
				Width="189px" Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" BackColor="Desktop" ForeColor="ActiveCaptionText" BorderStyle="Solid"
				BorderColor="ActiveCaption" BorderWidth="1px">VCode for the new application</asp:label><asp:label id="Label4" style="Z-INDEX: 112; LEFT: 284px; POSITION: absolute; TOP: 150px" runat="server"
				Width="188px" Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" BackColor="Desktop" ForeColor="ActiveCaptionText" BorderStyle="Solid" BorderColor="ActiveCaption" BorderWidth="1px">GCode for the new application</asp:label><asp:textbox id="txtCode1" style="Z-INDEX: 113; LEFT: 64px; POSITION: absolute; TOP: 167px" runat="server"
				Width="189" Height="55px" Font-Size="XX-Small" TextMode="MultiLine"></asp:textbox><asp:textbox id="txtCode2" style="Z-INDEX: 114; LEFT: 284px; POSITION: absolute; TOP: 167px"
				runat="server" Width="188px" Height="55px" Font-Size="XX-Small" TextMode="MultiLine"></asp:textbox><asp:button id="cmdAdd" style="Z-INDEX: 115; LEFT: 64px; POSITION: absolute; TOP: 233px" runat="server"
				Width="189px" Height="23px" Text="Add to Product List"></asp:button><asp:label id="Label5" style="Z-INDEX: 116; LEFT: 13px; POSITION: absolute; TOP: 290px" runat="server"
				Width="9px" Font-Size="X-Small" Font-Names="Microsoft Sans Serif">VCode</asp:label><asp:label id="Label6" style="Z-INDEX: 117; LEFT: 12px; POSITION: absolute; TOP: 313px" runat="server"
				Width="9px" Font-Size="X-Small" Font-Names="Microsoft Sans Serif">GCode</asp:label><asp:textbox id="txtCode1_2" style="Z-INDEX: 118; LEFT: 64px; POSITION: absolute; TOP: 287px"
				runat="server" Width="411px" Height="21px"></asp:textbox><asp:textbox id="txtCode2_2" style="Z-INDEX: 119; LEFT: 64px; POSITION: absolute; TOP: 310px"
				runat="server" Width="411" Height="21px"></asp:textbox><asp:label id="Label8" style="Z-INDEX: 120; LEFT: 64px; POSITION: absolute; TOP: 272px" runat="server"
				Width="411px" Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" BackColor="Desktop" ForeColor="ActiveCaptionText" BorderStyle="Solid" BorderColor="ActiveCaption"
				BorderWidth="1px">Product Codes for the Selected Product</asp:label><asp:label id="Label9" style="Z-INDEX: 121; LEFT: 64px; POSITION: absolute; TOP: 348px" runat="server"
				Width="60px" Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" BackColor="Desktop" ForeColor="ActiveCaptionText" BorderStyle="Solid" BorderColor="ActiveCaption"
				BorderWidth="1px">Product List</asp:label><asp:button id="cmdCodeGen" style="Z-INDEX: 122; LEFT: 504px; POSITION: absolute; TOP: 150px"
				runat="server" Width="97px" Height="23px" Text="Generate"></asp:button><asp:button id="cmdValidate" style="Z-INDEX: 123; LEFT: 504px; POSITION: absolute; TOP: 278px"
				runat="server" Width="97px" Height="23px" Text="Validate"></asp:button><asp:button id="cmdRemove" style="Z-INDEX: 124; LEFT: 504px; POSITION: absolute; TOP: 309px"
				runat="server" Width="97px" Height="23px" Text="Remove"></asp:button><asp:datagrid id="gridProds" style="Z-INDEX: 125; LEFT: 64px; POSITION: absolute; TOP: 368px"
				runat="server" Width="180px" Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" BackColor="White" BorderStyle="None" BorderColor="#CCCCCC" BorderWidth="1px" CellPadding="3"
				HorizontalAlign="Left">
				<FooterStyle ForeColor="#000066" BackColor="White"></FooterStyle>
				<SelectedItemStyle Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" Height="10px" ForeColor="White"
					Width="20px" VerticalAlign="Top" BackColor="CadetBlue"></SelectedItemStyle>
				<EditItemStyle Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" Height="10px" Width="20px"
					VerticalAlign="Top"></EditItemStyle>
				<AlternatingItemStyle Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" Height="10px" VerticalAlign="Top"></AlternatingItemStyle>
				<ItemStyle Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" Height="10px" ForeColor="#000066"
					Width="20px" VerticalAlign="Top"></ItemStyle>
				<HeaderStyle Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" Font-Bold="True" Height="10px"
					BorderWidth="5px" ForeColor="White" BorderStyle="Outset" Width="20px" VerticalAlign="Top"
					BackColor="#006699"></HeaderStyle>
				<Columns>
					<asp:ButtonColumn Visible="False" Text="Select" CommandName="Select"></asp:ButtonColumn>
				</Columns>
				<PagerStyle Font-Size="XX-Small" Font-Names="Microsoft Sans Serif" HorizontalAlign="Left" ForeColor="#000066"
					BackColor="White" Mode="NumericPages"></PagerStyle>
			</asp:datagrid>
            <asp:RadioButton ID="optStrenght0" runat="server" Checked="True" Font-Names="Microsoft Sans Serif"
                Font-Size="XX-Small" GroupName="strength" Style="z-index: 126; left: 59px; position: absolute;
                top: 126px" Text="ALCrypto 1024-bit" />
            <asp:RadioButton ID="optStrenght1" runat="server" Font-Names="Microsoft Sans Serif"
                Font-Size="XX-Small" GroupName="strength" Style="z-index: 127; left: 164px; position: absolute;
                top: 126px" Text="4096-bit" />
            <asp:RadioButton ID="optStrenght2" runat="server" Font-Names="Microsoft Sans Serif"
                Font-Size="XX-Small" GroupName="strength" Style="z-index: 128; left: 223px; position: absolute;
                top: 126px" Text="2048-bit" />
            <asp:RadioButton ID="optStrenght3" runat="server" Font-Names="Microsoft Sans Serif"
                Font-Size="XX-Small" GroupName="strength" Style="z-index: 129; left: 282px; position: absolute;
                top: 126px" Text="1536-bit" />
            <asp:RadioButton ID="optStrenght4" runat="server" Font-Names="Microsoft Sans Serif"
                Font-Size="XX-Small" GroupName="strength" Style="z-index: 130; left: 341px; position: absolute;
                top: 126px" Text="1024-bit" Width="64px" />
            <asp:RadioButton ID="optStrenght5" runat="server" Font-Names="Microsoft Sans Serif"
                Font-Size="XX-Small" GroupName="strength" Style="z-index: 131; left: 408px; position: absolute;
                top: 126px" Text="512-bit" Width="66px" />
        </form>
		<script>
<asp:Literal id="ltlAlert" runat="server"
        EnableViewState="False">
</asp:Literal>
		</script>
	</body>
</HTML>
