<%@ Register TagPrefix="cr" Namespace="CrystalDecisions.Web" Assembly="CrystalDecisions.Web, Version=9.1.3300.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" %>
<%@ Page Language="vb" AutoEventWireup="false" Codebehind="WebForm1.aspx.vb" Inherits="crnet_web_vbnet_SimplePreviewReport.WebForm1"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>WebForm1</title>
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<CR:CRYSTALREPORTVIEWER id="CrystalReportViewer1" style="Z-INDEX: 101; LEFT: 19px; POSITION: absolute; TOP: 225px" runat="server" Width="350px" Height="50px" PageToTreeRatio="4"></CR:CRYSTALREPORTVIEWER>
			<asp:TextBox id="TextBox1" style="Z-INDEX: 105; LEFT: 21px; POSITION: absolute; TOP: 72px" runat="server" Width="641px" Height="143px" TextMode="MultiLine" BorderStyle="Ridge"></asp:TextBox><asp:dropdownlist id="DropDownList1" style="Z-INDEX: 102; LEFT: 21px; POSITION: absolute; TOP: 34px" runat="server" Width="305px"></asp:dropdownlist>
			<asp:Label id="Label1" style="Z-INDEX: 103; LEFT: 21px; POSITION: absolute; TOP: 9px" runat="server" ToolTip="This sample demonstrates different ways of loading and viewing the same report.">Select a method for viewing the report</asp:Label>
			<asp:Button id="Button1" style="Z-INDEX: 104; LEFT: 351px; POSITION: absolute; TOP: 32px" runat="server" Text="Load Report"></asp:Button></form>
	</body>
</HTML>
