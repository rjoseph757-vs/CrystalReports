<%@ Page Title="Home Page" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Default.aspx.cs" Inherits="CrystalReportWebAppFW480._Default" %>

<%@ Register Assembly="CrystalDecisions.Web, Version=13.0.4000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304" Namespace="CrystalDecisions.Web" TagPrefix="CR" %>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">

    <CR:CrystalReportViewer ID="CrystalReportViewer1" runat="server" AutoDataBind="True"
        Height="1040px" ReportSourceID="CrystalReportSource1" Width="900px" ToolPanelView="None" />
    <CR:CrystalReportSource ID="CrystalReportSource1" runat="server">
        <%--<Report FileName="~\CrystalReport1.rpt">--%>
        <Report FileName="~\Reports\AADDonors.rpt">
        </Report>
    </CR:CrystalReportSource>

</asp:Content>
