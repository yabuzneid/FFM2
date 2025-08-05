<%@ Page Title="" Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="JobReport.aspx.vb" Inherits="ProjectReport" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" Runat="Server">
<link rel="stylesheet" type="text/css" href="scripts/datePicker.css" />
<script src="scripts/jquery.datePicker.js" type="text/javascript"></script>
<script src="scripts/reports.js" type="text/javascript"></script>
<script type="text/javascript" src="scripts/date.js"></script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CPH_Wide" Runat="Server">
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="CPH_MainContent" Runat="Server">
<asp:Literal ID="LIT_Report" runat="server"></asp:Literal>
</asp:Content>

