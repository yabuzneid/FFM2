<%@ Page Title="" Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="timesheets.aspx.vb" Inherits="timesheets" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CPH_MainContent" Runat="Server">

<div id="synccontrol"><asp:Literal ID="LIT_SyncData" runat="server"></asp:Literal></div>

<h1>Approve Time</h1>

<asp:Literal ID="LIT_DatesDropDown" runat="server"></asp:Literal>

<asp:DropDownList ID="DDL_DateRange" runat="server"></asp:DropDownList>

<asp:DropDownList ID="DDL_Offices" runat="server"></asp:DropDownList>

<asp:DropDownList ID="DDL_Type" runat="server"></asp:DropDownList>

<asp:DropDownList ID="DDL_ShowAsPM" runat="server"></asp:DropDownList>


<table style="margin:0 auto;">
<tr>
<td style="width:70%; vertical-align:top;">
<asp:Literal ID="LIT_WorkerList" runat="server"></asp:Literal>
<asp:Literal ID="LIT_ReportLink" runat="server"></asp:Literal>


<asp:Literal ID="LIT_SpeedReport" runat="server"></asp:Literal>
</td>
<td style="width:30%; vertical-align:top;"><asp:Literal ID="LIT_ProjectList" runat="server"></asp:Literal></td>
</tr>
</table>

</asp:Content>

