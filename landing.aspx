<%@ Page Title="" Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="landing.aspx.vb" Inherits="landing" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CPH_MainContent" Runat="Server">
<h1>Approved Time</h1>
<asp:DropDownList ID="DDL_DateRange" runat="server"></asp:DropDownList>

<asp:SqlDataSource ID="SDS_Office" runat="server" ConnectionString="<%$ConnectionStrings:FullConnection %>"
SelectCommand="select OfficeID, OfficeName from tOffices order by OfficeName"></asp:SqlDataSource>
<asp:DropDownList ID="DDL_Office" runat="server" DataSourceID="SDS_Office" 
AppendDataBoundItems="true" DataValueField="OfficeID" DataTextField="OfficeName">
<asp:ListItem Value="" Text=""  />
</asp:DropDownList>

<table style="width:550px; margin:0 auto; "><tr>

<td><asp:Literal ID="LIT_Main" runat="server"></asp:Literal></td>
<td style="vertical-align:top;"><asp:Literal ID="LIT_OfficeTotals" runat="server"></asp:Literal>
<!-- [<a href="csvlist.aspx">Saved CSV Exports</a>] -->
<asp:Literal ID="LIT_UnprocessedWarnings" runat="server"></asp:Literal>

    <a href="ExportList.aspx">List of exported timesheets</a>

</td>

</tr></table>


<asp:Button ID="BTN_Payroll" runat="server" Text="Open Payroll" OnClick="OpenPayroll" />
<asp:Button ID="BTN_Lock" runat="server" Text="Lock Week" OnClick="LockWeek" />
<asp:Button ID="BTN_Unlock" runat="server" Text="Unlock Week" OnClick="UnlockWeek" Visible="false" />


</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="CPH_Wide" Runat="Server">
</asp:Content>

