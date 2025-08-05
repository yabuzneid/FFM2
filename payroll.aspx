<%@ Page Title="" Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="payroll.aspx.vb" Inherits="payroll" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" Runat="Server">
<script src="scripts/jquery.tablesorter.js" type="text/javascript" language="javascript"></script>
<script src="scripts/payroll.js" type="text/javascript" language="javascript"></script>

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CPH_Wide" Runat="Server">

<div id="typefilter">
Filter by type: 
<asp:DropDownList ID="DDL_Type" runat="server" AutoPostBack="true">
<asp:ListItem Text="" Value=""></asp:ListItem>
<asp:ListItem Text="G" Value="G"></asp:ListItem>
<asp:ListItem Text="K" Value="K"></asp:ListItem>
</asp:DropDownList>

Filter by office:
<asp:SqlDataSource ID="SDS_Office" runat="server" ConnectionString="<%$ConnectionStrings:FullConnection %>"
SelectCommand="select OfficeID, OfficeName from tOffices order by OfficeName"></asp:SqlDataSource>
<asp:DropDownList ID="DDL_Office" runat="server" AutoPostBack="true" DataSourceID="SDS_Office" 
AppendDataBoundItems="true" DataValueField="OfficeID" DataTextField="OfficeName">
<asp:ListItem Value="" Text=""  />
</asp:DropDownList>
</div>

<asp:Literal ID="LIT_Grid" runat="server"></asp:Literal>

</asp:Content>

