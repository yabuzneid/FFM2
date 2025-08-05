<%@ Page Title="" Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="entertime.aspx.vb" Inherits="entertime" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" Runat="Server">
<script src="scripts/worksheets.js" type="text/javascript" language="javascript"></script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CPH_MainContent" Runat="Server">

<asp:DropDownList ID="DDL_DateRange" runat="server"></asp:DropDownList>
<asp:Label id="LBL_Info" runat="server" Visible="false" CssClass="infolabel"></asp:Label>
<asp:Literal ID="LIT_Main" runat="server"></asp:Literal>

</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="CPH_Wide" Runat="Server">
</asp:Content>

