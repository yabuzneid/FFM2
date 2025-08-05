<%@ Page Title="" Language="vb" AutoEventWireup="true" MasterPageFile="~/MasterPage.Master" CodeBehind="entertime.aspx.vb" Inherits="FFMWebApp.entertime" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" runat="Server">
    <script src="scripts/worksheets.js" type="text/javascript" language="javascript"></script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CPH_MainContent" runat="Server">

    <asp:DropDownList ID="DDL_DateRange" runat="server"></asp:DropDownList>
    <asp:Label ID="LBL_Info" runat="server" Visible="false" CssClass="infolabel"></asp:Label>
    <asp:Literal ID="LIT_Main" runat="server"></asp:Literal>

</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="CPH_Wide" runat="Server">
</asp:Content>
