<%@ Page Title="" Language="vb" AutoEventWireup="true" MasterPageFile="~/MasterPage.Master" CodeBehind="ExportList.aspx.vb" Inherits="FFMWebApp.ExportList" %>
<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" Runat="Server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="CPH_Wide" Runat="Server">

</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="CPH_MainContent" Runat="Server">
        <asp:GridView ID="GV_Files" runat="server" AllowPaging="true" PageSize="50" OnRowDataBound="GV_Files_RowDataBound" CssClass="listing">
    </asp:GridView>
</asp:Content>

