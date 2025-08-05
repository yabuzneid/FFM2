<%@ Page Title="" Language="vb" AutoEventWireup="true" MasterPageFile="~/MasterPage.Master" CodeBehind="report-notes.aspx.vb" Inherits="FFMWebApp.report_notes" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" runat="Server">
    <link rel="stylesheet" type="text/css" href="scripts/datePicker.css" />
    <script src="scripts/jquery.datePicker.js" type="text/javascript"></script>
    <script type="text/javascript" src="scripts/date.js"></script>
    <script type="text/javascript" src="scripts/reports.js"></script>
    <!--[if IE]><script type="text/javascript" src="scripts/jquery.bgiframe.js"></script><![endif]-->

</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="CPH_MainContent" runat="Server">
    <h1>Reports</h1>

    <asp:SqlDataSource ID="SDS_Users" SelectCommand="select EmployeeID, LastName + ', ' + FirstName as FullName from tEmployees where Active = 1 order by LastName asc" runat="server" ConnectionString="<%$ ConnectionStrings:Fullconnection %>" />
    <asp:SqlDataSource ID="SDS_PMs" SelectCommand="select Initials from tEmployees where Active = 1 and PM = 1 and not Initials = '' order by Initials asc" runat="server" ConnectionString="<%$ ConnectionStrings:Fullconnection %>" />


    <table class="daterange">
        <tr>
            <td class="label">Start</td>
            <td>
                <asp:TextBox ID="TXT_StartDate" runat="server" CssClass="pickdate"></asp:TextBox></td>
            <td class="label">End</td>
            <td>
                <asp:TextBox ID="TXT_EndDate" runat="server" CssClass="pickdate"></asp:TextBox></td>
            <td class="label">User</td>
            <td>
                <asp:DropDownList ID="DDL_User" runat="server" DataSourceID="SDS_Users" DataTextField="FullName" DataValueField="EmployeeID" AppendDataBoundItems="true">
                    <asp:ListItem Text="" Value="0" />
                </asp:DropDownList>
            </td>

            <td class="label">PM</td>
            <td>
                <asp:DropDownList ID="DDL_PM" runat="server" DataSourceID="SDS_PMs" DataTextField="Initials" DataValueField="Initials" AppendDataBoundItems="true">
                    <asp:ListItem Text="" Value="" />
                </asp:DropDownList>
            </td>

            <td>
                <asp:Button ID="BTN_SubmitDates" runat="server" Text="Show" /></td>
        </tr>
    </table>
    <h2>Employee comments report</h2>
    <asp:Literal ID="LIT_Results" runat="server"></asp:Literal>

</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="CPH_Wide" runat="Server">
</asp:Content>
