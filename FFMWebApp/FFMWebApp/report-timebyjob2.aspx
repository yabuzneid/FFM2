<%@ Page Title="" Language="vb" AutoEventWireup="true" MasterPageFile="~/MasterPage.Master" CodeBehind="report-timebyjob2.aspx.vb" Inherits="FFMWebApp.report_timebyjob2" %>

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

    <table class="daterange">
        <tr>
            <td class="label">Start</td>
            <td>
                <asp:TextBox ID="TXT_StartDate" runat="server" CssClass="pickdate"></asp:TextBox></td>

            <td class="label">End</td>
            <td>
                <asp:TextBox ID="TXT_EndDate" runat="server" CssClass="pickdate"></asp:TextBox></td>
        </tr>
        <tr>
            <td class="label">Job #</td>
            <td>
                <asp:TextBox ID="TXT_JobNumber" runat="server"></asp:TextBox></td>

            <td class="label" colspan="2">
                <asp:CheckBox ID="CHK_AVOnly" runat="server" Width="55px" />AV Employees Only</td>

        </tr>

        <tr>
            <td>
                <asp:Button ID="BTN_SubmitDates" runat="server" Text="Show" /></td>
        </tr>
    </table>
    <h2>Time by Job</h2>
    <asp:Literal ID="LIT_Results" runat="server"></asp:Literal>
    <a href="report-timebyjob2-excel.aspx">Open in Excel</a>
</asp:Content>

<asp:Content ID="Content3" ContentPlaceHolderID="CPH_Wide" runat="Server">
</asp:Content>

