<%@ Page Title="" Language="vb" AutoEventWireup="true" MasterPageFile="~/MasterPage.Master" CodeBehind="reports.aspx.vb" Inherits="FFMWebApp.reports" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" runat="Server">
    <link rel="stylesheet" type="text/css" href="scripts/datePicker.css" />
    <script src="scripts/jquery.datePicker.js" type="text/javascript"></script>
    <script type="text/javascript" src="scripts/date.js"></script>
    <script type="text/javascript" src="scripts/reports.js"></script>
    <!--[if IE]><script type="text/javascript" src="scripts/jquery.bgiframe.js"></script><![endif]-->

</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="CPH_MainContent" runat="Server">
    <h1>Reports</h1>


    <table>
        <tr>
            <td>
                <ul style="text-align: left; font-weight: bold; width: 500px; margin: 0 auto;">
                    <li><a href="report-timebyjob.aspx">Time by job report</a></li>
                    <li><a href="report-timebyjob2.aspx">Time by job report - AV dept</a></li>
                    <li><a href="report-notes.aspx">Employee comments report</a></li>
                </ul>

            </td>


            <td>


                <h2>Jump to Job Report</h2>

                <form name="JobForm" id="JobForm" method="get" action="jobreport.aspx">
                    <table class="listing" style="width: 400px;">
                        <tr>
                            <td>JUMP TO REPORT</td>
                            <td colspan="3">
                                <asp:Literal ID="LIT_JobNumberDropDown" runat="server" /></td>
                        </tr>
                        <tr>
                            <td>JUMP TO PERIOD</td>
                            <td>
                                <asp:TextBox ID="startdate" runat="server" CssClass="pickdate" Width="80px"></asp:TextBox>
                            </td>
                            <td colspan="2">
                                <asp:TextBox ID="enddate" runat="server" CssClass="pickdate" Width="80px"></asp:TextBox></td>
                        </tr>
                        <tr>
                            <td colspan="4">
                                <input type="button" id="submitjobnumber" value="Go" />
                            </td>
                        </tr>
                    </table>
                </form>
            </td>
        </tr>
    </table>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="CPH_Wide" runat="Server">
</asp:Content>

