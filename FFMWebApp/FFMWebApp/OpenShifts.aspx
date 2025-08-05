<%@ Page Title="" Language="vb" AutoEventWireup="true" MasterPageFile="~/MasterPage.Master" CodeBehind="OpenShifts.aspx.vb" Inherits="FFMWebApp.OpenShifts" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" runat="Server">
    <link rel="stylesheet" type="text/css" href="scripts/datePicker.css" />
    <script src="scripts/jquery.datePicker.js" type="text/javascript"></script>
    <script src="scripts/reports.js" type="text/javascript"></script>
    <script type="text/javascript" src="scripts/date.js"></script>
    <script type="text/javascript">

        $(document).ready(function () {
            //console.log("start");

            $("input.IgnoreButton").click(function () {
                //console.log("found button");
                //console.log("run " + $(this).attr("title"));
                $("input#actiontoignore").val($(this).attr("title"));
            });
        });

    </script>

</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CPH_Wide" runat="Server">
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="CPH_MainContent" runat="Server">
    <h1>Unprocessed Events</h1>
    <p class="info">
        These are the employee events that are older than 24 hours but have not been converted into shifts yet, usually meaning that those shifts seen here to be started were not closed yet. 
This could be legitimate (ex: an employee forgetting to punch out), but could also be the result of an error in communication with Field Force Manager, which although rare, can corrupt future time for that employee until adressed.<br />
        If you see an employee here with a shift open that you don't believe should be open, please alert Tina Merrifield of the situation.<br />
        For events that you know are valid (ex: employee stops working for Comnet, never punches out of last shift), please use the "Ignore" button to supress the warning on this page.
    </p>
    <asp:Literal ID="LIT_Report" runat="server"></asp:Literal>
</asp:Content>
