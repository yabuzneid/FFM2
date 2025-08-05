<%@ Page Title="" Language="vb" AutoEventWireup="true" MasterPageFile="~/MasterPage.Master" CodeBehind="generatecsv.aspx.vb" Inherits="FFMWebApp.generatecsv" %>
 
<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" Runat="Server">
<script src="scripts/jquery.tablesorter.js" type="text/javascript" language="javascript"></script>
<script>
    $(document).ready(function () {
        $("div#holder").css("display", "none");

        //validate form
        $("input#ctl00_CPH_Wide_BTN_Submit").click(function () {
            var ShowAlert = false;

            $("input.payidoverride").each(function (n, element) {
                if ($(this).val() == "") {
                    ShowAlert = true;
                    $(this).parent().parent().addClass("error");
                }
            });

            if (ShowAlert == true) {
                alert("You have not filled in all the required fields. \n Please fill in all the fields marked in red.");
                scroll(0, 0);
                return false;

            } else {
                document.aspnetForm.submit();
            }

        });

        $("table#generatecsv").tablesorter({
            headers: {
                5: { sorter: false },
                6: { sorter: false },
                7: { sorter: false },
                9: { sorter: false }
            }
        });

        $("input.payidoverride").change(function () {
            $(this).parent().parent().removeClass("error");
            $(this).parent().parent().removeClass("highlight");
            $(this).parent().parent().addClass("rowchanged");
        });

        //keep session alive
        setTimeout("callserver()", 6000);

    });


    function callserver() {
        var remoteURL = 'emptypage.aspx';
        $.get(remoteURL, function (data) { setTimeout("callserver()", 60000); });
    }

</script>
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="CPH_Wide" Runat="Server">
<h1>Export for Timberline</h1>

<p class="info" style="width:800px; text-align:left; margin:0 auto;">
<strong>Pay IDs that are pre-filled based on the type of time entered:</strong><br />
1 - Standard Time<br />
2 - Overtime (time and a half)<br />
11 - Per Diem<br />
18 - Travel/Company vehicle<br />
19 - Travel/Personal vehicle<br />
</p>
    <asp:Literal ID="LIT_Grid" runat="server"></asp:Literal>
    <asp:Button runat="server" ID="BTN_Submit" Text="Generate CSV" />
    <img src="/images/blank.gif" id="keepAliveIMG" width="0" height="0" />
</asp:Content>
