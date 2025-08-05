<%@ Page Language="vb" AutoEventWireup="true" CodeBehind="printreport.aspx.vb" Inherits="FFMWebApp.printreport" %>

<!DOCTYPE html>


<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">

    <style>
        table {
            width: 100%;
            border-collapse: collapse;
            font-family: Arial,helvetica,sans-serif;
        }

        td, th {
            text-align: left;
            border: 1px solid #000;
        }

        h3 {
            margin: 0;
        }

        table.info td {
            border: 0;
        }

        table.listing {
            margin-bottom: 20px;
            page-break-after: always;
        }

            table.listing td {
                font-size: 12px;
            }

        th.dayheading {
            color: #fff;
            background-color: #444;
        }

        tr.label td {
            font-weight: bold;
        }

        tr.shift td {
            background-color: #ddd;
        }

        img {
            border: 0;
            margin: 0 3px;
        }
    </style>
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
        <asp:Literal runat="server" ID="LIT_Page" />
    </form>
</body>
</html>
