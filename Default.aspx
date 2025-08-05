<%@ Page Language="VB" AutoEventWireup="true" CodeFile="Default.aspx.vb" 
Inherits="_Default" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Comnet Communications </title>
    <link href="outside.css" rel="Stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div id="heading"></div>


<asp:Label ID="LBL_Info" runat="server" CssClass="info" Visible="false"></asp:Label>

<h2>Log in</h2>
<table id="login">
<tr><td class="label">Username</td><td><asp:TextBox ID="TXT_Username" runat="server" style="width:160px"></asp:TextBox></td></tr>
<tr><td class="label">Password</td><td><asp:TextBox ID="TXT_Password" runat="server" TextMode="Password" style="width:160px"></asp:TextBox></td></tr>
<tr><td></td><td><asp:Button ID="BTN_Login" runat="server" Text="Log in" OnClick="ValidateLogin" /></td></tr>
</table>


</form>
</body>
</html>


