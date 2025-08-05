<%@ Page Title="" Language="VB" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="addtime.aspx.vb" Inherits="addtime" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" Runat="Server">
<link rel="stylesheet" type="text/css" href="scripts/datePicker.css" />
<script src="scripts/jquery.datePicker.js" type="text/javascript"></script>
<script src="scripts/reports.js" type="text/javascript"></script>
<script type="text/javascript" src="scripts/date.js"></script>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CPH_MainContent" Runat="Server">

<asp:ObjectDataSource ID="ODS_TimePicker" runat="server" TypeName="GlobalClass" SelectMethod="PickTime" />


<h1>Add Shift</h1>

<asp:Label ID="LBL_Info" Visible="false" runat="server" CssClass="infolabel"></asp:Label>


<table class="editform">

<asp:Panel ID="Panel_Employees" runat="server" Visible="false">
    <tr><td class="label">Employee</td>
    <td>
        <asp:ObjectDataSource ID="ODS_Employee" runat="server" TypeName="GlobalClass" SelectMethod="GetEmployeeList" />
        <asp:DropDownList ID="DDL_Employee" runat="server" DataSourceID="ODS_Employee" DataTextField="FullName" DataValueField="EmployeeID" AppendDataBoundItems="true">
        <asp:ListItem Text="- ALL EMPLOYEES -" Value="-1"></asp:ListItem>
        </asp:DropDownList>
    </td>
    <td class="notes"></td></tr>
    <tr>
    <td></td>
    <td>
        <asp:RadioButton ID="RB_SingleDay" GroupName="TimeType" runat="server" Checked="true" OnCheckedChanged="SwitchTimeType" AutoPostBack="true" />Single Shift
        <asp:RadioButton ID="RB_MultiDay" GroupName="TimeType" runat="server" OnCheckedChanged="SwitchTimeType" AutoPostBack="true" />Multiple Shifts
    </td>
    <td></td>
    </tr>
</asp:Panel>

<asp:Panel ID="Panel_SingleDay" runat="server" Visible="true">
    <tr><td class="label">Start Time</td>
    <td>
    <asp:TextBox ID="TXT_StartDate" runat="server" CssClass="pickdate" Width="80px"></asp:TextBox>
    <asp:DropDownList ID="DDL_StartTime" runat="server" DataSourceID="ODS_TimePicker" />
    </td>
    <td class="notes"></td></tr>

    <tr><td class="label">End Time</td>
    <td><asp:TextBox ID="TXT_EndDate" runat="server" CssClass="pickdate" Width="80px"></asp:TextBox>
    <asp:DropDownList ID="DDL_EndTime" runat="server" DataSourceID="ODS_TimePicker" />
    </td>
    <td class="notes"></td></tr>
</asp:Panel>

<asp:Panel ID="Panel_MultiDay" runat="server" Visible="false">
    <tr><td class="label">Start Date</td>
    <td>
    <asp:TextBox ID="TXT_StartDateMulti" runat="server" CssClass="pickdate" Width="80px"></asp:TextBox>
    </td>
    <td class="notes">(Weekends will be skipped automatically.)</td></tr>

    <tr><td class="label">End Date</td>
    <td><asp:TextBox ID="TXT_EndDateMulti" runat="server" CssClass="pickdate" Width="80px"></asp:TextBox>
    </td>
    <td class="notes"></td></tr>
    
    <tr>
        <td class="label">From</td>
        <td>
        <asp:DropDownList ID="DDL_StartTimeMulti" runat="server" DataSourceID="ODS_TimePicker" />
         To <asp:DropDownList ID="DDL_EndTimeMulti" runat="server" DataSourceID="ODS_TimePicker" />
        </td>
        <td></td>
    </tr>
</asp:Panel>

<tr><td class="label">Project Manager</td>
<td>
<asp:SqlDataSource ID="SDS_PMs" runat="server" ConnectionString="<%$ConnectionStrings:FullConnection %>"
SelectCommand="Select Initials from tEmployees where PM=1 and not Initials = '' and Active = 1 order by Initials"></asp:SqlDataSource>
<asp:DropDownList ID="DDL_PM" runat="server" DataSourceID="SDS_PMs" DataTextField="Initials" DataValueField="Initials" />

</td>
<td class="notes"></td></tr>

<tr><td class="label">State</td>
<td>
<asp:ObjectDataSource ID="ODS_ListStates" runat="server" TypeName="GlobalClass" SelectMethod="GetStates" />
<asp:DropDownList ID="DDL_State" runat="server" DataSourceID="ODS_ListStates" />

</td>
<td class="notes"></td></tr>

<tr><td class="label">Job #</td>
<td><asp:TextBox ID="TXT_JobNumber" runat="server"></asp:TextBox></td>
<td class="notes">(example: 01-03-001234)</td></tr>

<tr><td class="label">Prevailing Wage</td>
<td><asp:DropDownList ID="DDL_PW" runat="server">
    <asp:ListItem Text="" Value="" />
    <asp:ListItem Text="Yes" Value="Y" />
    <asp:ListItem Text="No" Value="N" />
</asp:DropDownList></td><td class="notes"></td></tr>

<tr><td class="label">Project</td>
<td><asp:TextBox ID="TXT_Project" runat="server"></asp:TextBox></td>
<td class="notes"></td></tr>


<tr><td class="label">Shift Type</td>
<td>
<asp:DropDownList ID="DDL_ShiftType" runat="server">
    <asp:ListItem Text="" Value="" />
    <asp:ListItem Text="Regular Shift" Value="Regular Shift" Selected="True" />
    <asp:ListItem Text="Travel" Value="Travel" />
    <asp:ListItem Text="Holiday" Value="Holiday" />
    <asp:ListItem Text="Personal" Value="Personal" />
    <asp:ListItem Text="Vacation" Value="Vacation" />
    <asp:ListItem Text="Sick" Value="Sick" />
    <asp:ListItem Text="PTO" Value="PTO" />
    <asp:ListItem Text="FMLA" Value="FMLA" />
    <asp:ListItem Text="FFCRA" Value="FFCRA" />
    <asp:ListItem Text="Bereavement" Value="Bereavement" />
</asp:DropDownList>
</td>
<td class="notes"></td></tr>
<!--
<tr><td class="label">Travel</td>
<td><asp:DropDownList ID="DDL_Travel" runat="server">
    <asp:ListItem Text="" Value="" />
    <asp:ListItem Text="Company" Value="Company" />
    <asp:ListItem Text="Personal" Value="Personal" />
</asp:DropDownList></td>
<td class="notes"></td></tr>
-->
<tr><td class="label">Per Diem</td>
<td><asp:DropDownList ID="DDL_PerDiem" runat="server">
    <asp:ListItem Text="" Value="" />
    <asp:ListItem Text="Yes" Value="Yes" />
    <asp:ListItem Text="No" Value="No" />
</asp:DropDownList></td>
<td class="notes"></td></tr>

<tr><td class="label">Injured</td>
<td><asp:DropDownList ID="DDL_Injured" runat="server">
    <asp:ListItem Text="" Value="" />
    <asp:ListItem Text="Yes" Value="Yes" />
    <asp:ListItem Text="No" Value="No" Selected="True" />
</asp:DropDownList></td><td class="notes"></td></tr>

<tr><td class="label">Comments</td>
<td colspan=2><asp:TextBox ID="TXT_Comments" runat="server" TextMode="MultiLine" Width="250px" Height="50px"></asp:TextBox></td>
</tr>

<tr>
<td colspan="3" class="actionbuttons">
<asp:Button ID="BTN_Submit1" Text="Save" runat="server" OnClick="SubmitShift" OnClientClick="return ValidateTimeEntry();"  /> 
<asp:Button ID="BTN_Cancel" Text="Cancel" runat="server" OnClick="CancelShift" />
</td></tr>

</table>

</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="CPH_Wide" Runat="Server">
</asp:Content>

