<%@ Page Title="" Language="vb" AutoEventWireup="true" MasterPageFile="~/MasterPage.Master" CodeBehind="employees.aspx.vb" Inherits="FFMWebApp.employees" %>

<asp:Content ID="Content1" ContentPlaceHolderID="CPH_Head" runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="CPH_MainContent" runat="Server">

    <asp:Label ID="LBL_Info" runat="server" Visible="false" CssClass="infolabel"></asp:Label>

    <asp:Panel ID="Panel_employeelist" runat="server">
        <h1>Employees</h1>
        <asp:SqlDataSource ID="SDS_EmployeeList" runat="server"
            ConnectionString="<%$ConnectionStrings:FullConnection %>"
            SelectCommand="select EmployeeID, Username, FirstName, LastName, PayrollID, Active, OfficeName, Initials from tEmployees, tOffices where tOffices.OfficeID = tEmployees.OfficeID order by LastName"></asp:SqlDataSource>
        <div id="newemployee">
            <img src="images/icon_add.gif" />
            <strong><a href="employees.aspx?id=-1">Add Employee</a></strong>
        </div>
        <asp:GridView ID="GW_EmployeeList" runat="server" DataSourceID="SDS_EmployeeList" AutoGenerateColumns="false" CssClass="listing" Width="450" BorderWidth="0" RowStyle-BorderWidth="0">
            <Columns>
                <asp:TemplateField HeaderText="Username">
                    <ItemTemplate>
                        <a href="employees.aspx?id=<%#Eval("EmployeeID")%>"><%#Eval("Username")%></a>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Name">
                    <ItemTemplate>
                        <%#Eval("LastName")%>, <%#Eval("FirstName")%>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Initials">
                    <ItemTemplate>
                        <%#Eval("Initials")%>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Payroll ID">
                    <ItemTemplate>
                        <%#Eval("PayrollID")%>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Office">
                    <ItemTemplate>
                        <%#Eval("OfficeName")%>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:TemplateField HeaderText="Active">
                    <ItemTemplate>
                        <%#ConvertCheck(Eval("Active"))%>
                    </ItemTemplate>
                </asp:TemplateField>
            </Columns>
        </asp:GridView>

    </asp:Panel>

    <asp:Panel ID="Panel_EditUser" runat="server" Visible="false">


        <h1>
            <asp:Literal ID="LIT_Heading" runat="server" Text="Edit User"></asp:Literal></h1>
        <table class="editform">
            <asp:Panel ID="Panel_UsernameBox" runat="server" Visible="false">
                <tr>
                    <td class="label">Username</td>
                    <td>
                        <asp:TextBox ID="TXT_Username" runat="server"></asp:TextBox></td>
                    <td class="notes"></td>
                </tr>
            </asp:Panel>

            <asp:Panel ID="Panel_UsernameLabel" runat="server">
                <tr>
                    <td class="label">Username</td>
                    <td><strong>
                        <asp:Literal ID="LIT_Username" runat="server"></asp:Literal></strong></td>
                    <td class="notes"></td>
                </tr>
            </asp:Panel>
            <tr>
                <td class="label">Email</td>
                <td>
                    <asp:TextBox ID="TXT_Email" runat="server"></asp:TextBox></td>

            </tr>
            <tr>
                <td class="label">Password</td>
                <td>
                    <asp:TextBox ID="TXT_password" runat="server"></asp:TextBox></td>
                <td class="notes">(Leave blank to keep the same password, enter a new one to change.)</td>
            </tr>
            <tr>
                <td class="label">First Name</td>
                <td>
                    <asp:TextBox ID="TXT_firstname" runat="server"></asp:TextBox></td>
                <td class="notes"></td>
            </tr>
            <tr>
                <td class="label">Last Name</td>
                <td>
                    <asp:TextBox ID="TXT_LastName" runat="server"></asp:TextBox></td>
                <td class="notes"></td>
            </tr>
            <tr>
                <td class="label">Initials</td>
                <td>
                    <asp:TextBox ID="TXT_Initials" runat="server"></asp:TextBox></td>
                <td class="notes">(Initials used via FieldForce Manager, only necessary for PMs. No two PMs can have the same initials.)</td>
            </tr>
            <tr>
                <td class="label">Cell Phone #</td>
                <td>
                    <asp:TextBox ID="TXT_Cellphone" runat="server"></asp:TextBox></td>
                <td class="notes">(Phone # used for FieldForce Manager, can be left blank if user does not enter time via FFM. Only numbers.)</td>
            </tr>
            <tr>
                <td class="label">Office</td>
                <td>
                    <asp:SqlDataSource ID="SDS_Offices" runat="server" ConnectionString="<%$ConnectionStrings:FullConnection %>"
                        SelectCommand="Select OfficeID, OfficeName from tOffices order by OfficeName"></asp:SqlDataSource>
                    <asp:DropDownList ID="DDL_Office" runat="server" DataSourceID="SDS_Offices" DataTextField="OfficeName" DataValueField="OfficeID" />
                </td>
                <td class="notes"></td>
            </tr>

            <tr>
                <td class="label">Department</td>
                <td>
                    <asp:DropDownList ID="DDL_Department" runat="server">
                        <asp:ListItem Value="" Text=""></asp:ListItem>
                        <asp:ListItem Value="AV" Text="AV"></asp:ListItem>
                    </asp:DropDownList></td>
                <td class="notes"></td>
            </tr>

            <tr>
                <td class="label">Type</td>
                <td>
                    <asp:DropDownList ID="DDL_Type" runat="server">
                        <asp:ListItem Value="" Text=""></asp:ListItem>
                        <asp:ListItem Value="G" Text="G"></asp:ListItem>
                        <asp:ListItem Value="K" Text="K"></asp:ListItem>
                    </asp:DropDownList></td>
                <td class="notes"></td>
            </tr>

            <tr>
                <td class="label">Employee Group</td>
                <td>
                    <asp:DropDownList ID="DDL_EmployeeGroup" runat="server">
                        <asp:ListItem Value="" Text=""></asp:ListItem>
                        <asp:ListItem Value="GCI" Text="GCI"></asp:ListItem>
                        <asp:ListItem Value="KCI" Text="KCI"></asp:ListItem>
                    </asp:DropDownList></td>
                <td class="notes"></td>
            </tr>


            <tr>
                <td class="label">Worker</td>
                <td>
                    <asp:CheckBox ID="CHK_Worker" runat="server" />
                </td>
                <td class="notes">(People not assigned Worker status will not appear in Employee lists showing total time. They can still log in.)</td>
            </tr>
            <tr>
                <td class="label">Part time</td>
                <td>
                    <asp:CheckBox ID="CHK_PartTime" runat="server" />
                </td>
                <td class="notes"></td>
            </tr>
            <tr>
                <td class="label">Project Manager</td>
                <td>
                    <asp:CheckBox ID="CHK_PM" runat="server" />
                </td>
                <td class="notes">(Project Managers can see and approve anyone's time, as well as run reports and manage other users.)</td>
            </tr>
            <tr>
                <td class="label">Clerk</td>
                <td>
                    <asp:CheckBox ID="CHK_Clerk" runat="server" />
                </td>
                <td class="notes">(Clerks can access the payroll spreadsheet as well as run reports and manage other users.)</td>
            </tr>
            <tr>
                <td class="label">Active</td>
                <td>
                    <asp:CheckBox ID="CHK_Active" runat="server" />
                </td>
                <td class="notes">(Users not marked as Active cannot log into the system at all.)</td>
            </tr>
            <tr>
                <td class="label">Read Only</td>
                <td>
                    <asp:CheckBox ID="CHK_ReadOnlyAccess" runat="server" />
                </td>
                <td class="notes">(Users marked as Read Only can access the system but will not be able to add or modify any entered time.)</td>
            </tr>
            <tr>
                <td class="label">Limited View</td>
                <td>
                    <asp:CheckBox ID="CHK_LimitedView" runat="server" />
                </td>
                <td class="notes">(PMs with Limited View can only see timecards entered with that PM's initials.)</td>
            </tr>
            <tr>
                <td class="label">Per Diem Rate</td>
                <td>
                    <asp:TextBox ID="TXT_PerDiem" runat="server"></asp:TextBox></td>
                <td class="notes"></td>
            </tr>
            <!--
<tr><td class="label">Travel Multiplier Company</td><td><asp:TextBox ID="TXT_TravelCompany" runat="server"></asp:TextBox></td><td class="notes">(Rate for Code 18)</td></tr>
<tr><td class="label">Travel Multiplier Personal</td><td><asp:TextBox ID="TXT_TravelPersonal" runat="server"></asp:TextBox></td><td class="notes">(Rate for Code 19)</td></tr>
-->
            <tr>
                <td class="label">Payroll ID #</td>
                <td>
                    <asp:TextBox ID="TXT_PayrollID" runat="server"></asp:TextBox></td>
                <td class="notes"></td>
            </tr>
            <tr>
                <td colspan="3" class="actionbuttons">
                    <asp:Button ID="BTN_Submit" runat="server" Text="Save" OnClick="SaveEmployeeData" />
                    <asp:Button ID="BTN_Cancel" runat="server" Text="Cancel" OnClick="CancelForm" /></td>
            </tr>
        </table>

    </asp:Panel>

</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="CPH_Wide" runat="Server">
</asp:Content>

