Imports System.Data.SqlClient


Partial Class MasterPage
    Inherits System.Web.UI.MasterPage
    Public sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("FullConnection").ConnectionString


    Sub page_load()

        Dim sMainNav As String = ""

        If Session("PM") = 1 Or Session("Clerk") = 1 Then
            sMainNav = sMainNav & "<div id=""mainnav"">"
            sMainNav = sMainNav & "<div id=""nav-corner-l""></div><div id=""nav-corner-r"">"
            sMainNav = sMainNav & "<a href=""entertime.aspx"">Enter Time</a>"
            If Session("PM") = 1 Then sMainNav = sMainNav & " | <a href=""timesheets.aspx"">Approve Time</a>"
            If Session("Clerk") = 1 Then sMainNav = sMainNav & " | <a href=""landing.aspx"">Payroll</a>"
            If Session("PM") = 1 Or Session("Clerk") = 1 Then sMainNav = sMainNav & " | <a href=""reports.aspx"">Reports</a>"
            If (Session("PM") = 1 Or Session("Clerk") = 1) And Not Session("LimitedView") = 1 Then sMainNav = sMainNav & " | <a href=""employees.aspx"">Employees</a>"
            'sMainNav = sMainNav & " | <a href=""Default.aspx?action=logout"">Log Out</a>"
            sMainNav = sMainNav & "</div></div>"
            LIT_MainNav.Text = sMainNav
        Else
            LIT_MainNav.Text = "<div style=""height:30px;""></div>"
        End If

        LIT_CurrentUser.Text = GetCurrentUser()

    End Sub

    Function GetCurrentUser() As String

        Dim UserID As Integer = Session("UserID")

        Dim FirstName As String = "", LastName As String = "", OfficeName As String = ""

        Dim qUser As String = "select FirstName, LastName, OfficeName from tEmployees, tOffices where tOffices.OfficeID = tEmployees.OfficeID and EmployeeID=@EmployeeID"
        Dim connUser As New SqlConnection(sConnection)
        Dim cmdUser As New SqlCommand(qUser, connUser)
        Dim drUser As SqlDataReader
        cmdUser.Parameters.Add(New SqlParameter("@EmployeeID", UserID))

        Try
            connUser.Open()
            drUser = cmdUser.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drUser.Read
                FirstName = drUser.Item("FirstName")
                LastName = drUser.Item("LastName")
                OfficeName = drUser.Item("OfficeName")
            End While
        Catch ex As Exception
            'LogError("MasterPage.Master.vb :: GetCurrentUser", ex.ToString)
            Response.Write("ERROR: " & ex.ToString)
        Finally
            connUser.Close()
        End Try

        Return "<strong>" & FirstName & " " & LastName & "</strong>, " & OfficeName & " [<a href=""Default.aspx?action=logout"">log out</a>]"

    End Function
End Class

