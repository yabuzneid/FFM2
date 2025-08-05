Imports System.Data.SqlClient


Partial Class _Default
    Inherits System.Web.UI.Page

    Sub page_load()

        Dim qsAction As String = Request.QueryString("action")

        If qsAction = "logout" Then
            Session.Clear()
            LBL_Info.Visible = True
            LBL_Info.Text = "You have successfully logged out, have a nice day."
        ElseIf qsAction = "expired" Then
            LBL_Info.Visible = True
            LBL_Info.Text = "Your session has expired due to inactivity, please log back in."
        End If

    End Sub

    Sub ValidateLogin(ByVal sender As Object, ByVal e As EventArgs)

        Dim Username As String = TXT_Username.Text
        Dim Password As String = TXT_Password.Text
        Dim AllowLogin As Boolean = False, dbUserID As Integer, dbClerk As Integer, dbWorker As Integer, dbInitials As String = "", dbPM As String = "", dbOffice As String = "", dbLimitedView As String = ""

        Dim sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("FullConnection").ConnectionString
        Dim LoginQuery As String = "select password, EmployeeID, Clerk, PM, Initials, Worker, OfficeID, LimitedView from tEmployees where Username=@Username and Active=1 and not Password=''"
        Dim connLogin As New SqlConnection(sConnection)
        Dim cmdLogin As New SqlCommand(LoginQuery, connLogin)
        Dim drLogin As SqlDataReader
        cmdLogin.Parameters.Add(New SqlParameter("@Username", Username))

        Try
            connLogin.Open()
            drLogin = cmdLogin.ExecuteReader(Data.CommandBehavior.CloseConnection)

            If drLogin.HasRows = True Then
                While drLogin.Read
                    If drLogin("Password") = Password Then
                        AllowLogin = True
                        dbUserID = drLogin("EmployeeID")
                        dbClerk = drLogin("Clerk")
                        dbPM = drLogin("PM")
                        dbInitials = drLogin("Initials")
                        dbWorker = drLogin("Worker")
                        dbOffice = drLogin("OfficeID")
                        dbLimitedView = drLogin("LimitedView")
                    End If
                End While
            End If

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString)
        Finally
            connLogin.Close()
        End Try

        If AllowLogin = True Then
            Session("UserID") = dbUserID

            Dim LandingPage As String

            LandingPage = "entertime.aspx"
            If dbPM = 1 Then LandingPage = "timesheets.aspx"
            If dbClerk = 1 Then LandingPage = "timesheets.aspx"

            Session("PM") = dbPM
            Session("Clerk") = dbClerk
            Session("Initials") = dbInitials
            Session("Worker") = dbWorker
            Session("SelectedOffice") = 0 'dbOffice
            Session("ShowAsPM") = dbInitials
            Session("LimitedView") = dbLimitedView

            Response.Redirect(LandingPage)

        Else
            LBL_Info.Visible = True
            LBL_Info.Text = "Login does not match our database, please try again."

        End If

    End Sub

End Class
