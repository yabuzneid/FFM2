Imports System.Data.SqlClient


Partial Class OpenShifts
    Inherits GlobalClass


    Sub page_load()

        If IsPostBack = True Then

            Dim ActionToIgnore As String = Request.Form("actiontoignore")

            IgnoreAction(ActionToIgnore)



        End If

        LIT_Report.Text = ShowUnprocessed()

    End Sub

    Sub IgnoreAction(ActionID As String)

        Dim sIgnoreQuery As String = "insert into tIgnoreUnprocessed (ActionID) values(@ActionID)"
        Dim connIgnore As New SqlConnection(sConnection)
        Dim cmdIgnore As New SqlCommand(sIgnoreQuery, connIgnore)
        Dim drIgnore As SqlDataReader
        cmdIgnore.Parameters.Add(New SqlParameter("@ActionID", ActionID))
        Try
            connIgnore.Open()
            drIgnore = cmdIgnore.ExecuteReader(Data.CommandBehavior.CloseConnection)



        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
            LogError("Openshifts.aspx :: IgnoreAction", ex.ToString)
        Finally
            connIgnore.Close()
        End Try



    End Sub

    Function ShowUnprocessed() As String

        Dim ReturnString As String = "<table class=""listing"">"
        ReturnString = ReturnString & "<tr><th>Time</th><th>Employee</th><th>Type&nbsp;of&nbsp;Event</th><th>Location</th><th>Job Number</th><th>Project</th><th>PM</th><th>Ignore Message</th></tr>"
        Dim EmployeeName As String = "", ActionEvent As String = "", ActionTime As DateTime, Address As String = "", JobNumber As String = "", ProjectName As String = "", PM As String = "", ActionID As Integer

        Dim sUnprocessedQuery As String = "select ActionID, LastName + ', ' + FirstName as EmployeeName, ActionEvent, ActionTime, Address, fJobNumber, fProjectName, fPM from tActions, tEmployees where tEmployees.CellPhone = tActions.UserPerformed And Processed = 0 And ActionTime < DateAdd(dd, -1, GETDATE()) and (ActionEvent = 'Start Shift' or ActionEvent = 'Start Break' or ActionEvent = 'End Break' or ActionEvent = 'End Shift') and not ActionID in (select ActionID from tIgnoreUnprocessed) order by ActionTime desc"
        Dim connUnprocessed As New SqlConnection(sConnection)
        Dim cmdUnprocessed As New SqlCommand(sUnprocessedQuery, connUnprocessed)
        Dim drUnprocessed As SqlDataReader
        Try
            connUnprocessed.Open()
            drUnprocessed = cmdUnprocessed.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drUnprocessed.Read

                EmployeeName = drUnprocessed.Item("EmployeeName")
                ActionEvent = drUnprocessed.Item("ActionEvent")
                ActionTime = drUnprocessed.Item("ActionTime")
                Address = drUnprocessed.Item("Address")
                JobNumber = drUnprocessed.Item("fJobNumber")
                ProjectName = drUnprocessed.Item("fProjectName")
                PM = drUnprocessed.Item("fPM")
                ActionID = drUnprocessed.Item("ActionID")

                ReturnString = ReturnString & "<tr><td>" & ActionTime.ToString & "</td><td>" & EmployeeName & "</td><td>" & ActionEvent & "</td><td>" & Address & "</td><td>" & JobNumber & "</td><td>" & ProjectName & "</td><td>" & UCase(PM) & "</td>"
                ReturnString = ReturnString & "<td>"
                ReturnString = ReturnString & "<input type=""submit"" value=""Ignore"" class=""IgnoreButton"" title=""" & ActionID & """ id=""ignore_" & ActionID & """ /></td></tr>"


            End While

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
            LogError("Openshifts.aspx :: ShowUnprocessed", ex.ToString)
        Finally
            connUnprocessed.Close()
        End Try

        ReturnString = ReturnString & "<input type=""hidden"" id=""actiontoignore"" name=""actiontoignore"" value=""""/>"
        ReturnString = ReturnString & "</table>"

        Return ReturnString

    End Function


End Class
