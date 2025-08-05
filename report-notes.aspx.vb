Imports System.Data.SqlClient

Partial Class reportnotes
    Inherits GlobalClass

    'Dim sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("Fullconnection").ConnectionString

    Sub page_load()

        If IsPostBack = False Then
            TXT_EndDate.Text = Now().Date
            TXT_StartDate.Text = DateAdd(DateInterval.Day, -7, Now().Date)
            DDL_User.SelectedValue = 0
        End If

        LIT_Results.Text = GenerateReport()

    End Sub

    Function GenerateReport() As String

        Dim StartDate As Date = TXT_StartDate.Text
        Dim EndDate As Date = TXT_EndDate.Text
        Dim UserPerformed As Integer = DDL_User.SelectedValue
        Dim PMSelected As String = DDL_PM.SelectedValue
        Dim ReturnString As String = ""
        Dim EmployeeID As Integer = 0, FirstName As String = "", LastName As String = "", Comments As String = "", startTime As Date, TimeSpent As Double, PM As String = "", Approved As Integer
        Dim EmployeeIDLoop As Integer = 0

        'Dim JobNumber As String
        Dim strResult As String = "<table class=""listing"" id=""ListEmployees"">"
        Dim strEmployeeList As String = ""
        strResult = strResult & "<tr><th>Date</th><th>PM</th><th>Hours</th><th>Comments</th><th>Approved</th></tr>"

        Dim sEmployeeQuery As String = ""

        If Not UserPerformed = 0 Then
            sEmployeeQuery = " and EmployeeID=@EmployeeID "
        End If

        If Not PMSelected = "" Then
            sEmployeeQuery = sEmployeeQuery & " and fPM=@PMSelected "
        End If



        If sEmployeeQuery = "" Then
            sEmployeeQuery = "select EmployeeID, FirstName, LastName, fComments, startTime, sum(DATEDIFF(mi,StartTime,EndTime)) as TimeSpent, fPM, Approved from tTimeEntry, tShifts, tEmployees where tTimeEntry.ShiftID = tShifts.ShiftID And tEmployees.EmployeeID = tShifts.UserPerformed and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME) >= @startTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME) <= @endTime and not fComments = '' group by tShifts.ShiftId, EmployeeId, FirstName, LastName, startTime, fComments, fPM, Approved order by LastName, FirstName, StartTime asc"
        Else
            sEmployeeQuery = "select EmployeeID, FirstName, LastName, fComments, startTime, sum(DATEDIFF(mi,StartTime,EndTime)) as TimeSpent, fPM, Approved from tTimeEntry, tShifts, tEmployees where tTimeEntry.ShiftID = tShifts.ShiftID And tEmployees.EmployeeID = tShifts.UserPerformed and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME) >= @startTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME) <= @endTime and not fComments = '' " & sEmployeeQuery & " group by tShifts.ShiftId, EmployeeId, FirstName, LastName, startTime, fComments, fPM, Approved order by LastName, FirstName, StartTime asc"
        End If

        Dim connEmployees As New SqlConnection(sConnection)
        Dim cmdEmplyees As New SqlCommand(sEmployeeQuery, connEmployees)
        Dim drEmployees As SqlDataReader
        cmdEmplyees.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdEmplyees.Parameters.Add(New SqlParameter("@EndTime", EndDate))
        cmdEmplyees.Parameters.Add(New SqlParameter("@EmployeeID", UserPerformed))
        cmdEmplyees.Parameters.Add(New SqlParameter("@PMSelected", PMSelected))

        Try
            connEmployees.Open()

            drEmployees = cmdEmplyees.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drEmployees.Read

                'JobNumber = drEmployees.Item("fJobNumber")

                EmployeeID = drEmployees.Item("EmployeeID")
                FirstName = drEmployees.Item("FirstName")
                LastName = drEmployees.Item("LastName")
                Comments = drEmployees.Item("fComments")
                startTime = drEmployees.Item("startTime")
                TimeSpent = drEmployees.Item("TimeSpent")
                PM = drEmployees.Item("fPM")
                Approved = drEmployees.Item("Approved")

                If Not EmployeeID = EmployeeIDLoop Then
                    strResult = strResult & "<tr><td colspan=5 style=""background-color:#d2cedf;""><a href=""worker.aspx?startdate=" & StartDate & "&enddate=" & EndDate & "&id=" & EmployeeID & """><strong>" & LastName & ", " & FirstName & "</strong></a></td></tr>"
                    EmployeeIDLoop = EmployeeID
                End If

                'strResult = strResult & "<tr><td>PM</td><td>Hours</td><td>Comments</td><td>Approved</td></tr>"
                strResult = strResult & "<tr style=""border-bottom: 1px solid #d2cedf;"">"
                strResult = strResult & "<td>" & startTime.Date & "</td>"
                strResult = strResult & "<td>" & UCase(PM) & "</td>"
                strResult = strResult & "<td>" & TimeSpent / 60 & "</td>"
                strResult = strResult & "<td>" & Comments & "</td>"
                strResult = strResult & "<td>"
                If Approved = 1 Then strResult = strResult & "<img src=""images/tick.png"">"
                strResult = strResult & "</td></tr>"

            End While

            strResult = strResult & "</table>"
        Catch ex As Exception
            LogError("report-notes.aspx.vb :: GenerateReport", ex.ToString)
            Response.Write("ERROR:" & ex.ToString & "<br>")
        Finally
            connEmployees.Close()
        End Try

        Return strResult

    End Function

End Class
