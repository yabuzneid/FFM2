Imports System.Data.SqlClient

Partial Class report_timebyjob
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
        Dim ReturnString As String = ""
        Dim TotalHours As Double, OTHours As Double, StandardHours As Double

        Dim JobNumber As String
        Dim strResult As String = "<table class=""listing"" style=""width:250px;"" id=""ListEmployees"">"
        Dim strEmployeeList As String = ""
        strResult = strResult & "<tr><th>Job Number</th><th>Standard Time</th><th>Overtime</th></tr>"

        Dim sEmployeeQuery As String = ""

        If UserPerformed = 0 Then
            sEmployeeQuery = "select distinct fJobNumber from tShifts where ShiftID in (select distinct ShiftID from tTimeEntry where CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime) order by fJobNumber asc"
        Else
            sEmployeeQuery = "select distinct fJobNumber from tShifts where ShiftID in (select distinct ShiftID from tTimeEntry where CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime) and UserPerformed=@UserPerformed order by fJobNumber asc"
        End If



        Dim connEmployees As New SqlConnection(sConnection)
        Dim cmdEmplyees As New SqlCommand(sEmployeeQuery, connEmployees)
        Dim drEmployees As SqlDataReader
        cmdEmplyees.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdEmplyees.Parameters.Add(New SqlParameter("@EndTime", EndDate))
        cmdEmplyees.Parameters.Add(New SqlParameter("@UserPerformed", UserPerformed))

        Try
            connEmployees.Open()

            drEmployees = cmdEmplyees.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drEmployees.Read

                JobNumber = drEmployees.Item("fJobNumber")

                TotalHours = GetTotalHours(JobNumber, StartDate, EndDate, UserPerformed)
                OTHours = GetOvertimeByJob(JobNumber, StartDate, EndDate, UserPerformed)
                StandardHours = TotalHours - OTHours

                strResult = strResult & "<tr>"
                strResult = strResult & "<td class=""jobnumber""><a href=""jobreport.aspx?startdate=" & StartDate & "&enddate=" & EndDate & "&jobnumber=" & JobNumber & """>" & JobNumber & "</a></td>"
                strResult = strResult & "<td>" & StandardHours & "</td>"
                strResult = strResult & "<td>" & OTHours & "</td>"
                strResult = strResult & "</tr>"

            End While

            strResult = strResult & "</table>"
        Catch ex As Exception
            LogError("timesheets.aspx.vb :: GenerateProjectWeekTotals", ex.ToString)
            Response.Write("ERROR:" & ex.ToString & "<br>")
        Finally
            connEmployees.Close()
        End Try

        Return strResult



    End Function

    Function GetOvertimeByJob(ByVal JobNumber As String, ByVal StartDate As Date, ByVal EndDate As Date, UserPerformed As Integer) As Double

        Dim returnstring As String = "", TotalTime As Double

        Dim sTimeQuery As String = ""

        If UserPerformed = 0 Then
            sTimeQuery = "select isNull(sum(Overtime),0) as TotalOT from tTimeAdjustments where JobNumber = @JobNumber and CAST(FLOOR(CAST(DateAdjusted AS FLOAT))AS DATETIME)>=@StartDate and CAST(FLOOR(CAST(DateAdjusted AS FLOAT))AS DATETIME)<=@EndDate"
        Else
            sTimeQuery = "select isNull(sum(Overtime),0) as TotalOT from tTimeAdjustments where JobNumber = @JobNumber and CAST(FLOOR(CAST(DateAdjusted AS FLOAT))AS DATETIME)>=@StartDate and CAST(FLOOR(CAST(DateAdjusted AS FLOAT))AS DATETIME)<=@EndDate and EmployeeID=@UserPerformed"
        End If

        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndDate", EndDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@UserPerformed", UserPerformed))

        Dim drTimeQuery As SqlDataReader

        Try
            connTimeQuery.Open()

            drTimeQuery = cmdTimeQuery.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drTimeQuery.Read()
                TotalTime = drTimeQuery.Item("TotalOT")
            End While

            returnstring = TotalTime

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

        Return TotalTime


    End Function

    Function GetTotalHours(ByVal JobNumber As String, ByVal StartDate As Date, ByVal EndDate As Date, UserPerformed As Integer) As Double

        Dim returnstring As String = "", TotalTime As Double, StartTime As DateTime, EndTime As DateTime
        Dim TimeSpan As Double

        Dim sTimeQuery As String = ""

        If UserPerformed = 0 Then
            sTimeQuery = "select startTime, endTime from tTimeEntry, tShifts where tShifts.fJobNumber = @JobNumber and tShifts.ShiftID = tTimeEntry.ShiftID and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"
        Else
            sTimeQuery = "select startTime, endTime from tTimeEntry, tShifts where tShifts.fJobNumber = @JobNumber and tShifts.ShiftID = tTimeEntry.ShiftID and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime and UserPerformed=@UserPerformed"
        End If

        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndTime", EndDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@UserPerformed", UserPerformed))

        Dim drTimeQuery As SqlDataReader

        Try
            connTimeQuery.Open()

            drTimeQuery = cmdTimeQuery.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drTimeQuery.Read()

                StartTime = drTimeQuery.Item("StartTime")
                EndTime = drTimeQuery.Item("EndTime")

                TimeSpan = Math.Round(DateDiff(DateInterval.Minute, StartTime, EndTime) / 60, 2)
                TotalTime = TotalTime + TimeSpan

            End While

            'returnstring = TotalTime

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

        Return TotalTime

    End Function

End Class
