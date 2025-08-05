Imports System.Data.SqlClient

Partial Class report_timebyjob2
    Inherits GlobalClass

    'Dim sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("Fullconnection").ConnectionString


    Sub page_load()

        If IsPostBack = True Then
            LIT_Results.Text = GenerateReport()
        Else
            TXT_EndDate.Text = Now().Date
            TXT_StartDate.Text = DateAdd(DateInterval.Day, -7, Now().Date)
            TXT_JobNumber.Text = ""
            CHK_AVOnly.Checked = False
        End If

    End Sub

    Function GenerateReport() As String

        Dim StartDate As Date = TXT_StartDate.Text
        Dim EndDate As Date = TXT_EndDate.Text
        Dim JobNumber As String = TXT_JobNumber.Text
        Dim AVOnly As Boolean = CHK_AVOnly.Checked

        Dim ReturnString As String = ""
        Dim TotalHours As Double, OTHours As Double, StandardHours As Double

        Dim strResult As String = "<table class=""listing"" style=""width:250px;"" id=""ListEmployees"">"
        Dim strEmployeeList As String = ""
        strResult = strResult & "<tr><th>Job Number</th><th>Employee</th><th>Standard Time</th><th>Overtime</th></tr>"

        Dim sFilter As String = ""

        If JobNumber <> "" Then
            sFilter = sFilter & " and fJobNumber like '%" & JobNumber & "%'"
        End If

        If AVOnly = True Then
            sFilter = sFilter & " and EmployeeID in (select EmployeeID from tEmployees where Department = 'AV')"
        End If


        Dim sEmployeeQuery As String = "select distinct fJobNumber, UserPerformed, tEmployees.FirstName + ' ' + tEmployees.LastName as EmployeeFullName from tShifts, tEmployees where tShifts.UserPerformed = tEmployees.EmployeeID " & sFilter & " and ShiftID in (select distinct ShiftID from tTimeEntry where CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime) order by fJobNumber"

        Session("AVDeptReportQuery") = sEmployeeQuery
        Session("AVDeptReportStartTime") = StartDate
        Session("AVDeptReportEndTime") = EndDate

        Dim connEmployees As New SqlConnection(sConnection)
        Dim cmdEmplyees As New SqlCommand(sEmployeeQuery, connEmployees)
        Dim drEmployees As SqlDataReader
        cmdEmplyees.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdEmplyees.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Try
            connEmployees.Open()

            drEmployees = cmdEmplyees.ExecuteReader(Data.CommandBehavior.CloseConnection)

            Dim loopJobNumber As String, loopEmployeeID As Integer, loopEmployeeFullName As String = ""

            While drEmployees.Read

                loopJobNumber = drEmployees.Item("fJobNumber")
                loopEmployeeID = drEmployees.Item("UserPerformed")
                loopEmployeeFullName = drEmployees.Item("EmployeeFullName")

                TotalHours = GetTotalHours(loopJobNumber, StartDate, EndDate, loopEmployeeID)
                OTHours = GetOvertimeByJob(loopJobNumber, StartDate, EndDate, loopEmployeeID)
                StandardHours = TotalHours - OTHours

                strResult = strResult & "<tr>"
                strResult = strResult & "<td class=""jobnumber"">" & loopJobNumber & "</td>"
                strResult = strResult & "<td>" & loopEmployeeFullName & "</td>"
                strResult = strResult & "<td>" & StandardHours & "</td>"
                strResult = strResult & "<td>" & OTHours & "</td>"
                strResult = strResult & "</tr>"

            End While

            strResult = strResult & "</table>"
        Catch ex As Exception
            LogError("report-timebyjob2.aspx.vb :: GenerateReport", ex.ToString)
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
