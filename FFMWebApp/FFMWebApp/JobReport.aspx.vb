Imports System.Data.SqlClient

Public Class JobReport
    Inherits GlobalClass


    Sub page_load()

        Dim StartDate As DateTime = Request.QueryString("startdate")
        Dim EndDate As DateTime = Request.QueryString("enddate")
        Dim JobNumber As String = Request.QueryString("jobnumber")

        If IsPostBack = True Then
            JobNumber = Request.Form("JumpToJob")
            StartDate = Request.Form("datejump_start")
            EndDate = Request.Form("datejump_end")
            Response.Redirect("jobreport.aspx?startdate=" & StartDate & "&enddate=" & EndDate & "&jobnumber=" & JobNumber)
        End If

        LIT_Report.Text = GenerateJobReport(StartDate, EndDate, JobNumber)

    End Sub


    Function GenerateJobReport(ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal JobNumber As String) As String

        Dim ReturnString As String = ""
        Dim EmployeeName As String = ""
        Dim TotalTime As Double = 0, OverTime As Double = 0, RegularTime As Double = 0, EmployeeID As Integer
        Dim SumRegular As Double = 0, SumOvertime As Double = 0
        Dim ApprovedBy As String


        Dim sJobTimeQuery As String = "select EmployeeID, FirstName, LastName, sum(DATEDIFF(mi,StartTime,EndTime)) as TimeSpent from tTimeEntry, tShifts, tEmployees where fJobNumber = @JobNumber and tTimeEntry.ShiftID = tShifts.ShiftID and tEmployees.EmployeeID = tShifts.UserPerformed and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME) >= @StartDate and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME) <= @EndDate and Approved = 1 group by EmployeeID, LastName, FirstName order by LastName"
        Dim connJobTime As New SqlConnection(sConnection)
        Dim cmdJobTime As New SqlCommand(sJobTimeQuery, connJobTime)
        cmdJobTime.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdJobTime.Parameters.Add(New SqlParameter("@EndDate", EndDate))
        cmdJobTime.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
        Dim drJobTime As SqlDataReader
        Try
            connJobTime.Open()
            drJobTime = cmdJobTime.ExecuteReader(Data.CommandBehavior.CloseConnection)

            ReturnString = ReturnString & "<h1>Job Report</h1>"
            ReturnString = ReturnString & "<h2 style=""text-align:left; margin-left:250px;"">Job #: " & JobNumber & "<br>"
            ReturnString = ReturnString & StartDate & " through " & EndDate & "<br>"
            ReturnString = ReturnString & "(approved time only)</h2>"

            ReturnString = ReturnString & "<table class=""listing"" style=""width:400px;"">"
            ReturnString = ReturnString & "<tr><th>Employee</th><th>Regular Time</th><th>Overtime</th><th>Approved By</th></tr>"

            While drJobTime.Read()

                EmployeeID = drJobTime("EmployeeID")
                EmployeeName = drJobTime("FirstName") & " " & drJobTime("LastName")
                TotalTime = drJobTime("TimeSpent") / 60

                OverTime = GetPeriodOvertime(StartDate, EndDate, JobNumber, EmployeeID)

                RegularTime = TotalTime - OverTime

                ApprovedBy = GetApprover(EmployeeID, StartDate, EndDate, JobNumber)


                TotalTime = drJobTime.Item("TimeSpent")
                TotalTime = drJobTime.Item("TimeSpent")
                TotalTime = drJobTime.Item("TimeSpent")


                ReturnString = ReturnString & "<tr>"
                ReturnString = ReturnString & "<td><a href=""worker.aspx?id=" & EmployeeID & "&startdate=" & StartDate & "&enddate=" & EndDate & """>" & EmployeeName & "</a></td>"
                ReturnString = ReturnString & "<td>" & RegularTime & "</td>"
                ReturnString = ReturnString & "<td>" & OverTime & "</td>"
                ReturnString = ReturnString & "<td>" & ApprovedBy & "</td>"
                ReturnString = ReturnString & "</tr>"

                SumRegular = SumRegular + RegularTime
                SumOvertime = SumOvertime + OverTime

            End While

            ReturnString = ReturnString & "<tr><td><strong>TOTAL</strong></td><td><strong>" & SumRegular & "</strong></td><td><strong>" & SumOvertime & "</strong></td><td></td></tr></table>"

            ReturnString = ReturnString & "<form name=""JobForm"">"
            ReturnString = ReturnString & "<table  class=""listing"" style=""width:400px;"">"
            ReturnString = ReturnString & "<tr><td>JUMP TO REPORT</td><td colspan=3>" & GenerateJobDropdown(StartDate, EndDate, JobNumber) & "</td></tr>"
            ReturnString = ReturnString & "<tr><td>JUMP TO PERIOD</td><td><input type=""text"" name=""datejump_start"" class=""pickdate"" style=""width:80px;"" value=""" & StartDate & """ /></td>"
            ReturnString = ReturnString & "<td colspan=""2""><input type=""text"" name=""datejump_end"" class=""pickdate""  style=""width:80px;"" value=""" & EndDate & """ /></td></tr>"
            ReturnString = ReturnString & "<tr><td colspan=""4""><input type=""submit"" value=""Go"" /></td></tr>"
            ReturnString = ReturnString & "</form>"
            ReturnString = ReturnString & "</table>"
            ReturnString = ReturnString & "<strong><a style=""font-size:14px;"" href=""timesheets.aspx""><< Back to listing</a><strong>"


        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
            LogError("JobReport.aspx.vb", ex.ToString)
        Finally
            connJobTime.Close()
        End Try





        Return ReturnString

    End Function

    Function GenerateJobDropdown(StartDate As DateTime, EndDate As DateTime, CurrentJobNumber As String) As String

        Dim ReturnString As String = ""
        Dim JobNumber As String = ""

        Dim sEmployeeQuery As String = "select distinct fJobNumber, sum(DateDiff(n,startTime, endTime)) as TotalTime from tShifts, tTimeEntry where tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.ShiftID in (select distinct ShiftID from tTimeEntry where CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime) and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime group by fJobNumber order by fJobNumber"
        Dim connEmployees As New SqlConnection(sConnection)
        Dim cmdEmplyees As New SqlCommand(sEmployeeQuery, connEmployees)
        Dim drEmployees As SqlDataReader
        cmdEmplyees.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdEmplyees.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Try
            connEmployees.Open()

            drEmployees = cmdEmplyees.ExecuteReader(Data.CommandBehavior.CloseConnection)

            'ReturnString = ReturnString & "<form name=""JobForm"">"
            ReturnString = ReturnString & "<select name=""JumpToJob"">"

            While drEmployees.Read

                JobNumber = drEmployees.Item("fJobNumber")
                'TotalJobTime = drEmployees.Item("TotalTime") / 60

                ReturnString = ReturnString & "<option value=""" & JobNumber & """"

                If JobNumber = CurrentJobNumber Then ReturnString = ReturnString & " selected"

                ReturnString = ReturnString & ">" & JobNumber & "</option>"

            End While

            ReturnString = ReturnString & "</select>"
            'ReturnString = ReturnString & "<input type=""submit"" value=""Go"">"
            'ReturnString = ReturnString & "</form>"

        Catch ex As Exception
            LogError("jobreport.aspx.vb :: GenerateJobDropdown", ex.ToString)
            Response.Write("ERROR:" & ex.ToString & "<br>")
        Finally
            connEmployees.Close()
        End Try

        Return ReturnString

    End Function

    Function GetPeriodOvertime(ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal JobNumber As String, ByVal EmployeeID As Integer) As Double

        Dim Overtime As Double = 0

        Dim sOverTimeQuery As String = "select isNull(sum(Overtime),0) as SumOvertime from tTimeAdjustments where CAST(FLOOR(CAST(DateAdjusted AS FLOAT))AS DATETIME) >= @StartDate and CAST(FLOOR(CAST(DateAdjusted AS FLOAT))AS DATETIME) <= @EndDate and EmployeeID = @EmployeeID and JobNumber = @JobNumber"
        Dim connOverTime As New SqlConnection(sConnection)
        Dim cmdOverTime As New SqlCommand(sOverTimeQuery, connOverTime)
        cmdOverTime.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdOverTime.Parameters.Add(New SqlParameter("@EndDate", EndDate))
        cmdOverTime.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
        cmdOverTime.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))

        Dim drOverTime As SqlDataReader
        Try

            connOverTime.Open()
            drOverTime = cmdOverTime.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drOverTime.Read()

                Overtime = drOverTime.Item("SumOvertime")

            End While


        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
            LogError("JobReport.aspx.vb", ex.ToString)
        Finally
            connOverTime.Close()
        End Try

        Return Overtime

    End Function

    Function GetApprover(ByVal EmployeeID As Integer, ByVal StartDate As Date, ByVal EndDate As Date, ByVal JobNumber As String) As String

        Dim returnString As String = ""

        Dim sJobTimeQuery As String = "Select distinct Initials from tEmployees, tTimeEntry, tShifts where tEmployees.EmployeeID = tTimeEntry.ApprovedBy and tShifts.ShiftID = tTimeEntry.ShiftID and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME) >= @StartDate and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME) <= @EndDate and tShifts.fJobNumber = @JobNumber and tShifts.UserPerformed = @EmployeeID"
        Dim connJobTime As New SqlConnection(sConnection)
        Dim cmdJobTime As New SqlCommand(sJobTimeQuery, connJobTime)
        cmdJobTime.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
        cmdJobTime.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdJobTime.Parameters.Add(New SqlParameter("@EndDate", EndDate))
        cmdJobTime.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
        Dim drJobTime As SqlDataReader
        Try
            connJobTime.Open()

            drJobTime = cmdJobTime.ExecuteReader()

            While drJobTime.Read()

                If returnString = "" Then
                    returnString = drJobTime.Item("Initials")
                Else
                    returnString = returnString & ", " & drJobTime.Item("Initials")
                End If

            End While

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
            LogError("JobReport.aspx.vb :: GetApprover", ex.ToString)
        Finally
            connJobTime.Close()
        End Try

        GetApprover = returnString


    End Function

End Class