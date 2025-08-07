
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices.ComTypes

Partial Class report_timebyjob2_excel
    Inherits GlobalClass

    Sub page_load()
        Dim FileName As String = "AV-Dept-report.xls"

        Response.Clear()
        Response.AddHeader("content-disposition", "attachment;filename=" & FileName)
        Response.ContentType = "application/ms-excel"
        Response.ContentEncoding = System.Text.Encoding.UTF8

        Response.Write(GenerateReport())

    End Sub

    Function GenerateReport() As String

        Dim sEmployeeQuery As String = TryCast(Session("AVDeptReportQuery"), String)
        Dim StartDate As String
        Dim EndDate As String
        If Session("AVDeptReportStartTime") IsNot Nothing Then
            StartDate = Convert.ToDateTime(Session("AVDeptReportStartTime")).ToString("yyyy-MM-dd HH:mm:ss")
        End If

        If Session("AVDeptReportEndTime") IsNot Nothing Then
            EndDate = Convert.ToDateTime(Session("AVDeptReportEndTime")).ToString("yyyy-MM-dd HH:mm:ss")
        End If
        If String.IsNullOrEmpty(sEmployeeQuery) OrElse String.IsNullOrEmpty(StartDate) OrElse String.IsNullOrEmpty(EndDate) Then
            Return "ERROR: Missing session data (query or dates)."
        End If

        Session.Remove("AVDeptReportQuery")
        Session.Remove("AVDeptReportStartTime")
        Session.Remove("AVDeptReportEndTime")

        Dim strResult As String = "<table class=""listing"" style=""width:250px;"" id=""ListEmployees"">"
        strResult &= "<tr><th>Job Number</th><th>Employee</th><th>Standard Time</th><th>Overtime</th></tr>"

        Dim connEmployees As New SqlConnection(sConnection)
        Dim cmdEmplyees As New SqlCommand(sEmployeeQuery, connEmployees)
        cmdEmplyees.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdEmplyees.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Try
            connEmployees.Open()

            Using drEmployees As SqlDataReader = cmdEmplyees.ExecuteReader(CommandBehavior.CloseConnection)
                While drEmployees.Read()
                    Dim loopJobNumber As String = drEmployees("fJobNumber").ToString()
                    Dim loopEmployeeID As Integer = Convert.ToInt32(drEmployees("UserPerformed"))
                    Dim loopEmployeeFullName As String = drEmployees("EmployeeFullName").ToString()

                    Dim TotalHours As Double = GetTotalHours(loopJobNumber, StartDate, EndDate, loopEmployeeID)
                    Dim OTHours As Double = GetOvertimeByJob(loopJobNumber, StartDate, EndDate, loopEmployeeID)
                    Dim StandardHours As Double = TotalHours - OTHours

                    strResult &= "<tr>"
                    strResult &= "<td class=""jobnumber"">" & loopJobNumber & "</td>"
                    strResult &= "<td>" & loopEmployeeFullName & "</td>"
                    strResult &= "<td>" & StandardHours.ToString("F2") & "</td>"
                    strResult &= "<td>" & OTHours.ToString("F2") & "</td>"
                    strResult &= "</tr>"
                End While
            End Using

            strResult &= "</table>"

        Catch ex As Exception
            LogError("report-timebyjob2-excel.aspx.vb :: GenerateReport", ex.ToString())
            Response.Write("ERROR: " & Server.HtmlEncode(ex.ToString()) & "<br>")
        Finally
            If connEmployees.State <> ConnectionState.Closed Then
                connEmployees.Close()
            End If
        End Try

        Return strResult

    End Function

    'Function GenerateReport() As String

    '    Dim sEmployeeQuery As String = Session("AVDeptReportQuery")
    '    Dim StartDate As String = Session("AVDeptReportStartTime")
    '    Dim EndDate As String = Session("AVDeptReportEndTime")

    '    Session.Remove("AVDeptReportQuery")
    '    Session.Remove("AVDeptReportStartTime")
    '    Session.Remove("AVDeptReportEndTime")

    '    Dim ReturnString As String = ""
    '    Dim TotalHours As Double, OTHours As Double, StandardHours As Double

    '    Dim strResult As String = "<table class=""listing"" style=""width:250px;"" id=""ListEmployees"">"
    '    Dim strEmployeeList As String = ""
    '    strResult = strResult & "<tr><th>Job Number</th><th>Employee</th><th>Standard Time</th><th>Overtime</th></tr>"

    '    Dim sFilter As String = ""


    '    Dim connEmployees As New SqlConnection(sConnection)
    '    Dim cmdEmplyees As New SqlCommand(sEmployeeQuery, connEmployees)
    '    Dim drEmployees As SqlDataReader
    '    cmdEmplyees.Parameters.Add(New SqlParameter("@StartTime", StartDate))
    '    cmdEmplyees.Parameters.Add(New SqlParameter("@EndTime", EndDate))

    '    Try
    '        connEmployees.Open()

    '        drEmployees = cmdEmplyees.ExecuteReader(Data.CommandBehavior.CloseConnection)

    '        Dim loopJobNumber As String, loopEmployeeID As Integer, loopEmployeeFullName As String = ""

    '        While drEmployees.Read

    '            loopJobNumber = drEmployees.Item("fJobNumber")
    '            loopEmployeeID = drEmployees.Item("UserPerformed")
    '            loopEmployeeFullName = drEmployees.Item("EmployeeFullName")

    '            TotalHours = GetTotalHours(loopJobNumber, StartDate, EndDate, loopEmployeeID)
    '            OTHours = GetOvertimeByJob(loopJobNumber, StartDate, EndDate, loopEmployeeID)
    '            StandardHours = TotalHours - OTHours

    '            strResult = strResult & "<tr>"
    '            strResult = strResult & "<td class=""jobnumber"">" & loopJobNumber & "</td>"
    '            strResult = strResult & "<td>" & loopEmployeeFullName & "</td>"
    '            strResult = strResult & "<td>" & StandardHours & "</td>"
    '            strResult = strResult & "<td>" & OTHours & "</td>"
    '            strResult = strResult & "</tr>"

    '        End While

    '        strResult = strResult & "</table>"
    '    Catch ex As Exception
    '        LogError("report-timebyjob2-excel.aspx.vb :: GenerateReport", ex.ToString)
    '        Response.Write("ERROR:" & ex.ToString & "<br>")
    '    Finally
    '        connEmployees.Close()
    '    End Try

    '    Return strResult

    'End Function

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
