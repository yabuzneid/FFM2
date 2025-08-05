Imports System.Data.SqlClient


Partial Class landing
    Inherits GlobalClass


    Sub page_load()

        If IsPostBack = False Then
            MakeDateDropdown()
        End If

        Dim sDateRange As String = DDL_DateRange.SelectedValue
        Dim arrDateRange As Array = sDateRange.Split("-")
        Dim StartDate As Date = arrDateRange(0)
        Dim EndDate As Date = arrDateRange(1)

        LIT_Main.Text = PMList(StartDate, EndDate)
        LIT_OfficeTotals.Text = GetOfficeTotals(StartDate, EndDate)
        LIT_UnprocessedWarnings.Text = CheckUnprocessed()

        ToggleButton(StartDate, EndDate)


    End Sub

    Function PMList(ByVal StartDate As Date, ByVal EndDate As Date) As String

        Dim PMListInitials As String = ""

        Dim strResult As New System.Text.StringBuilder()

        strResult.Append("<table class=""listing"" id=""ListPMs"">")
        strResult.Append("<tr><th>Project Manager</th><th>Total Time</th><th>Left Unapproved</th></tr>")

        Dim FirstName As String = "", LastName As String = "", Initials As String = ""
        Dim PMQuery As String = "select FirstName, LastName, Initials from tEmployees where Active = 1 and PM = 1 and not Initials = '' order by LastName asc"
        Dim connPM As New SqlConnection(sConnection)
        Dim cmdPM As New SqlCommand(PMQuery, connPM)
        Dim drPM As SqlDataReader

        Try
            connPM.Open()
            drPM = cmdPM.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drPM.Read()

                FirstName = drPM.Item("FirstName")
                LastName = drPM.Item("LastName")
                Initials = drPM.Item("Initials")

                strResult.Append("<tr><td>" & LastName & ", " & FirstName & "</td>")
                strResult.Append(GetPMTime(StartDate, EndDate, Initials))
                strResult.Append("</tr>")

                If PMListInitials = "" Then
                    PMListInitials = "'" & Initials & "'"
                Else
                    PMListInitials = PMListInitials & ",'" & Initials & "'"
                End If

            End While

        Catch ex As Exception
            LogError("landing.aspx.vb :: PMList", ex.ToString)
            Response.Write("ERROR: " & ex.ToString)
        Finally
            connPM.Close()
        End Try


        strResult.Append("<tr><td>(no PM specified)</td>" & GetPMTime(StartDate, EndDate, "") & "</tr>")
        strResult.Append("<tr><td>(PMs not in list)</td>" & GetInvalidPM(StartDate, EndDate, PMListInitials) & "</tr>")
        strResult.Append("</table>")

        PMList = strResult.ToString()

    End Function

    Sub MakeDateDropdown()

        'Dim returnstring As String = ""
        Dim EarliestRecord As DateTime

        Dim sStartTimeQuery As String = "select top 1 startTime as EarliestStartTime from tTimeEntry where startTime > '2009-11-1' order by startTime asc"
        Dim connStartTime As New SqlConnection(sConnection)
        Dim cmdStartTime As New SqlCommand(sStartTimeQuery, connStartTime)
        Dim drStartTime As SqlDataReader
        Try
            connStartTime.Open()
            drStartTime = cmdStartTime.ExecuteReader(Data.CommandBehavior.CloseConnection)

            If drStartTime.HasRows = True Then
                While drStartTime.Read
                    EarliestRecord = drStartTime.Item("EarliestStartTime")
                End While
            Else
                EarliestRecord = "11/1/2009"
            End If

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connStartTime.Close()
        End Try

        Dim DaysToStartOfWeek As Integer = EarliestRecord.Date.DayOfWeek + 2

        If DaysToStartOfWeek = 0 Then DaysToStartOfWeek = 7
        DaysToStartOfWeek = DaysToStartOfWeek - 1

        Dim Today As DateTime = Now()

        Dim StartFirstWeek As DateTime = DateAdd(DateInterval.Day, -DaysToStartOfWeek, EarliestRecord)

        Dim WeekStart As Date = StartFirstWeek
        Dim WeekEnd As Date = StartFirstWeek.AddDays(6)
        Dim EndWeekDiff As Integer


        While WeekStart.Date <= Now().Date

            EndWeekDiff = DateDiff(DateInterval.Day, Now().Date, WeekEnd.Date)

            DDL_DateRange.Items.Add(FormatDateTime(WeekStart, DateFormat.LongDate) & " - " & FormatDateTime(WeekEnd, DateFormat.LongDate))

            If 0 <= EndWeekDiff And EndWeekDiff < 7 Then
                DDL_DateRange.SelectedValue = FormatDateTime(WeekStart, DateFormat.LongDate) & " - " & FormatDateTime(WeekEnd, DateFormat.LongDate)
            End If

            If WeekEnd.Date = "12/24/2010" Then
                WeekStart = WeekStart.AddDays(7)
                WeekEnd = WeekEnd.AddDays(2)
            ElseIf WeekEnd.Date = "12/26/2010" Then
                WeekStart = WeekStart.AddDays(2)
                WeekEnd = WeekEnd.AddDays(7)
            Else
                WeekStart = WeekStart.AddDays(7)
                WeekEnd = WeekEnd.AddDays(7)
            End If
        End While

        DDL_DateRange.AutoPostBack = True

    End Sub

    Function GetPMTime(ByVal StartDate As Date, ByVal EndDate As Date, ByVal PMInitials As String) As String

        Dim returnstring As String = "", TotalTime As Double, Unapproved As Double
        Dim TimeSpan As Double

        Dim Approved As Integer, StartTime As DateTime, EndTime As DateTime

        Dim sTimeQuery As String = "select startTime, endTime, Approved from tTimeEntry, tShifts where tShifts.fPM = @fPM and tShifts.ShiftID = tTimeEntry.ShiftID and not fInjured is null and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@fPM", PMInitials))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndTime", EndDate))
        Dim drTimeQuery As SqlDataReader

        Try
            connTimeQuery.Open()

            drTimeQuery = cmdTimeQuery.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drTimeQuery.Read()

                StartTime = drTimeQuery.Item("StartTime")
                EndTime = drTimeQuery.Item("EndTime")
                Approved = drTimeQuery.Item("Approved")

                TimeSpan = Math.Round(DateDiff(DateInterval.Minute, StartTime, EndTime) / 60, 2)
                TotalTime = TotalTime + TimeSpan
                If Approved <> 1 Then Unapproved = Unapproved + TimeSpan

            End While

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

        returnstring = returnstring & "<td>" & TotalTime & "</td>"
        returnstring = returnstring & "<td>" & Unapproved & "</td>"

        GetPMTime = returnstring

    End Function

    Function GetInvalidPM(ByVal StartDate As Date, ByVal endDate As Date, ByVal PMList As String) As String

        Dim returnstring As String = "", TotalTime As Double, Unapproved As Double
        Dim TimeSpan As Double

        Dim Approved As Integer, StartTime As DateTime, EndTime As DateTime

        Dim sTimeQuery As String = "select startTime, endTime, Approved from tTimeEntry, tShifts where NOT LTRIM(RTRIM(tShifts.fPM)) in (" & PMList & ",'') and tShifts.ShiftID = tTimeEntry.ShiftID and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        'cmdTimeQuery.Parameters.Add(New SqlParameter("@PMList", PMList))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndTime", endDate))
        Dim drTimeQuery As SqlDataReader
        Try
            connTimeQuery.Open()

            drTimeQuery = cmdTimeQuery.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drTimeQuery.Read()

                StartTime = drTimeQuery.Item("StartTime")
                EndTime = drTimeQuery.Item("EndTime")
                Approved = drTimeQuery.Item("Approved")

                TimeSpan = Math.Round(DateDiff(DateInterval.Minute, StartTime, EndTime) / 60, 2)
                TotalTime = TotalTime + TimeSpan
                If Approved <> 1 Then Unapproved = Unapproved + TimeSpan

            End While

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

        returnstring = returnstring & "<td>" & TotalTime & "</td>"
        returnstring = returnstring & "<td>" & Unapproved & "</td>"

        GetInvalidPM = returnstring


    End Function

    Sub OpenPayroll(ByVal Sender As Object, ByVal e As EventArgs)

        Dim sDateRange As String = DDL_DateRange.SelectedValue
        Dim arrDateRange As Array = sDateRange.Split("-")
        Dim StartDate As Date = arrDateRange(0)
        Dim EndDate As Date = arrDateRange(1)
        Dim OfficeID As String = DDL_Office.SelectedValue()

        Response.Redirect("payroll.aspx?startdate=" & StartDate & "&enddate=" & EndDate & "&officeID=" & OfficeID)

    End Sub

    Sub LockWeek(ByVal sender As Object, ByVal e As EventArgs)

        Dim WeekExists As Boolean = False

        Dim sDateRange As String = DDL_DateRange.SelectedValue
        Dim arrDateRange As Array = sDateRange.Split("-")
        Dim StartDate As Date = arrDateRange(0)
        Dim EndDate As Date = arrDateRange(1)

        Dim qCheck As String = "select Locked from tLockedWeeks where StartDate=@StartDate and EndDate=@EndDate"
        Dim connCheck As New SqlConnection(sConnection)
        Dim cmdCheck As New SqlCommand(qCheck, connCheck)
        Dim drCheck As SqlDataReader
        cmdCheck.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdCheck.Parameters.Add(New SqlParameter("@EndDate", EndDate))

        Try
            connCheck.Open()
            drCheck = cmdCheck.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheck.Read
                If drCheck.HasRows = True Then
                    WeekExists = True
                End If
            End While
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("landing.aspx.vb :: LockWeek", ex.ToString)
        Finally
            connCheck.Close()
        End Try

        Dim qLockWeek As String

        If WeekExists = True Then
            qLockWeek = "update tLockedWeeks set Locked = 1 where StartDate=@StartDate and EndDate=@EndDate"
        Else
            qLockWeek = "insert into tLockedWeeks(StartDate,EndDate,Locked) values (@StartDate,@EndDate,1)"
        End If

        Dim connLock As New SqlConnection(sConnection)
        Dim cmdLock As New SqlCommand(qLockWeek, connLock)
        cmdLock.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdLock.Parameters.Add(New SqlParameter("@EndDate", EndDate))

        Try
            connLock.Open()
            cmdLock.ExecuteNonQuery()

        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("landing.aspx.vb :: LockWeek2", ex.ToString)
        Finally
            connLock.Close()
        End Try

        ToggleButton(StartDate, EndDate)

    End Sub

    Sub UnlockWeek(ByVal sender As Object, ByVal e As EventArgs)

        Dim sDateRange As String = DDL_DateRange.SelectedValue
        Dim arrDateRange As Array = sDateRange.Split("-")
        Dim StartDate As Date = arrDateRange(0)
        Dim EndDate As Date = arrDateRange(1)

        Dim qUnlockWeek As String = "update tLockedWeeks set Locked=0 where startDate=@StartDate and EndDate=@EndDate"
        Dim connUnlock As New SqlConnection(sConnection)
        Dim cmdUnlock As New SqlCommand(qUnlockWeek, connUnlock)
        cmdUnlock.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdUnlock.Parameters.Add(New SqlParameter("@EndDate", EndDate))

        Try
            connUnlock.Open()
            cmdUnlock.ExecuteNonQuery()

        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("landing.aspx.vb :: UnlockWeek", ex.ToString)
        Finally
            connUnlock.Close()
        End Try

        ToggleButton(StartDate, EndDate)

    End Sub

    Function GetOfficeTotals(ByVal StartDate As Date, ByVal EndDate As Date) As String

        Dim OfficeID As Integer, OfficeName As String, TotalOfficeTime As Double, OfficeOvertime As Double

        Dim returnstring As String = "<table class=""listing"">"
        returnstring = returnstring & "<tr><th>Office</th><th>Total Time</th><th>Overtime</th></tr>"

        Dim qOffices As String = "Select OfficeID, OfficeName from tOffices order by OfficeName"
        Dim connOffices As New SqlConnection(sConnection)
        Dim cmdOffices As New SqlCommand(qOffices, connOffices)
        Dim drOffices As SqlDataReader

        Try
            connOffices.Open()
            drOffices = cmdOffices.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drOffices.Read()

                OfficeID = drOffices.Item("OfficeID")
                OfficeName = drOffices.Item("OfficeName")


                Dim qTotalOfficeTime As String = "select IsNull(Sum(DateDiff(mi,starttime,endtime)),0) as TotalTime from tTimeEntry, tShifts, tEmployees where  tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.UserPerformed = tEmployees.EmployeeID and OfficeID = @OfficeID and not fInjured is null and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartDate and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndDate"
                Dim connOfficeTotal As New SqlConnection(sConnection)
                Dim cmdOfficeTotal As New SqlCommand(qTotalOfficeTime, connOfficeTotal)
                Dim drOfficeTotal As SqlDataReader
                cmdOfficeTotal.Parameters.Add(New SqlParameter("@OfficeID", OfficeID))
                cmdOfficeTotal.Parameters.Add(New SqlParameter("@StartDate", StartDate))
                cmdOfficeTotal.Parameters.Add(New SqlParameter("@EndDate", EndDate))

                Try

                    connOfficeTotal.Open()

                    drOfficeTotal = cmdOfficeTotal.ExecuteReader(Data.CommandBehavior.CloseConnection)

                    While drOfficeTotal.Read()

                        TotalOfficeTime = drOfficeTotal.Item("TotalTime") / 60


                    End While

                Catch ex As Exception
                    LogError("landing.aspx.vb :: GetOfficeTotals 2", ex.ToString)
                    Response.Write("ERROR: " & ex.Message)
                Finally
                    connOfficeTotal.Close()
                End Try

                OfficeOvertime = GetOfficeOvertime(OfficeID, StartDate, EndDate)

                returnstring = returnstring & "<tr><td>" & OfficeName & "</td><td>" & TotalOfficeTime & "</td><td>" & OfficeOvertime & "</td></tr>"

            End While

        Catch ex As Exception
            LogError("landing.aspx.vb :: GetOfficeTotals", ex.ToString)
            Response.Write("ERROR: " & ex.Message)
        Finally
            connOffices.Close()
        End Try

        returnstring = returnstring & "</table>"

        GetOfficeTotals = returnstring

    End Function

    Function GetOfficeOvertime(ByVal OfficeID As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As Double

        Dim Overtime As Double
        Dim sAdjustmentQuery As String = "select isNull(sum(Overtime),0) as Overtime from tTimeAdjustments, tEmployees where tEmployees.EmployeeID = tTimeAdjustments.EmployeeID and CAST(FLOOR(CAST(DateAdjusted AS FLOAT))AS DATETIME)>=@StartDate and CAST(FLOOR(CAST(DateAdjusted AS FLOAT))AS DATETIME)<=@EndDate and OfficeID = @OfficeID"
        Dim connAdjustment As New SqlConnection(sConnection)
        Dim cmdAdjustment As New SqlCommand(sAdjustmentQuery, connAdjustment)
        Dim drAdjustment As SqlDataReader
        cmdAdjustment.Parameters.Add(New SqlParameter("@OfficeID", OfficeID))
        cmdAdjustment.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdAdjustment.Parameters.Add(New SqlParameter("@EndDate", EndDate))

        Try
            connAdjustment.Open()
            drAdjustment = cmdAdjustment.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drAdjustment.Read
                Overtime = drAdjustment.Item("Overtime")
            End While
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message & "<br>")
            LogError("landing.aspx.vb :: GetOfficeOvertime", ex.ToString)
        Finally
            connAdjustment.Close()
        End Try

        GetOfficeOvertime = Overtime

    End Function

    Sub ToggleButton(ByVal StartDate As Date, ByVal EndDate As Date)

        If IsWeekLocked(StartDate, EndDate) = True Then
            BTN_Lock.Visible = False
            BTN_Unlock.Visible = True
        Else
            BTN_Lock.Visible = True
            BTN_Unlock.Visible = False
        End If

    End Sub

    Function CheckUnprocessed() As String

        Dim returnString As String = ""

        Dim sUnprocessedQuery As String = "select count(ActionID) as UnprocCount from tActions, tEmployees where tEmployees.CellPhone = tACtions.UserPerformed and Processed = 0 And ActionTime < DateAdd(dd, -1, GETDATE()) and (ActionEvent = 'Start Shift' or ActionEvent = 'Start Break' or ActionEvent = 'End Break' or ActionEvent = 'End Shift') and not ActionID in (select ActionID from tIgnoreUnprocessed)"
        Dim connUnprocessed As New SqlConnection(sConnection)
        Dim cmdUnprocessed As New SqlCommand(sUnprocessedQuery, connUnprocessed)
        Dim drUnprocessed As SqlDataReader
        Try
            connUnprocessed.Open()
            drUnprocessed = cmdUnprocessed.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drUnprocessed.Read

                returnString = "<span class=""alertlabelsmall"">There are " & drUnprocessed.Item("UnprocCount") & " unprocessed events older than 24 hours. <a href=""Openshifts.aspx"">Click here to see them</a>.</span>"

            End While

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
            LogError("Openshifts.aspx :: ShowUnprocessed", ex.ToString)
        Finally
            connUnprocessed.Close()
        End Try



        Return returnString

    End Function

End Class
