Imports System.Data.SqlClient


Partial Class entertime
    Inherits GlobalClass


    Sub page_load()

        Dim StartDate As Date, EndDate As Date
        Try

            If IsPostBack = False Then

                MakeDateDropdown()

                Dim SessStartDate As String = ""
                SessStartDate = Session("DateRangeStart")

                If Not SessStartDate = "" Then
                    StartDate = CDate(Session("DateRangeStart"))
                    EndDate = CDate(Session("DateRangeEnd"))

                    DDL_DateRange.SelectedValue = FormatDateTime(StartDate, DateFormat.LongDate) & " - " & FormatDateTime(EndDate, DateFormat.LongDate)
                Else
                    Dim sDateRange As String = DDL_DateRange.SelectedValue
                    Dim arrDateRange As Array = sDateRange.Split("-")

                    StartDate = CDate(arrDateRange(0))
                    EndDate = CDate(arrDateRange(1))
                End If

            Else
                Dim sDateRange As String = DDL_DateRange.SelectedValue
                Dim arrDateRange As Array = sDateRange.Split("-")

                StartDate = arrDateRange(0)
                EndDate = arrDateRange(1)

                Session("DateRangeStart") = StartDate
                Session("DateRangeEnd") = EndDate
            End If
        Catch ex As Exception
            LogError("entertime.aspx.vb :: page_load", ex.ToString)
        End Try

        LIT_Main.Text = MakeTimeSheet(Session("UserID"), StartDate, EndDate)

        If IsDayLocked(StartDate) = True Then
            LBL_Info.visible = True
            LBL_Info.Text = "Week has been closed by payroll."
        Else
            LBL_Info.Visible = False
        End If

    End Sub

    Sub MakeDateDropdown()

        'Dim returnstring As String = ""
        Dim EarliestRecord As DateTime

        Dim sStartTimeQuery As String = "select top 1 startTime as EarliestStartTime from tTimeEntry where startTime > '2000-1-1' order by startTime asc"
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

    Function MakeTimeSheet(ByVal EmployeeID As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As String

        Dim returnString As String = ""

        returnString = returnString & "<h1>Weekly Time Sheet</h1>"

        Dim FirstName As String = "", LastName As String = "", Username As String = "", PayrollID As String = "", isPM As Integer, ReadOnlyAccess As Integer = 0, FlagReadOnly As Boolean = False
        Dim StartLatitude As String = "", StartLongitude As String = "", EndLatitude As String = "", EndLongitude As String = ""
        Dim StartMap As String = "", EndMap As String = ""


        Dim sWorkerQuery As String = "select FirstName, LastName, Username, PayrollID, PM, ReadOnlyAccess from tEmployees where EmployeeID=@EmployeeID"
        Dim connWorker As New SqlConnection(sConnection)
        Dim cmdWorker As New SqlCommand(sWorkerQuery, connWorker)
        Dim drWorker As SqlDataReader
        cmdWorker.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
        Try
            connWorker.Open()
            drWorker = cmdWorker.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drWorker.Read

                FirstName = drWorker.Item("FirstName")
                LastName = drWorker.Item("LastName")
                Username = drWorker.Item("Username")
                PayrollID = drWorker.Item("PayrollID")
                isPM = drWorker.Item("PM")
                ReadOnlyAccess = drWorker.Item("ReadOnlyAccess")

            End While

        Catch ex As Exception
            returnString = returnString & "ERROR: " & ex.ToString
        Finally
            connWorker.Close()
        End Try

        If Session("UserID") = EmployeeID And ReadOnlyAccess = 1 Then
            FlagReadOnly = True
        End If

        'returnString = returnString & "<h3>Employee Name: " & FirstName & " " & LastName & "</h3>" & vbCrLf
        'returnString = returnString & "<h3>Employee/Payroll #: " & PayrollID & "</h3>" & vbCrLf
        'returnString = returnString & "<h3>" & StartDate & " - " & EndDate & "</h3>" & vbCrLf & vbCrLf
        If FlagReadOnly = False Then
            returnString = returnString & "<div id=""newshift""><img src=""images/icon_add.gif"" /> <strong><a href=""addtime.aspx"">Add Shift</a></strong></div>"
        End If

        Dim TimeEntryID As Integer, fPM As String, fJobNumber As String, fState As String, StartTime As DateTime, EndTime As DateTime, fTravelTime As String, fPerDiem As String, Approved As Integer, StartAddress As String, EndAddress As String, ShiftID As Integer, fType As String, fShiftType As String, fInjured As String, fComments As String, fProjectName As String, fPW As String
        Dim LoopDate As Date, DayName As String = "", Hours As Double, LastDay As String = "", LastShift As Integer = 0, ShiftCounter As Integer = 1, Overtime As String, JobList As String = ""
        Dim WeekTotalTime As Double = 0, WeekOvertime As Double = 0

        Dim TimeSheetQuery As String = "select TimeEntryID, fPM, fJobNumber, fState, StartTime, EndTime, fTravelTime, fPerDiem, Approved, StartAddress, EndAddress, tTimeEntry.ShiftID, fType, fShiftType, fInjured, fComments, fProjectName, StartLatitude, StartLongitude, EndLatitude, EndLongitude, fPW from tTimeEntry, tShifts where tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.UserPerformed = @EmployeeID and not fInjured is null and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartDate and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndDate order by StartTime asc"
        Dim connTimeSheet As New SqlConnection(sConnection)
        Dim cmdTimeSheet As New SqlCommand(TimeSheetQuery, connTimeSheet)
        Dim drTimeSheets As SqlDataReader
        cmdTimeSheet.Parameters.Add(New SqlParameter("@UserPerformed", Username))
        cmdTimeSheet.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdTimeSheet.Parameters.Add(New SqlParameter("@EndDate", EndDate))
        cmdTimeSheet.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
        Try
            connTimeSheet.Open()
            drTimeSheets = cmdTimeSheet.ExecuteReader(Data.CommandBehavior.CloseConnection)

            returnString = returnString & "<table class=""listing"">" & vbCrLf
            'returnString = returnString & "<tr><th>Day</th><th>PM</th><th>Job #</th><th>State</th><th colspan=""2"">In</th><th colspan=""2"">Out</th><th>Hours</th>" & vbCrLf
            'returnString = returnString & "<th colspan=""2"">Start</th><th colspan=""2"">End</th><th>C/P</th><th>Total</th><th>Per Diem</th><th>Approval</th></tr>" & vbCrLf

            While drTimeSheets.Read

                TimeEntryID = drTimeSheets.Item("TimeEntryID")
                fPM = drTimeSheets.Item("fPM")
                fJobNumber = drTimeSheets.Item("fJobNumber")
                fState = drTimeSheets.Item("fState")
                StartTime = drTimeSheets.Item("StartTime")
                EndTime = drTimeSheets.Item("EndTime")
                fTravelTime = drTimeSheets.Item("fTravelTime")
                fPerDiem = drTimeSheets.Item("fPerDiem")
                Approved = drTimeSheets.Item("Approved")
                StartAddress = drTimeSheets.Item("StartAddress")
                EndAddress = drTimeSheets.Item("EndAddress")
                ShiftID = drTimeSheets.Item("ShiftID")
                fType = drTimeSheets.Item("fType")
                fShiftType = drTimeSheets.Item("fShiftType")
                fInjured = drTimeSheets.Item("fInjured")
                fComments = drTimeSheets.Item("fComments")
                fProjectName = drTimeSheets.Item("fProjectName")
                StartLatitude = drTimeSheets.Item("StartLatitude")
                StartLongitude = drTimeSheets.Item("StartLongitude")
                EndLatitude = drTimeSheets.Item("EndLatitude")
                EndLongitude = drTimeSheets.Item("EndLongitude")
                fPW = drTimeSheets.Item("fPW")

                Dim sStartTime As String, sEndTime As String

                sStartTime = FormatTime(StartTime)
                sEndTime = FormatTime(EndTime)

                LoopDate = StartTime.Date
                DayName = LoopDate.DayOfWeek.ToString

                Dim EditClass As String = "", EditClassDate As String = "", EditClassType As String = "", EditClassShiftType As String = "", EditClassTravel As String = "", EditClassPerDiem As String = "", EditClassInjured As String = "", EditClassComments As String = "", EditClassPW As String = ""


                If Not Approved = 1 And IsDayLocked(LoopDate) = False And Overtime = 0 And isPM = 1 And FlagReadOnly = False Then
                    EditClass = "edit"
                    EditClassDate = "edit_date"
                    EditClassType = "edit"
                    EditClassShiftType = "edit_shifttype"
                    EditClassTravel = "edit_travel"
                    EditClassPerDiem = "edit_yesno"
                    EditClassInjured = "edit_yesno"
                    EditClassComments = "edit_comments"
                    EditClassPW = "edit_yesno"
                End If

                If StartAddress <> "" Then
                    StartMap = "<a target=""_blank"" href=""" & GoogleMapsURL & StartLatitude & "+" & StartLongitude & """ ><img src=""images/map.png"" title=""" & StartAddress & """/></a>"
                Else
                    StartMap = ""
                End If

                If EndAddress <> "" Then
                    EndMap = "<a target=""_blank"" href=""" & GoogleMapsURL & EndLatitude & "+" & EndLongitude & """><img src=""images/map.png""  title=""" & EndAddress & """/></a>"
                Else
                    EndMap = ""
                End If

                Hours = DateDiff(DateInterval.Minute, StartTime, EndTime) / 60

                If Not DayName = LastDay Then
                    'WRITE DAY HEADING
                    returnString = returnString & "<tr><th colspan=13>" & DayName & " " & LoopDate & "</th></tr>"
                    ShiftCounter = 1
                    JobList = ""
                End If

                If Not ShiftID = LastShift Then
                    'WRITE SHIFT HEADING

                    Dim strOvertime As String = GetOvertime(EmployeeID, LoopDate, fJobNumber, fState, fPW, JobList)
                    'Response.Write(":" & EmployeeID & ":" & LoopDate & ":" & fJobNumber & ":" & fState & ":" & fPW & ":" & JobList & ":<br><br>")


                    If strOvertime = "NoOT" Then
                        Overtime = 0
                    Else
                        Overtime = CDbl(strOvertime)
                    End If

                    returnString = returnString & "<tr class=""shift label""><td rowspan=2>Shift<br>" & ShiftCounter & "</td><td>PM</td><td>State</td><td>Job&nbsp;#</td><td>PW</td><td>Project</td><td style=""width:80px;"">Shift&nbsp;Type</td><td>Travel</td><td>Per&nbsp;Diem</td><td>Injured</td><td>Comments</td>"

                    If Overtime = "NoOT" Then
                        returnString = returnString & "<td>&nbsp;</td>"
                    Else
                        returnString = returnString & "<td style=""background-color:#A01D1E;color:#fff; font-weight:bold;text-align:center;"">OT&nbsp;Hrs</td>"
                    End If
                    returnString = returnString & "<td></td></tr>"
                    'returnString = returnString & "<td style=""background-color:#A01D1E;color:#fff; font-weight:bold;text-align:center;"">OT&nbsp;Hrs</td><td></td></tr>"

                    returnString = returnString & "<tr class=""shift"">"
                    returnString = returnString & "<td class=""" & EditClass & """ id=""pm_" & TimeEntryID & """>" & UCase(fPM) & "</td>"
                    returnString = returnString & "<td class=""" & EditClass & """ id=""state_" & TimeEntryID & """>" & UCase(fState) & "</td>"
                    returnString = returnString & "<td class=""" & EditClass & """ id=""jobnumber_" & TimeEntryID & """>" & fJobNumber & "</td>"
                    returnString = returnString & "<td class=""" & EditClassPW & """ id=""pw_" & TimeEntryID & """>" & fPW & "</td>"
                    returnString = returnString & "<td class=""" & EditClass & """ id=""projectname_" & TimeEntryID & """>" & fProjectName & "</td>"
                    'returnString = returnString & "<td class=""" & EditClassType & """ id=""type_" & TimeEntryID & """>" & fType & "</td>"
                    returnString = returnString & "<td class=""" & EditClassShiftType & """ id=""shifttype_" & TimeEntryID & """>" & fShiftType & "</td>"
                    returnString = returnString & "<td class=""" & EditClassTravel & """ id=""travel_" & TimeEntryID & """>" & fTravelTime & "</td>"
                    returnString = returnString & "<td class=""" & EditClassPerDiem & """ id=""perdiem_" & TimeEntryID & """>" & fPerDiem & "</td>"
                    returnString = returnString & "<td class=""" & EditClassInjured & """ id=""injured_" & TimeEntryID & """>" & fInjured & "</td>"
                    returnString = returnString & "<td class=""" & EditClassComments & """ id=""comments_" & TimeEntryID & """>" & fComments & "</td>"

                    If Overtime = "NoOT" Then
                        returnString = returnString & "<td>&nbsp;</td>"
                    Else
                        returnString = returnString & "<td id=""ot_" & Username & "_" & fJobNumber & "_" & LoopDate & """ style=""background-color:#A01D1E;color:#fff; font-weight:bold;text-align:center;"">" & Overtime & "</td>"
                    End If
                    returnString = returnString & "<td></td>"

                    returnString = returnString & "</tr>"

                    ShiftCounter = ShiftCounter + 1
                    JobList = JobList & "%" & fJobNumber & "%"

                End If

                If Not Overtime = "NoOT" Then WeekOvertime = WeekOvertime + Overtime

                WeekTotalTime = WeekTotalTime + Hours


                returnString = returnString & "<tr>"
                returnString = returnString & "<td>" & Hours & "hrs</td><td colspan=4>START:"
                returnString = returnString & "<span class=""" & EditClassDate & """ id=""starttime_" & TimeEntryID & """>" & sStartTime & "</span>"
                returnString = returnString & StartMap & "</td>"
                returnString = returnString & "<td colspan=5>END:"
                returnString = returnString & "<span class=""" & EditClassDate & """ id=""endtime_" & TimeEntryID & """>" & sEndTime & "</span>"
                returnString = returnString & EndMap & "</td>"
                returnString = returnString & "<td colspan=2>&nbsp;</td>"
                returnString = returnString & "<td class=""approval"">"

                If Approved = 1 Then
                    returnString = returnString & "<img src=""images/tick.png"">"
                Else
                    returnString = returnString & ""
                End If

                returnString = returnString & "</td>" & vbCrLf
                returnString = returnString & "</tr>"

                LastDay = DayName
                LastShift = ShiftID

            End While

            returnString = returnString & "<tr class=""totalsrow""><td colspan=""3""><strong>Week Totals</strong></td><td colspan=""3"">TOTAL TIME: <strong>" & WeekTotalTime & "</strong></td><td colspan=""3"">STANDARD TIME: <strong>" & WeekTotalTime - WeekOvertime & "</strong></td><td colspan=""3"">OVERTIME: <strong>" & WeekOvertime & "</strong></td><td colspan=""2""></td></tr>"
            returnString = returnString & "</table>"
            'returnString = returnString & "<div id=""ddl_holder"">" & GenerateWorkerDropdown(StartDate, EndDate, EmployeeID) & "</div>"

            If FlagReadOnly = False Then
                returnString = returnString & "<div id=""newshift""><img src=""images/icon_add.gif"" /> <strong><a href=""addtime.aspx"">Add Shift</a></strong></div>"
            End If
            returnString = returnString & "<p class=""info""><strong>NOTES:</strong><br>"
            returnString = returnString & "- Please include your initials with any information you enter into the Comments box.<br>"
            returnString = returnString & "- You can edit all Shift fields by clicking on them, as well as the Start and End times of each time entry.<br>"
            returnString = returnString & "- No changes can be made once time has been approved by a Project Manager.<br>"
            returnString = returnString & "- The red Overtime field represents the number of hours out of the total hours for the shift that are to be processed as overtime.<br>"
            returnString = returnString & "- The Overtime value can also be adjusted by the clerk prior to processing payroll, and in that case, the change will be reflected here.<br>"
            returnString = returnString & "</p>"


        Catch ex As Exception
            returnString = returnString & "ERROR: " & ex.ToString
        Finally
            connTimeSheet.Close()

        End Try

        MakeTimeSheet = returnString

    End Function

End Class
