Imports System.Data.SqlClient

Partial Class worker
    Inherits GlobalClass


    Sub page_Load(ByVal sender As Object, ByVal e As EventArgs)

        Dim StartDate As DateTime = Request.QueryString("startdate")
        Dim EndDate As DateTime = Request.QueryString("enddate")
        Dim EmployeeID As Integer = Request.QueryString("id")

        'LIT_Main.Text = MakeTimeSheet(EmployeeID, StartDate, EndDate)

        LIT_Main.Text = MakeTimeSheet2(EmployeeID, StartDate, EndDate)

        Session("RedirectBack") = "worker.aspx?id=" & EmployeeID & "&startdate=" & StartDate & "&enddate=" & EndDate

    End Sub

    Function MakeTimeSheet(ByVal EmployeeID As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As String

        Dim returnString As String = ""

        returnString = returnString & "<h1>Weekly Time Sheet</h1>"

        Dim FirstName As String = "", LastName As String = "", PayrollID As String = ""
        Dim StartMap As String = "", EndMap As String = ""

        Dim sWorkerQuery As String = "select FirstName, LastName, Username, PayrollID from tWorkers where WorkerID=@WorkerID"
        Dim connWorker As New SqlConnection(sConnection)
        Dim cmdWorker As New SqlCommand(sWorkerQuery, connWorker)
        Dim drWorker As SqlDataReader
        cmdWorker.Parameters.Add(New SqlParameter("@WorkerID", EmployeeID))
        Try
            connWorker.Open()
            drWorker = cmdWorker.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drWorker.Read

                FirstName = drWorker.Item("FirstName")
                LastName = drWorker.Item("LastName")
                'Username = drWorker.Item("Username")
                PayrollID = drWorker.Item("PayrollID")

            End While

        Catch ex As Exception
            LogError("ax-worker.aspx.vb :: MakeTimeSheet", ex.ToString)
            returnString = returnString & "ERROR: " & ex.ToString
        Finally
            connWorker.Close()
        End Try

        returnString = returnString & "<h3>Employee Name: " & FirstName & " " & LastName & "</h3>" & vbCrLf
        returnString = returnString & "<h3>Employee/Payroll #: " & PayrollID & "</h3>" & vbCrLf
        returnString = returnString & "<h3>" & StartDate & " - " & EndDate & "</h3>" & vbCrLf & vbCrLf


        Dim TimeEntryID As Integer, fPM As String, fJobNumber As String, fState As String, StartTime As DateTime, EndTime As DateTime, fTravelTime As String, fPerDiem As String, Approved As Integer, StartAddress As String, EndAddress As String
        Dim StartLatitude As String = "", StartLongitude As String = "", EndLatitude As String = "", EndLongitude As String = ""
        Dim LoopDate As Date, DayName As String = "", Hours As Double, PerDiemTotal As Double
        Dim ApprovedInitials As String = "", LimitedViewClause As String = ""

        If Session("LimitedView") = 1 Then
            LimitedViewClause = " and fPM = '" & Session("PM") & "'"
        End If

        Dim TimeSheetQuery As String = "select TimeEntryID, fPM, fJobNumber, fState, StartTime, EndTime, fTravelTime, fPerDiem, Approved, StartAddress, EndAddress, StartLatitude, StartLongitude, EndLatitude, EndLongitude, Initials from tTimeEntry, tShifts, tEmployees where tEmployees.EmployeeID = tTimeEntry.ApprovedBy and tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.UserPerformed = @UserPerformed and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartDate and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndDate " & LimitedViewClause & " order by StartTime asc"
        Response.Write(TimeSheetQuery)
        Dim connTimeSheet As New SqlConnection(sConnection)
        Dim cmdTimeSheet As New SqlCommand(TimeSheetQuery, connTimeSheet)
        Dim drTimeSheets As SqlDataReader
        cmdTimeSheet.Parameters.Add(New SqlParameter("@UserPerformed", EmployeeID))
        cmdTimeSheet.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdTimeSheet.Parameters.Add(New SqlParameter("@EndDate", EndDate))
        Try
            connTimeSheet.Open()
            drTimeSheets = cmdTimeSheet.ExecuteReader(Data.CommandBehavior.CloseConnection)

            returnString = returnString & "<table class=""listing"">" & vbCrLf
            returnString = returnString & "<tr><th>Day</th><th>PM</th><th>Job #</th><th>State</th><th colspan=""2"">In</th><th colspan=""2"">Out</th><th>Hours</th>" & vbCrLf
            returnString = returnString & "<th colspan=""2"">Start</th><th colspan=""2"">End</th><th>C/P</th><th>Total</th><th>Per Diem</th><th>Approval</th></tr>" & vbCrLf

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
                StartLatitude = drTimeSheets.Item("StartLatitude")
                StartLongitude = drTimeSheets.Item("StartLongitude")
                EndLatitude = drTimeSheets.Item("EndLatitude")
                EndLongitude = drTimeSheets.Item("EndLongitude")
                ApprovedInitials = drTimeSheets.Item("Initials")


                Dim sStartTime As String, sEndTime As String

                sStartTime = FormatTime(StartTime)
                sEndTime = FormatTime(EndTime)

                If Not LoopDate = StartTime.Date Then
                    LoopDate = StartTime.Date
                    DayName = LoopDate.DayOfWeek.ToString
                Else
                    DayName = "&nbsp;"
                End If

                Dim EditClass As String = "", EditClassDate As String = ""

                If Not Approved = 1 Then
                    EditClass = "edit"
                    EditClassDate = "edit_date"
                End If

                If StartAddress <> "" Then
                    StartMap = "123<a target=""_blank"" href=""" & GoogleMapsURL & StartLatitude & "+" & StartLongitude & """ ><img src=""images/map.png"" title=""" & StartAddress & """/></a>"
                Else
                    StartMap = ""
                End If

                If EndAddress <> "" Then
                    EndMap = "<a target=""_blank"" href=""" & GoogleMapsURL & EndLatitude & "+" & EndLongitude & """><img src=""images/map.png""  title=""" & EndAddress & """/></a>"
                Else
                    EndMap = ""
                End If



                returnString = returnString & "<tr>"

                returnString = returnString & "<td>" & DayName & "</td>"
                returnString = returnString & "<td class=""" & EditClass & """>" & UCase(fPM) & "</td>"
                returnString = returnString & "<td class=""" & EditClass & """>" & fJobNumber & "</td>"
                returnString = returnString & "<td class=""" & EditClass & """>" & UCase(fState) & "</td>" & vbCrLf

                Hours = DateDiff(DateInterval.Minute, StartTime, EndTime) / 60

                If fTravelTime = "N/A" Then
                    returnString = returnString & "<td class=""" & EditClassDate & """ id=""starttime_" & TimeEntryID & """>" & sStartTime & "</td>"
                    returnString = returnString & "<td class=""map"">" & StartMap & "</td>"
                    returnString = returnString & "<td class=""" & EditClassDate & """ id=""endtime_" & TimeEntryID & """>" & sEndTime & "</td>"
                    returnString = returnString & "<td class=""map"">" & EndMap & "</td>"
                    returnString = returnString & "<td>" & Hours & "</td>"
                    returnString = returnString & "<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
                Else
                    returnString = returnString & "<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>"
                    returnString = returnString & "<td class=""" & EditClassDate & """ id=""starttime_" & TimeEntryID & """>" & sStartTime & "</td>"
                    returnString = returnString & "<td class=""map"">" & StartMap & "</td>"
                    returnString = returnString & "<td class=""" & EditClassDate & """ id=""endtime_" & TimeEntryID & """>" & sEndTime & "</td>"
                    returnString = returnString & "<td class=""map"">" & EndMap & "</td>"
                    returnString = returnString & "<td>" & Left(fTravelTime, 1) & "</td>"
                    returnString = returnString & "<td>" & Hours & "</td>"
                    'returnString = returnString & "<td>&nbsp;</td>" & vbCrLf
                End If

                returnString = returnString & "<td class=""edit_pd"" id=""perdiem_" & TimeEntryID & """>" & vbCrLf
                If fPerDiem = "Yes" Then
                    PerDiemTotal = Hours * 4.5
                    'Response.Write("$" & PerDiemTotal.ToString(0.0))
                    returnString = returnString & String.Format("{0:C}", PerDiemTotal)
                Else
                    returnString = returnString & "&nbsp;"
                End If
                returnString = returnString & "</td>"
                returnString = returnString & "<td class=""approval"">"

                If Approved = 1 Then
                    returnString = returnString & "<img src=""images/tick.png""><strong>123" & ApprovedInitials & "</strong>"
                Else
                    returnString = returnString & "<input type=""button"" class=""approvebutton"" id=""approve_" & TimeEntryID & """ value=""Approve"">"
                End If

                returnString = returnString & "</td>" & vbCrLf
                returnString = returnString & "</tr>" & vbCrLf

            End While

            returnString = returnString & "</table>"
            returnString = returnString & "<div id=""ddl_holder"">" & GenerateWorkerDropdown(StartDate, EndDate, EmployeeID) & "</div>"

        Catch ex As Exception
            LogError("ax-worker.aspx.vb :: MakeTimeSheet 2", ex.ToString)
            returnString = returnString & "ERROR: " & ex.ToString
        Finally
            connTimeSheet.Close()

        End Try

        MakeTimeSheet = returnString

    End Function

    Function MakeTimeSheet2(ByVal EmployeeID As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As String

        Dim returnString As String = ""

        returnString = returnString & "<h1>Weekly Time Sheet</h1>"

        Dim FirstName As String = "", LastName As String = "", Username As String = "", PayrollID As String = ""
        Dim StartLatitude As String = "", StartLongitude As String = "", EndLatitude As String = "", EndLongitude As String = ""
        Dim StartMap As String = "", EndMap As String = ""
        Dim ApprovedInitials As String = "", RowHighlightClass As String = ""
        Dim EnteredOn As String = "", EnteredManually As Integer, ApprovedOn As String = ""


        Dim sWorkerQuery As String = "select FirstName, LastName, Username, PayrollID from tEmployees where EmployeeID=@EmployeeID"
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

            End While

        Catch ex As Exception
            LogError("ax-worker.aspx.vb :: MakeTimeSheet2", ex.ToString)
            returnString = returnString & "ERROR: " & ex.ToString
        Finally
            connWorker.Close()
        End Try

        returnString = returnString & "<table style=""width:100%;""><tr><td style=""width:500px;"">"

        returnString = returnString & "<h3>Employee Name: " & FirstName & " " & LastName & "</h3>" & vbCrLf
        returnString = returnString & "<h3>Employee/Payroll #: " & PayrollID & "</h3>" & vbCrLf
        returnString = returnString & "<h3>" & StartDate & " - " & EndDate & "</h3>" & vbCrLf & vbCrLf

        returnString = returnString & "</td>"

        returnString = returnString & "<td colspan=""6"" style=""width:360px;""><h3>Notes:</h3><div id=""saveweeknotes_" & EmployeeID & "_" & StartDate & """ class=""weeknotes"">" & GetWeekNotes(EmployeeID, StartDate) & "</div>"
        returnString = returnString & "</table>"

        Dim TimeEntryID As Integer, fPM As String, fJobNumber As String, fState As String, StartTime As DateTime, EndTime As DateTime, fTravelTime As String, fPerDiem As String, Approved As Integer, StartAddress As String, EndAddress As String, ShiftID As Integer, fType As String, fShiftType As String, fInjured As String, fComments As String, fProjectName As String, fPW As String
        Dim LoopDate As Date, DayName As String = "", Hours As Double, LastDay As String = "", LastShift As Integer = 0, ShiftCounter As Integer = 1, Overtime As String, JobList As String = ""
        Dim WeekTotalTime As Double = 0, WeekOvertime As Double = 0, WeekTravelTime As Double = 0, WeekStandardTime As Double = 0
        Dim EditClass As String = "", EditClassDate As String = "", EditClassType As String = "", EditClassShiftType As String = "", EditClassTravel As String = "", EditClassPerDiem As String = "", EditClassInjured As String = "", EditClassComments As String = "", EditClassPW As String = "", EditClassOT As String = "", EditClassOTLock As String = "", EditClassJobNumber As String = ""
        Dim LimitedViewClause As String = ""

        If Session("LimitedView") = 1 Then
            LimitedViewClause = " and fPM = '" & Session("ShowAsPM") & "'"
        End If

        Dim TimeSheetQuery As String = "select TimeEntryID, fPM, fJobNumber, fState, StartTime, EndTime, fTravelTime, fPerDiem, Approved, StartAddress, EndAddress, tTimeEntry.ShiftID, fType, fShiftType, fInjured, fComments, fProjectName, StartLatitude, StartLongitude, EndLatitude, EndLongitude, fPW, Initials, IsNull(EnteredOn,'') as EnteredOn, isNull(EnteredManually,0) as EnteredManually, IsNull(ApprovedOn,'') as ApprovedOn from tTimeEntry, tShifts, tEmployees where tEmployees.EmployeeID = isNull(tTimeEntry.ApprovedBy,1) and tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.UserPerformed = @EmployeeID and not fInjured is null and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartDate and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndDate " & LimitedViewClause & " order by StartTime asc"
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
                ApprovedInitials = drTimeSheets.Item("Initials")
                EnteredOn = drTimeSheets.Item("EnteredOn")
                EnteredManually = drTimeSheets.Item("EnteredManually")
                ApprovedOn = drTimeSheets.Item("ApprovedOn")


                Dim sStartTime As String, sEndTime As String

                sStartTime = FormatTime(StartTime)
                sEndTime = FormatTime(EndTime)

                LoopDate = StartTime.Date
                DayName = LoopDate.DayOfWeek.ToString


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

                    EditClassDate = ""

                    Dim strOvertime As String = GetOvertime(EmployeeID, LoopDate, fJobNumber, fState, fPW, JobList)
                    'Response.Write(":" & EmployeeID & ":" & LoopDate & ":" & fJobNumber & ":" & fState & ":" & fPW & ":" & JobList & ":<br><br>")


                    If strOvertime = "NoOT" Then
                        Overtime = 0
                    Else
                        Overtime = CDbl(strOvertime)
                    End If

                    'LOCK FOR OVERTIME
                    If IsDayLocked(LoopDate) = False And Overtime = 0 And IsShiftLocked(TimeEntryID) = False Then

                        EditClassOTLock = "edit"
                        EditClassDate = "edit_date"
                        EditClassPW = "edit_yesno"

                        EditClass = "edit"
                        EditClassType = "edit"
                        EditClassShiftType = "edit_shifttype"
                        EditClassTravel = "edit_travel"
                        EditClassPerDiem = "edit_yesno"
                        EditClassInjured = "edit_yesno"
                        EditClassComments = "edit_comments"
                        EditClassJobNumber = "edit_jobnumber"

                    Else
                        EditClassOTLock = ""
                        EditClassDate = ""
                        EditClassPW = ""

                        EditClass = ""
                        EditClassType = ""
                        EditClassShiftType = ""
                        EditClassTravel = ""
                        EditClassPerDiem = ""
                        EditClassInjured = ""
                        EditClassComments = ""
                        EditClassJobNumber = ""

                    End If

                    'LOCK FOR APPROVED OR LOCKED WEEK
                    If IsDayLocked(LoopDate) = False And IsShiftLocked(TimeEntryID) = False Then

                        EditClass = "edit"
                        EditClassType = "edit"
                        EditClassShiftType = "edit_shifttype"
                        EditClassTravel = "edit_travel"
                        EditClassPerDiem = "edit_yesno"
                        EditClassInjured = "edit_yesno"
                        EditClassComments = "edit_comments"

                        EditClassOT = "edit_overtime"

                    Else
                        EditClass = ""
                        EditClassType = ""
                        EditClassShiftType = ""
                        EditClassTravel = ""
                        EditClassPerDiem = ""
                        EditClassInjured = ""
                        EditClassComments = ""

                        EditClassOT = ""

                    End If

                    Dim JobNumberDash As String = fJobNumber

                    If Not fComments = "" Then
                        EditClassComments = EditClassComments & " highlight_comments"
                    End If


                    If IsJobNumberValid(fJobNumber) = False Then
                        EditClassJobNumber = EditClassJobNumber & " invalidjobnumber"
                        'Else
                        'JobNumberDash = AddDashes(fJobNumber)
                        'JobNumberDash = fJobNumber
                    End If

                    RowHighlightClass = ""

                    If (fShiftType = "Holiday") Or (fShiftType = "Vacation") Or (fShiftType = "Sick") Or (fShiftType = "PTO") Or (fShiftType = "Personal") Then
                        RowHighlightClass = " highlight"
                    End If

                    If Left(fJobNumber, 5) = "01-15" Or Left(fJobNumber, 4) = "0115" Then
                        RowHighlightClass = " highlight3"
                    End If

                    If Left(fJobNumber, 5) = "01-14" Or Left(fJobNumber, 4) = "0114" Then
                        RowHighlightClass = " highlight4"
                    End If

                    returnString = returnString & "<tr class=""shift label" & RowHighlightClass & """><td rowspan=2>Shift<br>" & ShiftCounter & "</td><td>PM</td><td>State</td><td>Job&nbsp;#</td><td>PW</td><td>Project</td><td style=""width:80px;"">Shift&nbsp;Type</td><td>Travel</td><td>Per&nbsp;Diem</td><td>Injured</td><td>Comments</td>"

                    If Overtime = "NoOT" Then
                        returnString = returnString & "<td>&nbsp;</td>"
                    Else
                        returnString = returnString & "<td style=""background-color:#A01D1E;color:#fff; font-weight:bold;text-align:center;"">OT&nbsp;Hrs</td>"
                    End If
                    returnString = returnString & "<td></td></tr>"
                    'returnString = returnString & "<td style=""background-color:#A01D1E;color:#fff; font-weight:bold;text-align:center;"">OT&nbsp;Hrs</td><td></td></tr>"



                    returnString = returnString & "<tr class=""shift" & RowHighlightClass & """>"
                    returnString = returnString & "<td class=""" & EditClass & """ id=""pm_" & TimeEntryID & """>" & UCase(fPM) & "</td>"
                    returnString = returnString & "<td class=""" & EditClassOTLock & """ id=""state_" & TimeEntryID & """>" & UCase(fState) & "</td>"
                    returnString = returnString & "<td class=""" & EditClassJobNumber & """ id=""jobnumber_" & TimeEntryID & """>" & fJobNumber & "</td>"
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
                        returnString = returnString & "<td class=""" & EditClassOT & """ id=""ot_" & EmployeeID & "_" & fJobNumber & "_" & LoopDate & "_" & fPW & "_" & fState & """ style=""background-color:#A01D1E;color:#fff; font-weight:bold;text-align:center;"" alt=""" & TimeEntryID & """>" & Overtime & "</td>"
                    End If
                    returnString = returnString & "<td></td>"

                    returnString = returnString & "</tr>"

                    ShiftCounter = ShiftCounter + 1
                    JobList = JobList & "%" & fJobNumber & "%"

                    If Not Overtime = "NoOT" Then
                        WeekOvertime = WeekOvertime + Overtime
                    Else
                        Overtime = 0
                    End If

                End If

                If ShiftID = LastShift Then Overtime = 0

                If fTravelTime = "Company" Or fTravelTime = "Personal" Then
                    WeekTravelTime = WeekTravelTime + Hours
                Else
                    WeekStandardTime = WeekStandardTime + Hours - CDbl(Overtime)
                End If

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
                    If IsDayLocked(LoopDate) = False Then
                        returnString = returnString & "<img src=""images/tick.png""><strong>" & ApprovedInitials & "</strong><br><span class=""unapprove"">[<a href=""#"" class=""unapprove"" id=""unapprove_" & TimeEntryID & """>unapprove</a>]</span>"
                    Else
                        returnString = returnString & "<img src=""images/tick.png""><strong>" & ApprovedInitials & "</strong><br><span class=""unapprove"">[week locked]</span>"
                    End If
                Else
                    If IsDayLocked(LoopDate) = False Then
                        returnString = returnString & "<input type=""button"" class=""approvebutton"" id=""approve_" & TimeEntryID & """ value=""Approve"">"
                    Else
                        returnString = returnString & "<input type=""button"" class=""approvebutton"" value=""Approve"" disabled><br><span class=""unapprove"">[week locked]</span>"
                    End If

                End If

                If (EnteredOn = "" Or EnteredOn = "1/1/1900") And (ApprovedOn = "" Or ApprovedOn = "1/1/1900") And EnteredManually = 0 Then
                    'do nothing
                Else
                    returnString = returnString & "&nbsp;<a href=""#"" class=""tei-trigger"" id=""tei-trigger-" & TimeEntryID & """>[?]</a>"
                    returnString = returnString & "<div class=""tei-holder"" id=""tei-holder-" & TimeEntryID & """>"



                    If EnteredManually = 1 Then
                        returnString = returnString & "Entered via FFM phone app<br />"
                        If Not EnteredOn = "" And Not EnteredOn = "1/1/1900" Then returnString = returnString & "FFM synced: " & EnteredOn & "<br />"
                    ElseIf EnteredManually = 2 Then
                        returnString = returnString & "Entered via web<br />"
                        If Not EnteredOn = "" And Not EnteredOn = "1/1/1900" Then returnString = returnString & "Entered: " & EnteredOn & "<br />"
                    End If

                    If Not ApprovedOn = "" And Not ApprovedOn = "1/1/1900" Then returnString = returnString & "Approved: " & ApprovedOn

                    returnString = returnString & ""
                    returnString = returnString & "</div>"

                End If

                returnString = returnString & "</td>" & vbCrLf
                returnString = returnString & "</tr>"

                LastDay = DayName
                LastShift = ShiftID

            End While

            returnString = returnString & "<tr class=""totalsrow""><td></td><td colspan=""6""><strong>Week Totals</strong><br />"
            returnString = returnString & "TOTAL TIME: <strong>" & WeekTotalTime & "</strong><br />"
            returnString = returnString & "STANDARD TIME: <strong>" & WeekStandardTime & "</strong><br />"
            returnString = returnString & "OVERTIME: <strong>" & WeekOvertime & "</strong><br>"
            returnString = returnString & "TRAVEL TIME: <strong>" & WeekTravelTime & "</strong>"

            returnString = returnString & "<td colspan=""6"">"
            'returnString = returnString & "<textarea id=""weeknotes"" class=""weeknotes"">" & GetWeekNotes(EmployeeID, StartDate) & "</textarea><br/>"
            'returnString = returnString & "<input type=""button"" class=""saveweeknotes"" id=""saveweeknotes_" & EmployeeID & "_" & StartDate & """ value=""Save Notes"" />"
            returnString = returnString & "</td></tr></table>"
            returnString = returnString & "<div id=""ddl_holder"">" & GenerateWorkerDropdown(StartDate, EndDate, EmployeeID) & "</div>"

            returnString = returnString & "<div id=""newshift""><img src=""images/icon_add.gif"" /> <strong><a href=""addtime.aspx?action=redirectback"">Add Shift</a></strong></div>"
            returnString = returnString & "<p class=""info""><strong>NOTES:</strong><br>"
            returnString = returnString & "- Please include your initials with any information you enter into the Comments box.<br>"
            returnString = returnString & "- You can edit all Shift fields by clicking on them, as well as the Start and End times of each time entry.<br>"
            returnString = returnString & "- The red Overtime field represents the number of hours out of the total hours for the shift that are to be processed as overtime.<br>"
            returnString = returnString & "- Please note that the total hour count per shift will not update until the page is refreshed. You can use the Refresh button of your browser  or <a href=""#"" Onclick=""javascript:location.reload(true)"">click here</a>.<br>"
            returnString = returnString & "- The Overtime value can also be adjusted by the clerk prior to processing payroll, and in that case, the change will be reflected here.<br>"
            returnString = returnString & "</p>"


        Catch ex As Exception
            LogError("ax-worker.aspx.vb :: MakeTimeSheet2 2", ex.ToString & " --- " & ex.Source & " --- " & ex.Message & " --- " & ex.StackTrace)
            returnString = returnString & "ERROR: " & ex.ToString
        Finally
            connTimeSheet.Close()

        End Try

        MakeTimeSheet2 = returnString

    End Function

    Function AddDashes(Jobnumber As String) As String

        Dim returnstring As String = ""

        Jobnumber = Regex.Replace(Jobnumber, "[^0-9]", "")

        Dim JN1 As String = "", JN2 As String = "", JN3 As String = "", JN4 As String = ""

        JN1 = Jobnumber.Substring(0, 2)
        JN2 = Jobnumber.Substring(2, 2)
        JN3 = Jobnumber.Substring(4, 6)

        returnstring = JN1 & "-" & JN2 & "-" & JN3

        If Jobnumber.Length > 10 Then
            JN4 = Jobnumber.Substring(10, Len(Jobnumber) - 10)
            returnstring = returnstring & "-" & JN4
        End If

        Return returnstring

    End Function

    Function GenerateWorkerDropdown(ByVal StartDate As Date, ByVal EndDate As Date, ByVal CurrentEmployee As Integer) As String

        Dim returnstring As String = "Select Worker: <select id=""worker_ddl"">"


        Dim sWorkerQuery As String = "select EmployeeID, LastName, FirstName from tEmployees where Worker = 1 and Active = 1 order by LastName asc"
        Dim connWorker As New SqlConnection(sConnection)
        Dim cmdWorker As New SqlCommand(sWorkerQuery, connWorker)
        Dim drWorker As SqlDataReader

        Try
            connWorker.Open()
            drWorker = cmdWorker.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drWorker.Read()

                returnstring = returnstring & "<option value=""worker.aspx?id=" & drWorker.Item("EmployeeID") & "&startdate=" & StartDate & "&enddate=" & EndDate & """"
                If drWorker.Item("EmployeeID") = CurrentEmployee Then returnstring = returnstring & " selected"
                returnstring = returnstring & ">" & drWorker.Item("LastName") & ", " & drWorker.Item("FirstName") & "</option>"

            End While
        Catch ex As Exception
            LogError("ax-worker.aspx.vb :: GenerateWorkerDropdown", ex.ToString)
            returnstring = returnstring & "ERROR: " & ex.ToString
        Finally
            connWorker.Close()
        End Try


        returnstring = returnstring & "</select>"

        GenerateWorkerDropdown = returnstring

    End Function

    Function GetWeekNotes(ByVal EmployeeID As Integer, ByVal DateStart As Date) As String

        Dim NotesText As String = ""
        Dim qCheckNotes As String = "select NotesText  from tNotes where EmployeeID=@EmployeeID and WeekStartDate=@DateStart"
        Dim ConnCheckNotes As New SqlConnection(sConnection)
        Dim CmdCheckNotes As New SqlCommand(qCheckNotes, ConnCheckNotes)
        Dim drCheckNotes As SqlDataReader

        CmdCheckNotes.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
        CmdCheckNotes.Parameters.Add(New SqlParameter("@DateStart", DateStart))

        Try
            ConnCheckNotes.Open()
            drCheckNotes = CmdCheckNotes.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheckNotes.Read()
                NotesText = drCheckNotes.Item("NotesText")
            End While
        Catch ex As Exception
            LogError("ax-worker.aspx.vb :: GetWeekNotes", ex.ToString)
            Response.Write("ERROR: " & ex.ToString)
        Finally
            ConnCheckNotes.Close()
        End Try

        Return NotesText

    End Function

    Function IsJobNumberValid(JobNumber As String) As Boolean

        Dim JobNum1 As String = "01", JobNum2 As String = "", JobNum3 As String = "", JobNum4 As String = ""
        Dim IsItValid As Boolean = True
        Dim JobNumberOriginal As String = JobNumber

        JobNumber = Regex.Replace(JobNumber, "[^0-9]", "")

        If Not Len(JobNumber) = 10 And Not Len(JobNumber) = 13 Then
            IsItValid = False
        Else
            If Not JobNumber.Substring(0, 2) = "01" Then
                IsItValid = False
            Else
                If Not JobNumber.Substring(2, 2) = "01" And Not JobNumber.Substring(2, 2) = "02" And Not JobNumber.Substring(2, 2) = "03" And Not JobNumber.Substring(2, 2) = "04" And Not JobNumber.Substring(2, 2) = "12" Then
                    IsItValid = False
                End If
            End If
        End If

        If IsItValid = True Then

            Dim JobNumberDashed As String

            JobNum1 = JobNumber.Substring(0, 2)
            JobNum2 = JobNumber.Substring(2, 2)
            JobNum3 = JobNumber.Substring(4, 6)
            JobNumberDashed = JobNum1 & "-" & JobNum2 & "-" & JobNum3
            If Len(JobNumber) = 13 Then
                JobNum4 = JobNumber.Substring(10, 3)
                JobNumberDashed = JobNumberDashed & "-" & JobNum4
            End If

            If Not JobNumberOriginal = JobNumberDashed Then
                IsItValid = False
            End If

        End If

        Return IsItValid

    End Function


End Class


