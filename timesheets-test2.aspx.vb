Imports System.Data.SqlClient
Imports System.Data


Partial Class timesheets
    Inherits GlobalClass

    Public sWorkersWithoutTime As String = ""
    Public PMList As String = GenerateListOfPMs()
    Public CurrentWorkerList As String = ""
    Public dtShifts As New DataTable
    Public dtOvertime As New DataTable
    Public dtEmployees As New DataTable

    Public timerStart As DateTime, timerFinish As DateTime, TotalTimerSpan As TimeSpan, TimerSpan As TimeSpan

    Sub page_load()

        Dim PMLoggedIn As String = Session("Initials")

        Dim StartDate As Date, EndDate As Date
        Try

            If IsPostBack = False Then

                MakeDateDropdown()
                MakeOfficeDropDown()
                MakeTypeDropDown()
                MakeShowAsPMDropDown()

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

                If Not Request.QueryString("page") Is Nothing Then
                    Session("ShowPage") = Request.QueryString("page")
                Else
                    Session("ShowPage") = 1
                End If

            Else
                Dim sDateRange As String = DDL_DateRange.SelectedValue
                Dim arrDateRange As Array = sDateRange.Split("-")

                StartDate = arrDateRange(0)
                EndDate = arrDateRange(1)

                Session("DateRangeStart") = StartDate
                Session("DateRangeEnd") = EndDate

                Session("ShowPage") = 1

            End If

        Catch ex As Exception
            LogError("timesheets.aspx.vb :: page_load", ex.ToString)
            Response.Write("ERROR: " & ex.ToString)
        End Try

        Session("SelectedOffice") = DDL_Offices.SelectedValue
        Session("SelectedType") = DDL_Type.SelectedValue

        Session("ShowAsPM") = Session("Initials")

        'If DDL_ShowAsPM.SelectedValue <> "0" Then
        Session("ShowAsPM") = DDL_ShowAsPM.SelectedValue
        'Else
        '    DDL_ShowAsPM.SelectedValue = Session("ShowAsPM")
        'End If

        LIT_WorkerList.Text = GenerateEmployeeWeekTotals(StartDate, EndDate, PMLoggedIn)

        LIT_ProjectList.Text = GenerateProjectWeekTotals(StartDate, EndDate)

        LIT_SyncData.Text = GenerateSyncPanel()


        LIT_ReportLink.Text = "[<a target=""_blank"" href=""printreport.aspx?startdate=" & StartDate & "&enddate=" & EndDate & """>Printable Payroll Report</a>]"

    End Sub

    Sub MakeDateDropdown()

        'Dim returnstring As String = ""
        Dim EarliestRecord As DateTime '= DateAdd(DateInterval.Month, -12, Now())

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
            LogError("timesheets.aspx.vb :: MakeDateDropdown", ex.ToString)
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

    Sub MakeOfficeDropDown()

        DDL_Offices.Items.Add(New ListItem("[All Offices]", 0))


        Dim qryOffices As String = "select OfficeName, OfficeID from tOffices order by OfficeName"
        Dim connOffices As New SqlConnection(sConnection)
        Dim CmdOffices As New SqlCommand(qryOffices, connOffices)
        Dim drOffices As SqlDataReader

        Try
            connOffices.Open()
            drOffices = CmdOffices.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drOffices.Read()

                DDL_Offices.Items.Add(New ListItem(drOffices("OfficeName"), drOffices("OfficeID")))

            End While
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("timesheets.aspx.vb :: MakeOfficeDropdown", ex.ToString)
        Finally
            connOffices.Close()

        End Try

        DDL_Offices.SelectedValue = Session("SelectedOffice")

        DDL_Offices.AutoPostBack = True

    End Sub

    Sub MakeTypeDropDown()

        DDL_Type.Items.Add(New ListItem("[Any]", 0))


        Dim qryType As String = "select distinct tEmployees.Type as Type from tEmployees"
        Dim connType As New SqlConnection(sConnection)
        Dim CmdType As New SqlCommand(qryType, connType)
        Dim drType As SqlDataReader

        Try
            connType.Open()
            drType = CmdType.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drType.Read()

                DDL_Type.Items.Add(New ListItem(drType("Type"), drType("Type")))

            End While
            drType.Close()
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("timesheets.aspx.vb :: MakeTypeDropdown", ex.ToString)
        Finally
            connType.Close()

        End Try

        DDL_Type.SelectedValue = Session("SelectedType")

        DDL_Type.AutoPostBack = True

    End Sub

    Sub MakeShowAsPMDropDown()

        DDL_ShowAsPM.Items.Add(New ListItem("[Show All]", 0))

        Dim qryPMs As String = "select distinct tEmployees.Initials as Initials from tEmployees where not Initials = '' and Active = 1 order by Initials"
        Dim connPMs As New SqlConnection(sConnection)
        Dim CmdPMs As New SqlCommand(qryPMs, connPMs)
        Dim drPMs As SqlDataReader

        Try
            connPMs.Open()
            drPMs = CmdPMs.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drPMs.Read()

                DDL_ShowAsPM.Items.Add(New ListItem(drPMs("Initials"), drPMs("Initials")))

            End While
            drPMs.Close()
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("timesheets.aspx.vb :: MakeShowAsPMDropDown", ex.ToString)
        Finally
            connPMs.Close()

        End Try

        DDL_ShowAsPM.SelectedValue = Session("ShowAsPM")
        DDL_ShowAsPM.AutoPostBack = True

        If Session("LimitedView") = 1 Then
            DDL_ShowAsPM.Enabled = False
        End If


    End Sub



    Function GenerateEmployeeWeekTotals(ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal PMLoggedIn As String) As String

        Dim Username As String, FirstName As String, LastName As String, PayrollID As String, WorkerID As Integer
        Dim strResult As String = "<table class=""listing"" id=""ListEmployees"" style=""width:500px;"">"
        Dim strEmployeeList As String = ""

        strResult = strResult & "<tr><th rowspan=""2"">Payroll&nbsp;#</th><th rowspan=""2"">Employee</th><th colspan=""5"" style=""text-align:center;"">Total Time</th><th colspan=""3"" style=""text-align:center;"">Unapproved</th></tr>"
        strResult = strResult & "<tr><th style=""text-align:center;"">Reg</th><th style=""text-align:center;"">OT</th><th style=""text-align:center;"">TT</th><th style=""text-align:center;"">No PM</th><th style=""text-align:center;"">No Job#</th><th style=""text-align:center;"">All</th><th style=""text-align:center;"">Me as PM</th><th style=""text-align:center;"">No PM</th></tr>"

        Dim OfficeFilter As String = DDL_Offices.SelectedValue
        Dim TypeFilter As String = DDL_Type.SelectedValue
        Dim sEmployeeQuery As String = ""

        Dim QueryFilter As String = "", PMFilter As String = ""

        If OfficeFilter <> "0" Then
            QueryFilter = QueryFilter & " and tEmployees.OfficeID=@OfficeID"
        End If

        If TypeFilter <> "0" Then
            QueryFilter = QueryFilter & " and tEmployees.Type=@Type"
        End If

        If Not Session("ShowAsPM") = "" And Not Session("ShowAsPM") Is Nothing And Not Session("ShowAsPM") = "0" Then
            PMFilter = " and fPM = '" & Session("ShowAsPM") & "'"
        End If

        Dim TotalEmployees As Integer = GetEmployeeCount(PMFilter, QueryFilter, StartDate, EndDate, OfficeFilter, TypeFilter)
        Dim StartRecordNumber As Integer = 0

        If Not Session("ShowPage") Is Nothing Then
            'StartRecordNumber = (Session("ShowPage") - 1) * 25 + 1
            StartRecordNumber = CInt(Session("ShowPage")) * 20 - 19
        End If

        If TypeFilter <> "0" Then
            sEmployeeQuery = "WITH OrderedEmployees AS (select EmployeeID, Username, FirstName, LastName, PayrollID, ROW_NUMBER() OVER (ORDER BY LastName) AS 'RowNumber' from tEmployees where Active = 1 and not LastName = '' " & PMFilter & " " & QueryFilter & ") SELECT * FROM OrderedEmployees WHERE RowNumber BETWEEN " & StartRecordNumber & " AND " & (StartRecordNumber + 19)
        Else
            sEmployeeQuery = "WITH OrderedEmployees AS (select EmployeeID, Username, FirstName, LastName, PayrollID, ROW_NUMBER() OVER (ORDER BY LastName) AS 'RowNumber' from tEmployees where EmployeeID in (select distinct UserPerformed from tShifts where ShiftID in (select distinct ShiftID from tTimeEntry where CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime)" & PMFilter & ") " & QueryFilter & ") SELECT * FROM OrderedEmployees WHERE RowNumber BETWEEN " & StartRecordNumber & " AND " & (StartRecordNumber + 19)
        End If


        'Response.Write(sEmployeeQuery)

        Dim connEmployees As New SqlConnection(sConnection)
        Dim cmdEmplyees As New SqlCommand(sEmployeeQuery, connEmployees)
        Dim drEmployees As SqlDataReader
        cmdEmplyees.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdEmplyees.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        If OfficeFilter <> "0" Then
            cmdEmplyees.Parameters.Add(New SqlParameter("@OfficeID", OfficeFilter))
        End If

        If TypeFilter <> "0" Then
            cmdEmplyees.Parameters.Add(New SqlParameter("@Type", TypeFilter))
        End If

        Try
            connEmployees.Open()

            Using daEmployees As New SqlDataAdapter(cmdEmplyees)
                daEmployees.Fill(dtEmployees)
            End Using

            'drEmployees = cmdEmplyees.ExecuteReader(Data.CommandBehavior.CloseConnection)

            'While drEmployees.Read

            '    WorkerID = drEmployees.Item("EmployeeID")
            '    Username = drEmployees.Item("Username")
            '    FirstName = drEmployees.Item("FirstName")
            '    LastName = drEmployees.Item("LastName")
            '    PayrollID = drEmployees.Item("PayrollID")

            '    'strResult = strResult & GetWorkerTime(WorkerID, StartDate, EndDate, Session("ShowAsPM")) & "</tr>" & vbCrLf

            '    If strEmployeeList = "" Then
            '        strEmployeeList = WorkerID
            '    Else
            '        strEmployeeList = strEmployeeList & "," & WorkerID
            '    End If

            'End While

            For Each Employee As DataRow In dtEmployees.Rows()
                If strEmployeeList = "" Then
                    strEmployeeList = Employee("EmployeeID")
                Else
                    strEmployeeList = strEmployeeList & "," & Employee("EmployeeID")
                End If
            Next

        Catch ex As Exception
            LogError("timesheets.aspx.vb :: GenerateEmployeeWeekTotals", ex.ToString)
            Response.Write("ERROR:" & ex.ToString & "<br>")
        Finally
            connEmployees.Close()
        End Try

        '------------ NEW

        If strEmployeeList.Length > 0 Then

            GetAllWorkersTime(strEmployeeList, StartDate, EndDate)
            GetAllOvertime(strEmployeeList, StartDate, EndDate)

            strResult = strResult & MakeTimeListing(strEmployeeList, PMLoggedIn, StartDate, EndDate)
        Else
            strResult = strResult & "<tr><td colspan=""10"">No users matching the criteria</td></tr>"
        End If

        '----------------

        strResult = strResult & "<tr><td colspan=""10"" class=""pagination"">Page:&nbsp;&nbsp;"

        Dim PageCounter As Integer = 0

        While PageCounter < (TotalEmployees / 20)
            PageCounter = PageCounter + 1

            If PageCounter = Session("ShowPage") Then
                strResult = strResult & " <span class=""current"">" & PageCounter & "</span>&nbsp;&nbsp;"
            Else
                strResult = strResult & " <a href=""timesheets-test2.aspx?page=" & PageCounter & """>" & PageCounter & "</a>&nbsp;&nbsp;"
            End If


        End While



        strResult = strResult & "</td></tr>"
        If TypeFilter = "0" Then
            strResult = strResult & "<tr><td colspan=""10"" style=""border-top:1px dashed #3F317A""><strong>Users that have not entered any time this week:</strong><br>" & EmployeesWithNoTime(StartDate, EndDate) & "</td></tr>"
        End If
        strResult = strResult & "</table>"



        Return strResult

    End Function

    Sub GetAllWorkersTime(sListOfWorkers As String, StartDate As DateTime, EndDate As DateTime)

        Dim sTimeQuery As String = "select startTime, endTime, fPM, Approved, fJobNumber, fTravelTime, UserPerformed, FirstName, LastName, PayrollID from tTimeEntry, tShifts, tEmployees where tShifts.UserPerformed in (" & sListOfWorkers & ") and (not tShifts.fInjured is NULL) and tShifts.ShiftID = tTimeEntry.ShiftID and tEmployees.EmployeeId = tShifts.UserPerformed and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Try
            connTimeQuery.Open()
            Using daShifts As New SqlDataAdapter(cmdTimeQuery)
                daShifts.Fill(dtShifts)
            End Using
        Catch ex As Exception
            LogError("timesheets.aspx.vb :: GetAllWorkersTime", ex.ToString)
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

    End Sub

    Sub GetAllOvertime(sListOfWorkers As String, StartDate As DateTime, EndDate As DateTime)

        Dim sOvertimeQuery As String = "select EmployeeID, IsNull(sum(Overtime),0) as TotalOvertime from tTimeAdjustments where DateAdjusted >= @StartTime and DateAdjusted <= @EndTime and EmployeeID in (" & sListOfWorkers & ") group by EmployeeID"
        Dim connOvertimeQuery As New SqlConnection(sConnection)
        Dim cmdOvertimeQuery As New SqlCommand(sOvertimeQuery, connOvertimeQuery)
        cmdOvertimeQuery.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdOvertimeQuery.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Try
            connOvertimeQuery.Open()
            Using daOvertime As New SqlDataAdapter(cmdOvertimeQuery)
                daOvertime.Fill(dtOvertime)
            End Using
        Catch ex As Exception
            LogError("timesheets.aspx.vb :: GetAllOvertime", ex.ToString)
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connOvertimeQuery.Close()
        End Try

    End Sub

    Function MakeTimeListing(sListOfWorkers As String, PMLoggedIn As String, DateRangeStart As DateTime, DateRangeEnd As DateTime) As String

        Dim HilightClass As String = "", PayrollID As String = "", FirstName As String = "", LastName As String = ""
        Dim Approved As Integer, fPM As String, StartTime As DateTime, EndTime As DateTime, fJobNumber As String = "", fTravelTime As String = ""
        Dim returnstring As String = ""
        Dim TimeSpan As Double

        Dim workers As String() = sListOfWorkers.Split(New Char() {","c})

        Dim worker As String
        For Each Employee As DataRow In dtEmployees.Rows()

            PayrollID = Employee("PayrollID")
            FirstName = Employee("FirstName")
            LastName = Employee("LastName")
            worker = Employee("EmployeeID")

            returnstring = returnstring & "<tr " & HilightClass & "><td>" & PayrollID & "</td><td style=""width:120px""><a href=""worker.aspx?id=" & worker & "&startdate=" & DateRangeStart & "&enddate=" & DateRangeEnd & """>" & LastName & ", " & FirstName & "</a></td>"


            Dim rows() As DataRow = dtShifts.Select("UserPerformed=" & worker)

            If rows.Length > 0 Then



                Dim TotalTime As Double = 0, TotalUnapproved As Double = 0, MeAsPM As Double = 0, NoPMTotal As Double = 0, NoPMUnapproved As Double = 0, NoJobTotal As Double = 0, OvertimeTotal As Double = 0, RegularTimeTotal As Double = 0, TravelTimeTotal As Double = 0

                Dim OvertimeRows() As DataRow = dtOvertime.Select("EmployeeID=" & worker)

                If OvertimeRows.Length > 0 Then
                    OvertimeTotal = OvertimeRows(0).Item("TotalOvertime")
                End If



                If Not rows Is Nothing Then

                    For Each row As DataRow In rows

                        StartTime = row.Item("StartTime")
                        EndTime = row.Item("EndTime")
                        fPM = row.Item("fPM")
                        Approved = row.Item("Approved")
                        fJobNumber = row.Item("fJobNumber")
                        fTravelTime = row.Item("fTravelTime")
                        fPM = CheckForPM(fPM)

                        TimeSpan = Math.Round(DateDiff(DateInterval.Minute, StartTime, EndTime) / 60, 2)
                        TotalTime = TotalTime + TimeSpan

                        If Approved <> 1 Then TotalUnapproved = TotalUnapproved + TimeSpan

                        If fPM = "" Then
                            NoPMTotal = NoPMTotal + TimeSpan
                            If Approved <> 1 Then
                                NoPMUnapproved = NoPMUnapproved + TimeSpan
                            End If
                        End If

                        If fJobNumber = "" Then
                            NoJobTotal = NoJobTotal + TimeSpan
                        End If

                        If fTravelTime = "Personal" Or fTravelTime = "Company" Then
                            TravelTimeTotal = TravelTimeTotal + TimeSpan
                        End If

                        If UCase(Trim(fPM)) = PMLoggedIn And Approved <> 1 Then MeAsPM = MeAsPM + TimeSpan


                    Next

                    HilightClass = ""

                    'If CheckUnbillableTime(WorkerID, StartDate, EndDate) = True Then
                    '    HilightClass = "class=""hilight"""
                    'End If

                    'If CheckTimeOff(WorkerID, StartDate, EndDate) = True Then
                    '    HilightClass = "class=""hilight2"""
                    'End If

                    'If CheckAVJob(WorkerID, StartDate, EndDate) = True Then
                    '    HilightClass = "class=""hilight3"""
                    'End If

                    'If CheckSecurityJob(WorkerID, StartDate, EndDate) = True Then
                    '    HilightClass = "class=""hilight4"""
                    'End If





                    RegularTimeTotal = TotalTime - OvertimeTotal - TravelTimeTotal

                    returnstring = returnstring & "<td>" & RegularTimeTotal & "</td>"

                    returnstring = returnstring & "<td>" & OvertimeTotal & "</td>"

                    returnstring = returnstring & "<td>" & TravelTimeTotal & "</td>"

                    If NoPMTotal > 0 Then
                        returnstring = returnstring & "<td><strong>" & NoPMTotal & "</strong></td>"
                    Else
                        returnstring = returnstring & "<td>" & NoPMTotal & "</td>"
                    End If

                    If NoJobTotal > 0 Then
                        returnstring = returnstring & "<td><strong>" & NoJobTotal & "</strong></td>"
                    Else
                        returnstring = returnstring & "<td>" & NoJobTotal & "</td>"
                    End If

                    If TotalUnapproved > 0 Then
                        returnstring = returnstring & "<td><strong>" & TotalUnapproved & "</strong></td>"
                    Else
                        returnstring = returnstring & "<td>" & TotalUnapproved & "</td>"
                    End If

                    If MeAsPM > 0 Then
                        returnstring = returnstring & "<td><strong>" & MeAsPM & "</strong></td>"
                    Else
                        returnstring = returnstring & "<td>" & MeAsPM & "</td>"
                    End If

                    If NoPMUnapproved > 0 Then
                        returnstring = returnstring & "<td><strong>" & NoPMUnapproved & "</strong></td>"
                    Else
                        returnstring = returnstring & "<td>" & NoPMUnapproved & "</td>"
                    End If



                End If
            Else
                returnstring = returnstring & "<td colspan=""8""></td>"
            End If
            returnstring = returnstring & "</tr>"

        Next

        Return returnstring

    End Function

    Function GetWorkerTime(ByVal UserID As Integer, ByVal StartDate As DateTime, ByVal EndDate As DateTime, ByVal PMLoggedIn As String) As String

        Dim returnstring As String = "", TotalTime As Double, TotalUnapproved As Double, MeAsPM As Double = 0, NoPMTotal As Double = 0, NoPMUnapproved As Double = 0, NoJobTotal As Double = 0, OvertimeTotal As Double = 0, RegularTimeTotal As Double = 0, TravelTimeTotal As Double = 0
        Dim TimeSpan As Double

        Dim Approved As Integer, fPM As String, StartTime As DateTime, EndTime As DateTime, fJobNumber As String = "", fTravelTime As String = ""

        OvertimeTotal = GetWeeklyOvertime(StartDate, EndDate, UserID)

        Dim sTimeQuery As String = "select startTime, endTime, fPM, Approved, fJobNumber, fTravelTime from tTimeEntry, tShifts where tShifts.UserPerformed = @UserID and (not tShifts.fInjured is NULL) and tShifts.ShiftID = tTimeEntry.ShiftID and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@UserID", UserID))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndTime", EndDate))
        Dim drTimeQuery As SqlDataReader

        Try
            connTimeQuery.Open()

            drTimeQuery = cmdTimeQuery.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drTimeQuery.Read()

                StartTime = drTimeQuery.Item("StartTime")
                EndTime = drTimeQuery.Item("EndTime")
                fPM = drTimeQuery.Item("fPM")
                Approved = drTimeQuery.Item("Approved")
                fJobNumber = drTimeQuery.Item("fJobNumber")
                fTravelTime = drTimeQuery.Item("fTravelTime")

                fPM = CheckForPM(fPM)

                TimeSpan = Math.Round(DateDiff(DateInterval.Minute, StartTime, EndTime) / 60, 2)
                TotalTime = TotalTime + TimeSpan

                If Approved <> 1 Then TotalUnapproved = TotalUnapproved + TimeSpan

                If fPM = "" Then
                    NoPMTotal = NoPMTotal + TimeSpan
                    If Approved <> 1 Then
                        NoPMUnapproved = NoPMUnapproved + TimeSpan
                    End If
                End If

                If fJobNumber = "" Then
                    NoJobTotal = NoJobTotal + TimeSpan
                End If

                If fTravelTime = "Personal" Or fTravelTime = "Company" Then
                    TravelTimeTotal = TravelTimeTotal + TimeSpan
                End If

                If UCase(Trim(fPM)) = PMLoggedIn And Approved <> 1 Then MeAsPM = MeAsPM + TimeSpan

                'If fPM = "" And Approved <> 1 Then NoPM = NoPM + TimeSpan

            End While

        Catch ex As Exception
            LogError("timesheets.aspx.vb :: GetWorkerTime", ex.ToString)
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try



        RegularTimeTotal = TotalTime - OvertimeTotal - TravelTimeTotal

        returnstring = returnstring & "<td>" & RegularTimeTotal & "</td>"

        returnstring = returnstring & "<td>" & OvertimeTotal & "</td>"

        returnstring = returnstring & "<td>" & TravelTimeTotal & "</td>"

        If NoPMTotal > 0 Then
            returnstring = returnstring & "<td><strong>" & NoPMTotal & "</strong></td>"
        Else
            returnstring = returnstring & "<td>" & NoPMTotal & "</td>"
        End If

        If NoJobTotal > 0 Then
            returnstring = returnstring & "<td><strong>" & NoJobTotal & "</strong></td>"
        Else
            returnstring = returnstring & "<td>" & NoJobTotal & "</td>"
        End If

        If TotalUnapproved > 0 Then
            returnstring = returnstring & "<td><strong>" & TotalUnapproved & "</strong></td>"
        Else
            returnstring = returnstring & "<td>" & TotalUnapproved & "</td>"
        End If

        If MeAsPM > 0 Then
            returnstring = returnstring & "<td><strong>" & MeAsPM & "</strong></td>"
        Else
            returnstring = returnstring & "<td>" & MeAsPM & "</td>"
        End If

        If NoPMUnapproved > 0 Then
            returnstring = returnstring & "<td><strong>" & NoPMUnapproved & "</strong></td>"
        Else
            returnstring = returnstring & "<td>" & NoPMUnapproved & "</td>"
        End If


        GetWorkerTime = returnstring

    End Function




    Function GetWeeklyOvertime(ByVal StartDate As Date, ByVal EndDate As Date, ByVal EmployeeID As Integer) As Double

        Dim OvertimeTotal As Double

        Dim sTimeQuery As String = "select IsNull(sum(Overtime),0) as TotalOvertime from tTimeAdjustments where DateAdjusted >= @StartDate and DateAdjusted <= @EndDate and EmployeeID = @EmployeeID"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndDate", EndDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))

        Dim drTimeQuery As SqlDataReader

        Try
            connTimeQuery.Open()

            drTimeQuery = cmdTimeQuery.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drTimeQuery.Read()
                OvertimeTotal = drTimeQuery.Item("TotalOvertime")
            End While

        Catch ex As Exception
            LogError("timesheets.aspx.vb :: GetWeeklyOvertime", ex.ToString)
            Response.Write("ERROR: " & ex.Message & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

        Return OvertimeTotal


    End Function

    Function EmployeesWithNoTime(ByVal StartDate As DateTime, ByVal EndDate As DateTime) As String
        Dim returnstring As String = ""

        'If EmployeesWithTime = "" Then
        '    returnstring = ""
        'Else


        Dim sNoTimeQuery As String = "select FirstName, LastName from tEmployees where Active = 1 and Worker = 1 and not EmployeeID in (select distinct UserPerformed from tShifts, tTimeEntry where tShifts.ShiftID = tTimeEntry.ShiftID and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime) order by LastName asc"
        Dim connNoTime As New SqlConnection(sConnection)
        Dim cmdNoTime As New SqlCommand(sNoTimeQuery, connNoTime)
        Dim drNoTime As SqlDataReader
        cmdNoTime.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdNoTime.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Try
            connNoTime.Open()
            drNoTime = cmdNoTime.ExecuteReader(Data.CommandBehavior.CloseConnection)


            While drNoTime.Read()
                If returnstring = "" Then
                    returnstring = drNoTime.Item("FirstName") & " " & drNoTime.Item("LastName")
                Else
                    returnstring = returnstring & ", " & drNoTime.Item("FirstName") & " " & drNoTime.Item("LastName")

                End If
            End While

        Catch ex As Exception
            LogError("timesheets.aspx.vb :: EmployeesWithNoTime", ex.ToString)
            returnstring = returnstring & "ERROR: " & ex.ToString
        Finally
            connNoTime.Close()
        End Try
        'End If


        Return "<div id=""userswithouttime"">" & returnstring & "</div>"

    End Function

    Function GenerateProjectWeekTotals(ByVal StartDate As Date, ByVal enddate As Date) As String


        Dim JobNumber As String, TotalJobTime As Double
        Dim strResult As String = "<table class=""listing"" style=""width:200px;"" id=""ListEmployees"">"
        Dim strEmployeeList As String = ""
        strResult = strResult & "<tr><th>Job Number</th><th>Total Time</th></tr>"

        Dim sEmployeeQuery As String = "select distinct fJobNumber, sum(DateDiff(n,startTime, endTime)) as TotalTime from tShifts, tTimeEntry where tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.ShiftID in (select distinct ShiftID from tTimeEntry where CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime) and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime group by fJobNumber order by fJobNumber"
        Dim connEmployees As New SqlConnection(sConnection)
        Dim cmdEmplyees As New SqlCommand(sEmployeeQuery, connEmployees)
        Dim drEmployees As SqlDataReader
        cmdEmplyees.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdEmplyees.Parameters.Add(New SqlParameter("@EndTime", enddate))

        Try
            connEmployees.Open()

            drEmployees = cmdEmplyees.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drEmployees.Read

                JobNumber = drEmployees.Item("fJobNumber")
                TotalJobTime = drEmployees.Item("TotalTime") / 60

                strResult = strResult & "<tr><td><a href=""jobreport.aspx?startdate=" & StartDate & "&enddate=" & enddate & "&jobnumber=" & JobNumber & """>" & JobNumber & "</a></td><td>" & TotalJobTime & "</td></tr>"

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

    Function GetTotalHours(ByVal JobNumber As String, ByVal StartDate As Date, ByVal EndDate As Date) As String

        Dim returnstring As String = "", TotalTime As Double, StartTime As DateTime, EndTime As DateTime
        Dim TimeSpan As Double

        Dim sTimeQuery As String = "select startTime, endTime from tTimeEntry, tShifts where tShifts.fJobNumber = @JobNumber and tShifts.ShiftID = tTimeEntry.ShiftID and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndTime", EndDate))

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

            returnstring = TotalTime

        Catch ex As Exception
            LogError("timesheets.aspx.vb :: GetTotalHours", ex.ToString)
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

        Return returnstring

    End Function

    Function GenerateSyncPanel() As String

        Dim LastSyncTime As DateTime
        Dim Datecolor As String = "#000"

        Dim qLastSync As String = "select top 1 RequestTime from tSOAPLog order by requestTime Desc"
        Dim connLastSync As New SqlConnection(sConnection)
        Dim cmdLastSync As New SqlCommand(qLastSync, connLastSync)
        Dim drLastSync As SqlDataReader

        Try
            connLastSync.Open()
            drLastSync = cmdLastSync.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drLastSync.Read
                LastSyncTime = drLastSync.Item("RequestTime")
            End While
        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString)
            LogError("timesheets.aspx.vb :: GenerateSyncPanel", ex.ToString)
        Finally
            connLastSync.Close()
        End Try

        If DateAdd(DateInterval.Hour, 24, LastSyncTime) < Now() Then Datecolor = "#f00"

        Dim ReturnString As String = ""
        ReturnString = ReturnString & "Last FFM Update:<br/> <strong style=""color:" & Datecolor & """>" & LastSyncTime & "</strong><br>"
        ReturnString = ReturnString & "<input type=""button"" Onclick=""javascript:window.open('LoadData.aspx?action=sync','Spreadsheet'); "" value=""Sync Now"" />"
        Return ReturnString

    End Function

    Function CheckForPM(ByVal OriginalPM As String) As String

        'Dim PMCount As Integer

        If InStr(PMList, UCase(OriginalPM)) <> 0 Then
            Return OriginalPM
        Else
            Return ""
        End If

        'Dim qCheckPM As String = "select count(EmployeeID) as PMCount from tEmployees where Initials=@PM"
        'Dim connCheckPM As New SqlConnection(sConnection)
        'Dim CmdCheckPM As New SqlCommand(qCheckPM, connCheckPM)
        'Dim drCheckPM As SqlDataReader
        'CmdCheckPM.Parameters.Add(New SqlParameter("@PM", Trim(OriginalPM)))

        'Try
        '    connCheckPM.Open()
        '    drCheckPM = CmdCheckPM.ExecuteReader(Data.CommandBehavior.CloseConnection)
        '    While drCheckPM.Read
        '        PMCount = drCheckPM.Item("PMCount")
        '    End While
        'Catch ex As Exception
        '    LogError("timesheets.aspx.vb :: CheckForPM", ex.ToString)
        '    Response.Write("ERROR: " & ex.Message)
        'Finally
        '    connCheckPM.Close()
        'End Try

        'If PMCount > 0 Then
        '    Return OriginalPM
        'Else
        '    Return ""
        'End If


    End Function

    Function GenerateListOfPMs() As String

        Dim PMList As String = ""

        Dim qCheckPM As String = "select distinct Initials from tEmployees where not Initials = '' order by Initials"
        Dim connCheckPM As New SqlConnection(sConnection)
        Dim CmdCheckPM As New SqlCommand(qCheckPM, connCheckPM)
        Dim drCheckPM As SqlDataReader
        'CmdCheckPM.Parameters.Add(New SqlParameter("@PM", Trim(OriginalPM)))

        Try
            connCheckPM.Open()
            drCheckPM = CmdCheckPM.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheckPM.Read

                If PMList = "" Then
                    PMList = drCheckPM.Item(0)
                Else
                    PMList = PMList & "," & UCase(drCheckPM.Item(0))
                End If

            End While
        Catch ex As Exception
            LogError("timesheets.aspx.vb :: CheckForPM", ex.ToString)
            Response.Write("ERROR: " & ex.Message)
        Finally
            connCheckPM.Close()
        End Try

        Return PMList

    End Function

    Function CheckUnbillableTime(ByVal WorkerID As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean

        Dim HasUnbillableTime As Boolean = False



        Dim qCheckUnbill As String = "select count(tShifts.ShiftID) as ShiftCount from tShifts, tTimeEntry where tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.UserPerformed = @UserID and (not tShifts.fInjured is NULL) and Approved <> 1 and StartTime = EndTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"

        Dim connCheckUnbill As New SqlConnection(sConnection)
        Dim CmdCheckUnbill As New SqlCommand(qCheckUnbill, connCheckUnbill)
        Dim drCheckUnbill As SqlDataReader
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@UserID", WorkerID))
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Try
            connCheckUnbill.Open()
            drCheckUnbill = CmdCheckUnbill.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheckUnbill.Read
                If drCheckUnbill.Item("ShiftCount") > 0 Then
                    HasUnbillableTime = True
                End If
            End While
        Catch ex As Exception
            LogError("timesheets.aspx.vb :: CheckUnbillableTime", ex.ToString)
            Response.Write("ERROR: " & ex.Message)
        Finally
            connCheckUnbill.Close()
        End Try

        Return HasUnbillableTime

    End Function

    Function CheckTimeOff(ByVal WorkerID As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean

        Dim HasUnbillableTime As Boolean = False

        Dim qCheckUnbill As String = "select count(tShifts.ShiftID) as ShiftCount from tShifts, tTimeEntry where tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.UserPerformed = @UserID and (not tShifts.fInjured is NULL) and Approved <> 1 and tShifts.fShiftType in ('Vacation','Sick','Personal','PTO') and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"

        Dim connCheckUnbill As New SqlConnection(sConnection)
        Dim CmdCheckUnbill As New SqlCommand(qCheckUnbill, connCheckUnbill)
        Dim drCheckUnbill As SqlDataReader
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@UserID", WorkerID))
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Try
            connCheckUnbill.Open()
            drCheckUnbill = CmdCheckUnbill.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheckUnbill.Read
                If drCheckUnbill.Item("ShiftCount") > 0 Then
                    HasUnbillableTime = True
                End If
            End While
        Catch ex As Exception
            LogError("timesheets.aspx.vb :: CheckTimeOff", ex.ToString)
            Response.Write("ERROR: " & ex.Message)
        Finally
            connCheckUnbill.Close()
        End Try

        Return HasUnbillableTime

    End Function

    Function CheckAVJob(ByVal WorkerID As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean

        Dim HasUnbillableTime As Boolean = False

        Dim qCheckUnbill As String = "select count(tShifts.ShiftID) as ShiftCount from tShifts, tTimeEntry where tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.UserPerformed = @UserID and (not tShifts.fInjured is NULL) and Approved <> 1 and (tShifts.fJobNumber like ('01-15%') OR tShifts.fJobNumber like ('0115%')) and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"

        Dim connCheckUnbill As New SqlConnection(sConnection)
        Dim CmdCheckUnbill As New SqlCommand(qCheckUnbill, connCheckUnbill)
        Dim drCheckUnbill As SqlDataReader
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@UserID", WorkerID))
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Try
            connCheckUnbill.Open()
            drCheckUnbill = CmdCheckUnbill.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheckUnbill.Read
                If drCheckUnbill.Item("ShiftCount") > 0 Then
                    HasUnbillableTime = True
                End If
            End While
        Catch ex As Exception
            LogError("timesheets.aspx.vb :: CheckAVJob", ex.ToString)
            Response.Write("ERROR: " & ex.Message)
        Finally
            connCheckUnbill.Close()
        End Try

        Return HasUnbillableTime

    End Function

    Function CheckSecurityJob(ByVal WorkerID As Integer, ByVal StartDate As Date, ByVal EndDate As Date) As Boolean

        Dim HasUnbillableTime As Boolean = False

        Dim qCheckUnbill As String = "select count(tShifts.ShiftID) as ShiftCount from tShifts, tTimeEntry where tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.UserPerformed = @UserID and (not tShifts.fInjured is NULL) and Approved <> 1 and (tShifts.fJobNumber like ('01-14%') OR tShifts.fJobNumber like ('0114%')) and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"

        Dim connCheckUnbill As New SqlConnection(sConnection)
        Dim CmdCheckUnbill As New SqlCommand(qCheckUnbill, connCheckUnbill)
        Dim drCheckUnbill As SqlDataReader
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@UserID", WorkerID))
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        CmdCheckUnbill.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Try
            connCheckUnbill.Open()
            drCheckUnbill = CmdCheckUnbill.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheckUnbill.Read
                If drCheckUnbill.Item("ShiftCount") > 0 Then
                    HasUnbillableTime = True
                End If
            End While
        Catch ex As Exception
            LogError("timesheets.aspx.vb :: CheckAVJob", ex.ToString)
            Response.Write("ERROR: " & ex.Message)
        Finally
            connCheckUnbill.Close()
        End Try

        Return HasUnbillableTime

    End Function

    Function GetEmployeeCount(PMFilter As String, QueryFilter As String, StartDate As DateTime, EndDate As DateTime, OfficeFilter As String, TypeFilter As String) As Integer

        Dim EmployeeCount As Integer = 0, sEmployeeQuery As String = ""

        If TypeFilter <> "0" Then
            sEmployeeQuery = "select count(EmployeeID) as EmployeeCount from tEmployees where Active = 1 and not LastName = '' " & QueryFilter
        Else
            sEmployeeQuery = "select count(EmployeeID) as EmployeeCount from tEmployees where EmployeeID in (select distinct UserPerformed from tShifts where ShiftID in (select distinct ShiftID from tTimeEntry where CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime)" & PMFilter & ") " & QueryFilter
        End If

        Dim connEmployees As New SqlConnection(sConnection)
        Dim cmdEmplyees As New SqlCommand(sEmployeeQuery, connEmployees)
        Dim drEmployees As SqlDataReader
        cmdEmplyees.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdEmplyees.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        If OfficeFilter <> "0" Then
            cmdEmplyees.Parameters.Add(New SqlParameter("@OfficeID", OfficeFilter))
        End If

        If TypeFilter <> "0" Then
            cmdEmplyees.Parameters.Add(New SqlParameter("@Type", TypeFilter))
        End If

        Try
            connEmployees.Open()

            drEmployees = cmdEmplyees.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drEmployees.Read

                EmployeeCount = drEmployees.Item("EmployeeCount")

            End While

        Catch ex As Exception
            LogError("timesheets.aspx.vb :: GetEmployeeCount", ex.ToString)
            Response.Write("ERROR:" & ex.ToString & "<br>")
        Finally
            connEmployees.Close()
        End Try

        Return EmployeeCount

    End Function

End Class
