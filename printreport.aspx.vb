Imports System.Data.SqlClient

Partial Class printreport
    Inherits GlobalClass

    Sub page_load()
        LIT_Page.Text = EmployeeList()
    End Sub


    Function EmployeeList() As String

        Dim StartDate As String = Request.QueryString("startdate")
        Dim EndDate As String = Request.QueryString("enddate")
        Dim returnString As String = ""

        Dim sEmployeeQuery As String = "select EmployeeID, FirstName, LastName from tEmployees where Active = 1 and Worker = 1 order by LastName"
        Dim connEmployees As New SqlConnection(sConnection)
        Dim cmdEmployees As New SqlCommand(sEmployeeQuery, connEmployees)
        Dim drEmployees As SqlDataReader
        Try
            connEmployees.Open()
            drEmployees = cmdEmployees.ExecuteReader()

            While drEmployees.Read()

                returnString = returnString & GenerateWeekTotal(drEmployees("EmployeeID"), StartDate, EndDate)

            End While

        Catch ex As Exception
            Response.Write("error:<br /> " & ex.ToString)
        Finally
            connEmployees.Close()
        End Try

        EmployeeList = returnString

    End Function


    Function GenerateWeekTotal(ByVal EmployeeID As Integer, ByVal StartDate As DateTime, ByVal EndDate As DateTime) As String

        Dim ReturnString As String = ""

        Dim FirstName As String = "", LastName As String = "", Username As String = "", PayrollID As String = ""
        Dim StartLatitude As String = "", StartLongitude As String = "", EndLatitude As String = "", EndLongitude As String = ""
        Dim StartMap As String = "", EndMap As String = ""
        Dim ApprovedInitials As String = ""


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
            returnString = returnString & "ERROR: " & ex.ToString
        Finally
            connWorker.Close()
        End Try

        ReturnString = ReturnString & "<table class=""info""><tr><td style=""width:500px;"">"

        returnString = returnString & "<h3>Employee Name: " & FirstName & " " & LastName & "</h3>" & vbCrLf
        returnString = returnString & "<h3>Employee/Payroll #: " & PayrollID & "</h3>" & vbCrLf
        returnString = returnString & "<h3>" & StartDate & " - " & EndDate & "</h3>" & vbCrLf & vbCrLf

        returnString = returnString & "</td>"

        ReturnString = ReturnString & "<td colspan=""6"" style=""width:360px;""><h3>Notes:</h3><div class=""weeknotes"">" & GetWeekNotes(EmployeeID, StartDate) & "</div>"
        ReturnString = ReturnString & "</table>"


        Dim TimeEntryID As Integer, fPM As String, fJobNumber As String, fState As String, StartTime As DateTime, EndTime As DateTime, fTravelTime As String, fPerDiem As String, Approved As Integer, StartAddress As String, EndAddress As String, ShiftID As Integer, fType As String, fShiftType As String, fInjured As String, fComments As String, fProjectName As String, fPW As String
        Dim LoopDate As Date, DayName As String = "", Hours As Double, LastDay As String = "", LastShift As Integer = 0, ShiftCounter As Integer = 1, Overtime As String, JobList As String = ""
        Dim WeekTotalTime As Double = 0, WeekOvertime As Double = 0, WeekTravelTime As Double = 0, WeekStandardTime As Double = 0
        Dim EditClass As String = "", EditClassDate As String = "", EditClassType As String = "", EditClassShiftType As String = "", EditClassTravel As String = "", EditClassPerDiem As String = "", EditClassInjured As String = "", EditClassComments As String = "", EditClassPW As String = "", EditClassOT As String = "", EditClassOTLock As String = ""


        Dim TimeSheetQuery As String = "select TimeEntryID, fPM, fJobNumber, fState, StartTime, EndTime, fTravelTime, fPerDiem, Approved, StartAddress, EndAddress, tTimeEntry.ShiftID, fType, fShiftType, fInjured, fComments, fProjectName, StartLatitude, StartLongitude, EndLatitude, EndLongitude, fPW, Initials from tTimeEntry, tShifts, tEmployees where tEmployees.EmployeeID = isNull(tTimeEntry.ApprovedBy,1) and tShifts.ShiftID = tTimeEntry.ShiftID and tShifts.UserPerformed = @EmployeeID and not fInjured is null and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartDate and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndDate order by StartTime asc"
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

                Dim sStartTime As String, sEndTime As String

                sStartTime = FormatTime(StartTime)
                sEndTime = FormatTime(EndTime)

                LoopDate = StartTime.Date
                DayName = LoopDate.DayOfWeek.ToString

                If StartAddress <> "" Then
                    StartMap = "<img src=""images/map.png"" title=""" & StartAddress & """/>"
                Else
                    StartMap = ""
                End If

                If EndAddress <> "" Then
                    EndMap = "<img src=""images/map.png""  title=""" & EndAddress & """/>"
                Else
                    EndMap = ""
                End If

                Hours = DateDiff(DateInterval.Minute, StartTime, EndTime) / 60

                If Not DayName = LastDay Then
                    'WRITE DAY HEADING
                    ReturnString = ReturnString & "<tr><th colspan=13 class=""dayheading"">" & DayName & " " & LoopDate & "</th></tr>"
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


                    ReturnString = ReturnString & "<tr class=""shift label""><td rowspan=2>Shift<br/>" & ShiftCounter & "</td><td>PM</td><td>State</td><td>Job&nbsp;#</td><td>PW</td><td>Project</td><td style=""width:80px;"">Shift&nbsp;Type</td><td>Travel</td><td>Per&nbsp;Diem</td><td>Injured</td><td>Comments</td>"

                    If Overtime = "NoOT" Then
                        returnString = returnString & "<td>&nbsp;</td>"
                    Else
                        returnString = returnString & "<td style=""background-color:#A01D1E;color:#fff; font-weight:bold;text-align:center;"">OT&nbsp;Hrs</td>"
                    End If
                    returnString = returnString & "<td></td></tr>"

                    returnString = returnString & "<tr class=""shift"">"
                    ReturnString = ReturnString & "<td>" & UCase(fPM) & "</td>"
                    ReturnString = ReturnString & "<td>" & UCase(fState) & "</td>"
                    ReturnString = ReturnString & "<td>" & fJobNumber & "</td>"
                    ReturnString = ReturnString & "<td>" & fPW & "</td>"
                    ReturnString = ReturnString & "<td>" & fProjectName & "</td>"
                    ReturnString = ReturnString & "<td>" & fShiftType & "</td>"
                    ReturnString = ReturnString & "<td>" & fTravelTime & "</td>"
                    ReturnString = ReturnString & "<td>" & fPerDiem & "</td>"
                    ReturnString = ReturnString & "<td>" & fInjured & "</td>"
                    ReturnString = ReturnString & "<td>" & fComments & "</td>"

                    If Overtime = "NoOT" Then
                        returnString = returnString & "<td>&nbsp;</td>"
                    Else
                        ReturnString = ReturnString & "<td style=""background-color:#A01D1E;color:#fff; font-weight:bold;text-align:center;"">" & Overtime & "</td>"
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
                ReturnString = ReturnString & "<span>" & sStartTime & "</span>"
                returnString = returnString & StartMap & "</td>"
                returnString = returnString & "<td colspan=5>END:"
                ReturnString = ReturnString & "<span>" & sEndTime & "</span>"
                returnString = returnString & EndMap & "</td>"
                returnString = returnString & "<td colspan=2>&nbsp;</td>"
                ReturnString = ReturnString & "<td>"


                If Approved = 1 Then
                    ReturnString = ReturnString & "<img src=""images/tick.png"" /><strong>" & ApprovedInitials & "</strong>"

                End If

                ReturnString = ReturnString & "</td>" & vbCrLf
                ReturnString = ReturnString & "</tr>"

                LastDay = DayName
                LastShift = ShiftID

            End While

            returnString = returnString & "<tr class=""totalsrow""><td></td><td colspan=""6""><strong>Week Totals</strong><br />"
            returnString = returnString & "TOTAL TIME: <strong>" & WeekTotalTime & "</strong><br />"
            returnString = returnString & "STANDARD TIME: <strong>" & WeekStandardTime & "</strong><br />"
            ReturnString = ReturnString & "OVERTIME: <strong>" & WeekOvertime & "</strong><br />"
            returnString = returnString & "TRAVEL TIME: <strong>" & WeekTravelTime & "</strong>"

            returnString = returnString & "<td colspan=""6"">"
            ReturnString = ReturnString & "</td></tr></table>"



        Catch ex As Exception
            returnString = returnString & "ERROR: " & ex.ToString
        Finally
            connTimeSheet.Close()

        End Try

        GenerateWeekTotal = ReturnString


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
            LogError("ax-edittimesheet.aspx.vb :: SaveNotes", ex.ToString)
            Response.Write("ERROR: " & ex.ToString)
        Finally
            ConnCheckNotes.Close()
        End Try

        Return NotesText

    End Function


End Class
