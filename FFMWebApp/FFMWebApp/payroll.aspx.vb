Imports System.Data.SqlClient
Imports System.Data

Partial Class payroll
    Inherits GlobalClass

    Dim RecTotStandard As Double = 0, RecTotOvertime As Double = 0, RecTotCompany As Double = 0, RecTotPersonal As Double = 0, RecTotPerDiem As Double = 0, RecTotPerDiemDays As Integer = 0
    Dim RecTotCompanyAdjusted As Double, recTotPersonalAdjusted As Double, RecTotPerDiemAdjusted As Double


    Dim multipilerCompany As Double = 4.0
    Dim multiplierPersonal As Double = 7.0
    Dim multiplierPerDiem As Double
    Dim multiplierPerDiemDaily As Double = 44

    Public sEmployeeList As String = ""
    Public dtEmployees As New DataTable
    Public dtJobs As New DataTable
    Public dtAdjustments As New DataTable
    Public dtTimeEntries As New DataTable
    Public dtBonus As New DataTable

    Dim ClassEditPayroll As String = "", ClassEditNotes As String = ""


    Sub page_load()

        Dim StartDate As DateTime = Request.QueryString("startdate")
        Dim EndDate As DateTime = Request.QueryString("enddate")

        Dim sOfficeID As String = Request.QueryString("officeid")
        Dim OfficeID As Integer

        If IsPostBack = False Then
            DDL_Office.SelectedValue = sOfficeID
        End If

        If sOfficeID = "" Then
            OfficeID = 0
        Else
            OfficeID = CInt(sOfficeID)
        End If

        GetEmployeeData(StartDate, EndDate, OfficeID)
        sEmployeeList = GenerateListOfEmployees()
        GetJobData(StartDate, EndDate)
        GetAdjustmentsData(StartDate, EndDate)
        GetTimeEntryData(StartDate, EndDate)
        GetBonusData(StartDate, EndDate)

        LIT_Grid.Text = GenerateGrid(StartDate, EndDate, OfficeID)


    End Sub

    Sub GetEmployeeData(StartDate As DateTime, EndDate As DateTime, PrevPageOfficeID As Integer)

        Dim QueryCondition As String = ""

        Dim fType As String = DDL_Type.SelectedValue
        Dim OfficeID As String = DDL_Office.SelectedValue


        If Not fType = "" Then
            QueryCondition = QueryCondition & " and tEmployees.Type='" & fType & "' "
        End If

        If OfficeID = "" And PrevPageOfficeID <> 0 Then
            OfficeID = PrevPageOfficeID
        End If

        If Not OfficeID = "" Then
            QueryCondition = QueryCondition & " and OfficeID=" & OfficeID & " "
        End If

        Dim sTimeQuery As String = "select distinct FirstName, LastName, PayrollID, EmployeeID, isNull(PerDiem,0) as PerDiem from tEmployees, tShifts, tTimeEntry where tEmployees.EmployeeID = tShifts.UserPerformed and tShifts.ShiftID = tTimeEntry.ShiftID and CAST(tTimeEntry.StartTime as DATE) >= @StartDate and CAST(tTimeEntry.StartTime as DATE) <= @EndDate " & QueryCondition & " order by LastName asc"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndDate", EndDate))

        Try
            connTimeQuery.Open()
            Using daShifts As New SqlDataAdapter(cmdTimeQuery)
                daShifts.Fill(dtEmployees)
            End Using
        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

    End Sub

    Sub GetJobData(StartDate As DateTime, EndDate As DateTime)

        Dim sTimeQuery As String = "SELECT DISTINCT UserPerformed, fJobNumber, fPW, fState " &
    "FROM tShifts INNER JOIN tTimeEntry ON tShifts.ShiftID = tTimeEntry.ShiftID " &
    "WHERE CAST(tTimeEntry.StartTime AS DATE) >= @StartDate " &
    "AND CAST(tTimeEntry.StartTime AS DATE) <= @EndDate " &
    "AND Approved = 1"

        If Not String.IsNullOrWhiteSpace(sEmployeeList) Then
            sTimeQuery &= " AND UserPerformed IN (" & sEmployeeList & ")"
        End If
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndDate", EndDate))

        Try
            connTimeQuery.Open()
            Using daShifts As New SqlDataAdapter(cmdTimeQuery)
                daShifts.Fill(dtJobs)
            End Using
        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

    End Sub

    Sub GetAdjustmentsData(StartDate As DateTime, EndDate As DateTime)

        Dim sTimeQuery As String = "select DateAdjusted, JobNumber, EmployeeID, PW, State, isNull(Overtime,0) as Overtime, isNull(Standardtime,0) as StandardTime, isNull(Company,0) as Company, isNull(Personal,0) as Personal from tTimeAdjustments where CAST(tTimeAdjustments.DateAdjusted as DATE) >= @StartDate and CAST(tTimeAdjustments.DateAdjusted as DATE) <= @EndDate"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndDate", EndDate))

        Try
            connTimeQuery.Open()
            Using daShifts As New SqlDataAdapter(cmdTimeQuery)
                daShifts.Fill(dtAdjustments)
            End Using
        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

    End Sub

    Sub GetBonusData(StartDate As DateTime, EndDate As DateTime)

        Dim sTimeQuery As String = "select isNull(Bonus,0) as Bonus, isNull(Notes,'') as Notes, EmployeeID, JobNumber, State, StartDate, PW from tBonus where CAST(tBonus.StartDate as DATE) >= @StartDate and CAST(tBonus.StartDate as DATE) <= @EndDate"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndDate", EndDate))

        Try
            connTimeQuery.Open()
            Using daShifts As New SqlDataAdapter(cmdTimeQuery)
                daShifts.Fill(dtBonus)
            End Using
        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

    End Sub

    Sub GetTimeEntryData(StartDate As DateTime, EndDate As DateTime)

        Dim sTimeQuery As String = "select UserPerformed, fJobNumber, fPW, fState, StartTime, EndTime, fTravelTime, fPerDiem from tTimeEntry, tShifts where Approved = 1 and tTimeEntry.ShiftID = tShifts.ShiftID and CAST(tTimeEntry.StartTime as DATE) >= @StartDate and CAST(tTimeEntry.StartTime as DATE) <= @EndDate"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndDate", EndDate))

        Try
            connTimeQuery.Open()
            Using daShifts As New SqlDataAdapter(cmdTimeQuery)
                daShifts.Fill(dtTimeEntries)
            End Using
        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

    End Sub

    Function GenerateGrid(ByVal StartDate As Date, ByVal EndDate As Date, ByVal PrevPageOfficeID As Integer) As String

        Dim fType As String = DDL_Type.SelectedValue
        Dim OfficeID As String = DDL_Office.SelectedValue

        Dim strResult As New System.Text.StringBuilder()

        strResult.Append("<div id=""payrollholder""><h1 class=""payroll"">Payroll " & StartDate & " - " & EndDate & "</h1>")
        strResult.Append("<table class=""payroll"">")
        strResult.Append("<thead>" & GenerateHeader(StartDate, EndDate) & "</thead><tbody>")
        Dim FirstName As String = "", LastName As String = "", PayrollID As String = "", EmployeeID As Integer
        Dim fJobNumber As String = "", fPW As String, fState As String
        Dim DateRange As Integer = DateDiff(DateInterval.Day, StartDate, EndDate)

        If IsWeekLocked(StartDate, EndDate) = False Then
            ClassEditPayroll = "edit_payroll"
            ClassEditNotes = "edit_payroll_notes"
        End If

        Try

            For Each drEmployee As DataRow In dtEmployees.Rows

                FirstName = drEmployee.Item("FirstName")
                LastName = drEmployee.Item("LastName")
                PayrollID = drEmployee.Item("PayrollID")
                EmployeeID = drEmployee.Item("EmployeeID")
                multiplierPerDiem = drEmployee.Item("PerDiem")

                Dim filteredJobs() As DataRow = dtJobs.Select("UserPerformed=" & EmployeeID)

                For Each drJob As DataRow In filteredJobs

                    fJobNumber = drJob.Item("fJobNumber")
                    fPW = drJob.Item("fPW")
                    fState = drJob.Item("fState")

                    strResult.Append("<tr><td class=""name"">" & LastName & ", " & FirstName & "</td><td>" & PayrollID & "</td><td>" & fJobNumber.Replace("-", "") & "</td><td>" & fPW & "</td><td class=""divider"">" & fState.ToUpper & "</td>")

                    For x As Integer = 0 To DateRange

                        strResult.Append(CalculateDay(EmployeeID, DateAdd(DateInterval.Day, x, StartDate), fJobNumber, fPW, fState))

                    Next

                    Dim Bonus As Double = 0, Notes As String = ""

                    Dim filteredBonuses() As DataRow = dtBonus.Select("EmployeeID=" & EmployeeID & " and StartDate = '" & StartDate & "' and JobNumber='" & fJobNumber & "' and PW='" & fPW & "' and State='" & fState & "'")

                    For Each drBonus As DataRow In filteredBonuses
                        Bonus = drBonus.Item("Bonus")
                        Notes = drBonus.Item("Notes")
                    Next


                    RecTotCompanyAdjusted = RecTotCompany * multipilerCompany
                    recTotPersonalAdjusted = RecTotPersonal * multiplierPersonal
                    'RecTotPerDiemAdjusted = RecTotPerDiem * multiplierPerDiem
                    RecTotPerDiemAdjusted = RecTotPerDiemDays * multiplierPerDiemDaily

                    strResult.Append("<td>" & RecTotStandard & "</td>")
                    strResult.Append("<td>" & RecTotOvertime & "</td>")
                    strResult.Append("<td>" & RecTotCompany & "</td>")
                    strResult.Append("<td>$" & RecTotCompanyAdjusted.ToString("0.00") & "</td>")
                    strResult.Append("<td>" & RecTotPersonal & "</td>")
                    strResult.Append("<td>$" & recTotPersonalAdjusted.ToString("0.00") & "</td>")
                    strResult.Append("<td>" & RecTotPerDiem & "</td>")
                    strResult.Append("<td>$" & RecTotPerDiemAdjusted.ToString("0.00") & "</td>")
                    strResult.Append("<td class=""" & ClassEditPayroll & """ id=""bonus_" & EmployeeID & "_" & fJobNumber & "_" & StartDate & "_" & fPW & "_" & fState & """>" & Bonus.ToString("0.00") & "</td>")
                    strResult.Append("<td class=""" & ClassEditNotes & """ id=""notes_" & EmployeeID & "_" & fJobNumber & "_" & StartDate & "_" & fPW & "_" & fState & """>" & Notes & "</td>")

                    RecTotStandard = 0
                    RecTotOvertime = 0
                    RecTotCompany = 0
                    RecTotPersonal = 0
                    RecTotPerDiem = 0
                    RecTotPerDiemDays = 0

                    strResult.Append("</tr>" & vbCrLf)

                Next

            Next

        Catch ex As Exception
            strResult.Append("ERROR: " & ex.ToString)

        End Try

        Dim SpreadsheetURL As String = "payroll_spreadsheet.aspx?startdate=" & StartDate & "&enddate=" & EndDate & "&office=" & OfficeID & "&type=" & fType
        Dim Spreadsheetv2URL As String = "payroll_spreadsheetv2.aspx?startdate=" & StartDate & "&enddate=" & EndDate & "&office=" & OfficeID & "&type=" & fType
        Dim TimberlineURL As String = "generatecsv.aspx?start=" & StartDate & "&end=" & EndDate & "&office=" & OfficeID & "&type=" & fType

        strResult.Append("</tbody></table>")
        strResult.Append("<p class=""info""><strong>NOTES:</strong><br>")
        strResult.Append("- All fields showing daily totals (except Standard Time) can be edited by clicking on them. The Bonus and Notes fields are also editable. Totals are calculated automatically based on values in other fields.<br>")
        strResult.Append("- Please note that the totals will not reflect the changes until the page is refreshed. To see updated totals use your browser Refresh button, press F5 on your keyboard or use the Recalculate button below.</p>")
        strResult.Append("<input type=""button"" value=""Recalculate"" Onclick=""javascript:location.reload(true)"">")
        strResult.Append("<input type=""button"" value=""Generate Excel"" Onclick=""javascript:window.open('" & SpreadsheetURL & "','Spreadsheet'); "">")
        strResult.Append("<input type=""button"" value=""Generate Excel V2"" Onclick=""javascript:window.open('" & Spreadsheetv2URL & "','Spreadsheet'); "">")
        strResult.Append("<input type=""button"" value=""Timberline Export"" Onclick=""javascript:window.open('" & TimberlineURL & "','Timberline'); ""></div>")

        Return strResult.ToString()

    End Function

    Function CalculateDay(ByVal EmployeeID As Integer, ByVal DateToRun As Date, ByVal JobNumber As String, ByVal PW As String, ByVal state As String) As String

        Dim ReturnString As String = ""
        Dim StartTime As DateTime, Endtime As DateTime, Travel As String, PerDiem As String = ""
        Dim TotStandardTime As Double, TotOverTime As Double, TotCompany As Double, TotPersonal As Double, Timedifference As Double, TotPerDiem As Double, PerDiemDays As Integer = 0
        Dim adjStandardTime As Double = 0, adjOvertime As Double = 0, adjCompany As Double = 0, adjPersonal As Double = 0

        Try

            Dim filteredAdjustments() As DataRow = dtAdjustments.Select("EmployeeID=" & EmployeeID & " and DateAdjusted = '" & DateToRun.ToString("MM-dd-yyyy") & "' and JobNumber='" & JobNumber & "' and PW='" & PW & "' and State='" & state & "'")

            For Each drAdjustment As DataRow In filteredAdjustments
                adjCompany = drAdjustment.Item("Company")
                adjPersonal = drAdjustment.Item("Personal")
                adjStandardTime = drAdjustment.Item("StandardTime")
                adjOvertime = drAdjustment.Item("Overtime")
            Next

            '----- LOOP THROUGH TIME ENTRIES FOR THAT DAY+PROJECT+WORKER(+State+PW)
            Dim filteredTimeEntries() As DataRow = dtTimeEntries.Select("UserPerformed=" & EmployeeID & " and StartTime >= '" & DateToRun.ToString("MM-dd-yyyy") & "' and StartTime < '" & DateAdd(DateInterval.Day, 1, DateToRun).ToString("MM-dd-yyyy") & "' and fJobNumber='" & JobNumber & "' and fPW='" & PW & "' and fState='" & state & "'")

            For Each drDay As DataRow In filteredTimeEntries


                StartTime = drDay.Item("StartTime")
                Endtime = drDay.Item("EndTime")
                'Overtime = drDay.Item("Overtime")
                Travel = drDay.Item("fTravelTime")
                PerDiem = drDay.Item("fPerDiem")

                Timedifference = DateDiff(DateInterval.Minute, StartTime, Endtime) / 60

                '---- ADD UP TIME ENTRIES FOR THAT DAY
                If Travel = "N/A" Or Travel = "" Then
                    TotStandardTime = TotStandardTime + Timedifference
                    'TotOverTime = TotOverTime + Overtime
                ElseIf Travel = "Company" Then
                    TotCompany = TotCompany + Timedifference
                ElseIf Travel = "Personal" Then
                    TotPersonal = TotPersonal + Timedifference
                End If

                If PerDiem = "Yes" Then
                    TotPerDiem = TotPerDiem + Timedifference
                    PerDiemDays = PerDiemDays + 1
                End If

            Next

            '--- ADJUST TIME
            If adjStandardTime > 0 Then
                TotStandardTime = adjStandardTime
            End If

            If adjOvertime > 0 Then
                TotOverTime = adjOvertime
                TotStandardTime = TotStandardTime - adjOvertime
            End If

            If adjCompany > 0 Then
                TotCompany = adjCompany
            End If

            If adjPersonal > 0 Then
                TotPersonal = adjPersonal
            End If



            ReturnString = "<td>" & IsZero(TotStandardTime) & "</td>"
            ReturnString = ReturnString & "<td class=""" & ClassEditPayroll & """ id=""ot_" & EmployeeID & "_" & JobNumber & "_" & DateToRun & "_" & PW & "_" & state & """>" & IsZero(TotOverTime) & "</td>"
            ReturnString = ReturnString & "<td class=""" & ClassEditPayroll & """ id=""c_" & EmployeeID & "_" & JobNumber & "_" & DateToRun & "_" & PW & "_" & state & """>" & IsZero(TotCompany) & "</td>"
            ReturnString = ReturnString & "<td class=""divider " & ClassEditPayroll & """ id=""p_" & EmployeeID & "_" & JobNumber & "_" & DateToRun & "_" & PW & "_" & state & """>" & IsZero(TotPersonal) & "</td>"


        Catch ex As Exception
            ReturnString = "ERROR: " & ex.ToString & "<br>"
        End Try

        RecTotStandard = RecTotStandard + TotStandardTime
        RecTotOvertime = RecTotOvertime + TotOverTime
        RecTotCompany = RecTotCompany + TotCompany
        RecTotPersonal = RecTotPersonal + TotPersonal
        RecTotPerDiem = RecTotPerDiem + TotPerDiem
        RecTotPerDiemDays = RecTotPerDiemDays + PerDiemDays


        Return ReturnString


    End Function

    Function GenerateHeader(ByVal StartDate As Date, ByVal EndDate As Date) As String

        Dim DateRange As Integer = DateDiff(DateInterval.Day, StartDate, EndDate)

        Dim HeaderLine1 As String = "<tr><th rowspan=""3"">Employee&nbsp;Name</th><th rowspan=""3"">Payroll&nbsp;#</th><th rowspan=""3"">Job&nbsp;#</th><th rowspan=""3"">PW</th><th rowspan=""3"">State</th>"
        Dim HeaderLine2 As String = "<tr>"
        Dim HeaderLine3 As String = "<tr>"

        For x As Integer = 0 To DateRange

            HeaderLine1 = HeaderLine1 & "<th colspan=""4"">" & DateAdd(DateInterval.Day, x, StartDate).DayOfWeek.ToString
            HeaderLine1 = HeaderLine1 & "<br>" & DateAdd(DateInterval.Day, x, StartDate) & "</th>"
            HeaderLine2 = HeaderLine2 & "<th colspan=""2"">WORK</th><th colspan=""2"">TRAVEL</th>"
            HeaderLine3 = HeaderLine3 & "<th>ST</th><th>OT</th><th>C</th><th>P</th>"

        Next

        HeaderLine1 = HeaderLine1 & "<th colspan=""2"" rowspan=""2"">TOTALS</th><th colspan=""4"">VEHICLE</th><th colspan=2>PER DIEM</th><th rowspan=3>Bonus</th><th rowspan=3 style=""min-width:100px;"">Notes</th></tr>"
        HeaderLine2 = HeaderLine2 & "<th>4</th><th>Code 18</th><th>7</th><th>Code 19</th><th rowspan=2>PD Hrs</th><th>Code<br>11</th></tr>"
        HeaderLine3 = HeaderLine3 & "<th>ST</th><th>OT</th><th>COM</th><th>$</th><th>PER</th><th>$</th><th>PD$</th></tr>"

        Return HeaderLine1 & HeaderLine2 & HeaderLine3

    End Function

    Function IsZero(ByVal InputValue As Double) As String

        If InputValue = 0 Then
            Return " "
        Else
            Return InputValue.ToString
        End If

    End Function


    Function GenerateListOfEmployees() As String

        Dim PMList As String = ""

        For Each drEmployee As DataRow In dtEmployees.Rows
            If PMList = "" Then
                PMList = drEmployee.Item("EmployeeID")
            Else
                PMList = PMList & "," & UCase(drEmployee.Item("EmployeeID"))
            End If
        Next

        Return PMList

    End Function

End Class
