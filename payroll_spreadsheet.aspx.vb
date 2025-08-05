Imports System.Data.SqlClient
Imports System.IO
Imports System.Data

Partial Class payroll_spreadsheet
    Inherits GlobalClass

    'Public sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("FullConnection").ConnectionString

    Dim RecTotStandard As Double = 0, RecTotOvertime As Double = 0, RecTotCompany As Double = 0, RecTotPersonal As Double = 0, RecTotPerDiem As Double = 0, RecTotPerDiemDays As Integer = 0
    Dim RecTotCompanyAdjusted As Double, recTotPersonalAdjusted As Double, RecTotPerDiemAdjusted As Double


    Dim multipilerCompany As Double = 4.0
    Dim multiplierPersonal As Double = 7.0
    Dim multiplierPerDiemDaily As Double = 44
    Dim multiplierPerDiem As Double

    Public sEmployeeList As String = ""
    Public dtEmployees As New DataTable
    Public dtJobs As New DataTable
    Public dtAdjustments As New DataTable
    Public dtTimeEntries As New DataTable
    Public dtBonus As New DataTable


    Sub page_load()

        Dim StartDate As DateTime = Request.QueryString("startdate")
        Dim EndDate As DateTime = Request.QueryString("enddate")
        Dim QSType As String = Request.QueryString("type")
        Dim QSOffice As String = Request.QueryString("office")

        If QSOffice Is Nothing Or QSOffice = "" Then QSOffice = 0

        Response.ContentType = "application/vnd.ms-excel"

        GetEmployeeData(StartDate, EndDate, QSOffice)
        sEmployeeList = GenerateListOfEmployees()
        GetJobData(StartDate, EndDate)
        GetAdjustmentsData(StartDate, EndDate)
        GetTimeEntryData(StartDate, EndDate)
        GetBonusData(StartDate, EndDate)

        LIT_Grid.Text = GenerateGrid(StartDate, EndDate, QSType, QSOffice)

    End Sub

    Sub GetEmployeeData(StartDate As DateTime, EndDate As DateTime, PrevPageOfficeID As Integer)

        Dim QueryCondition As String = ""

        Dim fType As String = Request.QueryString("type")
        Dim OfficeID As String = Request.QueryString("office")


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

        Dim sTimeQuery As String = "select distinct UserPerformed, fJobNumber, fPW, fState from tShifts, tTimeEntry where tShifts.ShiftID = tTimeEntry.ShiftID and UserPerformed in (" & sEmployeeList & ") and CAST(tTimeEntry.StartTime as DATE) >= @StartDate and CAST(tTimeEntry.StartTime as DATE) <= @EndDate and Approved = 1"
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

    Function GenerateGrid(ByVal StartDate As Date, ByVal EndDate As Date, ByVal fType As String, ByVal OfficeID As String) As String

        Dim strResult As New System.Text.StringBuilder()

        strResult.Append("<h1 class=""payroll"">Payroll " & StartDate & " - " & EndDate & "</h1>")
        strResult.Append("<table class=""payroll"" border=""1"">")
        strResult.Append("<thead>" & GenerateHeader(StartDate, EndDate) & "</thead><tbody>")
        Dim FirstName As String = "", LastName As String = "", PayrollID As String = "", EmployeeID As Integer
        Dim fJobNumber As String = "", fPW As String, fState As String
        Dim DateRange As Integer = DateDiff(DateInterval.Day, StartDate, EndDate)
        Dim RowColor As String = ""

        Dim QueryCondition As String = ""
        Try
            For Each drEmployee As DataRow In dtEmployees.Rows

                FirstName = drEmployee.Item("FirstName")
                LastName = drEmployee.Item("LastName")
                PayrollID = drEmployee.Item("PayrollID")
                EmployeeID = drEmployee.Item("EmployeeID")
                multiplierPerDiem = drEmployee.Item("PerDiem")

                If PayrollID = "99027" Then
                    Dim s As String = ""
                End If
                Dim filteredJobs() As DataRow = dtJobs.Select("UserPerformed=" & EmployeeID)

                For Each drJob As DataRow In filteredJobs

                    fJobNumber = drJob.Item("fJobNumber")
                    fPW = drJob.Item("fPW")
                    fState = drJob.Item("fState")

                    Dim arrJobNumber As String() = fJobNumber.Split("-")
                    Dim jobP1 As String = "", jobP2 As String = "", jobP3 As String = "", jobP4 As String = ""

                    If arrJobNumber.Length > 0 Then
                        If Not arrJobNumber(0) = Nothing Then jobP1 = arrJobNumber(0)
                    End If

                    If arrJobNumber.Length > 1 Then
                        If Not arrJobNumber(1) = Nothing Then jobP2 = arrJobNumber(1)
                    End If

                    If arrJobNumber.Length > 2 Then
                        If Not arrJobNumber(2) = Nothing Then jobP3 = arrJobNumber(2)
                    End If

                    If arrJobNumber.Length > 3 Then
                        If Not arrJobNumber(3) = Nothing Then jobP4 = arrJobNumber(3)
                    End If

                    If fState = "PTO" Then
                        RowColor = " style=""background-color:#C6E0B4"""
                    Else
                        RowColor = ""
                    End If

                    strResult.Append("<tr" & RowColor & ">")
                    strResult.Append("<td Class=""name"">" & LastName & ", " & FirstName & "</td>")
                    strResult.Append("<td>" & PayrollID & "</td>")
                    'strResult.Append("<td>" & fJobNumber.Replace("-", "") & "&nbsp;</td>")
                    strResult.Append("<td>" & jobP1 & "&nbsp;</td>")
                    strResult.Append("<td>" & jobP2 & "&nbsp;</td>")
                    strResult.Append("<td>" & jobP3 & "&nbsp;</td>")
                    strResult.Append("<td>" & jobP4 & "&nbsp;</td>")
                    strResult.Append("<td>" & fPW & "</td>")
                    strResult.Append("<td Class= ""divider"">" & fState.ToUpper & "</td>")

                    For x As Integer = 0 To DateRange

                        strResult.Append(CalculateDay(EmployeeID, DateAdd(DateInterval.Day, x, StartDate), fJobNumber, fPW, fState))

                    Next

                    Dim Bonus As Double = 0, Notes As String = ""
                    Dim filteredBonuses() As DataRow = dtBonus.Select("EmployeeID=" & EmployeeID & " And StartDate = '" & StartDate & "' and JobNumber='" & fJobNumber & "' and PW='" & fPW & "' and State='" & fState & "'")

                    For Each drBonus As DataRow In filteredBonuses
                        Bonus = drBonus.Item("Bonus")
                        Notes = drBonus.Item("Notes")
                    Next

                    RecTotCompanyAdjusted = RecTotCompany * multipilerCompany
                    recTotPersonalAdjusted = RecTotPersonal * multiplierPersonal
                    'RecTotPerDiemAdjusted = RecTotPerDiem * multiplierPerDiem
                    RecTotPerDiemAdjusted = RecTotPerDiemDays * multiplierPerDiemDaily

                    strResult.Append("<td style=""background-color:#8EA9DB"">" & RecTotStandard & "</td>")
                    strResult.Append("<td style=""background-color:#8EA9DB"">" & RecTotOvertime & "</td>")
                    strResult.Append("<td>" & RecTotCompany & "</td>")
                    strResult.Append("<td>$" & RecTotCompanyAdjusted.ToString("0.00") & "</td>")
                    strResult.Append("<td>" & RecTotPersonal & "</td>")
                    strResult.Append("<td>$" & recTotPersonalAdjusted.ToString("0.00") & "</td>")
                    strResult.Append("<td>" & RecTotPerDiem & "</td>")
                    strResult.Append("<td>$" & RecTotPerDiemAdjusted.ToString("0.00") & "</td>")
                    strResult.Append("<td class=""edit_payroll"" id=""bonus_" & EmployeeID & "_" & fJobNumber & "_" & StartDate & "_" & fPW & "_" & fState & """>" & Bonus.ToString("0.00") & "</td>")
                    strResult.Append("<td class=""edit_payroll_notes"" id=""notes_" & EmployeeID & "_" & fJobNumber & "_" & StartDate & "_" & fPW & "_" & fState & """>" & Notes & "</td>")

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
            LogError("payroll_spreadsheet.aspx :: GenerateGrid", ex.ToString)
            strResult.Append("ERROR: " & ex.ToString)
        End Try

        strResult.Append("</tbody></table>")

        LogSpreadsheet(strResult.ToString())

        Return strResult.ToString()

    End Function

    Function CalculateDay(ByVal EmployeeID As Integer, ByVal DateToRun As Date, ByVal JobNumber As String, ByVal PW As String, ByVal state As String) As String

        Dim ReturnString As String = ""
        Dim DateParsed As String = DateToRun.ToString("MM-dd-yyyy")
        Dim StartTime As DateTime, Endtime As DateTime, Travel As String, PerDiem As String = ""
        Dim TotStandardTime As Double, TotOverTime As Double, TotCompany As Double, TotPersonal As Double, Timedifference As Double, TotPerDiem As Double, PerDiemDays As Integer = 0
        Dim adjStandardTime As Double = 0, adjOvertime As Double = 0, adjCompany As Double = 0, adjPersonal As Double = 0

        '---- GET ADJUSTMENTS
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
                    PerDiemDays = 1
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
            ReturnString = ReturnString & "<td style=""background-color:#BDD7EE"" class=""edit_payroll"" id=""ot_" & EmployeeID & "_" & JobNumber & "_" & DateToRun & "_" & PW & "_" & state & """>" & IsZero(TotOverTime) & "</td>"
            ReturnString = ReturnString & "<td class=""edit_payroll"" id=""c_" & EmployeeID & "_" & JobNumber & "_" & DateToRun & "_" & PW & "_" & state & """>" & IsZero(TotCompany) & "</td>"
            ReturnString = ReturnString & "<td class=""divider edit_payroll"" id=""p_" & EmployeeID & "_" & JobNumber & "_" & DateToRun & "_" & PW & "_" & state & """>" & IsZero(TotPersonal) & "</td>"

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

        Dim HeaderLine1 As String = "<tr><th rowspan=""3"" style=""background-color:#333333;color:#fff;"">Employee&nbsp;Name</th><th rowspan=""3"" style=""background-color:#333333;color:#fff;"">Payroll&nbsp;#</th><th rowspan=""3"" colspan=""4"" style=""background-color:#333333;color:#fff;"">Job&nbsp;#</th><th rowspan=""3"" style=""background-color:#333333;color:#fff;"">PW</th><th rowspan=""3"" style=""background-color:#333333;color:#fff;"">State</th>"
        Dim HeaderLine2 As String = "<tr>"
        Dim HeaderLine3 As String = "<tr>"

        For x As Integer = 0 To DateRange

            HeaderLine1 = HeaderLine1 & "<th colspan=""4"" style=""background-color:#333333;color:#fff;"">" & DateAdd(DateInterval.Day, x, StartDate).DayOfWeek.ToString
            HeaderLine1 = HeaderLine1 & "<br>" & DateAdd(DateInterval.Day, x, StartDate) & "</th>"
            HeaderLine2 = HeaderLine2 & "<th colspan=""2"" style=""background-color:#333333;color:#fff;"">WORK</th><th colspan=""2"" style=""background-color:#333333;color:#fff;"">TRAVEL</th>"
            HeaderLine3 = HeaderLine3 & "<th style=""background-color:#333333;color:#fff;"">ST</th><th style=""background-color:#333333;color:#fff;"">OT</th><th style=""background-color:#333333;color:#fff;"">C</th><th style=""background-color:#333333;color:#fff;"">P</th>"

        Next

        HeaderLine1 = HeaderLine1 & "<th colspan=""2"" rowspan=""2"" style=""background-color:#333333;color:#fff;"">TOTALS</th><th colspan=""4"" style=""background-color:#333333;color:#fff;"">VEHICLE</th><th colspan=2 style=""background-color:#333333;color:#fff;"">PER DIEM</th><th rowspan=3 style=""background-color:#333333;color:#fff;"">Bonus</th><th rowspan=3 style=""min-width:100px; background-color:#333333;color:#fff;"" >Notes</th></tr>"
        HeaderLine2 = HeaderLine2 & "<th style=""background-color:#333333;color:#fff;"">4</th><th style=""background-color:#333333;color:#fff;"">Code 18</th><th style=""background-color:#333333;color:#fff;"">7</th><th style=""background-color:#333333;color:#fff;"">Code 19</th><th rowspan=2 style=""background-color:#333333;color:#fff;"">PD Hrs</th><th style=""background-color:#333333;color:#fff;"">Code<br>11</th></tr>"
        HeaderLine3 = HeaderLine3 & "<th style=""background-color:#333333;color:#fff;"">ST</th><th style=""background-color:#333333;color:#fff;"">OT</th><th style=""background-color:#333333;color:#fff;"">COM</th><th style=""background-color:#333333;color:#fff;"">$</th><th style=""background-color:#333333;color:#fff;"">PER</th><th style=""background-color:#333333;color:#fff;"">$</th><th style=""background-color:#333333;color:#fff;"">PD$</th></tr>"

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

    Sub LogSpreadsheet(ByVal ReturnString As String)

        Dim fp As StreamWriter
        Dim TimeStamp As String = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
        Dim Filename As String = TimeStamp & "_" & Session("UserID") & ".xls"

        Try
            fp = File.CreateText(Server.MapPath(".\SpreadsheetLogs\") & Filename)
            fp.WriteLine(ReturnString)
            fp.Close()
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("payroll_spreadsheet.aspx :: LogSpreadsheet", ex.ToString)
        Finally

        End Try

    End Sub


End Class
