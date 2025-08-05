Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Partial Class payroll_spreadsheetv2
    Inherits GlobalClass

    'Public sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("FullConnection").ConnectionString

    Dim RecTotStandard As Double = 0, RecTotOvertime As Double = 0, RecTotCompany As Double = 0, RecTotPersonal As Double = 0, RecTotPerDiem As Double = 0, RecTotPerDiemDays As Integer = 0
    Dim RecTotCompanyAdjusted As Double, recTotPersonalAdjusted As Double, RecTotPerDiemAdjusted As Double, PerDiemTot As Boolean = False

    Dim PDHrs As String = ""

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

        'Response.ContentType = "application/vnd.ms-excel"

        GetEmployeeData(StartDate, EndDate, QSOffice)
        sEmployeeList = GenerateListOfEmployees()
        GetJobData(StartDate, EndDate)
        GetAdjustmentsData(StartDate, EndDate)
        GetTimeEntryData(StartDate, EndDate)
        GetBonusData(StartDate, EndDate)

        LIT_Grid.Text = GenerateGrid(StartDate, EndDate, QSType, QSOffice)

        ExportDataToCSV(LIT_Grid.Text)
    End Sub



    Public Sub ExportDataToCSV(ByVal data As String)
        Dim dataTable As New DataTable()

        HttpContext.Current.Response.Clear()

        HttpContext.Current.Response.ContentType = "text/csv"
        HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment; filename=Data.xls")
        HttpContext.Current.Response.ContentEncoding = Encoding.UTF8
        HttpContext.Current.Response.Write(data.ToString())
        HttpContext.Current.Response.End()
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

        Dim sTimeQuery As String = "select distinct FirstName, LastName, PayrollID, EmployeeID, isNull(PerDiem,0) as PerDiem , Type from tEmployees, tShifts, tTimeEntry where tEmployees.EmployeeID = tShifts.UserPerformed and tShifts.ShiftID = tTimeEntry.ShiftID and CAST(tTimeEntry.StartTime as DATE) >= @StartDate and CAST(tTimeEntry.StartTime as DATE) <= @EndDate " & QueryCondition & " order by LastName asc"
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

        Dim sTimeQuery As String = "select distinct UserPerformed, fJobNumber, fShiftType ,fPW, fState from tShifts, tTimeEntry where tShifts.ShiftID = tTimeEntry.ShiftID and UserPerformed in (" & sEmployeeList & ") and CAST(tTimeEntry.StartTime as DATE) >= @StartDate and CAST(tTimeEntry.StartTime as DATE) <= @EndDate and Approved = 1"
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
    Function FormatAsString(value As Object) As String
        Return "=""" & value.ToString() & """"
    End Function


    Function GenerateGrid(ByVal StartDate As Date, ByVal EndDate As Date, ByVal fType As String, ByVal OfficeID As String) As String

        Dim strResult As New System.Text.StringBuilder()

        strResult.Append("<h1 class=""payroll"">Payroll " & StartDate & " - " & EndDate & "</h1>")
        strResult.Append("<table class=""payroll"" border=""1"">")
        strResult.Append("<thead>" & GenerateHeader(StartDate, EndDate) & "</thead><tbody>")

        Dim FirstName As String = "", LastName As String = "", PayrollID As String = "", EmployeeID As Integer
        Dim fJobNumber As String = "", fPW As String, fState As String, fShiftType As String, Type As String = ""
        Dim DateRange As Integer = DateDiff(DateInterval.Day, StartDate, EndDate)
        Dim RowColor As String = ""

        Dim QueryCondition As String = ""
        Dim duplicateRow As Boolean = False


        ' This will track the current employee to alternate row colors when employee changes
        Dim currentEmployeeID As Integer = -1
        Dim rowColorToggle As Boolean = False ' This will control row color alternation
        Dim emplyeeAdded As Boolean = False

        Try
            Dim resPerDiem As String = GetPerDiem("PerDiem")
            Dim PerDiemValue As Double
            If Double.TryParse(resPerDiem, PerDiemValue) Then

            End If

            Dim odd As Boolean = False
            For Each drEmployee As DataRow In dtEmployees.Rows

                FirstName = drEmployee.Item("FirstName")
                LastName = drEmployee.Item("LastName")
                PayrollID = drEmployee.Item("PayrollID")
                EmployeeID = drEmployee.Item("EmployeeID")
                Type = drEmployee.Item("Type")


                If PayrollID = "98563" Then
                    Dim s As String = ""
                End If

                'If FirstName = "Robin" Then
                '    Dim ssss As String = ""
                'End If

                Dim multiplierPerDiem = drEmployee.Item("PerDiem")
                Dim filteredJobs() As DataRow = dtJobs.Select("UserPerformed=" & EmployeeID)

                For Each drJob As DataRow In filteredJobs


                    fJobNumber = drJob.Item("fJobNumber")
                    fPW = drJob.Item("fPW")


                    fShiftType = drJob.Item("fShiftType")
                    fState = drJob.Item("fState")

                    Dim arrJobNumber As String() = fJobNumber.Split("-")
                    Dim jobP1 As String = "", jobP2 As String = "", jobP3 As String = "", jobP4 As String = "", fullcode As String = ""

                    If Len(Trim(fShiftType)) > 0 Then
                        Select Case fShiftType
                            Case "PTO", "Vacation"
                                jobP1 = "V"
                            Case "Holiday"
                                jobP1 = "H"
                        End Select
                    End If

                    If arrJobNumber.Length > 1 Then
                        jobP2 = arrJobNumber(1)
                    End If

                    If arrJobNumber.Length > 2 Then
                        If Not arrJobNumber(2) = Nothing Then jobP3 = arrJobNumber(2)
                    End If

                    If arrJobNumber.Length > 3 Then
                        If Not arrJobNumber(3) = Nothing Then
                            jobP4 = arrJobNumber(3)
                        End If
                    End If

                    Dim trimmedString As String = jobP3.TrimStart("0"c) 'job#
                    Dim digitCount As Integer = trimmedString.Length
                    Dim taskId As String = ""
                    If digitCount = 5 Then
                        '' fullcode = trimmedString + "-" + jobP2
                        If jobP4.Length = 0 Then
                            fullcode = trimmedString
                        Else
                            fullcode = trimmedString + "-" + jobP4
                        End If
                        taskId = "10-9500"
                    ElseIf digitCount = 4 Then
                        If jobP2 = "12" Or jobP2 = "14" Then
                            fullcode = "70-" + trimmedString
                        ElseIf jobP2 = "15" Then
                            fullcode = "50-" + trimmedString
                        Else
                            Dim p2 As Integer
                            Try
                                p2 = Integer.Parse(jobP2)
                            Catch ex As Exception
                                p2 = 0
                            End Try
                            If p2 = 0 Then
                                fullcode = trimmedString
                            ElseIf p2 > 9 Then
                                fullcode = jobP2 + "-" + trimmedString
                            Else
                                fullcode = p2.ToString + "0" + "-" + trimmedString
                            End If
                        End If
                    End If

                    Dim Bonus As Double = 0, Notes As String = ""
                    Dim filteredBonuses() As DataRow = dtBonus.Select("EmployeeID=" & EmployeeID & " And StartDate = '" & StartDate & "' and JobNumber='" & fJobNumber & "' and PW='" & fPW & "' and State='" & fState & "'")
                    For Each drBonus As DataRow In filteredBonuses
                        Bonus = drBonus.Item("Bonus")
                        Notes = drBonus.Item("Notes")
                    Next

                    Dim empStr As New List(Of EmployeeRow)()

                    For x As Integer = 0 To DateRange
                        Dim dayCalculations As String = CalculateDay(EmployeeID, Microsoft.VisualBasic.DateAndTime.DateAdd(DateInterval.Day, x, StartDate), fJobNumber, fPW, fState)
                        If Not String.IsNullOrEmpty(dayCalculations) Then
                            Dim overtime As Double = GetOvertimer(dayCalculations)


                            If fState = "CA" Then
                                If IsDoubleTime(dayCalculations) Then
                                    jobP1 = "3"
                                End If
                            End If

                            If overtime > 0 Then
                                duplicateRow = True
                                If (fState = "NY NIGHT DIFF" Or fState = "NY NIGT DIFF") Then
                                    jobP1 = "S2"
                                Else
                                    jobP1 = "2"
                                End If
                            Else
                                duplicateRow = False
                                If (fState = "NY NIGHT DIFF" Or fState = "NY NIGT DIFF") Then
                                    jobP1 = "S1"
                                Else
                                    jobP1 = "1"
                                End If
                            End If
                            If fState = "PTO" Or fState = "13P" Or fState = "7" Or fState = "14V" Then

                                jobP1 = "V"

                            End If
                            If fState = "15" Or fState = "HOL" Then

                                jobP1 = "H"

                            End If



                            Dim empRec As New StringBuilder()



                            If EmployeeID <> currentEmployeeID Then

                                ' If a new employee, toggle the row color

                                currentEmployeeID = EmployeeID ' Update the current employee
                                rowColorToggle = Not rowColorToggle



                            End If

                            ' Determine the row color based on the toggle state
                            If rowColorToggle Then
                                RowColor = " style=""background-color:#C6E0B4""" ' Greenish background
                            Else
                                RowColor = "" ' No background color
                            End If



                            empRec.Append("<tr" & RowColor & ">")

                            empRec.Append("<td >" & GetCompanycode(Type) & "</td>")

                            empRec.Append("<td style='mso-number-format:\@;'>" & taskId & "</td>")
                            empRec.Append("<td>" & LastName & ", " & FirstName & "</td>")
                            empRec.Append("<td>" & PayrollID & "</td>")
                            empRec.Append("<td>" & jobP1 & " </td>") 'Paycode column G
                            empRec.Append("<td>" & jobP2 & " </td>")
                            empRec.Append("<td>" & jobP3 & " </td>") 'Job# column H
                            empRec.Append("<td>" & FormatAsString(String.Format("{0:D3}", jobP4)) & " </td>")
                            empRec.Append("<td style='mso-number-format:\@;'>" & fullcode & " </td>")
                            If fPW = "" Or fPW = "N" Then
                                empRec.Append("<td>" & "No" & "</td>")
                            Else fPW = "Y"
                                empRec.Append("<td>" & "Yes" & "</td>")
                            End If
                            empRec.Append("<td>" & fState.ToUpper & "</td>")
                            empRec.Append(dayCalculations)

                            empStr.Add(New EmployeeRow(empRec, duplicateRow))


                        End If
                    Next

                    RecTotCompanyAdjusted = RecTotCompany * multipilerCompany
                    recTotPersonalAdjusted = RecTotPersonal * multiplierPersonal
                    RecTotPerDiemAdjusted = RecTotPerDiemDays * multiplierPerDiemDaily
                    If PerDiemTot Then
                        RecTotPerDiemAdjusted = PerDiemValue
                    Else
                        RecTotPerDiemAdjusted = 0
                    End If

                    For Each empRow As EmployeeRow In empStr
                        Dim empRec As StringBuilder = empRow.EmployeeRecord
                        Dim isDuplicateRow As Boolean = empRow.DuplicateRow

                        empRec.Append("<td>" & RecTotCompany & "</td>")
                        empRec.Append("<td>$" & RecTotCompanyAdjusted.ToString("0.00") & "</td>")
                        empRec.Append("<td>" & RecTotPersonal & "</td>")
                        empRec.Append("<td>$" & recTotPersonalAdjusted.ToString("0.00") & "</td>")
                        empRec.Append("<td>" & PDHrs & "</td>")
                        empRec.Append("<td>$" & RecTotPerDiemAdjusted.ToString("0.00") & "</td>")
                        empRec.Append("<td>" & Bonus.ToString("0.00") & "</td>")
                        empRec.Append("<td>" & Notes & "</td>")
                        empRec.Append("</tr>" & vbCrLf)

                        If isDuplicateRow Then
                            Dim tmpRow As String = empRec.ToString()
                            strResult.Append(ReplaceduplicateNormalHours(empRec.ToString(), 14, 5, "1")) 'overtime'
                            strResult.Append(ReplaceduplicateOvertime(empRec.ToString(), 13, 14)) 'normal hours'
                            'strResult.Append(empRec.ToString())
                        Else
                            strResult.Append(empRec.ToString())
                        End If
                    Next

                    RecTotStandard = 0
                    RecTotOvertime = 0
                    RecTotCompany = 0
                    RecTotPersonal = 0
                    RecTotPerDiem = 0
                    RecTotPerDiemDays = 0
                    duplicateRow = False
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

    Function GetOvertimer(html As String) As Double
        Dim pattern As String = "<td>(.*?)</td>"

        ' Match all <td> elements
        Dim matches As MatchCollection = Regex.Matches(html, pattern)

        ' Check if there are enough matches to get the 3rd <td> element
        If matches.Count >= 3 Then
            ' Get the value from the 3rd <td> element
            Dim value As String = matches(2).Groups(1).Value

            ' Try to parse the value as a double
            Dim overtimeValue As Double
            If Double.TryParse(value, overtimeValue) Then
                Return overtimeValue
            End If
        End If

        ' Return 0 if not found or if the value is not a valid number
        Return 0
    End Function
    Function IsDoubleTime(html As String) As Boolean
        Dim pattern As String = "<td>(.*?)</td>"

        ' Match all <td> elements
        Dim matches As MatchCollection = Regex.Matches(html, pattern)

        ' Check if there are enough matches to get the 3rd <td> element
        If matches.Count >= 3 Then
            ' Get the value from the 3rd <td> element
            Dim valueDate As String = matches(0).Groups(1).Value
            Dim valueWorkingHour As String = matches(1).Groups(1).Value
            Dim valueOvertime As String = matches(2).Groups(1).Value

            Dim hoursWorked As Double
            Double.TryParse(valueWorkingHour, hoursWorked)

            ' Try to parse the value as a double
            Dim overtimeValue As Double
            Double.TryParse(valueOvertime, overtimeValue)


            Dim workDate As DateTime
            Try
                workDate = DateTime.Parse(valueDate)
                Dim dayOfWeek As DayOfWeek = workDate.DayOfWeek

                Select Case dayOfWeek
                    Case DayOfWeek.Monday, DayOfWeek.Tuesday, DayOfWeek.Wednesday, DayOfWeek.Thursday, DayOfWeek.Friday
                        If hoursWorked > 12 Then
                            Return True
                        End If

                    Case DayOfWeek.Saturday
                        If hoursWorked > 12 Or overtimeValue > 12 Then
                            Return True
                        End If

                    Case DayOfWeek.Sunday
                        If hoursWorked > 8 Or overtimeValue > 8 Then
                            Return True
                        End If

                End Select

            Catch ex As FormatException
                MsgBox("Invalid date format: " & ex.Message)
            End Try

            ' Try to parse the value as a double
            Return False

        End If


        Return False

    End Function

    Function ReplaceduplicateOvertime(html As String, normalHourIndex As Integer, overTimeHourIndex As Integer) As String
        Dim pattern As String = "(<td.*?>.*?</td>)"
        Dim matches As MatchCollection = Regex.Matches(html, pattern)

        ' Extract the part of the string before the <td> elements
        Dim startTag As String = html.Substring(0, html.IndexOf("<td"))

        ' Check if there are enough matches to access the specified column
        If matches.Count >= normalHourIndex Then
            ' Rebuild the HTML string with the replacement
            Dim modifiedHtml As New Text.StringBuilder()
            modifiedHtml.Append(startTag) ' Append the original start tag (e.g., <tr style="background-color:#C6E0B4">)

            For i As Integer = 0 To matches.Count - 1
                Dim nomralHourvalue As String = ""
                Dim overTimeValue As String = ""
                If i + 1 = normalHourIndex Then
                    modifiedHtml.Append(matches(i + 1).Value)
                ElseIf i + 1 = overTimeHourIndex Then
                    modifiedHtml.Append("<td></td>")
                ElseIf i + 1 = 21 Then  'PD Hrs
                    modifiedHtml.Append("<td>No</td>")
                ElseIf i + 1 = 22 Then  'Code 11
                    modifiedHtml.Append("<td>$0.00</td>")
                Else
                    modifiedHtml.Append(matches(i).Value)
                End If
            Next

            ' Append the rest of the HTML after the last <td> element
            Dim endTagIndex As Integer = html.IndexOf("</tr>")
            If endTagIndex > -1 Then
                modifiedHtml.Append(html.Substring(endTagIndex))
            End If

            Return modifiedHtml.ToString()
        End If

        ' Return the original HTML if no modification was made
        Return html
    End Function

    Function ReplaceduplicateNormalHours(html As String, columnIndex As Integer, PaycodeIndex As Integer, PaycodeNewValue As String) As String
        Dim pattern As String = "(<td.*?>.*?</td>)"
        Dim matches As MatchCollection = Regex.Matches(html, pattern)

        ' Extract the part of the string before the <td> elements
        Dim startTag As String = html.Substring(0, html.IndexOf("<td"))

        ' Check if there are enough matches to access the specified column
        If matches.Count >= columnIndex Then
            ' Rebuild the HTML string with the replacement
            Dim modifiedHtml As New Text.StringBuilder()
            modifiedHtml.Append(startTag) ' Append the original start tag (e.g., <tr style="background-color:#C6E0B4">)

            For i As Integer = 0 To matches.Count - 1
                If i + 1 = columnIndex Then
                    modifiedHtml.Append("<td></td>")
                ElseIf i + 1 = PaycodeIndex Then
                    modifiedHtml.Append("<td>" + PaycodeNewValue + "</td>")
                Else
                    modifiedHtml.Append(matches(i).Value)
                End If
            Next

            ' Append the rest of the HTML after the last <td> element
            Dim endTagIndex As Integer = html.IndexOf("</tr>")
            If endTagIndex > -1 Then
                modifiedHtml.Append(html.Substring(endTagIndex))
            End If

            Return modifiedHtml.ToString()
        End If

        ' Return the original HTML if no modification was made
        Return html
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
            Dim filteredTimeEntries() As DataRow = dtTimeEntries.Select("UserPerformed=" & EmployeeID & " and StartTime >= '" & DateToRun.ToString("MM-dd-yyyy") & "' and StartTime < '" & Microsoft.VisualBasic.DateAndTime.DateAdd(DateInterval.Day, 1, DateToRun).ToString("MM-dd-yyyy") & "' and fJobNumber='" & JobNumber & "' and fPW='" & PW & "' and fState='" & state & "'")

            For Each drDay As DataRow In filteredTimeEntries


                StartTime = drDay.Item("StartTime")
                Endtime = drDay.Item("EndTime")
                'Overtime = drDay.Item("Overtime")
                Travel = drDay.Item("fTravelTime")
                PerDiem = drDay.Item("fPerDiem")

                If PerDiem = "" Then
                    PDHrs = "No"
                Else
                    PDHrs = PerDiem
                End If


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
                    ' TotPerDiem = TotPerDiem + Timedifference
                    TotPerDiem = 48.0
                    PerDiemDays = 1
                    PerDiemTot = True
                Else
                    PerDiemTot = False

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
            'If TotStandardTime = 0 Then
            '    Return ""

            'End If
            ReturnString = "<td>" & DateParsed & "</td>"
            ReturnString = ReturnString & "<td>" & IsZero(TotStandardTime) & "</td>"
            ReturnString = ReturnString & "<td>" & IsZero(TotOverTime) & "</td>"
            ReturnString = ReturnString & "<td>" & IsZero(TotCompany) & "</td>"
            ReturnString = ReturnString & "<td>" & IsZero(TotPersonal) & "</td>"

        Catch ex As Exception
            ReturnString = "ERROR: " & ex.ToString & "<br>"
        End Try

        RecTotStandard = RecTotStandard + TotStandardTime
        RecTotOvertime = RecTotOvertime + TotOverTime
        RecTotCompany = RecTotCompany + TotCompany
        RecTotPersonal = RecTotPersonal + TotPersonal
        '  RecTotPerDiem = RecTotPerDiem + TotPerDiem
        RecTotPerDiem = TotPerDiem
        RecTotPerDiemDays = RecTotPerDiemDays + PerDiemDays


        Return ReturnString


    End Function

    Function GenerateHeader(ByVal StartDate As Date, ByVal EndDate As Date) As String

        Dim DateRange As Integer = DateDiff(DateInterval.Day, StartDate, EndDate)

        Dim HeaderLine1 As String = "<tr><th>Company Code</th><th>Task ID</th><th>Employee Name</th><th>Payroll #</th><th>Paycode</th><th></th><th >Job #</th><th></th><th>Full Code</th><th>PW</th><th>State</th><th>Date Worked</th> <th>ST</th><th>OT</th><th>C</th><th>P</th><th>COM</th><th>Code 18</th><th>PER</th><th>Code 19</th><th>PD Hrs</th><th>Code 11</th><th>Bonus</th><th>Notes</th></tr>"



        Return HeaderLine1

    End Function

    Function IsZero(ByVal InputValue As Double) As String

        If InputValue = 0 Then
            Return " "
        Else
            Return InputValue.ToString
        End If

    End Function

    Function GetCompanycode(ByVal InputValue As String) As String

        If InputValue = "G" Then
            Return "8SL"
        Else
            Return "8SK"
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



    Function GetPerDiem(ByVal SettingKey As String) As String

        Dim SettingsText As String = ""
        Dim qCheckSettings As String = "select SettingValue  from tSettings where SettingKey=@SettingKey"
        Dim ConnCheckSettings As New SqlConnection(sConnection)
        Dim CmdCheckSettings As New SqlCommand(qCheckSettings, ConnCheckSettings)
        Dim drCheckSettings As SqlDataReader

        CmdCheckSettings.Parameters.Add(New SqlParameter("@SettingKey", SettingKey))


        Try
            ConnCheckSettings.Open()
            drCheckSettings = CmdCheckSettings.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheckSettings.Read()
                SettingsText = drCheckSettings.Item("SettingValue")
            End While
        Catch ex As Exception

        Finally
            ConnCheckSettings.Close()
        End Try

        Return SettingsText

    End Function


End Class

' Define a custom class to hold the employee record and duplicate row status
Public Class EmployeeRow
    ' Property to store the employee's record as a StringBuilder
    Private _employeeRecord As StringBuilder
    Public Property EmployeeRecord As StringBuilder
        Get
            Return _employeeRecord
        End Get
        Set(ByVal value As StringBuilder)
            _employeeRecord = value
        End Set
    End Property

    ' Property to indicate if the row is a duplicate
    Private _duplicateRow As Boolean
    Public Property DuplicateRow As Boolean
        Get
            Return _duplicateRow
        End Get
        Set(ByVal value As Boolean)
            _duplicateRow = value
        End Set
    End Property

    ' Constructor to initialize the properties
    Public Sub New(ByVal empRecord As StringBuilder, ByVal isDuplicateRow As Boolean)
        _employeeRecord = empRecord
        _duplicateRow = isDuplicateRow
    End Sub
End Class
