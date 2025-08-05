Imports System.Data.SqlClient
Imports System.IO


Partial Class generatecsv
    Inherits GlobalClass

    Dim OutputCSV As Boolean = False
    Dim RowCounter As Integer = 0

    Sub page_load()

        Dim StartDate As Date = Request.QueryString("start")
        Dim EndDate As Date = Request.QueryString("end")
        Dim OfficeID As String = Request.QueryString("office")
        Dim Type As String = Request.QueryString("group")

        Dim CSVstring As String = ""

        If IsPostBack = True Then

            Context.Response.Clear()
            Context.Response.ContentType = "text/csv"
            Context.Response.AddHeader("Content-Disposition", "attachment; filename=payroll_timberline_export_" & StartDate & "-" & EndDate & ".csv")

            'OutputCSV = True
            'CSVstring = GetRecords(StartDate, EndDate, OfficeID, Type)
            CSVstring = GetFormData()

            Dim FileName As String = Replace(StartDate, "/", "") & "-" & Replace(EndDate, "/", "") & "_" & OfficeID & "_" & Type
            SaveCSVFile(CSVstring, FileName)
            Response.Write(CSVstring)
            Response.End()

        Else
            LIT_Grid.Text = GenerateHeader() & GetRecords(StartDate, EndDate, OfficeID, Type) & GenerateFooter()
        End If





    End Sub


    Function GetRecords(StartDate As Date, EndDate As Date, OfficeID As String, Type As String) As String

        Dim QueryCondition As String = "", ReturnString As String = ""

        If Not Type = "" Then
            QueryCondition = QueryCondition & " and tEmployees.Type='" & Type & "' "
        End If

        If Not OfficeID = "" Then
            QueryCondition = QueryCondition & " and OfficeID=" & OfficeID & " "
        End If

        Dim EmployeeID As Integer, JobNumber As String, State As String, DateOfWork As Date, PayrollID As String, EmployeeName As String

        If Not Type = "" Then
            QueryCondition = QueryCondition & " and tEmployees.Type=@Type "
        End If

        If Not OfficeID = "" Then
            QueryCondition = QueryCondition & " and OfficeID=@OfficeID "
        End If

        Dim sEmployeeQuery As String = "select distinct tEmployees.EmployeeID, FirstName + ' ' + LastName as EmployeeName, PayrollID, fJobNumber, fState, CAST(tTimeEntry.StartTime as DATE) as DateOfWork from tEmployees, tShifts, tTimeEntry where tEmployees.EmployeeID = tShifts.UserPerformed and tShifts.ShiftID = tTimeEntry.ShiftID " & QueryCondition & " and CAST(tTimeEntry.StartTime as DATE) >= @StartDate and CAST(tTimeEntry.StartTime as DATE) <= @EndDate order by tEmployees.EmployeeID, DateOfWork"
        Dim connEmployee As New SqlConnection(sConnection)
        Dim cmdEmployee As New SqlCommand(sEmployeeQuery, connEmployee)
        Dim drEmployee As SqlDataReader
        cmdEmployee.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdEmployee.Parameters.Add(New SqlParameter("@EndDate", EndDate))
        If Not OfficeID = "" Then cmdEmployee.Parameters.Add(New SqlParameter("@OfficeID", OfficeID))
        If Not Type = "" Then cmdEmployee.Parameters.Add(New SqlParameter("@Type", Type))




        connEmployee.Open()
        drEmployee = cmdEmployee.ExecuteReader(Data.CommandBehavior.CloseConnection)

        Try
            Dim TimeBeforeFlag1 As Integer = Environment.TickCount

            While drEmployee.Read

                EmployeeID = drEmployee.Item("EmployeeID")
                JobNumber = drEmployee.Item("fJobNumber")
                State = drEmployee.Item("fState")
                DateOfWork = drEmployee.Item("DateOfWork")
                PayrollID = drEmployee.Item("PayrollID")
                EmployeeName = drEmployee.Item("EmployeeName")

                ReturnString = ReturnString & GetTimeForRecord(EmployeeID, JobNumber, State, DateOfWork, PayrollID, EmployeeName)
            End While



        Catch ex As Exception

        Finally
            drEmployee.Close()
            connEmployee.Close()
        End Try

        Return ReturnString


    End Function

    Function GetTimeForRecord(EmployeeID As Integer, JobNumber As String, State As String, DateOfWork As Date, PayrollID As String, EmployeeName As String) As String
        Dim StartTime As DateTime, Endtime As DateTime, Travel As String, PerDiem As String = "", PW As String = "", Comments As String = "", ShiftID As String = ""
        Dim TotStandardTime As Double, TotCompany As Double, TotPersonal As Double, Timedifference As Double, TotPerDiem As Double, TotPW As Double, TotOverTime As Double
        Dim ReturnString As String = ""

        Dim sDayQuery As String = "select StartTime, EndTime, fTravelTime, fPerDiem, fPW, fComments, tShifts.ShiftID from tTimeEntry, tShifts where Approved = 1 and tTimeEntry.ShiftID = tShifts.ShiftID and fJobNumber = @JobNumber and fState=@State and UserPerformed = @EmployeeID and CAST(StartTime as DATE) = @DateOfWork"
        Dim connDay As New SqlConnection(sConnection)
        Dim cmdDay As New SqlCommand(sDayQuery, connDay)
        Dim drDay As SqlDataReader
        cmdDay.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
        cmdDay.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
        cmdDay.Parameters.Add(New SqlParameter("@DateOfWork", DateOfWork))
        cmdDay.Parameters.Add(New SqlParameter("@State", State))

        Try
            connDay.Open()
            drDay = cmdDay.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drDay.Read

                StartTime = drDay.Item("StartTime")
                Endtime = drDay.Item("EndTime")
                Travel = drDay.Item("fTravelTime")
                PerDiem = drDay.Item("fPerDiem")
                PW = drDay.Item("fPW")
                Comments = drDay.Item("fComments")
                ShiftID = drDay.Item("ShiftID")

                Timedifference = DateDiff(DateInterval.Minute, StartTime, Endtime) / 60

                '---- ADD UP TIME ENTRIES 
                If Travel = "N/A" Or Travel = "" Then
                    TotStandardTime = TotStandardTime + Timedifference
                ElseIf Travel = "Company" Then
                    TotCompany = TotCompany + Timedifference
                ElseIf Travel = "Personal" Then
                    TotPersonal = TotPersonal + Timedifference
                End If

                If PerDiem = "Yes" Then
                    TotPerDiem = TotPerDiem + Timedifference
                End If

                'If PW = "Yes" Then
                '    TotPW = TotPW + Timedifference
                'End If


            End While




            '---- get adjustments form adjustments table
            Dim adjStandardTime As Double = 0, adjOvertime As Double = 0, adjCompany As Double = 0, adjPersonal As Double = 0
            Dim sAdjustmentQuery As String = "select isNull(Overtime,0) as Overtime, isNull(Standardtime,0) as StandardTime, isNull(Company,0) as Company, isNull(Personal,0) as Personal from tTimeAdjustments where EmployeeID=@EmployeeID and DateAdjusted=@DateParsed and JobNumber=@JobNumber and PW=@PW and State=@State"
            Dim connAdjustment As New SqlConnection(sConnection)
            Dim cmdAdjustment As New SqlCommand(sAdjustmentQuery, connAdjustment)
            Dim drAdjustment As SqlDataReader
            cmdAdjustment.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
            cmdAdjustment.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
            cmdAdjustment.Parameters.Add(New SqlParameter("@DateParsed", DateOfWork))
            cmdAdjustment.Parameters.Add(New SqlParameter("@PW", PW))
            cmdAdjustment.Parameters.Add(New SqlParameter("@State", State))

            Try
                connAdjustment.Open()
                drAdjustment = cmdAdjustment.ExecuteReader(Data.CommandBehavior.CloseConnection)
                While drAdjustment.Read

                    adjCompany = drAdjustment.Item("Company")
                    adjPersonal = drAdjustment.Item("Personal")
                    adjStandardTime = drAdjustment.Item("StandardTime")
                    adjOvertime = drAdjustment.Item("Overtime")

                End While
            Catch ex As Exception
                Response.Write("ERROR: " & ex.ToString & "<br>")
            Finally
                connAdjustment.Close()
            End Try

            If adjStandardTime > 0 Then
                TotStandardTime = adjStandardTime
            End If

            If adjOvertime > 0 Then
                TotOverTime = adjOvertime
                TotStandardTime = TotStandardTime - adjOvertime
                'add row for overtime
                ReturnString = ReturnString & AddRow(PayrollID, JobNumber, DateOfWork, TotOverTime, 2, EmployeeName, State, Comments, PW, ShiftID)
            End If

            If adjCompany > 0 Then
                TotCompany = adjCompany
            End If

            If adjPersonal > 0 Then
                TotPersonal = adjPersonal
            End If



            If TotStandardTime > 0 Then
                'add row for standard time
                ReturnString = ReturnString & AddRow(PayrollID, JobNumber, DateOfWork, TotStandardTime, 1, EmployeeName, State, Comments, PW, ShiftID)
            End If

            If TotCompany > 0 Then
                'add row for company travel
                ReturnString = ReturnString & AddRow(PayrollID, JobNumber, DateOfWork, TotCompany, 18, EmployeeName, State, Comments, PW, ShiftID)
            End If

            If TotPersonal > 0 Then
                'add row for personal travel
                ReturnString = ReturnString & AddRow(PayrollID, JobNumber, DateOfWork, TotPersonal, 19, EmployeeName, State, Comments, PW, ShiftID)
            End If

            If TotPerDiem > 0 Then
                'add row for per diem
                ReturnString = ReturnString & AddRow(PayrollID, JobNumber, DateOfWork, TotPerDiem, 11, EmployeeName, State, Comments, PW, ShiftID)
            End If

            'If TotPW > 0 Then
            '    'add row for pw
            '    ReturnString = ReturnString & AddRow(PayrollID, JobNumber, DateOfWork, TotPW, 24, EmployeeName, State, Notes, PW)
            'End If

        Catch ex As Exception
            Response.Write(ex.ToString)

        Finally
            drDay.Close()
            connDay.Close()

        End Try

        Return ReturnString


    End Function

    Function AddRow(PayrollID As String, JobNumber As String, DateofWork As Date, Hours As Double, PayID As Integer, employeeName As String, State As String, Comments As String, PW As String, shiftID As String) As String

        RowCounter = RowCounter + 1

        Dim HighlightClass As String = "", PayIDPrint As String = "", ReturnString As String = ""

        Try

            DateofWork = DateofWork.ToString("MM/dd/yyyy")

            'parse job# to look for extra digits
            JobNumber = JobNumber.Trim()
            'JobNumber = Regex.Replace(JobNumber, "[0-9]+", "")

            Dim sbReplace As New StringBuilder(JobNumber.Length)
            For Each ch As Char In JobNumber
                If Char.IsDigit(ch) Then
                    sbReplace.Append(ch)
                End If
            Next

            JobNumber = sbReplace.ToString

            Dim JobCostExtra As String = ""

            If JobNumber.Length > 10 Then
                JobCostExtra = Right(JobNumber, JobNumber.Length - 10)
                JobNumber = Left(JobNumber, 10)
            End If

            PayIDPrint = PayID.ToString

            If IsStateValid(State) = False Then
                PayIDPrint = ""
                HighlightClass = " class=""highlight"""
            End If

            If UCase(PW) = "Y" Or UCase(PW) = "YES" Then
                PayIDPrint = ""
                HighlightClass = " class=""highlight"""
            End If

            'If OutputCSV = True Then

            '    'generate csv line
            '    'we use PayID val from form if we're exporting to CSV
            '    Dim FormPayID As String = Request.Form("payid_" & PayrollID & "_" & JobNumber & "_" & JobCostExtra & "_" & State & "_" & DateofWork & "_" & shiftID & "_" & PayIDPrint)
            '    Dim FormRate As String = Request.Form("rate_" & PayrollID & "_" & JobNumber & "_" & JobCostExtra & "_" & State & "_" & DateofWork & "_" & shiftID & "_" & PayIDPrint)
            '    ReturnString = ReturnString & PayrollID & "," & JobNumber & "," & JobCostExtra & ",,,,,,,,,,,,,,,,," & FormPayID & "," & DateofWork & ",,,,,," & Hours & "," & FormRate & "," & vbCrLf
            'Else
            'generate html
            ReturnString = ReturnString & "<tr" & HighlightClass & ">"
            ReturnString = ReturnString & "<td>" & employeeName & "</td>"
            ReturnString = ReturnString & "<td><input type=""text"" class=""noedit"" readonly=""readonly"" name=""payrollid_" & RowCounter & """ value=""" & PayrollID & """ /></td>"
            ReturnString = ReturnString & "<td><input type=""text"" class=""noedit"" readonly=""readonly"" name=""jobnumber_" & RowCounter & """ value=""" & JobNumber.Replace("-", "") & """ /></td>"
            ReturnString = ReturnString & "<td><input type=""text"" class=""noedit"" readonly=""readonly"" name=""jobcostextra_" & RowCounter & """ value=""" & JobCostExtra & """ /></td>"
            ReturnString = ReturnString & "<td>" & State & "</td>"
            ReturnString = ReturnString & "<td><input type=""text"" class=""noedit"" readonly=""readonly"" name=""dateofwork_" & RowCounter & """ value=""" & DateofWork & """ /></td>"
            ReturnString = ReturnString & "<td><input type=""text"" class=""noedit"" readonly=""readonly"" name=""hours_" & RowCounter & """ value=""" & Hours & """ /></td>"
            ReturnString = ReturnString & "<td><input type=""text"" class=""payidoverride"" name=""payid_" & RowCounter & """ value=""" & PayIDPrint & """ /></td>"
            ReturnString = ReturnString & "<td>" & PW & "</td>"
            ReturnString = ReturnString & "<td style=""text-align:left"">" & Comments & "</td>"
            ReturnString = ReturnString & "<td><input type=""text"" class=""rate"" readonly=""readonly"" name=""rate_" & RowCounter & """ value="""" /></td>"
            ReturnString = ReturnString & "</tr>"
            'End If
        Catch ex As Exception

            ReturnString = ReturnString & "ERROR:" & ex.ToString
            LogError("generatecsv.aspx.vb :: AddRow", ex.ToString)

        End Try

        Return ReturnString

    End Function

    Function GetFormData() As String

        Dim ReturnString As String = ""
        Dim PayrollID As String = "", JobNumber As String = "", JobCostExtra As String = "", DateOfWork As String = "", Hours As String = "", PayID As String = "", Rate As String = ""

        Dim TotalRows As Integer = Request.Form("totalrows")
        Dim x As Integer = 1

        While x <= TotalRows

            PayrollID = Request.Form("payrollid_" & x.ToString)
            JobNumber = Request.Form("jobnumber_" & x.ToString)
            JobCostExtra = Request.Form("jobcostextra_" & x.ToString)
            DateOfWork = Request.Form("dateofwork_" & x.ToString)
            Hours = Request.Form("hours_" & x.ToString)
            PayID = Request.Form("payid_" & x.ToString)
            Rate = Request.Form("rate_" & x.ToString)

            ReturnString = ReturnString & PayrollID & "," & JobNumber.Replace("-", "") & "," & JobCostExtra & ",,,,,,,,,,,,,,,,," & PayID & "," & DateOfWork & ",,,,,," & Hours & "," & Rate & "," & vbCrLf

            x = x + 1
        End While


        Return ReturnString

    End Function

    Function IsStateValid(fState As String) As Boolean

        Dim statesList As String, states As Object
        statesList = ("AL,AK,AZ,AR,CA,CO,CT,DC,DE,FL,GA,HI,ID,IL,IN,IA,KS,KY,LA,ME,MD,MA,MI,MN,MS,MO,MT,NE,NV,NH,NJ,NM,NY,NC,ND,OH,OK,OR,PA,RI,SC,SD,TN,TX,UT,VT,VA,WA,WV,WI,WY,")
        states = Split(statesList, ",")

        If Not Array.IndexOf(states, UCase(fState)) > 0 Then
            Return False
        Else
            Return True
        End If

    End Function


    Function GenerateHeader() As String

        Return "<table class=""payroll"" id=""generatecsv"" style=""margin:30px auto 0 auto;""><thead><tr><th>Employee</th><th>Payroll #</th><th>Job #</th><th>Job Cost Extra</th><th>State</th><th>Date of Work</th><th>Hours</th><th>Pay ID</th><th>PW</th><th>Comments</th><th>Rate</th></tr></thead><tbody>"

    End Function

    Function GenerateFooter() As String

        Return "</tbody></table><input type=""hidden"" name=""totalrows"" value=""" & RowCounter & """ />"

    End Function

    Sub SaveCSVFile(ByVal ReturnString As String, FileNamePart As String)

        Dim fp As StreamWriter
        Dim TimeStamp As String = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
        Dim Filename As String = TimeStamp & "_" & Session("UserID") & "_" & FileNamePart & ".csv"

        Try
            fp = File.CreateText(Server.MapPath(".\CSVFiles\") & Filename)
            fp.WriteLine(ReturnString)
            fp.Close()
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("generatecsv.aspx :: SaveCSVFile", ex.ToString)
        Finally

        End Try

    End Sub

End Class