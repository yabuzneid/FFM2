Imports System.Data.SqlClient
Imports System.Data


Partial Class addtime
    Inherits GlobalClass

    Sub page_load()


        AddOfficesToDropdown()


        If IsPostBack = False Then
            DDL_State.SelectedValue = "CT"
            DDL_StartTime.SelectedValue = "09:00"
            DDL_EndTime.SelectedValue = "17:00"
            TXT_StartDate.Text = Now().Date
            TXT_EndDate.Text = Now().Date

            PopulateFromCookies()

            TXT_StartDateMulti.Text = Now().Date
            TXT_EndDateMulti.Text = Now().Date



            If Session("PM") = 1 Then
                If Session("Worker") = 1 Then
                    DDL_Employee.SelectedValue = Session("UserID")
                End If
            End If


        End If

        If Session("PM") = 1 Then
            Panel_Employees.Visible = True
        End If


    End Sub


    Sub SubmitShift(ByVal sender As Object, ByVal e As EventArgs)

        Dim StartDateTime As DateTime, EndDateTime As DateTime


        Dim CheckStartDate As DateTime

        If RB_MultiDay.Checked = True Then
            CheckStartDate = TXT_StartDateMulti.Text
        Else
            CheckStartDate = TXT_StartDate.Text
        End If


        If IsDayLocked(CheckStartDate) = True Then

            LBL_Info.Visible = True
            LBL_Info.Text = "You cannot add time into a week that has already been closed by payroll."

        Else

            Dim EmployeeID As String
            EmployeeID = Session("UserID")

            If Session("PM") = 1 Then
                EmployeeID = DDL_Employee.SelectedValue
            End If

            Dim OfficeClause As String = ""

            If EmployeeID = "-1" Or Left(EmployeeID, 6) = "office" Or Left(EmployeeID, 5) = "group" Then

                If Not EmployeeID = "-1" Then

                    'only for specific office
                    If Left(EmployeeID, 6) = "office" Then
                        OfficeClause = " and officeID = " & Right(EmployeeID, Len(EmployeeID) - 6)
                    End If

                    'only for specific group
                    If Left(EmployeeID, 5) = "group" Then
                        OfficeClause = " and EmployeeGroup = '" & Right(EmployeeID, Len(EmployeeID) - 5) & "'"
                    End If
                End If


                Dim qEmployees As String = "select EmployeeID from tEmployees where Active = 1 and Worker = 1 and PartTime <> 1 " & OfficeClause
                Dim connEmployee As New SqlConnection(sConnection)
                Dim cmdEmployee As New SqlCommand(qEmployees, connEmployee)
                Dim drEmployee As SqlDataReader

                Try
                    connEmployee.Open()
                    drEmployee = cmdEmployee.ExecuteReader(CommandBehavior.CloseConnection)
                    While drEmployee.Read()

                        GenerateShifts(StartDateTime, EndDateTime, drEmployee("EmployeeID"))

                    End While


                Catch ex As Exception
                    Response.Write(ex.Message)
                    LogError("addtime.aspx.vb :: SubmitShift", ex.ToString)
                Finally
                    connEmployee.Close()
                End Try

            Else

                GenerateShifts(StartDateTime, EndDateTime, EmployeeID)


            End If

            If Request.QueryString("action") = "redirectback" Then
                Response.Redirect(Session("RedirectBack"))
            Else
                Response.Redirect("entertime.aspx")
            End If


            End If


    End Sub

    Sub GenerateShifts(ByVal StartDateTime As DateTime, ByVal EndDateTime As DateTime, ByVal EmployeeID As Integer)


        If RB_MultiDay.Checked = True Then

            Dim StartDate As Date, EndDate As Date

            StartDate = TXT_StartDateMulti.Text
            EndDate = TXT_EndDateMulti.Text

            Dim RunDate As Date = StartDate

            While RunDate <= EndDate
                If Not RunDate.DayOfWeek = 6 And Not RunDate.DayOfWeek = 0 Then
                    Dim NewShiftID As Integer = CreateShift(EmployeeID)

                    StartDateTime = RunDate & " " & DDL_StartTimeMulti.SelectedValue
                    EndDateTime = RunDate & " " & DDL_EndTimeMulti.SelectedValue

                    AddTimeEntry(NewShiftID, StartDateTime, EndDateTime)

                End If

                RunDate = DateAdd(DateInterval.Day, 1, RunDate)
            End While

        Else

            Dim NewShiftID As Integer = CreateShift(EmployeeID)

            StartDateTime = TXT_StartDate.Text & " " & DDL_StartTime.SelectedValue
            EndDateTime = TXT_EndDate.Text & " " & DDL_EndTime.SelectedValue

            AddTimeEntry(NewShiftID, StartDateTime, EndDateTime)

        End If

        WriteCookies()





    End Sub

    Sub CancelShift(ByVal sender As Object, ByVal e As EventArgs)
        Response.Redirect("entertime.aspx")
    End Sub

    Function CreateShift(ByVal EmployeeID As Integer) As Integer

        Dim PM As String, State As String, JobNumber As String, PW As String, Project As String, Type As String, ShiftType As String, Travel As String, PerDiem As String, Injured As String, Comments As String

        PM = DDL_PM.SelectedValue
        State = DDL_State.SelectedValue
        JobNumber = TXT_JobNumber.Text
        PW = DDL_PW.SelectedValue
        Project = TXT_Project.Text
        Type = "" 'TXT_Type.Text
        ShiftType = DDL_ShiftType.SelectedValue
        Travel = "" 'DDL_Travel.SelectedValue
        PerDiem = DDL_PerDiem.SelectedValue
        Injured = DDL_Injured.SelectedValue
        Comments = TXT_Comments.Text


        Dim ReturnedShiftID As Integer
        Dim sQuery As String = "exec CreateShift @UserPerformed, @fType, @fShiftType, @fJobNumber, @fProjectName, @fPM, @fPerDiem, @fTravelTime, @fState, @fPW, @fInjured, @fComments"
        Dim connNewShift As New SqlConnection(sConnection)
        Dim cmdNewShift As New SqlCommand(sQuery, connNewShift)
        Dim drNewShift As SqlDataReader
        cmdNewShift.Parameters.Add(New SqlParameter("@UserPerformed", EmployeeID))
        cmdNewShift.Parameters.Add(New SqlParameter("@fType", Type))
        cmdNewShift.Parameters.Add(New SqlParameter("@fShiftType", ShiftType))
        cmdNewShift.Parameters.Add(New SqlParameter("@fJobNumber", JobNumber))
        cmdNewShift.Parameters.Add(New SqlParameter("@fProjectName", Project))
        cmdNewShift.Parameters.Add(New SqlParameter("@fPM", PM))
        cmdNewShift.Parameters.Add(New SqlParameter("@fPerDiem", PerDiem))
        cmdNewShift.Parameters.Add(New SqlParameter("@fTravelTime", Travel))
        cmdNewShift.Parameters.Add(New SqlParameter("@fState", State))
        cmdNewShift.Parameters.Add(New SqlParameter("@fPW", PW))
        cmdNewShift.Parameters.Add(New SqlParameter("@fInjured", Injured))
        cmdNewShift.Parameters.Add(New SqlParameter("@fComments", Comments))

        Try
            connNewShift.Open()
            drNewShift = cmdNewShift.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            While drNewShift.Read()
                ReturnedShiftID = drNewShift.Item(0)
            End While

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString)
        Finally
            connNewShift.Close()
        End Try

        Return ReturnedShiftID

    End Function

    Function AddTimeEntry(ByVal ShiftID As Integer, ByVal StartDateTime As DateTime, ByVal EndDateTime As DateTime) As String


        Dim sQuery As String = "insert into tTimeEntry (ShiftID,StartTime,EndTime,StartAddress,EndAddress, StartLatitude, StartLongitude, EndLatitude, EndLongitude, EnteredOn, EnteredManually) values(@ShiftID,@StartTime,@Endtime,@StartAddress,@EndAddress,@StartLatitude,@StartLongitude,@EndLatitude,@EndLongitude, @EnteredOn, @EnteredManually)"
        Dim connNewShift As New SqlConnection(sConnection)
        Dim cmdNewShift As New SqlCommand(sQuery, connNewShift)
        cmdNewShift.Parameters.Add(New SqlParameter("@ShiftID", ShiftID))
        cmdNewShift.Parameters.Add(New SqlParameter("@StartTime", StartDateTime))
        cmdNewShift.Parameters.Add(New SqlParameter("@EndTime", EndDateTime))
        cmdNewShift.Parameters.Add(New SqlParameter("@StartAddress", ""))
        cmdNewShift.Parameters.Add(New SqlParameter("@EndAddress", ""))
        cmdNewShift.Parameters.Add(New SqlParameter("@StartLatitude", ""))
        cmdNewShift.Parameters.Add(New SqlParameter("@StartLongitude", ""))
        cmdNewShift.Parameters.Add(New SqlParameter("@EndLatitude", ""))
        cmdNewShift.Parameters.Add(New SqlParameter("@EndLongitude", ""))
        cmdNewShift.Parameters.Add(New SqlParameter("@EnteredOn", Now()))
        cmdNewShift.Parameters.Add(New SqlParameter("@EnteredManually", 2))

        Try
            connNewShift.Open()
            cmdNewShift.ExecuteNonQuery()
        Catch ex As Exception
            LogError("addtime.aspx.vb :: AddTimeEntry", ex.ToString)
            Response.Write("ERROR: " & ex.ToString & "<br><br>")

        Finally
            connNewShift.Close()
        End Try


    End Function

    Sub SwitchTimeType(ByVal sender As Object, ByVal e As EventArgs)

        If RB_MultiDay.Checked = True Then
            Panel_SingleDay.Visible = False
            Panel_MultiDay.Visible = True

            DDL_StartTimeMulti.SelectedValue = "09:00"
            DDL_EndTimeMulti.SelectedValue = "17:00"


        Else
            Panel_SingleDay.Visible = True
            Panel_MultiDay.Visible = False
        End If
    End Sub

    Sub WriteCookies()

        Response.Cookies("enter_starttime").Value = DDL_StartTime.SelectedValue
        Response.Cookies("enter_endtime").Value = DDL_EndTime.SelectedValue
        Response.Cookies("enter_pm").Value = DDL_PM.SelectedValue
        Response.Cookies("enter_state").Value = DDL_State.SelectedValue
        Response.Cookies("enter_jobnumber").Value = TXT_JobNumber.Text
        Response.Cookies("enter_pw").Value = DDL_PW.SelectedValue
        Response.Cookies("enter_project").Value = TXT_Project.Text
        Response.Cookies("enter_shifttype").Value = DDL_ShiftType.SelectedValue
        Response.Cookies("enter_travel").Value = DDL_Travel.SelectedValue
        Response.Cookies("enter_perdiem").Value = DDL_PerDiem.SelectedValue
        Response.Cookies("enter_injured").Value = DDL_Injured.SelectedValue

        Dim ExpirationDate As DateTime = DateAdd(DateInterval.Year, 1, Now())

        Response.Cookies("enter_starttime").Expires = ExpirationDate
        Response.Cookies("enter_endtime").Expires = ExpirationDate
        Response.Cookies("enter_pm").Expires = ExpirationDate
        Response.Cookies("enter_state").Expires = ExpirationDate
        Response.Cookies("enter_jobnumber").Expires = ExpirationDate
        Response.Cookies("enter_pw").Expires = ExpirationDate
        Response.Cookies("enter_project").Expires = ExpirationDate
        Response.Cookies("enter_shifttype").Expires = ExpirationDate
        Response.Cookies("enter_travel").Expires = ExpirationDate
        Response.Cookies("enter_perdiem").Expires = ExpirationDate
        Response.Cookies("enter_injured").Expires = ExpirationDate


    End Sub

    Sub PopulateFromCookies()

        Try

            If Not Request.Cookies("enter_starttime") Is Nothing Then DDL_StartTime.SelectedValue = Request.Cookies("enter_starttime").Value
            If Not Request.Cookies("enter_endtime") Is Nothing Then DDL_EndTime.SelectedValue = Request.Cookies("enter_endtime").Value
            If Not Request.Cookies("enter_pm") Is Nothing And CheckIfPMExists(Request.Cookies("enter_pm").Value) = True Then DDL_PM.SelectedValue = Request.Cookies("enter_pm").Value
            If Not Request.Cookies("enter_state") Is Nothing Then DDL_State.SelectedValue = Request.Cookies("enter_state").Value
            If Not Request.Cookies("enter_jobnumber") Is Nothing Then TXT_JobNumber.Text = Request.Cookies("enter_jobnumber").Value
            If Not Request.Cookies("enter_pw") Is Nothing Then DDL_PW.SelectedValue = Request.Cookies("enter_pw").Value
            If Not Request.Cookies("enter_project") Is Nothing Then TXT_Project.Text = Request.Cookies("enter_project").Value
            If Not Request.Cookies("enter_shifttype") Is Nothing Then DDL_ShiftType.SelectedValue = Request.Cookies("enter_shifttype").Value
            If Not Request.Cookies("enter_travel") Is Nothing Then DDL_Travel.SelectedValue = Request.Cookies("enter_travel").Value
            If Not Request.Cookies("enter_perdiem") Is Nothing Then DDL_PerDiem.SelectedValue = Request.Cookies("enter_perdiem").Value
            If Not Request.Cookies("enter_injured") Is Nothing Then DDL_Injured.SelectedValue = Request.Cookies("enter_injured").Value
        Catch ex As exception
            'error populating form, so nevermind
        End Try
    End Sub

    Sub AddOfficesToDropDown()

                Dim OfficeID As Integer, OfficeName As String

        Dim sqlOffices As String = "Select OfficeID, OfficeName from tOffices order by OfficeName"
        Dim connOffices As New SqlConnection(sConnection)
        Dim cmdOffices As New SqlCommand(sqlOffices, connOffices)
        Dim drOffices As SqlDataReader


        Try
            connOffices.Open()
            drOffices = cmdOffices.ExecuteReader(CommandBehavior.CloseConnection)
            While drOffices.Read()
                OfficeID = drOffices.Item("OfficeID")
                OfficeName = drOffices.Item("OfficeName")

                DDL_Employee.Items.Add(New ListItem("- ALL " & OfficeName & " -", "office" & OfficeID.ToString))

            End While

            DDL_Employee.Items.Add(New ListItem("- ALL GCI -", "groupGCI"))
            DDL_Employee.Items.Add(New ListItem("- ALL KCI -", "groupKCI"))

        Catch ex As Exception

        End Try

    End Sub

    Function CheckIfPMExists(Initials As String) As Boolean

        Dim PMExists As Boolean = False

        Dim qCheck As String = "select count(EmployeeID) as PMCount from tEmployees where Active = 1 and PM = 1 and Initials=@Initials"
        Dim connCheck As New SqlConnection(sConnection)
        Dim cmdCheck As New SqlCommand(qCheck, connCheck)
        Dim drCheck As SqlDataReader
        cmdCheck.Parameters.Add(New SqlParameter("@Initials", Initials))

        Try
            connCheck.Open()
            drCheck = cmdCheck.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheck.Read
                If drCheck.Item(0) > 0 Then
                    PMExists = True
                End If
            End While
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("addtime.aspx.vb :: CheckIfPMExists", ex.ToString)
        Finally
            connCheck.Close()
        End Try

        Return PMExists

    End Function


End Class
