'Imports gwUserStatus
'Imports gwCompanyAdmin
Imports prodCompanyAdmin
Imports prodUserStatus
Imports System.Data.SqlClient

Partial Class LoadData
    Inherits GlobalClass

    Sub page_load()

        If Request.QueryString("action") = "sync" Then
            LoadTimeCards()
        End If

    End Sub

    Sub GenerateDates(ByVal sender As Object, ByVal e As System.EventArgs)

        Dim datePatt As String = "yyyy-MM-dd HH:mm:ss.fff"
        Dim dLastDate As DateTime = GetLastdate()

        dLastDate = dLastDate.ToUniversalTime()

        'dLastDate = dLastDate.AddSeconds(1)

        Dim sStartDate As String = CDate(dLastDate).ToString(datePatt)
        Dim dEndDate As DateTime = DateAdd(DateInterval.Day, 2, dLastDate)

        If dEndDate > Now() Then dEndDate = Now().ToUniversalTime

        Dim sEndDate As String = dEndDate.ToString(datePatt)

        'LBL_test.Text = "Last Date (date):" & GetLastdate() & "<br>"
        'LBL_test.Text = LBL_test.Text & "Last Date (string): " & sStartDate & "<br>"
        'LBL_test.Text = LBL_test.Text & "End Date (string): " & sEndDate & "<br><br>"


    End Sub

    Sub LoadTimeCards()

        'Dim StartAndEnd As String = ""

        Dim datePatt As String = "yyyy-MM-dd HH:mm:ss.fff"
        Dim dLastDate As DateTime = GetLastdate()

        'dLastDate = dLastDate.ToUniversalTime()

        'dLastDate = dLastDate.AddSeconds(1)

        Dim sStartDate As String = CDate(dLastDate).ToString(datePatt)
        Dim dEndDate As DateTime = DateAdd(DateInterval.Hour, 6, dLastDate)

        If dEndDate > Now().ToUniversalTime Then dEndDate = Now().ToUniversalTime

        Dim sEndDate As String = dEndDate.ToString(datePatt)

        Dim startDate As New prodUserStatus.SvDate
        startDate.dateString = sStartDate
        Dim EndDate As New prodUserStatus.SvDate
        EndDate.dateString = sEndDate

        'StartAndEnd = " :: " & startDate.ToString & " :: " & EndDate.ToString

        Try

            Dim wsConnection As New prodUserStatus.UserStatusWebServiceProxyService

            Dim login As New System.Net.NetworkCredential
            login.UserName = "comnetws@95333"
            login.Password = "c0MmN35"

            wsConnection.Credentials = login

            Dim myResultSet As SvUserAction()
            Dim strResults As String = ""

            Dim myResult As New SvUserAction

            'get timecards
            myResultSet = wsConnection.getTimecardActions(startDate, EndDate)

            'check the soap log again to make sure we're not making a duplicate
            Dim LastStartDate As DateTime = CheckLastSOAPCall()
            If Not LastStartDate = dLastDate Then

                LogSOAP(sStartDate, sEndDate, "")

                Dim ActionEvent As String = "", ActionTime As String = "", UserPerformed As String = "", Address As String = "", fType As String = "", fShiftType As String = "", fJobNumber As String = "", fProjectName As String = "", fPM As String = "", fPerDiem As String = "", fTravelTime As String = "", fState As String = "", fInjured As String = "", fComments As String = "", Latitude As String = "", Longitude As String = "", fPW As String = ""

                For Each myResult In myResultSet
                    ActionEvent = myResult.actionName.ToString
                    ActionTime = myResult.executionDateTime.dateString
                    UserPerformed = myResult.workerName

                    fType = ""
                    fShiftType = ""
                    fJobNumber = ""
                    fProjectName = ""
                    fPM = ""
                    fPerDiem = ""
                    fTravelTime = ""
                    fState = ""
                    fInjured = ""
                    fComments = ""
                    fPW = ""

                    Dim myLoc As prodUserStatus.SvLocation = myResult.location

                    If myLoc Is Nothing Then
                        Address = ""
                        Latitude = ""
                        Longitude = ""
                    Else
                        If Not myResult.location.address Is Nothing Then
                            Dim myAddress As prodUserStatus.SvAddress = myLoc.address
                            Address = myAddress.streetAddress.ToString & ", " & myAddress.city.ToString & " " & myAddress.state.ToString & ", " & myAddress.zip.ToString
                        Else
                            Address = ""
                        End If

                        If Not myResult.location.position Is Nothing Then
                            Latitude = myLoc.position.latitude.ToString
                            Longitude = myLoc.position.longitude.ToString
                        Else
                            Latitude = ""
                            Longitude = ""
                        End If

                    End If

                    Dim MyForm As SvForm = myResult.form

                    If Not MyForm Is Nothing Then

                        Dim MyFormField As SvFormField

                        If MyForm.formName = "ShiftInformationForm" Then
                            'process shift start form

                            For Each MyFormField In MyForm.formData
                                If MyFormField.fieldName = "Type" Then
                                    fType = MyFormField.fieldValue
                                ElseIf MyFormField.fieldName = "ShiftType" Then
                                    fShiftType = MyFormField.fieldValue
                                ElseIf MyFormField.fieldName = "JobNumber" Then
                                    fJobNumber = MyFormField.fieldValue
                                ElseIf MyFormField.fieldName = "ProjectName" Then
                                    fProjectName = MyFormField.fieldValue
                                ElseIf MyFormField.fieldName = "PM" Then
                                    fPM = MyFormField.fieldValue
                                ElseIf MyFormField.fieldName = "PerDiem" Then
                                    fPerDiem = MyFormField.fieldValue
                                ElseIf MyFormField.fieldName = "TravelTime" Then
                                    fTravelTime = MyFormField.fieldValue
                                ElseIf MyFormField.fieldName = "State" Then
                                    fState = MyFormField.fieldValue
                                ElseIf MyFormField.fieldName = "PW" Then
                                    fPW = MyFormField.fieldValue
                                End If
                            Next

                        ElseIf MyForm.formName = "InjuryForm" Then
                            'process injury form

                            For Each MyFormField In MyForm.formData
                                If MyFormField.fieldName = "WereYouHurtToday" Then
                                    fInjured = MyFormField.fieldValue
                                ElseIf MyFormField.fieldName = "Comments" Then
                                    fComments = MyFormField.fieldValue
                                End If
                            Next

                        End If

                    End If

                    Dim sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("FullConnection").ConnectionString
                    Dim sqlActions As String = "insert into tActions(ActionEvent, ActionTime, UserPerformed, Address, fType, fShiftType, fJobNumber, fProjectName, fPM, fPerDiem, fTravelTime, fState, fInjured, fComments, gpsLatitude, gpsLongitude, fPW) values (@ActionEvent, @ActionTime, @UserPerformed, @Address, @fType, @fShiftType, @fJobNumber, @fProjectName, @fPM, @fPerDiem, @fTravelTime, @fState, @fInjured, @fComments, @Latitude, @Longitude, @fPW)"
                    Dim connActions As New SqlConnection(sConnection)
                    Dim cmdActions As New SqlCommand(sqlActions, connActions)
                    cmdActions.Parameters.Add(New SqlParameter("ActionEvent", ActionEvent))
                    cmdActions.Parameters.Add(New SqlParameter("ActionTime", ActionTime))
                    cmdActions.Parameters.Add(New SqlParameter("UserPerformed", UserPerformed))
                    cmdActions.Parameters.Add(New SqlParameter("Address", Address))
                    cmdActions.Parameters.Add(New SqlParameter("fType", fType))
                    cmdActions.Parameters.Add(New SqlParameter("fShiftType", fShiftType))
                    cmdActions.Parameters.Add(New SqlParameter("fJobNumber", fJobNumber))
                    cmdActions.Parameters.Add(New SqlParameter("fProjectName", fProjectName))
                    cmdActions.Parameters.Add(New SqlParameter("fPM", fPM))
                    cmdActions.Parameters.Add(New SqlParameter("fPerDiem", fPerDiem))
                    cmdActions.Parameters.Add(New SqlParameter("fTravelTime", fTravelTime))
                    cmdActions.Parameters.Add(New SqlParameter("fState", fState))
                    cmdActions.Parameters.Add(New SqlParameter("fInjured", fInjured))
                    cmdActions.Parameters.Add(New SqlParameter("fComments", fComments))
                    cmdActions.Parameters.Add(New SqlParameter("Latitude", Latitude))
                    cmdActions.Parameters.Add(New SqlParameter("Longitude", Longitude))
                    cmdActions.Parameters.Add(New SqlParameter("fPW", fPW))


                    connActions.Open()

                    Try
                        cmdActions.ExecuteNonQuery()
                    Catch ex As Exception
                        Response.Write("ERROR: " & ex.ToString)
                        LogError("LoadData.aspx.vb :: LoadTimeCards 1 :: " & startDate.ToString & " :: " & EndDate.ToString, ex.ToString)

                    Finally
                        cmdActions.Connection.Close()
                        connActions.Close()
                    End Try
                Next

                strResults = strResults & "</table>"

                'LBL_test.Text = strResults



                ParseTimeCards()
            Else

                LogSOAP(sStartDate, sEndDate, "Skipped due to duplicate")

            End If
        Catch ex As Exception

            'LBL_test.Text = ex.ToString
            LIT_Status.Text = "Could not sync to FieldForce Manager"
            LogError("LoadData.aspx.vb :: LoadTimeCards 2 :: " & startDate.dateString & " :: " & EndDate.dateString, ex.Message & " :: " & ex.ToString & " :: start stack trace :: " & ex.StackTrace)


        End Try

    End Sub

    Function GetLastdate() As DateTime

        Dim ReturnDate As DateTime

        Dim sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("FullConnection").ConnectionString
        Dim sqlDate As String = "select top 1 IsNull(DateRangeEnd,'11/1/2009') as ActionTime from tSoapLog order by DateRangeEnd desc"
        Dim connDate As New SqlConnection(sConnection)
        Dim cmdDate As New SqlCommand(sqlDate, connDate)
        Dim drDate As SqlDataReader

        Try
            connDate.Open()
            drDate = cmdDate.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            If drDate.HasRows = True Then
                While drDate.Read()
                    ReturnDate = drDate.Item("ActionTime") '.ToString("yyyy-MM-dd HH:mm:ss.fff")
                    'GetLastdate = drDate.Item("ActionTime")
                End While
            Else
                ReturnDate = DateSerial(2009, 11, 1)
            End If

        Catch ex As Exception
            LogError("LoadData.aspx :: GetLastDate", ex.ToString)
            'ReturnString = "ERROR: " & ex.ToString
        Finally
            connDate.Close()
        End Try

        GetLastdate = ReturnDate

    End Function

    Function CheckLastSOAPCall() As DateTime

        Dim ReturnDate As DateTime

        Dim sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("FullConnection").ConnectionString
        Dim sqlDate As String = "select top 1 IsNull(DateRangeStart,'11/1/2009') as DateRangeStart from tSoapLog order by RequestTime desc"
        Dim connDate As New SqlConnection(sConnection)
        Dim cmdDate As New SqlCommand(sqlDate, connDate)
        Dim drDate As SqlDataReader

        Try
            connDate.Open()
            drDate = cmdDate.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            If drDate.HasRows = True Then
                While drDate.Read()
                    ReturnDate = drDate.Item("DateRangeStart") '.ToString("yyyy-MM-dd HH:mm:ss.fff")
                    'GetLastdate = drDate.Item("ActionTime")
                End While
            Else
                ReturnDate = DateSerial(2009, 11, 1)
            End If

        Catch ex As Exception
            LogError("LoadData.aspx :: GetLastDate", ex.ToString)
            'ReturnString = "ERROR: " & ex.ToString
        Finally
            connDate.Close()
        End Try

        CheckLastSOAPCall = ReturnDate

    End Function

    Sub ParseTimeCards()

        Dim sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("FullConnection").ConnectionString
        Dim sQuery As String = "select ActionID, ActionEvent, ActionTime, UserPerformed, Address, gpsLatitude, gpsLongitude from tActions where (ActionEvent = 'Start Shift' OR ActionEvent = 'End Shift' OR ActionEvent = 'Start Break' OR ActionEvent = 'End Break') and NOT Processed = 1 order by UserPerformed, ActionTime"
        Dim connActions As New SqlConnection(sConnection)
        Dim cmdActions As New SqlCommand(sQuery, connActions)
        Dim drActions As SqlDataReader

        Try
            connActions.Open()

            drActions = cmdActions.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            Dim ActionID As Integer, ActionEvent As String, ActionTime As DateTime, UserPerformed As String, LastUser As String = "", Address As String, Latitude As String, Longitude As String, TimeEntryStartedLatitude As String = "", TimeEntryStartedLongitude As String = "", LastActionEvent As String = ""
            Dim ShiftID As Integer
            Dim TimeEntryStartedTime As DateTime, TimeEntryStartedAddress As String = "", TimeEntryStartedActionID As Integer
            Dim MarkProcessed As Boolean = False
            'Response.Write("<table border=1>")

            While drActions.Read()

                ActionID = drActions.Item("ActionID")
                ActionEvent = drActions.Item("ActionEvent")
                ActionTime = drActions.Item("ActionTime")
                UserPerformed = drActions.Item("UserPerformed")
                Address = drActions.Item("Address")
                Latitude = drActions.Item("gpsLatitude")
                Longitude = drActions.Item("gpsLongitude")
                MarkProcessed = False

                'Response.Write("<tr><td>" & ActionID & " - " & ActionEvent & "</td>")
                'Response.Write("<td>" & ShiftID & "</td>")
                'Response.Write("<td>last:<br>" & LastUser & "</td>")
                'Response.Write("<td>curr:<br>" & UserPerformed & "</td>")
                'Response.Write("<td>" & ActionTime & "</td>")


                If LastActionEvent = "Start Shift" And Not UserPerformed = LastUser Then
                    DeleteShift(ShiftID)
                End If


                If ActionEvent = "Start Shift" Then

                    ShiftID = CreateShift(ActionID)

                    TimeEntryStartedTime = ActionTime
                    TimeEntryStartedAddress = Address
                    TimeEntryStartedLatitude = Latitude
                    TimeEntryStartedLongitude = Longitude
                    TimeEntryStartedActionID = ActionID

                ElseIf ActionEvent = "Start Break" Then

                    If UserPerformed = LastUser Then
                        InsertTimeEntry(ShiftID, Roundtime(TimeEntryStartedTime), Roundtime(ActionTime), TimeEntryStartedAddress, Address, TimeEntryStartedLatitude, TimeEntryStartedLongitude, Latitude, Longitude)
                        MarkProcessed = True
                        'Response.Write("<td>starting break on " & ShiftID & "</td>")
                    End If

                ElseIf ActionEvent = "End Break" Then
                    'Response.Write("<td>ending break on " & ShiftID)

                    If Not LastUser = UserPerformed Then
                        ShiftID = FindLastShift(GetUserID(UserPerformed), ActionID)
                        'Response.Write("FINDING LAST=" & ShiftID)
                    End If
                    'Response.Write("</td>")
                    TimeEntryStartedTime = ActionTime
                    TimeEntryStartedAddress = Address
                    TimeEntryStartedLatitude = Latitude
                    TimeEntryStartedLongitude = Longitude
                    TimeEntryStartedActionID = ActionID

                ElseIf ActionEvent = "End Shift" Then
                    If UserPerformed = LastUser Then
                        'Response.Write("<td>ending shift id=" & ShiftID & "::actionid=" & ActionID & "</td>")
                        EndShift(ActionID, ShiftID)
                        InsertTimeEntry(ShiftID, Roundtime(TimeEntryStartedTime), Roundtime(ActionTime), TimeEntryStartedAddress, Address, TimeEntryStartedLatitude, TimeEntryStartedLongitude, Latitude, Longitude)
                        MarkProcessed = True
                    End If
                End If

                'Response.Write("<td>MarkProcessed:" & MarkProcessed & "</td><td>ShiftID:" & ShiftID & "</td></tr>")
                'Response.Write("<tr><td colspan=8><hr></td></tr>")

                If MarkProcessed = True Then
                    MarkAsProcessed(ActionID)
                    MarkAsProcessed(TimeEntryStartedActionID)

                End If

                LastUser = UserPerformed
                LastActionEvent = ActionEvent

            End While

            If LastActionEvent = "Start Shift" Then
                DeleteShift(ShiftID)
            End If

            'Response.Write("</table>")

            LIT_Status.Text = "Sync completed. You can close this window."

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString)
            LogError("LoadData.aspx.vb :: ParseTimeCards", ex.ToString)
        Finally
            connActions.Close()
        End Try


    End Sub

    Function Roundtime(ByVal OriginalTime As DateTime) As DateTime

        'Dim RoundedTime As DateTime = DateAdd(DateInterval.Hour, -5, OriginalTime)  'Adjust from Universal
        Dim RoundedTime As DateTime = OriginalTime.ToLocalTime()
        Dim RoundedMins As Integer
        Dim RoundedHour As Integer = OriginalTime.Hour

        Dim OrMinsSecs As Integer = OriginalTime.Minute * 60 + OriginalTime.Second

        If OrMinsSecs <= (7 * 60 + 29) Then
            RoundedMins = 0
        ElseIf OrMinsSecs >= (7 * 60 + 30) And OrMinsSecs <= (22 * 60 + 29) Then
            RoundedMins = 15
        ElseIf OrMinsSecs >= (22 * 60 + 30) And OrMinsSecs <= (37 * 60 + 29) Then
            RoundedMins = 30
        ElseIf OrMinsSecs >= (37 * 60 + 30) And OrMinsSecs <= (52 * 60 + 29) Then
            RoundedMins = 45
        ElseIf OrMinsSecs >= (52 * 60 + 30) Then
            RoundedMins = 0
            RoundedTime = RoundedTime.AddHours(1)
        End If

        RoundedTime = CDate(RoundedTime.Year & "-" & RoundedTime.Month & "-" & RoundedTime.Day & " " & RoundedTime.Hour & ":" & RoundedMins & ":00")

        'LBL_test.Text = LBL_test.Text & OriginalTime & "==" & RoundedTime & "<br>"

        Return RoundedTime

    End Function

    Function CreateShift(ByVal ActionID As Integer) As Integer

        Dim UserPerformed As String = "", fType As String = "", fShiftType As String = "", fJobNumber As String = "", fProjectName As String = "", fPM As String = "", fPerDiem As String = "", fTravelTime As String = "", fState As String = "", fPW As String = ""

        Dim ReturnedShiftID As Integer = -1, UserPerformedID As Integer

        Dim sQuery As String = "select UserPerformed, fType, fShiftType, fJobNumber, fProjectName, fPM, fPerDiem, fTravelTime, fState, fPW from tActions where ActionID = @ActionID"
        Dim connShiftData As New SqlConnection(sConnection)
        Dim cmdShiftData As New SqlCommand(sQuery, connShiftData)
        Dim drShiftData As SqlDataReader
        cmdShiftData.Parameters.Add(New SqlParameter("@ActionID", ActionID))

        Try
            connShiftData.Open()
            drShiftData = cmdShiftData.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            While drShiftData.Read

                UserPerformed = drShiftData.Item("UserPerformed")
                fType = drShiftData.Item("fType")
                fShiftType = drShiftData.Item("fShiftType")
                fJobNumber = drShiftData.Item("fJobNumber")
                fProjectName = drShiftData.Item("fProjectName")
                fPM = drShiftData.Item("fPM")
                fPerDiem = drShiftData.Item("fPerDiem")
                fTravelTime = drShiftData.Item("fTravelTime")
                fState = drShiftData.Item("fState")
                fPW = drShiftData.Item("fPW")

                UserPerformedID = GetUserID(UserPerformed)

            End While

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString)
            LogError("LoadData.aspx.vb :: CreateShift", ex.ToString)
        Finally
            connShiftData.Close()
        End Try

        sQuery = "exec CreateShift @UserPerformed, @fType, @fShiftType, @fJobNumber, @fProjectName, @fPM, @fPerDiem, @fTravelTime, @fState, @fPW, NULL, NULL"
        Dim connNewShift As New SqlConnection(sConnection)
        Dim cmdNewShift As New SqlCommand(sQuery, connNewShift)
        Dim drNewShift As SqlDataReader
        cmdNewShift.Parameters.Add(New SqlParameter("@UserPerformed", UserPerformedID))
        cmdNewShift.Parameters.Add(New SqlParameter("@fType", fType))
        cmdNewShift.Parameters.Add(New SqlParameter("@fShiftType", fShiftType))
        cmdNewShift.Parameters.Add(New SqlParameter("@fJobNumber", fJobNumber))
        cmdNewShift.Parameters.Add(New SqlParameter("@fProjectName", fProjectName))
        cmdNewShift.Parameters.Add(New SqlParameter("@fPM", fPM))
        cmdNewShift.Parameters.Add(New SqlParameter("@fPerDiem", fPerDiem))
        cmdNewShift.Parameters.Add(New SqlParameter("@fTravelTime", fTravelTime))
        cmdNewShift.Parameters.Add(New SqlParameter("@fState", fState))
        cmdNewShift.Parameters.Add(New SqlParameter("@fPW", fPW))

        Try
            connNewShift.Open()
            drNewShift = cmdNewShift.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            While drNewShift.Read()
                ReturnedShiftID = drNewShift.Item(0)
            End While

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString)
            LogError("LoadData.aspx.vb :: CreateShift", ex.ToString)
        Finally
            connNewShift.Close()
        End Try
        Return ReturnedShiftID

    End Function

    Function GetUserID(ByVal UserPhone As String) As Integer

        Dim sUserID As String = "select EmployeeID from tEmployees where Cellphone=@UserPhone"
        Dim connUserID As New SqlConnection(sConnection)
        Dim cmdUserID As New SqlCommand(sUserID, connUserID)
        Dim drUserID As SqlDataReader
        cmdUserID.Parameters.Add(New SqlParameter("@UserPhone", UserPhone))

        Try
            connUserID.Open()
            drUserID = cmdUserID.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drUserID.Read
                Return drUserID.Item(0)
            End While
        Catch ex As Exception
            LogError("LoadData.aspx.vb :: GetUserID", ex.ToString)
            Response.Write("ERROR: " & ex.ToString)
        Finally
            connUserID.Close()
        End Try



    End Function

    Function FindLastShift(ByVal UserPerformed As String, ByVal ActionID As Integer) As Integer

        Dim ShiftID As Integer

        'Dim LastShiftQuery As String = "select top 1 shiftID from tShifts where UserPerformed = @UserPerformed and fInjured is null and fComments is null"
        Dim LastShiftQuery As String = "select top 1 tShifts.ShiftID from tShifts, tTimeEntry where tShifts.ShiftID = tTimeEntry.ShiftID and fInjured is null and fComments is null and UserPerformed=@UserPerformed order by tTimeEntry.EndTime DESC"
        'Above Find the most recent OpenShift based on Time Entry End Date
        Dim connLastShift As New SqlConnection(sConnection)
        Dim cmdLastShift As New SqlCommand(LastShiftQuery, connLastShift)
        Dim drLastShift As SqlDataReader
        cmdLastShift.Parameters.Add(New SqlParameter("@UserPerformed", UserPerformed))

        Try
            connLastShift.Open()
            drLastShift = cmdLastShift.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drLastShift.Read()
                ShiftID = drLastShift.Item("ShiftID")
            End While

        Catch ex As Exception
            LogError("LoadData.aspx.vb :: FindLastShift", ex.ToString)
            Response.Write("ERROR: " & ex.ToString)
        Finally
            connLastShift.Close()

        End Try

        If ShiftID = 0 Then ShiftID = CreateShift(ActionID)

        FindLastShift = ShiftID

    End Function

    Sub DeleteShift(ByVal ShiftID As Integer)

        Dim DeleteQuery As String = "delete from tShifts where shiftID=@ShiftID"
        Dim connDelete As New SqlConnection(sConnection)
        Dim cmdDelete As New SqlCommand(DeleteQuery, connDelete)
        cmdDelete.Parameters.Add(New SqlParameter("@ShiftID", ShiftID))

        Try
            connDelete.Open()
            cmdDelete.ExecuteNonQuery()
        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString)
            LogError("LoadData.aspx.vb :: DeleteShift", ex.ToString)
        Finally
            connDelete.Close()
        End Try

    End Sub

    Sub MarkAsProcessed(ByVal ActionID As Integer)

        Dim sQuery As String = "update tActions set Processed = 1 where ActionID=@ActionID"
        Dim connNewShift As New SqlConnection(sConnection)
        Dim cmdNewShift As New SqlCommand(sQuery, connNewShift)
        cmdNewShift.Parameters.Add(New SqlParameter("@ActionID", ActionID))

        Try
            connNewShift.Open()
            cmdNewShift.ExecuteNonQuery()
        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString)
            LogError("LoadData.aspx.vb :: MarkAsProcessed", ex.ToString)
        Finally
            connNewShift.Close()
        End Try

    End Sub

    Sub InsertTimeEntry(ByVal ShiftID As Integer, ByVal StartTime As DateTime, ByVal EndTime As DateTime, ByVal StartAddress As String, ByVal EndAddress As String, ByVal StartLat As String, ByVal StartLong As String, ByVal EndLat As String, ByVal EndLong As String)

        Dim sQuery As String = "insert into tTimeEntry (ShiftID,StartTime,EndTime,StartAddress,EndAddress, StartLatitude, StartLongitude, EndLatitude, EndLongitude, EnteredOn, EnteredManually) values(@ShiftID,@StartTime,@Endtime,@StartAddress,@EndAddress,@StartLatitude,@StartLongitude,@EndLatitude,@EndLongitude, @EnteredOn, @EnteredManually)"
        Dim connNewShift As New SqlConnection(sConnection)
        Dim cmdNewShift As New SqlCommand(sQuery, connNewShift)
        cmdNewShift.Parameters.Add(New SqlParameter("@ShiftID", ShiftID))
        cmdNewShift.Parameters.Add(New SqlParameter("@StartTime", StartTime))
        cmdNewShift.Parameters.Add(New SqlParameter("@EndTime", EndTime))
        cmdNewShift.Parameters.Add(New SqlParameter("@StartAddress", CleanAddress(StartAddress)))
        cmdNewShift.Parameters.Add(New SqlParameter("@EndAddress", CleanAddress(EndAddress)))
        cmdNewShift.Parameters.Add(New SqlParameter("@StartLatitude", ConvertCoords(StartLat)))
        cmdNewShift.Parameters.Add(New SqlParameter("@StartLongitude", ConvertCoords(StartLong)))
        cmdNewShift.Parameters.Add(New SqlParameter("@EndLatitude", ConvertCoords(EndLat)))
        cmdNewShift.Parameters.Add(New SqlParameter("@EndLongitude", ConvertCoords(EndLong)))
        cmdNewShift.Parameters.Add(New SqlParameter("@EnteredOn", Now()))
        cmdNewShift.Parameters.Add(New SqlParameter("@EnteredManually", 1))

        Try
            connNewShift.Open()
            cmdNewShift.ExecuteNonQuery()
            'Response.Write("START:" & StartAddress & "<br>END: " & EndAddress & "<br><br>")
        Catch ex As Exception
            'Response.Write("Start: " & StartAddress & "<Br>End: " & EndAddress & "<br>")
            Response.Write("ERROR: " & ex.ToString & "<br><br>")
            LogError("LoadData.aspx.vb :: InsertTimeEntry :: Shift=" & ShiftID & " :: StartTime=" & StartTime, ex.ToString)

        Finally
            connNewShift.Close()
        End Try
    End Sub

    Function ConvertCoords(ByVal RadianCoord As String) As Double

        If RadianCoord = "" Then
            Return 0
        Else
            Return Convert.ToDouble(RadianCoord) * (180 / Math.PI)
        End If

    End Function

    Sub EndShift(ByVal ActionID As Integer, ByVal ShiftID As Integer)

        Dim fInjured As String = "", fComments As String = ""

        Dim sQuery As String = "select fInjured, fComments from tActions where ActionID = @ActionID"
        Dim connShiftData As New SqlConnection(sConnection)
        Dim cmdShiftData As New SqlCommand(sQuery, connShiftData)
        Dim drShiftData As SqlDataReader
        cmdShiftData.Parameters.Add(New SqlParameter("@ActionID", ActionID))

        Try
            connShiftData.Open()
            drShiftData = cmdShiftData.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            While drShiftData.Read

                fInjured = drShiftData.Item("fInjured")
                fComments = drShiftData.Item("fComments")

            End While

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString)
            LogError("LoadData.aspx.vb :: EndShift", ex.ToString)
        Finally
            connShiftData.Close()
        End Try


        sQuery = "update tShifts set fInjured=@fInjured, fComments=@fComments where ShiftID=@ShiftID"
        Dim connUpdateShift As New SqlConnection(sConnection)
        Dim cmdUpdateShift As New SqlCommand(sQuery, connUpdateShift)
        cmdUpdateShift.Parameters.Add(New SqlParameter("@fInjured", fInjured))
        cmdUpdateShift.Parameters.Add(New SqlParameter("@fComments", fComments))
        cmdUpdateShift.Parameters.Add(New SqlParameter("@ShiftID", ShiftID))

        Try
            connUpdateShift.Open()
            cmdUpdateShift.ExecuteNonQuery()


        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString)
            LogError("LoadData.aspx.vb :: EndShift", ex.ToString)
        Finally
            connUpdateShift.Close()
        End Try

    End Sub

    Sub GetWorkerInfo(ByVal sender As Object, ByVal e As System.EventArgs)


        Dim wsWorkersConnection As New prodCompanyAdmin.CompanyAdminWebServiceProxyService


        Dim login As New System.Net.NetworkCredential
        login.UserName = "comnetws@95333"
        login.Password = "c0MmN35"

        wsWorkersConnection.Credentials = login

        Dim arrWorkers As Array
        Dim myWorker As SvWorker
        Dim SecFilter(0) As String
        SecFilter.SetValue("Mobile Worker", 0)
        Dim sResult As String = "<table border=1>"
        Dim Firstname As String = "", LastName As String = "", Username As String = ""

        arrWorkers = wsWorkersConnection.getWorkers(SecFilter)

        FlagAllAsDeleted()

        For Each myWorker In arrWorkers
            sResult = sResult & "<tr>"
            sResult = sResult & "<td>" & myWorker.firstName & "</td>"
            sResult = sResult & "<td>" & myWorker.lastName & "</td>"
            sResult = sResult & "<td>" & myWorker.userName & "</td>"
            sResult = sResult & "<td>" & myWorker.mobileNumber & "</td>"
            sResult = sResult & "</tr>"

            Firstname = myWorker.firstName
            LastName = myWorker.lastName
            Username = myWorker.userName

            AddWorker(Username, Firstname, LastName)



        Next

        sResult = sResult & "</table>"

        'LBL_test.Text = sResult

    End Sub

    Sub AddWorker(ByVal Username As String, ByVal Firstname As String, ByVal Lastname As String)

        Dim WorkerCount As Integer = 0
        Dim sqlCheckExisting As String = "select count(WorkerID) as WorkerCount from tWorkers where Username=@Username"
        Dim connCheckExisting As New SqlConnection(sConnection)
        Dim cmdCheckExisting As New SqlCommand(sqlCheckExisting, connCheckExisting)
        cmdCheckExisting.Parameters.Add(New SqlParameter("@Username", Username))
        Dim drCheckExisting As SqlDataReader

        Try
            connCheckExisting.Open()
            drCheckExisting = cmdCheckExisting.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            While drCheckExisting.Read()
                WorkerCount = drCheckExisting.Item("WorkerCount").ToString
            End While

        Catch ex As Exception
            LogError("LoadData.aspx.vb :: AddWorker", ex.ToString)
            Response.Write("ERROR: " & ex.ToString)
        Finally
            connCheckExisting.Close()
        End Try

        Dim sQuery As String = ""

        If WorkerCount = 0 Then
            sQuery = "insert into tWorkers (Username, FirstName, LastName, Deleted) values (@UserName, @FirstName, @LastName, 0)"
        Else
            sQuery = "update tWorkers set FirstName=@FirstName, LastName=@LastName, Deleted=0 where Username=@UserName"
        End If

        Dim connAddUser As New SqlConnection(sConnection)
        Dim cmdAddUser As New SqlCommand(sQuery, connAddUser)
        cmdAddUser.Parameters.Add(New SqlParameter("@Username", Username))
        cmdAddUser.Parameters.Add(New SqlParameter("@FirstName", Firstname))
        cmdAddUser.Parameters.Add(New SqlParameter("@LastName", Lastname))

        Try
            connAddUser.Open()
            cmdAddUser.ExecuteNonQuery()
        Catch ex As Exception
            LogError("LoadData.aspx.vb :: AddWorker", ex.ToString)
            Response.Write("ERROR: " & ex.ToString)
        Finally
            connAddUser.Close()
        End Try


    End Sub

    Sub FlagAllAsDeleted()

        Dim DeletedQuery As String = "update tWorkers set Deleted=1"
        Dim connDeleted As New SqlConnection(sConnection)
        Dim cmdDeleted As New SqlCommand(DeletedQuery, connDeleted)
        Try
            connDeleted.Open()
            cmdDeleted.ExecuteNonQuery()
        Catch ex As Exception
            LogError("LoadData.aspx.vb :: FlagAllAsDeleted", ex.ToString)
            Response.Write("ERROR:" & ex.ToString)
        Finally
            connDeleted.Close()
        End Try

    End Sub

    Public Sub LogSOAP(ByVal DateRangeStart As DateTime, ByVal DateRangeEnd As DateTime, ByVal ResponseText As String)

        Dim sLogError As String = "insert into tSOAPLog (RequestTime, DateRangeStart,DateRangeEnd,Response) values(@requestTime,@DateRangeStart,@DateRangeEnd,@Response)"
        Dim connLogError As New SqlConnection(sConnection)
        Dim cmdLogError As New SqlCommand(sLogError, connLogError)
        cmdLogError.Parameters.Add(New SqlParameter("@RequestTime", Now()))
        cmdLogError.Parameters.Add(New SqlParameter("@DateRangeStart", DateRangeStart))
        cmdLogError.Parameters.Add(New SqlParameter("@DateRangeEnd", DateRangeEnd))
        cmdLogError.Parameters.Add(New SqlParameter("@Response", ResponseText))

        Try
            connLogError.Open()
            cmdLogError.ExecuteNonQuery()
        Catch ex As Exception
            LogError("LoadData.aspx.vb :: LogSOAP", ex.ToString)
            Response.Write("::::::::::" & ex.ToString & ":::::::::::::")
        Finally
            connLogError.Close()
        End Try

    End Sub

    Function CleanAddress(ByVal OldAddress As String) As String
        If Not OldAddress = "" Then
            CleanAddress = Replace(OldAddress, ", TOWN OF", "")
        Else
            CleanAddress = OldAddress
        End If
    End Function

End Class
