Imports System.Data.SqlClient


Partial Class ax_edittimesheet
    Inherits GlobalClass


    Sub page_load()

        If Not Session("UserID") > 0 Then
            Response.Write("%EXPIRED%")
        Else

            Dim ResponseString As String = ""

            If Request.QueryString("action") = "approve" Then

                Dim TimeEntryID As Integer = Request.QueryString("id")

                Response.Write(ApproveTimeEntry(TimeEntryID, 1))

            ElseIf Request.QueryString("action") = "unapprove" Then

                Dim TimeEntryID As Integer = Request.QueryString("id")

                Response.Write(ApproveTimeEntry(TimeEntryID, 0))

            ElseIf Request.QueryString("action") = "savenotes" Then

                Dim EmployeeID As Integer = Request.Form("employeeid")
                Dim WeekStart As Date = Request.Form("weekstart")
                Dim NotesText As String = Request.Form("value")

                Response.Write(SaveNotes(EmployeeID, WeekStart, NotesText))

            ElseIf Request.QueryString("action") = "showjobnumberform" Then

                Dim TimeEntryID As Integer = request.form("timeentryid")
                Dim JobNumber As String = request.form("jobnumber")

                Response.write(ShowJobNumberForm(JobNumber, TimeEntryID))

            ElseIf Request.QueryString("action") = "savejobnumberform" Then

                Dim TimeEntryID As Integer = request.form("timeentryid")
                Dim JobNum1 As String = request.form("jobnum1")
                Dim JobNum2 As String = request.form("jobnum2")
                Dim JobNum3 As String = request.form("jobnum3")
                Dim JobNum4 As String = request.form("jobnum4")

                Response.write(SaveJobNumberForm(TimeEntryID, JobNum1, JobNum2, JobNum3, JobNum4))

            Else


                Dim ElemetID As String = Request.Form("id")
                Dim ElementValue As String = Request.Form("value")

                Dim ArrElementID As Array = ElemetID.Split("_")

                Dim FieldToEdit As String = ArrElementID(0)
                Dim TimeEntryID As String = ArrElementID(1)

                Dim UpdateCommand As String = "", selectCommand As String = ""

                If FieldToEdit = "starttime" Then
                    UpdateCommand = "update tTimeEntry set StartTime=@NewValue where TimeEntryID=@TimeEntryID"
                    selectCommand = "select StartTime from tTimeEntry where TimeEntryID=@TimeEntryID"
                    ResponseString = UpdateTimeEntry(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "endtime" Then
                    UpdateCommand = "update tTimeEntry set EndTime=@NewValue where TimeEntryID=@TimeEntryID"
                    selectCommand = "select EndTime from tTimeEntry where TimeEntryID=@TimeEntryID"
                    ResponseString = UpdateTimeEntry(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "pm" Then
                    UpdateCommand = "update tShifts set fPM=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fPM from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "state" Then
                    UpdateCommand = "update tShifts set fState=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fState from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "jobnumber" Then
                    UpdateCommand = "update tShifts set fJobNumber=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fJobNumber from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "projectname" Then
                    UpdateCommand = "update tShifts set fProjectName=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fProjectName from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "type" Then
                    UpdateCommand = "update tShifts set fType=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fType from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "shifttype" Then
                    UpdateCommand = "update tShifts set fShiftType=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fShiftType from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "travel" Then
                    UpdateCommand = "update tShifts set fTravelTime=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fTravelTime from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "type" Then
                    UpdateCommand = "update tShifts set fType=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fType from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "perdiem" Then
                    UpdateCommand = "update tShifts set fPerDiem=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fPerDiem from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "injured" Then
                    UpdateCommand = "update tShifts set fInjured=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fInjured from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "comments" Then
                    UpdateCommand = "update tShifts set fComments=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fComments from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                ElseIf FieldToEdit = "pw" Then
                    UpdateCommand = "update tShifts set fPW=@NewValue where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    selectCommand = "select fPW from tShifts where shiftID in (select shiftID from tTimeEntry where TimeEntryID = @TimeEntryID)"
                    ResponseString = UpdateShift(UpdateCommand, selectCommand, TimeEntryID, ElementValue)
                End If

            End If
            Response.Write(ResponseString)
        End If
    End Sub

    Function UpdateTimeEntry(ByVal UpdateCommand As String, ByVal SelectCommand As String, ByVal timeEntryID As Integer, ByVal NewValue As String) As String

        Dim returnstring As String = ""

        Dim connUpdate As New SqlConnection(sConnection)
        Dim cmdUpdate As New SqlCommand(UpdateCommand, connUpdate)
        cmdUpdate.Parameters.Add(New SqlParameter("@TimeEntryID", timeEntryID))
        cmdUpdate.Parameters.Add(New SqlParameter("@NewValue", NewValue))

        Try
            connUpdate.Open()
            cmdUpdate.ExecuteNonQuery()

            Dim connSelect As New SqlConnection(sConnection)
            Dim cmdSelect As New SqlCommand(SelectCommand, connSelect)
            Dim drSelect As SqlDataReader
            cmdSelect.Parameters.Add(New SqlParameter("@TimeEntryID", timeEntryID))

            Try
                connSelect.Open()
                drSelect = cmdSelect.ExecuteReader(Data.CommandBehavior.CloseConnection)

                While drSelect.Read
                    returnstring = drSelect.Item(0)
                End While

            Catch ex As Exception
                LogError("ax-edittimesheet.aspx.vb :: UpdateTimeEntry", ex.ToString)
                returnstring = "ERROR: " & ex.ToString
            Finally
                connSelect.Close()
            End Try

        Catch ex As Exception
            LogError("ax-edittimesheet.aspx.vb :: UpdateTimeEntry", ex.ToString)
            returnstring = "ERROR: The system could not save the value you entered. Please make sure that start and end times are in the correct Date and Time format."
        Finally
            connUpdate.Close()
        End Try

        Return returnstring
    End Function

    Function UpdateShift(ByVal UpdateCommand As String, ByVal selectCommand As String, ByVal TimeEntryID As String, ByVal NewValue As String) As String

        Dim returnstring As String = ""

        Dim connUpdate As New SqlConnection(sConnection)
        Dim cmdUpdate As New SqlCommand(UpdateCommand, connUpdate)
        cmdUpdate.Parameters.Add(New SqlParameter("@TimeEntryID", TimeEntryID))
        cmdUpdate.Parameters.Add(New SqlParameter("@NewValue", NewValue))

        Try
            connUpdate.Open()
            cmdUpdate.ExecuteNonQuery()

            Dim connSelect As New SqlConnection(sConnection)
            Dim cmdSelect As New SqlCommand(selectCommand, connSelect)
            Dim drSelect As SqlDataReader
            cmdSelect.Parameters.Add(New SqlParameter("@TimeEntryID", TimeEntryID))

            Try
                connSelect.Open()
                drSelect = cmdSelect.ExecuteReader(Data.CommandBehavior.CloseConnection)

                While drSelect.Read
                    returnstring = drSelect.Item(0)
                End While

            Catch ex As Exception
                LogError("ax-edittimesheet.aspx.vb :: UpdatePerDiem", ex.ToString)
                returnstring = "ERROR: " & ex.ToString
            Finally
                connSelect.Close()
            End Try

        Catch ex As Exception
            LogError("ax-edittimesheet.aspx.vb :: UpdatePerDiem", ex.ToString)
            returnstring = "ERROR: " & ex.ToString
        Finally
            connUpdate.Close()
        End Try

        Return returnstring

    End Function

    Function SaveJobNumberForm(TimeEntryID As Integer, JobNum1 As String, JobNum2 As String, JobNum3 As String, JobNum4 As String) As String

        Dim returnstring As String = ""

        Dim FullJobNumber As String = JobNum1 & "-" & JobNum2 & "-" & JobNum3

        If len(JobNum4) > 0 Then
            FullJobNumber = FullJobNumber & "-" & JobNum4
        End If

        Dim UpdateCommand As String = "update tShifts set fJobNumber=@JobNumber where ShiftID in (select ShiftID from tTimeEntry where TimeEntryID=@TimeEntryID)"
        Dim connUpdate As New SqlConnection(sConnection)
        Dim cmdUpdate As New SqlCommand(UpdateCommand, connUpdate)
        cmdUpdate.Parameters.Add(New SqlParameter("@TimeEntryID", TimeEntryID))
        cmdUpdate.Parameters.Add(New SqlParameter("@JobNumber", FullJobNumber))

        Try
            connUpdate.Open()
            cmdUpdate.ExecuteNonQuery()

            returnstring = FullJobNumber

        Catch ex As Exception
            LogError("ax-edittimesheet.aspx.vb :: SaveJobNumberForm", ex.ToString)
            returnstring = "ERROR: " & ex.ToString
        Finally
            connUpdate.Close()
        End Try

        Return returnstring

    End Function

    Function ShowJobNumberForm(JobNumber As String, TimeEntryID As Integer) As String

        Dim returnstring As String = ""
        Dim JobNum1 As String = "01", JobNum2 As String = "", JobNum3 As String = "", JobNum4 As String = ""
        Dim IsItValid As Boolean = True

        'FIGURE OUT IF NUMBER IS VALID
        JobNumber = Regex.Replace(JobNumber, "[^0-9]", "")

        If Not len(JobNumber) = 10 And Not len(JobNumber) = 13 Then
            IsItValid = False
        Else
            If Not JobNumber.substring(0, 2) = "01" Then
                IsItValid = False
            Else
                If Not JobNumber.Substring(2, 2) = "01" And Not JobNumber.Substring(2, 2) = "02" And Not JobNumber.Substring(2, 2) = "03" And Not JobNumber.Substring(2, 2) = "04" And Not JobNumber.Substring(2, 2) = "12" And Not JobNumber.Substring(2, 2) = "14" And Not JobNumber.Substring(2, 2) = "15" Then
                    IsItValid = False
                End If
            End If
        End If

        If IsItValid = True Then
            JobNum1 = JobNumber.substring(0, 2)
            JobNum2 = JobNumber.substring(2, 2)
            JobNum3 = JobNumber.substring(4, 6)
            If len(JobNumber) = 13 Then
                JobNum4 = JobNumber.substring(10, 3)
            End If
        End If

        'CREATE FORM
        returnstring = JobNumber
        returnstring = returnstring & "<form name=""editjobnumber_" & TimeEntryID & """ class=""editjobnumber"" action=""ax-edittimesheet.aspx?action=savejobnumberform"">"
        returnstring = returnstring & "<input type=""text"" name=""jobnum_1_" & TimeEntryID & """ id=""jobnum_1_" & TimeEntryID & """ value=""" & JobNum1 & """ maxlength=""2"" size=""2"" />-"

        returnstring = returnstring & "<select name=""jobnum_2_" & TimeEntryID & """ id=""jobnum_2_" & TimeEntryID & """>"
        returnstring = returnstring & "<option value=""""></option>"
        returnstring = returnstring & "<option value=""01""" & IIf(JobNum2 = "01", " selected", " ") & ">01</option>"
        returnstring = returnstring & "<option value=""02""" & IIf(JobNum2 = "02", " selected", " ") & ">02</option>"
        returnstring = returnstring & "<option value=""03""" & IIf(JobNum2 = "03", " selected", " ") & ">03</option>"
        returnstring = returnstring & "<option value=""04""" & IIf(JobNum2 = "04", " selected", " ") & ">04</option>"
        returnstring = returnstring & "<option value=""12""" & IIf(JobNum2 = "12", " selected", " ") & ">12</option>"
        returnstring = returnstring & "<option value=""14""" & IIf(JobNum2 = "14", " selected", " ") & ">14</option>"
        returnstring = returnstring & "<option value=""15""" & IIf(JobNum2 = "15", " selected", " ") & ">15</option>"
        returnstring = returnstring & "</select>-"

        returnstring = returnstring & "<input type=""text"" name=""jobnum_3_" & TimeEntryID & """ id=""jobnum_3_" & TimeEntryID & """ value=""" & JobNum3 & """ maxlength=""6"" size=""6"" />-"
        returnstring = returnstring & "<input type=""text"" name=""jobnum_4_" & TimeEntryID & """ id=""jobnum_4_" & TimeEntryID & """ value=""" & JobNum4 & """ maxlength=""4"" size=""4"" />"
        returnstring = returnstring & "<input type=""button"" class=""jobnum_submit"" name=""jobnum_submit_" & TimeEntryID & """ id=""jobnum_submit_" & TimeEntryID & """ value=""OK"" />"
        returnstring = returnstring & "<input type=""button"" class=""jobnum_cancel"" name=""jobnum_cancel_" & TimeEntryID & """ id=""jobnum_cancel_" & TimeEntryID & """ value=""Cancel"" />"
        returnstring = returnstring & "</form>"

        Return returnstring

    End Function

    Function ApproveTimeEntry(ByVal TimeEntryID As Integer, ByVal NewValue As Integer) As String

        Dim ApproveQuery As String = "update tTimeEntry set Approved = @NewValue, ApprovedBy=@PM, ApprovedOn=@Now where TimeEntryID=@TimeEntryID"
        Dim connApprove As New SqlConnection(sConnection)
        Dim cmdApprove As New SqlCommand(ApproveQuery, connApprove)
        cmdApprove.Parameters.Add(New SqlParameter("@TimeEntryID", TimeEntryID))
        cmdApprove.Parameters.Add(New SqlParameter("@PM", Session("UserID")))
        cmdApprove.Parameters.Add(New SqlParameter("@Now", Now()))
        cmdApprove.Parameters.Add(New SqlParameter("@NewValue", NewValue))
        Try
            connApprove.Open()
            cmdApprove.ExecuteNonQuery()
            'CheckApprovalFinished()
            If NewValue = 1 Then
                Return "<img src=""images/tick.png""><br><span class=""unapprove"">[<a href=""#"" class=""unapprove"" id=""unapprove_" & TimeEntryID & """>unapprove</a>]</span>"
            Else
                Return "<input type=""button"" class=""approvebutton"" id=""approve_" & TimeEntryID & """ value=""Approve"">"
            End If

        Catch ex As Exception
            LogError("ax-edittimesheet.aspx.vb :: ApproveTimeQuery", ex.ToString)
            Return "ERROR: " & ex.ToString
        Finally
            connApprove.Close()
        End Try

    End Function

    Sub CheckApprovalFinished()

        Dim DaysToStartOfWeek As Integer = Now().Date.DayOfWeek + 2

        If DaysToStartOfWeek = 0 Then DaysToStartOfWeek = 7
        DaysToStartOfWeek = DaysToStartOfWeek

        Dim EndLastWeek As Date = DateAdd(DateInterval.Day, -DaysToStartOfWeek, Now())

        Dim WeekStart As Date = EndLastWeek.AddDays(-6)
        Dim WeekEnd As Date = EndLastWeek

        Dim sTimeQuery As String = "select count(TimeEntryID) as CountUnapproved from tTimeEntry, tShifts where Approved <> 1 and tShifts.ShiftID = tTimeEntry.ShiftID and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime"
        Dim connTimeQuery As New SqlConnection(sConnection)
        Dim cmdTimeQuery As New SqlCommand(sTimeQuery, connTimeQuery)
        cmdTimeQuery.Parameters.Add(New SqlParameter("@StartTime", WeekStart))
        cmdTimeQuery.Parameters.Add(New SqlParameter("@EndTime", WeekEnd))
        Dim drTimeQuery As SqlDataReader

        Try
            connTimeQuery.Open()
            drTimeQuery = cmdTimeQuery.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drTimeQuery.Read()
                If Not drTimeQuery.Item("CountUnapproved") > 0 Then

                    Dim msgSubject As String = "Time approved notification"
                    Dim msgBody As String = "All time for the week ending " & FormatDateTime(WeekEnd, DateFormat.LongDate) & " has been approved by Project Managers and is ready to process into payroll."
                    Dim EmailResponse As String = SendEmail("lmiseljic@cispectrum.com", "donotreply@comnetcomm.com", msgSubject, msgBody)

                End If
            End While

        Catch ex As Exception
            LogError("ax-edittimesheet.aspx.vb :: CheckApprovalFinished", ex.ToString)
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connTimeQuery.Close()
        End Try

    End Sub

    Function SaveNotes(ByVal EmployeeID As Integer, ByVal DateStart As Date, ByVal NotesText As String) As String

        Dim NotesCount As Integer = 0, ErrorString As String = ""
        Dim qCheckNotes As String = "select count(EmployeeID) as NotesCount from tNotes where EmployeeID=@EmployeeID and WeekStartDate=@DateStart"
        Dim ConnCheckNotes As New SqlConnection(sConnection)
        Dim CmdCheckNotes As New SqlCommand(qCheckNotes, ConnCheckNotes)
        Dim drCheckNotes As SqlDataReader

        CmdCheckNotes.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
        CmdCheckNotes.Parameters.Add(New SqlParameter("@DateStart", DateStart))

        Try
            ConnCheckNotes.Open()
            drCheckNotes = CmdCheckNotes.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheckNotes.Read()
                NotesCount = drCheckNotes.Item("NotesCount")
            End While
        Catch ex As Exception
            LogError("ax-edittimesheet.aspx.vb :: SaveNotes", ex.ToString)
            ErrorString = ex.ToString
        Finally
            ConnCheckNotes.Close()
        End Try


        Dim qSaveNotes As String

        If NotesCount > 0 Then
            qSaveNotes = "update tNotes set NotesText=@NotesText where EmployeeID=@EmployeeID and WeekStartDate=@DateStart"
        Else
            qSaveNotes = "insert into tNotes (NotesText,EmployeeID,WeekStartDate) values (@NotesText,@EmployeeID,@DateStart)"
        End If

        Dim connSaveNotes As New SqlConnection(sConnection)
        Dim cmdSaveNotes As New SqlCommand(qSaveNotes, connSaveNotes)
        cmdSaveNotes.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
        cmdSaveNotes.Parameters.Add(New SqlParameter("@DateStart", DateStart))
        cmdSaveNotes.Parameters.Add(New SqlParameter("@NotesText", NotesText))

        Try
            connSaveNotes.Open()
            cmdSaveNotes.ExecuteNonQuery()

        Catch ex As Exception
            LogError("ax-edittimesheet.aspx.vb :: SaveNotes", ex.ToString)
            ErrorString = ex.ToString
        Finally
            connSaveNotes.Close()
        End Try


        If ErrorString = "" Then
            Return "SUCCESS"
        Else
            Return ErrorString
        End If


    End Function

End Class
