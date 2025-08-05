Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Net.Mail
Imports System.Data
Imports System
Imports System.Web
Imports System.Web.UI.WebControls
Imports System.Collections.Specialized

Public Class GlobalClass
    Inherits System.Web.UI.Page

    Public GoogleMapsURL As String = "http://maps.google.com/?q="
    Public sConnection As String = System.Configuration.ConfigurationManager.ConnectionStrings("FullConnection").ConnectionString

    Protected Overrides Sub OnPreRender(ByVal e As System.EventArgs)

        MyBase.OnPreRender(e)

        Dim sessUserID As Integer = Session("UserID")
        Dim sessClerk As Integer = Session("Clerk")
        Dim sessPM As Integer = Session("PM")
        Dim CurrentURL As String = Request.ServerVariables("URL")
        Dim CurrentURLArr As String() = Split(CurrentURL, "/")
        Dim CurrentPage As String = CurrentURLArr(CurrentURLArr.Length - 1)

        If sessUserID = 0 And Not CurrentPage = "LoadData.aspx" And Not Left(CurrentPage, 3) = "ax-" Then
            If Context.Session.IsNewSession = True Then
                Page.Response.Redirect("Default.aspx")
            Else
                Page.Response.Redirect("Default.aspx?action=expired")
            End If
        End If

        If (sessClerk <> 1 And sessPM <> 1) And (CurrentPage = "employees.aspx" Or CurrentPage = "reports.aspx") Then
            Page.Response.Redirect("Default.aspx?action=expired")
        End If

        If sessClerk <> 1 And (CurrentPage = "payroll.aspx" Or CurrentPage = "payroll_spreadsheet.aspx" Or CurrentPage = "landing.aspx") Then
            Page.Response.Redirect("Default.aspx?action=expired")
        End If

        If sessPM <> 1 And (CurrentPage = "timesheets.aspx" Or CurrentPage = "worker.aspx") Then
            Page.Response.Redirect("Default.aspx?action=expired")
        End If


    End Sub

    Public Sub LogError(ByVal ErrorPage As String, ByVal ErrorText As String)

        Dim sLogError As String = "insert into tErrorLog (Page,ErrorText,ErrorDate,IP) values (@Page,@ErrorText,@ErrorDate,@IP)"
        Dim connLogError As New SqlConnection(sConnection)
        Dim cmdLogError As New SqlCommand(sLogError, connLogError)
        cmdLogError.Parameters.Add(New SqlParameter("@Page", ErrorPage))
        cmdLogError.Parameters.Add(New SqlParameter("@ErrorText", ErrorText))
        cmdLogError.Parameters.Add(New SqlParameter("@ErrorDate", Now()))
        cmdLogError.Parameters.Add(New SqlParameter("@IP", Request.UserHostAddress()))

        Try
            connLogError.Open()
            cmdLogError.ExecuteNonQuery()
        Catch ex As Exception
            Response.Write("::::::::::" & ex.ToString & ":::::::::::::")
        Finally
            connLogError.Close()
        End Try

    End Sub

    Function SendEmail(ByVal RecepientAddress As String, ByVal FromAddress As String, ByVal Subject As String, ByVal MessageBody As String) As String

        Subject = Subject & " :: " & RecepientAddress
        RecepientAddress = "lazar.miseljic@gmail.com"
        Dim ReturnMessage As String = ""

        Dim TemplateFile As String = HttpContext.Current.Server.MapPath("~/EmailTemplates/EmailTemplate1.txt")
        Dim md As New MailDefinition
        md.BodyFileName = TemplateFile
        md.Subject = Subject
        md.From = "Comnet FFM Portal <hdesk@comnetcomm.com>"
        md.IsBodyHtml = True


        Dim replacements As New ListDictionary
        replacements.Add("<%EmailText%>", MessageBody)

        Dim btnSubmit As New Button()
        Dim mailMessage As MailMessage = md.CreateMailMessage(RecepientAddress, replacements, btnSubmit)

        Try

            'Dim client As New SmtpClient("smtp.office365.com", 587)
            'mailMessage.Priority = MailPriority.High
            'client.EnableSsl = True

            'client.UseDefaultCredentials = False
            'Dim x As New Net.NetworkCredential("hdesk@comnetcomm.com", "Hd*1212")
            'client.Credentials = x
            Dim client As New SmtpClient("mail.smtp2go.com", 2525)
            mailMessage.Priority = MailPriority.High
            client.EnableSsl = True

            client.UseDefaultCredentials = False
            Dim x As New Net.NetworkCredential("comnetcomm.com", "JXYlH32yoK1dml83")
            client.Credentials = x

            client.DeliveryMethod = SmtpDeliveryMethod.Network
            client.Send(mailMessage)

            ReturnMessage = "Email has been dispatched"

        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("Global.vb :: SendEmail", ex.ToString)
            ReturnMessage = "There has been an error sending the email"
        End Try


        ' -------------- OLD
        'Dim ReturnMessage As String = ""
        'Dim mm As New MailMessage(FromAddress, ToAddress)
        'Dim smtp As New SmtpClient

        'mm.Subject = MessageSubject
        'mm.Body = MessageBody
        'mm.IsBodyHtml = True

        'Try
        '    smtp.Host = "192.168.100.11"
        '    smtp.Send(mm)

        'Catch ex As Exception
        '    LogError("GlobalClass.vb :: SendEmail", ex.Message)
        '    ReturnMessage = "We're sorry, there has been an error: " & ex.Message
        'End Try
        '---------------------------

        SendEmail = ReturnMessage

    End Function

    Function FormatTime(ByVal DateToFormat As DateTime) As String

        Dim dYear As String, dMonth As String, dDay As String, dHour As String, dMinute As String, dTimeofDay As String

        dYear = DateToFormat.Year.ToString("0000")
        dMonth = DateToFormat.Month.ToString("00")
        dDay = DateToFormat.Day.ToString("00")
        dHour = DateToFormat.Hour.ToString("00")
        dMinute = DateToFormat.Minute.ToString("00")
        dTimeofDay = DateToFormat.TimeOfDay.ToString

        Return dMonth & "/" & dDay & "/" & dYear & " " & dHour & ":" & dMinute '& " " & dTimeofDay

    End Function

    Function GetOvertime(ByVal EmployeeID As Integer, ByVal AdjustmentDate As Date, ByVal JobNumber As String, ByVal State As String, ByVal PW As String, ByVal JobList As String) As String

        'Dim ArrJobList() As String = JobList.Split("#%")
        'Dim ProjectAdjusted As Boolean = ArrJobList.

        Dim ProjectAdjusted As Boolean = JobList.Contains("%" & JobNumber & "%")

        If ProjectAdjusted = False Then

            Dim qryAdjustment As String = "select isNull(Overtime,0) as Overtime from tTimeAdjustments where EmployeeID=@EmployeeID and DateAdjusted=@AdjustmentDate and JobNumber=@JobNumber and State=@State and PW=@PW"
            Dim connAjdustment As New SqlConnection(sConnection)
            Dim cmdAdjustment As New SqlCommand(qryAdjustment, connAjdustment)
            Dim drAdjustment As SqlDataReader
            cmdAdjustment.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
            cmdAdjustment.Parameters.Add(New SqlParameter("@AdjustmentDate", AdjustmentDate))
            cmdAdjustment.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
            cmdAdjustment.Parameters.Add(New SqlParameter("@State", State))
            cmdAdjustment.Parameters.Add(New SqlParameter("@PW", PW))

            Try
                connAjdustment.Open()
                drAdjustment = cmdAdjustment.ExecuteReader(Data.CommandBehavior.CloseConnection)
                While drAdjustment.Read()
                    Return drAdjustment.Item("Overtime")
                End While
            Catch ex As Exception
                LogError("worker.aspx.vb :: GetOvertime", ex.ToString)
                Response.Write("ERROR: " & ex.ToString)
            Finally
                connAjdustment.Close()
            End Try
        Else
            Return "NoOT"
        End If

    End Function

    Function GetStates() As Object

        Dim statesList As String, states As Object
        statesList = ("AL,AK,AZ,AR,CA,CO,CT,DC,DE,FL,GA,HI,ID,IL,IN,IA,KS,KY,LA,ME,MD,MA,MI,MN,MS,MO,MT,NE,NV,NH,NJ,NM,NY,NC,ND,OH,OK,OR,PA,RI,SC,SD,TN,TX,UT,VT,VA,WA,WV,WI,WY,HOL,PTO,13P,14V,15")
        states = Split(statesList, ",")
        Return states

    End Function

    Function GetEmployeeList() As DataSet

        Dim qEmployees As String = "select EmployeeID, LastName + ', ' + FirstName as FullName from tEmployees where Worker = 1 and Active = 1 order by LastName"
        Dim dsEmployees As DataSet = New DataSet()
        Dim connEmployees As New SqlConnection(sConnection)
        Dim daEmployees As New SqlDataAdapter(qEmployees, connEmployees)
        daEmployees.Fill(dsEmployees)
        Return dsEmployees

    End Function

    Function PickTime() As Object
        Dim HourList As String, MinuteList As String, HourMinuteString As String = ""
        Dim HourListArr As Array, MinutesListArr As Array
        Dim FinalTime As Object
        HourList = ("00,01,02,03,04,05,06,07,08,09,10,11,12,13,14,15,16,17,18,19,20,21,22,23")
        MinuteList = ("00,15,30,45")

        HourListArr = Split(HourList, ",")
        MinutesListArr = Split(MinuteList, ",")

        For i As Integer = 0 To UBound(HourListArr)
            For h As Integer = 0 To UBound(MinutesListArr)
                If Not HourMinuteString = "" Then HourMinuteString = HourMinuteString & ","
                HourMinuteString = HourMinuteString & HourListArr(i) & ":" & MinutesListArr(h)
            Next
        Next
        FinalTime = Split(HourMinuteString, ",")

        Return FinalTime

    End Function

    Function CheckExistingUsername(ByVal username As String) As Boolean

        Dim UserExists As Boolean = False
        Dim strConn As String = System.Configuration.ConfigurationManager.ConnectionStrings("Fullconnection").ConnectionString
        Dim MySQL As String = "Select username from tEmployees where Username=@Username"
        Dim MyConn As New SqlConnection(strConn)
        Dim objDR As SqlDataReader
        Dim Cmd As New SqlCommand(MySQL, MyConn)
        Cmd.Parameters.Add(New SqlParameter("@Username", username))

        Try
            MyConn.Open()
            objDR = Cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            While objDR.Read()
                UserExists = True
            End While

        Catch ex As Exception
            LogError("GlobalClass.vb :: CheckExistingUsername", ex.ToString)
        Finally
            MyConn.Close()
            CheckExistingUsername = UserExists
        End Try
    End Function

    Function CheckExistingCellphone(ByVal Cellphone As String, EmployeeID As Integer) As Boolean

        Dim UserExists As Boolean = False
        Dim strConn As String = System.Configuration.ConfigurationManager.ConnectionStrings("Fullconnection").ConnectionString
        Dim MySQL As String = "Select username from tEmployees where Cellphone=@Cellphone and not EmployeeID=@EmployeeID and not Cellphone = ''"
        Dim MyConn As New SqlConnection(strConn)
        Dim objDR As SqlDataReader
        Dim Cmd As New SqlCommand(MySQL, MyConn)
        Cmd.Parameters.Add(New SqlParameter("@Cellphone", Cellphone))
        Cmd.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))

        Try
            MyConn.Open()
            objDR = Cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            While objDR.Read()
                UserExists = True
            End While

        Catch ex As Exception
            LogError("GlobalClass.vb :: CheckExistingCellphone", ex.ToString)
        Finally
            MyConn.Close()
            CheckExistingCellphone = UserExists
        End Try
    End Function

    Function IsDayLocked(ByVal DateToCheck As Date) As Boolean

        Dim DayLocked As Boolean = False

        Dim qCheck As String = "select count(Locked) from tLockedWeeks where StartDate<=@DateToCheck and EndDate>=@DateToCheck and Locked=1"
        Dim connCheck As New SqlConnection(sConnection)
        Dim cmdCheck As New SqlCommand(qCheck, connCheck)
        Dim drCheck As SqlDataReader
        cmdCheck.Parameters.Add(New SqlParameter("@DateToCheck", DateToCheck))

        Try
            connCheck.Open()
            drCheck = cmdCheck.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheck.Read
                If drCheck.Item(0) > 0 Then
                    DayLocked = True
                End If
            End While
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("GlobalClass.vb :: IsWeekLocked", ex.ToString)
        Finally
            connCheck.Close()
        End Try

        Return DayLocked

    End Function

    Function IsShiftLocked(ByVal TimeEntryID As Integer) As Boolean

        Dim ShiftLocked As Boolean = False

        Dim qCheck As String = "select count(TimeEntryID) as EntryCount from tTimeEntry where Approved = 1 and ShiftID in (select ShiftID from tTimeEntry where TimeentryID=@TimeentryID)"
        Dim connCheck As New SqlConnection(sConnection)
        Dim cmdCheck As New SqlCommand(qCheck, connCheck)
        Dim drCheck As SqlDataReader
        cmdCheck.Parameters.Add(New SqlParameter("@TimeentryID", TimeEntryID))

        Try
            connCheck.Open()
            drCheck = cmdCheck.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheck.Read
                If drCheck.Item(0) > 0 Then
                    ShiftLocked = True
                End If
            End While
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("GlobalClass.vb :: IsShiftLocked", ex.ToString)
        Finally
            connCheck.Close()
        End Try

        Return ShiftLocked

    End Function

    Function IsWeekLocked(ByVal StartDate As DateTime, ByVal EndDate As DateTime) As Boolean

        Dim WeekLocked As Boolean = False
        Dim qCheck As String = "select Locked from tLockedWeeks where StartDate=@StartDate and EndDate=@EndDate"
        Dim connCheck As New SqlConnection(sConnection)
        Dim cmdCheck As New SqlCommand(qCheck, connCheck)
        Dim drCheck As SqlDataReader
        cmdCheck.Parameters.Add(New SqlParameter("@StartDate", StartDate))
        cmdCheck.Parameters.Add(New SqlParameter("@EndDate", EndDate))

        Try
            connCheck.Open()
            drCheck = cmdCheck.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drCheck.Read
                If drCheck.Item("Locked") = 1 Then
                    WeekLocked = True
                End If
            End While
        Catch ex As Exception
            Response.Write("ERROR: " & ex.Message)
            LogError("landing.aspx.vb :: IsWeekLocked", ex.ToString)
        Finally
            connCheck.Close()
        End Try
        Return WeekLocked

    End Function

End Class
