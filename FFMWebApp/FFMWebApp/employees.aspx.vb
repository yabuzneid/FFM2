Imports System.Data.SqlClient

Partial Class employees
    Inherits GlobalClass


    Sub page_load()

        LBL_Info.Visible = False

        If Request.QueryString("id") <> 0 Then

            Panel_employeelist.Visible = False
            Panel_EditUser.Visible = True

            If IsPostBack = True Then

                'SaveEmployeeData()

            Else

                PopulateEmployeeForm(Request.QueryString("id"))

            End If

            If Request.QueryString("id") = -1 Then
                Panel_UsernameBox.Visible = True
                Panel_UsernameLabel.Visible = False
                LIT_Heading.Text = "Add User"
            Else
                Panel_UsernameBox.Visible = False
                Panel_UsernameLabel.Visible = True
            End If

        End If

        If Request.QueryString("action") = "saved" Then
            LBL_Info.Visible = True
            LBL_Info.Text = "Your changes have been saved."
        End If

    End Sub

    Sub PopulateEmployeeForm(ByVal EmployeeID As Integer)

        Dim Username As String = "", Password As String = "", Email As String = "", FirstName As String = "", LastName As String = "", Initials As String = "", CellPhone As String = "", PayrollID As String = "", Worker As Integer, PM As Integer, Clerk As Integer, Active As Integer = 1, PerDiem As Double = 4.5, OfficeID As Integer = 1, Type As String = "", TravelMultiplierCompany As Double = 0, TravelMultiplierPersonal As Double = 0, PartTime As Integer = 0, ReadOnlyAccess As Integer = 0, EmployeeGroup As String = "", LimitedView As Integer = 0, Department As String = ""

        Dim qGetUser As String = "select Username,Email, Password, FirstName, LastName, Initials, CellPhone, PayrollID, Worker, PM, Clerk, Active, PerDiem, OfficeID, Type, TravelMultiplierCompany, TravelMultiplierPersonal, PartTime, ReadOnlyAccess, EmployeeGroup, LimitedView, IsNull(Department,'') as Department from tEmployees where EmployeeID=@EmployeeID"
        Dim connGetUser As New SqlConnection(sConnection)
        Dim cmdGetUser As New SqlCommand(qGetUser, connGetUser)
        Dim drGetUser As SqlDataReader
        cmdGetUser.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))

        Try
            connGetUser.Open()
            drGetUser = cmdGetUser.ExecuteReader(Data.CommandBehavior.CloseConnection)
            While drGetUser.Read

                Username = drGetUser.Item("Username")
                Password = drGetUser.Item("Password")
                FirstName = drGetUser.Item("FirstName")
                LastName = drGetUser.Item("LastName")
                Email = If(IsDBNull(drGetUser.Item("Email")), "", drGetUser.Item("Email").ToString())
                Initials = drGetUser.Item("Initials")
                CellPhone = drGetUser.Item("CellPhone")
                PayrollID = drGetUser.Item("PayrollID")
                Worker = drGetUser.Item("Worker")
                PM = drGetUser.Item("PM")
                Clerk = drGetUser.Item("Clerk")
                Active = drGetUser.Item("Active")
                PerDiem = drGetUser.Item("PerDiem")
                OfficeID = drGetUser.Item("OfficeID")
                Type = drGetUser.Item("Type")
                TravelMultiplierCompany = drGetUser.Item("TravelMultiplierCompany")
                TravelMultiplierPersonal = drGetUser.Item("TravelMultiplierPersonal")
                PartTime = drGetUser.Item("PartTime")
                ReadOnlyAccess = drGetUser.Item("ReadOnlyAccess")
                EmployeeGroup = drGetUser.Item("EmployeeGroup")
                LimitedView = drGetUser.Item("LimitedView")
                Department = drGetUser.Item("Department")

            End While
        Catch ex As Exception
            LogError("Employees.aspx.vb :: GenerateUserForm", ex.ToString)
            Response.Write("ERROR: " & ex.ToString)
        Finally
            connGetUser.Close()
        End Try

        LIT_Username.Text = Username
        TXT_firstname.Text = FirstName
        TXT_LastName.Text = LastName
        TXT_Email.Text = Email
        TXT_Initials.Text = Initials
        TXT_Cellphone.Text = CellPhone
        If Worker = 1 Then CHK_Worker.Checked = True
        If PartTime = 1 Then CHK_PartTime.Checked = True
        If PM = 1 Then CHK_PM.Checked = True
        If Clerk = 1 Then CHK_Clerk.Checked = True
        If Active = 1 Then CHK_Active.Checked = True
        If ReadOnlyAccess = 1 Then CHK_ReadOnlyAccess.Checked = True
        If LimitedView = 1 Then CHK_LimitedView.Checked = True
        TXT_PerDiem.Text = PerDiem
        TXT_PayrollID.Text = PayrollID
        DDL_Office.SelectedValue = OfficeID
        DDL_Type.SelectedValue = Type
        DDL_EmployeeGroup.SelectedValue = EmployeeGroup
        DDL_Department.SelectedValue = Department

        TXT_TravelCompany.Text = TravelMultiplierCompany
        TXT_TravelPersonal.Text = TravelMultiplierPersonal

    End Sub

    Sub SaveEmployeeData(ByVal sender As Object, ByVal e As EventArgs)

        Dim EmployeeID As Integer = Request.QueryString("id")
        Dim Redirect As Boolean = False

        If (EmployeeID = -1 And (CheckExistingUsername(TXT_Username.Text) = True Or TXT_Username.Text = "")) Or CheckExistingCellphone(TXT_Cellphone.Text, EmployeeID) = True Then
            LBL_Info.Visible = True

            If CheckExistingCellphone(TXT_Cellphone.Text, EmployeeID) = True Then
                LBL_Info.Text = "Cellphone number is already assigned to a user. Please remove the number from that employee's profile first."
            Else
                If TXT_Username.Text = "" Then
                    LBL_Info.Text = "Username cannot be blank. Please enter a valid username."
                Else
                    LBL_Info.Text = "This Username is already in use. Please choose a distinct username."
                End If
            End If

        Else

            Dim Username As String, Email As String, Password As String = "", FirstName As String, Lastname As String, Initials As String, CellPhone As String, PerDiem As Double, PayrollID As String, OfficeID As Integer, Type As String, TravelMultiplierCompany As Double = 0, TravelMultiplierPersonal As Double = 0, PartTime As Integer = 0, ReadOnlyAccess As Integer = 0
            Dim Active As Integer = 0, Worker As Integer = 0, PM As Integer = 0, Clerk As Integer = 0, sPerDiem As String, EmployeeGroup As String, LimitedView As Integer = 0, Department As String = ""

            EmployeeID = Request.QueryString("id")

            Username = TXT_Username.Text
            Password = TXT_password.Text
            Email = TXT_Email.Text
            FirstName = TXT_firstname.Text
            Lastname = TXT_LastName.Text
            Initials = TXT_Initials.Text
            CellPhone = TXT_Cellphone.Text
            sPerDiem = TXT_PerDiem.Text
            PayrollID = TXT_PayrollID.Text
            If CHK_Active.Checked = True Then Active = 1
            If CHK_Worker.Checked = True Then Worker = 1
            If CHK_PartTime.Checked = True Then PartTime = 1
            If CHK_PM.Checked = True Then PM = 1
            If CHK_Clerk.Checked = True Then Clerk = 1
            If CHK_ReadOnlyAccess.Checked = True Then ReadOnlyAccess = 1
            If CHK_LimitedView.Checked = True Then LimitedView = 1
            OfficeID = DDL_Office.SelectedValue
            Type = DDL_Type.SelectedValue
            EmployeeGroup = DDL_EmployeeGroup.SelectedValue
            Department = DDL_Department.SelectedValue

            TravelMultiplierCompany = 0 'TXT_TravelCompany.Text
            TravelMultiplierPersonal = 0 'TXT_TravelPersonal.Text

            sPerDiem = System.Text.RegularExpressions.Regex.Replace(sPerDiem, "[^0-9\.]", "")
            If sPerDiem = "" Then sPerDiem = "0"
            PerDiem = CDbl(sPerDiem)

            CellPhone = System.Text.RegularExpressions.Regex.Replace(CellPhone, "[^0-9]", "")
            'PerDiem = System.Text.RegularExpressions.Regex.Replace(PerDiem, "[^0-9\.]", "")

            Dim qEmployee As String

            If EmployeeID = -1 Then
                qEmployee = "insert into tEmployees (Username,Email,FirstName,LastName,Initials,Password,CellPhone,Worker,PM,Clerk,Active,PerDiem,PayrollID,OfficeID,Type, TravelMultiplierCompany, TravelMultiplierPersonal, PartTime, ReadOnlyAccess, EmployeeGroup, LimitedView, Department) values (@Username,@Email,@FirstName,@LastName,@Initials,@Password,@CellPhone,@Worker,@PM,@Clerk,@Active,@PerDiem,@PayrollID,@OfficeID,@Type, @TravelMultiplierCompany, @TravelMultiplierPersonal, @PartTime, @ReadOnlyAccess, @EmployeeGroup, @LimitedView, @Department)"
            Else
                If Password = "" Then
                    qEmployee = "update tEmployees set Email=@Email,FirstName=@FirstName,LastName=@LastName,Initials=@Initials,Cellphone=@Cellphone,Worker=@Worker,PM=@PM,Clerk=@Clerk,Active=@Active,PerDiem=@PerDiem,PayrollID=@PayrollID, OfficeID=@OfficeID, Type=@Type, TravelMultiplierCompany=@TravelMultiplierCompany, TravelMultiplierPersonal=@TravelMultiplierPersonal, PartTime=@PartTime, ReadOnlyAccess=@ReadOnlyAccess, EmployeeGroup=@EmployeeGroup, LimitedView=@LimitedView, Department=@Department where EmployeeID=@EmployeeID"
                Else
                    qEmployee = "update tEmployees set Email=@Email, Password=@Password,FirstName=@FirstName,LastName=@LastName,Initials=@Initials,Cellphone=@Cellphone,Worker=@Worker,PM=@PM,Clerk=@Clerk,Active=@Active,PerDiem=@PerDiem,PayrollID=@PayrollID, OfficeID=@OfficeID, Type=@Type, TravelMultiplierCompany=@TravelMultiplierCompany, TravelMultiplierPersonal=@TravelMultiplierPersonal, PartTime=@PartTime, ReadOnlyAccess=@ReadOnlyAccess, EmployeeGroup=@EmployeeGroup, LimitedView=@LimitedView, Department=@Department where EmployeeID=@EmployeeID"
                End If
            End If

            Dim connEmployee As New SqlConnection(sConnection)
            Dim cmdEmployee As New SqlCommand(qEmployee, connEmployee)

            cmdEmployee.Parameters.Add(New SqlParameter("@Email", Email))
            cmdEmployee.Parameters.Add(New SqlParameter("@FirstName", FirstName))
            cmdEmployee.Parameters.Add(New SqlParameter("@LastName", Lastname))
            cmdEmployee.Parameters.Add(New SqlParameter("@Initials", Initials))
            cmdEmployee.Parameters.Add(New SqlParameter("@CellPhone", CellPhone))
            cmdEmployee.Parameters.Add(New SqlParameter("@Worker", Worker))
            cmdEmployee.Parameters.Add(New SqlParameter("@PM", PM))
            cmdEmployee.Parameters.Add(New SqlParameter("@Clerk", Clerk))
            cmdEmployee.Parameters.Add(New SqlParameter("@Active", Active))
            cmdEmployee.Parameters.Add(New SqlParameter("@PerDiem", PerDiem))
            cmdEmployee.Parameters.Add(New SqlParameter("@PayrollID", PayrollID))
            cmdEmployee.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
            cmdEmployee.Parameters.Add(New SqlParameter("@OfficeID", OfficeID))
            cmdEmployee.Parameters.Add(New SqlParameter("@Type", Type))
            cmdEmployee.Parameters.Add(New SqlParameter("@TravelMultiplierCompany", TravelMultiplierCompany))
            cmdEmployee.Parameters.Add(New SqlParameter("@TravelMultiplierPersonal", TravelMultiplierPersonal))
            cmdEmployee.Parameters.Add(New SqlParameter("@PartTime", PartTime))
            cmdEmployee.Parameters.Add(New SqlParameter("@ReadOnlyAccess", ReadOnlyAccess))
            cmdEmployee.Parameters.Add(New SqlParameter("@EmployeeGroup", EmployeeGroup))
            cmdEmployee.Parameters.Add(New SqlParameter("@LimitedView", LimitedView))
            cmdEmployee.Parameters.Add(New SqlParameter("@Department", Department))


            If EmployeeID = -1 Then
                cmdEmployee.Parameters.Add(New SqlParameter("@Username", Username))
                cmdEmployee.Parameters.Add(New SqlParameter("@Password", Password))
            Else
                If Not Password = "" Then
                    cmdEmployee.Parameters.Add(New SqlParameter("@Password", Password))
                End If
            End If

            Try
                connEmployee.Open()

                cmdEmployee.ExecuteNonQuery()


                Redirect = True

            Catch ex As Exception
                LogError("employees.aspx.vb :: SaveEmployeeData", ex.ToString)
                Response.Write("ERROR: " & ex.ToString)
            Finally
                connEmployee.Close()
            End Try

            If Redirect = True Then Response.Redirect("employees.aspx?action=saved")

        End If

    End Sub

    Sub CancelForm(ByVal sender As Object, ByVal e As EventArgs)

        Response.Redirect("employees.aspx")

    End Sub

    Function ConvertCheck(ByVal OrigValue As Integer) As String
        If OrigValue = 1 Then
            Return "<img src=""images/icon-checked-16.png"" />"
        Else
            Return ""
        End If
    End Function

End Class
