Imports System.Data.SqlClient


Partial Class ax_editpayroll
    Inherits GlobalClass


    Sub page_load()

        If Not Session("UserID") > 0 Then
            Response.Write("%EXPIRED%")
        Else
            Dim ElemetID As String = Request.Form("id")
            Dim ElementValue As String = Request.Form("value")

            Dim ArrElementID As Array = ElemetID.Split("_")

            Dim FieldType As String = ArrElementID(0)
            Dim EmployeeID As String = ArrElementID(1)
            Dim JobNumber As String = ArrElementID(2)
            Dim DateAdjusted As String = ArrElementID(3)
            Dim PW As String = ArrElementID(4)
            Dim State As String = ArrElementID(5)
            Dim Travel As String = ""

            Dim sUpdateQuery As String = "", sInsertQuery As String = "", sSelectQuery As String = "", sCheckQuery As String = ""

            If FieldType = "st" Then

                sUpdateQuery = "update tTimeAdjustments set StandardTime=@Value, LastModifiedBy=@LastModifiedBy, LastModifiedOn=@LastModifiedOn where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"
                sInsertQuery = "insert into tTimeAdjustments (DateAdjusted,EmployeeID,JobNumber,StandardTime, LastModifiedBy, LastModifiedOn, PW, State) values (@DateAdjusted,@EmployeeID,@JobNumber,@Value, @LastModifiedBy, @LastModifiedOn,@PW,@State)"
                sSelectQuery = "select Standardtime from tTimeAdjustments where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"
                sCheckQuery = "select count(EmployeeID) as RecordCount from tTimeAdjustments where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"

            ElseIf FieldType = "ot" Then

                sUpdateQuery = "update tTimeAdjustments set Overtime=@Value, LastModifiedBy=@LastModifiedBy, LastModifiedOn=@LastModifiedOn where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"
                sInsertQuery = "insert into tTimeAdjustments (DateAdjusted,EmployeeID,JobNumber,Overtime, LastModifiedBy, LastModifiedOn, PW, State) values (@DateAdjusted,@EmployeeID,@JobNumber,@Value, @LastModifiedBy, @LastModifiedOn,@PW,@State)"
                sSelectQuery = "select Overtime from tTimeAdjustments where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"
                sCheckQuery = "select count(EmployeeID) as RecordCount from tTimeAdjustments where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"

            ElseIf FieldType = "c" Then

                sUpdateQuery = "update tTimeAdjustments set Company=@Value, LastModifiedBy=@LastModifiedBy, LastModifiedOn=@LastModifiedOn where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"
                sInsertQuery = "insert into tTimeAdjustments (DateAdjusted,EmployeeID,JobNumber,Company, LastModifiedBy, LastModifiedOn, PW, State) values (@DateAdjusted,@EmployeeID,@JobNumber,@Value, @LastModifiedBy, @LastModifiedOn,@PW,@State)"
                sSelectQuery = "select Company from tTimeAdjustments where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"
                sCheckQuery = "select count(EmployeeID) as RecordCount from tTimeAdjustments where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"

            ElseIf FieldType = "p" Then

                sUpdateQuery = "update tTimeAdjustments set Personal=@Value, LastModifiedBy=@LastModifiedBy, LastModifiedOn=@LastModifiedOn where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"
                sInsertQuery = "insert into tTimeAdjustments (DateAdjusted,EmployeeID,JobNumber,Personal, LastModifiedBy, LastModifiedOn, PW, State) values (@DateAdjusted,@EmployeeID,@JobNumber,@Value, @LastModifiedBy, @LastModifiedOn,@PW,@State)"
                sSelectQuery = "select Personal from tTimeAdjustments where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"
                sCheckQuery = "select count(EmployeeID) as RecordCount from tTimeAdjustments where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted and PW=@PW and State=@State"

            ElseIf FieldType = "bonus" Then

                sUpdateQuery = "update tBonus set Bonus=@Value, LastModifiedBy=@LastModifiedBy, LastModifiedOn=@LastModifiedOn where EmployeeID=@EmployeeID and JobNumber=@JobNumber and StartDate=@DateAdjusted and PW=@PW and State=@State"
                sInsertQuery = "insert into tBonus (StartDate,EmployeeID,JobNumber,Bonus, LastModifiedBy, LastModifiedOn, PW, State) values (@DateAdjusted,@EmployeeID,@JobNumber,@Value, @LastModifiedBy, @LastModifiedOn,@PW,@State)"
                sSelectQuery = "select Bonus from tBonus where EmployeeID=@EmployeeID and JobNumber=@JobNumber and StartDate=@DateAdjusted and PW=@PW and State=@State"
                sCheckQuery = "select count(EmployeeID) as RecordCount from tbonus where EmployeeID=@EmployeeID and JobNumber=@JobNumber and StartDate=@DateAdjusted and PW=@PW and State=@State"

            ElseIf FieldType = "notes" Then

                sUpdateQuery = "update tBonus set Notes=@Value, LastModifiedBy=@LastModifiedBy, LastModifiedOn=@LastModifiedOn where EmployeeID=@EmployeeID and JobNumber=@JobNumber and StartDate=@DateAdjusted and PW=@PW and State=@State"
                sInsertQuery = "insert into tBonus (StartDate,EmployeeID,JobNumber,Notes, LastModifiedBy, LastModifiedOn, PW, State) values (@DateAdjusted,@EmployeeID,@JobNumber,@Value, @LastModifiedBy, @LastModifiedOn,@PW,@State)"
                sSelectQuery = "select Notes from tBonus where EmployeeID=@EmployeeID and JobNumber=@JobNumber and StartDate=@DateAdjusted and PW=@PW and State=@State"
                sCheckQuery = "select count(EmployeeID) as RecordCount from tbonus where EmployeeID=@EmployeeID and JobNumber=@JobNumber and StartDate=@DateAdjusted and PW=@PW and State=@State"

            End If

            Response.Write(AdjustTime(sUpdateQuery, sInsertQuery, sSelectQuery, sCheckQuery, EmployeeID, JobNumber, DateAdjusted, ElementValue, PW, State))
        End If

    End Sub

    Function AdjustTime(ByVal sUpdateQuery As String, ByVal sInsertQuery As String, ByVal sSelectQuery As String, ByVal sCheckQuery As String, ByVal EmployeeID As Integer, ByVal JobNumber As String, ByVal DateAdjusted As DateTime, ByVal value As String, ByVal PW As String, ByVal State As String) As String

        Dim returnstring As String = ""
        Dim RecordCount As Integer
        Dim sAdjustmentQuery As String

        Dim connCheck As New SqlConnection(sConnection)
        'Dim sCheckQuery As String = "select count(Username) as RecordCount from tTimeAdjustments where EmployeeID=@EmployeeID and JobNumber=@JobNumber and DateAdjusted=@DateAdjusted"
        Dim cmdCheck As New SqlCommand(sCheckQuery, connCheck)
        cmdCheck.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
        cmdCheck.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
        cmdCheck.Parameters.Add(New SqlParameter("@DateAdjusted", DateAdjusted))
        cmdCheck.Parameters.Add(New SqlParameter("@PW", PW))
        cmdCheck.Parameters.Add(New SqlParameter("@State", State))
        Dim drCheck As SqlDataReader

        Try
            connCheck.Open()
            drCheck = cmdCheck.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drCheck.Read
                RecordCount = drCheck.Item("RecordCount")
            End While

            If RecordCount > 0 Then
                sAdjustmentQuery = sUpdateQuery
            Else
                sAdjustmentQuery = sInsertQuery
            End If

            Dim connUpdate As New SqlConnection(sConnection)
            Dim cmdUpdate As New SqlCommand(sAdjustmentQuery, connUpdate)
            cmdUpdate.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
            cmdUpdate.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
            cmdUpdate.Parameters.Add(New SqlParameter("@DateAdjusted", DateAdjusted))
            cmdUpdate.Parameters.Add(New SqlParameter("@Value", value))
            cmdUpdate.Parameters.Add(New SqlParameter("@LastModifiedBy", Session("UserID")))
            cmdUpdate.Parameters.Add(New SqlParameter("@LastModifiedOn", Now()))
            cmdUpdate.Parameters.Add(New SqlParameter("@PW", PW))
            cmdUpdate.Parameters.Add(New SqlParameter("@State", State))
            Try
                connUpdate.Open()
                cmdUpdate.ExecuteNonQuery()

                Dim connSelect As New SqlConnection(sConnection)
                Dim cmdSelect As New SqlCommand(sSelectQuery, connSelect)
                Dim drSelect As SqlDataReader
                cmdSelect.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))
                cmdSelect.Parameters.Add(New SqlParameter("@JobNumber", JobNumber))
                cmdSelect.Parameters.Add(New SqlParameter("@DateAdjusted", DateAdjusted))
                cmdSelect.Parameters.Add(New SqlParameter("@PW", PW))
                cmdSelect.Parameters.Add(New SqlParameter("@State", State))
                Try
                    connSelect.Open()
                    drSelect = cmdSelect.ExecuteReader(Data.CommandBehavior.CloseConnection)

                    While drSelect.Read
                        returnstring = drSelect.Item(0)
                    End While

                Catch ex As Exception
                    LogError("ax-editpayroll.aspx.vb", ex.ToString)
                    returnstring = "ERROR: " & ex.ToString
                Finally
                    connSelect.Close()
                End Try

            Catch ex As Exception
                LogError("ax-editpayroll.aspx.vb", ex.ToString)
                returnstring = "ERROR: " & ex.ToString
            Finally
                connUpdate.Close()
            End Try

        Catch ex As Exception
            LogError("ax-editpayroll.aspx.vb", ex.ToString)
            returnstring = "ERROR: " & ex.ToString
        Finally
            connCheck.Close()
        End Try


        Return returnstring

    End Function

End Class
