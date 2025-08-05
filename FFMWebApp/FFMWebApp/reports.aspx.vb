Imports System.Data.SqlClient

Partial Class reports
    Inherits GlobalClass

    Sub page_load()

        If IsPostBack = False Then
            enddate.Text = Now().Date
            startdate.Text = DateAdd(DateInterval.Day, -7, Now().Date)
        End If

        LIT_JobNumberDropDown.Text = GenerateJobNumberList()

    End Sub

    Function GenerateJobNumberList() As String

        Dim returnstring As String = ""

        Dim StartDate As Date = DateAdd(DateInterval.Day, -14, Now().Date)
        Dim EndDate As Date = Now().Date

        Dim sJobNumberQuery As String = "select distinct fJobNumber from tShifts where ShiftID in (select distinct ShiftID from tTimeEntry where CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)>=@StartTime and CAST(FLOOR(CAST(StartTime AS FLOAT))AS DATETIME)<=@EndTime) order by fJobNumber asc"
        Dim connJobNumberQuery As New SqlConnection(sConnection)
        Dim cmdJobNumberQuery As New SqlCommand(sJobNumberQuery, connJobNumberQuery)
        cmdJobNumberQuery.Parameters.Add(New SqlParameter("@StartTime", StartDate))
        cmdJobNumberQuery.Parameters.Add(New SqlParameter("@EndTime", EndDate))

        Dim drJobNumberQuery As SqlDataReader

        Try
            connJobNumberQuery.Open()

            drJobNumberQuery = cmdJobNumberQuery.ExecuteReader(Data.CommandBehavior.CloseConnection)

            While drJobNumberQuery.Read()
                returnstring = returnstring & "<option>" & drJobNumberQuery.Item("fJobNumber") & "</option>"
            End While

        Catch ex As Exception
            Response.Write("ERROR: " & ex.ToString & "<br>")
        Finally
            connJobNumberQuery.Close()
        End Try

        returnstring = "<select name=""jobnumber"" id=""jobnumber"">" & returnstring & "</select>"

        Return returnstring

    End Function


End Class
