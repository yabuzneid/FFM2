Imports System.IO
Imports System.Data
Imports System.Data.SqlClient

Partial Class ExportList
    Inherits GlobalClass
    Sub page_load()

        'LIT_FileListing.Text = GetFileList()

        GV_Files.DataSource = Fileinfo_To_DataTable()
        GV_Files.DataBind()

        'Dim DateColumn As DataGridColumn = GV_Files

        'GV_Files.Sort("Last Write Time", Descending)

    End Sub

    Function GetFileList() As String

        Dim returnstring As String = "<table>"


        Dim Files As String() = IO.Directory.GetFiles(Server.MapPath(".\SpreadsheetLogs\"))

        System.Array.Sort(Of String)(Files)

        For Each file As String In Files
            returnstring = returnstring & "<tr><td>" & file.ToString & "</td></tr>"
        Next

        returnstring = returnstring & "</table>"

        Return returnstring

    End Function

    Private Function Fileinfo_To_DataTable() As DataView
        Try
            Dim dt As DataTable = New DataTable
            Dim FileName As String, CreatedBy As String = "", CutUnderscore As Integer, CutPeriod As Integer
            dt.Columns.Add(New DataColumn("Name"))
            dt.Columns.Add(New DataColumn("User"))
            dt.Columns.Add(New DataColumn("Date"))

            Dim directoryPath As String = Server.MapPath(".\SpreadsheetLogs\")

            If Not IO.Directory.Exists(directoryPath) Then
                IO.Directory.CreateDirectory(directoryPath)
                Dim emptyDv As DataView = New DataView(dt)
                Return emptyDv
            End If

            Dim dirInfo As New IO.DirectoryInfo(directoryPath)

            Dim files() As IO.FileInfo = dirInfo.GetFiles()
            If files.Length = 0 Then
                Dim emptyDv As DataView = New DataView(dt)
                Return emptyDv
            End If

            For Each file As IO.FileInfo In files
                Dim dr As DataRow = dt.NewRow
                FileName = file.Name

                CutUnderscore = InStr(FileName, "_")
                If CutUnderscore > 0 Then
                    CreatedBy = FileName.Substring(CutUnderscore, FileName.Length - CutUnderscore)
                    CutPeriod = InStr(CreatedBy, ".")
                    If CutPeriod > 0 Then
                        CreatedBy = CreatedBy.Substring(0, CutPeriod - 1)
                    End If
                Else
                    CreatedBy = "Unknown"
                End If

                dr(0) = "<a href=""/ffm/SpreadSheetLogs/" & FileName & """ target=""_blank"">" & file.Name & "</a>"
                dr(1) = GetUsername(CreatedBy)
                dr(2) = file.LastWriteTime
                dt.Rows.Add(dr)
            Next

            Dim dv As DataView = New DataView(dt)
            dv.Sort = "Name desc"
            Return dv

        Catch ex As Exception
            Response.Write("Error in Fileinfo_To_DataTable: " & ex.ToString)

            Dim dt As DataTable = New DataTable
            dt.Columns.Add(New DataColumn("Name"))
            dt.Columns.Add(New DataColumn("User"))
            dt.Columns.Add(New DataColumn("Date"))
            Return New DataView(dt)
        End Try
    End Function

    Function GetUsername(EmployeeID As String) As String

        Dim returnstring As String = ""

        Dim strConn As String = System.Configuration.ConfigurationManager.ConnectionStrings("Fullconnection").ConnectionString
        Dim MySQL As String = "Select FirstName + ' ' + LastName as FullName from tEmployees where EmployeeID=@EmployeeID"
        Dim MyConn As New SqlConnection(strConn)
        Dim objDR As SqlDataReader
        Dim Cmd As New SqlCommand(MySQL, MyConn)
        Cmd.Parameters.Add(New SqlParameter("@EmployeeID", EmployeeID))

        Try
            MyConn.Open()
            objDR = Cmd.ExecuteReader(System.Data.CommandBehavior.CloseConnection)

            While objDR.Read()
                returnstring = objDR.Item("FullName")
            End While

        Catch ex As Exception
            LogError("ExportList.vb :: GetUsername", ex.ToString)
        Finally
            MyConn.Close()
        End Try

        Return returnstring

    End Function


    Protected Sub GV_Files_RowDataBound(sender As Object, e As GridViewRowEventArgs)

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim cells As TableCellCollection = e.Row.Cells

            For Each cell As TableCell In cells
                cell.Text = Server.HtmlDecode(cell.Text)
            Next
        End If

    End Sub
End Class




