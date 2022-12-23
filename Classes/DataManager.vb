Imports System.Data
Imports System.Data.OleDb


Public Class DataManager

    Public Shared Function ExecuteQuery(ByVal ConnectionString As String, ByVal query As String, Optional ByVal tableName As String = "Table1") As DataTable
        Try
            Dim myConnection As OleDbConnection = New OleDbConnection(ConnectionString)

            Dim myAdapter As OleDbDataAdapter = New OleDbDataAdapter(query, myConnection)
            Dim ds As DataSet = New DataSet
            myAdapter.Fill(ds, tableName)

            ds.Tables(0).TableName = tableName

            Return ds.Tables(0)
        Catch ex As Exception
            Try
                Utils.LogError(ex, "Error in ExecuteQuery() method in DataManager. <BR><BR>" & query)
            Catch
            End Try

            Throw
        End Try
    End Function

    Public Shared Sub ExecuteNonQuery(ByVal ConnectionString As String, ByVal query As String)
        Dim myConnection As OleDbConnection

        Try
            myConnection = New OleDbConnection(ConnectionString)
            myConnection.Open()

            Dim myCommand As New OleDbCommand(query, myConnection)
            myCommand.ExecuteNonQuery()
        Catch ex As Exception
            Try
                Utils.LogError(ex, "Error in ExecuteQuery() method in DataManager. <BR><BR>" & query)
            Catch
            End Try

            Throw
        Finally
            If myConnection.State = ConnectionState.Open Then
                myConnection.Close()
            End If
        End Try
    End Sub
End Class
