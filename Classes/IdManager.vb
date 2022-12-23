Imports System
Imports System.Data
Imports System.Data.OleDb


Public Class IdManager

    Public Shared Function GetNextId(ByVal tableName As String, ByVal idField As String) As Int32
        ' Look for the upgrade records for any profile of this user.
        Dim ConnectionString As String = Configuration.ConfigurationSettings.AppSettings("dsn")
        Dim myConnection As OleDbConnection = New OleDbConnection(ConnectionString)

        Dim Query As String = "select max(" & idField & ") from " & tableName

        myConnection.Open()
        Dim myCommand As OleDbCommand = New OleDbCommand(Query, myConnection)
        Dim maxValue As Object = myCommand.ExecuteScalar()
        myConnection.Close()

        If maxValue Is DBNull.Value Then
            Return 1
        Else
            Return Int32.Parse(maxValue) + 1
        End If
    End Function


End Class
