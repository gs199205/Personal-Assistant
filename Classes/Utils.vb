Imports System.Data
Imports System.Data.OleDb
Imports System.Text
Imports System.Diagnostics
Imports System


Public Class Utils

    Public Shared Function LogError(ByVal Ex As Exception, Optional ByVal Message As String = "") As Boolean
        Try
            Dim msg As String

            If Not Ex Is Nothing Then
                msg = msg & "<BR><BR><b>Exception : </b><font color=RED>" & Ex.Message & "</font>"
            End If

            msg = msg & "<BR><BR><b>Description : </b><font color=RED>" & Message & "</font><BR><BR>"

        Catch ex1 As Exception
        End Try
        Return True
    End Function

    Public Shared Function ReplaceEscapeChars(ByVal Str As String) As String
        If Str = "" Then
            Return Str
        End If

        Str = Str.Replace("'", "''")
        Str = Str.Replace("""", """")
        Return Str
    End Function

    Public Shared Function PopulateCombo(ByVal Combo As System.Windows.Forms.ComboBox, ByVal ConnectionString As String, ByVal TableName As String, ByVal TextField As String, ByVal ValueField As String, Optional ByVal DefaultValue As String = Nothing, Optional ByVal Criteria As String = Nothing) As Boolean
        Combo.DataSource = Nothing

        Dim myConnection As OleDbConnection = New OleDbConnection(ConnectionString)

        Dim Query As String = "select * from " & TableName
        If Criteria <> "" AndAlso Not Criteria Is Nothing Then
            Query = Query & " where " & Criteria
        End If
        Query = Query & " order by " & TextField

        Dim myAdapter As OleDbDataAdapter = New OleDbDataAdapter(Query, myConnection)
        Dim ds As DataSet = New DataSet
        myAdapter.Fill(ds, TableName)
        Dim Dt As DataTable = ds.Tables(TableName)
        If (Dt.Rows.Count = 0) Then
            Return False
        End If

        Combo.DataSource = Dt
        Combo.DisplayMember = TextField
        Combo.ValueMember = ValueField
        Combo.SelectedValue = DefaultValue

        Return True
    End Function

End Class
