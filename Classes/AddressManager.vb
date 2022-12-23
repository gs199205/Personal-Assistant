Imports System
Imports System.Data
Imports System.Data.OleDb

Public Class AddressManager

    Public Shared Sub DeleteAddress(ByVal ID As System.Int32)
        Dim ConnectionString As String = System.Configuration.ConfigurationSettings.AppSettings("DSN")
        Dim myConnection As OleDbConnection = New OleDbConnection(ConnectionString)
        Dim query As String = "delete from Address where ID = " & ID
        DataManager.ExecuteNonQuery(ConnectionString, query)
    End Sub


    Public Shared Sub CreateAddress(ByVal AddressObj As Address)
        Dim ConnectionString As String = System.Configuration.ConfigurationSettings.AppSettings("DSN")
        Dim query As String = "insert into ADDRESS ( ID, Name, Address, Email, URL, HomePhone, CellPhone, WorkPhone, CategoryId, DateOfBirth, Remarks, Status ) values ( " & AddressObj.ID & ", '" & AddressObj.Name & "', '" & AddressObj.Address & "', '" & AddressObj.Email & "', '" & AddressObj.URL & "', '" & AddressObj.HomePhone & "', '" & AddressObj.CellPhone & "', '" & AddressObj.WorkPhone & "', " & AddressObj.CategoryId & ", '" & AddressObj.DateOfBirth & "', '" & AddressObj.Remarks & "', " & AddressObj.Status & " ) "

        DataManager.ExecuteNonQuery(ConnectionString, query)
    End Sub


    Public Shared Sub UpdateAddress(ByVal AddressObj As Address)
        Dim ConnectionString As String = System.Configuration.ConfigurationSettings.AppSettings("DSN")
        Dim query As String = "update ADDRESS set ID = " & AddressObj.ID & ", Name = '" & AddressObj.Name & "', Address = '" & AddressObj.Address & "', Email = '" & AddressObj.Email & "', URL = '" & AddressObj.URL & "', HomePhone = '" & AddressObj.HomePhone & "', CellPhone = '" & AddressObj.CellPhone & "', WorkPhone = '" & AddressObj.WorkPhone & "', CategoryId = " & AddressObj.CategoryId & ", DateOfBirth = '" & AddressObj.DateOfBirth & "', Remarks = '" & AddressObj.Remarks & "', Status = " & AddressObj.Status & " where ID = " & AddressObj.ID

        DataManager.ExecuteNonQuery(ConnectionString, query)
    End Sub


    Public Shared Function GetAddress(ByVal ID As System.Int32) As Address
        Dim ConnectionString As String = System.Configuration.ConfigurationSettings.AppSettings("DSN")
        Dim query As String = "select * from ADDRESS where ID = " & ID

        Dim Dt As DataTable = DataManager.ExecuteQuery(ConnectionString, query)
        If (Dt.Rows.Count = 0) Then
            Return Nothing
        End If

        Return New Address(Dt.Rows(0))
    End Function


    Public Shared Function GetAddresses(ByVal criteria As String) As DataTable
        Dim ConnectionString As String = System.Configuration.ConfigurationSettings.AppSettings("DSN")
        Dim query As String = "select * from ADDRESS"

        If criteria <> "" Then
            query = query & " where " & criteria
        End If

        Dim Dt As DataTable = DataManager.ExecuteQuery(ConnectionString, query, "Address")
        Return Dt
    End Function

End Class

