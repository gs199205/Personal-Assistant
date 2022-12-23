Imports System

Public Class Address

    'Public Attributes
    Public ID As Int32
    Public Name As String
    Public Address As String
    Public Email As String
    Public URL As String
    Public HomePhone As String
    Public CellPhone As String
    Public WorkPhone As String
    Public CategoryId As Int32
    Public DateOfBirth As DateTime = Nothing
    Public Remarks As String
    Public Status As eStatus = eStatus.Approved


    ' Default constructor
    Public Sub New()
    End Sub

    ' Constructor
    Public Sub New(ByVal dr As DataRow)
        If Not dr("ID") Is DBNull.Value Then
            Me.ID = dr("ID")
        End If

        If Not dr("Name") Is DBNull.Value Then
            Me.Name = dr("Name")
        End If

        If Not dr("Address") Is DBNull.Value Then
            Me.Address = dr("Address")
        End If

        If Not dr("Email") Is DBNull.Value Then
            Me.Email = dr("Email")
        End If

        If Not dr("URL") Is DBNull.Value Then
            Me.URL = dr("URL")
        End If

        If Not dr("HomePhone") Is DBNull.Value Then
            Me.HomePhone = dr("HomePhone")
        End If

        If Not dr("CellPhone") Is DBNull.Value Then
            Me.CellPhone = dr("CellPhone")
        End If

        If Not dr("WorkPhone") Is DBNull.Value Then
            Me.WorkPhone = dr("WorkPhone")
        End If

        If Not dr("CategoryId") Is DBNull.Value Then
            Me.CategoryId = dr("CategoryId")
        End If

        If Not dr("DateOfBirth") Is DBNull.Value Then
            Me.DateOfBirth = dr("DateOfBirth")
        End If

        If Not dr("Remarks") Is DBNull.Value Then
            Me.Remarks = dr("Remarks")
        End If

        If Not dr("Status") Is DBNull.Value Then
            Me.Status = dr("Status")
        End If

    End Sub

End Class

