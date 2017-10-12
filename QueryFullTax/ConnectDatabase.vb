Imports System.Data.SqlClient

Public Class ConnectDatabase

    Dim _Server As String = ""
    Dim _Database As String = ""
    Dim _Username As String = ""
    Dim _Password As String = ""
    Dim _Port As String = ""

    Public Property Server As String
        Get
            Return _Server
        End Get
        Set(value As String)
            _Server = value
        End Set
    End Property

    Public Property Database As String
        Get
            Return _Database
        End Get
        Set(value As String)
            _Database = value
        End Set
    End Property

    Public Property Username As String
        Get
            Return _Username
        End Get
        Set(value As String)
            _Username = value
        End Set
    End Property

    Public Property Password As String
        Get
            Return _Password
        End Get
        Set(value As String)
            _Password = value
        End Set
    End Property

    Public Property Port As String
        Get
            Return _Port
        End Get
        Set(value As String)
            _Port = value
        End Set
    End Property


    Function FillData(sql As String) As ProcessReturnInfo
        Dim ret As New ProcessReturnInfo
        Try
            Dim connstring As String = [String].Format("Server={0};Port={1};" + "User Id={2};Password={3};Database={4};CommandTimeout={5};", Server, Port, Username, Password, Database, 500)
            Dim myConnectionBase As New Npgsql.NpgsqlConnection(connstring)
            myConnectionBase.Open()

            Dim myCommand As New Npgsql.NpgsqlCommand(sql, myConnectionBase)
            Dim ad As New Npgsql.NpgsqlDataAdapter(myCommand)
            Dim dt As New DataTable
            ad.Fill(dt)

            With ret
                .DT = dt
                .IsSuccess = True
                .ErrorMessage = ""
            End With
        Catch ex As Exception
            With ret
                .DT = Nothing
                .IsSuccess = False
                .ErrorMessage = ex.ToString
            End With
        End Try
        Return ret
    End Function

End Class
