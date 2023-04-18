Imports MySql.Data.MySqlClient

Module Module1
    Public myconn As New MySql.Data.MySqlClient.MySqlConnection
    Public myConnectionString As String
    Public strSQL As String

    Public Sub Connect2DB()
        myConnectionString = "server=localhost;" _
                & "user id=root;" _
                & "port=3306;" _
                & "password=mariahannah25;" _
                & "database=database"
        Try
            myconn.ConnectionString = myConnectionString
            myconn.Open()

        Catch ex As MySqlException
            Select Case ex.Number
                Case 0
                    MsgBox("Cannot connect to the server!")
                Case 1045
                    MsgBox("Invalid username or password")
            End Select

        End Try
    End Sub

    Public Sub Disconnect2DB()
        myconn.Close()
        myconn.Dispose()

    End Sub

End Module

