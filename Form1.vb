Imports MySql.Data.MySqlClient
Public Class Form1


    Private Sub PictureBox3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles ButtonLogin.Click
        If (TextBox1.Text = "") Then
            MsgBox("Enter Username")
            Exit Sub
        End If
        If (TextBox2.Text = "") Then
            MsgBox("Enter Password")
        End If
        With Me
            Call Connect2DB()
            Dim mycomm As New MySqlCommand
            Dim myreader As MySqlDataReader

            Dim mysql As String
            mysql = "select * from student where fname = '" & TextBox1.Text & "'and stud_password = '" & TextBox2.Text & "'"
            mycomm.Connection = myconn
            mycomm.CommandText = mysql

            myreader = mycomm.ExecuteReader
            If myreader.HasRows Then
                Form3.Show()
                .Hide()
                myconn.Close()
            Else
                MessageBox.Show("Login Sucessful!")
                myconn.Close()


            End If
        End With
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub
End Class