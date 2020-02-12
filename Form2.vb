Imports System.Data.SqlClient

Public Class Form2

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        'Dim dLastTimeAnythingIsDone As Date
        'dLastTimeAnythingIsDone = Now()
        Button1.Select()
        TextBox1.Text = Environment.UserName.ToUpper
        'Label1.Font = New Font(Label1.Font, FontStyle.Bold)
        'MsgBox(DateDiff("s", dLastTimeAnythingIsDone, Now), MsgBoxStyle.Exclamation)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable


        cmd.Parameters.Add(New SqlParameter("@user", TextBox1.Text))


        cmd.CommandText = "SELECT USERID, BKUSER_ID FROM BKTIME_USERS where USERID=@user"


        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()

        reader = cmd.ExecuteReader()


        If reader.HasRows Then
            Form1.Show()
            Me.Hide()
            reader.Close()
            Dim cmd1 As SqlCommand = New SqlCommand("UPDATE BKTIME_USERS SET LAST_LOGIN = GETDATE() WHERE USERID =@user", sqlConnection1)
            cmd1.Parameters.Add(New SqlParameter("@user", TextBox1.Text))
            cmd1.ExecuteNonQuery()

        Else
            MsgBox("Sorry, username not found; Register Your Usename First!", MsgBoxStyle.Exclamation, "User Not Found")
            TextBox1.Text = Environment.UserName
        End If

        sqlConnection1.Close()
    End Sub

End Class