Imports System.Data.SqlClient

Public Class Form3

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles Me.Load
        Label1.Select()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Form1.Show()
        Me.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        'Save the address of the 1st display cell
        Dim i As Integer = Form1.DataGridView1.CurrentRow.Index()




        If Me.MaskedTextBox1.MaskCompleted = False Or Me.MaskedTextBox2.MaskCompleted = False Then
            'MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Remove User!")
            MsgBox("Please fill-up all fields!", MessageBoxIcon.Error, "Remove User!")
            'lblErrorIndicator.Visible = True
            'lblErrorIndicator.Text = "Bad Number"
        Else

            cmd.Parameters.Add(New SqlParameter("@timei", MaskedTextBox1.Text))
            cmd.Parameters.Add(New SqlParameter("@timej", MaskedTextBox2.Text))
            cmd.Parameters.Add(New SqlParameter("@tabid", TextBox5.Text))
            cmd.Parameters.Add(New SqlParameter("@user", TextBox4.Text))
            cmd.CommandText = "UPDATE BKTIME_LOG " _
                            & "SET LOGIN_TIME = CONCAT(CAST(LOGIN_TIME AS Date), ' ', +@timei +'.0000000')," _
                            & "LOGOUT_TIME = CONCAT(CAST(LOGOUT_TIME AS Date), ' ', +@timej +'.0000000'), UPDATED_BY = @user WHERE BKTIME_ID =@tabid"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            ' Remove user
            Dim mResult
            mResult = MsgBox("Are you sure you want to change the time of  " + TextBox1.Text + "?", vbYesNo + vbQuestion, "Change Confirmation")
            If mResult = vbNo Then
                Exit Sub
            Else
                cmd.ExecuteNonQuery()
                MsgBox("Records Successfully changed!", MsgBoxStyle.Information, "Change Time!")

            End If
       
            sqlConnection1.Close()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            MaskedTextBox1.Text = ""
            MaskedTextBox2.Text = ""

            'Reload data!
            display()
            Me.Hide()
            Form1.Show()


            'set the grid back the the first cell displayed
            Form1.DataGridView1.Rows(i).Selected = True
            'set color of the modified cell
            Form1.DataGridView1.Rows(i).DefaultCellStyle.SelectionBackColor = Color.IndianRed
            'Form1.DataGridView1.RowsDefaultCellStyle.SelectionBackColor = Color.IndianRed

            End If


    End Sub
    Sub display()
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable

        cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, CONVERT(varchar(8),LOGIN_TIME, 108) AS LOGIN_TIME, CONVERT(varchar(8), LOGOUT_TIME, 108) AS LOGOUT_TIME, DATEDIFF(s, LOGIN_TIME,LOGOUT_TIME) AS LOGGED_IN_TIME, BKTIME_ID FROM BKTIME_LOG ORDER BY ROW_DATE DESC, AGENT_NAME ASC, LOGIN ASC, LOGOUT ASC"
        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()

        reader = cmd.ExecuteReader()

        Form1.DataGridView1.DataSource = Nothing


        'Do While reader.Read = True

        dt.Load(reader)


        Form1.DataGridView1.DataSource = dt
        Form1.DataGridView1.Columns("BKTIME_ID").Visible = False
        'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'DataGridView1.AutoResizeColumns()

        Form1.DataGridView1.ClearSelection()

        'Loop

        sqlConnection1.Close()
    End Sub
End Class