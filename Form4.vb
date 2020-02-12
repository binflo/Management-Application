Imports System.Data.SqlClient

Public Class Form4

    Private Sub Form4_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        Form1.Show()
        Form1.RadioButton6.Checked = False
    End Sub


    Private Sub Form4_Load(sender As Object, e As EventArgs) Handles Me.Load
        TextBox1.Visible = False
        TextBox2.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        'TextBox6.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        Button1.Visible = False
        GroupBox1.Visible = False
        TextBox6.Text = Environment.UserName.ToUpper
        TextBox7.Text = Environment.UserName.ToUpper
        ComboBox1.Visible = False
        Button2.Visible = False
        ComboBox2.Visible = False
        ComboBox3.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable

        TextBox1.Visible = False
        TextBox2.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        Label1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Button1.Visible = False
        GroupBox1.Visible = False
        ComboBox1.Visible = False
        Button2.Visible = False

        'cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, LOGID, LOGIN_TIME, LOGOUT_TIME, LOGGED_IN_TIME, #CALLS_IN, CALLS_IN_TIME, #CALLS_OUT, CALLS_OUT_TIME, #CALLS_TRANSFERRED, #CALLS_ONHOLD, HOLD_CALLS_TIME FROM BKTIME_LOG ORDER BY AGENT_NAME ASC"
        cmd.CommandText = " SELECT EMP_NAME AS [NAME],EMP_SUP_NAME AS [NEXT LEVEL SUPERVISOR], EMP_TEAM AS [TEAM], EMP_STATUS AS [FTE/Temp]," _
                        & " EMP_SITE AS [SITE] FROM BK_EMPLOYEES ORDER BY EMP_NAME ASC"
        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()

        reader = cmd.ExecuteReader()

        DataGridView1.DataSource = Nothing


        'Do While reader.Read = True

        dt.Load(reader)


        DataGridView1.DataSource = dt

        'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'DataGridView1.AutoResizeColumns()

        DataGridView1.ClearSelection()

        'Loop

        sqlConnection1.Close()
    End Sub


    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged

    End Sub
    Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
        GroupBox1.Text = "Add Employee"
        GroupBox1.Visible = True
        TextBox1.Visible = True
        TextBox2.Visible = True
        TextBox3.Visible = True
        TextBox4.Visible = True
        TextBox5.Visible = True
        Label1.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        Button1.Visible = True
        ComboBox1.Visible = False
        Button2.Visible = False
        DataGridView1.DataSource = Nothing
    End Sub
    Private Sub RadioButton3_Click(sender As Object, e As EventArgs) Handles RadioButton3.Click
        GroupBox1.Text = "Remove Employee"
        TextBox1.Visible = False
        TextBox2.Visible = False
        TextBox3.Visible = False
        TextBox4.Visible = False
        TextBox5.Visible = False
        Label1.Visible = True
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        Button1.Visible = False
        GroupBox1.Visible = True
        ComboBox1.Visible = True
        Button2.Visible = True
        DataGridView1.DataSource = Nothing
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable

        If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Then
            MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Add New Employee!")
        Else

            cmd.Parameters.Add(New SqlParameter("@winid", TextBox1.Text))
            cmd.CommandText = "SELECT EMP_NAME FROM BK_EMPLOYEES WHERE EMP_NAME =@winid"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader()

            If reader.HasRows Then
                ' User already exists
                MsgBox("Employee Already Exist!", MsgBoxStyle.Exclamation, "Add Employee!")
            Else
                reader.Close()
                ' User does not exist, add them
                Dim cmd1 As SqlCommand = New SqlCommand("INSERT INTO BK_EMPLOYEES (EMP_NAME, EMP_SUP_NAME, EMP_TEAM, EMP_STATUS, EMP_SITE, ADDED_BY, ADDED_DATE) VALUES(@par1, @par2, @par3, @par4, @par5, @par6, GETDATE())", sqlConnection1)
                cmd1.Parameters.Add(New SqlParameter("@par1", TextBox1.Text))
                cmd1.Parameters.Add(New SqlParameter("@par2", TextBox2.Text))
                cmd1.Parameters.Add(New SqlParameter("@par3", TextBox3.Text))
                cmd1.Parameters.Add(New SqlParameter("@par4", TextBox4.Text))
                cmd1.Parameters.Add(New SqlParameter("@par5", TextBox5.Text))
                cmd1.Parameters.Add(New SqlParameter("@par6", TextBox6.Text))
                cmd1.ExecuteNonQuery()
                MsgBox("Records Successfully Added!", MsgBoxStyle.Information, "Add New Employee!")
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""
            End If
            sqlConnection1.Close()
        End If
    End Sub

    Private Sub ComboBox1_DropDown(sender As Object, e As EventArgs) Handles ComboBox1.DropDown
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim dt As New DataTable
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        'Dim strSql As String = "SELECT EMP_NAME, EMP_SUP_NAME, BKEMP_ID FROM BK_EMPLOYEES ORDER BY EMP_NAME ASC"
        'Dim dad As SqlDataAdapter

        'cmd.Parameters.Add(New SqlParameter("@winid", ComboBox1.Text))
        cmd.CommandText = "SELECT EMP_NAME, EMP_SUP_NAME, BKEMP_ID FROM BK_EMPLOYEES ORDER BY EMP_NAME ASC;"

        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()
        'Console.WriteLine("Connection Opened")

        reader = cmd.ExecuteReader()


        ' Data is accessible through the DataReader object here.

        'ComboBox1.AutoCompleteSource = AutoCompleteSource.ListItems
        'ComboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend


        'cboColours.DisplayMember = dtColours.Columns("Colour").ToString
        'cboColours.ValueMember = dtColours.Columns("Colour_id").ToString

        'Fill a combo box with the datareader
        ComboBox1.Items.Clear()



        'dad = New SqlDataAdapter(strSql, sqlConnection1)

        'dad.Fill(dt)
        'ComboBox1.DataSource = dt

        'ComboBox1.Text = dt(1).ToString
        'TextBox8.Text = dt(2).ToString


        Do While reader.Read = True



            ComboBox1.Items.Add(reader.GetString(0))
            'TextBox8.Text = reader(2).ToString


        Loop

        sqlConnection1.Close()


        '------------------------------------------------------------------------------------------
        'NEW LOGIC
        '------------------------------------------------------------------------------------------
        'Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        'Dim qry As String = "SELECT EMP_NAME, EMP_SUP_NAME, BKEMP_ID FROM BK_EMPLOYEES ORDER BY EMP_NAME ASC;"
        'Dim dt As New DataTable
        'Using cmd As New SqlCommand(qry, sqlConnection1)
        '    sqlConnection1.Open()
        '    Using leitor As New SqlDataAdapter
        '        leitor.SelectCommand = cmd
        '        leitor.Fill(dt)



        '        ComboBox1.DataSource = dt
        '        ComboBox1.DisplayMember = "EMP_NAME"
        '        ComboBox1.ValueMember = "BKEMP_ID"
        '        'TextBox8.Text = CStr(dt.Rows.Item("BKEMP_ID").ToString)
        '        TextBox8.Text = dt.Rows(0).Item("BKEMP_ID")



        '    End Using
        'End Using
    End Sub

    Private Sub ComboBox3_DropDown(sender As Object, e As EventArgs) Handles ComboBox3.DropDown
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader

        cmd.CommandText = "SELECT EMP_NAME, EMP_SUP_NAME FROM BK_EMPLOYEES ORDER BY EMP_NAME ASC;"

        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()
        'Console.WriteLine("Connection Opened")

        reader = cmd.ExecuteReader()
        ' Data is accessible through the DataReader object here.

        'ComboBox1.AutoCompleteSource = AutoCompleteSource.ListItems
        'ComboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend

        'Fill a combo box with the datareader
        ComboBox3.Items.Clear()
        Do While reader.Read = True


            ComboBox3.Items.Add(reader.GetString(0))

        Loop

        sqlConnection1.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand


        cmd.Parameters.Add(New SqlParameter("@winid", ComboBox1.Text))
        cmd.Parameters.Add(New SqlParameter("@empid", TextBox8.Text))
        cmd.CommandText = "DELETE FROM BK_EMPLOYEES WHERE EMP_NAME =@winid AND BKEMP_ID = @empid"
        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()


        If ComboBox1.Text = "" Then
            MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Remove Employee!")
            ComboBox1.Items.Clear()
        Else
            ' Remove user
            Dim mResult
            mResult = MsgBox("Are you sure you want to remove the Employee " + ComboBox1.Text + "?", vbYesNo + vbQuestion, "Removal Confirmation")
            If mResult = vbNo Then
                ComboBox1.Items.Clear()
                Exit Sub
            Else
                cmd.ExecuteNonQuery()
                MsgBox("Records Successfully Removed!", MsgBoxStyle.Information, "Remove Employee!")
                ComboBox1.Text = ""
                TextBox8.Text = ""
            End If

        End If
        sqlConnection1.Close()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader


        cmd.Parameters.Add(New SqlParameter("@winid", ComboBox1.Text))
        cmd.CommandText = "SELECT EMP_NAME, EMP_SUP_NAME, BKEMP_ID FROM BK_EMPLOYEES WHERE EMP_NAME =@winid ORDER BY EMP_NAME"
        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()

        'Console.WriteLine("Connection Opened")

        reader = cmd.ExecuteReader()




        TextBox8.Clear()

        Do While reader.Read = True

            TextBox8.Text = Convert.ToString(reader(2))


        Loop

        sqlConnection1.Close()

    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable

        ComboBox2.Visible = False
        ComboBox3.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False



        'cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, LOGID, LOGIN_TIME, LOGOUT_TIME, LOGGED_IN_TIME, #CALLS_IN, CALLS_IN_TIME, #CALLS_OUT, CALLS_OUT_TIME, #CALLS_TRANSFERRED, #CALLS_ONHOLD, HOLD_CALLS_TIME FROM BKTIME_LOG ORDER BY AGENT_NAME ASC"
        cmd.CommandText = "SELECT EMP_NAME, EVENT_TYPE, START_TIME, END_TIME, MODIFIED_BY, MODIFIED_DATE, BKEVENT_ID FROM BKEVENT_LOG  ORDER BY EMP_NAME ASC"
        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()

        reader = cmd.ExecuteReader()

        DataGridView2.DataSource = Nothing


        'Do While reader.Read = True

        dt.Load(reader)


        DataGridView2.DataSource = dt
        DataGridView2.Columns("BKEVENT_ID").Visible = False
        'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'DataGridView1.AutoResizeColumns()

        DataGridView2.ClearSelection()

        'Loop

        sqlConnection1.Close()


    End Sub

    Private Sub RadioButton5_Click(sender As Object, e As EventArgs) Handles RadioButton5.Click
        ComboBox2.Visible = True
        ComboBox3.Visible = True
        MaskedTextBox1.Visible = True
        DataGridView2.DataSource = Nothing
        Button3.Visible = True
        Button4.Visible = False
        MaskedTextBox2.Visible = True
        Label6.Visible = True
        Label7.Visible = True
        Label8.Visible = True
        Label9.Visible = True
        DateTimePicker1.Visible = True
        DateTimePicker2.Visible = True
    End Sub

    Private Sub RadioButton6_Click(sender As Object, e As EventArgs) Handles RadioButton6.Click
        'ComboBox2.Visible = False
        'ComboBox3.Visible = True
        'MaskedTextBox1.Visible = False
        'MaskedTextBox2.Visible = False
        'DataGridView2.DataSource = Nothing
        'Button4.Visible = True
        'Button3.Visible = False
        'Label6.Visible = True
        'Label7.Visible = False
        'Label8.Visible = False
        'Label9.Visible = False

        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable

        ComboBox2.Visible = False
        ComboBox3.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        Button3.Visible = False
        Button4.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Label9.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False



        'cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, LOGID, LOGIN_TIME, LOGOUT_TIME, LOGGED_IN_TIME, #CALLS_IN, CALLS_IN_TIME, #CALLS_OUT, CALLS_OUT_TIME, #CALLS_TRANSFERRED, #CALLS_ONHOLD, HOLD_CALLS_TIME FROM BKTIME_LOG ORDER BY AGENT_NAME ASC"
        cmd.CommandText = "SELECT EMP_NAME, EVENT_TYPE, START_TIME, END_TIME, MODIFIED_BY, MODIFIED_DATE, BKEVENT_ID FROM BKEVENT_LOG  ORDER BY EMP_NAME ASC"
        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()

        reader = cmd.ExecuteReader()

        DataGridView2.DataSource = Nothing


        'Do While reader.Read = True

        dt.Load(reader)


        DataGridView2.DataSource = dt
        DataGridView2.Columns("BKEVENT_ID").Visible = False
        'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'DataGridView1.AutoResizeColumns()

        DataGridView2.ClearSelection()

        'Loop

        sqlConnection1.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        'Dim reader As SqlDataReader
        Dim dt As New DataTable

        If ComboBox2.Text = "" Or ComboBox3.Text = "" Or Me.MaskedTextBox1.MaskCompleted = False Or Me.MaskedTextBox2.MaskCompleted = False Then
            MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Add New Event!")
            MaskedTextBox1.Clear()
            MaskedTextBox2.Clear()
            ComboBox2.Items.Clear()
            ComboBox3.Items.Clear()
        Else
            ' User does not exist, add them
            Dim cmd1 As SqlCommand = New SqlCommand("INSERT INTO BKEVENT_LOG (EMP_NAME, EVENT_TYPE, START_TIME, END_TIME, MODIFIED_BY, MODIFIED_DATE) VALUES(@par, @par1, @par2, @par3, @user, GETDATE())", sqlConnection1)
            cmd1.Parameters.Add(New SqlParameter("@par", ComboBox3.Text))
            cmd1.Parameters.Add(New SqlParameter("@par1", ComboBox2.Text))
            cmd1.Parameters.Add(New SqlParameter("@par2", MaskedTextBox1.Text))
            cmd1.Parameters.Add(New SqlParameter("@par3", MaskedTextBox2.Text))
            cmd1.Parameters.Add(New SqlParameter("@user", TextBox7.Text))
            cmd1.ExecuteNonQuery()
            MsgBox("Records Successfully Added!", MsgBoxStyle.Information, "Add New Employee Event!")
            TextBox1.Text = ""
        End If
            sqlConnection1.Close()
        ComboBox2.Text = ""
        ComboBox3.Text = ""
        MaskedTextBox1.Text = ""
        MaskedTextBox2.Text = ""
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
     
    End Sub

    Private Sub DataGridView2_RowHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.RowHeaderMouseDoubleClick
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=MI_Temp_Data;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim i As Integer

        'If ComboBox2.Text = "" Then
        '    MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Remove User!")
        'Else

        cmd.Parameters.Add(New SqlParameter("@user", Me.DataGridView2.Item(0, i).Value))
        cmd.Parameters.Add(New SqlParameter("@winid", Me.DataGridView2.Item(6, i).Value))
        cmd.CommandText = "DELETE FROM BKEVENT_LOG WHERE EMP_NAME =@user and BKEVENT_ID =@winid"
        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        'MsgBox(Me.DataGridView2.Item(0, i).Value, MsgBoxStyle.Information)
        'MsgBox(Me.DataGridView2.Item(6, i).Value, MsgBoxStyle.Information)


        sqlConnection1.Open()

        ' Remove user
        Dim mResult
        mResult = MsgBox("Are you sure you want to remove this event from the list ?", vbYesNo + vbQuestion, "Removal Confirmation")
        If mResult = vbNo Then
            DataGridView2.ClearSelection()
            Exit Sub
        Else
            cmd.ExecuteNonQuery()
            MsgBox("Records Successfully Removed!", MsgBoxStyle.Information, "Remove Event!")
            ComboBox2.Items.Clear()
        End If
        'End If
        sqlConnection1.Close()
    End Sub

    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged

        MaskedTextBox1.Text = DateTimePicker1.Value.ToString("MM/dd/yyyy hh:mm")

    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        MaskedTextBox2.Text = DateTimePicker1.Value.ToString("MM/dd/yyyy hh:mm")
    End Sub

End Class
