Imports System.Data.SqlClient


Public Class Form1

    'Private WithEvents NonActivityTimer As New Timer With {.Interval = 5000} '5 seconds for testing
    'Private WithEvents NonActivityTimer As New Timer With {.Interval = 600000} '10 min
    ' Private WithEvents NonActivityTimer As New Timer With {.Interval = 1800000} '30 min

 

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        ComboBox1.Visible = False
        ComboBox2.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        TextBox2.Visible = False
        Label1.Visible = False
        TextBox1.Visible = False
        Button1.Visible = False
        Button4.Visible = False
        TextBox2.Text = (Environment.UserName).ToUpper
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False

        DateTimePicker3.Visible = False
        DateTimePicker4.Visible = False
        MaskedTextBox3.Visible = False
        MaskedTextBox4.Visible = False
        Button5.Visible = False
        Label10.Visible = False

        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Button6.Visible = False
        Label11.Visible = False
        'Label12.Visible = False

        CheckBox1.Visible = False


        'AddActivityHandler(Me)
        'AddHandler Me.MouseDown, AddressOf NonActivityTimerStop
        'AddHandler Me.MouseMove, AddressOf NonActivityTimerStop
        'AddHandler Me.KeyDown, AddressOf NonActivityTimerStop
        'NonActivityTimer.Start()


        'DateTimePicker1.Format = DateTimePickerFormat.Custom
        'DateTimePicker1.CustomFormat = " "
        'DateTimePicker2.Format = DateTimePickerFormat.Custom
        'DateTimePicker2.CustomFormat = " "
    End Sub

    'Private Sub AddActivityHandler(ByVal ThisParent As Control)
    '    For Each ctr As Control In ThisParent.Controls
    '        AddHandler ctr.MouseDown, AddressOf NonActivityTimerStop
    '        AddHandler ctr.KeyDown, AddressOf NonActivityTimerStop
    '        AddActivityHandler(ctr)
    '    Next
    'End Sub


    'Private Sub NonActivityTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NonActivityTimer.Tick
    '    NonActivityTimer.Stop()
    '    If MessageBox.Show("Application hasn't been used for 30 min, Continue?", "", MessageBoxButtons.YesNo) <> Windows.Forms.DialogResult.Yes Then
    '        'Application.Exit()
    '        Environment.Exit(0)

    '    ElseIf MessageBox.Show("Application hasn't been used for 30 min, Continue?", "", MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.None Then
    '        'Application.Exit()
    '        Environment.Exit(0)
    '    End If
    '    NonActivityTimer.Start()
    'End Sub
    'Private Sub NonActivityTimerStop(ByVal sender As Object, ByVal e As System.EventArgs)
    '    With NonActivityTimer
    '        .Stop()
    '        .Start()
    '    End With
    'End Sub

    Private Sub ComboBox1_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox1.DropDown

        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader

        cmd.CommandText = "SELECT DISTINCT AGENT_NAME FROM BKTIME_LOG ORDER BY AGENT_NAME ASC"
        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()
        'Console.WriteLine("Connection Opened")

        reader = cmd.ExecuteReader()
        ' Data is accessible through the DataReader object here.

        'ComboBox1.AutoCompleteSource = AutoCompleteSource.ListItems
        'ComboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend

        'Fill a combo box with the datareader
        ComboBox1.Items.Clear()
        Do While reader.Read = True


            ComboBox1.Items.Add(reader.GetString(0))

        Loop

        sqlConnection1.Close()
    End Sub

    Private Sub ComboBox2_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles ComboBox2.DropDown

        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader

        cmd.CommandText = "SELECT USERID, BKUSER_ID FROM BKTIME_USERS ORDER BY USERID ASC;"

        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()
        'Console.WriteLine("Connection Opened")

        reader = cmd.ExecuteReader()
        ' Data is accessible through the DataReader object here.

        'ComboBox1.AutoCompleteSource = AutoCompleteSource.ListItems
        'ComboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend

        'Fill a combo box with the datareader
        ComboBox2.Items.Clear()
        Do While reader.Read = True


            ComboBox2.Items.Add(reader.GetString(0))

        Loop

        sqlConnection1.Close()
    End Sub


    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged

        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable

        ComboBox1.Items.Clear()
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        ComboBox1.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        ComboBox2.Visible = False
        Button4.Visible = False
        Label8.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        TextBox1.Visible = False
        Label1.Visible = False
        Button1.Visible = False

        DateTimePicker3.Visible = False
        DateTimePicker4.Visible = False
        MaskedTextBox3.Visible = False
        MaskedTextBox4.Visible = False
        Button5.Visible = False
        Label10.Visible = False

        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Button6.Visible = False
        Label11.Visible = False
        'Label12.Visible = False
        CheckBox1.Visible = False

        'cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, LOGID, LOGIN_TIME, LOGOUT_TIME, LOGGED_IN_TIME, #CALLS_IN, CALLS_IN_TIME, #CALLS_OUT, CALLS_OUT_TIME, #CALLS_TRANSFERRED, #CALLS_ONHOLD, HOLD_CALLS_TIME FROM BKTIME_LOG ORDER BY AGENT_NAME ASC"
        cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, CONVERT(varchar(8),LOGIN_TIME, 108) AS LOGIN_TIME, CONVERT(varchar(8), LOGOUT_TIME, 108) AS LOGOUT_TIME, DATEDIFF(s, LOGIN_TIME,LOGOUT_TIME) AS LOGGED_IN_TIME, BKTIME_ID FROM BKTIME_LOG ORDER BY DAY_DATE DESC, AGENT_NAME ASC, LOGIN_TIME ASC, LOGOUT_TIME ASC"
        cmd.CommandType = CommandType.Text
        cmd.Connection = sqlConnection1

        sqlConnection1.Open()

        reader = cmd.ExecuteReader()

        DataGridView1.DataSource = Nothing


        'Do While reader.Read = True

        dt.Load(reader)


        DataGridView1.DataSource = dt
        DataGridView1.Columns("BKTIME_ID").Visible = False
        'DataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
        'DataGridView1.AutoResizeColumns()

        DataGridView1.ClearSelection()

        'Loop

        sqlConnection1.Close()


    End Sub

    Private Sub ComboBox1_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles ComboBox1.KeyDown
        e.Handled = True
    End Sub
    Private Sub ComboBox1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ComboBox1.KeyPress
        e.Handled = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable

        If ComboBox1.Text = "" Then
            MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Select Agent!")
            ComboBox1.Items.Clear()
        Else
            cmd.Parameters.Add(New SqlParameter("@winid", ComboBox1.Text))

            'cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, LOGID, LOGIN_TIME, LOGOUT_TIME, LOGGED_IN_TIME, #CALLS_IN, CALLS_IN_TIME, #CALLS_OUT, CALLS_OUT_TIME, #CALLS_TRANSFERRED, #CALLS_ONHOLD, HOLD_CALLS_TIME FROM BKTIME_LOG WHERE AGENT_NAME=@winid ORDER BY AGENT_NAME ASC"
            cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, CONVERT(varchar(8),LOGIN_TIME, 108) AS LOGIN_TIME, CONVERT(varchar(8), LOGOUT_TIME, 108) AS LOGOUT_TIME, DATEDIFF(s, LOGIN_TIME,LOGOUT_TIME) AS LOGGED_IN_TIME, BKTIME_ID FROM BKTIME_LOG WHERE AGENT_NAME=@winid ORDER BY DAY_DATE DESC, AGENT_NAME ASC, LOGIN_TIME ASC, LOGOUT_TIME ASC"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader()

            DataGridView1.DataSource = Nothing


            'Do While reader.Read = True

            dt.Load(reader)


            DataGridView1.DataSource = dt
            DataGridView1.Columns("BKTIME_ID").Visible = False

            DataGridView1.ClearSelection()

            'Loop
            'ComboBox1.Items.Clear()
        End If
        sqlConnection1.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable

        If ComboBox1.Text = "" Or MaskedTextBox1.Text = "" Or MaskedTextBox2.Text = "" Then
            MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Select Agent!")
            MaskedTextBox1.Clear()
            MaskedTextBox2.Clear()
            ComboBox1.Items.Clear()
        Else

            cmd.Parameters.Add(New SqlParameter("@winid", ComboBox1.Text))
            cmd.Parameters.Add(New SqlParameter("@date1", DateTimePicker1.Value.ToString("yyyy-MM-dd")))
            cmd.Parameters.Add(New SqlParameter("@date2", DateTimePicker2.Value.ToString("yyyy-MM-dd")))


            'cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, LOGID, LOGIN_TIME, LOGOUT_TIME, LOGGED_IN_TIME, #CALLS_IN, CALLS_IN_TIME, #CALLS_OUT, CALLS_OUT_TIME, #CALLS_TRANSFERRED, #CALLS_ONHOLD, HOLD_CALLS_TIME FROM BKTIME_LOG WHERE AGENT_NAME=@winid AND DAY_DATE BETWEEN @date1 AND @date2  ORDER BY AGENT_NAME ASC"
            cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, CONVERT(varchar(8),LOGIN_TIME, 108) AS LOGIN_TIME, CONVERT(varchar(8), LOGOUT_TIME, 108) AS LOGOUT_TIME, DATEDIFF(s, LOGIN_TIME,LOGOUT_TIME) AS LOGGED_IN_TIME, BKTIME_ID FROM BKTIME_LOG WHERE AGENT_NAME=@winid AND DAY_DATE BETWEEN @date1 AND @date2  ORDER BY DAY_DATE DESC, AGENT_NAME ASC, LOGIN_TIME ASC, LOGOUT_TIME ASC"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader()

            DataGridView1.DataSource = Nothing


            'Do While reader.Read = True

            dt.Load(reader)


            DataGridView1.DataSource = dt
            DataGridView1.Columns("BKTIME_ID").Visible = False

            DataGridView1.ClearSelection()

            'Loop
        End If

        sqlConnection1.Close()
    End Sub


    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        MaskedTextBox1.Text = DateTimePicker1.Value.ToString("MM/dd/yyyy")
    End Sub

    Private Sub DateTimePicker2_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker2.ValueChanged
        MaskedTextBox2.Text = DateTimePicker2.Value.ToString("MM/dd/yyyy")
    End Sub

    Private Sub DataGridView1_RowHeaderMouseDoubleClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.RowHeaderMouseDoubleClick
        Dim i As Integer
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim dt As New DataTable
        Dim dt2 As New DataTable
        Dim ds As New DataSet()

        i = Me.DataGridView1.CurrentRow.Index
        Form3.TextBox1.Text = Me.DataGridView1.Item(1, i).Value
        Form3.TextBox2.Text = Me.DataGridView1.Item(0, i).Value
        Form3.MaskedTextBox1.Text = Me.DataGridView1.Item(2, i).Value
        Form3.MaskedTextBox2.Text = Me.DataGridView1.Item(3, i).Value
        Form3.TextBox5.Text = Me.DataGridView1.Item(5, i).Value
        Form3.TextBox4.Text = Environment.UserName.ToUpper
        'Form3.DateTimePicker2.Value = DateTime.Parse(Me.DataGridView1.Item(13, i).Value.ToShortdatestring)
        'Form3.MaskedTextBox1.Text = Me.DataGridView1.Item(6, i).Value.ToString("MM/dd/yyyy")
        'Form3.MaskedTextBox2.Text = Me.DataGridView1.Item(13, i).Value '.ToString("MM/dd/yyyy"))
        'Form3.MaskedTextBox3.Text = Me.DataGridView1.Item(11, i).Value '.ToString("MM/dd/YYYY")
        'Form3.MaskedTextBox4.Text = Me.DataGridView1.Item(10, i).Value '.ToString("MM/dd/YYYY")
        'Form3.MaskedTextBox5.Text = Me.DataGridView1.Item(9, i).Value '.ToString("MM/dd/YYYY")
        'Form3.MaskedTextBox6.Text = Me.DataGridView1.Item(8, i).Value '.ToString("MM/dd/YYYY")



        Form3.Show()
        Me.Hide()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable

        If TextBox1.Text = "" Then
            MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Add New User!")
        Else

            cmd.Parameters.Add(New SqlParameter("@winid", TextBox1.Text))
            cmd.CommandText = "SELECT USERID, CREATED_BY, CREATED_DATE FROM BKTIME_USERS WHERE USERID =@winid"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader()

            If reader.HasRows Then
                ' User already exists
                MsgBox("User Already Exist!", MsgBoxStyle.Exclamation, "Add New User!")
            Else
                reader.Close()
                ' User does not exist, add them
                Dim cmd1 As SqlCommand = New SqlCommand("INSERT INTO BKTIME_USERS (USERID, CREATED_BY, CREATED_DATE) VALUES(@winid, @user, GETDATE())", sqlConnection1)
                cmd1.Parameters.Add(New SqlParameter("@winid", TextBox1.Text.ToUpper))
                cmd1.Parameters.Add(New SqlParameter("@user", TextBox2.Text))
                cmd1.ExecuteNonQuery()
                MsgBox("Records Successfully Added!", MsgBoxStyle.Information, "Add New Customer!")
                TextBox1.Text = ""
            End If
            sqlConnection1.Close()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand

        If ComboBox2.Text = "" Then
            MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Remove User!")
        Else

            cmd.Parameters.Add(New SqlParameter("@winid", ComboBox2.Text))
            cmd.CommandText = "DELETE FROM BKTIME_USERS WHERE USERID =@winid"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()


            If ComboBox2.Text.ToUpper = "BINYOMF" Or ComboBox2.Text.ToUpper = "BAKERJ" Or ComboBox2.Text.ToUpper = "ADMIN" Then
                ' Block user from removing admin/Mainteance user
                MsgBox("You cannot Remove the Admin User!", MsgBoxStyle.Exclamation, "Remove User!")
                ComboBox2.Items.Clear()
            Else
                ' Remove user
                Dim mResult
                mResult = MsgBox("Are you sure you want to remove the user " + ComboBox2.Text + "?", vbYesNo + vbQuestion, "Removal Confirmation")
                If mResult = vbNo Then
                    ComboBox2.Items.Clear()
                    Exit Sub
                Else
                    cmd.ExecuteNonQuery()
                    MsgBox("Records Successfully Removed!", MsgBoxStyle.Information, "Remove User!")
                    ComboBox2.Items.Clear()
                End If
            End If
            sqlConnection1.Close()
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        DataGridView1.DataSource = Nothing
        MaskedTextBox1.Clear()
        MaskedTextBox2.Clear()
        MaskedTextBox3.Clear()
        MaskedTextBox4.Clear()

        If ComboBox1.Text = "" Then
            Exit Sub
        Else

            cmd.Parameters.Add(New SqlParameter("@winid", ComboBox1.Text))
            cmd.CommandText = "SELECT AGENT_NAME, LOGID, MAX(DAY_DATE) FROM BKTIME_LOG WHERE AGENT_NAME =@winid GROUP BY AGENT_NAME, LOGID"
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()

            reader = cmd.ExecuteReader()

            Do While reader.Read = True


                TextBox3.Text = reader(1)

            Loop


        End If


    End Sub
    Private Sub RadioButton4_Click(sender As Object, e As EventArgs) Handles RadioButton4.Click
        Label1.Visible = False
        TextBox1.Visible = False
        Button1.Visible = False
        Label6.Visible = False
        Label6.Font = New Font(Label6.Font, FontStyle.Bold)
        Label7.Visible = False
        Label8.Visible = False
        ComboBox2.Visible = False
        Button4.Visible = False
        Label4.Visible = False
        ComboBox1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
        Label5.Visible = False
        Button3.Visible = False
        Button2.Visible = False
        DataGridView1.DataSource = Nothing

        DateTimePicker3.Visible = False
        DateTimePicker4.Visible = False
        MaskedTextBox3.Visible = False
        MaskedTextBox4.Visible = False
        Button5.Visible = False
        Label10.Visible = False

        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Button6.Visible = False
        Label11.Visible = False
        'Label12.Visible = False
        CheckBox1.Visible = False

        If TextBox2.Text.ToUpper = "BINYOMF" Or TextBox2.Text = "BAKERJ" Or ComboBox2.Text.ToUpper = "ADMIN" Then
            Label1.Visible = True
            TextBox1.Visible = True
            Button1.Visible = True
            Label6.Visible = True
            Label6.Font = New Font(Label6.Font, FontStyle.Bold)
            Label7.Visible = True
        Else
            MsgBox("You're not an Administrator, Only Administrator can Add Users!", MsgBoxStyle.Exclamation, "Add User")
            Exit Sub
        End If
    End Sub

    Private Sub RadioButton5_Click(sender As Object, e As EventArgs) Handles RadioButton5.Click
        ComboBox2.Visible = False
        Button4.Visible = False
        Label1.Visible = False
        Label6.Visible = False
        Label6.Font = New Font(Label6.Font, FontStyle.Bold)
        Label7.Visible = False
        Label8.Visible = False
        TextBox1.Visible = False
        Button1.Visible = False
        Label4.Visible = False
        ComboBox1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
        Label5.Visible = False
        Button3.Visible = False
        Button2.Visible = False
        DataGridView1.DataSource = Nothing

        DateTimePicker3.Visible = False
        DateTimePicker4.Visible = False
        MaskedTextBox3.Visible = False
        MaskedTextBox4.Visible = False
        Button5.Visible = False
        Label10.Visible = False

        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Button6.Visible = False
        Label11.Visible = False
        'Label12.Visible = False
        CheckBox1.Visible = False

        If TextBox2.Text.ToUpper = "BINYOMF" Or TextBox2.Text = "BAKERJ" Or ComboBox2.Text.ToUpper = "ADMIN" Then
            ComboBox2.Visible = True
            Button4.Visible = True
            Label1.Visible = False
            Label6.Visible = True
            Label6.Font = New Font(Label6.Font, FontStyle.Bold)
            Label7.Visible = True
            Label8.Visible = True
        Else
            MsgBox("You're not an Administrator, Only Administrator can Remove Users!", MsgBoxStyle.Exclamation, "Remove User")
            Exit Sub
        End If

    End Sub

    Private Sub RadioButton1_Click(sender As Object, e As EventArgs) Handles RadioButton1.Click
        'Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        'Dim cmd As New SqlCommand
        'Dim reader As SqlDataReader
        'Dim dt As New DataTable

        'ComboBox1.Items.Clear()
        'Label2.Visible = False
        'Label3.Visible = False
        'Label4.Visible = False
        'Label5.Visible = False
        'ComboBox1.Visible = False
        'DateTimePicker1.Visible = False
        'DateTimePicker2.Visible = False
        'MaskedTextBox1.Visible = False
        'MaskedTextBox2.Visible = False
        'Button2.Visible = False
        'Button3.Visible = False
        'ComboBox2.Visible = False
        'Button4.Visible = False
        'Label8.Visible = False
        'Label6.Visible = False
        'Label7.Visible = False
        'TextBox1.Visible = False
        'Label1.Visible = False
        'Button1.Visible = False

        ''cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, LOGID, LOGIN_TIME, LOGOUT_TIME, LOGGED_IN_TIME, #CALLS_IN, CALLS_IN_TIME, #CALLS_OUT, CALLS_OUT_TIME, #CALLS_TRANSFERRED, #CALLS_ONHOLD, HOLD_CALLS_TIME FROM BKTIME_LOG ORDER BY AGENT_NAME ASC"
        'cmd.CommandText = " SELECT DAY_DATE, AGENT_NAME, LOGID, CONVERT(varchar(8),LOGIN_TIME, 108) AS LOGIN_TIME, CONVERT(varchar(8), LOGOUT_TIME, 108) AS LOGOUT_TIME, LOGGED_IN_TIME, BKTIME_ID FROM BKTIME_LOG ORDER BY AGENT_NAME ASC"
        'cmd.CommandType = CommandType.Text
        'cmd.Connection = sqlConnection1

        'sqlConnection1.Open()

        'reader = cmd.ExecuteReader()

        'DataGridView1.DataSource = Nothing


        ''Do While reader.Read = True

        'dt.Load(reader)


        'DataGridView1.DataSource = dt
        'DataGridView1.Columns("BKTIME_ID").Visible = False

        ''DataGridView1.ClearSelection()

        ''Loop

        'sqlConnection1.Close()
    End Sub

    Private Sub RadioButton2_Click(sender As Object, e As EventArgs) Handles RadioButton2.Click
        ComboBox1.Items.Clear()
        'ComboBox1.ResetText
        ComboBox1.Visible = True
        Label2.Visible = False
        Label3.Visible = False
        Label5.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        Button2.Visible = True
        Button3.Visible = False
        ComboBox2.Visible = False
        Button4.Visible = False
        DataGridView1.DataSource = Nothing
        Label8.Visible = False
        Label4.Visible = True
        Label1.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        TextBox1.Visible = False
        Button1.Visible = False

        DateTimePicker3.Visible = False
        DateTimePicker4.Visible = False
        MaskedTextBox3.Visible = False
        MaskedTextBox4.Visible = False
        Button5.Visible = False
        Label10.Visible = False

        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Button6.Visible = False
        Label11.Visible = False
        'Label12.Visible = False
        CheckBox1.Visible = False
    End Sub

    Private Sub RadioButton3_Click(sender As Object, e As EventArgs) Handles RadioButton3.Click
        ComboBox1.Visible = True
        Label2.Visible = True
        Label3.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        DateTimePicker1.Visible = True
        DateTimePicker2.Visible = True
        MaskedTextBox1.Visible = True
        MaskedTextBox2.Visible = True
        Button2.Visible = False
        Button3.Visible = True
        DataGridView1.DataSource = Nothing
        MaskedTextBox1.Clear()
        MaskedTextBox2.Clear()
        ComboBox1.Items.Clear()
        'DateTimePicker1.Format = DateTimePickerFormat.Custom
        'DateTimePicker1.CustomFormat = " "
        'DateTimePicker2.Format = DateTimePickerFormat.Custom
        'DateTimePicker2.CustomFormat = " "
        TextBox1.Visible = False
        Label8.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Button1.Visible = False
        ComboBox2.Visible = False
        Button4.Visible = False
        Label1.Visible = False

        DateTimePicker3.Visible = False
        DateTimePicker4.Visible = False
        MaskedTextBox3.Visible = False
        MaskedTextBox4.Visible = False
        Button5.Visible = False
        Label10.Visible = False

        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Button6.Visible = False
        Label11.Visible = False
        'Label12.Visible = False
        CheckBox1.Visible = False
       
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        'Label9.Font = New Font(Label6.Font, FontStyle.Bold)
        'TextBox3.Text = Format(Now, "yyyy-MM-dd hh:mm:ss")
        Label9.Text = Format(Now, "MM/dd/yyyy hh:mm:ss")

    End Sub
    Private Sub RadioButton6_Click(sender As Object, e As EventArgs) Handles RadioButton6.Click
        Label2.Visible = False
        Label3.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        ComboBox1.Visible = False
        ComboBox2.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        TextBox2.Visible = False
        Label1.Visible = False
        TextBox1.Visible = False
        Button1.Visible = False
        Button4.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        Label8.Visible = False
        Form4.Show()
        Me.Hide()

        DateTimePicker3.Visible = False
        DateTimePicker4.Visible = False
        MaskedTextBox3.Visible = False
        MaskedTextBox4.Visible = False
        Button5.Visible = False
        Label10.Visible = False

        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Button6.Visible = False
        Label11.Visible = False
        'Label12.Visible = False
        CheckBox1.Visible = False
    End Sub

    Private Sub RadioButton7_Click(sender As Object, e As EventArgs) Handles RadioButton7.Click
        Label1.Visible = False
        TextBox1.Visible = False
        Button1.Visible = False
        Label6.Visible = False
        Label6.Font = New Font(Label6.Font, FontStyle.Bold)
        Label7.Visible = False
        Label8.Visible = False
        ComboBox2.Visible = False
        Button4.Visible = False
        Label4.Visible = True
        ComboBox1.Visible = True
        Label2.Visible = False
        Label3.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
        Label5.Visible = False
        Button3.Visible = False
        Button2.Visible = False
        DataGridView1.DataSource = Nothing
        Label1.Visible = False
        TextBox1.Visible = False
        Button1.Visible = False
        Label6.Visible = False
        Label6.Font = New Font(Label6.Font, FontStyle.Bold)

        DateTimePicker3.Visible = True
        DateTimePicker4.Visible = True
        MaskedTextBox3.Visible = True
        MaskedTextBox4.Visible = True
        Button5.Visible = True
        Label10.Visible = True
        Label5.Visible = True

        ComboBox3.Visible = False
        ComboBox4.Visible = False
        Button6.Visible = False
        Label11.Visible = False
        'Label12.Visible = False
        CheckBox1.Visible = False


    End Sub

    Private Sub DateTimePicker3_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker3.ValueChanged
        MaskedTextBox3.Text = DateTimePicker3.Value.ToString("MM/dd/yyyy hh:mm")
    End Sub

    Private Sub DateTimePicker4_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker4.ValueChanged
        MaskedTextBox4.Text = DateTimePicker4.Value.ToString("MM/dd/yyyy hh:mm")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        'Dim reader As SqlDataReader
        Dim dt As New DataTable

        If ComboBox1.Text = "" Or Me.MaskedTextBox3.MaskCompleted = False Or Me.MaskedTextBox4.MaskCompleted = False Then
            MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Add New Login!")
            MaskedTextBox3.Clear()
            MaskedTextBox4.Clear()
            ComboBox1.Items.Clear()
        Else
            Dim mResult
            mResult = MsgBox("Are you sure you want to add a new login for " + ComboBox1.Text + "?", vbYesNo + vbQuestion, "Add Confirmation")
            If mResult = vbNo Then
                ComboBox1.Items.Clear()
                MaskedTextBox3.Text = ""
                MaskedTextBox4.Text = ""
                Exit Sub
            Else
                sqlConnection1.Open()
                ' User does not exist, add them
                Dim cmd1 As SqlCommand = New SqlCommand("INSERT INTO BKTIME_LOG (AGENT_NAME, LOGID, LOGIN_TIME, LOGOUT_TIME, UPDATED_BY, UPDATED_DATE, DAY_DATE) VALUES(@par, @par3, @par1, @par2, @user, GETDATE(), CAST(@par1 AS DATE))", sqlConnection1)
                cmd1.Parameters.Add(New SqlParameter("@par", ComboBox1.Text))
                cmd1.Parameters.Add(New SqlParameter("@par1", MaskedTextBox3.Text))
                cmd1.Parameters.Add(New SqlParameter("@par2", MaskedTextBox4.Text))
                cmd1.Parameters.Add(New SqlParameter("@par3", TextBox3.Text))
                cmd1.Parameters.Add(New SqlParameter("@user", TextBox2.Text))
                cmd1.ExecuteNonQuery()
                MsgBox("Records Successfully Added!", MsgBoxStyle.Information, "Add New Agent Login!")
            End If
        End If
        sqlConnection1.Close()
        ComboBox1.Items.Clear()
        MaskedTextBox3.Text = ""
        MaskedTextBox4.Text = ""
    End Sub

    Private Sub ComboBox3_DropDown(sender As Object, e As EventArgs) Handles ComboBox3.DropDown


        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader

        cmd.CommandText = "SELECT DISTINCT MANAGER_NAME FROM BKTIME_LOG ORDER BY MANAGER_NAME ASC"
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

    'Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
    '    Application.Close()
    'End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader
        Dim dt As New DataTable

        'If ComboBox3.Text = "" Or ComboBox4.Text = "" Then
        If String.IsNullOrEmpty(ComboBox3.Text.ToString) Then
            MsgBox("Please fill-up all fields!", MsgBoxStyle.Exclamation, "Select Team!")
            'Exit Sub
            ComboBox3.Items.Clear()
            ComboBox4.Items.Clear()
        ElseIf String.IsNullOrEmpty(ComboBox4.Text.ToString) And String.IsNullOrEmpty(ComboBox3.Text.ToString) = False Then
            sqlConnection1.Open()
            ' User does not exist, add them
            Dim cmd1 As SqlCommand = New SqlCommand("SELECT DAY_DATE, AGENT_NAME, CONVERT(varchar(8),LOGIN_TIME, 108) AS LOGIN_TIME, CONVERT(varchar(8), LOGOUT_TIME, 108) AS LOGOUT_TIME, DATEDIFF(s, LOGIN_TIME,LOGOUT_TIME) AS LOGGED_IN_TIME, BKTIME_ID FROM BKTIME_LOG WHERE MANAGER_NAME=@winid ORDER BY DAY_DATE DESC, AGENT_NAME ASC, LOGIN_TIME ASC, LOGOUT_TIME ASC", sqlConnection1)
            cmd1.Parameters.Add(New SqlParameter("@winid", ComboBox3.Text))
            cmd1.Parameters.Add(New SqlParameter("@winid2", ComboBox4.Text))
            reader = cmd1.ExecuteReader
            DataGridView1.DataSource = Nothing
            dt.Load(reader)


            DataGridView1.DataSource = dt
            DataGridView1.Columns("BKTIME_ID").Visible = False

            DataGridView1.ClearSelection()

        ElseIf String.IsNullOrEmpty(ComboBox4.Text.ToString) = False And String.IsNullOrEmpty(ComboBox3.Text.ToString) = False Then
            sqlConnection1.Open()
            ' User does not exist, add them
            Dim cmd2 As SqlCommand = New SqlCommand("SELECT DAY_DATE, AGENT_NAME, CONVERT(varchar(8),LOGIN_TIME, 108) AS LOGIN_TIME, CONVERT(varchar(8), LOGOUT_TIME, 108) AS LOGOUT_TIME, DATEDIFF(s, LOGIN_TIME,LOGOUT_TIME) AS LOGGED_IN_TIME, BKTIME_ID FROM BKTIME_LOG WHERE MANAGER_NAME=@winid AND AGENT_NAME=@winid2 ORDER BY DAY_DATE DESC, AGENT_NAME ASC, LOGIN_TIME ASC, LOGOUT_TIME ASC", sqlConnection1)
            cmd2.Parameters.Add(New SqlParameter("@winid", ComboBox3.Text))
            cmd2.Parameters.Add(New SqlParameter("@winid2", ComboBox4.Text))
            reader = cmd2.ExecuteReader
            DataGridView1.DataSource = Nothing
            dt.Load(reader)


            DataGridView1.DataSource = dt
            DataGridView1.Columns("BKTIME_ID").Visible = False

            DataGridView1.ClearSelection()

        End If

        sqlConnection1.Close()
    End Sub

    Private Sub ComboBox4_DropDown(sender As Object, e As EventArgs) Handles ComboBox4.DropDown
        Dim sqlConnection1 As New SqlConnection("Data Source=COWVSQLD001;Initial Catalog=Bankruptcy;User ID =RecourseUser; Password=Recourse#123;")
        Dim cmd As New SqlCommand
        Dim reader As SqlDataReader

        'ComboBox3.Visible = False
        'ComboBox4.Visible = False
        'Button6.Visible = False

        If Me.ComboBox3.Text = "" Then
            MsgBox("Select a Manager First!", MsgBoxStyle.Exclamation, "Select Team!")
            ComboBox4.DropDownStyle = ComboBoxStyle.Simple
            ComboBox4.DropDownStyle = ComboBoxStyle.DropDownList
            ComboBox4.Items.Clear()
            ComboBox3.Items.Clear()
            Exit Sub

        Else
            cmd.CommandText = "SELECT DISTINCT AGENT_NAME FROM BKTIME_LOG WHERE MANAGER_NAME =@winid ORDER BY AGENT_NAME ASC"
            cmd.Parameters.Add(New SqlParameter("@winid", ComboBox3.Text))
            cmd.CommandType = CommandType.Text
            cmd.Connection = sqlConnection1

            sqlConnection1.Open()
            'Console.WriteLine("Connection Opened")

            reader = cmd.ExecuteReader()
            ' Data is accessible through the DataReader object here.

            'ComboBox1.AutoCompleteSource = AutoCompleteSource.ListItems
            'ComboBox1.AutoCompleteMode = AutoCompleteMode.SuggestAppend

            'Fill a combo box with the datareader
            ComboBox4.Items.Clear()
            Do While reader.Read = True


                ComboBox4.Items.Add(reader.GetString(0))

            Loop

            sqlConnection1.Close()
        End If
    End Sub

    Private Sub RadioButton8_Click(sender As Object, e As EventArgs) Handles RadioButton8.Click
        ComboBox1.Items.Clear()
        'ComboBox1.ResetText
        ComboBox1.Visible = False
        Label2.Visible = False
        Label3.Visible = False
        Label5.Visible = False
        DateTimePicker1.Visible = False
        DateTimePicker2.Visible = False
        MaskedTextBox1.Visible = False
        MaskedTextBox2.Visible = False
        Button2.Visible = False
        Button3.Visible = False
        ComboBox2.Visible = False
        Button4.Visible = False
        DataGridView1.DataSource = Nothing
        Label8.Visible = False
        Label4.Visible = False
        Label1.Visible = False
        Label6.Visible = False
        Label7.Visible = False
        TextBox1.Visible = False
        Button1.Visible = False

        DateTimePicker3.Visible = False
        DateTimePicker4.Visible = False
        MaskedTextBox3.Visible = False
        MaskedTextBox4.Visible = False
        Button5.Visible = False
        Label10.Visible = False

        ComboBox3.Visible = True
        ComboBox4.Visible = True
        ComboBox4.Enabled = False
        Button6.Visible = True
        Label11.Visible = True
        'Label12.Visible = True
        CheckBox1.Visible = True
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        ComboBox4.Items.Clear()
        CheckBox1.Checked = False
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.CheckState = CheckState.Checked Then
            ComboBox4.Enabled = True
        ElseIf CheckBox1.CheckState = CheckState.Unchecked Then
            ComboBox4.Items.Clear()
            ComboBox4.Enabled = False
        End If
    End Sub
End Class

