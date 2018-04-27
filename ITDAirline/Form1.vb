Public Class Form1
    Dim cn As New OleDb.OleDbConnection
    Dim da As OleDb.OleDbDataAdapter
    Dim ds As New DataSet
    Dim Maxrecord As Integer
    Dim CurrentRow As Integer
    'create an array of datarows 
    ' so the result of select method can live in this array
    Dim FoundRows() As Data.DataRow
    'for passengerTable
    Dim da2 As OleDb.OleDbDataAdapter
    Dim ds2 As New DataSet
    Dim Maxrecord2 As Integer
    'for flightTable
    Dim da3 As OleDb.OleDbDataAdapter
    Dim ds3 As New DataSet
    Dim Maxrecord3 As Integer
    Dim rd As OleDb.OleDbDataReader
    'for bookingTable
    Dim da4 As OleDb.OleDbDataAdapter
    Dim ds4 As New DataSet
    Dim Maxrecord4 As Integer

    Dim emptySeat
    Dim occupiedSeat

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'TODO: This line of code loads data into the table.
        cn.ConnectionString = "Provider = SQLNCLI11.0;Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\ITD\Term 3\Visual Basic.Net\assignment\assignment 5\ITDAirline\ITDAirline\ITDAirlineDB.mdf;Integrated Security = SSPI;database=thingy"

        cn.Open()
        'for passengerTable
        Dim SqlStr2 As String
        SqlStr2 = "Select * from passengerTable"
        da2 = New OleDb.OleDbDataAdapter(SqlStr2, cn)
        da2.Fill(ds2, "passengerTable")
        Maxrecord2 = ds2.Tables("passengerTable").Rows.Count - 1
        'for flightTable
        Dim SqlStr3 As String
        SqlStr3 = "Select * from flightTable"
        da3 = New OleDb.OleDbDataAdapter(SqlStr3, cn)
        da3.Fill(ds3, "flightTable")
        Maxrecord3 = ds3.Tables("flightTable").Rows.Count - 1
        'for bookingTable
        Dim SqlStr4 As String
        SqlStr4 = "Select * from bookingTable"
        da4 = New OleDb.OleDbDataAdapter(SqlStr3, cn)
        da4.Fill(ds4, "bookingTable")
        Maxrecord4 = ds4.Tables("bookingTable").Rows.Count - 1

        Call DateTimeFormat()

        'COMBO BOX for passengerPassport
        ComboBox2.DataSource = ds2.Tables(0)
        ComboBox2.DisplayMember = "passportNumber"
        'COMBO BOX for trip
        ComboBox3.Items.Add("Vancouver to Victoria")
        ComboBox3.Items.Add("Victoria to Vancouver")
        'controlling tab
        TabControl1.TabPages.Remove(TabControl1.TabPages("tabpage3"))
        TabControl1.TabPages.Remove(TabControl1.TabPages("tabpage4"))
        'SeatGroup.Visible = False

        DataGridView1.DataSource = ds3.Tables(0)
    End Sub
    'change DataTimePicker ToolBox to specific format
    Private Sub DateTimeFormat()
        DateTimePicker1.Format = DateTimePickerFormat.Custom
        DateTimePicker1.CustomFormat = "MMM d,yyyy"
        'DateTimePicker2.Format = DateTimePickerFormat.Custom
        'DateTimePicker2.CustomFormat = "hh:mm tt"
        'DateTimePicker2.ShowUpDown = True
        DateTimePicker3.Format = DateTimePickerFormat.Custom
        DateTimePicker3.CustomFormat = "MMM"
        DateTimePicker3.ShowUpDown = True
        DateTimePicker4.Format = DateTimePickerFormat.Custom
        DateTimePicker4.CustomFormat = "hh:mm tt"
        DateTimePicker4.ShowUpDown = True
        DateTimePicker5.Format = DateTimePickerFormat.Custom
        DateTimePicker5.CustomFormat = "hh:mm tt"
        DateTimePicker5.ShowUpDown = True
        DateTimePicker6.Format = DateTimePickerFormat.Custom
        DateTimePicker6.CustomFormat = "hh:mm tt"
        DateTimePicker6.ShowUpDown = True
        DateTimePicker7.Format = DateTimePickerFormat.Custom
        DateTimePicker7.CustomFormat = "hh:mm tt"
        DateTimePicker7.ShowUpDown = True

    End Sub
    'verify clientID and Password by reading the info from the database 
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim SqlStr As String
        SqlStr = "Select * from securityTable"
        da = New OleDb.OleDbDataAdapter(SqlStr, cn)
        da.Fill(ds, "securityTable")
        Maxrecord = ds.Tables("securityTable").Rows.Count - 1

        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Plz Fill Client ID and Password")
            Exit Sub
        End If
        'Dim criteria As String
        'criteria = "userName = '" + TextBox1.Text + "'"
        'FoundRows = ds.Tables("securityTable").Select(criteria)

        'If FoundRows.Length = 0 Then 'if the array is empty
        '    'which it means record not found 
        '    MessageBox.Show("Enter a valid username and password !")
        '    Exit Sub
        'Else
        Dim cmd As New OleDb.OleDbCommand("select * from securityTable where userName ='" & TextBox1.Text & "' AND password='" & TextBox2.Text & "' AND userType='admin';", cn)
        rd = cmd.ExecuteReader()
        If rd.HasRows = True Then
            MsgBox("Login Successfull", MsgBoxStyle.Information, "Login Passed")
            TabControl1.TabPages.Add(TabPage3)
            TabControl1.SelectedTab = TabControl1.TabPages("Flight Booking")
            TabControl1.TabPages.Remove(TabPage1)
            TextBox1.Clear()
            TextBox2.Clear()
            Exit Sub
        End If
        Dim cmd2 As New OleDb.OleDbCommand("select * from securityTable where userName ='" & TextBox1.Text & "' AND password='" & TextBox2.Text & "' AND userType='manager';", cn)
        rd = cmd2.ExecuteReader()
        If rd.HasRows = True Then
            MsgBox("Login Successfull", MsgBoxStyle.Information, "Login Passed")
            TabControl1.TabPages.Add(TabPage3)
            TabControl1.TabPages.Add(TabPage4)
            TabControl1.SelectedTab = TabControl1.TabPages("Flight Info. Management")
            TabControl1.TabPages.Remove(TabPage1)
            'TabControl1.TabPages(3).Enabled = Enabled

            TextBox1.Clear()
            TextBox2.Clear()
            Exit Sub
        Else
            MsgBox("Invalid Account", MsgBoxStyle.Critical, "Login Failed")
            TextBox1.Clear()
            TextBox2.Clear()
            'End If
        End If
    End Sub
    'REGISTER A NEW PASSENGER
    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim Criteria2 As String
        If TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MessageBox.Show("Incomplete Data, please fill all fields")
            Exit Sub
        End If
        Criteria2 = "passengerID ='" + TextBox3.Text + "'"
        FoundRows = ds2.Tables("passengerTable").Select(Criteria2)
        If FoundRows.Length = 0 Then
            Dim NewRow As DataRow = ds2.Tables("passengerTable").NewRow()
            'the next line is needed so VB Can execute the SQL statement 
            Dim Cb As New OleDb.OleDbCommandBuilder(da2)
            NewRow.Item("passengerID") = Convert.ToInt32(TextBox3.Text)
            NewRow.Item("name") = TextBox4.Text
            NewRow.Item("lastName") = TextBox5.Text
            NewRow.Item("passportNumber") = TextBox6.Text

            ds2.Tables.Item("passengerTable").Rows.Add(NewRow)
            da2.Update(ds2, "passengerTable")
            Maxrecord2 = Maxrecord2 + 1
            MessageBox.Show("Record Added succesfully!!!")
            Call Clearboxes()
        Else
            MessageBox.Show("Duplicate Record!!!, Try another passengerID")
        End If
    End Sub

    Public Sub ShowRecord(ByVal ThisRow As Integer)
        TextBox3.Text = ds.Tables("passengerTable").Rows(ThisRow).Item(0)
        TextBox4.Text = ds.Tables("passengerTable").Rows(ThisRow).Item(1)
        TextBox5.Text = ds.Tables("passengerTable").Rows(ThisRow).Item(2)
        TextBox6.Text = ds.Tables("passengerTable").Rows(ThisRow).Item(3)
    End Sub
    Public Sub Clearboxes()
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Clear()
        TextBox6.Clear()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)
        Call Clearboxes()
    End Sub
    'cancel the login
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        TextBox1.Clear()
        TextBox2.Clear()
    End Sub
    ' save a new flight schedual
    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Dim Criteria As String
        'If ComboBox4.Text = "" Then
        '    MessageBox.Show("Incomplete Data, please fill 'From' fields")
        '    Exit Sub
        'End If
        Criteria = "month ='" + DateTimePicker3.Text + "'"
        FoundRows = ds3.Tables("flightTable").Select(Criteria)
        If FoundRows.Length = 0 Then
            Dim NewRow As DataRow = ds3.Tables("flightTable").NewRow()
            'the next line is needed so VB Can execute the SQL statement 
            Dim Cb As New OleDb.OleDbCommandBuilder(da3)

            NewRow.Item("month") = DateTimePicker3.Text
            NewRow.Item("VanVicFrist") = DateTimePicker4.Text
            NewRow.Item("VanVicSecond") = DateTimePicker5.Text
            NewRow.Item("VicVanFrist") = DateTimePicker6.Text
            NewRow.Item("VicVanSecond") = DateTimePicker7.Text

            ds3.Tables.Item("flightTable").Rows.Add(NewRow)
            da3.Update(ds3, "flightTable")
            Maxrecord3 = Maxrecord3 + 1
            MessageBox.Show("Record Added succesfully!")
        Else
            MessageBox.Show("Duplicate Record!, Try another passengerID")
        End If
    End Sub
    'reserve seat
    Private Sub seat_Click(sender As Object, e As EventArgs) Handles Button8.Click, Button7.Click, Button6.Click, Button5.Click, Button12.Click, Button11.Click
        'thisPIC now is a reference to the box, you can use .Name, etc. to get it's properties.
        Dim thisPic As Button = sender
        'As a defult all free seat is green after click and confirm it reserves
        If thisPic.BackColor = Color.Green Then
            Dim verification2 = MsgBox("This seat " & thisPic.Name & ", is free. Would you like to reserve the seat?", MsgBoxStyle.YesNoCancel)
            If verification2 = MsgBoxResult.Yes Then
                thisPic.BackColor = Color.DarkRed
                MessageBox.Show("Your seat number is : " + thisPic.Name)

                Dim NewRow As DataRow = ds4.Tables("bookingTable").NewRow() 'Sending the data do Flight Table
                Dim Cb As New OleDb.OleDbCommandBuilder(da4)
                NewRow.Item("flightID") = 0
                NewRow.Item("tripDate") = Convert.ToDateTime(DateTimePicker1.Text)
                NewRow.Item("tripTime") = ComboBox4.Text
                NewRow.Item("trip") = ComboBox3.Text
                NewRow.Item("passportNumber") = ComboBox2.Text
                NewRow.Item("seatNumber") = thisPic.Text


                ds4.Tables.Item("bookingTable").Rows.Add(NewRow)
                da4.Update(ds4, "bookingTable")
            End If
        Else

            'Cancelling the reservation
            Dim Criteria As String
            Criteria = "passportNumber='" & ComboBox2.Text & "' AND reserveTime='" & ComboBox4.Text & "' And reserveDate='" & DateTimePicker1.Text & "' And seatNumber='" & thisPic.Text & "'"

            If ComboBox3.Text <> "" Then
                Criteria = Criteria & " And trip = Vancouver To Victoria"

            Else
                Criteria = Criteria & " And trip = Victoria To Vancouver"
            End If

            'Remove from dataset
            Dim FoundRow = ds4.Tables("bookingTable").Select(Criteria)
            If FoundRow.Length > 0 Then
                Dim msgResultCancel = MsgBox("This seat is reserved. Would you like to cancel the reservation?", MsgBoxStyle.YesNo, "Attention")
                If msgResultCancel = MsgBoxResult.Yes Then
                    FoundRow(0).Delete()

                    'Remove from DB
                    Dim Sql = "DELETE FROM bookingTable WHERE " & Criteria

                    da4.DeleteCommand = cn.CreateCommand
                    da4.DeleteCommand.CommandText = Sql
                    da4.DeleteCommand.ExecuteNonQuery()

                    ds4.AcceptChanges()

                    da4.Update(ds4, "Booking")

                    MessageBox.Show("Your reservation was canceled succesfully.")
                    thisPic.BackColor = Color.Green
                End If
            Else
                MessageBox.Show("Only the owner can cancel the reservation.")
            End If
        End If
    End Sub
    'logout by admin
    Private Sub Button9_Click_1(sender As Object, e As EventArgs) Handles Button9.Click
        TabControl1.TabPages.Remove(TabPage3)
        'TabControl1.TabPages.Remove(TabPage4)
        TabControl1.TabPages.Add(TabPage1)

    End Sub
    'logout by manager
    Private Sub Button10_Click_1(sender As Object, e As EventArgs) Handles Button10.Click
        TabControl1.TabPages.Remove(TabPage3)
        TabControl1.TabPages.Remove(TabPage4)
        TabControl1.TabPages.Add(TabPage1)
    End Sub
    'booking seat
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        'MessageBox.Show("Finding seats START!")
        If ComboBox2.Text = "" Or ComboBox3.Text = "" Or ComboBox4.Text = "" Or DateTimePicker1.Checked = False Then
            MessageBox.Show("Please fill all field.")
            Exit Sub
        End If

        'SeatGroup.Visible = True

        Dim Criteria As String
        Criteria = "reserveTime='" & ComboBox4.Text & "' And reserveDate='" & DateTimePicker1.Text & "'"

        'checking combobox value
        If ComboBox3.Text <> "" Then
            Criteria = Criteria & " And trip = Vancouver To Victoria"

        Else
            Criteria = Criteria & " And trip = Victoria To Vancouver"
        End If

        Dim FoundRows = ds4.Tables("bookingTable").Select(Criteria)
        If FoundRows.Length > 0 Then 'if you select an available seat and reserve it
            For Each row In FoundRows
                If row.Item("seatNumber") = 1 Then
                    Button5.BackColor = Color.DarkRed
                End If
                If row.Item("seatNumber") = 2 Then
                    Button6.BackColor = Color.DarkRed
                End If
                If row.Item("seatNumber") = 3 Then
                    Button7.BackColor = Color.DarkRed
                End If
                If row.Item("seatNumber") = 4 Then
                    Button8.BackColor = Color.DarkRed
                End If
                If row.Item("seatNumber") = 5 Then
                    Button11.BackColor = Color.DarkRed
                End If
                If row.Item("seatNumber") = 6 Then
                    Button12.BackColor = Color.DarkRed
                End If
            Next
        End If
        MessageBox.Show("Finding seats FINISHED!")
        Exit Sub
    End Sub
    ' save a new flight schedual
    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Dim Criteria As String
        'If ComboBox4.Text = "" Then
        '    MessageBox.Show("Incomplete Data, please fill 'From' fields")
        '    Exit Sub
        'End If
        Criteria = "month ='" + DateTimePicker3.Text + "'"
        FoundRows = ds3.Tables("flightTable").Select(Criteria)
        If FoundRows.Length = 0 Then
            Dim NewRow As DataRow = ds3.Tables("flightTable").NewRow()
            'the next line is needed so VB Can execute the SQL statement 
            Dim Cb As New OleDb.OleDbCommandBuilder(da3)

            NewRow.Item("month") = DateTimePicker3.Text
            NewRow.Item("VanVicFrist") = DateTimePicker4.Text
            NewRow.Item("VanVicSecond") = DateTimePicker5.Text
            NewRow.Item("VicVanFrist") = DateTimePicker6.Text
            NewRow.Item("VicVanSecond") = DateTimePicker7.Text

            ds3.Tables.Item("flightTable").Rows.Add(NewRow)
            da3.Update(ds3, "flightTable")
            Maxrecord3 = Maxrecord3 + 1
            MessageBox.Show("Record Added succesfully!")
        Else
            MessageBox.Show("Duplicate Record!")
        End If

    End Sub
    'register new passenger
    Private Sub Button2_Click_2(sender As Object, e As EventArgs)
        Dim Criteria2 As String
        If TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MessageBox.Show("Incomplete Data, please fill all fields")
            Exit Sub
        End If
        Criteria2 = "passengerID ='" + TextBox3.Text + "'"
        FoundRows = ds2.Tables("passengerTable").Select(Criteria2)
        If FoundRows.Length = 0 Then
            Dim NewRow As DataRow = ds2.Tables("passengerTable").NewRow()
            'the next line is needed so VB Can execute the SQL statement 
            Dim Cb As New OleDb.OleDbCommandBuilder(da2)
            NewRow.Item("passengerID") = Convert.ToInt32(TextBox3.Text)
            NewRow.Item("name") = TextBox4.Text
            NewRow.Item("lastName") = TextBox5.Text
            NewRow.Item("passportNumber") = TextBox6.Text

            ds2.Tables.Item("passengerTable").Rows.Add(NewRow)
            da2.Update(ds2, "passengerTable")
            Maxrecord2 = Maxrecord2 + 1
            MessageBox.Show("Record Added succesfully!!!")
            Call Clearboxes()
        Else
            MessageBox.Show("Duplicate Record!!!, Try another passengerID")
        End If
    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        Dim Criteria2 As String
        If TextBox3.Text = "" Or TextBox4.Text = "" Or TextBox5.Text = "" Or TextBox6.Text = "" Then
            MessageBox.Show("Incomplete Data, please fill all fields")
            Exit Sub
        End If
        Criteria2 = "passengerID ='" + TextBox3.Text + "'"
        FoundRows = ds2.Tables("passengerTable").Select(Criteria2)
        If FoundRows.Length = 0 Then
            Dim NewRow As DataRow = ds2.Tables("passengerTable").NewRow()
            'the next line is needed so VB Can execute the SQL statement 
            Dim Cb As New OleDb.OleDbCommandBuilder(da2)
            NewRow.Item("passengerID") = Convert.ToInt32(TextBox3.Text)
            NewRow.Item("name") = TextBox4.Text
            NewRow.Item("lastName") = TextBox5.Text
            NewRow.Item("passportNumber") = TextBox6.Text

            ds2.Tables.Item("passengerTable").Rows.Add(NewRow)
            da2.Update(ds2, "passengerTable")
            Maxrecord2 = Maxrecord2 + 1
            MessageBox.Show("Record Added succesfully!!!")
            Call Clearboxes()
        Else
            MessageBox.Show("Duplicate Record!!!, Try another passengerID")
        End If
    End Sub
End Class
