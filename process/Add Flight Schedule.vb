Imports System.Data.OleDb
Public Class Add_Flight_Schedule

    Dim conn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Dim da As OleDbDataAdapter
    Dim ds As New DataSet
    Private strDay As String
    Private strMth As String
    Private strYear As String
    Private dTime As Date
    Private aTime As Date

    Private Sub Add_Flight_Schedule_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        conn.Close()
    End Sub

    Private Sub Add_Flight_Schedule_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim connstring As String = ""
        connstring = strConnection
        conn = New OleDbConnection(connstring)
        conn.Open()
    End Sub

    Private Sub cmbDay_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbDay.LostFocus
        If cmbDay.SelectedIndex = -1 Then
            lblDate.Text = "The day of flight schedule is required."
        Else
            strDay = cmbDay.SelectedItem.ToString
            lblDate.Text = ""
        End If
    End Sub

    Private Sub cmbMth_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbMth.LostFocus
        Dim intDayLength As Integer = cmbDay.Items.Count

        If cmbMth.SelectedIndex = -1 Then
            lblDate.Text = "The month of flight schedule is required."
            Exit Sub
        Else
            lblDate.Text = ""
            strMth = cmbMth.SelectedItem.ToString
            cmbYear.Focus()
        End If

        Select Case cmbMth.SelectedIndex
            Case 0, 2, 4, 6, 7, 9, 11 '"JAN, MAR, MAY, JUL, AUG, OCT, DEC"
                If intDayLength = 28 Then
                    cmbDay.Items.Insert(28, "29")
                    cmbDay.Items.Insert(29, "30")
                    cmbDay.Items.Insert(30, "31")
                ElseIf intDayLength = 29 Then
                    cmbDay.Items.Insert(29, "30")
                    cmbDay.Items.Insert(30, "31")
                ElseIf intDayLength = 30 Then
                    cmbDay.Items.Insert(30, "31")
                End If

            Case 1 '"FEB"
                If cmbYear.SelectedIndex = -1 Then

                Else
                    If CInt(cmbYear.SelectedItem.ToString) Mod 4 = 0 Then '29th for FEB
                        If intDayLength = 31 Then
                            cmbDay.Items.RemoveAt(30) '31st
                            cmbDay.Items.RemoveAt(29) '30th
                        ElseIf intDayLength = 30 Then
                            cmbDay.Items.RemoveAt(29) '30th
                        ElseIf intDayLength = 28 Then
                            cmbDay.Items.Insert(28, "29")
                        End If
                    Else '28th for FEB
                        If intDayLength = 31 Then
                            cmbDay.Items.RemoveAt(30) '31st
                            cmbDay.Items.RemoveAt(29) '30th
                            cmbDay.Items.RemoveAt(28) '29th
                        ElseIf intDayLength = 30 Then
                            cmbDay.Items.RemoveAt(29) '30th
                            cmbDay.Items.RemoveAt(28) '29th
                        ElseIf intDayLength = 29 Then
                            cmbDay.Items.RemoveAt(28) '29th
                        End If
                    End If

                End If

            Case 3, 5, 8, 10 '"APR, JUN, SEP, NOV"
                If intDayLength = 28 Then
                    cmbDay.Items.Insert(28, "29")
                    cmbDay.Items.Insert(29, "30")
                    cmbDay.Items.Insert(30, "31")
                ElseIf intDayLength = 29 Then
                    cmbDay.Items.Insert(29, "30")
                    cmbDay.Items.Insert(30, "31")
                ElseIf intDayLength = 31 Then
                    cmbDay.Items.RemoveAt(30)
                End If

        End Select

        cmbDay.Focus()
    End Sub

    Private Sub cmbYear_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbYear.LostFocus
        If cmbYear.SelectedIndex = -1 Then
            lblDate.Text = "The year of flight schedule is required."
            Exit Sub
        Else
            lblDate.Text = ""
            strYear = (cmbYear.SelectedIndex).ToString
        End If
    End Sub

    Private Sub txtScheduleCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtScheduleCode.LostFocus
        Dim conn As OleDbConnection
        Dim cmdRead As OleDbCommand


        conn = New OleDbConnection(strConnection)
        conn.Open()

        cmdRead = New OleDbCommand("Select * from FlightSchedule", conn)

        dr = cmdRead.ExecuteReader
        While dr.Read
            If dr.GetString(0) = txtScheduleCode.Text Then
                lblScheduleCode.Text = "The schedule code is already exist.Please enter a new schedule code."
                Exit While
            ElseIf txtScheduleCode.Text.Chars(0) <> "A" Then
                lblScheduleCode.Text = "The initial character of schedule code must be A and followed by six numbers."
            ElseIf txtScheduleCode.Text = "" Then
                lblScheduleCode.Text = "Schedule code can't leave blank."
            Else
                lblScheduleCode.Text = ""
            End If
        End While

        conn.Close()
    End Sub

    Private Sub mskDepartTime_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskDepartTime.LostFocus
        Dim split As String() = mskDepartTime.Text.Split(New [Char]() {":"})
        Dim hour As String
        Dim min As String
        hour = split(0)
        min = split(1)

        If mskDepartTime.Text = "  :" Then
            lblDepartTime.Text = "The time can't leave blank!"
        ElseIf CInt(hour) < 1 Or CInt(hour) >= 13 Or CInt(min) < 0 Or CInt(min) > 59 Then
            lblDepartTime.Text = "Invalid Time!Please key in a proper time in 12 hour format."
        Else
            lblDepartTime.Text = ""
        End If
    End Sub

    Private Sub cmbDepart_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbDepart.LostFocus
        If cmbDepart.SelectedIndex = -1 Then
            lblDepartTime.Text = "Please choose whether the time is in AM or PM!"
        Else
            lblDepartTime.Text = ""
        End If
    End Sub

    Private Sub mskArriveTime_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles mskArriveTime.LostFocus
        Dim split As String() = mskArriveTime.Text.Split(New [Char]() {":"})
        Dim hour As String
        Dim min As String
        hour = split(0)
        min = split(1)

        If mskArriveTime.Text = "  :" Then
            lblArriveTime.Text = "The time can't leave blank!"
        ElseIf CInt(hour) < 1 Or CInt(hour) >= 13 Or CInt(min) < 0 Or CInt(min) > 59 Then
            lblArriveTime.Text = "Invalid Time!Please key in a proper time in 12 hour format."
        Else
            lblArriveTime.Text = ""
        End If
    End Sub

    Private Sub cmbArrive_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbArrive.LostFocus
        If cmbArrive.SelectedIndex = -1 Then
            lblArriveTime.Text = "Please choose whether the time is in AM or PM!"
        Else
            lblArriveTime.Text = ""
        End If
    End Sub

    Private Sub txtFlightCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtFlightCode.LostFocus
        If txtFlightCode.Text.Chars(0) <> "A" Then
            lblFlightCode.Text = "The Initial character of flight code must be A and followed by two numbers."
        ElseIf txtFlightCode.Text = "" Then
            lblFlightCode.Text = "Flight Code can't leave blank!"
        Else
            lblFlightCode.Text = ""
        End If
    End Sub

    Private Sub txtLocationCode_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtLocationCode.LostFocus
        Dim conn As OleDbConnection
        Dim cmdRead As OleDbCommand


        conn = New OleDbConnection(strConnection)
        conn.Open()

        cmdRead = New OleDbCommand("Select * from FlightSchedule", conn)

        dr = cmdRead.ExecuteReader
        While dr.Read
            If txtLocationCode.Text.Chars(0) <> "L" Then
                lblLocationCode.Text = "The Initial character of location code must be S and followed by two numbers."
            ElseIf txtLocationCode.Text = "" Then
                lblLocationCode.Text = "Location code can't leave blank."
            Else
                lblLocationCode.Text = ""
            End If
        End While

        conn.Close()
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        txtFlightCode.Clear()
        txtLocationCode.Clear()
        txtScheduleCode.Clear()
        txtScheduleCode.Focus()
        cmbArrive.SelectedIndex = -1
        cmbDay.SelectedIndex = -1
        cmbDepart.SelectedIndex = -1
        cmbMth.SelectedIndex = -1
        cmbYear.SelectedIndex = -1
        mskArriveTime.Clear()
        mskDepartTime.Clear()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Dim cmdAdd As OleDbCommand
        Dim sDate As Date
        Dim strDate As String
        Dim msgResult As Long
        Dim strTime1 As String
        Dim strTime2 As String
        Dim dTime As Date
        Dim aTime As Date
        Dim hour1 As String
        Dim hour2 As String
        Dim min1 As String
        Dim min2 As String
        Dim split1 As String() = mskDepartTime.Text.Split(New [Char]() {":"})
        Dim split2 As String() = mskArriveTime.Text.Split(New [Char]() {":"})
        Dim conn As OleDbConnection
        Dim cmdRead As OleDbCommand

        conn = New OleDbConnection(strConnection)
        conn.Open()

        cmdRead = New OleDbCommand("Select * from FlightSchedule", conn)

        strDate = cmbDay.SelectedItem & "/" & cmbMth.SelectedItem & "/" & cmbYear.SelectedItem
        sDate = CDate(strDate)

        hour1 = split1(0)
        min1 = split1(1)
        strTime1 = hour1 & ":" & min1 & cmbDepart.SelectedItem
        dTime = CDate(strTime1)

        hour2 = split2(0)
        min2 = split2(1)
        strTime2 = hour2 & ":" & min2 & cmbArrive.SelectedItem
        aTime = CDate(strTime2)

        dr = cmdRead.ExecuteReader

        While dr.Read
            If dr.GetDateTime(1).ToString("dd/\MM/\yyyy") = sDate.ToString("dd/\MM/\yyyy") Then
                If dr.GetDateTime(2).ToString("HHmm") = dTime.ToString("HHmm") Then
                    If dr.GetString(4) = txtFlightCode.Text Then
                        MessageBox.Show("The flight has be used in the same day and same depart time.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Exit While
                    Else
                        cmdAdd = New OleDbCommand("Insert into FlightSchedule(FSchedule,SCDate,SCDepartTime,SCArriveTime,FlightCode,LocationCode,StfID) values('" & txtScheduleCode.Text & "','" & sDate & "','" & dTime & "','" & aTime & "','" & txtFlightCode.Text & "','" & txtLocationCode.Text & "','" & LoginID & "')", conn)
                        msgResult = MessageBox.Show("Pleace check the following data:" & vbCrLf & "Location Code: " & txtLocationCode.Text & vbCrLf & "Date: " & sDate & vbCrLf & "Depart Time: " & dTime & vbCrLf & "Arrive Time: " & aTime & vbCrLf & "Flight Code: " & txtFlightCode.Text & vbCrLf & "Location Code: " & txtLocationCode.Text, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                        If msgResult = DialogResult.Yes Then
                            cmdAdd.ExecuteNonQuery()
                            MessageBox.Show("The new schedule is successfully added!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                            txtScheduleCode.Focus()
                            Exit While
                        ElseIf msgResult = DialogResult.No Then
                            Exit Sub
                        End If
                    End If
                Else
                    cmdAdd = New OleDbCommand("Insert into FlightSchedule(FSchedule,SCDate,SCDepartTime,SCArriveTime,FlightCode,LocationCode,StfID) values('" & txtScheduleCode.Text & "','" & sDate & "','" & dTime & "','" & aTime & "','" & txtFlightCode.Text & "','" & txtLocationCode.Text & "','" & LoginID & "')", conn)
                    msgResult = MessageBox.Show("Pleace check the following data:" & vbCrLf & "Location Code: " & txtLocationCode.Text & vbCrLf & "Date: " & sDate & vbCrLf & "Depart Time: " & dTime & vbCrLf & "Arrive Time: " & aTime & vbCrLf & "Flight Code: " & txtFlightCode.Text & vbCrLf & "Location Code: " & txtLocationCode.Text, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                    If msgResult = DialogResult.Yes Then
                        cmdAdd.ExecuteNonQuery()
                        MessageBox.Show("The new schedule is successfully added!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        txtScheduleCode.Focus()
                        Exit While
                    ElseIf msgResult = DialogResult.No Then
                        Exit Sub
                    End If
                End If
            Else
                cmdAdd = New OleDbCommand("Insert into FlightSchedule(FSchedule,SCDate,SCDepartTime,SCArriveTime,FlightCode,LocationCode,StfID) values('" & txtScheduleCode.Text & "','" & sDate & "','" & dTime & "','" & aTime & "','" & txtFlightCode.Text & "','" & txtLocationCode.Text & "','" & LoginID & "')", conn)
                msgResult = MessageBox.Show("Pleace check the following data:" & vbCrLf & "Location Code: " & txtLocationCode.Text & vbCrLf & "Date: " & sDate & vbCrLf & "Depart Time: " & dTime & vbCrLf & "Arrive Time: " & aTime & vbCrLf & "Flight Code: " & txtFlightCode.Text & vbCrLf & "Location Code: " & txtLocationCode.Text, "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
                If msgResult = DialogResult.Yes Then
                    cmdAdd.ExecuteNonQuery()
                    MessageBox.Show("The new schedule is successfully added!", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    txtScheduleCode.Focus()
                    Exit While
                ElseIf msgResult = DialogResult.No Then
                    Exit Sub
                End If
            End If
        End While

        conn.Close()
    End Sub

    Private Sub btnAdd_MouseDown(sender As Object, e As MouseEventArgs) Handles btnAdd.MouseDown
        Dim myfont As New Font(btnAdd.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnAdd.BackgroundImage = My.Resources.button2Normal
        btnAdd.Font = myfont
    End Sub

    Private Sub btnAdd_MouseEnter(sender As Object, e As EventArgs) Handles btnAdd.MouseEnter
        Dim myfont As New Font(btnAdd.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnAdd.BackgroundImage = My.Resources.button2normalBlackg
        btnAdd.Font = myfont
    End Sub

    Private Sub btnAdd_MouseUp(sender As Object, e As MouseEventArgs) Handles btnAdd.MouseUp
        btnAdd.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnAdd_MouseLeave(sender As Object, e As EventArgs) Handles btnAdd.MouseLeave
        Dim myfont As New Font(btnAdd.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnAdd.BackgroundImage = My.Resources.button2normalDark
        btnAdd.Font = myfont
    End Sub

    Private Sub btnClear_MouseDown(sender As Object, e As MouseEventArgs) Handles btnClear.MouseDown
        Dim myfont As New Font(btnClear.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnClear.BackgroundImage = My.Resources.button2Normal
        btnClear.Font = myfont
    End Sub

    Private Sub btnClear_MouseEnter(sender As Object, e As EventArgs) Handles btnClear.MouseEnter
        Dim myfont As New Font(btnClear.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnClear.BackgroundImage = My.Resources.button2normalBlackg
        btnClear.Font = myfont
    End Sub

    Private Sub btnClear_MouseUp(sender As Object, e As MouseEventArgs) Handles btnClear.MouseUp
        btnClear.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnClear_MouseLeave(sender As Object, e As EventArgs) Handles btnClear.MouseLeave
        Dim myfont As New Font(btnClear.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnClear.BackgroundImage = My.Resources.button2normalDark
        btnClear.Font = myfont
    End Sub

End Class