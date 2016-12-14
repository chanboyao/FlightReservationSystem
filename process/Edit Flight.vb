Imports System.Data.OleDb
Public Class Edit_Flight

    Dim conn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Private strday As String
    Private strmonth As String
    Private stryear As String

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        Dim flightdate As Date

        cmd = New OleDbCommand("Select FlightModel , Year , StfID from flight where flightcode = '" & txtFlightCode.Text & "'", conn)
        dr = cmd.ExecuteReader

        If dr.Read Then
            txtFlightModel.Text = dr.GetString(0)
            flightdate = dr.GetDateTime(1)
            txtStaffID.Text = dr.GetString(2)
            strday = flightdate.Day
            strmonth = flightdate.Month
            stryear = flightdate.Year
            If strmonth.Count = 1 And strday.Count = 1 Then
                mskTxtDate.Text = "0" & strday & "0" & strmonth & stryear
            ElseIf strmonth.Count = 1 Then
                mskTxtDate.Text = strday & "0" & strmonth & stryear
            ElseIf strday.Count = 1 Then
                mskTxtDate.Text = "0" & strday & strmonth & stryear
            Else
                mskTxtDate.Text = strday & strmonth & stryear
            End If


        Else
            MessageBox.Show("The flight code doesn't exist.Please make sure you have enter a correct flight code.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If


        mskTxtDate.Show()
        cmbDay.Hide()
        cmbYear.Hide()
        cmbMth.Hide()
        lblSymbol1.Hide()
        lblSymbol2.Hide()

        btnEdit.Enabled = True

    End Sub

    Private Sub Edit_Flight_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        conn.Close()
    End Sub

    Private Sub Edit_Flight_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim connstring As String = ""
        connstring = strConnection
        conn = New OleDbConnection(connstring)
        conn.Open()

    End Sub

    Private Sub btnEdit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEdit.Click
        Dim cmdUpdate As OleDbCommand
        Dim flightdate As Date
        Dim date1 As String
        Dim msgResult As Long

        If btnEdit.Text = "Edit" Then
            btnEdit.Text = "Save"
            txtFlightModel.Enabled = True
            mskTxtDate.Enabled = True
            cmbDay.SelectedIndex = CInt(strday - 1)
            cmbMth.SelectedIndex = CInt(strmonth - 1)
            cmbYear.SelectedItem = stryear
            btnCancel.Visible = True
            mskTxtDate.Hide()
            cmbDay.Show()
            cmbMth.Show()
            cmbYear.Show()
            lblSymbol1.Show()
            lblSymbol2.Show()

        ElseIf btnEdit.Text = "Save" Then
            btnEdit.Text = "Edit"
            txtFlightModel.Enabled = False
            txtStaffID.Enabled = False
            mskTxtDate.Enabled = False
            date1 = cmbDay.SelectedItem & "/" & cmbMth.SelectedItem & "/" & cmbYear.SelectedItem
            flightdate = CDate(date1)
            msgResult = MessageBox.Show("Are your sure you want to update?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If msgResult = DialogResult.Yes Then
                cmdUpdate = New OleDbCommand("Update Flight set FlightModel ='" & txtFlightModel.Text & "',[year] = '" & flightdate & "', StfID = '" & LoginID & "' where flightcode ='" & txtFlightCode.Text & "'", conn)
                cmdUpdate.ExecuteNonQuery()
                MessageBox.Show("Update Successful!")
            ElseIf msgResult = DialogResult.No Then
                txtFlightCode.Focus()
            End If

            btnCancel.Visible = False
            btnEdit.Enabled = False
            mskTxtDate.Show()
            cmbDay.Hide()
            cmbMth.Hide()
            cmbYear.Hide()
            lblSymbol1.Hide()
            lblSymbol2.Hide()
        End If

    End Sub

    Private Sub cmbMth_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbMth.LostFocus
        Dim intDayLength As Integer = cmbDay.Items.Count

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
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        txtFlightCode.Clear()
        mskTxtDate.Clear()
        mskTxtDate.Show()
        cmbDay.Hide()
        cmbMth.Hide()
        cmbYear.Hide()
        lblSymbol1.Hide()
        lblSymbol2.Hide()
        txtFlightModel.Clear()
        txtStaffID.Clear()
        txtFlightCode.Focus()
        btnEdit.Text = "Edit"
        btnEdit.Enabled = False
        btnCancel.Hide()
        mskTxtDate.Enabled = False
        txtFlightModel.Enabled = False
    End Sub

    Private Sub btnShow_MouseDown(sender As Object, e As MouseEventArgs) Handles btnShow.MouseDown
        Dim myfont As New Font(btnShow.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnShow.BackgroundImage = My.Resources.button2Normal
        btnShow.Font = myfont
    End Sub

    Private Sub btnShow_MouseEnter(sender As Object, e As EventArgs) Handles btnShow.MouseEnter
        Dim myfont As New Font(btnShow.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnShow.BackgroundImage = My.Resources.button2normalBlackg
        btnShow.Font = myfont
    End Sub

    Private Sub btnShow_MouseLeave(sender As Object, e As EventArgs) Handles btnShow.MouseLeave
        Dim myfont As New Font(btnShow.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnShow.BackgroundImage = My.Resources.button2normalDark
        btnShow.Font = myfont
    End Sub

    Private Sub btnShow_MouseUp(sender As Object, e As MouseEventArgs) Handles btnShow.MouseUp
        btnShow.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnEdit_MouseDown(sender As Object, e As MouseEventArgs) Handles btnEdit.MouseDown
        Dim myfont As New Font(btnEdit.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnEdit.BackgroundImage = My.Resources.button2Normal
        btnEdit.Font = myfont
    End Sub

    Private Sub btnEdit_MouseEnter(sender As Object, e As EventArgs) Handles btnEdit.MouseEnter
        Dim myfont As New Font(btnEdit.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnEdit.BackgroundImage = My.Resources.button2normalBlackg
        btnEdit.Font = myfont
    End Sub

    Private Sub btnEdit_MouseLeave(sender As Object, e As EventArgs) Handles btnEdit.MouseLeave
        Dim myfont As New Font(btnEdit.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnEdit.BackgroundImage = My.Resources.button2normalDark
        btnEdit.Font = myfont
    End Sub

    Private Sub btnEdit_MouseUp(sender As Object, e As MouseEventArgs) Handles btnEdit.MouseUp
        btnEdit.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnCancel_MouseDown(sender As Object, e As MouseEventArgs) Handles btnCancel.MouseDown
        Dim myfont As New Font(btnCancel.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCancel.BackgroundImage = My.Resources.button2Normal
        btnCancel.Font = myfont
    End Sub

    Private Sub btnCancel_MouseEnter(sender As Object, e As EventArgs) Handles btnCancel.MouseEnter
        Dim myfont As New Font(btnCancel.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnCancel.BackgroundImage = My.Resources.button2normalBlackg
        btnCancel.Font = myfont
    End Sub

    Private Sub btnCancel_MouseLeave(sender As Object, e As EventArgs) Handles btnCancel.MouseLeave
        Dim myfont As New Font(btnCancel.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCancel.BackgroundImage = My.Resources.button2normalDark
        btnCancel.Font = myfont
    End Sub

    Private Sub btnCancel_MouseUp(sender As Object, e As MouseEventArgs) Handles btnCancel.MouseUp
        btnCancel.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

End Class