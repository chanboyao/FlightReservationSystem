Imports System.Data.OleDb
Public Class Cancellation
    Dim conn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Dim da As OleDbDataAdapter
    Dim ds As New DataSet
    Dim strReservationID As String
    Dim strSchedule As String
    Dim strLocationCode As String
    Dim strSeatCode(5) As String
    Dim strSeatsReserved(5) As String
    Dim dblTotal As Double
    Dim dblCharges As Double
    Dim dblRefunable As Double
    Const dblCharge As Double = 0.1
    Dim intNumOfSeats As Integer = 0
    Dim blnFoundID As Boolean = False

    Private Sub Cancellation_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim constring As String
        constring = strConnection
        conn = New OleDbConnection(constring)
        conn.Open()
    End Sub

    Private Function clearReservatioID()
        txtFlightReCode.Text = ""
        txtFlightReCode.Focus()
        btnCancelRe.Enabled = False
        Return Nothing
    End Function

    Private Function showAllLabel()
        lblDestination.Show()
        lblFrom.Show()
        lblTo.Show()
        lblArrival.Show()
        lblDate.Show()
        lblGetDate.Show()
        lblDepartureTime.Show()
        lblDTime.Show()
        lblArrivalTime.Show()
        lblATime.Show()
        lblSeatsReserved.Show()
        lstSeatsReserved.Show()

        Return Nothing
    End Function

    Private Function hideAllLabel()
        lblDestination.Hide()
        lblFrom.Hide()
        lblTo.Hide()
        lblArrival.Hide()
        lblDate.Hide()
        lblGetDate.Hide()
        lblDepartureTime.Hide()
        lblDTime.Hide()
        lblArrivalTime.Hide()
        lblATime.Hide()
        lblSeatsReserved.Hide()
        lstSeatsReserved.Hide()

        Return Nothing
    End Function

    Private Sub btnShow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShow.Click
        strReservationID = txtFlightReCode.Text

        If strReservationID = "" Then
            MessageBox.Show("Please fill in the reservation ID to show the details.", "Blank Reservation ID", MessageBoxButtons.OK, MessageBoxIcon.Information)
            clearReservatioID()
            hideAllLabel()
            Exit Sub
        ElseIf strReservationID.Length < 9 Then
            MessageBox.Show("Reservation ID must have at least 9 character.", "Invalid Reservation ID", MessageBoxButtons.OK, MessageBoxIcon.Information)
            clearReservatioID()
            hideAllLabel()
            Exit Sub
        ElseIf strReservationID.Length = 9 Then
            Dim sqlChkReservationCode As String = "SELECT * FROM FlightReservation WHERE ReservationID = '" & strReservationID & "'"
            cmd = New OleDbCommand(sqlChkReservationCode, conn)
            dr = cmd.ExecuteReader

            While dr.Read
                Dim strGetID = dr.GetString(4)

                If strReservationID = strGetID Then
                    strSchedule = dr.GetString(1)

                    Dim sqlSearchSchedule As String = "SELECT SCDate, SCDepartTime, SCArriveTime, LocationCode FROM(FlightSchedule) WHERE FSchedule = '" & strSchedule & "'"
                    cmd = New OleDbCommand(sqlSearchSchedule, conn)
                    dr = cmd.ExecuteReader

                    dr.Read()
                    lblGetDate.Text = dr.GetDateTime(0)
                    lblDTime.Text = dr.GetDateTime(1)
                    lblATime.Text = dr.GetDateTime(2)
                    strLocationCode = dr.GetString(3)

                    Dim sqlSearchLocation As String = "SELECT * FROM(FlightLocation) WHERE LocationCode = '" & strLocationCode & "'"
                    cmd = New OleDbCommand(sqlSearchLocation, conn)
                    dr = cmd.ExecuteReader

                    dr.Read()

                    lblFrom.Text = dr.GetString(1)
                    lblArrival.Text = dr.GetString(2)

                    blnFoundID = True
                    Exit While
                End If
            End While

            Dim i As Integer = 0

            If blnFoundID = False Then
                MessageBox.Show("Reservation ID not found." & vbCrLf & "Please re-eneter the ID.", "Non-exist Reservation ID", MessageBoxButtons.OK, MessageBoxIcon.Information)
                clearReservatioID()
                hideAllLabel()
            Else
                lstSeatsReserved.Items.Clear()

                cmd = New OleDbCommand(sqlChkReservationCode, conn)
                dr = cmd.ExecuteReader

                While dr.Read
                    strSeatCode(i) = dr.GetString(2)

                    i += 1
                End While

                Dim j As Integer

                For j = 0 To (i - 1)
                    Dim sqlGetSeatNo As String = "SELECT * FROM FlightSeatPrice WHERE SeatCode = '" & strSeatCode(j) & "'"
                    cmd = New OleDbCommand(sqlGetSeatNo, conn)
                    dr = cmd.ExecuteReader

                    dr.Read()

                    strSeatsReserved(j) = dr.GetString(3)
                    lstSeatsReserved.Items.Add(strSeatsReserved(j))

                Next

                intNumOfSeats = i
                showAllLabel()
                btnCancelRe.Enabled = True
            End If

        End If

        blnFoundID = False
    End Sub

    Private Sub btnCancelRe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelRe.Click
        If lstSeatsReserved.SelectedIndex = -1 Then
            MessageBox.Show("Please select seat(s) to be cancelled.", "Cancellation Error", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub

        Else
            Dim result As DialogResult

            result = MessageBox.Show("Are you sure to cancel the reservation?", "Reconfirmation of Reservation Cancellation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                Dim sqlDeleteReservation As String = "DELETE FROM Reservation WHERE ReservationID = '" & strReservationID & "'"
                Dim sqlChkMemOrStaff As String = "SELECT * FROM MemberReservation WHERE ReservationID = '" & strReservationID & "'"
                Dim sqlDeleteMOrSReservation As String = "DELETE FROM "


                cmd = New OleDbCommand(sqlChkMemOrStaff, conn)
                dr = cmd.ExecuteReader

                If dr.Read() Then
                    sqlDeleteMOrSReservation += "MemberReservation WHERE ReservationID = '" & strReservationID & "'"
                Else
                    sqlDeleteMOrSReservation += "StaffReservation WHERE ReservationID = '" & strReservationID & "'"
                End If

                Dim intSelected As Integer = 0

                For Each intChose As Integer In lstSeatsReserved.SelectedIndices
                    Dim sqlDeleteFlightReservation As String = "DELETE * FROM FlightReservation WHERE SeatCode = '"
                    Dim sqlGetTotal As String = "SELECT * FROM FlightSeatPrice WHERE SeatCode = '"

                    sqlDeleteFlightReservation += strSeatCode(intChose) & "'"

                    cmd = New OleDbCommand(sqlDeleteFlightReservation, conn)
                    dr = cmd.ExecuteReader
                    dr.Read()

                    sqlGetTotal += strSeatCode(intChose) + "'"

                    cmd = New OleDbCommand(sqlGetTotal, conn)
                    dr = cmd.ExecuteReader
                    dr.Read()

                    dblTotal += CDbl(dr.GetDouble(1))

                    intSelected += 1
                Next

                dblCharges = dblTotal * dblCharge
                dblRefunable = dblTotal - dblCharges

                If intSelected = intNumOfSeats Then
                    cmd = New OleDbCommand(sqlDeleteReservation, conn)
                    dr = cmd.ExecuteReader
                    dr.Read()

                    cmd = New OleDbCommand(sqlDeleteMOrSReservation, conn)
                    dr = cmd.ExecuteReader
                    dr.Read()

                    MessageBox.Show("Successfully deleted all the reservation made." & vbCrLf & vbCrLf & _
                                    "Charges(10% of you reservation) : RM " & dblCharges.ToString("F2") & vbCrLf & _
                                    "Total refundable                : RM " & dblRefunable.ToString("F2"), "Deletion Completed", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("Successfully deleted " & intSelected.ToString & " of the reservation." & vbCrLf & vbCrLf & _
                                    "Charges(10% of you reservation) : RM " & dblCharges.ToString("F2") & vbCrLf & _
                                    "Total refundable                : RM " & dblRefunable.ToString("F2"), "Deletion Completed", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If

                intSelected = 0
                dblTotal = 0
                clearReservatioID()
                hideAllLabel()

            End If

        End If

    End Sub

    Private Sub txtFlightReCode_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFlightReCode.TextChanged
        hideAllLabel()

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

    Private Sub btnCancelRe_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles btnCancelRe.MouseDown
        Dim myfont As New Font(btnCancelRe.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCancelRe.BackgroundImage = My.Resources.button2Normal
        btnCancelRe.Font = myfont
    End Sub

    Private Sub btnCancelRe_MouseEnter(sender As Object, e As EventArgs) Handles btnCancelRe.MouseEnter
        Dim myfont As New Font(btnCancelRe.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnCancelRe.BackgroundImage = My.Resources.button2normalBlackg
        btnCancelRe.Font = myfont
    End Sub

    Private Sub btnCancelRe_MouseLeave(sender As Object, e As EventArgs) Handles btnCancelRe.MouseLeave
        Dim myfont As New Font(btnCancelRe.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCancelRe.BackgroundImage = My.Resources.button2normalDark
        btnCancelRe.Font = myfont
    End Sub

    Private Sub btnCancelRe_MouseUp(sender As Object, e As MouseEventArgs) Handles btnCancelRe.MouseUp
        btnCancelRe.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub lblFlightReCode_Click(sender As Object, e As EventArgs) Handles lblFlightReCode.Click

    End Sub
End Class