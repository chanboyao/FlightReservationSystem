Imports System.Data.OleDb
Public Class ShowMemberFlightDetail

    Dim conn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader

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
        lblamount.Show()
        lblTotal.Show()

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
        lblamount.Hide()
        lblTotal.Hide()

        Return Nothing
    End Function

    Private Sub ShowMemberFlightDetail_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim constring As String
        Dim intcount As Integer = 0
        Dim i As Integer = 0
        constring = strConnection
        conn = New OleDbConnection(constring)
        conn.Open()

        cmd = New OleDbCommand("select count(ReservationID) from MemberReservation where MemID = '" & LoginID & "';", conn)
        dr = cmd.ExecuteReader
        dr.Read()
        intcount = dr.GetInt32(0)
        If intcount = 0 Then
            MessageBox.Show("You did not make any reservation on flight before.")
            cmbFRSV.Enabled = False
        Else
            cmbFRSV.Enabled = True
            cmd = New OleDbCommand("select ReservationID from MemberReservation where MemID='" & LoginID & "';", conn)
            dr = cmd.ExecuteReader
            Dim IDarr(intcount) As String
            While dr.Read
                IDarr(i) = dr.GetString(0)
                i += 1
            End While
            For i = 0 To intcount - 1 Step 1
                cmbFRSV.Items.Add(IDarr(i))
            Next
        End If


    End Sub

    Private Sub cmbFRSV_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFRSV.SelectedIndexChanged
        Dim locto As String = ""
        Dim departtime As DateTime
        Dim arrivetime As DateTime
        Dim datDate As DateTime
        Dim dblamount As Double = 0
        Dim flightSC As String = ""
        Dim locCode As String = ""
        Dim intcount As Integer = 0
        Dim i As Integer = 0
        i = 0

        If cmbFRSV.SelectedIndex = -1 Then
            hideAllLabel()
        Else
            cmd = New OleDbCommand("select count(ReservationID) from MemberReservation where MemID = '" & LoginID & "';", conn)
            dr = cmd.ExecuteReader
            dr.Read()
            intcount = dr.GetInt32(0)
            cmd = New OleDbCommand("select ReservationID from MemberReservation where MemID='" & LoginID & "';", conn)
            dr = cmd.ExecuteReader
                Dim IDarr(intcount - 1) As String
                While dr.Read
                    IDarr(i) = dr.GetString(0)
                    i += 1
                End While

                cmd = New OleDbCommand("select count(Seatcode) from FlightReservation where ReservationID = '" & IDarr(cmbFRSV.SelectedIndex) & "';", conn)
                dr = cmd.ExecuteReader
                dr.Read()
                intcount = dr.GetInt32(0)
                Dim seat(intcount - 1) As String
                cmd = New OleDbCommand("select Seatcode,FSchedule from FlightReservation where ReservationID = '" & IDarr(cmbFRSV.SelectedIndex) & "';", conn)
                dr = cmd.ExecuteReader
                i = 0
                While dr.Read
                    seat(i) = dr.GetString(0)
                    flightSC = dr.GetString(1)
                    i += 1
                End While
                cmd = New OleDbCommand("select SCDate,SCDepartTime,SCArriveTime,LocationCode from FlightSchedule where FSchedule = '" & flightSC & "';", conn)
                dr = cmd.ExecuteReader
                dr.Read()
                datDate = dr.GetDateTime(0)
                departtime = dr.GetDateTime(1)
                arrivetime = dr.GetDateTime(2)
                locCode = dr.GetString(3)
                cmd = New OleDbCommand("select To from FlightLocation where LocationCode = '" & locCode & "';", conn)
                dr = cmd.ExecuteReader
                dr.Read()
                locto = dr.GetString(0)
                For i = 0 To intcount - 1 Step 1
                    cmd = New OleDbCommand("select SeatNo,SeatPrice from FlightSeatPrice where SeatCode = '" & seat(i) & "';", conn)
                    dr = cmd.ExecuteReader
                    dr.Read()
                    seat(i) = dr.GetString(0)
                    dblamount += dr.GetDouble(1)
                Next

                lstSeatsReserved.Items.Clear()
                lblFrom.Text = "PENANG"
                lblArrival.Text = locto
                lblGetDate.Text = datDate
                lblDTime.Text = departtime
                lblATime.Text = arrivetime
                lblTotal.Text = "RM " & dblamount.ToString("F2")
                For i = 0 To intcount - 1 Step 1
                    lstSeatsReserved.Items.Add(seat(i))
                Next
                showAllLabel()
            End If

    End Sub

End Class