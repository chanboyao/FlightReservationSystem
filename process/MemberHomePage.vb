Imports System.Data.OleDb
Public Class MemberHomePage
    Dim coon As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader

    Private Sub btnReservation_MouseHover(sender As Object, e As EventArgs) Handles btnReservation.MouseHover
        Me.BackgroundImage = My.Resources.memhomeBack1
        lblShow.Visible = False
    End Sub

    Private Sub btnShowReservationDetails_MouseHover(sender As Object, e As EventArgs) Handles btnShowReservationDetails.MouseHover
        Me.BackgroundImage = My.Resources.memback2
        lblShow.Visible = False
    End Sub

    Private Sub btnReservation_MouseDown(sender As Object, e As MouseEventArgs) Handles btnReservation.MouseDown
        Dim myfont As New Font(btnReservation.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnReservation.BackgroundImage = My.Resources.button2Normal
        btnReservation.Font = myfont
    End Sub

    Private Sub btnReservation_MouseEnter(sender As Object, e As EventArgs) Handles btnReservation.MouseEnter
        Dim myfont As New Font(btnReservation.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnReservation.BackgroundImage = My.Resources.button2normalBlackg
        btnReservation.Font = myfont
    End Sub

    Private Sub btnReservation_MouseUp(sender As Object, e As MouseEventArgs) Handles btnReservation.MouseUp
        btnReservation.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnReservation_MouseLeave(sender As Object, e As EventArgs) Handles btnReservation.MouseLeave
        Dim myfont As New Font(btnReservation.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnReservation.BackgroundImage = My.Resources.button2normalDark
        btnReservation.Font = myfont
        Me.BackgroundImage = My.Resources.memback3
        lblShow.Visible = True
    End Sub

    Private Sub btnShowReservationDetails_MouseDown(sender As Object, e As MouseEventArgs) Handles btnShowReservationDetails.MouseDown
        Dim myfont As New Font(btnShowReservationDetails.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnShowReservationDetails.BackgroundImage = My.Resources.button2Normal
        btnShowReservationDetails.Font = myfont
    End Sub

    Private Sub btnShowReservationDetails_MouseEnter(sender As Object, e As EventArgs) Handles btnShowReservationDetails.MouseEnter
        Dim myfont As New Font(btnShowReservationDetails.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnShowReservationDetails.BackgroundImage = My.Resources.button2normalBlackg
        btnShowReservationDetails.Font = myfont
    End Sub

    Private Sub btnShowReservationDetails_MouseUp(sender As Object, e As MouseEventArgs) Handles btnShowReservationDetails.MouseUp
        btnShowReservationDetails.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnShowReservationDetails_MouseLeave(sender As Object, e As EventArgs) Handles btnShowReservationDetails.MouseLeave
        Dim myfont As New Font(btnShowReservationDetails.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnShowReservationDetails.BackgroundImage = My.Resources.button2normalDark
        btnShowReservationDetails.Font = myfont
        Me.BackgroundImage = My.Resources.memback3
        lblShow.Visible = True
    End Sub

    Private Sub btnReservation_Click(sender As Object, e As EventArgs) Handles btnReservation.Click
        Reservation.Show()
    End Sub

    Private Sub btnShowReservationDetails_Click(sender As Object, e As EventArgs) Handles btnShowReservationDetails.Click
        ShowMemberFlightDetail.Show()
    End Sub

    Private Sub HomeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HomeToolStripMenuItem.Click
        Me.Refresh()
    End Sub

    Private Sub MemberHomePage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer1.Start()

    End Sub

    Private Sub AboutUsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutUsToolStripMenuItem.Click
        AboutUs.Show()
    End Sub

    Private Sub HelpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HelpToolStripMenuItem.Click

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        lblShow.Text = Today.ToString("dd/MM/yyyy")
        lblShow.Text &= "  " & TimeOfDay.ToString("h:mm:ss tt")
    End Sub

End Class