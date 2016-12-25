Public Class StaffHomePage

    Private Sub btnAddFlightDetails_Click(sender As Object, e As EventArgs) Handles btnAddFlightDetails.Click
        Me.Hide()
        Add_Flight_Detail.Show()
    End Sub

    Private Sub btnReservation_Click(sender As Object, e As EventArgs) Handles btnReservation.Click
        Me.Hide()
        Reservation.Show()
    End Sub

    Private Sub btnEditFlightDetails_Click(sender As Object, e As EventArgs) Handles btnEditFlightDetails.Click
        Me.Hide()
        Edit_Flight_Detail.Show()
    End Sub

    Private Sub btnCancellation_Click(sender As Object, e As EventArgs) Handles btnCancellation.Click
        Me.Hide()
        Cancellation.Show()
    End Sub

    Private Sub btnGenerateReport_Click(sender As Object, e As EventArgs) Handles btnGenerateReport.Click
        Me.Hide()
        FlightReport.Show()
    End Sub

    Private Sub btnAddFlightDetails_MouseDown(sender As Object, e As MouseEventArgs) Handles btnAddFlightDetails.MouseDown
        Dim myfont As New Font(btnAddFlightDetails.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnAddFlightDetails.BackgroundImage = My.Resources.button2Normal
        btnAddFlightDetails.Font = myfont
    End Sub

    Private Sub btnAddFlightDetails_MouseEnter(sender As Object, e As EventArgs) Handles btnAddFlightDetails.MouseEnter
        Dim myfont As New Font(btnAddFlightDetails.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnAddFlightDetails.BackgroundImage = My.Resources.button2normalBlackg
        btnAddFlightDetails.Font = myfont
    End Sub

    Private Sub btnAddFlightDetails_MouseUp(sender As Object, e As MouseEventArgs) Handles btnAddFlightDetails.MouseUp
        btnAddFlightDetails.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnAddFlightDetails_MouseLeave(sender As Object, e As EventArgs) Handles btnAddFlightDetails.MouseLeave
        Dim myfont As New Font(btnAddFlightDetails.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnAddFlightDetails.BackgroundImage = My.Resources.button2normalDark
        btnAddFlightDetails.Font = myfont
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
    End Sub

    Private Sub btnEditFlightDetails_MouseDown(sender As Object, e As MouseEventArgs) Handles btnEditFlightDetails.MouseDown
        Dim myfont As New Font(btnEditFlightDetails.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnEditFlightDetails.BackgroundImage = My.Resources.button2Normal
        btnEditFlightDetails.Font = myfont
    End Sub

    Private Sub btnEditFlightDetails_MouseEnter(sender As Object, e As EventArgs) Handles btnEditFlightDetails.MouseEnter
        Dim myfont As New Font(btnEditFlightDetails.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnEditFlightDetails.BackgroundImage = My.Resources.button2normalBlackg
        btnEditFlightDetails.Font = myfont
    End Sub

    Private Sub btnEditFlightDetails_MouseUp(sender As Object, e As MouseEventArgs) Handles btnEditFlightDetails.MouseUp
        btnEditFlightDetails.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnEditFlightDetails_MouseLeave(sender As Object, e As EventArgs) Handles btnEditFlightDetails.MouseLeave
        Dim myfont As New Font(btnEditFlightDetails.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnEditFlightDetails.BackgroundImage = My.Resources.button2normalDark
        btnEditFlightDetails.Font = myfont
    End Sub

    Private Sub btnCancellation_MouseDown(sender As Object, e As MouseEventArgs) Handles btnCancellation.MouseDown
        Dim myfont As New Font(btnCancellation.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCancellation.BackgroundImage = My.Resources.button2Normal
        btnCancellation.Font = myfont
    End Sub

    Private Sub btnCancellation_MouseEnter(sender As Object, e As EventArgs) Handles btnCancellation.MouseEnter
        Dim myfont As New Font(btnCancellation.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnCancellation.BackgroundImage = My.Resources.button2normalBlackg
        btnCancellation.Font = myfont
    End Sub

    Private Sub btnCancellation_MouseUp(sender As Object, e As MouseEventArgs) Handles btnCancellation.MouseUp
        btnCancellation.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnCancellation_MouseLeave(sender As Object, e As EventArgs) Handles btnCancellation.MouseLeave
        Dim myfont As New Font(btnCancellation.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCancellation.BackgroundImage = My.Resources.button2normalDark
        btnCancellation.Font = myfont
    End Sub

    Private Sub btnGenerateReport_MouseDown(sender As Object, e As MouseEventArgs) Handles btnGenerateReport.MouseDown
        Dim myfont As New Font(btnGenerateReport.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnGenerateReport.BackgroundImage = My.Resources.button2Normal
        btnGenerateReport.Font = myfont
    End Sub

    Private Sub btnGenerateReport_MouseEnter(sender As Object, e As EventArgs) Handles btnGenerateReport.MouseEnter
        Dim myfont As New Font(btnGenerateReport.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnGenerateReport.BackgroundImage = My.Resources.button2normalBlackg
        btnGenerateReport.Font = myfont
    End Sub

    Private Sub btnGenerateReport_MouseUp(sender As Object, e As MouseEventArgs) Handles btnGenerateReport.MouseUp
        btnGenerateReport.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnGenerateReport_MouseLeave(sender As Object, e As EventArgs) Handles btnGenerateReport.MouseLeave
        Dim myfont As New Font(btnGenerateReport.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnGenerateReport.BackgroundImage = My.Resources.button2normalDark
        btnGenerateReport.Font = myfont
    End Sub

    Private Sub AboutUsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutUsToolStripMenuItem.Click
        AboutUs.Show()
    End Sub

    Private Sub HomeToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HomeToolStripMenuItem.Click
        Me.Refresh()
    End Sub
End Class