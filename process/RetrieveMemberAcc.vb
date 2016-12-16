Public Class RetrieveMemberAcc

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        prntdlgMemberDetails.ShowDialog()
    End Sub

    Private Sub btnPrint_MouseDown(sender As Object, e As MouseEventArgs) Handles btnPrint.MouseDown
        Dim myfont As New Font(btnPrint.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnPrint.BackgroundImage = My.Resources.button2Normal
        btnPrint.Font = myfont
    End Sub

    Private Sub btnPrint_MouseEnter(sender As Object, e As EventArgs) Handles btnPrint.MouseEnter
        Dim myfont As New Font(btnPrint.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnPrint.BackgroundImage = My.Resources.button2normalBlackg
        btnPrint.Font = myfont
    End Sub

    Private Sub btnPrint_MouseLeave(sender As Object, e As EventArgs) Handles btnPrint.MouseLeave
        Dim myfont As New Font(btnPrint.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnPrint.BackgroundImage = My.Resources.button2normalDark
        btnPrint.Font = myfont
    End Sub

    Private Sub btnPrint_MouseUp(sender As Object, e As MouseEventArgs) Handles btnPrint.MouseUp
        btnPrint.BackgroundImage = My.Resources.button2normalBlackg
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