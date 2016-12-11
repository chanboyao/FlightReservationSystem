Imports System.Data.OleDb
Public Class ForgetPassword

    Private Sub llblForgetIDorSAns_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblForgetIDorSAns.LinkClicked
        Dim intAns As Integer = MessageBox.Show("Please proceed to the counter" & vbCrLf &
                        "to retrieve your account's details.","Forget ID/Security Answer", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)


        If intAns = DialogResult.OK Then
            Me.Close()
        End If

    End Sub

    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim strAns As String
        strID = txtID.Text
        strAns = txtAns.Text

        If strID = "MXXXXXX" Or strID = "" Then
            MessageBox.Show("Please enter your ID to proceed.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtID.Focus()
            txtID.ForeColor = Color.Black
        ElseIf strAns = "" Then
            MessageBox.Show("Please enter your answer to retrieve your password.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtAns.Focus()
        Else
            cmd = New OleDbCommand("select * from Member where MemID='" & strID & "'", coon)
            dr = cmd.ExecuteReader
            If dr.Read Then

                If strAns.ToLower = dr.GetString(8).ToLower Then
                    strUsername = dr.GetString(1)
                    strPass = dr.GetString(2)
                    Dim intAns As Integer = MessageBox.Show("Hi, " & strUsername & "!" & vbCrLf &
                                                            "Here's your password." & vbCrLf & vbCrLf &
                                                            "Password : " & strPass, "Retrieve Password", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    If intAns = DialogResult.OK Then
                        Me.Close()
                    End If
                Else
                    MessageBox.Show("Wrong Answer. Please try again.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    txtAns.Clear()
                    txtAns.Focus()
                End If
            End If

        End If

    End Sub

    Private Sub ForgetPassword_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim constring As String
        constring = strConnection
        coon = New OleDbConnection(constring)
        coon.Open()
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        txtID.Clear()
        txtAns.Clear()
        cbQandA.SelectedIndex = -1
        txtID.Focus()
        txtID.ForeColor = Color.Black
    End Sub

    Private Sub txtAns_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAns.GotFocus
        Dim intIndex As Integer
        If txtID.Text = "MXXXXXX" Then
            txtID.Focus()
            txtID.ForeColor = Color.Black
            MessageBox.Show("Please enter your Member ID to view your security question.", "Blank Member ID", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Else
            strID = txtID.Text

            If strID.Length < 7 Then
                txtID.Clear()
                txtID.Focus()
                txtID.ForeColor = Color.Black
                MessageBox.Show("Member ID must at least have 7 character. Please re-enter.", "Invalid Member ID", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Else
                cmd = New OleDbCommand("select * from Member where MemID='" & strID & "'", coon)
                dr = cmd.ExecuteReader
                If dr.Read Then
                    intIndex = dr.GetInt32(7)
                    cbQandA.SelectedIndex = intIndex

                Else
                    txtID.Clear()
                    txtID.Focus()
                    txtID.ForeColor = Color.Black
                    MessageBox.Show("Invalid Member ID. Please re-enter." & vbCrLf &
                                    "*Tips : Member ID is case sensitive and format is important" & vbCrLf &
                                    "            OR " & vbCrLf &
                                    "            You simply just entered the wrong ID.", "Invalid Member ID", MessageBoxButtons.OK, MessageBoxIcon.Error)

                End If

            End If

        End If
    End Sub

    Private Sub txtID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtID.GotFocus
        If txtID.Text = "MXXXXXX" Then
            txtID.ForeColor = Color.Black
            txtID.Text = ""
        End If
    End Sub

    Private Sub txtID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtID.LostFocus
        If txtID.Text = "" Then
            txtID.ForeColor = Color.Gray
            txtID.Text = "MXXXXXX"
        End If
    End Sub

    Private Sub txtID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtID.TextChanged
        cbQandA.SelectedIndex = -1
        txtAns.Clear()
    End Sub

    Private Sub btnSubmit_MouseDown(sender As Object, e As MouseEventArgs) Handles btnSubmit.MouseDown
        Dim myfont As New Font(btnSubmit.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnSubmit.BackgroundImage = My.Resources.button2Normal
        btnSubmit.Font = myfont
    End Sub

    Private Sub btnSubmit_MouseEnter(sender As Object, e As EventArgs) Handles btnSubmit.MouseEnter
        Dim myfont As New Font(btnSubmit.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnSubmit.BackgroundImage = My.Resources.button2normalBlackg
        btnSubmit.Font = myfont
    End Sub

    Private Sub btnSubmit_MouseLeave(sender As Object, e As EventArgs) Handles btnSubmit.MouseLeave
        Dim myfont As New Font(btnSubmit.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnSubmit.BackgroundImage = My.Resources.button2normalDark
        btnSubmit.Font = myfont
    End Sub

    Private Sub btnSubmit_MouseUp(sender As Object, e As MouseEventArgs) Handles btnSubmit.MouseUp
        btnSubmit.BackgroundImage = My.Resources.button2normalBlackg
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

    Private Sub btnClear_MouseLeave(sender As Object, e As EventArgs) Handles btnClear.MouseLeave
        Dim myfont As New Font(btnClear.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnClear.BackgroundImage = My.Resources.button2normalDark
        btnClear.Font = myfont
    End Sub

    Private Sub btnClear_MouseUp(sender As Object, e As MouseEventArgs) Handles btnClear.MouseUp
        btnClear.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub lblSecurityQuestion_Click(sender As Object, e As EventArgs) Handles lblSecurityQuestion.Click

    End Sub
End Class