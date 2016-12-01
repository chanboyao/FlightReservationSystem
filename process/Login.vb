Imports System.Data.OleDb
Public Class Login

    Dim coon As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader

    Private Function FStafflogin() As String
        Dim char1 As Char
        Dim charmax As Char = ""
        Dim int1 As Integer
        Dim intMax As Integer = -1
        Dim str1 As String = ""
        Dim str3 As String
        Dim i As Integer = 0
        Dim count As Integer
        Dim cmd1 As OleDbCommand
        Dim dr2 As OleDbDataReader

        cmd1 = New OleDbCommand("Select count(StfLoginID) From StaffLoginHistory", coon)
        dr2 = cmd1.ExecuteReader
        If dr2.Read Then
            count = dr2.GetInt32(0)
            If count > 0 Then
                Dim strArray(count) As String
                cmd = New OleDbCommand("Select StfLoginID From StaffLoginHistory", coon)
                dr = cmd.ExecuteReader

                i = 0
                While dr.Read
                    strArray(i) = dr.GetString(0)
                    i += 1
                End While

                For i = 0 To count - 1 Step 1
                    str3 = strArray(i)
                    char1 = str3.Chars(2)
                    If charmax = "" Then
                        charmax = char1
                    Else
                        If char1 > charmax Then
                            charmax = char1
                        Else
                            charmax = charmax
                        End If
                    End If
                Next

                For i = 0 To count - 1 Step 1
                    str3 = strArray(i)
                    char1 = str3.Chars(2)
                    If char1 = charmax Then
                        int1 = CInt(Microsoft.VisualBasic.Right(str3, 6))
                        If intMax = -1 Then
                            intMax = int1
                        Else
                            If int1 > intMax Then
                                intMax = int1
                            Else
                                intMax = intMax
                            End If
                        End If
                    End If
                Next

                If intMax = 999999 Then
                    If charmax = "Z" Then
                        MessageBox.Show("Insufficent place.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Me.Close()
                    Else
                        charmax = Chr(Asc(charmax) + 1)
                        str1 = "SL" & charmax & "000000"
                    End If
                Else
                    intMax += 1
                    If CStr(intMax).Length < 2 Then
                        str1 = "SL" & charmax & "00000" & intMax
                    ElseIf CStr(intMax).Length < 3 Then
                        str1 = "SL" & charmax & "0000" & intMax
                    ElseIf CStr(intMax).Length < 4 Then
                        str1 = "SL" & charmax & "000" & intMax
                    ElseIf CStr(intMax).Length < 5 Then
                        str1 = "SL" & charmax & "00" & intMax
                    ElseIf CStr(intMax).Length < 6 Then
                        str1 = "SL" & charmax & "0" & intMax
                    Else
                        str1 = "SL" & charmax & intMax
                    End If
                End If

            Else
                str1 = "SLA000001"
            End If

        End If
        Return str1
    End Function

    Private Function FMemberlogin() As String
        Dim char1 As Char
        Dim charmax As Char = ""
        Dim int1 As Integer
        Dim intMax As Integer = -1
        Dim str1 As String = ""
        Dim str3 As String
        Dim i As Integer = 0
        Dim count As Integer
        Dim cmd1 As OleDbCommand
        Dim dr2 As OleDbDataReader

        cmd1 = New OleDbCommand("Select count(MemLoginID) From MemberLoginHistory", coon)
        dr2 = cmd1.ExecuteReader
        If dr2.Read Then
            count = dr2.GetInt32(0)
            If count > 0 Then
                Dim strArray(count) As String
                cmd = New OleDbCommand("Select MemLoginID From MemberLoginHistory", coon)
                dr = cmd.ExecuteReader

                i = 0
                While dr.Read
                    strArray(i) = dr.GetString(0)
                    i += 1
                End While

                For i = 0 To count - 1 Step 1
                    str3 = strArray(i)
                    char1 = str3.Chars(2)
                    If charmax = "" Then
                        charmax = char1
                    Else
                        If char1 > charmax Then
                            charmax = char1
                        Else
                            charmax = charmax
                        End If
                    End If
                Next

                For i = 0 To count - 1 Step 1
                    str3 = strArray(i)
                    char1 = str3.Chars(2)
                    If char1 = charmax Then
                        int1 = CInt(Microsoft.VisualBasic.Right(str3, 6))
                        If intMax = -1 Then
                            intMax = int1
                        Else
                            If int1 > intMax Then
                                intMax = int1
                            Else
                                intMax = intMax
                            End If
                        End If
                    End If
                Next

                If intMax = 999999 Then
                    If charmax = "Z" Then
                        MessageBox.Show("Insufficent place.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Me.Close()
                    Else
                        charmax = Chr(Asc(charmax) + 1)
                        str1 = "ML" & charmax & "000000"
                    End If
                Else
                    intMax += 1
                    If CStr(intMax).Length < 2 Then
                        str1 = "ML" & charmax & "00000" & intMax
                    ElseIf CStr(intMax).Length < 3 Then
                        str1 = "ML" & charmax & "0000" & intMax
                    ElseIf CStr(intMax).Length < 4 Then
                        str1 = "ML" & charmax & "000" & intMax
                    ElseIf CStr(intMax).Length < 5 Then
                        str1 = "ML" & charmax & "00" & intMax
                    ElseIf CStr(intMax).Length < 6 Then
                        str1 = "ML" & charmax & "0" & intMax
                    Else
                        str1 = "ML" & charmax & intMax
                    End If
                End If

            Else
                str1 = "MLA000001"
            End If
        End If
        Return str1
    End Function

    Private Sub llblRegister_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblRegister.LinkClicked
        Register.ShowDialog()
    End Sub

    Private Sub llblForgetPassword_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles llblForgetPassword.LinkClicked
        ForgetPassword.ShowDialog()
    End Sub

    Private Sub btnLogin_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLogin.Click

        strID = txtID.Text
        strPass = txtPassword.Text
        If strID = "" Then
            MessageBox.Show("Please enter your ID to continue process.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtID.Focus()
            txtID.ForeColor = Color.Black
        ElseIf strPass = "" Then
            MessageBox.Show("Please enter your Password to continue process.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtPassword.Focus()
        ElseIf strID.Length < 7 Then
            MessageBox.Show("Please enter correct ID format to continue process.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            txtID.Clear()
            txtPassword.Clear()
            txtID.Focus()
            txtID.ForeColor = Color.Black
        Else
            cmd = New OleDbCommand("select * from Member where MemID='" & strID & "'", coon)
            dr = cmd.ExecuteReader
            If dr.Read Then
                If strPass = dr.GetString(2) Then
                    MessageBox.Show("Login Successful!!!", "Login", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    CurrTime = CDate(TimeOfDay.ToString("h:mm:ss tt"))
                    CurrDate = Date.Now().ToString("dd\/MM\/yyyy")
                    cmd = New OleDbCommand("insert into MemberLoginHistory(MemLoginID,MemID,MLTime,MLDate) values('" & FMemberlogin() & "','" & strID & "','" & CurrTime & "','" & CurrDate & "');", coon)
                    dr = cmd.ExecuteReader
                    LoginID = strID

                    MemberHomePage.Show()
                    Me.Close()

                Else
                    MessageBox.Show("Password and ID that you enter was not match." & vbCrLf & "Try again later.", "Login", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Else
                cmd = New OleDbCommand("select * from Staff where StfID='" & strID & "'", coon)
                dr = cmd.ExecuteReader
                If dr.Read Then
                    If strPass = dr.GetString(2) Then
                        MessageBox.Show("Login Successful!!!", "Login", MessageBoxButtons.OK, MessageBoxIcon.Information)
                        CurrTime = TimeOfDay.ToString("h:mm:ss tt")
                        CurrDate = Date.Now().ToString("dd\/MM\/yyyy")
                        cmd = New OleDbCommand("insert into StaffLoginHistory(StfLoginID,StfID,SLTime,SLDate) values('" & FStafflogin() & "','" & strID & "','" & CurrTime & "','" & CurrDate & "');", coon)
                        dr = cmd.ExecuteReader
                        LoginID = strID

                        StaffHomePage.Show()
                        Me.Close()
                    Else
                        MessageBox.Show("Password and ID that you enter was not match." & vbCrLf & "Try again later.", "Login", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                Else
                    MessageBox.Show("ID was not found." & vbCrLf & _
                                        "*Tips : Member ID is case sensitive and format is important" & vbCrLf & _
                                        "            OR " & vbCrLf & _
                                        "            You simply just entered the wrong ID.", "Login", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    txtID.Clear()
                    txtPassword.Clear()
                    txtID.Focus()
                    txtID.ForeColor = Color.Black
                End If
            End If
        End If
    End Sub

    Private Sub Login_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim constring As String
        constring = strConnection
        coon = New OleDbConnection(constring)
        coon.Open()
        FlightSplashScreen.Close()
    End Sub

    Private Sub Login_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        coon.Close()
    End Sub

    Private Sub txtID_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtID.GotFocus
        If txtID.Text = "MXXXXXX" Then
            txtID.Text = ""
            txtID.ForeColor = Color.Black
        End If
    End Sub

    Private Sub txtID_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtID.KeyPress
        If Asc(e.KeyChar) = 13 Then
            txtPassword.Focus()
        End If
    End Sub

    Private Sub txtID_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtID.LostFocus
        If txtID.Text = "" Or txtID.Text.ToLower = "mxxxxxx" Then
            txtID.Text = "MXXXXXX"
            txtID.ForeColor = Color.LightGray
            btnLogin.Enabled = False
        End If
    End Sub

    Private Sub txtID_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtID.TextChanged
        If txtID.Text = "" And txtPassword.Text = "" Then
            btnClear.Enabled = False
            btnLogin.Enabled = False
        ElseIf txtID.Text = "" Or txtPassword.Text = "" Then
            btnClear.Enabled = True
            btnLogin.Enabled = False
        Else
            btnClear.Enabled = True
            btnLogin.Enabled = True
        End If
    End Sub

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        txtID.Clear()
        txtPassword.Clear()
        btnLogin.Enabled = False
        btnClear.Enabled = False
        txtID.Focus()
        txtID.ForeColor = Color.Black
    End Sub

    Private Sub txtPassword_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPassword.GotFocus
        If txtID.Text = "MXXXXXX" Then
            btnClear.Enabled = False
            btnLogin.Enabled = False
        End If
    End Sub

    Private Sub txtPassword_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtPassword.KeyPress
        If Asc(e.KeyChar) = 13 Then
            btnLogin.PerformClick()
        End If
    End Sub

    Private Sub txtPassword_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtPassword.TextChanged
        If txtID.Text = "" And txtPassword.Text = "" Then
            btnClear.Enabled = False
            btnLogin.Enabled = False
        ElseIf txtID.Text = "" Or txtPassword.Text = "" Then
            btnClear.Enabled = True
            btnLogin.Enabled = False
        Else
            If txtID.Text = "MXXXXXX" Then
                btnClear.Enabled = True
                btnLogin.Enabled = False
            Else
                btnClear.Enabled = True
                btnLogin.Enabled = True
            End If

        End If
    End Sub

    Private Sub btnLogin_MouseLeave(sender As Object, e As EventArgs) Handles btnLogin.MouseLeave
        Dim myfont As New Font(btnLogin.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnLogin.BackgroundImage = My.Resources.button2lock
        btnLogin.Font = myfont

    End Sub

    Private Sub btnLogin_MouseEnter(sender As Object, e As EventArgs) Handles btnLogin.MouseEnter
        Dim myfont As New Font(btnLogin.Font.Name, 14, FontStyle.Bold Or FontStyle.Bold)
        btnLogin.BackgroundImage = My.Resources.button2lockDark
        btnLogin.Font = myfont
    End Sub

    Private Sub btnLogin_MouseDown(sender As Object, e As MouseEventArgs) Handles btnLogin.MouseDown
        Dim myfont As New Font(btnLogin.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnLogin.BackgroundImage = My.Resources.button2Black
        btnLogin.Font = myfont
    End Sub

    Private Sub btnLogin_MouseUp(sender As Object, e As MouseEventArgs) Handles btnLogin.MouseUp
        btnLogin.BackgroundImage = My.Resources.button2lockDark
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub

    Private Sub btnClear_MouseDown(sender As Object, e As MouseEventArgs) Handles btnClear.MouseDown
        Dim myfont As New Font(btnLogin.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnClear.BackgroundImage = My.Resources.buttonclearnormal
    End Sub

    Private Sub btnClear_MouseEnter(sender As Object, e As EventArgs) Handles btnClear.MouseEnter
        btnClear.BackgroundImage = My.Resources.buttonclearreddark
    End Sub

    Private Sub btnClear_MouseLeave(sender As Object, e As EventArgs) Handles btnClear.MouseLeave
        btnClear.BackgroundImage = My.Resources.button2
    End Sub

    Private Sub btnClear_MouseUp(sender As Object, e As MouseEventArgs) Handles btnClear.MouseUp
        btnClear.BackgroundImage = My.Resources.buttonclearreddark
    End Sub

End Class
