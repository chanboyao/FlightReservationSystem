Imports System.Data.OleDb
Public Class Register
    Dim conn As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Dim da As OleDbDataAdapter
    Dim ds As New DataSet
    Dim strUsername As String
    Dim strPassword As String
    Dim strRPassword As String
    Dim strBD As String
    Dim strDay As string
    Dim strMth As String
    Dim strYear As String
    Dim strEmail As String
    Dim strEmailType As String
    Dim strPassport As String
    Dim strCountry As String
    Dim intQstn As Integer
    Dim strAns As String

    Private Sub Register_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim constring As String
        constring = strConnection
        'constring = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\USER\Desktop\FlightReservationSystem\FlightReservation.accdb"
        conn = New OleDbConnection(constring)
        conn.Open()

    End Sub

    Private Sub Register_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        conn.Close()
    End Sub

    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim strSQLInsert As String

        txtPw.Focus()
        txtReenterPw.Focus()
        cmbDay.Focus()
        cmbMth.Focus()
        cmbYear.Focus()
        txtEmail.Focus()
        cmbEmail.Focus()
        txtPassportNo.Focus()
        cmbCountry.Focus()
        cmbQuestion.Focus()
        txtAns.Focus()

        If lblAUsername.Visible Or lblAPassword.Visible Or lblAReenterPw.Visible Or lblABirthDate.Visible Or lblAEmail.Visible Or lblAPassportNo.Visible Or lblACountry.Visible Or lblASecurityQuestion.Visible Or lblAAns.Visible Then
            Exit Sub
        Else
            strUsername = txtUsername.Text
            strPassword = txtPw.Text
            strBD = (cmbMth.SelectedIndex).ToString & "/" & cmbDay.SelectedItem.ToString & "/" & cmbYear.SelectedItem.ToString
            strEmail = txtEmail.Text & "@" & cmbEmail.SelectedItem.ToString
            strEmailType = cmbEmail.SelectedItem.ToString
            strPassport = txtPassportNo.Text
            strCountry = cmbCountry.SelectedItem.ToString
            intQstn = cmbQuestion.SelectedIndex
            strAns = txtAns.Text

            Dim lngAns As Integer = MessageBox.Show("Are you sure?", "Reconfirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

            If lngAns = DialogResult.Yes Then
                Dim strID As String = createID()

                strSQLInsert = "Insert into member" & _
                    "(MemID, MemUsername, MemPass, PassportNo, BirthDate, Email, Country, Question, Answer)values(" & _
                    "'" & strID & "'" & ", " & "'" & strUsername & "'" & ", " & _
                    "'" & strPassword & "'" & ", " & "'" & strPassport & "'" & ", " & "'" & strBD & "'" & ", " & "'" & strEmail & "'" & ", " & _
                    "'" & strCountry & "'" & ", " & "'" & intQstn & "'" & ", " & "'" & strAns & "'" & ");"

                cmd = New OleDb.OleDbCommand(strSQLInsert, conn)
                cmd.ExecuteNonQuery()

                lngAns = MessageBox.Show("Hi, " & strUsername & "!" & vbCrLf & "You have successfully register as our member!" & vbCrLf & vbCrLf &
                                "*********************************************************************" & vbCrLf &
                                "*Please remember the ID as it helps you to login as our member. *" & vbCrLf &
                                "*********************************************************************" & vbCrLf & vbCrLf &
                                "ID : " & strID, "Registration Successful!", MessageBoxButtons.OK, MessageBoxIcon.Information)

                If lngAns = DialogResult.OK Then
                    Me.Close()

                    Exit Sub
                End If
            End If
        End If

    End Sub

    Private Function createID() As String
        Dim strID As String = "M000000"
        Dim intTotal As Integer
        Dim cmdTotal As OleDbCommand
        Dim strTotal As String
        Dim intSubAt As Integer
        cmdTotal = New OleDbCommand _
            ("select count(*) from Member", conn)

        intTotal = cmdTotal.ExecuteScalar

        strTotal = CStr(intTotal + 1)

        intSubAt = 7 - strTotal.Length

        strID = (strID.Insert(intSubAt, strTotal)).Substring(0, 7)

        Return strID
    End Function

    Private Sub btnClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear.Click
        txtUsername.Clear()
        txtPw.Clear()
        txtReenterPw.Clear()
        cmbDay.SelectedIndex = -1
        cmbMth.SelectedIndex = -1
        cmbYear.SelectedIndex = -1
        txtEmail.Clear()
        cmbEmail.SelectedIndex = -1
        txtPassportNo.Clear()
        cmbCountry.SelectedIndex = -1
        cmbQuestion.SelectedIndex = -1
        txtAns.Clear()
        txtUsername.Focus()
        lblAUsername.Hide()
        lblAPassword.Hide()
        lblAReenterPw.Hide()
        lblABirthDate.Hide()
        lblAEmail.Hide()
        lblAPassportNo.Hide()
        lblACountry.Hide()
        lblASecurityQuestion.Hide()
        lblAAns.Hide()
    End Sub

    Private Sub txtUsername_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtUsername.LostFocus
        strUsername = txtUsername.Text

        If strUsername = "" Then
            lblAUsername.Show()
            lblAUsername.Text = "Your username is required."
        ElseIf strUsername.Length < 6 Then
            lblAUsername.Show()
            lblAUsername.Text = "Username must at least have 6 character."
        ElseIf strUsername.Length >= 6 Then
            Dim sqlChkUsername As String = "SELECT * FROM Member WHERE MemUsername = '" & txtUsername.Text & "'"
            cmd = New OleDbCommand(sqlChkUsername, conn)
            dr = cmd.ExecuteReader

            Dim blnGetID As Boolean = False

            If dr.Read Then
                blnGetID = True
            Else
                blnGetID = False
            End If

            If blnGetID Then
                lblAUsername.Show()
                lblAUsername.Text = "Oops. Someone else is using the same username, please enter a new one."
            ElseIf blnGetID = False Then
                lblAUsername.Hide()
            End If

        End If
    End Sub

    Private Sub txtPw_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPw.LostFocus
        strPassword = txtPw.Text

        If strPassword = "" Then
            lblAPassword.Show()
            lblAPassword.Text = "Please create a password."
        ElseIf strPassword.Length < 8 Then
            lblAPassword.Show()
            lblAPassword.Text = "Your password must at least have 8 character."
        ElseIf txtPw.TextLength >= 8 Then
            For Each c As Char In strPassword.ToLower
                If IsNumeric(c) Or Char.IsLetter(c) Then
                    lblAPassword.Hide()
                Else
                    lblAPassword.Show()
                    lblAPassword.Text = "Only alphebets and numbers allowed."
                    txtPw.Focus()
                    Exit Sub
                End If
            Next
        End If
    End Sub

    Private Sub txtReenterPw_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReenterPw.GotFocus
        If txtPw.Text = "" Then
            lblAPassword.Show()
            lblAPassword.Text = "Please create a password."
            txtPw.Focus()
        ElseIf strPassword.Length < 8 Then
            lblAPassword.Show()
            lblAPassword.Text = "Your password must at least have 8 character."
            txtPw.Focus()
        ElseIf txtPw.TextLength >= 8 Then
            For Each c As Char In strPassword.ToLower
                If IsNumeric(c) Or Char.IsLetter(c) Then
                    lblAPassword.Hide()
                    lblAReenterPw.Hide()
                Else
                    lblAPassword.Show()
                    lblAPassword.Text = "Only alphebets and numbers allowed."
                    txtPw.Focus()
                    Exit Sub
                End If
            Next
        End If

    End Sub

    Private Sub txtReenterPw_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtReenterPw.LostFocus
        strRPassword = txtReenterPw.Text

        If strRPassword = "" Then
            lblAReenterPw.Show()
            lblAReenterPw.Text = "Please re-enter password."
        ElseIf strPassword = strRPassword Then
            lblAReenterPw.Hide()
        Else
            lblAReenterPw.Show()
            lblAReenterPw.Text = "Please enter your password correctly."
        End If

    End Sub

    Private Sub cmbDay_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbDay.LostFocus
        If cmbDay.SelectedIndex = -1 Then
            'cmbDay.
            lblABirthDate.Show()
            lblABirthDate.Text = "The day of your birth is required."
        Else
            strDay = cmbDay.SelectedItem.ToString
            lblABirthDate.Hide()
        End If
    End Sub

    Private Sub cmbMth_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbMth.LostFocus
        Dim intDayLength As Integer = cmbDay.Items.Count

        If cmbMth.SelectedIndex = -1 Then
            lblABirthDate.Show()
            lblABirthDate.Text = "Your birth month is required."
            Exit Sub
        Else
            lblABirthDate.Hide()
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

        If cmbDay.SelectedIndex = -1 Then
            cmbDay.Focus()
        ElseIf cmbYear.SelectedIndex > -1 Then
            txtEmail.Focus()
        End If

    End Sub

    Private Sub cmbYear_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbYear.LostFocus
        If cmbYear.SelectedIndex = -1 Then
            lblABirthDate.Show()
            lblABirthDate.Text = "The year you were born is required."
            Exit Sub
        Else
            lblABirthDate.Hide()
            strYear = (cmbYear.SelectedIndex).ToString
        End If
    End Sub

    Private Sub txtEmail_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtEmail.LostFocus
        strEmail = txtEmail.Text

        If strEmail = "" Then
            lblAEmail.Show()
            lblAEmail.Text = "Please enter your e-mail."
        ElseIf strEmail.Length < 8 Then
            lblAEmail.Show()
            lblAEmail.Text = "Your e-mail must at least have 8 character."
        ElseIf txtEmail.TextLength >= 8 Then
            For Each c As Char In strEmail.ToLower
                If IsNumeric(c) Or Char.IsLetter(c) Then
                    If cmbEmail.SelectedIndex <> -1 Then
                        Dim strChkEmail = txtEmail.Text & "@" & cmbEmail.SelectedItem.ToString
                        Dim sqlChkEmail As String = "SELECT * FROM MEMBER WHERE Email = '" & strChkEmail & "'"
                        cmd = New OleDbCommand(sqlChkEmail, conn)
                        dr = cmd.ExecuteReader

                        While dr.Read
                            If strChkEmail = dr.GetString(5) Then
                                lblAEmail.Show()
                                lblAEmail.Text = "This e-mail has been used."
                                Exit Sub
                            End If
                        End While

                        lblAEmail.Hide()
                    End If
                Else
                    lblAEmail.Show()
                    lblAEmail.Text = "Only alphebets and numbers allowed."
                    txtEmail.Focus()
                    Exit Sub
                End If
            Next
        End If
    End Sub

    Private Sub cmbEmail_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbEmail.LostFocus
        If cmbEmail.SelectedIndex = -1 Then
            lblAEmail.Show()
            lblAEmail.Text = "The type of your e-mail is required."
        ElseIf txtEmail.Text <> "" Then
            Dim strChkEmail = txtEmail.Text & "@" & cmbEmail.SelectedItem.ToString
            Dim sqlChkEmail As String = "SELECT * FROM MEMBER WHERE Email = '" & strChkEmail & "'"
            cmd = New OleDbCommand(sqlChkEmail, conn)
            dr = cmd.ExecuteReader

            While dr.Read
                MessageBox.Show(dr.GetString(5))
                If strChkEmail = dr.GetString(5) Then
                    lblAEmail.Show()
                    lblAEmail.Text = "This e-mail has been used."
                    Exit Sub

                End If
            End While
            lblAEmail.Hide()
        End If
    End Sub

    Private Sub txtPassportNo_LostFocus1(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPassportNo.LostFocus
        strPassport = txtPassportNo.Text
        Dim strMsg As String = "Only "
        Dim blnChkLetter As Boolean = False
        Dim blnChkNum As Boolean = False
        Dim intAt As Integer = 0

        If strPassport = "" Then
            lblAPassportNo.Show()
            lblAPassportNo.Text = "Please enter your passport number."
        ElseIf strPassport.Length < 9 Then
            lblAPassportNo.Show()
            lblAPassportNo.Text = "Your passport number must at least have 9 character."
        ElseIf txtPassportNo.TextLength = 9 Then
            For Each c As Char In strPassport
                If intAt = 0 Then
                    If Char.IsLetter(c) And Char.IsUpper(c) Then
                        lblAPassportNo.Hide()
                        intAt = 1
                        blnChkLetter = True
                    Else
                        strMsg += "capital letters at the first character "
                        blnChkLetter = False
                        intAt = 1
                    End If
                Else
                    If IsNumeric(c) Then
                        lblAPassportNo.Hide()
                        blnChkNum = True
                    Else
                        If blnChkLetter = False Then
                            strMsg += "and numbers after the first letter "
                            blnChkNum = False
                            Exit For
                        Else
                            strMsg += "numbers after the first letter "
                            blnChkNum = False
                            Exit For
                        End If
                    End If
                End If
            Next

            If blnChkLetter = False Or blnChkNum = False Then
                lblAPassportNo.Show()
                lblAPassportNo.Text = strMsg & "allowed."
                Exit Sub
            Else
                Dim sqlChkPassport As String = "SELECT * FROM MEMBER WHERE PassportNo = '" & txtPassportNo.Text & "'"
                cmd = New OleDbCommand(sqlChkPassport, conn)
                dr = cmd.ExecuteReader

                Dim blnGetPassport As Boolean = False

                If dr.Read Then
                    blnGetPassport = True
                Else
                    blnGetPassport = False
                End If

                If blnGetPassport Then
                    lblAPassportNo.Show()
                    lblAPassportNo.Text = "This passport entered have been used. Please re-enter."
                ElseIf blnGetPassport = False Then
                    lblAPassportNo.Hide()
                End If

            End If
        End If

    End Sub

    Private Sub cmbCountry_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCountry.LostFocus
        If cmbCountry.SelectedIndex = -1 Then
            lblACountry.Show()
            lblACountry.Text = "Select the country you live in."
            Exit Sub
        ElseIf cmbCountry.SelectedItem.ToString = "*" Then
            lblACountry.Show()
            lblACountry.Text = "' * ' is just a divider. Please select the country you live in."
            Exit Sub
        Else
            lblACountry.Hide()
            strCountry = cmbCountry.SelectedItem.ToString
        End If
    End Sub

    Private Sub cmbQuestion_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbQuestion.LostFocus
        If cmbQuestion.SelectedIndex = -1 Then
            lblASecurityQuestion.Show()
            lblASecurityQuestion.Text = "Choose a question for future reference."
            Exit Sub
        Else
            lblASecurityQuestion.Hide()
            intQstn = cmbQuestion.SelectedIndex
        End If
    End Sub

    Private Sub txtAns_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAns.GotFocus
        If cmbQuestion.SelectedIndex = -1 Then
            lblASecurityQuestion.Show()
            lblASecurityQuestion.Text = "Choose a question for future reference."
            cmbQuestion.Focus()
            Exit Sub
        Else
            lblASecurityQuestion.Hide()
            intQstn = cmbQuestion.SelectedIndex
        End If
    End Sub

    Private Sub txtAns_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtAns.LostFocus
        strAns = txtAns.Text

        If cmbQuestion.SelectedIndex = -1 Then
            lblASecurityQuestion.Show()
            lblASecurityQuestion.Text = "Choose a question for future reference."
            cmbQuestion.Focus()
        Else
            If strAns = "" Then
                lblAAns.Show()
                lblAAns.Text = "Please create a password."
            ElseIf strAns.Length < 3 Then
                lblAAns.Show()
                lblAAns.Text = "Your answer must at least have 3 character."
            ElseIf txtAns.TextLength >= 3 Then
                For Each c As Char In strAns.ToLower
                    If IsNumeric(c) Or Char.IsLetter(c) Or c = " " Then
                        lblAAns.Hide()
                    Else
                        lblAAns.Show()
                        lblAAns.Text = "Only alphebets and numbers allowed."
                        txtAns.Focus()
                        Exit Sub
                    End If
                Next
            End If
        End If

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

End Class