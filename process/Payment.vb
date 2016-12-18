Imports System.Data.OleDb
Public Class Payment

    Dim coon As OleDbConnection
    Dim cmd As OleDbCommand
    Dim dr As OleDbDataReader
    Dim totalamount As Double = 0
    Dim dblpayment As Double = 0
    Dim rsvid As String = ""
    Dim payid As String = ""

    Private Function GetPaymentCode() As String
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

        cmd1 = New OleDbCommand("Select count(PayID) From Payment", coon)
        dr2 = cmd1.ExecuteReader
        If dr2.Read Then
            count = dr2.GetInt32(0)
            If count > 0 Then
                Dim strArray(count) As String
                cmd = New OleDbCommand("Select PayID From Payment", coon)
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
                        str1 = "PY" & charmax & "000000"
                    End If
                Else
                    intMax += 1
                    If CStr(intMax).Length < 2 Then
                        str1 = "PY" & charmax & "00000" & intMax
                    ElseIf CStr(intMax).Length < 3 Then
                        str1 = "PY" & charmax & "0000" & intMax
                    ElseIf CStr(intMax).Length < 4 Then
                        str1 = "PY" & charmax & "000" & intMax
                    ElseIf CStr(intMax).Length < 5 Then
                        str1 = "PY" & charmax & "00" & intMax
                    ElseIf CStr(intMax).Length < 6 Then
                        str1 = "PY" & charmax & "0" & intMax
                    Else
                        str1 = "PY" & charmax & intMax
                    End If
                End If

            Else
                str1 = "PYA000001"
            End If
        End If
        Return str1
    End Function

    Private Sub Payment_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim constring As String
        constring = strConnection
        coon = New OleDbConnection(constring)
        coon.Open()
        btnSubmit.Enabled = False
        cmbPaymentMethod.SelectedIndex = 0
        cmbPaymentMethod.SelectedIndex = -1
        If LoginID.Chars(0) = "M" Then
            radCash.Enabled = False
        End If

        If strReservationID = "" Then

        Else
            GPPay.BringToFront()

            txtRsvID.Text = strReservationID

            Dim checkexist As Integer
            rsvid = strReservationID

            cmd = New OleDbCommand("select count(ReservationID) from Reservation where ReservationID = '" & rsvid & "';", coon)
            dr = cmd.ExecuteReader
            dr.Read()
            checkexist = dr.GetInt32(0)
            If checkexist = 1 Then
                cmd = New OleDbCommand("select count(PayID) from Payment where ReservationID = '" & rsvid & "';", coon)
                dr = cmd.ExecuteReader
                dr.Read()
                checkexist = dr.GetInt32(0)
                If checkexist = 0 Then
                    cmd = New OleDbCommand("select SP.SeatPrice from FlightReservation FR , FlightSeatPrice SP where SP.SeatCode = FR.SeatCode and FR.ReservationID = '" & rsvid & "';", coon)
                    dr = cmd.ExecuteReader
                    While dr.Read
                        Dim price As Double
                        price = dr.GetDouble(0)
                        totalamount += price
                    End While
                    If totalamount > 0 Then
                        GPPay.BringToFront()
                        GPBlank.BringToFront()
                        txtRsvID.Text = rsvid
                        txtAmount.Text = "RM " & totalamount.ToString("F2")
                    Else
                        MessageBox.Show("Reservation ID error.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    End If
                Else
                    MessageBox.Show("Reservation has been paid.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    cmbRSX.SelectedIndex = -1
                    txtIdGain.Clear()
                End If
            Else
                MessageBox.Show("Reservation ID is not exist.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                cmbRSX.SelectedIndex = -1
                txtIdGain.Clear()
            End If
        End If
    End Sub

    Private Sub txtIdGain_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtIdGain.KeyPress
        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

            If Asc(e.KeyChar) <> 8 Then
                If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                    e.Handled = True
                End If
            End If
    End Sub

    Private Sub txtIdGain_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtIdGain.TextChanged


        If txtIdGain.Text.Count = 6 Then
            If cmbRSX.SelectedIndex <> -1 Then
                btnSubmit.Enabled = True
            Else
                btnSubmit.Enabled = False
            End If
        Else
            btnSubmit.Enabled = False
        End If

    End Sub

    Private Sub cmbRSX_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbRSX.SelectedIndexChanged


        If cmbRSX.SelectedIndex <> -1 Then
            If txtIdGain.Text.Count = 6 Then
                btnSubmit.Enabled = True
            Else
                btnSubmit.Enabled = False
            End If
        Else
            btnSubmit.Enabled = False
        End If

    End Sub

    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Dim checkexist As Integer
        rsvid = cmbRSX.SelectedItem.ToString & txtIdGain.Text

        cmd = New OleDbCommand("select count(ReservationID) from Reservation where ReservationID = '" & rsvid & "';", coon)
        dr = cmd.ExecuteReader
        dr.Read()
        checkexist = dr.GetInt32(0)
        If checkexist = 1 Then
            cmd = New OleDbCommand("select count(PayID) from Payment where ReservationID = '" & rsvid & "';", coon)
            dr = cmd.ExecuteReader
            dr.Read()
            checkexist = dr.GetInt32(0)
            If checkexist = 0 Then
                cmd = New OleDbCommand("select SP.SeatPrice from FlightReservation FR , FlightSeatPrice SP where SP.SeatCode = FR.SeatCode and FR.ReservationID = '" & rsvid & "';", coon)
                dr = cmd.ExecuteReader
                While dr.Read
                    Dim price As Double
                    price = dr.GetDouble(0)
                    totalamount += price
                End While
                If totalamount > 0 Then
                    GPPay.BringToFront()
                    GPBlank.BringToFront()
                    txtRsvID.Text = rsvid
                    txtAmount.Text = "RM " & totalamount.ToString("F2")
                Else
                    MessageBox.Show("Reservation ID error.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            Else
                MessageBox.Show("Reservation has been paid.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
                cmbRSX.SelectedIndex = -1
                txtIdGain.Clear()
            End If
        Else
            MessageBox.Show("Reservation ID is not exist.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmbRSX.SelectedIndex = -1
            txtIdGain.Clear()
        End If
    End Sub

    Private Sub radCash_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radCash.CheckedChanged
        If radCash.Checked Then
            GPCash.BringToFront()
            txtCashNeeded.Text = "RM " & totalamount.ToString("F2")
            txtCashPayment.Enabled = True
            txtCashPayment.Text = ""
            btnCalculate.Enabled = False
            txtCashBalance.Text = ""
            btnEnd.Enabled = False
        Else
            GPCash.SendToBack()
        End If
    End Sub

    Private Sub txtCashPayment_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCashPayment.Enter
        'btnSubmitPayment.PerformClick()
        'txtCashPayment.Text = "RM " & dblpayment.ToString("F2")
        'txtCashPayment.Enabled = False
    End Sub

    Private Sub txtCashPayment_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCashPayment.KeyPress
        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

        If Asc(e.KeyChar) = 13 Then
            dblpayment = CDbl(txtCashPayment.Text)
            txtCashPayment.Text = "RM " & dblpayment.ToString("F2")
            txtCashPayment.Enabled = False
            btnCalculate.PerformClick()
        End If

    End Sub

    Private Sub txtCashPayment_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCashPayment.TextChanged
        Dim dblpayment As Double
        If txtCashPayment.Text <> "" Then
            dblpayment = CDbl(txtCashPayment.Text)
        Else
            dblpayment = 0
        End If

        If dblpayment >= totalamount Then
            btnCalculate.Enabled = True
        Else
            btnCalculate.Enabled = False
        End If
    End Sub

    Private Sub btnCalculate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCalculate.Click
        If txtCashPayment.Text.Chars(0) = "R" Then
            txtCashBalance.Text = "RM " & (dblpayment - totalamount).ToString("F2")
        Else
            dblpayment = CDbl(txtCashPayment.Text)
            txtCashPayment.Text = "RM " & dblpayment.ToString("F2")
            txtCashPayment.Enabled = False
            txtCashBalance.Text = "RM " & (dblpayment - totalamount).ToString("F2")
        End If

        btnCalculate.Enabled = False
        btnEnd.Enabled = True
        btnEnd.Focus()
    End Sub

    Private Sub btnEnd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEnd.Click
        Dim ans As Long
        ans = MessageBox.Show("Are you sure want to end the procedure?", "Information", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If ans = DialogResult.OK Then
            CurrTime = TimeOfDay.ToString("hh:mm:ss tt")
            CurrDate = Date.Now().ToString("dd\/MM\/yyyy")
            payid = GetPaymentCode()
            cmd = New OleDbCommand("insert into Payment(PayID,Amount,PayDate,PayTime,ReservationID) values('" & payid & "','" & totalamount & "','" & CurrDate & "','" & CurrTime & "','" & rsvid & "');", coon)
            dr = cmd.ExecuteReader
            cmbRSX.SelectedIndex = -1
            txtIdGain.Text = ""
            If LoginID.Chars(0) = "M" Then
                Me.Hide()
                MemberHomePage.Show()
            Else
                Me.Hide()
                StaffHomePage.Show()
            End If
        End If
    End Sub

    Private Sub radCreditCard_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radCreditCard.CheckedChanged
        If radCreditCard.Checked Then
            GPCreditC.BringToFront()
            cmbPaymentMethod.SelectedIndex = -1
        Else
            GPCreditC.SendToBack()
        End If
    End Sub

    Private Sub txtCN1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCN1.KeyPress
        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtCN1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCN1.TextChanged
        If txtCN1.Text.Count = 4 Then
            txtCN2.Focus()
        End If

        If txtCN1.Text.Count = 4 And txtCN2.Text.Count = 4 And txtCN3.Text.Count = 4 And txtCN4.Text.Count = 4 And txtSecurity.Text.Count >= 3 And cmbMonth.SelectedIndex <> -1 And cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = True
            btnCCReset.Enabled = True
        ElseIf txtCN1.Text.Count <> 0 Or txtCN2.Text.Count <> 0 Or txtCN3.Text.Count <> 0 Or txtCN4.Text.Count <> 0 Or txtSecurity.Text.Count <> 0 Or cmbMonth.SelectedIndex <> -1 Or cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = True
        Else
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = False
        End If
    End Sub

    Private Sub txtCN2_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCN2.KeyPress
        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

        If txtCN2.Text.Count = 0 Then
            If Asc(e.KeyChar) = 8 Then
                txtCN1.Focus()
            End If
        End If

    End Sub

    Private Sub txtCN2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCN2.TextChanged
        If txtCN2.Text.Count = 4 Then
            txtCN3.Focus()
        End If

        If txtCN1.Text.Count = 4 And txtCN2.Text.Count = 4 And txtCN3.Text.Count = 4 And txtCN4.Text.Count = 4 And txtSecurity.Text.Count >= 3 And cmbMonth.SelectedIndex <> -1 And cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = True
            btnCCReset.Enabled = True
        ElseIf txtCN1.Text.Count <> 0 Or txtCN2.Text.Count <> 0 Or txtCN3.Text.Count <> 0 Or txtCN4.Text.Count <> 0 Or txtSecurity.Text.Count <> 0 Or cmbMonth.SelectedIndex <> -1 Or cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = True
        Else
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = False
        End If
    End Sub

    Private Sub txtCN3_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCN3.KeyPress
        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

        If txtCN3.Text.Count = 0 Then
            If Asc(e.KeyChar) = 8 Then
                txtCN2.Focus()
            End If
        End If

    End Sub

    Private Sub txtCN3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCN3.TextChanged
        If txtCN3.Text.Count = 4 Then
            txtCN4.Focus()
        End If

        If txtCN1.Text.Count = 4 And txtCN2.Text.Count = 4 And txtCN3.Text.Count = 4 And txtCN4.Text.Count = 4 And txtSecurity.Text.Count >= 3 And cmbMonth.SelectedIndex <> -1 And cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = True
            btnCCReset.Enabled = True
        ElseIf txtCN1.Text.Count <> 0 Or txtCN2.Text.Count <> 0 Or txtCN3.Text.Count <> 0 Or txtCN4.Text.Count <> 0 Or txtSecurity.Text.Count <> 0 Or cmbMonth.SelectedIndex <> -1 Or cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = True
        Else
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = False
        End If
    End Sub

    Private Sub txtCN4_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtCN4.KeyPress
        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If

        If txtCN4.Text.Count = 0 Then
            If Asc(e.KeyChar) = 8 Then
                txtCN3.Focus()
            End If
        End If

    End Sub

    Private Sub txtSecurity_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtSecurity.KeyPress
        '97 - 122 = Ascii codes for simple letters
        '65 - 90  = Ascii codes for capital letters
        '48 - 57  = Ascii codes for numbers

        If Asc(e.KeyChar) <> 8 Then
            If Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57 Then
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub txtSecurity_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSecurity.TextChanged
        If txtCN1.Text.Count = 4 And txtCN2.Text.Count = 4 And txtCN3.Text.Count = 4 And txtCN4.Text.Count = 4 And txtSecurity.Text.Count >= 3 And cmbMonth.SelectedIndex <> -1 And cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = True
            btnCCReset.Enabled = True
        ElseIf txtCN1.Text.Count <> 0 Or txtCN2.Text.Count <> 0 Or txtCN3.Text.Count <> 0 Or txtCN4.Text.Count <> 0 Or txtSecurity.Text.Count <> 0 Or cmbMonth.SelectedIndex <> -1 Or cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = True
        Else
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = False
        End If
    End Sub

    Private Sub txtCN4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtCN4.TextChanged
        If txtCN1.Text.Count = 4 And txtCN2.Text.Count = 4 And txtCN3.Text.Count = 4 And txtCN4.Text.Count = 4 And txtSecurity.Text.Count >= 3 And cmbMonth.SelectedIndex <> -1 And cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = True
            btnCCReset.Enabled = True
        ElseIf txtCN1.Text.Count <> 0 Or txtCN2.Text.Count <> 0 Or txtCN3.Text.Count <> 0 Or txtCN4.Text.Count <> 0 Or txtSecurity.Text.Count <> 0 Or cmbMonth.SelectedIndex <> -1 Or cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = True
        Else
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = False
        End If
    End Sub

    Private Sub cmbMonth_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbMonth.SelectedIndexChanged
        If txtCN1.Text.Count = 4 And txtCN2.Text.Count = 4 And txtCN3.Text.Count = 4 And txtCN4.Text.Count = 4 And txtSecurity.Text.Count >= 3 And cmbMonth.SelectedIndex <> -1 And cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = True
            btnCCReset.Enabled = True
        ElseIf txtCN1.Text.Count <> 0 Or txtCN2.Text.Count <> 0 Or txtCN3.Text.Count <> 0 Or txtCN4.Text.Count <> 0 Or txtSecurity.Text.Count <> 0 Or cmbMonth.SelectedIndex <> -1 Or cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = True
        Else
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = False
        End If
    End Sub

    Private Sub cmbYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbYear.SelectedIndexChanged
        If txtCN1.Text.Count = 4 And txtCN2.Text.Count = 4 And txtCN3.Text.Count = 4 And txtCN4.Text.Count = 4 And txtSecurity.Text.Count >= 3 And cmbMonth.SelectedIndex <> -1 And cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = True
            btnCCReset.Enabled = True
        ElseIf txtCN1.Text.Count <> 0 Or txtCN2.Text.Count <> 0 Or txtCN3.Text.Count <> 0 Or txtCN4.Text.Count <> 0 Or txtSecurity.Text.Count <> 0 Or cmbMonth.SelectedIndex <> -1 Or cmbYear.SelectedIndex <> -1 Then
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = True
        Else
            btnCCSubmit.Enabled = False
            btnCCReset.Enabled = False
        End If
    End Sub

    Private Sub btnCCReset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCCReset.Click
        cmbPaymentMethod.SelectedIndex = -1
    End Sub

    Private Sub cmbPaymentMethod_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbPaymentMethod.SelectedIndexChanged
        If cmbPaymentMethod.SelectedIndex <> -1 Then
            txtCN1.Enabled = True
            txtCN2.Enabled = True
            txtCN3.Enabled = True
            txtCN4.Enabled = True
            txtSecurity.Enabled = True
            cmbMonth.Enabled = True
            cmbYear.Enabled = True

        Else
            txtCN1.Enabled = False
            txtCN2.Enabled = False
            txtCN3.Enabled = False
            txtCN4.Enabled = False
            txtSecurity.Enabled = False
            cmbMonth.Enabled = False
            cmbYear.Enabled = False
            txtCN1.Clear()
            txtCN2.Clear()
            txtCN3.Clear()
            txtCN4.Clear()
            txtSecurity.Clear()
            cmbMonth.SelectedIndex = -1
            cmbYear.SelectedIndex = -1
        End If
    End Sub

    Private Sub btnCCSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCCSubmit.Click
        Dim ans As Long
        ans = MessageBox.Show("Are you sure want to submit payment?", "Information", MessageBoxButtons.OKCancel, MessageBoxIcon.Information)
        If ans = DialogResult.OK Then
            CurrTime = TimeOfDay.ToString("hh:mm:ss tt")
            CurrDate = Date.Now().ToString("dd\/MM\/yyyy")
            payid = GetPaymentCode()
            cmd = New OleDbCommand("insert into Payment(PayID,Amount,PayDate,PayTime,ReservationID) values('" & payid & "','" & totalamount & "','" & CurrDate & "','" & CurrTime & "','" & rsvid & "');", coon)
            dr = cmd.ExecuteReader
            MessageBox.Show("Payment Successful.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            cmbPaymentMethod.SelectedIndex = -1
            txtIdGain.Text = ""
            If LoginID.Chars(0) = "M" Then
                Me.Hide()
                MemberHomePage.Show()
            Else
                Me.Hide()
                StaffHomePage.Show()
            End If
        End If
    End Sub

    Private Sub LLSecurity_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LLSecurity.LinkClicked
        Security_Code.ShowDialog()
    End Sub

    Private Sub btnCCSubmit_MouseDown(sender As Object, e As MouseEventArgs) Handles btnCCSubmit.MouseDown
        Dim myfont As New Font(btnCCSubmit.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCCSubmit.BackgroundImage = My.Resources.button2Normal
        btnCCSubmit.Font = myfont
    End Sub

    Private Sub btnCCSubmit_MouseEnter(sender As Object, e As EventArgs) Handles btnCCSubmit.MouseEnter
        Dim myfont As New Font(btnCCSubmit.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnCCSubmit.BackgroundImage = My.Resources.button2normalBlackg
        btnCCSubmit.Font = myfont
    End Sub

    Private Sub btnCCSubmit_MouseLeave(sender As Object, e As EventArgs) Handles btnCCSubmit.MouseLeave
        Dim myfont As New Font(btnCCSubmit.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCCSubmit.BackgroundImage = My.Resources.button2normalDark
        btnCCSubmit.Font = myfont
    End Sub

    Private Sub btnCCSubmit_MouseUp(sender As Object, e As MouseEventArgs) Handles btnCCSubmit.MouseUp
        btnCCSubmit.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnCCReset_MouseDown(sender As Object, e As MouseEventArgs) Handles btnCCReset.MouseDown
        Dim myfont As New Font(btnCCReset.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCCReset.BackgroundImage = My.Resources.button2Normal
        btnCCReset.Font = myfont
    End Sub

    Private Sub btnCCReset_MouseEnter(sender As Object, e As EventArgs) Handles btnCCReset.MouseEnter
        Dim myfont As New Font(btnCCReset.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnCCReset.BackgroundImage = My.Resources.button2normalBlackg
        btnCCReset.Font = myfont
    End Sub

    Private Sub btnCCReset_MouseLeave(sender As Object, e As EventArgs) Handles btnCCReset.MouseLeave
        Dim myfont As New Font(btnCCReset.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCCReset.BackgroundImage = My.Resources.button2normalDark
        btnCCReset.Font = myfont
    End Sub

    Private Sub btnCCReset_MouseUp(sender As Object, e As MouseEventArgs) Handles btnCCReset.MouseUp
        btnCCReset.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnCalculate_MouseDown(sender As Object, e As MouseEventArgs) Handles btnCalculate.MouseDown
        Dim myfont As New Font(btnCalculate.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCalculate.BackgroundImage = My.Resources.button2Normal
        btnCalculate.Font = myfont
    End Sub

    Private Sub btnCalculate_MouseEnter(sender As Object, e As EventArgs) Handles btnCalculate.MouseEnter
        Dim myfont As New Font(btnCalculate.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnCalculate.BackgroundImage = My.Resources.button2normalBlackg
        btnCalculate.Font = myfont
    End Sub

    Private Sub btnCalculate_MouseLeave(sender As Object, e As EventArgs) Handles btnCalculate.MouseLeave
        Dim myfont As New Font(btnCalculate.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnCalculate.BackgroundImage = My.Resources.button2normalDark
        btnCalculate.Font = myfont
    End Sub

    Private Sub btnCalculate_MouseUp(sender As Object, e As MouseEventArgs) Handles btnCalculate.MouseUp
        btnCalculate.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnEnd_MouseDown(sender As Object, e As MouseEventArgs) Handles btnEnd.MouseDown
        Dim myfont As New Font(btnEnd.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnEnd.BackgroundImage = My.Resources.button2Normal
        btnEnd.Font = myfont
    End Sub

    Private Sub btnEnd_MouseEnter(sender As Object, e As EventArgs) Handles btnEnd.MouseEnter
        Dim myfont As New Font(btnEnd.Font.Name, 13, FontStyle.Bold Or FontStyle.Bold)
        btnEnd.BackgroundImage = My.Resources.button2normalBlackg
        btnEnd.Font = myfont
    End Sub

    Private Sub btnEnd_MouseLeave(sender As Object, e As EventArgs) Handles btnEnd.MouseLeave
        Dim myfont As New Font(btnEnd.Font.Name, 12, FontStyle.Bold Or FontStyle.Bold)
        btnEnd.BackgroundImage = My.Resources.button2normalDark
        btnEnd.Font = myfont
    End Sub

    Private Sub btnEnd_MouseUp(sender As Object, e As MouseEventArgs) Handles btnEnd.MouseUp
        btnEnd.BackgroundImage = My.Resources.button2normalBlackg
    End Sub

    Private Sub btnCancelRe_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancelRe.Click
        Dim result As DialogResult

        result = MessageBox.Show("Are you sure to cancel the reservation?", "Reconfirmation of Reservation Cancellation", MessageBoxButtons.YesNo, MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            Dim sqlDeleteReservation As String = "DELETE FROM Reservation WHERE ReservationID = '" & strReservationID & "'"
            Dim sqlChkMemOrStaff As String = "SELECT * FROM MemberReservation WHERE ReservationID = '" & strReservationID & "'"
            Dim sqlDeleteMOrSReservation As String = "DELETE FROM "
            Dim sqlDeleteFlightReservation As String = "DELETE * FROM FlightReservation WHERE ReservationID = '" & strReservationID & "'"

            cmd = New OleDbCommand(sqlChkMemOrStaff, coon)
            dr = cmd.ExecuteReader

            If dr.Read() Then
                sqlDeleteMOrSReservation += "MemberReservation WHERE ReservationID = '" & strReservationID & "'"
            Else
                sqlDeleteMOrSReservation += "StaffReservation WHERE ReservationID = '" & strReservationID & "'"
            End If

            Dim intChose As Integer = 0

            cmd = New OleDbCommand(sqlDeleteFlightReservation, coon)
            dr = cmd.ExecuteReader
            dr.Read()

            cmd = New OleDbCommand(sqlDeleteReservation, coon)
            dr = cmd.ExecuteReader
            dr.Read()

            cmd = New OleDbCommand(sqlDeleteMOrSReservation, coon)
            dr = cmd.ExecuteReader
            dr.Read()

            MessageBox.Show("Successfully cancelled the reservation.", "Deletion Completed", MessageBoxButtons.OK, MessageBoxIcon.Information)

            If LoginID.Chars(0) = "M" Then
                Me.Hide()
                MemberHomePage.Show()
            Else
                Me.Hide()
                StaffHomePage.Show()
            End If


        End If

    End Sub

End Class