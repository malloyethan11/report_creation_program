Public Class frmTaxReportInfo
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click

        ' reset control colors
        txtYear.BackColor = Color.White
        txtEmail.BackColor = Color.White

        ' declare variables
        Dim strToEmail As String
        Dim strYear As String
        Dim strMonth As String
        Dim strDay As String
        Dim strFile As String

        ' validate input
        If txtEmail.Text = "" Then
            txtEmail.BackColor = Color.Yellow
            MessageBox.Show("Please enter an email address.")

        Else
            strToEmail = txtEmail.Text

            If txtYear.Text = "" Then
                txtYear.BackColor = Color.Yellow
                MessageBox.Show("Please enter a year.")

            Else
                strYear = txtYear.Text

                If txtMonth.Text = "" Then
                    txtMonth.BackColor = Color.Yellow
                    MessageBox.Show("Please enter a month.")

                Else
                    strMonth = txtMonth.Text

                    If txtDay.Text = "" Then
                        txtDay.BackColor = Color.Yellow
                        MessageBox.Show("Please enter a day.")

                    Else
                        strDay = txtDay.Text
                        RunTaxReport(Me, False, strYear, strMonth, strDay)

                        GC.Collect()
                        GC.WaitForPendingFinalizers()

                        'strFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\TaxReport.xlsx"

                        'strFile = strFile.Remove(0, 6)

                        Threading.Thread.Sleep(3000)

                        SendMail(strToEmail, "TeamBeesCapstone@gmail.com", "Tax Report", "", "TeamBeesCapstone@gmail.com", "cincystate123", "TaxReport.xlsx", False)

                    End If
                End If
            End If
        End If

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        ' close the form
        Me.Close()

    End Sub

    Private Sub tmrUpdateButtonImage_Tick(sender As Object, e As EventArgs) Handles tmrUpdateButtonImage.Tick

        For Each Control In Controls
            If Control.GetType() = GetType(Button) Then
                ButtonColor(MousePosition, Control, Me, btmButtonShortGray, btmButtonShort)
            End If
        Next

    End Sub

    Private Sub frmTaxReportInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.CenterToScreen()

        For Each Control In Controls
            If Control.GetType() = GetType(Button) Then
                Control.FlatStyle = FlatStyle.Flat
                Control.ForeColor = BackColor
                Control.FlatAppearance.BorderColor = BackColor
                Control.FlatAppearance.MouseOverBackColor = BackColor
                Control.FlatAppearance.MouseDownBackColor = BackColor
            End If
        Next

    End Sub
End Class