

Public Class frmSalesReportInfo
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click

        ' reset control colors
        cboTimePeriod.BackColor = Color.White
        txtEmail.BackColor = Color.White

        ' declare variables
        Dim strTimePeriod As String

        ' get time period in which to run sales report
        If cboTimePeriod.SelectedItem Is Nothing Then
            cboTimePeriod.BackColor = Color.Yellow
            MessageBox.Show("Please select the time period for which you want to view sales data")

        Else

            strTimePeriod = cboTimePeriod.SelectedItem

            If txtEmail.Text = "" Then

                txtEmail.BackColor = Color.Yellow
                MessageBox.Show("Please enter the email you want to send the report to.")

            Else

                ' create sales report
                CreateSalesReport(Me, strTimePeriod, False)

            End If

        End If

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        ' close the form
        Me.Close()

    End Sub

End Class