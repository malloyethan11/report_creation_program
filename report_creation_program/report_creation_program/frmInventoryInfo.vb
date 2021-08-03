Public Class frmInventoryInfo
    Private Sub frmInventoryInfo_Load(sender As Object, e As EventArgs) Handles MyBase.Load

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

    Private Sub tmrUpdateButtonImage_Tick(sender As Object, e As EventArgs) Handles tmrUpdateButtonImage.Tick

        For Each Control In Controls
            If Control.GetType() = GetType(Button) Then
                ButtonColor(MousePosition, Control, Me, btmButtonShortGray, btmButtonShort)
            End If
        Next

    End Sub

    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click

        Dim strToEmail As String

        ' validate input
        If txtEmail.Text = "" Or System.Text.RegularExpressions.Regex.IsMatch(txtEmail.Text, "^([0-9a-zA-Z]([-\.\w]*[0-9a-zA-Z])*@([0-9a-zA-Z][-\w]*[0-9a-zA-Z]\.)+[a-zA-Z]{2,9})$") = False Then
            txtEmail.BackColor = Color.Yellow
            MessageBox.Show("Please enter a valid email address.")

        Else

            strToEmail = txtEmail.Text

            txtEmail.BackColor = Color.White

            ' run inventory report
            Dim blnResult As Boolean = RunInventoryReport(Me, False)

            GC.Collect()
            GC.WaitForPendingFinalizers()

            '' email inventory report
            'strFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\InventoryReport.xlsx"

            'strFile = strFile.Remove(0, 6)

            Threading.Thread.Sleep(3000)

            If blnResult = True Then
                SendMail(txtEmail.Text, "TeamBeesCapstone@gmail.com", "Inventory Report", "", "TeamBeesCapstone@gmail.com", "cincystate123", "InventoryReport.xlsx", False)
            End If
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub
End Class