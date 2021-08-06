
Public Class frmMain

    Private Sub btnSalesReport_Click(sender As Object, e As EventArgs) Handles btnSalesReport.Click

        ' declare variables
        Dim frmSalesReportInfo = New frmSalesReportInfo()

        ' Opens form to get user input and run sales report
        frmSalesReportInfo.ShowDialog()

    End Sub

    Private Sub btnTaxReport_Click(sender As Object, e As EventArgs) Handles btnTaxReport.Click

        ' declare variables
        Dim frmTaxReportInfo = New frmTaxReportInfo()

        ' open child form to get user input and run tax report
        frmTaxReportInfo.ShowDialog()

    End Sub

    Private Sub btnInventoryReport_Click(sender As Object, e As EventArgs) Handles btnInventoryReport.Click

        '' declare variables
        'Dim strToEmail As String
        'Dim strFile As String

        '' show messagebox asking user where they want to send the report
        'strToEmail = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

        '' don't progress if user enters blank input/presses cancel
        'If strToEmail = "" Then
        '    MessageBox.Show("You failed to enter an email address or clicked cancel. The operation will terminate, and no report will be generated.")
        '    Exit Sub
        'End If

        '' run inventory report
        'RunInventoryReport(Me, False)

        'GC.Collect()
        'GC.WaitForPendingFinalizers()

        ''' email inventory report
        ''strFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\InventoryReport.xlsx"

        ''strFile = strFile.Remove(0, 6)

        'Threading.Thread.Sleep(3000)

        'SendMail(strToEmail, "TeamBeesCapstone@gmail.com", "Inventory Report", "", "TeamBeesCapstone@gmail.com", "cincystate123", "InventoryReport.xlsx", False)

        ' declare variables
        Dim frmInventoryInfo = New frmInventoryInfo()

        ' Opens form to get user input and run deposits report
        frmInventoryInfo.ShowDialog()

    End Sub

    Private Sub btnCashCreditDepositReport_Click(sender As Object, e As EventArgs) Handles btnCashCreditDepositReport.Click

        ' declare variables
        Dim frmCashCreditDepositReportInfo = New frmCashCreditDepositReportInfo()

        ' Opens form to get user input and run deposits report
        frmCashCreditDepositReportInfo.ShowDialog()

    End Sub


    Friend WithEvents btnExit As Button

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        ' close the form
        Me.Close()

    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Me.CenterToScreen()

        For Each Control In Controls
            If Control.GetType() = GetType(System.Windows.Forms.Button) Then
                Control.FlatStyle = FlatStyle.Flat
                Control.ForeColor = BackColor
                Control.FlatAppearance.BorderColor = BackColor
                Control.FlatAppearance.MouseOverBackColor = BackColor
                Control.FlatAppearance.MouseDownBackColor = BackColor
            End If
        Next

    End Sub

    Private Sub tmrUpdateButtonImage_Tick(sender As Object, e As EventArgs) Handles tmrUpdateButtonImage.Tick

        ButtonColor(MousePosition, btnExit, Me, btmButtonShortGray, btmButtonShort)
        ButtonColor(MousePosition, btnInventoryReport, Me, btmButtonShortGray, btmButtonShort)
        ButtonColor(MousePosition, btnSalesReport, Me, btmButtonShortGray, btmButtonShort)
        ButtonColor(MousePosition, btnTaxReport, Me, btmButtonShortGray, btmButtonShort)
        ButtonColor(MousePosition, btnCashCreditDepositReport, Me, btmButtonShortGray, btmButtonShort)

    End Sub
End Class