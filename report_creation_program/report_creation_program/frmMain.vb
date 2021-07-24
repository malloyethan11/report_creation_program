
Public Class frmMain

    Dim intElapsedTimeUpdateConfig As Integer = 1000000
    Dim intElapsedTimeRunReport As Integer = 0

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

        ' declare variables
        Dim strToEmail As String
        Dim strFile As String

        ' show messagebox asking user where they want to send the report
        strToEmail = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

        ' don't progress if user enters blank input/presses cancel
        If strToEmail = "" Then
            MessageBox.Show("You failed to enter an email address or clicked cancel. The operation will terminate, and no report will be generated.")
            Exit Sub
        End If

        ' run inventory report
        RunInventoryReport(Me, False)

        ' email inventory report
        strFile = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().CodeBase) + "\InventoryReport.xlsx"

        strFile = strFile.Remove(0, 6)

        Threading.Thread.Sleep(3000)

        SendMail(strToEmail, "TeamBeesCapstone@gmail.com", "Inventory Report", "", "TeamBeesCapstone@gmail.com", "cincystate123", strFile)

    End Sub

    Private Sub btnCashCreditDepositReport_Click(sender As Object, e As EventArgs) Handles btnCashCreditDepositReport.Click

        ' declare variables
        Dim strUserInput As String

        ' show messagebox asking user where they want to send the report
        strUserInput = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

        ' don't progress if user enters blank input/presses cancel
        If strUserInput = "" Then
            MessageBox.Show("You failed to enter an email address or clicked cancel. The operation will terminate, and no report will be generated.")
            Exit Sub
        End If

        RunCashCreditReport()

    End Sub


    Friend WithEvents btnExit As Button

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        ' close the form
        Me.Close()

    End Sub

    ' ┍━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┑
    ' │ From here down is the functionality to automate the reports │
    ' ┕━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━┙

    Private Sub tmrUpdateLocalConfig_Tick(sender As Object, e As EventArgs) Handles tmrUpdateLocalConfig.Tick

        ' Elapse time
        intElapsedTimeUpdateConfig += 1

        ' Has 5 minutes elapsed?
        If (intElapsedTimeUpdateConfig >= 60 * 5 * 2) Then
            ' Get db update
            If OpenDatabaseConnectionSQLServer() = False Then

                ' The database is not open
                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                            "The form will now close.",
                            Me.Text + " Error",
                            MessageBoxButtons.OK, MessageBoxIcon.Error)

                ' Close the form/application
                Me.Close()

            End If

            ' Init select statement string
            Dim strSelect As String = ""
            ' Init select statement Db command
            Dim cmdSelect As OleDb.OleDbCommand
            ' Init data reader
            Dim drSourceTable As OleDb.OleDbDataReader
            ' Init data table
            Dim dt As DataTable = New DataTable

            ' Build the select statement
            strSelect = "SELECT * FROM TReports"

            ' Retrieve all the records 
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader
            ' Get data
            dt.Load(drSourceTable)

            ' Load from data table
            strEmailSalesReport = dt(0)("strTargetEmail").ToString
            strEmailSalesTaxReport = dt(2)("strTargetEmail").ToString
            strEmailInventoryReport = dt(1)("strTargetEmail").ToString
            strEmailDespositReport = dt(3)("strTargetEmail").ToString

            blnSalesReportDaily = dt(0)("blnDaily").ToString
            blnSalesReportMonthly = dt(0)("blnMonthly").ToString
            blnSalesReportWeekly = dt(0)("blnWeekly").ToString
            blnSalesReportYearly = dt(0)("blnYearly").ToString

            blnInventoryReportDaily = dt(1)("blnDaily").ToString
            blnInventoryReportMonthly = dt(1)("blnMonthly").ToString
            blnInventoryReportWeekly = dt(1)("blnWeekly").ToString
            blnInventoryReportYearly = dt(1)("blnYearly").ToString

            blnSalesTaxReportDaily = dt(2)("blnDaily").ToString
            blnSalesTaxReportMonthly = dt(2)("blnMonthly").ToString
            blnSalesTaxReportWeekly = dt(2)("blnWeekly").ToString
            blnSalesTaxReportYearly = dt(2)("blnYearly").ToString

            blnDepositReportDaily = dt(3)("blnDaily").ToString
            blnDepositReportMonthly = dt(3)("blnMonthly").ToString
            blnDepositReportWeekly = dt(3)("blnWeekly").ToString
            blnDepositReportYearly = dt(3)("blnYearly").ToString

            dtmDailySalesReport = Convert.ToDateTime(dt(0)("dtDailyReportDate").ToString())
            dtmDailySalesTaxReport = Convert.ToDateTime(dt(2)("dtDailyReportDate").ToString())
            dtmDailyInventoryReport = Convert.ToDateTime(dt(1)("dtDailyReportDate").ToString())
            dtmDailyDepositReport = Convert.ToDateTime(dt(3)("dtDailyReportDate").ToString())

            dtmWeeklySalesReport = Convert.ToDateTime(dt(0)("dtWeeklyReportDate").ToString())
            dtmWeeklySalesTaxReport = Convert.ToDateTime(dt(2)("dtWeeklyReportDate").ToString())
            dtmWeeklyInventoryReport = Convert.ToDateTime(dt(1)("dtWeeklyReportDate").ToString())
            dtmWeeklyDepositReport = Convert.ToDateTime(dt(3)("dtWeeklyReportDate").ToString())

            dtmMonthlySalesReport = Convert.ToDateTime(dt(0)("dtMonthlyReportDate").ToString())
            dtmMonthlySalesTaxReport = Convert.ToDateTime(dt(2)("dtMonthlyReportDate").ToString())
            dtmMonthlyInventoryReport = Convert.ToDateTime(dt(1)("dtMonthlyReportDate").ToString())
            dtmMonthlyDepositReport = Convert.ToDateTime(dt(3)("dtMonthlyReportDate").ToString())

            dtmYearlySalesReport = Convert.ToDateTime(dt(0)("dtYearlyReportDate").ToString())
            dtmYearlySalesTaxReport = Convert.ToDateTime(dt(2)("dtYearlyReportDate").ToString())
            dtmYearlyInventoryReport = Convert.ToDateTime(dt(1)("dtYearlyReportDate").ToString())
            dtmYearlyDepositReport = Convert.ToDateTime(dt(3)("dtYearlyReportDate").ToString())

            ' Reset
            intElapsedTimeUpdateConfig = 0
        End If

    End Sub

    Private Sub tmrCheckIfReportRun_Tick(sender As Object, e As EventArgs) Handles tmrCheckIfReportRun.Tick

        ' Elapse time
        intElapsedTimeRunReport += 1

        ' Has 5 minutes elapsed?
        If (intElapsedTimeRunReport >= 1) Then

            ' Run sales report
            AutomateSalesReport()

            ' Reset
            intElapsedTimeRunReport = 0

        End If

    End Sub

    Private Sub AutomateSalesReport()

        ' Sales report
        SalesDaily()
        SalesWeekly()
        SalesMonthly()
        SalesYearly()

        ' Inventory report
        InventoryDaily()
        InventoryWeekly()
        InventoryMonthly()
        InventoryYearly()

    End Sub

    Private Sub SalesDaily()

        ' Run daily?
        If (blnSalesReportDaily = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = dtmDailySalesReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(0, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Okay, we're in. Run the report in quiet mode
                    CreateSalesReport(Me, "last day", True)

                    ' email sales report
                    SendMail(strEmailSalesReport, "TeamBeesCapstone@gmail.com", "Daily Sales Report", "Automated message: See attached sales report.", "TeamBeesCapstone@gmail.com", "cincystate123", "SalesReport.xlsx", True)

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(0, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(0, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub SalesWeekly()

        ' Run weekly?
        If (blnSalesReportWeekly = True) Then

            ' Time period
            Dim strNow As String = WeekdayName(Weekday(DateTime.Now)) & " " & DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = WeekdayName(Weekday(dtmWeeklySalesReport)) & " " & dtmWeeklySalesReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(1, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Okay, we're in. Run the report in quiet mode
                    CreateSalesReport(Me, "last week", True)

                    ' email sales report
                    SendMail(strEmailSalesReport, "TeamBeesCapstone@gmail.com", "Weekly Sales Report", "Automated message: See attached sales report.", "TeamBeesCapstone@gmail.com", "cincystate123", "SalesReport.xlsx", True)

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(1, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(1, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub SalesMonthly()

        ' Run weekly?
        If (blnSalesReportMonthly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmMonthlySalesReport.Day Then
                strTarget = dtmMonthlySalesReport.ToString(intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmMonthlySalesReport.ToString("dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(2, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Okay, we're in. Run the report in quiet mode
                    CreateSalesReport(Me, "last month (30 days)", True)

                    ' email sales report
                    SendMail(strEmailSalesReport, "TeamBeesCapstone@gmail.com", "Monthly Sales Report", "Automated message: See attached sales report.", "TeamBeesCapstone@gmail.com", "cincystate123", "SalesReport.xlsx", True)

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(2, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(2, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub SalesYearly()

        ' Run weekly?
        If (blnSalesReportYearly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("MM/dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, dtmYearlySalesReport.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmYearlySalesReport.Day Then
                strTarget = dtmYearlySalesReport.ToString("MM/" & intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmYearlySalesReport.ToString("MM/dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(3, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Okay, we're in. Run the report in quiet mode
                    CreateSalesReport(Me, "last year (365 days)", True)

                    ' email sales report
                    SendMail(strEmailSalesReport, "TeamBeesCapstone@gmail.com", "Yearly Sales Report", "Automated message: See attached sales report.", "TeamBeesCapstone@gmail.com", "cincystate123", "SalesReport.xlsx", True)

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(3, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(3, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub


    Private Sub InventoryDaily()

        ' Run daily?
        If (blnInventoryReportDaily = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = dtmDailyInventoryReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(0, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Okay, we're in. Run the report in quiet mode
                    RunInventoryReport(Me, True)

                    ' email Inventory report
                    SendMail(strEmailInventoryReport, "TeamBeesCapstone@gmail.com", "Daily Inventory Report", "Automated message: See attached Inventory report.", "TeamBeesCapstone@gmail.com", "cincystate123", "InventoryReport.xlsx", True)

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(0, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(0, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub InventoryWeekly()

        ' Run weekly?
        If (blnInventoryReportWeekly = True) Then

            ' Time period
            Dim strNow As String = WeekdayName(Weekday(DateTime.Now)) & " " & DateTime.Now.ToString("HH:mm")
            Dim strTarget As String = WeekdayName(Weekday(dtmWeeklyInventoryReport)) & " " & dtmWeeklyInventoryReport.ToString("HH:mm")

            ' Is the flag false?
            If (aastrCSVFile(1, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Okay, we're in. Run the report in quiet mode
                    RunInventoryReport(Me, True)

                    ' email Inventory report
                    SendMail(strEmailInventoryReport, "TeamBeesCapstone@gmail.com", "Weekly Inventory Report", "Automated message: See attached Inventory report.", "TeamBeesCapstone@gmail.com", "cincystate123", "InventoryReport.xlsx", True)

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(1, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(1, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub InventoryMonthly()

        ' Run weekly?
        If (blnInventoryReportMonthly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmMonthlyInventoryReport.Day Then
                strTarget = dtmMonthlyInventoryReport.ToString(intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmMonthlyInventoryReport.ToString("dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(2, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Okay, we're in. Run the report in quiet mode
                    RunInventoryReport(Me, True)

                    ' email Inventory report
                    SendMail(strEmailInventoryReport, "TeamBeesCapstone@gmail.com", "Monthly Inventory Report", "Automated message: See attached Inventory report.", "TeamBeesCapstone@gmail.com", "cincystate123", "InventoryReport.xlsx", True)

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(2, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(2, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub InventoryYearly()

        ' Run weekly?
        If (blnInventoryReportYearly = True) Then

            ' Time period
            Dim strNow As String = DateTime.Now.ToString("MM/dd HH:mm")
            Dim strTarget As String
            Dim intDaysInMonth As Integer = DateTime.DaysInMonth(DateTime.Now.Year, dtmYearlyInventoryReport.Month)

            ' Clamp the days within this month
            If intDaysInMonth < dtmYearlyInventoryReport.Day Then
                strTarget = dtmYearlyInventoryReport.ToString("MM/" & intDaysInMonth & " HH:mm")
            Else
                strTarget = dtmYearlyInventoryReport.ToString("MM/dd HH:mm")
            End If

            ' Is the flag false?
            If (aastrCSVFile(3, 1) = "false") Then
                ' If so, is it time to run?
                If (strNow = strTarget) Then

                    ' Okay, we're in. Run the report in quiet mode
                    RunInventoryReport(Me, True)

                    ' email Inventory report
                    SendMail(strEmailInventoryReport, "TeamBeesCapstone@gmail.com", "Yearly Inventory Report", "Automated message: See attached Inventory report.", "TeamBeesCapstone@gmail.com", "cincystate123", "InventoryReport.xlsx", True)

                    ' Alright, we're all done. Let's set the flag and keep on keepin' on.
                    aastrCSVFile(3, 1) = "true"
                    ' Commit changes
                    WriteCSVFile()

                End If
            Else
                ' If not, should I reset the flag?
                If (strNow <> strTarget) Then

                    ' Okay, reset
                    aastrCSVFile(3, 1) = "false"
                    ' Commit changes
                    WriteCSVFile()

                End If
            End If
        End If

    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        If (My.Computer.FileSystem.FileExists("ReportFlags.csv") = False) Then

            ' Create the record file if it does not exist
            WriteCSVFile()

        Else

            ' Read in the CSV file
            ReadCSVFile()

        End If

    End Sub


End Class