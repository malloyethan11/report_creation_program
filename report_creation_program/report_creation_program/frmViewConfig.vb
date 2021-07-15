Public Class frmViewConfig
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click

        ' close the form
        Me.Close()

    End Sub

    Private Sub frmViewConfig_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ' declare variables
        Dim intReportTypeID As Integer
        Dim strReportType As String = frmConfigureReports.cboReportTypes.SelectedItem
        Dim strDaily As String
        Dim strWeekly As String
        Dim strMonthly As String
        Dim strYearly As String
        Dim strDay As String
        Dim strTime As String
        Dim strSelect As String = ""
        Dim cmdSelect As OleDb.OleDbCommand
        Dim drSourceTable As OleDb.OleDbDataReader
        Dim dt As DataTable = New DataTable

        Try

            ' Open the DB
            If OpenDatabaseConnectionSQLServer() = False Then

                ' The database is not open
                MessageBox.Show(Me, "Database connection error." & vbNewLine &
                                "The form will now close.",
                                Me.Text + " Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error)

                ' Close the form/application
                Me.Close()

            End If

            ' Build the select statement
            strSelect = "SELECT intReportConfigID, strReportType, strIsDaily, strIsWeekly, strIsMonthly, strIsYearly, strDateStarted, strTimeOfDayRun FROM TReportConfigs WHERE strReportType = '" & strReportType & "'"

            ' Retrieve all the records 
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader

            ' load table from data reader
            dt.Load(drSourceTable)

            Dim results As Object = dt.Rows.Item(0).ItemArray
            lblReportType.Text = results(1)
            lblDaily.Text = results(2)
            lblWeekly.Text = results(3)
            lblMonthly.Text = results(4)
            lblYearly.Text = results(5)
            lblDayofWeek.Text = results(6)
            lblTime.Text = results(7)

            ' Clean up
            drSourceTable.Close()

            ' close the database connection
            CloseDatabaseConnection()

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try

    End Sub
End Class