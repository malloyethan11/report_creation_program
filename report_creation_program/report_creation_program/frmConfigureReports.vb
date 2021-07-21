Imports System.Data.OleDb

Public Class frmConfigureReports
    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        ' close the form
        Me.Close()

    End Sub

    Private Sub btnAccept_Click(sender As Object, e As EventArgs) Handles btnAccept.Click

        ' reset control colors
        cboReportTypes.BackColor = Color.White
        chklstFrequency.BackColor = Color.White
        txtHour.BackColor = Color.White
        txtMinute.BackColor = Color.White
        rdoAM.BackColor = Color.White
        rdoPM.BackColor = Color.White

        ' declare variables
        Dim strReportType As String
        Dim dtmReportStartDate As String = dtmPickDate.Value
        Dim strDaily As String = "NO"
        Dim strWeekly As String = "NO"
        Dim strMonthly As String = "NO"
        Dim strYearly As String = "NO"
        Dim intHours As Integer
        Dim intMinutes As Integer
        Dim strAMorPM As String
        Dim strReportTime As String

        ' input validation
        If InputValidation() = True Then

            strReportType = cboReportTypes.SelectedItem
            AssignFrequencies(strDaily, strWeekly, strMonthly, strYearly)
            intHours = txtHour.Text
            intMinutes = txtMinute.Text
            strAMorPM = "AM"
            strAMorPM = "PM"
            strReportTime = intHours & ":" & intMinutes & " " & strAMorPM

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

                ' insert config data
                Dim cmdInsert = New OleDbCommand("INSERT INTO TReportConfigs(strReportType, strIsDaily, strIsWeekly, strIsMonthly, strIsYearly, strDateStarted, strTimeOfDayRun) VALUES(?,?,?,?,?,?,?)")
                cmdInsert.CommandType = CommandType.Text
                cmdInsert.Connection = m_conAdministrator
                cmdInsert.Parameters.AddWithValue("strReportType", strReportType)
                cmdInsert.Parameters.AddWithValue("strIsDaily", strDaily)
                cmdInsert.Parameters.AddWithValue("strIsWeekly", strWeekly)
                cmdInsert.Parameters.AddWithValue("strIsMonthly", strMonthly)
                cmdInsert.Parameters.AddWithValue("strIsYearly", strYearly)
                cmdInsert.Parameters.AddWithValue("strDateStarted", dtmReportStartDate)
                cmdInsert.Parameters.AddWithValue("strTimeOfDayRun", strReportTime)
                Dim result = cmdInsert.ExecuteNonQuery()

                If result > 0 Then
                    MessageBox.Show("Configuration successfully added.")
                End If

                ' close database connection
                CloseDatabaseConnection()

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End If


    End Sub

    Private Sub btnViewConfig_Click(sender As Object, e As EventArgs) Handles btnViewConfig.Click

        ' reset control color
        cboReportTypes.BackColor = Color.White

        ' declare variable
        Dim strReportType As String

        If cboReportTypes.SelectedItem Is Nothing Then
            cboReportTypes.BackColor = Color.Yellow
            MessageBox.Show("Please select the report type configuration to view.")

        Else
            strReportType = cboReportTypes.SelectedItem
            frmViewConfig.ShowDialog()

        End If

    End Sub

    Private Function InputValidation() As Boolean

        Dim strReportType As String
        Dim intHours As Integer
        Dim intMinutes As Integer
        Dim strAMorPM As String
        Dim blnResult As Boolean = False

        If cboReportTypes.SelectedItem Is Nothing Then
            cboReportTypes.BackColor = Color.Yellow
            MessageBox.Show("Please select a report type to configure.")

        Else
            If chklstFrequency.SelectedItem Is Nothing Then
                chklstFrequency.BackColor = Color.Yellow
                MessageBox.Show("Please select the frequency at which the report should be run.")

            Else
                If txtHour.Text = "" Then
                    txtHour.BackColor = Color.Yellow
                    MessageBox.Show("Please enter the hour at which the report should be run.")

                Else
                    If txtMinute.Text = "" Then
                        txtMinute.BackColor = Color.Yellow
                        MessageBox.Show("Please enter the minute at which the report should be run.")

                    Else
                        If rdoAM.Checked = False And rdoPM.Checked = False Then
                            rdoAM.BackColor = Color.Yellow
                            rdoPM.BackColor = Color.Yellow
                            MessageBox.Show("Please select AM or PM.")

                        Else
                            If rdoAM.Checked = True Then
                                blnResult = True
                            Else
                                blnResult = True
                            End If

                        End If

                    End If

                End If

            End If

        End If

        Return blnResult

    End Function

    Private Sub AssignFrequencies(ByRef strDaily As String, ByRef strWeekly As String, ByRef strMonthly As String, ByRef strYearly As String)

        If chklstFrequency.GetItemChecked(0) = True Then
            strDaily = "YES"
        End If

        If chklstFrequency.GetItemChecked(1) = True Then
            strWeekly = "YES"
        End If

        If chklstFrequency.GetItemChecked(2) = True Then
            strMonthly = "YES"
        End If

        If chklstFrequency.GetItemChecked(3) = True Then
            strYearly = "YES"
        End If

    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click

        ' reset control colors
        cboReportTypes.BackColor = Color.White
        chklstFrequency.BackColor = Color.White
        txtHour.BackColor = Color.White
        txtMinute.BackColor = Color.White
        rdoAM.BackColor = Color.White
        rdoPM.BackColor = Color.White

        ' declare variables
        Dim strReportType As String
        Dim dtmReportStartDate As String = dtmPickDate.Value
        Dim strDaily As String = "NO"
        Dim strWeekly As String = "NO"
        Dim strMonthly As String = "NO"
        Dim strYearly As String = "NO"
        Dim intHours As Integer
        Dim intMinutes As Integer
        Dim strAMorPM As String
        Dim strReportTime As String

        ' input validation
        If InputValidation() = True Then

            strReportType = cboReportTypes.SelectedItem
            AssignFrequencies(strDaily, strWeekly, strMonthly, strYearly)
            intHours = txtHour.Text
            intMinutes = txtMinute.Text
            strAMorPM = "AM"
            strAMorPM = "PM"
            strReportTime = intHours & ":" & intMinutes & " " & strAMorPM

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

                ' insert config data
                Dim cmdInsert = New OleDbCommand("UPDATE TReportConfigs SET strIsDaily = ?, strIsWeekly = ?, strIsMonthly = ?, strIsYearly = ?, strDateStarted = ?, strTimeOfDayRun = ? WHERE strReportType = '" & strReportType & "'")
                cmdInsert.CommandType = CommandType.Text
                cmdInsert.Connection = m_conAdministrator
                cmdInsert.Parameters.AddWithValue("@strIsDaily", strDaily)
                cmdInsert.Parameters.AddWithValue("@strIsWeekly", strWeekly)
                cmdInsert.Parameters.AddWithValue("@strIsMonthly", strMonthly)
                cmdInsert.Parameters.AddWithValue("@strIsYearly", strYearly)
                cmdInsert.Parameters.AddWithValue("@strDateStarted", dtmReportStartDate)
                cmdInsert.Parameters.AddWithValue("@strTimeOfDayRun", strReportTime)
                Dim result = cmdInsert.ExecuteNonQuery()

                If result > 0 Then
                    MessageBox.Show("Configuration successfully updated.")
                End If

                ' close database connection
                CloseDatabaseConnection()

            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try

        End If

    End Sub
End Class