﻿Imports Excel = Microsoft.Office.Interop.Excel

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
                CreateSalesReport(strTimePeriod)

            End If

        End If

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click

        ' close the form
        Me.Close()

    End Sub

    Private Sub CreateSalesReport(ByVal strTimePeriod As String)

        ' instantiate excel objects and declare variables
        ' THE EXCEL CODE IS BASED ON THIS TUTORIAL: https://www.tutorialspoint.com/vb.net/vb.net_excel_sheet.htm
        Dim ExcelApp As Excel.Application
        Dim ExcelWkBk As Excel.Workbook
        Dim ExcelWkSht As Excel.Worksheet
        Dim ExcelRange As Excel.Range

        ' start excel and get application object
        ExcelApp = CreateObject("Excel.Application")
        ExcelApp.Visible = True ' for testing only, set to false when go to prod

        ' Add a new workbook
        ExcelWkBk = ExcelApp.Workbooks.Add
        ExcelWkSht = ExcelWkBk.ActiveSheet

        ' add table headers going cell by cell
        ExcelWkSht.Cells(1, 1) = "Sales for the " & strTimePeriod
        ExcelWkSht.Cells(2, 1) = "Medals"
        ExcelWkSht.Cells(2, 2) = "Statues"
        ExcelWkSht.Cells(2, 3) = "Books"
        ExcelWkSht.Cells(2, 4) = "Church Goods"
        ExcelWkSht.Cells(2, 5) = "Tokens"
        ExcelWkSht.Cells(2, 6) = "Baptism"
        ExcelWkSht.Cells(2, 7) = "Rosary"
        ExcelWkSht.Cells(2, 8) = "Wedding"
        ExcelWkSht.Cells(2, 9) = "Anniversary"
        ExcelWkSht.Cells(2, 10) = "Cards"
        ExcelWkSht.Cells(2, 11) = "Holy Cards"
        ExcelWkSht.Cells(2, 12) = "Pictures/Artwork"
        ExcelWkSht.Cells(2, 13) = "Confirmation"
        ExcelWkSht.Cells(2, 14) = "First Communion"

        ' add table data
        ExcelWkSht.Cells(3, 1) = GetSales(1, strTimePeriod)
        ExcelWkSht.Cells(3, 2) = GetSales(2, strTimePeriod)
        ExcelWkSht.Cells(3, 3) = GetSales(3, strTimePeriod)
        ExcelWkSht.Cells(3, 4) = GetSales(4, strTimePeriod)
        ExcelWkSht.Cells(3, 5) = GetSales(5, strTimePeriod)
        ExcelWkSht.Cells(3, 6) = GetSales(6, strTimePeriod)
        ExcelWkSht.Cells(3, 7) = GetSales(7, strTimePeriod)
        ExcelWkSht.Cells(3, 8) = GetSales(8, strTimePeriod)
        ExcelWkSht.Cells(3, 9) = GetSales(9, strTimePeriod)
        ExcelWkSht.Cells(3, 10) = GetSales(10, strTimePeriod)
        ExcelWkSht.Cells(3, 11) = GetSales(11, strTimePeriod)
        ExcelWkSht.Cells(3, 12) = GetSales(12, strTimePeriod)
        ExcelWkSht.Cells(3, 13) = GetSales(13, strTimePeriod)
        ExcelWkSht.Cells(3, 14) = GetSales(14, strTimePeriod)

        ' Release object references.
        ExcelRange = Nothing
        ExcelWkSht = Nothing
        ExcelWkBk = Nothing
        ExcelApp.Quit()
        ExcelApp = Nothing
        Exit Sub
Err_Handler:
        MsgBox(Err.Description, vbCritical, "Error: " & Err.Number)

    End Sub

    Private Function GetSales(ByVal intCategory As Integer, ByVal strTimePeriod As String)

        Dim dblTotalSales As Double

        Try

            Dim strSelect As String
            Dim cmdSelect As OleDb.OleDbCommand
            Dim dt As DataTable = New DataTable

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

            ' Build the select statement based on user-selected time period
            If strTimePeriod = "last day" Then
                strSelect = "SELECT SUM(decItemPrice) from ItemsSoldByCategoryWithPriceAndDate where intCategoryID = " & intCategory & " AND strPurchaseDate > (DATEADD(DAY, -1, GETDATE()))"

            ElseIf strTimePeriod = "last week" Then
                strSelect = "Select SUM(decItemPrice) from ItemsSoldByCategoryWithPriceAndDate where intCategoryID = " & intCategory & " and strpurchasedate > (DATEADD(DAY, -7, GETDATE()))"

            ElseIf strTimePeriod = "last month (30 days)" Then
                strSelect = "SELECT SUM(decItemPrice) from ItemsSoldByCategoryWithPriceAndDate where intCategoryID = " & intCategory & " and strpurchasedate > (DATEADD(DAY, -30, GETDATE()))"

            ElseIf strTimePeriod = "last year (365 days)" Then
                strSelect = "SELECT SUM(decItemPrice) from ItemsSoldByCategoryWithPriceAndDate where intCategoryID = " & intCategory & " and strpurchasedate > (DATEADD(DAY, -365, GETDATE()))"

            End If

            ' Retrieve all the records 
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            Dim objTotalSales As Object = cmdSelect.ExecuteScalar

            ' check for null entries (zeroes), set to zero
            If IsDBNull(objTotalSales) Then
                dblTotalSales = 0
            Else
                dblTotalSales = CDbl(objTotalSales)
            End If

            ' close the database connection
            CloseDatabaseConnection()

        Catch excError As Exception

            ' Log and display error message
            MessageBox.Show(excError.Message)

        End Try

        Return dblTotalSales

    End Function

End Class