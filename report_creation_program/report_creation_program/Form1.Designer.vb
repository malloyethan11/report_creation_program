Imports Excel = Microsoft.Office.Interop.Excel

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnCreditDepositReport = New System.Windows.Forms.Button()
        Me.btnCashDepositReport = New System.Windows.Forms.Button()
        Me.btnInventoryReport = New System.Windows.Forms.Button()
        Me.btnTaxReport = New System.Windows.Forms.Button()
        Me.btnSalesReport = New System.Windows.Forms.Button()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'btnCreditDepositReport
        '
        Me.btnCreditDepositReport.Location = New System.Drawing.Point(54, 290)
        Me.btnCreditDepositReport.Name = "btnCreditDepositReport"
        Me.btnCreditDepositReport.Size = New System.Drawing.Size(135, 46)
        Me.btnCreditDepositReport.TabIndex = 0
        Me.btnCreditDepositReport.Text = "Credit Deposit Report"
        Me.btnCreditDepositReport.UseVisualStyleBackColor = True
        '
        'btnCashDepositReport
        '
        Me.btnCashDepositReport.Location = New System.Drawing.Point(54, 225)
        Me.btnCashDepositReport.Name = "btnCashDepositReport"
        Me.btnCashDepositReport.Size = New System.Drawing.Size(135, 46)
        Me.btnCashDepositReport.TabIndex = 1
        Me.btnCashDepositReport.Text = "Cash Deposit Report"
        Me.btnCashDepositReport.UseVisualStyleBackColor = True
        '
        'btnInventoryReport
        '
        Me.btnInventoryReport.Location = New System.Drawing.Point(54, 157)
        Me.btnInventoryReport.Name = "btnInventoryReport"
        Me.btnInventoryReport.Size = New System.Drawing.Size(135, 46)
        Me.btnInventoryReport.TabIndex = 2
        Me.btnInventoryReport.Text = "Inventory Report"
        Me.btnInventoryReport.UseVisualStyleBackColor = True
        '
        'btnTaxReport
        '
        Me.btnTaxReport.Location = New System.Drawing.Point(54, 93)
        Me.btnTaxReport.Name = "btnTaxReport"
        Me.btnTaxReport.Size = New System.Drawing.Size(135, 46)
        Me.btnTaxReport.TabIndex = 3
        Me.btnTaxReport.Text = "Tax Report"
        Me.btnTaxReport.UseVisualStyleBackColor = True
        '
        'btnSalesReport
        '
        Me.btnSalesReport.Location = New System.Drawing.Point(54, 29)
        Me.btnSalesReport.Name = "btnSalesReport"
        Me.btnSalesReport.Size = New System.Drawing.Size(135, 46)
        Me.btnSalesReport.TabIndex = 4
        Me.btnSalesReport.Text = "Sales Report"
        Me.btnSalesReport.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.Location = New System.Drawing.Point(54, 376)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(135, 46)
        Me.btnExit.TabIndex = 5
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(51, 345)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(139, 13)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "______________________"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(242, 434)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnExit)
        Me.Controls.Add(Me.btnSalesReport)
        Me.Controls.Add(Me.btnTaxReport)
        Me.Controls.Add(Me.btnInventoryReport)
        Me.Controls.Add(Me.btnCashDepositReport)
        Me.Controls.Add(Me.btnCreditDepositReport)
        Me.Name = "frmMain"
        Me.Text = "Home"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnCreditDepositReport As Button
    Friend WithEvents btnCashDepositReport As Button
    Friend WithEvents btnInventoryReport As Button
    Friend WithEvents btnTaxReport As Button
    Friend WithEvents btnSalesReport As Button

    Private Sub btnSalesReport_Click(sender As Object, e As EventArgs) Handles btnSalesReport.Click

        ' declare variables
        Dim frmSalesReportInfo = New frmSalesReportInfo()

        ' open intermediate form to get email address and time period
        frmSalesReportInfo.ShowDialog()

    End Sub

    Private Sub btnTaxReport_Click(sender As Object, e As EventArgs) Handles btnTaxReport.Click

        ' declare variables
        Dim strUserInput As String

        ' show messagebox asking user where they want to send the report
        strUserInput = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

        ' don't progress if user enters blank input/presses cancel
        If strUserInput = "" Then
            Exit Sub
        End If

        ' get tax report
        RunTaxReport()

    End Sub

    Private Sub btnInventoryReport_Click(sender As Object, e As EventArgs) Handles btnInventoryReport.Click

        ' declare variables
        Dim strUserInput As String

        ' show messagebox asking user where they want to send the report
        strUserInput = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

        ' don't progress if user enters blank input/presses cancel
        If strUserInput = "" Then
            Exit Sub
        End If

        ' run inventory report
        RunInventoryReport()

    End Sub

    Private Sub btnCashDepositReport_Click(sender As Object, e As EventArgs) Handles btnCashDepositReport.Click

        ' declare variables
        Dim strUserInput As String

        ' show messagebox asking user where they want to send the report
        strUserInput = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

        ' don't progress if user enters blank input/presses cancel
        If strUserInput = "" Then
            Exit Sub
        End If

    End Sub

    Private Sub btnCreditDepositReport_Click(sender As Object, e As EventArgs) Handles btnCreditDepositReport.Click

        ' declare variables
        Dim strUserInput As String

        ' show messagebox asking user where they want to send the report
        strUserInput = InputBox("Please enter the email address you want to send the report to.", "User Input Required")

    End Sub

    Friend WithEvents btnExit As Button

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        ' close the form
        Me.Close()

    End Sub

    Private Sub RunTaxReport()



    End Sub

    Private Sub RunInventoryReport()

        ' instantiate excel objects and declare variables
        ' THE EXCEL CODE IS BASED ON THIS TUTORIAL: https://www.tutorialspoint.com/vb.net/vb.net_excel_sheet.htm
        Dim ExcelApp As Excel.Application
        Dim ExcelWkBk As Excel.Workbook
        Dim ExcelWkSht As Excel.Worksheet
        Dim ExcelRange As Excel.Range
        Dim intNumRecords As Integer
        Dim intIndex As Integer = 2 ' starts at 2 to account for header row, Excel rows are also 1-based
        Dim intRecordIndex As Integer = 0

        ' start excel and get application object
        ExcelApp = CreateObject("Excel.Application")
        ExcelApp.Visible = True ' for testing only, set to false when go to prod

        ' Add a new workbook
        ExcelWkBk = ExcelApp.Workbooks.Add
        ExcelWkSht = ExcelWkBk.ActiveSheet

        ' add table headers
        ExcelWkSht.Cells(1, 1) = "SKU"
        ExcelWkSht.Cells(1, 2) = "Item Name"
        ExcelWkSht.Cells(1, 3) = "Item Description"
        ExcelWkSht.Cells(1, 4) = "Vendor Name"
        ExcelWkSht.Cells(1, 5) = "Retail Price"
        ExcelWkSht.Cells(1, 6) = "Current Inventory"
        ExcelWkSht.Cells(1, 7) = "Safety Stock"
        ExcelWkSht.Cells(1, 8) = "UPC"


        ' add table data
        Try

            ' Init select statement string
            Dim strSelect As String = ""
            ' Init select statement Db command
            Dim cmdSelect As OleDb.OleDbCommand
            ' Init data reader
            Dim drSourceTable As OleDb.OleDbDataReader
            ' Init data table
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
            strSelect = "SELECT strSKU, strItemName, strItemDesc, strVendorName, decItemPrice, intInventoryAmt, intSafetyStockAmt, strUPC FROM TItems, TVendors WHERE intInventoryAmt <= intSafetyStockAmt and TItems.intVendorID = TVendors.intVendorID"

            ' Retrieve all the records 
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)
            drSourceTable = cmdSelect.ExecuteReader

            ' load table from data reader
            dt.Load(drSourceTable)

            ' add data to excel spreadsheet
            intNumRecords = dt.Rows.Count

            While intIndex <= (intNumRecords + 1)

                ExcelWkSht.Cells(intIndex, 1).Value = dt.Rows.Item(intRecordIndex).ItemArray(0)
                ExcelWkSht.Cells(intIndex, 2).Value = dt.Rows.Item(intRecordIndex).ItemArray(1)
                ExcelWkSht.Cells(intIndex, 3).Value = dt.Rows.Item(intRecordIndex).ItemArray(2)
                ExcelWkSht.Cells(intIndex, 4).Value = dt.Rows.Item(intRecordIndex).ItemArray(3)
                ExcelWkSht.Cells(intIndex, 5).Value = dt.Rows.Item(intRecordIndex).ItemArray(4)
                ExcelWkSht.Cells(intIndex, 6).Value = dt.Rows.Item(intRecordIndex).ItemArray(5)
                ExcelWkSht.Cells(intIndex, 7).Value = dt.Rows.Item(intRecordIndex).ItemArray(6)
                ExcelWkSht.Cells(intIndex, 8).Value = dt.Rows.Item(intRecordIndex).ItemArray(7)
                intIndex += 1
                intRecordIndex += 1

            End While

            ' close the database connection
            CloseDatabaseConnection()

        Catch excError As Exception

            ' Log and display error message
            MessageBox.Show(excError.Message)

        End Try

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

    Friend WithEvents Label1 As Label
End Class
